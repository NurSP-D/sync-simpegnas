// ============================================================================
// 1. PUBLIC API & FRONT-END CONTROLLER
// ============================================================================

/**
 * Mengambil status/progress sinkronisasi saat ini dari ScriptProperties untuk UI
 */
function getSyncProgress() {
  const props = PropertiesService.getScriptProperties();
  const status = props.getProperty('SYNC_STATUS') || 'IDLE';
  
  const processedKantor = parseInt(props.getProperty('SYNC_PROCESSED_KANTOR') || '0', 10);
  const totalKantor = parseInt(props.getProperty('SYNC_TOTAL_KANTOR') || '0', 10);
  const totalPegawai = parseInt(props.getProperty('SYNC_TOTAL_PEGAWAI') || '0', 10);
  const bulan = parseInt(props.getProperty('SYNC_BULAN') || (new Date().getMonth() + 1), 10);
  const tahun = parseInt(props.getProperty('SYNC_TAHUN') || new Date().getFullYear(), 10);
  const errorMsg = props.getProperty('SYNC_ERROR_MSG') || '';

  // Hitung persentase
  let percent = 0;
  if (totalKantor > 0) {
    percent = Math.floor((processedKantor / totalKantor) * 100);
    if (percent > 100) percent = 100;
  }

  return {
    status: status,
    percent: percent,
    processedKantor: processedKantor,
    totalKantor: totalKantor,
    totalPegawai: totalPegawai,
    bulan: bulan,
    tahun: tahun,
    message: errorMsg
  };
}

/**
 * Memulai / Melanjutkan Sinkronisasi Presensi dari UI
 */
function startSyncPresensi(token, bulan, tahun, forceReset) {
  try {
    const props = PropertiesService.getScriptProperties();

    if (forceReset === true || forceReset === "true") resetSyncStatusProperties();

    let activeToken = token && token.trim() !== "" ? token.trim() : (props.getProperty('CURRENT_TOKEN') || props.getProperty('SYNC_TOKEN'));
    if (!activeToken) throw new Error("Token presensi kosong!");
    
    props.setProperty('CURRENT_TOKEN', activeToken);
    props.setProperty('SYNC_TOKEN', activeToken);

    const sheetKantor = getSheetDataKantorMaster();
    if (!sheetKantor || sheetKantor.getLastRow() < CONFIG_SYNC.ROW_DATA_KANTOR_START) {
      throw new Error("Data Master Kantor tidak ditemukan!");
    }

    initSheetLog();

    const prevStatus = props.getProperty('SYNC_STATUS');
    const isResume = (!forceReset && ['PAUSED_TOKEN_EXPIRED', 'ERROR_TOKEN', 'ERROR_SYSTEM', 'PAUSED_MANUAL'].includes(prevStatus));

    if (isResume) {
      const bTarget = props.getProperty('SYNC_BULAN') || bulan;
      const tTarget = props.getProperty('SYNC_TAHUN') || tahun;
      const lastLog = getLastUnfinishedLog(bTarget, tTarget);
      const lastRowProcessed = parseInt(props.getProperty('SYNC_LAST_ROW') || CONFIG_SYNC.ROW_DATA_KANTOR_START, 10);

      if (lastLog && lastLog.startRowRekap >= CONFIG_SYNC.ROW_DATA_REKAP_START) {
        const ssTarget = getOrCreateTargetSpreadsheet(bTarget, tTarget);
        const sheetRekap = ssTarget.getSheetByName(getNamaSheetRekap(bTarget, tTarget));
        if (sheetRekap && sheetRekap.getLastRow() >= lastLog.startRowRekap) {
          const numRowsToDelete = sheetRekap.getLastRow() - lastLog.startRowRekap + 1;
          if (numRowsToDelete > 0) sheetRekap.deleteRows(lastLog.startRowRekap, numRowsToDelete);
        }
      }

      props.setProperties({ 'SYNC_STATUS': 'PROCESSING', 'SYNC_LAST_ROW': String(lastRowProcessed), 'SYNC_ERROR_MSG': '' });
    } else {
      resetSyncStatusProperties();
      const totalKantor = sheetKantor.getLastRow() - (CONFIG_SYNC.ROW_DATA_KANTOR_START - 1);

      props.setProperties({
        'SYNC_STATUS': 'PROCESSING',
        'SYNC_BULAN': String(bulan),
        'SYNC_TAHUN': String(tahun),
        'SYNC_LAST_ROW': String(CONFIG_SYNC.ROW_DATA_KANTOR_START),
        'SYNC_PROCESSED_KANTOR': '0',
        'SYNC_TOTAL_KANTOR': String(totalKantor),
        'SYNC_TOTAL_PEGAWAI': '0',
        'SYNC_ERROR_MSG': ''
      });

      // 1. Inisialisasi Sheet Rekap
      initSheetRekap(bulan, tahun);

      // 2. [PERBAIKAN] KOSONGKAN DATA LAMA PADA SHEET REKAP
      const ssTarget = getOrCreateTargetSpreadsheet(bulan, tahun);
      if (ssTarget) {
        const sheetRekap = ssTarget.getSheetByName(getNamaSheetRekap(bulan, tahun));
        const startRow = CONFIG_SYNC.ROW_DATA_REKAP_START || 2; // Default baris ke-2 (di bawah header)
        
        if (sheetRekap && sheetRekap.getLastRow() >= startRow) {
          const numRows = sheetRekap.getLastRow() - startRow + 1;
          const numCols = sheetRekap.getLastColumn();
          
          // Menghapus isi sel dari baris awal data sampai baris terakhir tanpa merusak format/header
          sheetRekap.getRange(startRow, 1, numRows, numCols).clearContent();
        }
      }
    }

    createSyncTrigger();
    processBatchPresensi();

    return { success: true, message: isResume ? "Sync dilanjutkan!" : "Sync baru dimulai!", isResume };
  } catch (error) {
    removeSyncTrigger();
    PropertiesService.getScriptProperties().setProperties({ 'SYNC_STATUS': 'ERROR_SYSTEM', 'SYNC_ERROR_MSG': error.message });
    return { success: false, message: error.message };
  }
}

/**
 * Membatalkan / Menghentikan Sinkronisasi Presensi dari UI
 */
function stopSyncPresensi() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SYNC_STATUS', 'PAUSED_MANUAL');
    removeSyncTrigger();
    return { success: true, message: "Sinkronisasi berhasil dihentikan." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================================
// 2. CORE BATCH PROCESSING
// ============================================================================
function processBatchPresensi() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  const props = PropertiesService.getScriptProperties();

  try {
    const startTime = Date.now();
    if (props.getProperty('SYNC_STATUS') !== 'PROCESSING') {
      removeSyncTrigger();
      return;
    }

    const token = props.getProperty('SYNC_TOKEN');
    const bulan = parseInt(props.getProperty('SYNC_BULAN'), 10);
    const tahun = parseInt(props.getProperty('SYNC_TAHUN'), 10);
    const startRow = parseInt(props.getProperty('SYNC_LAST_ROW'), 10);
    const processedKantor = parseInt(props.getProperty('SYNC_PROCESSED_KANTOR') || '0', 10);
    const totalKantor = parseInt(props.getProperty('SYNC_TOTAL_KANTOR') || '0', 10);

    const ssTarget = getOrCreateTargetSpreadsheet(bulan, tahun);
    const sheetKantor = getSheetDataKantorMaster();
    const sheetRekap = ssTarget.getSheetByName(getNamaSheetRekap(bulan, tahun));
    
    if (!sheetKantor || !sheetRekap) return;

    const maxRowInSheet = sheetKantor.getLastRow();
    if (startRow > maxRowInSheet) {
      completeSyncProcess();
      return;
    }

    const kantorData = sheetKantor.getRange(startRow, 1, maxRowInSheet - startRow + 1, 2).getValues();
    const totalHari = new Date(tahun, bulan, 0).getDate();
    
    let kantorBerhasil = 0;
    let isTokenInvalid = false;
    let nextTargetRow = Math.max(sheetRekap.getLastRow() + 1, CONFIG_SYNC.ROW_DATA_REKAP_START);

    for (let i = 0; i < kantorData.length; i++) {
      if (Date.now() - startTime >= CONFIG_SYNC.MAX_EXECUTION_TIME_MS) break;

      const idKantor = kantorData[i][0];
      const namaKantor = kantorData[i][1];
      if (!idKantor) { kantorBerhasil++; continue; }

      logSyncStatus(idKantor, namaKantor, bulan, tahun, 'IN_PROGRESS', nextTargetRow, 0, 'Memproses API');

      let apiRes = null, attempt = 0;
      while (attempt < 3) {
        attempt++;
        apiRes = fetchRekapPresensiByKantor(idKantor, tahun, bulan, token);
        if (apiRes && (apiRes.statusCode === 200 || apiRes.statusCode === 401)) break;
        Utilities.sleep(1000);
      }

      if (!apiRes) {
        logSyncStatus(idKantor, namaKantor, bulan, tahun, 'FAILED', nextTargetRow, 0, 'Gagal API');
        kantorBerhasil++;
        continue;
      }

      if (apiRes.statusCode === 401) {
        logSyncStatus(idKantor, namaKantor, bulan, tahun, 'TOKEN_EXPIRED', nextTargetRow, 0, 'Token kadaluwarsa');
        isTokenInvalid = true;
        break;
      }

      if (apiRes.statusCode === 200 && apiRes.json) {
        const listPegawai = extractPegawaiList(apiRes.json);
        if (listPegawai.length > 0) {
          const formattedRows = formatDataRekapPegawai(namaKantor, listPegawai, totalHari, tahun, bulan);
          if (formattedRows.length > 0) {
            // Pastikan jumlah kolom di sheet cukup sebelum menulis data
            const reqCols = formattedRows[0].length;
            if (sheetRekap.getMaxColumns() < reqCols) {
              sheetRekap.insertColumnsAfter(sheetRekap.getMaxColumns(), reqCols - sheetRekap.getMaxColumns());
            }

            sheetRekap.getRange(nextTargetRow, 1, formattedRows.length, reqCols).setValues(formattedRows);
            SpreadsheetApp.flush();
            logSyncStatus(idKantor, namaKantor, bulan, tahun, 'SUCCESS', nextTargetRow, formattedRows.length, 'Selesai');
            nextTargetRow += formattedRows.length;
          }
        } else {
          logSyncStatus(idKantor, namaKantor, bulan, tahun, 'SUCCESS', nextTargetRow, 0, 'Data kosong');
        }
      }

      kantorBerhasil++;
      props.setProperties({
        'SYNC_PROCESSED_KANTOR': String(processedKantor + kantorBerhasil),
        'SYNC_TOTAL_PEGAWAI': String(Math.max(0, nextTargetRow - CONFIG_SYNC.ROW_DATA_REKAP_START)),
        'SYNC_LAST_ROW': String(startRow + kantorBerhasil)
      });
    }

    if (isTokenInvalid) {
      props.setProperty('SYNC_STATUS', 'ERROR_TOKEN');
      removeSyncTrigger();
      return;
    }

    if (processedKantor + kantorBerhasil >= totalKantor || startRow + kantorBerhasil > maxRowInSheet) {
      completeSyncProcess();
    }
  } catch (err) {
    props.setProperties({ 'SYNC_STATUS': 'ERROR_SYSTEM', 'SYNC_ERROR_MSG': err.message });
    removeSyncTrigger();
  } finally {
    lock.releaseLock();
  }
}

function completeSyncProcess() {
  const props = PropertiesService.getScriptProperties();
  const bulan = parseInt(props.getProperty('SYNC_BULAN'), 10);
  const tahun = parseInt(props.getProperty('SYNC_TAHUN'), 10);
  
  const ssTarget = getOrCreateTargetSpreadsheet(bulan, tahun);
  const namaSheetPresensi = getNamaSheetRekap(bulan, tahun);

  // Jalankan PEMBERSIHAN DUPLIKAT & STATISTIK di file spreadsheet baru
  hapusDuplikatPresensi(ssTarget, namaSheetPresensi);
  generateStatistikPresensi(ssTarget, bulan, tahun);
  generateStatistikASN(ssTarget, bulan, tahun);

  props.setProperty('SYNC_STATUS', 'COMPLETED');
  removeSyncTrigger();
}

// ============================================================================
// 3. GENERATOR SHEET & POST-PROCESSING
// ============================================================================
function initSheetRekap(bulan, tahun) {
  const ssTarget = getOrCreateTargetSpreadsheet(bulan, tahun);
  const namaSheet = getNamaSheetRekap(bulan, tahun);

  let sheet = ssTarget.getSheetByName(namaSheet) || ssTarget.insertSheet(namaSheet);

  // Hapus sheet default bawaan (misal: "Sheet1" atau "Lembar1") jika ada sheet lain
  const defaultSheet = ssTarget.getSheetByName("Sheet1") || ssTarget.getSheetByName("Lembar1");
  if (defaultSheet && ssTarget.getSheets().length > 1) {
    try { ssTarget.deleteSheet(defaultSheet); } catch(e){}
  }

  if (sheet.getLastRow() < CONFIG_SYNC.ROW_HEADER_REKAP) {
    sheet.clear();
    
    // Hitung total hari secara dinamis sesuai bulan dan tahun
    const totalHari = new Date(tahun, bulan, 0).getDate();
    
    // Header langsung di Baris 1
    const headers = ["Nama Kantor / SKPD", "Nama Pegawai", "NIP", "Tahun", "Bulan"];
    for (let d = 1; d <= totalHari; d++) {
      headers.push(
        `Tgl ${d} (WF)`, 
        `Tgl ${d} (In)`, 
        `Tgl ${d} (Status In)`, 
        `Tgl ${d} (Out)`, 
        `Tgl ${d} (Status Out)`, 
        `Tgl ${d} (Late)`, 
        `Tgl ${d} (Status Harian)`
      );
    }

    // --- TAMBAHAN KUNCI: Tambah kolom otomatis jika kolom sheet bawaan kurang ---
    const maxCols = sheet.getMaxColumns();
    if (maxCols < headers.length) {
      sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#e0e0e0");
    sheet.setFrozenRows(1);
  }
}

function hapusDuplikatPresensi(ssTarget, namaSheetTarget) {
  if (!ssTarget) return;
  const sheet = ssTarget.getSheetByName(namaSheetTarget);
  if (sheet && sheet.getLastRow() >= CONFIG_SYNC.ROW_DATA_REKAP_START) {
    // Mengecek duplikasi berdasarkan kombinasi: Kolom 1 (Kantor), Kolom 2 (Nama), Kolom 3 (NIP)
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).removeDuplicates([1, 2, 3]);
  }
}

function generateStatistikPresensi(ssTarget, bulan, tahun) {
  const sheetPresensi = ssTarget.getSheetByName(getNamaSheetRekap(bulan, tahun));
  if (!sheetPresensi || sheetPresensi.getLastRow() < CONFIG_SYNC.ROW_DATA_REKAP_START) return;

  const data = sheetPresensi.getRange(CONFIG_SYNC.ROW_DATA_REKAP_START, 1, sheetPresensi.getLastRow() - 1, sheetPresensi.getLastColumn()).getValues();
  const rekapKantor = {}; 

  for (let i = 0; i < data.length; i++) {
    const namaKantor = data[i][0];
    if (!namaKantor) continue;
    if (!rekapKantor[namaKantor]) rekapKantor[namaKantor] = { jumlahPegawai: 0, totalHadir: 0, totalTerlambat: 0, totalTanpaKeterangan: 0 };

    rekapKantor[namaKantor].jumlahPegawai++;
    for (let col = CONFIG_SYNC.COL_DAY_START_INDEX; col < data[i].length; col += CONFIG_SYNC.COL_HARIAN_STEP) {
      const statusIn = String(data[i][col + 2] || "").toLowerCase();     
      const late = String(data[i][col + 5] || "");                    
      const statusHarian = String(data[i][col + 6] || "").toLowerCase(); 

      if (statusIn !== "" || statusHarian !== "") rekapKantor[namaKantor].totalHadir++;
      if ((late !== "" && late !== "0" && late !== "00:00:00") || statusIn.includes("terlambat")) rekapKantor[namaKantor].totalTerlambat++;
      if (statusHarian.includes("alpa") || statusHarian.includes("tanpa keterangan") || statusHarian === "tk") rekapKantor[namaKantor].totalTanpaKeterangan++;
    }
  }

  const sheetStat = ssTarget.getSheetByName(getNamaSheetStatistik(bulan, tahun)) || ssTarget.insertSheet(getNamaSheetStatistik(bulan, tahun));
  sheetStat.clear();

  const headers = ["No", "Nama Kantor / SKPD", "Jumlah Pegawai", "Total Record Kehadiran", "Total Terlambat", "Total Tanpa Keterangan / Alpa"];
  sheetStat.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
  sheetStat.setFrozenRows(1);

  const outputRows = [];
  let no = 1, grandPegawai = 0, grandHadir = 0, grandTerlambat = 0, grandAlpa = 0;

  for (const k in rekapKantor) {
    const item = rekapKantor[k];
    outputRows.push([no++, k, item.jumlahPegawai, item.totalHadir, item.totalTerlambat, item.totalTanpaKeterangan]);
    grandPegawai += item.jumlahPegawai; grandHadir += item.totalHadir; grandTerlambat += item.totalTerlambat; grandAlpa += item.totalTanpaKeterangan;
  }

  if (outputRows.length > 0) {
    sheetStat.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    sheetStat.getRange(2 + outputRows.length, 1, 1, headers.length).setValues([["", "TOTAL KESELURUHAN", grandPegawai, grandHadir, grandTerlambat, grandAlpa]]).setFontWeight("bold").setBackground("#fff2cc");
  }
}

/**
 * REVISI: generateStatistikASN
 * Mendukung header baru (ITM, IPC, ITMPC, IDLI, IDLO) dan parsing status kombinasi/komposit
 */
function generateStatistikASN(ssTarget, bulan, tahun) {
  const sheetData = ssTarget.getSheetByName(getNamaSheetRekap(bulan, tahun));
  if (!sheetData || sheetData.getLastRow() < CONFIG_SYNC.ROW_DATA_REKAP_START) return;

  const sheetStat = ssTarget.getSheetByName(getNamaSheetStatistikASN(bulan, tahun)) || ssTarget.insertSheet(getNamaSheetStatistikASN(bulan, tahun));
  sheetStat.clear();

  // Header baru dengan penambahan: ITM, IPC, ITMPC, IDLI, IDLO
  const headers = [
    "Nama Kantor / SKPD", "Nama Pegawai", "NIP", "Tahun", "Bulan", 
    "HN", "TK", "TM1", "TM2", "TM3", "TMM", "PC1", "PC2", "PC3", "PCM", 
    "ITM", "IPC", "IDL", "IDLI", "IDLO", 
    "DL", "CT", "CB", "CS", "CM", "CKAP", "LJ", "LN"
  ];
  sheetStat.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#405d72").setFontColor("#ffffff");
  sheetStat.setFrozenRows(1);

  const rawData = sheetData.getRange(CONFIG_SYNC.ROW_DATA_REKAP_START, 1, sheetData.getLastRow() - 1, sheetData.getLastColumn()).getValues();
  const outputRows = [];

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row[2] && !row[1]) continue;

    // Inisialisasi counter status
    const counts = { 
      HN: 0, TK: 0, TM1: 0, TM2: 0, TM3: 0, TMM: 0, PC1: 0, PC2: 0, PC3: 0, PCM: 0, 
      ITM: 0, IPC: 0, IDL: 0, IDLI: 0, IDLO: 0,
      DL: 0, CT: 0, CB: 0, CS: 0, CM: 0, CKAP: 0, LJ: 0, LN: 0 
    };

    for (let col = CONFIG_SYNC.COL_STATUS_HARIAN_START_INDEX; col < row.length; col += CONFIG_SYNC.COL_HARIAN_STEP) {
      const rawKode = String(row[col] || "").trim().toUpperCase();
      if (!rawKode) continue;

      // Jika ada status gabungan yang dipisah tanda '-', kita bedah jadi komponen individu
      const subKodes = rawKode.split('-').map(function(s) { return s.trim(); });

      // Track kuesioner terhitung per hari agar tidak terhitung ganda untuk status tunggal
      var processedInThisDay = {};

      for (var k = 0; k < subKodes.length; k++) {
        var kD = subKodes[k];
        if (counts.hasOwnProperty(kD) && !processedInThisDay[kD]) {
          counts[kD]++;
          processedInThisDay[kD] = true;
        }
      }

      // Penanganan khusus untuk penanda khusus seperti ITMPC yang juga dihitung sebagai ITM & IPC
      if (rawKode.indexOf("ITMPC") !== -1) {
        if (!processedInThisDay["ITMPC"]) { counts.ITMPC++; processedInThisDay["ITMPC"] = true; }
        if (!processedInThisDay["ITM"]) { counts.ITM++; processedInThisDay["ITM"] = true; }
        if (!processedInThisDay["IPC"]) { counts.IPC++; processedInThisDay["IPC"] = true; }
      }
    }

    outputRows.push([
      row[0], row[1], `'${row[2]}`, tahun, padZero(bulan), 
      counts.HN, counts.TK, counts.TM1, counts.TM2, counts.TM3, counts.TMM, 
      counts.PC1, counts.PC2, counts.PC3, counts.PCM, 
      counts.ITM, counts.IPC, counts.IDL, counts.IDLI, counts.IDLO,
      counts.DL, counts.CT, counts.CB, counts.CS, counts.CM, counts.CKAP, counts.LJ, counts.LN
    ]);
  }

  if (outputRows.length > 0) {
    sheetStat.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  }
}

function extractPegawaiList(json) {
  if (!json) return [];
  if (Array.isArray(json)) return json;
  if (Array.isArray(json.data)) return json.data;
  if (json.data && Array.isArray(json.data.pegawai)) return json.data.pegawai;
  if (Array.isArray(json.pegawai)) return json.pegawai;
  return [];
}

/**
 * ============================================================================
 * API Simpegnas
 * ============================================================================
 */

function fetchRekapPresensiByKantor(idKantor, tahun, bulan, token) {
  try {
    var cleanId = String(idKantor).trim();
    var cleanTahun = String(tahun).trim();
    var cleanBulan = String(bulan).trim();
    var cleanToken = String(token).trim();

    var url = "https://api-absensi.simpegnas.go.id/absensi/api/get/rekap-bulanan-by-kantor" +
              "?kantor_id=" + encodeURIComponent(cleanId) +
              "&tahun=" + encodeURIComponent(cleanTahun) +
              "&bulan=" + encodeURIComponent(cleanBulan);

    var options = {
      "method": "get",
      "headers": {
        "presensi-key": cleanToken,
        "accept": "application/json"
      },
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();
    var jsonResult = null;

    if (statusCode === 200) {
      try {
        jsonResult = JSON.parse(contentText);
      } catch (e) {
        Logger.log("Gagal parsing JSON Kantor ID [" + cleanId + "]: " + e.message);
      }
    }

    return {
      statusCode: statusCode,
      json: jsonResult,
      errorText: statusCode !== 200 ? contentText : ""
    };

  } catch (error) {
    Logger.log("Exception pada fetchRekapPresensiByKantor ID [" + idKantor + "]: " + error.message);
    return {
      statusCode: 500,
      json: null,
      errorText: error.message
    };
  }
}

function formatDataRekapPegawai(namaKantor, listPegawai, totalHari, tahunParam, bulanParam) {
  var rows = [];

  if (!Array.isArray(listPegawai) || listPegawai.length === 0) {
    return rows;
  }

  listPegawai.forEach(function(pegawai) {
    var namaPegawai = pegawai.nama || pegawai.nama_pegawai || pegawai.name || "";
    var nipPegawai = pegawai.nip || pegawai.nip_baru || pegawai.pegawai_nip || "";
    
    // Gunakan fallback ke parameter pilihan jika dari API kosong
    var thnValue = pegawai.tahun || pegawai.tahun_baru || pegawai.pegawai_tahun || tahunParam;
    var blnValue = pegawai.bulan || pegawai.bulan_baru || pegawai.pegawai_bulan || bulanParam;

    var rowValues = [namaKantor, namaPegawai, nipPegawai, thnValue, blnValue];

    var daftarPresensi = pegawai.presensi || pegawai.rekap || pegawai.data_presensi || [];

    for (var d = 1; d <= totalHari; d++) {
      var presensiHari = null;

      if (Array.isArray(daftarPresensi)) {
        presensiHari = daftarPresensi.find(function(p) { 
          return parseInt(p.day || p.tgl || p.tanggal, 10) === d; 
        });
      }

      if (presensiHari) {
        var cIn = presensiHari.checkIn || presensiHari.check_in || {};
        var cOut = presensiHari.checkOut || presensiHari.check_out || {};

        rowValues.push(cIn.work_from || cIn.wf || "");
        rowValues.push(cIn.time_with_timezone || cIn.jam || cIn.time || "");
        rowValues.push(cIn.status || "");
        rowValues.push(cOut.time_with_timezone || cOut.jam || cOut.time || "");
        rowValues.push(cOut.status || "");
        rowValues.push(presensiHari.late || presensiHari.terlambat || "");
        rowValues.push(presensiHari.status || presensiHari.keterangan || "");
      } else {
        rowValues.push("", "", "", "", "", "", "");
      }
    }

    rows.push(rowValues);
  });

  return rows;
}
