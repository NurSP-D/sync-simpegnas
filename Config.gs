// ============================================================================
// 1. KONFIGURASI UTAMA
// ============================================================================
const CONFIG_SYNC = {
  SHEET_KANTOR: "DataKantor",           // Nama sheet master daftar kantor
  SHEET_LOG: "Log_Sync",                // Sheet log di Spreadsheet Master

  MAX_EXECUTION_TIME_MS: 4 * 60 * 1000, // Time Guard: Max 4 menit per batch
  TRIGGER_INTERVAL_MINUTES: 1,          // Interval pemicu otomatis

  ROW_DATA_KANTOR_START: 2,             // Baris awal data kantor di DataKantor
  ROW_HEADER_REKAP: 1,                  // Header langsung di Baris 1
  ROW_DATA_REKAP_START: 2,              // Data pegawai mulai Baris 2

  COL_DAY_START_INDEX: 5,                
  COL_STATUS_HARIAN_START_INDEX: 11,    
  COL_HARIAN_STEP: 7                    
};

const CONFIG = {
  FOLDER_ID: "ID_FOLDER_PADA_DRIVE_GOOGLE", // Masukkan ID Folder 'REKAP PRESENSI'
  DATA_MASTER_ID: "ID_SPREADSHEET ",  // Masukkan ID SPREADSHEET 'Data Master'
  SHEET_DATA_KANTOR: "DataKantor",    // Nama Sheet Data kantor
  SHEET_PPPKPW: "PPPKPW"                         // Nama Sheet Data master PPPKPW
};

// KONSTANTA CONFIGURASI REKAP PRESENSI
var CONFIG_REKAP = {
  FOLDER_ID: "ID_FOLDER_PADA_DRIVE_GOOGLE", // ID Folder REKAP PRESENSI
  ROW_HEADER: 1,                                  // Baris tempat nama-nama kolom berada
  ROW_DATA_START: 2                               // Baris awal data pegawai dimulai
};

// KONSTANTA CONFIGURASI REKAP STATISTIK PRESENSI
var CONFIG_STATISTIK_DRIVE = {
  FOLDER_ID: "ID_FOLDER_PADA_DRIVE_GOOGLE", // ID Folder REKAP PRESENSI
  ROW_HEADER: 1,      // Baris letak Header
  ROW_DATA_START: 2  // Baris awal data pegawai
};

// Header Standar / Output Statistik
var HEADERS_STATISTIK_DEFAULT = [
  "Nama Pegawai", "NIP", 
  "HN", "TK", "TM1", "TM2", "TM3", "TMM", "PC1", "PC2", "PC3", "PCM", 
  "ITM", "IPC", "IDL", "IDLI", "IDLO",
  "DL", "CT", "CB", "CS", "CM", "CKAP", "LJ", "LN"
];


// ============================================================================
// 2. HELPER MANAGEMENT FILE & SPREADSHEET
// ============================================================================
const padZero = (val) => String(parseInt(val, 10)).padStart(2, '0');

function getFolderRekap() {
  try {
    if (CONFIG.FOLDER_ID) return DriveApp.getFolderById(CONFIG.FOLDER_ID);
  } catch (err) {
    Logger.log("⚠️ Gagal mengambil folder: " + err.message);
  }
  return null;
}

function getSheetDataKantorMaster() {
  try {
    if (CONFIG.DATA_MASTER_ID) {
      const ssMaster = SpreadsheetApp.openById(CONFIG.DATA_MASTER_ID);
      const sheet = ssMaster.getSheetByName(CONFIG.SHEET_DATA_KANTOR);
      if (sheet) return sheet;
    }
  } catch (err) {
    Logger.log("⚠️ Gagal membuka Master Spreadsheet via ID: " + err.message);
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SYNC.SHEET_KANTOR);
}

/**
 * Otomatis mencari atau membuat Spreadsheet baru "Presensi_YYYY_MM" di folder target
 */
function getOrCreateTargetSpreadsheet(bulan, tahun) {
  const folder = getFolderRekap();
  if (!folder) {
    throw new Error("Folder 'REKAP PRESENSI' tidak ditemukan! Periksa CONFIG.FOLDER_ID.");
  }

  const targetName = `Presensi_${tahun}_${padZero(bulan)}`;
  const files = folder.getFilesByName(targetName);

  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    // Buat spreadsheet baru
    const newSS = SpreadsheetApp.create(targetName);
    const file = DriveApp.getFileById(newSS.getId());
    
    // Pindahkan file ke folder target secara modern
    file.moveTo(folder);
    
    return newSS;
  }
}

function getNamaSheetRekap(bulan, tahun) {
  return `Rekap_Presensi_${tahun}_${padZero(bulan)}`;
}

function getNamaSheetStatistik(bulan, tahun) {
  return `Statistik_${tahun}_${padZero(bulan)}`;
}

function getNamaSheetStatistikASN(bulan, tahun) {
  return `Statistik_ASN_${tahun}_${padZero(bulan)}`;
}

// ============================================================================
// 3. HELPER & UTILITAS STATUS ASN (PNS, PPPK, PPPKPW)
// ============================================================================

/**
 * Mengambil daftar NIP PPPKPW dari Sheet PPPKPW di Spreadsheet Master.
 * Mengembalikan objek Set agar proses pencarian (lookup) super cepat O(1).
 * 
 * @returns {Set<string>} Set berisi daftar NIP PPPKPW
 */
function getSetNIPPPPKPW() {
  try {
    if (!CONFIG.DATA_MASTER_ID) return new Set();
    
    const ssMaster = SpreadsheetApp.openById(CONFIG.DATA_MASTER_ID);
    const sheetPPPKPW = ssMaster.getSheetByName(CONFIG.SHEET_PPPKPW);
    
    if (!sheetPPPKPW) {
      Logger.log("⚠️ Sheet '" + CONFIG.SHEET_PPPKPW + "' tidak ditemukan di Data Master.");
      return new Set();
    }

    const lastRow = sheetPPPKPW.getLastRow();
    if (lastRow < 2) return new Set(); // Hanya header atau kosong

    // Ambil Kolom A mulai baris ke-2
    const dataNip = sheetPPPKPW.getRange(2, 1, lastRow - 1, 1).getValues();
    const setNip = new Set();

    dataNip.forEach(row => {
      let nip = String(row[0] || '').trim();
      if (nip !== '') {
        setNip.add(nip);
      }
    });

    return setNip;
  } catch (err) {
    Logger.log("⚠️ Gagal membaca data NIP PPPKPW: " + err.message);
    return new Set();
  }
}

/**
 * Menentukan Kategori ASN (PNS, PPPK, atau PPPKPW) dari NIP
 * 
 * @param {string} nip - NIP Pegawai
 * @param {Set<string>} setNipPPPKPW - Object Set hasil pemanggilan getSetNIPPPPKPW()
 * @returns {string} 'PNS', 'PPPK', atau 'PPPKPW'
 */
function getKategoriASN(nip, setNipPPPKPW) {
  const strNip = String(nip || '').trim();
  
  // Digit ke-13 dan 14 (indeks string 12 dan 13)
  const kodePengangkatan = strNip.substring(12, 14);

  // 1. Jika digit 13-14 BUKAN 21 -> PNS/CPNS
  if (kodePengangkatan !== '21') {
    return 'PNS';
  }

  // 2. Jika digit 13-14 ADALAH 21, cek di Set PPPKPW
  if (setNipPPPKPW && setNipPPPKPW.has(strNip)) {
    return 'PPPKPW'; // Ada di sheet PPPKPW (True)
  } else {
    return 'PPPK';   // Tidak ada di sheet PPPKPW (False)
  }
}

/**
 * ----------------------------------------------------------------------------
 * 4. FUNGSI DEBUGGER: testDebugDataRekapFolder
 * ----------------------------------------------------------------------------
 * Uji coba pembacaan file dari folder secara independen di Apps Script Editor.
 */
function testDebugDataRekapFolder() {
  Logger.log("=== MEMULAI PENGUJIAN BACKEND DINAMIS ===");
  
  var testBulan = "01";
  var testTahun = "2026";

  // Test 1: Buka Sheet Periode 01-2026
  var sheet = getTargetSheetRekap(testBulan, testTahun);
  if (!sheet) {
    Logger.log("❌ ERROR: File/Sheet untuk periode " + testBulan + "-" + testTahun + " tidak dapat dibuka!");
    return;
  }

  Logger.log("✅ Berhasil Akses Sheet: " + sheet.getName());
  Logger.log("✅ Total Baris: " + sheet.getLastRow());
  Logger.log("✅ Total Kolom: " + sheet.getLastColumn());

  // Test 2: Mengambil Daftar Kantor
  var listKantor = getDaftarKantorRekap(testBulan, testTahun);
  Logger.log("\n--- DAFTAR KANTOR UNIK ---");
  Logger.log("Jumlah Kantor Ditemukan: " + listKantor.length);
  Logger.log("Sampel Kantor: " + JSON.stringify(listKantor.slice(0, 3)));

  // Test 3: Filter Rekap Data PPPKPW
  var resData = getRekapDataFiltered({
    bulan: testBulan,
    tahun: testTahun,
    namaKantor: "",
    statusASN: "PPPKPW" // Menguji filter baru
  });

  Logger.log("\n--- HASIL FILTER PPPKPW ---");
  Logger.log("Jumlah Baris Header: " + resData.headers.length);
  Logger.log("Jumlah Baris Data PPPKPW Ditemukan: " + resData.data.length);

  Logger.log("\n=== PENGUJIAN SELESAI ===");
}

// ============================================================================
// 4. LOGGING & STATE MANAGEMENT
// ============================================================================
function initSheetLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SYNC.SHEET_LOG);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SYNC.SHEET_LOG);
  
  if (sheet.getLastRow() < 1) {
    const headers = ["Timestamp", "ID Kantor", "Nama Kantor", "Bulan", "Tahun", "Status", "Start Row Rekap", "Jumlah Pegawai", "Pesan Keterangan"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#cfe2f3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logSyncStatus(idKantor, namaKantor, bulan, tahun, status, startRowRekap, jumlahPegawai, pesan) {
  const sheet = initSheetLog();
  const lastRow = sheet.getLastRow();
  const timestamp = new Date();
  
  const data = lastRow > 1 ? sheet.getRange(2, 2, lastRow - 1, 4).getValues() : [];
  let targetRow = -1;

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(idKantor) && String(data[i][2]) === String(bulan) && String(data[i][3]) === String(tahun)) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow !== -1) {
    sheet.getRange(targetRow, 1).setValue(timestamp);
    sheet.getRange(targetRow, 6).setValue(status);
    if (startRowRekap) sheet.getRange(targetRow, 7).setValue(startRowRekap);
    if (jumlahPegawai !== null && jumlahPegawai !== undefined) sheet.getRange(targetRow, 8).setValue(jumlahPegawai);
    if (pesan !== undefined) sheet.getRange(targetRow, 9).setValue(pesan);
  } else {
    sheet.appendRow([timestamp, idKantor, namaKantor, bulan, tahun, status, startRowRekap || 0, jumlahPegawai || 0, pesan || ""]);
  }
  SpreadsheetApp.flush();
}

function getLastUnfinishedLog(bulan, tahun) {
  const sheet = initSheetLog();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][3]) === String(bulan) && String(data[i][4]) === String(tahun)) {
      if (['IN_PROGRESS', 'FAILED', 'TOKEN_EXPIRED'].includes(data[i][5])) {
        return { startRowRekap: parseInt(data[i][6], 10) };
      }
    }
  }
  return null;
}

function resetSyncStatusProperties() {
  removeSyncTrigger();
  const props = PropertiesService.getScriptProperties();
  ['SYNC_STATUS', 'SYNC_LAST_ROW', 'SYNC_PROCESSED_KANTOR', 'SYNC_TOTAL_KANTOR', 'SYNC_TOTAL_PEGAWAI', 'SYNC_BULAN', 'SYNC_TAHUN'].forEach(p => props.deleteProperty(p));
}

function createSyncTrigger() {
  removeSyncTrigger();
  ScriptApp.newTrigger('processBatchPresensi').timeBased().everyMinutes(CONFIG_SYNC.TRIGGER_INTERVAL_MINUTES).create();
}

function removeSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processBatchPresensi') ScriptApp.deleteTrigger(trigger);
  });
}

// ============================================================================
// 5. DATA TAHUN SERVER & PERMISSION
// ============================================================================
function getCurrentServerYear() {
  return new Date().getFullYear();
}

// KODE PEMICU OTORISASI (JALANKAN FUNGSI INI SEKALI)
function paksaIzinDrive() {
  // Tanpa try-catch agar Apps Script dipaksa meminta izin OAuth
  var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  Logger.log("Akses Berhasil! Nama Folder: " + folder.getName());
}

// KODE PEMICU OTORISASI DRIVE PERMISSION (JALANKAN SEKALI)
function paksaIzinDriveFull() {
  var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  // Panggil dummy file untuk memicu scope DriveApp.File.moveTo
  var tempFile = DriveApp.createFile("Temp_Test_Perm.txt", "test");
  tempFile.moveTo(folder);
  tempFile.setTrashed(true); // Hapus kembali file tes
  Logger.log("Berhasil Otorisasi Drive Full!");
}
