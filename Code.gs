// Fungsi utama untuk menampilkan Web App
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Aplikasi Rekap Presensi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi pembantu untuk memuat file HTML terpisah ke dalam index.html
function includePage(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
