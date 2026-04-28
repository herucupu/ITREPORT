/**
 * KONFIGURASI SISTEM
 */
const CONFIG = {
  PASSWORD: "vander", // Ganti password di sini
  SHEET_NAME: "DAILY_REPORT_DB"
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IT Daily Report System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi Inisialisasi Database
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = ["ID", "Tanggal", "Pekerjaan", "Status", "Keterangan", "Dokumen", "Validasi", "Catatan_Atasan", "Timestamp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight("bold").setBackground("#2563eb").setFontColor("white");
    sheet.setFrozenRows(1);
  }
  return "Database Berhasil Disiapkan!";
}

function checkLogin(pass) {
  return pass === CONFIG.PASSWORD;
}

function fetchData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();
}

function saveReport(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const id = "ID-" + new Date().getTime();
    sheet.appendRow([id, data.tanggal, data.detail, data.status, data.keterangan, data.dokumen, "Waiting", "", new Date()]);
    return { success: true, message: "Laporan berhasil disimpan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateStatusOnly(id, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      return "Status Diperbarui";
    }
  }
}

function updateApproval(id, status, catatan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 7).setValue(status);
      sheet.getRange(i + 1, 8).setValue(catatan);
      return { success: true, message: "Laporan " + status };
    }
  }
}