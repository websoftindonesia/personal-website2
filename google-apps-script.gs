/**
 * Database pendaftaran Learning Session -> Google Spreadsheet
 *
 * Cara pakai:
 * 1. Buat Google Sheet baru, lalu buka Extensions > Apps Script.
 * 2. Tempel seluruh kode ini, simpan.
 * 3. Deploy > New deployment > jenis "Web app".
 *      Execute as : Me
 *      Who has access : Anyone
 * 4. Salin URL Web app (berakhiran /exec) dan tempel ke variabel SCRIPT_URL
 *    pada <script id="learning-session"> di index.html.
 * 5. Setiap kali kode ini diubah, buat "New deployment" / "Manage deployments > Edit > New version".
 */
var SHEET_NAME = 'Pendaftar';
var HEADERS = ['Timestamp', 'Kursus', 'Nama', 'Email', 'WhatsApp', 'Instansi', 'Jurusan/Semester'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      clean(data.kursus), clean(data.nama), clean(data.email),
      "'" + clean(data.whatsapp),          // apostrof: jaga angka 0 di depan
      clean(data.instansi), clean(data.jurusan)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Cegah formula injection di spreadsheet (=, +, -, @ di awal teks)
function clean(v) {
  v = String(v == null ? '' : v).trim().slice(0, 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}