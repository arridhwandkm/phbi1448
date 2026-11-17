/**
 * Backend konfirmasi wakaf Masjid Ar Ridhwan.
 * Menyimpan bukti transfer ke Google Drive dan menambahkan data ke tab
 * "Form Responses 1" dengan urutan kolom A-H agar laporan lama tetap bekerja:
 * A Timestamp | B Tanggal | C Nama | D Nominal | E Pilihan Wakaf |
 * F Jenis Transaksi | G Nama File Bukti | H URL Bukti
 *
 * SETUP:
 * 1. Buka https://script.google.com/ dan buat project baru.
 * 2. Ganti seluruh isi Code.gs dengan kode ini.
 * 3. Jalankan setupWakafBackend() satu kali dari editor, lalu izinkan akses.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me; Who has access: Anyone.
 * 5. Salin URL Web App yang berakhiran /exec ke WAKAF_WEB_APP_URL di test.html.
 * 6. Unggah ulang test.html ke hosting website Anda.
 *
 * Batasi ukuran unggahan di frontend (maks. 5 MB). Web App yang terbuka untuk
 * publik dapat menerima spam; tinjau izin dan pantau data yang masuk.
 */
const SPREADSHEET_ID = '1F9HUDR27NEVIKH3lAg-VKdgjl7gvQVc0opfqqnzAdGk';
const RESPONSE_SHEET_NAME = 'Form Responses 1';
const DRIVE_FOLDER_NAME = 'Bukti Transfer Wakaf Masjid Ar Ridhwan';

function setupWakafBackend() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(RESPONSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Tanggal Wakaf', 'Nama Lengkap', 'Nominal Wakaf', 'Wakaf yang dipilih', 'Pilihan Transaksi', 'Nama File Bukti Transfer', 'URL Bukti Transfer']);
  } else {
    if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('Nama File Bukti Transfer');
    if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('URL Bukti Transfer');
  }
  getOrCreateDriveFolder_();
  return 'Setup selesai.';
}

function doPost(e) {
  let uploadedFile = null;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Data permintaan kosong.');
    const data = JSON.parse(e.postData.contents);
    validatePayload_(data);

    const bytes = Utilities.base64Decode(data.fileData);
    const safeName = String(data.fileName).replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 150);
    const blob = Utilities.newBlob(bytes, data.fileMimeType, safeName);
    const folder = getOrCreateDriveFolder_();
    uploadedFile = folder.createFile(blob);
    uploadedFile.setDescription('Bukti transfer wakaf dari formulir konfirmasi mandiri Masjid Ar Ridhwan.');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(RESPONSE_SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Tanggal Wakaf', 'Nama Lengkap', 'Nominal Wakaf', 'Wakaf yang dipilih', 'Pilihan Transaksi', 'Nama File Bukti Transfer', 'URL Bukti Transfer']);
    } else {
      if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('Nama File Bukti Transfer');
      if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('URL Bukti Transfer');
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      sheet.appendRow([new Date(), data.tanggal, data.nama.trim(), Number(data.nominal), data.pilihanWakaf, 'Transfer', uploadedFile.getName(), uploadedFile.getUrl()]);
    } finally {
      lock.releaseLock();
    }
    return jsonOutput_({ok: true, message: 'Konfirmasi tersimpan.'});
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    // Jika gagal sebelum pencatatan sheet, hapus file yang telanjur dibuat.
    if (uploadedFile) {
      try { uploadedFile.setTrashed(true); } catch (ignore) {}
    }
    return jsonOutput_({ok: false, message: String(err && err.message ? err.message : err)});
  }
}

function validatePayload_(d) {
  if (!d || !d.tanggal || !d.nama || !d.nominal || !d.pilihanWakaf) throw new Error('Data wajib belum lengkap.');
  if (d.jenisTransaksi !== 'Transfer') throw new Error('Form mandiri hanya menerima transaksi Transfer.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.tanggal)) throw new Error('Format tanggal tidak valid.');
  if (Number(d.nominal) <= 0 || !isFinite(Number(d.nominal))) throw new Error('Nominal tidak valid.');
  if (!['Renovasi Masjid', 'Sound System'].includes(d.pilihanWakaf)) throw new Error('Pilihan wakaf tidak valid.');
  if (!d.fileData || !d.fileName || !d.fileMimeType) throw new Error('Bukti transfer wajib dilampirkan.');
  if (!['image/jpeg', 'application/pdf'].includes(d.fileMimeType)) throw new Error('Bukti harus JPG/JPEG atau PDF.');
  const decodedLength = Utilities.base64Decode(d.fileData).length;
  if (decodedLength > 5 * 1024 * 1024) throw new Error('Ukuran bukti transfer melebihi 5 MB.');
}

function getOrCreateDriveFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
