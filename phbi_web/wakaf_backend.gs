/**
 * Backend konfirmasi wakaf Masjid Ar Ridhwan.
 * Mendukung Konfirmasi Mandiri (Transfer + Bukti) dan Mode Bendahara (Tunai / Transfer, Bukti Opsional).
 */
const SPREADSHEET_ID = '1F9HUDR27NEVIKH3lAg-VKdgjl7gvQVc0opfqqnzAdGk';
const RESPONSE_SHEET_NAME = 'Form Responses 1';
const DRIVE_FOLDER_NAME = 'Bukti Transfer Wakaf Masjid Ar Ridhwan';
const MAX_AUTH_ATTEMPTS = 3;

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

/**
 * Menangani permintaan JSONP (GET) untuk autentikasi password Bendahara.
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = String(params.callback || '').replace(/[^a-zA-Z0-9_$.]/g, '');

  function sendJsonp(obj) {
    const payload = JSON.stringify(obj);
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${payload})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOutput_(obj);
  }

  try {
    if (params.action === 'auth') {
      const inputPassword = String(params.password || '');
      const clientId = String(params.clientId || 'default').slice(0, 64);
      const props = PropertiesService.getScriptProperties();

      // Membaca dari BENDAHARA_PASSWORD (atau fallback jika masih huruf kecil)
      const savedPassword = props.getProperty('BENDAHARA_PASSWORD') || props.getProperty('BENDahara_PASSWORD');

      if (!savedPassword) {
        return sendJsonp({ ok: false, message: 'Password bendahara belum diatur di Script Properties.' });
      }

      const cache = CacheService.getScriptCache();
      const attemptKey = 'auth_fail_' + clientId;
      let attempts = Number(cache.get(attemptKey) || 0);

      if (attempts >= MAX_AUTH_ATTEMPTS) {
        return sendJsonp({ ok: false, locked: true, remaining: 0 });
      }

      if (inputPassword === savedPassword) {
        cache.remove(attemptKey);
        const token = Utilities.getUuid();
        // Token sesi berlaku 2 jam (7200 detik)
        cache.put('token_' + token, 'valid', 7200);
        return sendJsonp({ ok: true, token: token });
      } else {
        attempts += 1;
        cache.put(attemptKey, String(attempts), 1800); // Blokir 30 menit jika 3x salah
        const remaining = Math.max(0, MAX_AUTH_ATTEMPTS - attempts);
        return sendJsonp({
          ok: false,
          locked: attempts >= MAX_AUTH_ATTEMPTS,
          remaining: remaining
        });
      }
    }

    return sendJsonp({ ok: false, message: 'Aksi tidak dikenali.' });
  } catch (err) {
    return sendJsonp({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}

/**
 * Menangani pengiriman data konfirmasi wakaf (POST).
 */
function doPost(e) {
  let uploadedFile = null;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Data permintaan kosong.');
    const data = JSON.parse(e.postData.contents);

    const isBendahara = verifyBendaharaToken_(data.token);
    validatePayload_(data, isBendahara);

    let fileName = '-';
    let fileUrl = '-';

    // Proses upload file jika ada fileData yang dikirim
    if (data.fileData && data.fileName && data.fileMimeType) {
      const bytes = Utilities.base64Decode(data.fileData);
      const safeName = String(data.fileName).replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 150);
      const blob = Utilities.newBlob(bytes, data.fileMimeType, safeName);
      const folder = getOrCreateDriveFolder_();
      uploadedFile = folder.createFile(blob);
      uploadedFile.setDescription(
        isBendahara
          ? 'Bukti transaksi wakaf diinput oleh Bendahara Masjid Ar Ridhwan.'
          : 'Bukti transfer wakaf dari formulir konfirmasi mandiri Masjid Ar Ridhwan.'
      );
      fileName = uploadedFile.getName();
      fileUrl = uploadedFile.getUrl();
    } else if (isBendahara) {
      fileName = 'Dicatat Bendahara (Tanpa Lampiran)';
    }

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
      sheet.appendRow([
        new Date(),
        data.tanggal,
        data.nama.trim(),
        Number(data.nominal),
        data.pilihanWakaf,
        data.jenisTransaksi,
        fileName,
        fileUrl
      ]);
    } finally {
      lock.releaseLock();
    }

    return jsonOutput_({ ok: true, message: 'Konfirmasi tersimpan.' });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    if (uploadedFile) {
      try { uploadedFile.setTrashed(true); } catch (ignore) {}
    }
    return jsonOutput_({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}

function verifyBendaharaToken_(token) {
  if (!token) return false;
  const cache = CacheService.getScriptCache();
  return cache.get('token_' + String(token)) === 'valid';
}

function validatePayload_(d, isBendahara) {
  if (!d || !d.tanggal || !d.nama || !d.nominal || !d.pilihanWakaf) {
    throw new Error('Data wajib belum lengkap.');
  }
  if (!['Tunai', 'Transfer'].includes(d.jenisTransaksi)) {
    throw new Error('Jenis transaksi tidak valid.');
  }
  if (d.jenisTransaksi === 'Tunai' && !isBendahara) {
    throw new Error('Transaksi Tunai hanya dapat diinput oleh Bendahara.');
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.tanggal)) {
    throw new Error('Format tanggal tidak valid.');
  }
  if (Number(d.nominal) <= 0 || !isFinite(Number(d.nominal))) {
    throw new Error('Nominal tidak valid.');
  }
  if (!['Renovasi Masjid', 'Sound System'].includes(d.pilihanWakaf)) {
    throw new Error('Pilihan wakaf tidak valid.');
  }

  // Untuk Muwakif biasa (bukan Bendahara), bukti transfer wajib dilampirkan
  if (!isBendahara && d.jenisTransaksi === 'Transfer') {
    if (!d.fileData || !d.fileName || !d.fileMimeType) {
      throw new Error('Bukti transfer wajib dilampirkan.');
    }
  }

  // Jika melampirkan file (baik Muwakif maupun Bendahara), validasi tipe dan ukuran
  if (d.fileData) {
    if (!['image/jpeg', 'application/pdf'].includes(d.fileMimeType)) {
      throw new Error('Bukti harus JPG/JPEG atau PDF.');
    }
    const decodedLength = Utilities.base64Decode(d.fileData).length;
    if (decodedLength > 5 * 1024 * 1024) {
      throw new Error('Ukuran bukti transfer melebihi 5 MB.');
    }
  }
}

function getOrCreateDriveFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
