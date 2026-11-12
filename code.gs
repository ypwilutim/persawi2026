const CONFIG_PLACEHOLDER = 'TEMPEL_FOLDER_ID_DISINI';
const DEFAULT_MAX_MB = 25;
const DEFAULT_MAX_R2_MB = 2048;
const SESSION_HOURS = 12;
const CONFIG_KEYS = ['FOLDER_ID', 'EVENT_NAME', 'EVENT_SUBTITLE', 'EVENT_SLOGAN', 'SEKOLAH', 'USER_COUNTER', 'MAX_FILE_MB', 'MAX_R2_MB'];
const R2_REGION = 'auto';
const R2_SERVICE = 's3';
const R2_UPLOAD_EXPIRY_SECONDS = 6 * 60 * 60;
const R2_PROPERTY_KEYS = ['R2_ACCOUNT_ID', 'R2_ACCESS_KEY_ID', 'R2_SECRET_ACCESS_KEY', 'R2_BUCKET', 'R2_PUBLIC_BASE'];
const API_VERSION = '2026-11-12-r2-album';

function setupDatabaseSheets() {
  ensureDatabase_();
}

function setupR2Config() {
  const props = PropertiesService.getScriptProperties();
  R2_PROPERTY_KEYS.forEach(function (key) {
    if (!props.getProperty(key)) props.setProperty(key, '');
  });
  return 'Script Property R2 siap. Isi nilainya di Project Settings → Script Properties.';
}

function testR2Connection() {
  const lines = [];
  let key = '';

  try {
    const config = getR2Config_();
    lines.push('OK: Script Property R2 lengkap.');

    key = 'test/connection-' + Date.now() + '.txt';
    const signed = r2Presign_('PUT', key, 'PutObject');
    lines.push('OK: Presigned PUT URL dibuat.');
    lines.push('Access key length: ' + config.accessKeyId.length);
    lines.push('Secret length: ' + config.secretAccessKey.length);
    lines.push('PUT URL: ' + signed.url);

    const put = UrlFetchApp.fetch(signed.url, {
      method: 'put',
      contentType: 'text/plain',
      payload: 'persawi-r2-test-' + Date.now(),
      muteHttpExceptions: true
    });
    lines.push('PUT status: ' + put.getResponseCode());
    if (put.getResponseCode() < 200 || put.getResponseCode() >= 300) {
      lines.push('PUT body: ' + put.getContentText().substring(0, 2000));
      lines.push('GAGAL: Upload test ditolak. Periksa Access Key, Secret, dan bucket.');
      return lines.join('\n');
    }

    const publicUrl = config.publicBase + '/' + key;
    const head = UrlFetchApp.fetch(publicUrl, {
      method: 'get',
      headers: { Range: 'bytes=0-0' },
      muteHttpExceptions: true,
      followRedirects: true
    });
    lines.push('GET range public status: ' + head.getResponseCode());
    if (head.getResponseCode() !== 200 && head.getResponseCode() !== 206) {
      lines.push('GAGAL: File tidak dapat diakses publik. Aktifkan public access R2.');
      return lines.join('\n');
    }
    lines.push('OK: File dapat diakses di ' + publicUrl);

    const del = r2Presign_('DELETE', key, 'DeleteObject');
    const delRes = UrlFetchApp.fetch(del.url, { method: 'delete', muteHttpExceptions: true });
    lines.push('DELETE status: ' + delRes.getResponseCode());
    if (delRes.getResponseCode() < 200 || delRes.getResponseCode() >= 300) {
      lines.push('GAGAL: Delete test ditolak. Periksa izin token R2.');
      return lines.join('\n');
    }

    lines.push('SUKSES: R2 terhubung dan berfungsi.');
  } catch (e) {
    lines.push('ERROR: ' + e.message);
  }

  const result = lines.join('\n');
  Logger.log(result);
  return result;
}

function ensureDatabase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan.');

  let sheetConfig = ss.getSheetByName('Config');
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet('Config');
    sheetConfig.appendRow(['Key', 'Value']);
    [
      ['FOLDER_ID', CONFIG_PLACEHOLDER],
      ['EVENT_NAME', 'PERSAWI 2026'],
      ['EVENT_SUBTITLE', 'Perkemahan Santri Wahdah Islamiyah'],
      ['EVENT_SLOGAN', 'Tunas Muda Berjiwa Rabbani Menuju Luwu Timur Juara'],
      ['SEKOLAH', 'SDIT Insan Rabbani Malili'],
      ['USER_COUNTER', '0'],
      ['MAX_FILE_MB', String(DEFAULT_MAX_MB)],
      ['MAX_R2_MB', String(DEFAULT_MAX_R2_MB)]
    ].forEach(row => sheetConfig.appendRow(row));
  }

  const configRows = sheetConfig.getDataRange().getValues();
  let hasSekolah = false;
  let hasUserCounter = false;
  let hasMaxR2Mb = false;
  for (let i = 1; i < configRows.length; i++) {
    const key = String(configRows[i][0] == null ? '' : configRows[i][0]).trim();
    if (key === 'SEKOLAH') hasSekolah = true;
    if (key === 'USER_COUNTER') hasUserCounter = true;
    if (key === 'MAX_R2_MB') hasMaxR2Mb = true;
  }
  if (!hasSekolah) {
    sheetConfig.appendRow(['SEKOLAH', 'SDIT Insan Rabbani Malili']);
  }
  if (!hasUserCounter) {
    sheetConfig.appendRow(['USER_COUNTER', '0']);
  }
  if (!hasMaxR2Mb) {
    sheetConfig.appendRow(['MAX_R2_MB', String(DEFAULT_MAX_R2_MB)]);
  }

  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['username', 'password', 'role', 'unit', 'nama', 'telepon']);
    sheetUsers.appendRow(['admin', 'admin2026', 'admin', 'Koordinator YPWI', 'Akbar Irwansya', '']);
    sheetUsers.appendRow(['malili1', 'media2026', 'media', 'SDIT Insan Rabbani Malili', 'Tim Media 1', '']);
  } else {
    const header = sheetUsers.getRange(1, 1, 1, Math.max(sheetUsers.getLastColumn(), 1)).getValues()[0];
    if (header.indexOf('telepon') === -1) {
      sheetUsers.getRange(1, header.length + 1).setValue('telepon');
    }
  }

  let sheetMedia = ss.getSheetByName('Media');
  if (!sheetMedia) {
    sheetMedia = ss.insertSheet('Media');
    sheetMedia.appendRow([
      'id', 'timestamp', 'unit', 'uploader', 'kategori',
      'filename', 'filetype', 'drive_url', 'file_id', 'keterangan', 'task_id', 'storage_type', 'batch_id'
    ]);
  } else {
    const header = sheetMedia.getRange(1, 1, 1, Math.max(sheetMedia.getLastColumn(), 1)).getValues()[0];
    if (header.indexOf('task_id') === -1) {
      sheetMedia.getRange(1, header.length + 1).setValue('task_id');
    }
    if (header.indexOf('storage_type') === -1) {
      sheetMedia.getRange(1, Math.max(sheetMedia.getLastColumn(), header.length) + 1).setValue('storage_type');
    }
    if (header.indexOf('batch_id') === -1) {
      sheetMedia.getRange(1, Math.max(sheetMedia.getLastColumn(), header.length) + 1).setValue('batch_id');
    }
  }

  let sheetRundown = ss.getSheetByName('Rundown');
  if (!sheetRundown) {
    sheetRundown = ss.insertSheet('Rundown');
    sheetRundown.appendRow(['id', 'tanggal', 'jam', 'judul', 'kategori', 'lokasi', 'keterangan', 'created_at']);
  }

  let sheetProgress = ss.getSheetByName('TaskProgress');
  if (!sheetProgress) {
    sheetProgress = ss.insertSheet('TaskProgress');
    sheetProgress.appendRow(['task_id', 'unit', 'status', 'completed_at', 'completed_by', 'media_id', 'keterangan']);
  }

  getAuthSecret_();
}

function getAppConfig() {
  ensureDatabase_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const rows = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0] == null ? '' : String(rows[i][0]).trim();
    if (key) {
      config[key] = rows[i][1] == null ? '' : String(rows[i][1]).trim();
    }
  }
  return config;
}

function doGet(e) {
  try {
    ensureDatabase_();
    const params = (e && e.parameter) || {};
    const action = params.action || '';

    if (action === 'getMedia') return handleGetMedia(params);
    if (action === 'getRundown') return handleGetRundown();
    if (action === 'verifySession') return handleVerifySession(params.token);
    if (action === 'getMyProgress') return handleGetMyProgress(params.token);
    if (action === 'getTaskProgress') return handleGetTaskProgress(params.token);
    if (action === 'getUsers') return handleGetUsers(params.token);
    if (action === 'getConfig') return responseJSON({ status: 'success', config: getAppConfig() });
    if (action === 'ping') return responseJSON({ status: 'ok', version: API_VERSION, build: 'sha256' });
    return responseJSON({ status: 'ok', message: 'API Media Center Aktif', version: API_VERSION });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    ensureDatabase_();

    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const action = data.action;
    const protection = {
      upload: ['media', 'admin'],
      addRundown: ['admin'],
      deleteRundown: ['admin'],
      addUser: ['admin'],
      deleteMedia: ['admin'],
      updateConfig: ['admin'],
      changePassword: ['media', 'admin'],
      completeR2Upload: ['media', 'admin'],
      getUploadUrl: ['media', 'admin'],
      deleteR2Object: ['admin']
    };

    if (Object.prototype.hasOwnProperty.call(protection, action)) {
      const auth = verifyToken_(data.token, protection[action]);
      if (!auth) {
        return responseJSON({ status: 'error', message: 'Sesi berakhir atau tidak valid. Silakan login kembali.' });
      }
      data._auth = auth;
    }

    if (action === 'login') {
      return handleLogin(data.username, data.password);
    } else if (action === 'registerMedia') {
      return handleRegisterMedia(data);
    } else if (action === 'upload') {
      return handleUpload(data);
    } else if (action === 'completeR2Upload') {
      return handleCompleteR2Upload(data);
    } else if (action === 'getUploadUrl') {
      return handleGetUploadUrl(data);
    } else if (action === 'deleteR2Object') {
      return handleDeleteR2Object(data);
    } else if (action === 'addRundown') {
      return handleAddRundown(data);
    } else if (action === 'deleteRundown') {
      return handleDeleteRundown(data.id);
    } else if (action === 'addUser') {
      return handleAddUser(data);
    } else if (action === 'deleteMedia') {
      return handleDeleteMedia(data.id, data.fileId);
    } else if (action === 'updateConfig') {
      return handleUpdateConfig(data.configs);
    } else if (action === 'changePassword') {
      return handleChangePassword(data);
    }

    return responseJSON({ status: 'error', message: 'Aksi tidak dikenal' });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.message });
  }
}

function handleLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const rows = sheet.getDataRange().getValues();
  const inputUser = String(username == null ? '' : username).trim();
  const inputPass = String(password == null ? '' : password);
  const inputPhone = normalizePhone_(inputUser);

  for (let i = 1; i < rows.length; i++) {
    const rowUser = String(rows[i][0] == null ? '' : rows[i][0]).trim();
    const rowPass = String(rows[i][1] == null ? '' : rows[i][1]);
    const rowPhone = normalizePhone_(rowUser) || normalizePhone_(rows[i][5]);
    const matchesUser = rowUser === inputUser || (inputPhone && rowPhone && rowPhone === inputPhone);
    if (matchesUser && rowPass === inputPass) {
      const user = {
        username: rowUser,
        role: String(rows[i][2] == null ? '' : rows[i][2]).trim(),
        unit: String(rows[i][3] == null ? '' : rows[i][3]).trim(),
        nama: String(rows[i][4] == null ? '' : rows[i][4]).trim()
      };
      return responseJSON({
        status: 'success',
        token: createToken_(user),
        user: user
      });
    }
  }
  return responseJSON({ status: 'error', message: 'Username atau Password salah!' });
}

function handleRegisterMedia(data) {
  const nama = String(data.nama == null ? '' : data.nama).trim();
  const unit = String(data.unit == null ? '' : data.unit).trim();
  const telepon = String(data.telepon == null ? '' : data.telepon).trim();

  if (!nama || !unit || !telepon) {
    return responseJSON({ status: 'error', message: 'Nama, asal sekolah, dan nomor telepon wajib diisi.' });
  }

  const normalizedPhone = normalizePhone_(telepon);
  if (!normalizedPhone || !/^\+628\d{7,12}$/.test(normalizedPhone)) {
    return responseJSON({ status: 'error', message: 'Nomor telepon tidak valid. Gunakan nomor Indonesia, contoh 081234567890.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const rows = sheet.getDataRange().getValues();
  let existingUnit = '';
  let existingPhone = '';
  const takenUsernames = {};

  for (let i = 1; i < rows.length; i++) {
    const role = cell_(rows[i][2]).trim();
    const rowUnit = cell_(rows[i][3]).trim();
    const rowUser = cell_(rows[i][0]).trim();
    const rowPhone = normalizePhone_(rowUser) || normalizePhone_(cell_(rows[i][5]));
    if (rowUser) takenUsernames[rowUser.toLowerCase()] = true;
    if (role === 'media' && rowUnit.toLowerCase() === unit.toLowerCase()) {
      existingUnit = rowUnit;
    }
    if (rowPhone && rowPhone === normalizedPhone) {
      existingPhone = rowPhone;
    }
  }

  if (existingUnit) {
    return responseJSON({ status: 'error', message: 'Unit ' + existingUnit + ' sudah terdaftar. Silakan login atau hubungi admin.' });
  }
  if (existingPhone) {
    return responseJSON({ status: 'error', message: 'Nomor ' + existingPhone + ' sudah terdaftar.' });
  }

  let userId = getNextUserId_();
  let username = generateRegistrationUsername_(nama, userId);
  while (takenUsernames[username.toLowerCase()]) {
    userId = getNextUserId_();
    username = generateRegistrationUsername_(nama, userId);
  }

  const password = 'persawi2026';
  sheet.appendRow([username, password, 'media', unit, nama, normalizedPhone]);

  return responseJSON({
    status: 'success',
    message: 'Pendaftaran berhasil',
    user: { username: username, password: password, unit: unit, nama: nama, telepon: normalizedPhone }
  });
}

function handleUpload(data) {
  const auth = data._auth;
  const config = getAppConfig();
  const folderId = getFolderId_(config);
  const maxMb = parseMaxMb_(config.MAX_FILE_MB);
  const base64 = String(data.base64 == null ? '' : data.base64);

  if (!base64) {
    return responseJSON({ status: 'error', message: 'File kosong.' });
  }

  const approxBytes = Math.floor(base64.length * 0.75);
  if (approxBytes > maxMb * 1024 * 1024) {
    return responseJSON({
      status: 'error',
      message: 'Ukuran file melebihi batas ' + maxMb + 'MB.'
    });
  }

  const taskId = String(data.task_id == null ? '' : data.task_id).trim();
  const task = taskId ? getTaskById_(taskId) : null;
  if (taskId && !task) {
    return responseJSON({ status: 'error', message: 'Tugas rundown tidak ditemukan.' });
  }

  const contentType = String(data.mimeType || 'application/octet-stream');
  const filename = sanitizeFilename_(data.filename || ('file_' + Date.now()));
  const unit = auth.un;
  const uploader = auth.n || String(data.uploader == null ? '' : data.uploader).trim();
  const kategori = task ? task.kategori : (String(data.kategori == null ? '' : data.kategori).trim() || 'Candid_Lainnya');
  const keterangan = String(data.keterangan == null ? '' : data.keterangan).trim();
  const batchId = String(data.batch_id == null ? '' : data.batch_id).trim() || ('BATCH-' + Date.now() + '-' + Math.floor(Math.random() * 1000));

  if (!unit) {
    return responseJSON({ status: 'error', message: 'Unit pengguna tidak ditemukan.' });
  }

  const mainFolder = DriveApp.getFolderById(folderId);
  const unitFolder = getOrCreateUnitFolder_(mainFolder, unit);

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, contentType, filename);
  const file = unitFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const driveUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const id = 'MED-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
  const timestamp = nowText_();

  sheet.appendRow([
    id,
    timestamp,
    unit,
    uploader,
    kategori,
    filename,
    contentType.indexOf('video') === 0 ? 'video' : 'image',
    driveUrl,
    fileId,
    keterangan,
    taskId,
    'drive',
    batchId
  ]);

  if (taskId) {
    markTaskComplete_(taskId, unit, uploader, fileId, keterangan);
  }

  return responseJSON({ status: 'success', message: 'Berhasil diunggah!', url: driveUrl });
}

function handleCompleteR2Upload(data) {
  const auth = data._auth;
  const config = getAppConfig();
  const maxMb = parseMaxR2Mb_(config.MAX_R2_MB);
  const key = String(data.key == null ? '' : data.key).trim();
  const publicUrl = String(data.publicUrl == null ? '' : data.publicUrl).trim();
  const mimeType = String(data.mimeType || 'application/octet-stream');
  const size = parseInt(data.size, 10);

  if (!key || !publicUrl) {
    return responseJSON({ status: 'error', message: 'Key atau URL file tidak valid.' });
  }
  if (!/^https:\/\//i.test(publicUrl)) {
    return responseJSON({ status: 'error', message: 'URL file harus HTTPS.' });
  }

  const unitSlug = r2UnitSlug_(auth.un);
  if (!unitSlug || key.indexOf(unitSlug + '/') !== 0) {
    return responseJSON({ status: 'error', message: 'Key file tidak sesuai unit pengguna.' });
  }
  if (!publicUrl.split('?')[0].endsWith('/' + key)) {
    return responseJSON({ status: 'error', message: 'URL file tidak sesuai key.' });
  }
  if (!Number.isFinite(size) || size <= 0) {
    return responseJSON({ status: 'error', message: 'Ukuran file tidak valid.' });
  }
  if (size > maxMb * 1024 * 1024) {
    return responseJSON({ status: 'error', message: 'Ukuran file melebihi batas ' + maxMb + 'MB.' });
  }

  try {
    const head = UrlFetchApp.fetch(publicUrl, {
      method: 'get',
      headers: { Range: 'bytes=0-0' },
      muteHttpExceptions: true,
      followRedirects: true
    });
    const code = head.getResponseCode();
    if (code !== 200 && code !== 206) {
      return responseJSON({ status: 'error', message: 'File tidak dapat diverifikasi di penyimpanan (HTTP ' + code + ').' });
    }
  } catch (e) {
    return responseJSON({ status: 'error', message: 'File tidak dapat diverifikasi di penyimpanan.' });
  }

  const taskId = String(data.task_id == null ? '' : data.task_id).trim();
  const task = taskId ? getTaskById_(taskId) : null;
  if (taskId && !task) {
    return responseJSON({ status: 'error', message: 'Tugas rundown tidak ditemukan.' });
  }

  const unit = auth.un;
  const uploader = auth.n || String(data.uploader == null ? '' : data.uploader).trim();
  const kategori = task ? task.kategori : (String(data.kategori == null ? '' : data.kategori).trim() || 'Candid_Lainnya');
  const keterangan = String(data.keterangan == null ? '' : data.keterangan).trim();
  const filename = sanitizeFilename_(data.filename || key.split('/').pop() || ('file_' + Date.now()));
  const batchId = String(data.batch_id == null ? '' : data.batch_id).trim() || ('BATCH-' + Date.now() + '-' + Math.floor(Math.random() * 1000));

  if (!unit) {
    return responseJSON({ status: 'error', message: 'Unit pengguna tidak ditemukan.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const id = 'MED-' + Date.now() + '-' + Math.floor(Math.random() * 1000);

  sheet.appendRow([
    id,
    nowText_(),
    unit,
    uploader,
    kategori,
    filename,
    mimeType.indexOf('video') === 0 ? 'video' : 'image',
    publicUrl,
    key,
    keterangan,
    taskId,
    'r2',
    batchId
  ]);

  if (taskId) {
    markTaskComplete_(taskId, unit, uploader, key, keterangan);
  }

  return responseJSON({ status: 'success', message: 'Berhasil diunggah!', url: publicUrl });
}

function handleVerifySession(token) {
  const auth = verifyToken_(token, ['media', 'admin']);
  if (!auth) {
    return responseJSON({ status: 'error', message: 'Sesi berakhir atau tidak valid.' });
  }
  return responseJSON({
    status: 'success',
    user: {
      username: auth.u,
      role: auth.r,
      unit: auth.un,
      nama: auth.n
    }
  });
}

function handleGetUploadUrl(data) {
  const auth = data._auth;
  const config = getAppConfig();
  const maxMb = parseMaxR2Mb_(config.MAX_R2_MB);
  const filename = sanitizeFilename_(data.filename || ('file_' + Date.now()));
  const mimeType = String(data.mimeType || 'application/octet-stream');
  const size = parseInt(data.size, 10);

  if (!/^(image|video)\//.test(mimeType)) {
    return responseJSON({ status: 'error', message: 'Tipe file harus gambar atau video.' });
  }
  if (!Number.isFinite(size) || size <= 0) {
    return responseJSON({ status: 'error', message: 'Ukuran file tidak valid.' });
  }
  if (size > maxMb * 1024 * 1024) {
    return responseJSON({ status: 'error', message: 'Ukuran file melebihi batas ' + maxMb + 'MB.' });
  }

  const slug = r2UnitSlug_(auth.un);
  if (!slug) {
    return responseJSON({ status: 'error', message: 'Unit pengguna tidak ditemukan.' });
  }

  const key = r2KeyFor_(auth.un, filename);
  let signed;
  try {
    signed = r2Presign_('PUT', key, 'PutObject');
  } catch (e) {
    return responseJSON({ status: 'error', message: e.message });
  }

  return responseJSON({
    status: 'success',
    uploadUrl: signed.url,
    key: key,
    publicUrl: signed.config.publicBase + '/' + key
  });
}

function handleDeleteR2Object(data) {
  const key = String(data.key == null ? '' : data.key).trim();
  if (!key) {
    return responseJSON({ status: 'error', message: 'Key file tidak valid.' });
  }

  let signed;
  try {
    signed = r2Presign_('DELETE', key, 'DeleteObject');
  } catch (e) {
    return responseJSON({ status: 'error', message: e.message });
  }

  try {
    const res = UrlFetchApp.fetch(signed.url, { method: 'delete', muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      return responseJSON({ status: 'error', message: 'Gagal menghapus file R2 (HTTP ' + code + ').' });
    }
    return responseJSON({ status: 'success', message: 'File R2 dihapus.' });
  } catch (e) {
    return responseJSON({ status: 'error', message: 'Gagal menghapus file R2.' });
  }
}

function getR2Config_() {
  const props = PropertiesService.getScriptProperties();
  const config = {
    accountId: String(props.getProperty('R2_ACCOUNT_ID') || '').trim(),
    accessKeyId: String(props.getProperty('R2_ACCESS_KEY_ID') || '').trim(),
    secretAccessKey: String(props.getProperty('R2_SECRET_ACCESS_KEY') || '').trim(),
    bucket: String(props.getProperty('R2_BUCKET') || '').trim(),
    publicBase: String(props.getProperty('R2_PUBLIC_BASE') || '').trim().replace(/\/+$/, '')
  };

  const missing = [];
  if (!config.accountId) missing.push('R2_ACCOUNT_ID');
  if (!config.accessKeyId) missing.push('R2_ACCESS_KEY_ID');
  if (!config.secretAccessKey) missing.push('R2_SECRET_ACCESS_KEY');
  if (!config.bucket) missing.push('R2_BUCKET');
  if (!config.publicBase) missing.push('R2_PUBLIC_BASE');
  if (missing.length) {
    throw new Error('Script Property R2 belum lengkap: ' + missing.join(', '));
  }
  return config;
}

function r2KeyFor_(unit, filename) {
  const now = new Date();
  const pad = function (value) { return ('0' + value).slice(-2); };
  const stamp = now.getUTCFullYear() + '/' + pad(now.getUTCMonth() + 1) + '/' + pad(now.getUTCDate());
  const random = Math.random().toString(36).slice(2, 10);
  return r2UnitSlug_(unit) + '/' + stamp + '/' + Date.now() + '-' + random + '-' + sanitizeFilename_(filename);
}

function r2Presign_(method, key, xid) {
  const config = getR2Config_();
  const host = config.accountId + '.r2.cloudflarestorage.com';
  const now = new Date();
  const pad = function (value) { return ('0' + value).slice(-2); };
  const amzDate = now.getUTCFullYear() + pad(now.getUTCMonth() + 1) + pad(now.getUTCDate()) +
    'T' + pad(now.getUTCHours()) + pad(now.getUTCMinutes()) + pad(now.getUTCSeconds()) + 'Z';
  const dateStamp = amzDate.substring(0, 8);
  const scope = dateStamp + '/' + R2_REGION + '/' + R2_SERVICE + '/aws4_request';
  const canonicalUri = '/' + config.bucket + '/' + r2EncodeKeyPath_(key);
  const params = {
    'X-Amz-Algorithm': 'AWS4-HMAC-SHA256',
    'X-Amz-Content-Sha256': 'UNSIGNED-PAYLOAD',
    'X-Amz-Credential': config.accessKeyId + '/' + scope,
    'X-Amz-Date': amzDate,
    'X-Amz-Expires': String(R2_UPLOAD_EXPIRY_SECONDS),
    'X-Amz-SignedHeaders': 'host',
    'x-id': xid
  };

  const canonicalQuery = Object.keys(params).sort().map(function (name) {
    return rfc3986_(name) + '=' + rfc3986_(params[name]);
  }).join('&');
  const canonicalHeaders = 'host:' + host + '\n';
  const canonicalRequest = [method, canonicalUri, canonicalQuery, canonicalHeaders, 'host', 'UNSIGNED-PAYLOAD'].join('\n');
  const stringToSign = ['AWS4-HMAC-SHA256', amzDate, scope, sha256Hex_(canonicalRequest)].join('\n');

  const kDate = hmacSha256Bytes_(utf8Bytes_('AWS4' + config.secretAccessKey), utf8Bytes_(dateStamp));
  const kRegion = hmacSha256Bytes_(kDate, utf8Bytes_(R2_REGION));
  const kService = hmacSha256Bytes_(kRegion, utf8Bytes_(R2_SERVICE));
  const kSigning = hmacSha256Bytes_(kService, utf8Bytes_('aws4_request'));
  const signature = bytesToHex_(hmacSha256Bytes_(kSigning, utf8Bytes_(stringToSign)));

  return {
    url: 'https://' + host + canonicalUri + '?' + canonicalQuery + '&X-Amz-Signature=' + signature,
    config: config
  };
}

function hmacSha256Bytes_(keyBytes, valueBytes) {
  let key = unsignedBytes_(keyBytes);
  if (key.length > 64) key = sha256Bytes_(key);

  const padded = [];
  for (let i = 0; i < 64; i++) padded.push(i < key.length ? key[i] : 0);

  const inner = sha256Bytes_(concatBytes_(xorBytes_(padded, 0x36), valueBytes));
  const outer = sha256Bytes_(concatBytes_(xorBytes_(padded, 0x5c), inner));
  return outer;
}

function sha256Bytes_(bytes) {
  const k = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ];

  const msg = unsignedBytes_(bytes).slice(0);
  const bitLength = msg.length * 8;
  msg.push(0x80);
  while (msg.length % 64 !== 56) msg.push(0);

  const high = Math.floor(bitLength / 0x100000000);
  const low = bitLength >>> 0;
  msg.push((high >>> 24) & 0xff, (high >>> 16) & 0xff, (high >>> 8) & 0xff, high & 0xff);
  msg.push((low >>> 24) & 0xff, (low >>> 16) & 0xff, (low >>> 8) & 0xff, low & 0xff);

  const h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const w = new Array(64);

  for (let offset = 0; offset < msg.length; offset += 64) {
    for (let i = 0; i < 16; i++) {
      const j = offset + i * 4;
      w[i] = (msg[j] << 24) | (msg[j + 1] << 16) | (msg[j + 2] << 8) | msg[j + 3];
    }

    for (let i = 16; i < 64; i++) {
      const s0 = rotr_(w[i - 15], 7) ^ rotr_(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr_(w[i - 2], 17) ^ rotr_(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
    }

    let a = h[0];
    let b = h[1];
    let c = h[2];
    let d = h[3];
    let e = h[4];
    let f = h[5];
    let g = h[6];
    let hh = h[7];

    for (let i = 0; i < 64; i++) {
      const s1 = rotr_(e, 6) ^ rotr_(e, 11) ^ rotr_(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (hh + s1 + ch + k[i] + w[i]) | 0;
      const s0 = rotr_(a, 2) ^ rotr_(a, 13) ^ rotr_(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (s0 + maj) | 0;

      hh = g;
      g = f;
      f = e;
      e = (d + temp1) | 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) | 0;
    }

    h[0] = (h[0] + a) | 0;
    h[1] = (h[1] + b) | 0;
    h[2] = (h[2] + c) | 0;
    h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0;
    h[5] = (h[5] + f) | 0;
    h[6] = (h[6] + g) | 0;
    h[7] = (h[7] + hh) | 0;
  }

  return wordsToBytes_(h);
}

function rotr_(value, shift) {
  return (value >>> shift) | (value << (32 - shift));
}

function wordsToBytes_(words) {
  const bytes = [];
  for (let i = 0; i < words.length; i++) {
    const value = words[i] >>> 0;
    bytes.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
  }
  return bytes;
}

function sha256Hex_(value) {
  return bytesToHex_(sha256Bytes_(utf8Bytes_(value)));
}

function unsignedBytes_(bytes) {
  return bytes.map(function (value) { return (value + 256) % 256; });
}

function xorBytes_(bytes, mask) {
  return bytes.map(function (value) { return ((value + 256) % 256) ^ mask; });
}

function concatBytes_(first, second) {
  const result = first.slice(0);
  for (let i = 0; i < second.length; i++) result.push((second[i] + 256) % 256);
  return result;
}

function utf8Bytes_(value) {
  const str = String(value);
  const bytes = [];

  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);

    if (code < 0x80) {
      bytes.push(code);
    } else if (code < 0x800) {
      bytes.push(0xc0 | (code >> 6), 0x80 | (code & 0x3f));
    } else if (code >= 0xd800 && code <= 0xdbff && i + 1 < str.length) {
      const next = str.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        const codePoint = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
        bytes.push(
          0xf0 | (codePoint >> 18),
          0x80 | ((codePoint >> 12) & 0x3f),
          0x80 | ((codePoint >> 6) & 0x3f),
          0x80 | (codePoint & 0x3f)
        );
        i++;
      } else {
        bytes.push(0xef, 0xbf, 0xbd);
      }
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      bytes.push(0xef, 0xbf, 0xbd);
    } else {
      bytes.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 0x3f), 0x80 | (code & 0x3f));
    }
  }

  return bytes;
}

function bytesToHex_(bytes) {
  let result = '';
  for (let i = 0; i < bytes.length; i++) {
    const value = (bytes[i] + 256) % 256;
    result += (value < 16 ? '0' : '') + value.toString(16);
  }
  return result;
}

function rfc3986_(value) {
  return encodeURIComponent(value).replace(/[!'()*]/g, function (char) {
    return '%' + char.charCodeAt(0).toString(16).toUpperCase();
  });
}

function r2EncodeKeyPath_(key) {
  return String(key).split('/').map(rfc3986_).join('/');
}

function handleGetMedia(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const rows = sheet.getDataRange().getValues();
  const filterUnit = params.unit ? String(params.unit).trim() : '';
  const filterKategori = params.kategori ? String(params.kategori).trim() : '';
  const taskMap = getRundownMap_();
  const data = [];

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;

    const taskId = cell_(rows[i][10]);
    const task = taskMap[taskId];
    const item = {
      id: cell_(rows[i][0]),
      timestamp: cell_(rows[i][1]),
      unit: cell_(rows[i][2]),
      uploader: cell_(rows[i][3]),
      kategori: cell_(rows[i][4]),
      filename: cell_(rows[i][5]),
      filetype: cell_(rows[i][6]),
      drive_url: cell_(rows[i][7]),
      file_id: cell_(rows[i][8]),
      keterangan: cell_(rows[i][9]),
      task_id: taskId,
      storage_type: cell_(rows[i][11]).trim() || 'drive',
      batch_id: cell_(rows[i][12]).trim() || ('single-' + cell_(rows[i][0])),
      task_judul: task ? task.judul : '',
      task_schedule: task ? (task.tanggal + ' ' + task.jam) : ''
    };

    if (filterUnit && item.unit !== filterUnit) continue;
    if (filterKategori && item.kategori !== filterKategori) continue;
    data.push(item);
  }

  return responseJSON({ status: 'success', data: data.reverse() });
}

function handleGetRundown() {
  return responseJSON({ status: 'success', data: getRundownTasks_() });
}

function handleAddRundown(data) {
  const tanggal = String(data.tanggal == null ? '' : data.tanggal).trim();
  const jam = String(data.jam == null ? '' : data.jam).trim();
  const judul = String(data.judul == null ? '' : data.judul).trim();
  const kategori = String(data.kategori == null ? '' : data.kategori).trim() || 'Candid_Lainnya';
  const lokasi = String(data.lokasi == null ? '' : data.lokasi).trim();
  const keterangan = String(data.keterangan == null ? '' : data.keterangan).trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) {
    return responseJSON({ status: 'error', message: 'Tanggal harus format YYYY-MM-DD.' });
  }
  if (!/^\d{2}:\d{2}$/.test(jam)) {
    return responseJSON({ status: 'error', message: 'Jam harus format HH:MM.' });
  }
  if (!judul) {
    return responseJSON({ status: 'error', message: 'Judul tugas wajib diisi.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rundown');
  const id = 'RUN-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
  const createdAt = nowText_();

  sheet.appendRow([id, tanggal, jam, judul, kategori, lokasi, keterangan, createdAt]);
  return responseJSON({
    status: 'success',
    message: 'Rundown berhasil ditambahkan',
    data: { id: id, tanggal: tanggal, jam: jam, judul: judul, kategori: kategori, lokasi: lokasi, keterangan: keterangan }
  });
}

function handleDeleteRundown(id) {
  const targetId = String(id == null ? '' : id).trim();
  if (!targetId) {
    return responseJSON({ status: 'error', message: 'ID rundown tidak valid.' });
  }

  const rundownSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rundown');
  const rundownRows = rundownSheet.getDataRange().getValues();
  let found = false;
  for (let i = rundownRows.length - 1; i >= 1; i--) {
    if (cell_(rundownRows[i][0]) === targetId) {
      rundownSheet.deleteRow(i + 1);
      found = true;
    }
  }
  if (!found) {
    return responseJSON({ status: 'error', message: 'Rundown tidak ditemukan.' });
  }

  const progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskProgress');
  const progressRows = progressSheet.getDataRange().getValues();
  for (let i = progressRows.length - 1; i >= 1; i--) {
    if (cell_(progressRows[i][0]) === targetId) {
      progressSheet.deleteRow(i + 1);
    }
  }

  return responseJSON({ status: 'success', message: 'Rundown dihapus' });
}

function handleGetMyProgress(token) {
  if (!verifyToken_(token, ['media', 'admin'])) {
    return responseJSON({ status: 'error', message: 'Tidak diizinkan.' });
  }

  const auth = verifyToken_(token, ['media', 'admin']);
  const progress = getProgressMapForUnit_(auth.un);
  const tasks = getRundownTasks_().map(task => {
    const row = progress[task.id];
    return Object.assign({}, task, {
      status: row ? row.status : 'Belum',
      completed_at: row ? row.completed_at : '',
      media_id: row ? row.media_id : ''
    });
  });

  const completed = tasks.filter(task => task.status === 'Selesai').length;
  const total = tasks.length;
  const percent = total ? Math.round((completed / total) * 100) : 0;

  return responseJSON({
    status: 'success',
    data: tasks,
    summary: { unit: auth.un, total: total, completed: completed, percent: percent }
  });
}

function handleGetTaskProgress(token) {
  if (!verifyToken_(token, ['admin'])) {
    return responseJSON({ status: 'error', message: 'Tidak diizinkan.' });
  }

  const tasks = getRundownTasks_();
  const units = getMediaUnits_();
  const progressRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskProgress').getDataRange().getValues();
  const completedMap = {};

  for (let i = 1; i < progressRows.length; i++) {
    const taskId = cell_(progressRows[i][0]).trim();
    const unit = cell_(progressRows[i][1]).trim();
    const status = cell_(progressRows[i][2]).trim();
    if (!taskId || status !== 'Selesai') continue;
    if (units.indexOf(unit) === -1) continue;
    if (!completedMap[taskId]) completedMap[taskId] = {};
    completedMap[taskId][unit] = true;
  }

  const data = tasks.map(task => {
    const completedUnits = completedMap[task.id] ? Object.keys(completedMap[task.id]).length : 0;
    const totalUnits = units.length;
    return Object.assign({}, task, {
      completed_units: completedUnits,
      total_units: totalUnits,
      percent: totalUnits ? Math.round((completedUnits / totalUnits) * 100) : 0
    });
  });

  const totalAssignments = tasks.length * units.length;
  const totalCompleted = data.reduce((sum, task) => sum + task.completed_units, 0);
  const percent = totalAssignments ? Math.round((totalCompleted / totalAssignments) * 100) : 0;

  return responseJSON({
    status: 'success',
    data: data,
    summary: {
      total_tasks: tasks.length,
      total_units: units.length,
      total_assignments: totalAssignments,
      total_completed: totalCompleted,
      percent: percent
    }
  });
}

function handleGetUsers(token) {
  if (!verifyToken_(token, ['admin'])) {
    return responseJSON({ status: 'error', message: 'Tidak diizinkan.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const rows = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    users.push({
      username: cell_(rows[i][0]),
      role: cell_(rows[i][2]),
      unit: cell_(rows[i][3]),
      nama: cell_(rows[i][4]),
      telepon: cell_(rows[i][5])
    });
  }
  return responseJSON({ status: 'success', users: users });
}

function handleAddUser(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const username = String(data.username == null ? '' : data.username).trim();
  const password = String(data.password == null ? '' : data.password);
  const role = data.role === 'admin' ? 'admin' : 'media';
  const unit = String(data.unit == null ? '' : data.unit).trim();
  const nama = String(data.nama == null ? '' : data.nama).trim();
  const telepon = String(data.telepon == null ? '' : data.telepon).trim();

  if (!username || !password || !unit || !nama) {
    return responseJSON({ status: 'error', message: 'Username, password, unit, dan nama wajib diisi.' });
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]).toLowerCase() === username.toLowerCase()) {
      return responseJSON({ status: 'error', message: 'Username sudah digunakan.' });
    }
  }

  sheet.appendRow([username, password, role, unit, nama, telepon]);
  return responseJSON({ status: 'success', message: 'User berhasil ditambahkan' });
}

function handleDeleteMedia(id, fileId) {
  const targetId = String(id == null ? '' : id).trim();
  if (!targetId) {
    return responseJSON({ status: 'error', message: 'ID media tidak valid.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const rows = sheet.getDataRange().getValues();
  let targetRow = -1;
  let storedFileId = '';
  let taskId = '';
  let unit = '';
  let storageType = 'drive';

  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]) === targetId) {
      targetRow = i;
      storedFileId = cell_(rows[i][8]);
      taskId = cell_(rows[i][10]).trim();
      unit = cell_(rows[i][2]).trim();
      storageType = cell_(rows[i][11]).trim() || 'drive';
      break;
    }
  }

  if (targetRow === -1) {
    return responseJSON({ status: 'error', message: 'Data media tidak ditemukan.' });
  }

  if (storageType !== 'r2' && fileId && fileId === storedFileId) {
    try {
      DriveApp.getFileById(storedFileId).setTrashed(true);
    } catch (e) {}
  }

  sheet.deleteRow(targetRow + 1);

  if (taskId) {
    recalculateTaskProgress_(taskId, unit);
  }

  return responseJSON({ status: 'success', message: 'Data dihapus' });
}

function handleUpdateConfig(configs) {
  const input = configs || {};
  const updates = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');

  CONFIG_KEYS.forEach(key => {
    if (Object.prototype.hasOwnProperty.call(input, key)) {
      updates[key] = String(input[key] == null ? '' : input[key]).trim();
    }
  });

  if (updates.FOLDER_ID === CONFIG_PLACEHOLDER) {
    return responseJSON({ status: 'error', message: 'FOLDER_ID belum diisi dengan ID folder yang benar.' });
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'MAX_FILE_MB')) {
    const maxMb = parseInt(updates.MAX_FILE_MB, 10);
    if (!maxMb || maxMb < 1 || maxMb > 25) {
      return responseJSON({ status: 'error', message: 'Maksimal ukuran harus antara 1 dan 25 MB.' });
    }
    updates.MAX_FILE_MB = String(maxMb);
  }

  if (Object.prototype.hasOwnProperty.call(updates, 'MAX_R2_MB')) {
    const maxR2Mb = parseInt(updates.MAX_R2_MB, 10);
    if (!maxR2Mb || maxR2Mb < 1 || maxR2Mb > 5120) {
      return responseJSON({ status: 'error', message: 'Maksimal ukuran R2 harus antara 1 dan 5120 MB.' });
    }
    updates.MAX_R2_MB = String(maxR2Mb);
  }

  Object.keys(updates).forEach(key => setConfigValue_(sheet, key, updates[key]));
  return responseJSON({ status: 'success', message: 'Konfigurasi diperbarui' });
}

function handleChangePassword(data) {
  const auth = data._auth;
  const oldPassword = String(data.oldPassword == null ? '' : data.oldPassword);
  const newPassword = String(data.newPassword == null ? '' : data.newPassword);

  if (newPassword.length < 6) {
    return responseJSON({ status: 'error', message: 'Password baru minimal 6 karakter.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const rows = sheet.getDataRange().getValues();
  const targetUser = String(auth.u || '').trim();

  for (let i = 1; i < rows.length; i++) {
    const rowUser = cell_(rows[i][0]).trim();
    if (rowUser !== targetUser) continue;
    if (cell_(rows[i][1]) !== oldPassword) {
      return responseJSON({ status: 'error', message: 'Password lama salah.' });
    }
    sheet.getRange(i + 1, 2).setValue(newPassword);
    return responseJSON({ status: 'success', message: 'Password berhasil diubah.' });
  }

  return responseJSON({ status: 'error', message: 'User tidak ditemukan.' });
}

function getRundownTasks_() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rundown').getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    tasks.push({
      id: cell_(rows[i][0]),
      tanggal: cell_(rows[i][1]),
      jam: cell_(rows[i][2]),
      judul: cell_(rows[i][3]),
      kategori: cell_(rows[i][4]),
      lokasi: cell_(rows[i][5]),
      keterangan: cell_(rows[i][6]),
      created_at: cell_(rows[i][7])
    });
  }
  tasks.sort((a, b) => (a.tanggal + ' ' + a.jam).localeCompare(b.tanggal + ' ' + b.jam));
  return tasks;
}

function getRundownMap_() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rundown').getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    map[cell_(rows[i][0])] = {
      id: cell_(rows[i][0]),
      tanggal: cell_(rows[i][1]),
      jam: cell_(rows[i][2]),
      judul: cell_(rows[i][3]),
      kategori: cell_(rows[i][4]),
      lokasi: cell_(rows[i][5]),
      keterangan: cell_(rows[i][6])
    };
  }
  return map;
}

function getTaskById_(id) {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rundown').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]) === id) {
      return {
        id: cell_(rows[i][0]),
        tanggal: cell_(rows[i][1]),
        jam: cell_(rows[i][2]),
        judul: cell_(rows[i][3]),
        kategori: cell_(rows[i][4]),
        lokasi: cell_(rows[i][5]),
        keterangan: cell_(rows[i][6])
      };
    }
  }
  return null;
}

function getProgressMapForUnit_(unit) {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskProgress').getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][1]).trim() !== unit) continue;
    const taskId = cell_(rows[i][0]).trim();
    if (!taskId) continue;
    map[taskId] = {
      task_id: taskId,
      unit: cell_(rows[i][1]).trim(),
      status: cell_(rows[i][2]),
      completed_at: cell_(rows[i][3]),
      completed_by: cell_(rows[i][4]),
      media_id: cell_(rows[i][5]),
      keterangan: cell_(rows[i][6])
    };
  }
  return map;
}

function markTaskComplete_(taskId, unit, uploader, mediaId, keterangan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskProgress');
  const rows = sheet.getDataRange().getValues();
  const completedAt = nowText_();

  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]) === taskId && cell_(rows[i][1]) === unit) {
      sheet.getRange(i + 1, 3).setValue('Selesai');
      sheet.getRange(i + 1, 4).setValue(completedAt);
      sheet.getRange(i + 1, 5).setValue(uploader);
      sheet.getRange(i + 1, 6).setValue(mediaId);
      sheet.getRange(i + 1, 7).setValue(keterangan);
      return;
    }
  }

  sheet.appendRow([taskId, unit, 'Selesai', completedAt, uploader, mediaId, keterangan]);
}

function recalculateTaskProgress_(taskId, unit) {
  const mediaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const mediaRows = mediaSheet.getDataRange().getValues();
  let latestMedia = null;

  for (let i = 1; i < mediaRows.length; i++) {
    if (cell_(mediaRows[i][10]).trim() !== taskId || cell_(mediaRows[i][2]).trim() !== unit) continue;
    if (cell_(mediaRows[i][8])) {
      latestMedia = {
        id: cell_(mediaRows[i][0]),
        fileId: cell_(mediaRows[i][8]),
        uploader: cell_(mediaRows[i][3]),
        keterangan: cell_(mediaRows[i][9])
      };
    }
  }

  const progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TaskProgress');
  const progressRows = progressSheet.getDataRange().getValues();

  for (let i = progressRows.length - 1; i >= 1; i--) {
    if (cell_(progressRows[i][0]).trim() !== taskId || cell_(progressRows[i][1]).trim() !== unit) continue;
    if (latestMedia) {
      progressSheet.getRange(i + 1, 6).setValue(latestMedia.fileId);
      progressSheet.getRange(i + 1, 5).setValue(latestMedia.uploader);
      progressSheet.getRange(i + 1, 7).setValue(latestMedia.keterangan);
    } else {
      progressSheet.deleteRow(i + 1);
    }
  }
}

function getMediaUnits_() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues();
  const units = {};
  for (let i = 1; i < rows.length; i++) {
    const role = cell_(rows[i][2]).trim();
    const unit = cell_(rows[i][3]).trim();
    if (role === 'media' && unit) units[unit] = true;
  }
  return Object.keys(units);
}

function getAuthSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('AUTH_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('AUTH_SECRET', secret);
  }
  return secret;
}

function createToken_(user) {
  const payload = {
    u: user.username,
    r: user.role,
    un: user.unit,
    n: user.nama,
    exp: Date.now() + (SESSION_HOURS * 60 * 60 * 1000)
  };
  const body = base64UrlEncode_(JSON.stringify(payload));
  const signature = base64UrlEncode_(Utilities.computeHmacSha256Signature(getAuthSecret_(), body));
  return body + '.' + signature;
}

function verifyToken_(token, roles) {
  if (!token || typeof token !== 'string') return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;

  const expected = base64UrlEncode_(Utilities.computeHmacSha256Signature(getAuthSecret_(), parts[0]));
  if (expected !== parts[1]) return null;

  try {
    const payload = JSON.parse(base64UrlDecode_(parts[0]));
    if (!payload.exp || payload.exp < Date.now()) return null;
    if (roles && roles.indexOf(payload.r) === -1) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

function base64UrlEncode_(value) {
  return Utilities.base64Encode(value)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function base64UrlDecode_(value) {
  let base64 = value.replace(/-/g, '+').replace(/_/g, '/');
  while (base64.length % 4) base64 += '=';
  return Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString();
}

function getFolderId_(config) {
  const folderId = String((config && config.FOLDER_ID) || '').trim();
  if (!folderId || folderId === CONFIG_PLACEHOLDER) {
    throw new Error('FOLDER_ID belum diatur di sheet Config.');
  }
  return folderId;
}

function parseMaxMb_(value) {
  const maxMb = parseInt(value, 10);
  if (!Number.isFinite(maxMb) || maxMb < 1) return DEFAULT_MAX_MB;
  return Math.min(maxMb, DEFAULT_MAX_MB);
}

function parseMaxR2Mb_(value) {
  const maxMb = parseInt(value, 10);
  if (!Number.isFinite(maxMb) || maxMb < 1) return DEFAULT_MAX_R2_MB;
  return Math.min(maxMb, 5120);
}

function r2UnitSlug_(unit) {
  return String(unit || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 60);
}

function getOrCreateUnitFolder_(mainFolder, unit) {
  const folders = mainFolder.getFoldersByName(unit);
  if (folders.hasNext()) return folders.next();
  return mainFolder.createFolder(unit);
}

function normalizePhone_(value) {
  const phone = String(value || '').replace(/[^0-9+]/g, '');
  if (/^\+?62/.test(phone)) {
    return '+' + phone.replace(/^\+/, '');
  }
  if (/^0/.test(phone)) {
    return '+62' + phone.slice(1);
  }
  return '';
}

function getNextUserId_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const config = getAppConfig();
  const current = parseInt(config.USER_COUNTER, 10);
  const next = Number.isFinite(current) && current >= 0 ? current + 1 : 1;
  setConfigValue_(sheet, 'USER_COUNTER', String(next));
  return next;
}

function generateRegistrationUsername_(nama, id) {
  const firstName = String(nama || '').trim().split(/\s+/)[0] || 'USER';
  const slug = firstName.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 20) || 'USER';
  return slug + '_' + id;
}

function sanitizeFilename_(name) {
  const safe = String(name || 'file')
    .replace(/[\\/:*?"<>|\r\n]+/g, '_')
    .slice(0, 120)
    .trim();
  return safe || 'file';
}

function setConfigValue_(sheet, key, value) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] == null ? '' : rows[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function nowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function cell_(value) {
  return value == null ? '' : String(value);
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}