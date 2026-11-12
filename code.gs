const CONFIG_PLACEHOLDER = 'TEMPEL_FOLDER_ID_DISINI';
const DEFAULT_MAX_MB = 25;
const SESSION_HOURS = 12;
const CONFIG_KEYS = ['FOLDER_ID', 'EVENT_NAME', 'EVENT_SUBTITLE', 'EVENT_SLOGAN', 'MAX_FILE_MB'];

function setupDatabaseSheets() {
  ensureDatabase_();
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
      ['MAX_FILE_MB', String(DEFAULT_MAX_MB)]
    ].forEach(row => sheetConfig.appendRow(row));
  }

  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['username', 'password', 'role', 'unit', 'nama']);
    sheetUsers.appendRow(['admin', 'admin2026', 'admin', 'Koordinator YPWI', 'Akbar Irwansya']);
    sheetUsers.appendRow(['malili1', 'media2026', 'media', 'SDIT Insan Rabbani Malili', 'Tim Media 1']);
  }

  let sheetMedia = ss.getSheetByName('Media');
  if (!sheetMedia) {
    sheetMedia = ss.insertSheet('Media');
    sheetMedia.appendRow([
      'id', 'timestamp', 'unit', 'uploader', 'kategori',
      'filename', 'filetype', 'drive_url', 'file_id', 'keterangan'
    ]);
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
    if (action === 'getUsers') return handleGetUsers(params.token);
    if (action === 'getConfig') return responseJSON({ status: 'success', config: getAppConfig() });
    return responseJSON({ status: 'ok', message: 'API Media Center Aktif' });
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
      addUser: ['admin'],
      deleteMedia: ['admin'],
      updateConfig: ['admin']
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
    } else if (action === 'upload') {
      return handleUpload(data);
    } else if (action === 'addUser') {
      return handleAddUser(data);
    } else if (action === 'deleteMedia') {
      return handleDeleteMedia(data.id, data.fileId);
    } else if (action === 'updateConfig') {
      return handleUpdateConfig(data.configs);
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

  for (let i = 1; i < rows.length; i++) {
    const rowUser = String(rows[i][0] == null ? '' : rows[i][0]).trim();
    const rowPass = String(rows[i][1] == null ? '' : rows[i][1]);
    if (rowUser === inputUser && rowPass === inputPass) {
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

  const contentType = String(data.mimeType || 'application/octet-stream');
  const filename = sanitizeFilename_(data.filename || ('file_' + Date.now()));
  const unit = auth.un;
  const uploader = auth.n || String(data.uploader == null ? '' : data.uploader).trim();
  const kategori = String(data.kategori == null ? '' : data.kategori).trim();
  const keterangan = String(data.keterangan == null ? '' : data.keterangan).trim();

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
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

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
    keterangan
  ]);

  return responseJSON({ status: 'success', message: 'Berhasil diunggah!', url: driveUrl });
}

function handleGetMedia(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const rows = sheet.getDataRange().getValues();
  const filterUnit = params.unit ? String(params.unit).trim() : '';
  const filterKategori = params.kategori ? String(params.kategori).trim() : '';
  const data = [];

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;

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
      keterangan: cell_(rows[i][9])
    };

    if (filterUnit && item.unit !== filterUnit) continue;
    if (filterKategori && item.kategori !== filterKategori) continue;
    data.push(item);
  }

  return responseJSON({ status: 'success', data: data.reverse() });
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
      nama: cell_(rows[i][4])
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

  if (!username || !password || !unit || !nama) {
    return responseJSON({ status: 'error', message: 'Username, password, unit, dan nama wajib diisi.' });
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]).toLowerCase() === username.toLowerCase()) {
      return responseJSON({ status: 'error', message: 'Username sudah digunakan.' });
    }
  }

  sheet.appendRow([username, password, role, unit, nama]);
  return responseJSON({ status: 'success', message: 'User berhasil ditambahkan' });
}

function handleDeleteMedia(id, fileId) {
  const targetId = String(id == null ? '' : id).trim();
  if (!targetId) {
    return responseJSON({ status: 'error', message: 'ID media tidak valid.' });
  }

  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {}
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Media');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cell_(rows[i][0]) === targetId) {
      sheet.deleteRow(i + 1);
      break;
    }
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

  Object.keys(updates).forEach(key => setConfigValue_(sheet, key, updates[key]));
  return responseJSON({ status: 'success', message: 'Konfigurasi diperbarui' });
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
  return Number.isFinite(maxMb) && maxMb > 0 ? maxMb : DEFAULT_MAX_MB;
}

function getOrCreateUnitFolder_(mainFolder, unit) {
  const folders = mainFolder.getFoldersByName(unit);
  if (folders.hasNext()) return folders.next();
  return mainFolder.createFolder(unit);
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

function cell_(value) {
  return value == null ? '' : String(value);
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}