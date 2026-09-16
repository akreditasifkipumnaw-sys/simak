/**************************************************************
 * SIMAK — WEBSITE PENDUKUNG AKREDITASI PRODI FKIP
 * Arsitektur: 1 project Apps Script = WEBSITE + API + DATABASE
 *
 * FILE 1 (file ini) : Code.gs     → server / API
 * FILE 2            : index.html  → tampilan website
 *
 * PERUBAHAN versi ini:
 *   1. Password akses data dukung kini = PASSWORD USER yang
 *      memiliki akses ke sheet/prodi TERKAIT (bukan hanya admin).
 *      Password ADMIN tetap sah untuk semua sheet.
 *      Token akses berlaku PER-SHEET (buka PAUD ≠ buka BK).
 *   2. Durasi akses LIHAT data: 8 JAM.
 *      Token disimpan di PropertiesService (CacheService
 *      maksimal 6 jam, tidak cukup untuk 8 jam).
 *
 * DEPLOY (WAJIB setelah edit):
 *   Deploy → Manage deployments → ✏️ → Version: New version → Deploy
 **************************************************************/

const SHEET_ID  = '1adjFxPTccZ-DbbsHeEfAExrFvg09DZg2S0FkHTdH0IA';
const FOLDER_ID = '1ZWn8CGRekdw9cKP_EkWLxd1g5-m5bO9X';

const PRODI_SHEETS = ['FKIP','BK','PBSI','PBIG','PGSD','PAUD','PPKN','DIKMAT','DIKFIS','DIKO','PPG'];
const ALL_SHEETS   = ['DASBORD'].concat(PRODI_SHEETS);
const TOKEN_TTL    = 21600; // sesi login user (± 6 jam, batas CacheService)
const VIEW_TTL     = 28800; // ★ akses LIHAT data dukung: 8 JAM

/* Pemetaan nama sheet lama → baru (untuk hak akses & inisial lama) */
const SHEET_ALIAS_GS = { 'PG-PAUD': 'PAUD', 'FAKULTAS': 'FKIP', 'PGIG': 'PBIG' };

/* ============ ROUTER: WEBSITE + API ============ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action || p.callback) return handle(p, {});

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SIMAK — Monitoring Akreditasi FKIP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // agar bisa di-iframe (Blogger dll)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handle((e && e.parameter) || {}, body);
}

function handle(params, body) {
  try {
    const action = String(params.action || body.action || '').toUpperCase();
    let result;
    switch (action) {
      case 'PING':        result = { success: true, message: 'API aktif' }; break;
      case 'LOGIN':       result = apiLogin(body); break;
      case 'VIEWAUTH':    result = apiVerifyViewPassword(body); break;   /* ★ password user per sheet */
      case 'READ':        result = apiRead(params.sheet || body.sheet); break;
      case 'SAVE':        result = apiSave(body); break;
      case 'ADDROW':      result = apiAddRow(body); break;
      case 'DELETEROW':   result = apiDeleteRow(body); break;
      case 'REPLACE':     result = apiReplace(body); break;
      case 'UPLOAD':      result = apiUpload(body); break;
      case 'USERS':       result = apiListUsers(params.token ? params : body); break;
      case 'USERSAVE':    result = apiUserSave(body); break;
      case 'USERDELETE':  result = apiUserDelete(body); break;
      default:            result = { success: false, message: 'Aksi tidak dikenal: ' + action };
    }
    return respond(result, params.callback);
  } catch (err) {
    return respond({ success: false, message: 'Kesalahan server: ' + err }, params.callback);
  }
}

function respond(result, callback) {
  const json = JSON.stringify(result);
  if (callback) {
    const cb = String(callback).replace(/[^a-zA-Z0-9_.]/g, '');
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================= UTILITAS ======================= */
function sha256(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function getSS() { return SpreadsheetApp.openById(SHEET_ID); }

/* Konversi nilai sel → teks aman (agar Date/number tidak bermasalah di client) */
function plain(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return v;
}

/* Normalisasi daftar sheet: map nama lama → baru, buang duplikat/invalid */
function normalizeSheets(arr) {
  const out = [];
  (arr || []).forEach(function (s) {
    const k = String(s).trim().toUpperCase();
    const mapped = SHEET_ALIAS_GS[k] || k;
    if (PRODI_SHEETS.indexOf(mapped) > -1 && out.indexOf(mapped) === -1) out.push(mapped);
  });
  return out;
}

/* Sheet USERS dibuat otomatis (tersembunyi) saat login pertama kali.
   Akun default: admin / admin123 — SEGERA ganti passwordnya! */
function usersSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName('USERS');
  if (!sh) {
    sh = ss.insertSheet('USERS');
    sh.getRange(1, 1, 1, 5).setValues([['USERNAME','NAMA','ROLE','SHEETS','PASSHASH']]).setFontWeight('bold');
    sh.appendRow(['admin', 'Administrator', 'ADMIN', PRODI_SHEETS.join(','), sha256('admin123')]);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

/* ====== TOKEN LOGIN (CacheService) ====== */
function auth(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('tk_' + String(token));
  if (!raw) return null;
  try {
    const u = JSON.parse(raw);
    if (u && u.sheets) u.sheets = normalizeSheets(u.sheets); // map nama lama → baru
    return u;
  } catch (e) { return null; }
}

function makeToken(user) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('tk_' + token, JSON.stringify(user), TOKEN_TTL);
  return token;
}

/* ★★★ TOKEN AKSES LIHAT DATA (PropertiesService, 8 JAM, PER-SHEET)
   Disimpan sebagai map { "vt_<uuid>": { sheets:[...], exp:<ms> } }
   agar melewati batas 6 jam CacheService. */
function viewTokenMap() {
  const props = PropertiesService.getScriptProperties();
  try { return JSON.parse(props.getProperty('VIEW_TOKENS') || '{}'); }
  catch (e) { return {}; }
}
function viewTokenSave(map) {
  PropertiesService.getScriptProperties().setProperty('VIEW_TOKENS', JSON.stringify(map));
}
function purgeExpiredViewTokens(map) {
  const now = Date.now();
  let changed = false;
  Object.keys(map).forEach(function (k) {
    if (!map[k] || !map[k].exp || map[k].exp <= now) { delete map[k]; changed = true; }
  });
  if (changed) viewTokenSave(map);
}

function makeViewToken(sheet) {
  const token = Utilities.getUuid();
  const map = viewTokenMap();
  purgeExpiredViewTokens(map);
  map['vt_' + token] = { sheets: [String(sheet).toUpperCase()], exp: Date.now() + VIEW_TTL * 1000 };
  viewTokenSave(map);
  return token;
}

/* Valid: token ada, belum kedaluwarsa, DAN mencakup sheet yang diminta */
function authView(token, sheet) {
  if (!token || !sheet) return false;
  const map = viewTokenMap();
  const o = map['vt_' + String(token)];
  if (!o) return false;
  if (Date.now() >= (o.exp || 0)) { purgeExpiredViewTokens(map); return false; }
  const target = String(sheet).toUpperCase();
  return (o.sheets || []).map(function (s) { return String(s).toUpperCase(); }).indexOf(target) > -1;
}

function checkAccess(info, sheet, adminOnly) {
  if (!info) return 'Sesi berakhir. Silakan login kembali.';
  if (adminOnly && info.role !== 'ADMIN') return 'Aksi ini hanya untuk Administrator.';
  if (info.role === 'ADMIN') return null;
  const allowed = (info.sheets || []).map(function (s) { return String(s).toUpperCase(); });
  if (allowed.indexOf(String(sheet).toUpperCase()) === -1)
    return 'Anda tidak memiliki akses ke sheet ' + sheet + '.';
  return null;
}

function sheetGuard(sheet) {
  return ALL_SHEETS.indexOf(String(sheet)) === -1 ? 'Sheet tidak valid: ' + sheet : null;
}

function normalizeRow(vals, len) {
  const out = [];
  for (let i = 0; i < len; i++) out.push(vals && vals[i] !== undefined && vals[i] !== null ? vals[i] : '');
  return out;
}

function sanitizeFilename(name) {
  return String(name || 'bukti').replace(/[\\\/:*?"<>|]/g, '-').slice(0, 150);
}

/* File upload disimpan ke subfolder per-sheet/prodi dalam FOLDER_ID */
function targetFolder(sheetName) {
  const parent = DriveApp.getFolderById(FOLDER_ID);
  const name = sanitizeFilename(sheetName || 'UMUM');
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ======================= API ======================= */
function apiLogin(body) {
  const username = String((body && body.username) || '').trim().toLowerCase();
  const password = String((body && body.password) || '');
  if (!username || !password) return { success: false, message: 'Username dan password wajib diisi.' };
  const rows = usersSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === username && String(rows[i][4]) === sha256(password)) {
      const user = {
        username: rows[i][0],
        nama: rows[i][1],
        role: String(rows[i][2]).toUpperCase(),
        sheets: normalizeSheets(String(rows[i][3]).split(',')) /* auto-map PG-PAUD→PAUD, FAKULTAS→FKIP */
      };
      return { success: true, token: makeToken(user), user: user };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

/* ★★★ VERIFIKASI PASSWORD AKSES DATA DUKUNG ★★★
   Menerima { sheet, password }.
   Dicek terhadap SEMUA akun: jika password cocok dengan akun yang
   berhak atas sheet tsb (USER pemilik sheet / ADMIN) → berikan
   viewToken yang HANYA berlaku untuk sheet itu, selama 8 JAM. */
function apiVerifyViewPassword(body) {
  const sheetRaw = String((body && body.sheet) || '').trim().toUpperCase();
  const password = String((body && body.password) || '');
  if (!sheetRaw) return { success: false, message: 'Sheet tujuan tidak diketahui.' };
  if (!password) return { success: false, message: 'Password wajib diisi.' };
  const target = SHEET_ALIAS_GS[sheetRaw] || sheetRaw;
  if (PRODI_SHEETS.indexOf(target) === -1) return { success: false, message: 'Sheet tidak valid: ' + target };

  const rows = usersSheet().getDataRange().getValues();
  let passwordMatched = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]) === sha256(password)) {
      passwordMatched = true;
      const role = String(rows[i][2]).toUpperCase();
      const sheets = normalizeSheets(String(rows[i][3]).split(','));
      if (role === 'ADMIN' || sheets.indexOf(target) > -1) {
        return {
          success: true,
          viewToken: makeViewToken(target),
          sheet: target,
          expiresIn: VIEW_TTL,
          nama: rows[i][1],
          message: 'Akses data dukung diberikan.'
        };
      }
    }
  }
  return { success: false, message: passwordMatched
    ? 'Password ini tidak memiliki akses ke ' + target + '. Gunakan password user ' + target + '.'
    : 'Password salah.' };
}

/* Menerima string ("DASBORD") atau objek ({sheet, token, viewToken}) */
function apiRead(arg) {
  const obj = (arg && typeof arg === 'object') ? arg : { sheet: arg };
  const sheet = obj.sheet;
  const err = sheetGuard(sheet);
  if (err) return { success: false, message: err };

  /* Kontrol akses baca:
     - DASBORD = publik (dashboard umum)
     - sheet lain = token login user yang berhak ATAU viewToken
       (password user pemilik sheet / admin) yang mencakup sheet ini */
  if (String(sheet).toUpperCase() !== 'DASBORD') {
    let allowed = authView(obj.viewToken, sheet);
    if (!allowed) {
      const info = auth(obj.token);
      if (info) {
        if (info.role === 'ADMIN') allowed = true;
        else {
          const mine = (info.sheets || []).map(function (s) { return String(s).toUpperCase(); });
          allowed = mine.indexOf(String(sheet).toUpperCase()) > -1;
        }
      }
    }
    if (!allowed) return { success: false, message: 'Akses data dukung terkunci. Masukkan password user ' + sheet + ' untuk melihat data.' };
  }

  const sh = getSS().getSheetByName(sheet);
  if (!sh) return { success: false, message: 'Sheet "' + sheet + '" tidak ditemukan.' };
  if (sh.getLastRow() === 0) return { success: true, sheet: sheet, headers: [], rows: [] };
  const values = sh.getDataRange().getValues();
  const rows = values.slice(1).map(function (r) { return r.map(plain); });
  return { success: true, sheet: sheet, headers: values[0].map(String), rows: rows };
}

function apiSave(body) {
  const info = auth(body.token);
  const err = sheetGuard(body.sheet) || checkAccess(info, body.sheet, false);
  if (err) return { success: false, message: err };
  const sh = getSS().getSheetByName(body.sheet);
  const idx = parseInt(body.rowIndex, 10);
  if (!sh || isNaN(idx) || idx < 0 || idx + 2 > sh.getLastRow())
    return { success: false, message: 'Baris tidak ditemukan (data mungkin sudah berubah, muat ulang).' };
  const width = Math.max(sh.getLastColumn(), (body.values || []).length);
  sh.getRange(idx + 2, 1, 1, width).setValues([normalizeRow(body.values, width)]);
  return { success: true, message: 'Perubahan berhasil disimpan.' };
}

function apiAddRow(body) {
  const info = auth(body.token);
  const err = sheetGuard(body.sheet) || checkAccess(info, body.sheet, false);
  if (err) return { success: false, message: err };
  const sh = getSS().getSheetByName(body.sheet);
  if (!sh) return { success: false, message: 'Sheet tidak ditemukan.' };
  sh.appendRow(normalizeRow(body.values, Math.max(sh.getLastColumn(), (body.values || []).length)));
  return { success: true, message: 'Data berhasil ditambahkan.' };
}

function apiDeleteRow(body) {
  const info = auth(body.token);
  const err = sheetGuard(body.sheet) || checkAccess(info, body.sheet, true);
  if (err) return { success: false, message: err };
  const sh = getSS().getSheetByName(body.sheet);
  const idx = parseInt(body.rowIndex, 10);
  if (!sh || isNaN(idx) || idx < 0 || idx + 2 > sh.getLastRow())
    return { success: false, message: 'Baris tidak ditemukan.' };
  sh.deleteRow(idx + 2);
  return { success: true, message: 'Data berhasil dihapus.' };
}

/* Impor Excel: menimpa header + seluruh isi sheet (sinkron penuh) */
function apiReplace(body) {
  const info = auth(body.token);
  const err = sheetGuard(body.sheet) || checkAccess(info, body.sheet, true);
  if (err) return { success: false, message: err };
  const sh = getSS().getSheetByName(body.sheet);
  if (!sh) return { success: false, message: 'Sheet tidak ditemukan.' };
  const headers = body.headers || [];
  const rows = (body.rows || []).filter(function (r) {
    return (r || []).some(function (c) { return String(c).trim() !== ''; });
  });
  let width = Math.max(sh.getLastColumn(), headers.length, 1);
  rows.forEach(function (r) { width = Math.max(width, r.length); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  sh.getRange(1, 1, 1, width).setValues([normalizeRow(headers, width)]);
  if (rows.length) {
    const matrix = rows.map(function (r) { return normalizeRow(r, width); });
    sh.getRange(2, 1, matrix.length, width).setValues(matrix);
  }
  return { success: true, message: 'Impor berhasil: ' + rows.length + ' baris tersinkron ke database.' };
}

/* Upload file bukti (base64) → Google Drive → mengembalikan link */
function apiUpload(body) {
  const info = auth(body.token);
  if (!info) return { success: false, message: 'Sesi berakhir. Silakan login kembali.' };
  if (!body.data) return { success: false, message: 'Data file kosong.' };
  const name = sanitizeFilename(body.filename || ('bukti-' + Date.now()));
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.mimeType || 'application/octet-stream', name);
    const file = targetFolder(body.sheet).createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { success: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view', name: name };
  } catch (err) {
    return { success: false, message: 'Gagal mengunggah: ' + err };
  }
}

/* ============ MANAJEMEN PENGGUNA (ADMIN) ============ */
function apiListUsers(arg) {
  const info = auth((arg && arg.token) || '');
  if (!info || info.role !== 'ADMIN') return { success: false, message: 'Akses ditolak.' };
  const rows = usersSheet().getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    users.push({
      username: rows[i][0],
      nama: rows[i][1],
      role: String(rows[i][2]).toUpperCase(),
      sheets: normalizeSheets(String(rows[i][3]).split(',')) /* tampilkan nama sheet baru */
    });
  }
  return { success: true, users: users };
}

function apiUserSave(body) {
  const info = auth(body.token);
  if (!info || info.role !== 'ADMIN') return { success: false, message: 'Akses ditolak.' };
  const username = String(body.username || '').trim().toLowerCase();
  if (!username) return { success: false, message: 'Username wajib diisi.' };
  const role = String(body.role || 'USER').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'USER';
  let sheets = (body.sheets || []).map(String).map(function (s) { return s.trim().toUpperCase(); })
    .map(function (s) { return SHEET_ALIAS_GS[s] || s; })
    .filter(function (s) { return PRODI_SHEETS.indexOf(s) > -1; });
  if (role === 'ADMIN' && sheets.length === 0) sheets = PRODI_SHEETS.slice();
  const sh = usersSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === username) {
      sh.getRange(i + 1, 2, 1, 3).setValues([[body.nama || username, role, sheets.join(',')]]);
      if (body.password) sh.getRange(i + 1, 5).setValue(sha256(body.password));
      return { success: true, message: 'Pengguna berhasil diperbarui.' };
    }
  }
  if (!body.password) return { success: false, message: 'Password wajib diisi untuk pengguna baru.' };
  sh.appendRow([username, body.nama || username, role, sheets.join(','), sha256(body.password)]);
  return { success: true, message: 'Pengguna baru berhasil ditambahkan.' };
}

function apiUserDelete(body) {
  const info = auth(body.token);
  if (!info || info.role !== 'ADMIN') return { success: false, message: 'Akses ditolak.' };
  const username = String(body.username || '').trim().toLowerCase();
  if (username === info.username) return { success: false, message: 'Tidak dapat menghapus akun yang sedang login.' };
  const sh = usersSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === username) {
      sh.deleteRow(i + 1);
      return { success: true, message: 'Pengguna berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Pengguna tidak ditemukan.' };
}
