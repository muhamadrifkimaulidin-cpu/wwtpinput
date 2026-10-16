/**
 * ============================================
 * Auth.gs — Login, Sesi Server-side, Role & Approver
 * Identitas user selalu diambil dari token sesi di server,
 * bukan dari user_id/nama yang dikirim browser.
 * ============================================
 */

// Default approver jika Config 'approver_ids' kosong: Rifki M., Taryana, Hari Nuralam
var DEFAULT_APPROVER_IDS_ = ['USR005', 'USR006', 'USR003'];
var SESSION_TTL_ = 21600; // 6 jam (maksimum CacheService), diperpanjang tiap request
var MAX_LOGIN_FAIL_ = 5;

function loginUser(username, password) {
  try {
    var name = String(username || '').trim().toLowerCase();
    var pwd = String(password || '');
    if (!name || !pwd) return { success: false, message: 'Mohon isi nama dan password.' };

    var cache = CacheService.getScriptCache();
    var failKey = 'loginfail_' + Utilities.base64EncodeWebSafe(name);
    var fails = parseInt(cache.get(failKey) || '0', 10);
    if (fails >= MAX_LOGIN_FAIL_) return { success: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.' };

    var users = getUsers_();
    if (!users) return { success: false, message: 'Sheet Users tidak ditemukan.' };

    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      if (u.nama.toLowerCase() !== name) continue;
      if (!checkPassword_(pwd, u.pwd)) break;
      if (!u.active) return { success: false, message: 'Akun tidak aktif. Hubungi admin.' };
      cache.remove(failKey);
      var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      cache.put('sess_' + token, String(u.id), SESSION_TTL_);
      return { success: true, token: token, user: publicUser_(u) };
    }
    cache.put(failKey, String(fails + 1), 600);
    return { success: false, message: 'Nama atau password salah.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function logoutUser(token) {
  try { if (token) CacheService.getScriptCache().remove('sess_' + token); } catch (e) {}
  return { success: true };
}

// ========== USERS ==========
function getUsers_() {
  var users = propGetJson_('users_cache');
  if (!users) {
    users = loadUsersFromSheet_();
    if (users) propSetJson_('users_cache', users);
  }
  return users;
}

function loadUsersFromSheet_() {
  var sheet = getSheet_('Users');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var data = sheet.getDataRange().getValues();
  var ci = colIndex_(data[0].map(String));
  var c = {
    id: ci.hasOwnProperty('user_id') ? ci.user_id : 0,
    nama: ci.hasOwnProperty('nama') ? ci.nama : 1,
    pwd: ci.hasOwnProperty('password') ? ci.password : 2,
    role: ci.hasOwnProperty('role') ? ci.role : 3,
    active: ci.hasOwnProperty('active') ? ci.active : 4
  };
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][c.nama]).trim();
    if (!nama) continue;
    users.push({
      id: String(data[i][c.id]).trim(), nama: nama, pwd: String(data[i][c.pwd]).trim(),
      role: String(data[i][c.role]).trim().toLowerCase(), active: isActiveValue_(data[i][c.active])
    });
  }
  return users;
}

function isActiveValue_(v) {
  if (v === false || v === '' || v === 0 || v === null) return false;
  return !/^(false|tidak|no|0|nonaktif)$/i.test(String(v).trim());
}

/**
 * Panggil setelah mengubah sheet Users (tambah user, ganti password, nonaktifkan).
 */
function invalidateUsersCache() {
  propDel_('users_cache');
}

// ========== PASSWORD ==========
// Format hash: sha256$<salt>$<base64>. Password plaintext lama tetap diterima
// sampai di-hash lewat menu WWTP Tools › Admin › Amankan Password Users.
function hashPassword_(pwd, salt) {
  salt = salt || Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pwd, Utilities.Charset.UTF_8);
  return 'sha256$' + salt + '$' + Utilities.base64Encode(digest);
}

function checkPassword_(input, stored) {
  if (!stored) return false;
  if (stored.indexOf('sha256$') === 0) {
    var salt = stored.split('$')[1];
    return hashPassword_(input, salt) === stored;
  }
  return input === stored;
}

// ========== ROLE & APPROVER ==========
function getApproverIds_() {
  var cfg = getConfig_('approver_ids');
  if (cfg) {
    var ids = cfg.split(',').map(function(s) { return s.trim(); }).filter(String);
    if (ids.length) return ids;
  }
  return DEFAULT_APPROVER_IDS_;
}

function isApprover_(userId) { return getApproverIds_().indexOf(String(userId)) >= 0; }

function publicUser_(u) {
  return { user_id: u.id, nama: u.nama, role: u.role, isApprover: isApprover_(u.id) };
}

// ========== SESSION GUARD ==========
function requireSession_(token) {
  if (!token || typeof token !== 'string') throw new Error('AUTH');
  var cache = CacheService.getScriptCache();
  var uid = cache.get('sess_' + token);
  if (!uid) throw new Error('AUTH');
  var users = getUsers_() || [];
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].id) !== uid) continue;
    if (!users[i].active) break;
    cache.put('sess_' + token, uid, SESSION_TTL_);
    return publicUser_(users[i]);
  }
  cache.remove('sess_' + token);
  throw new Error('AUTH');
}

/**
 * Pembungkus semua fungsi yang dipanggil dari web app.
 * opts.write    → visitor ditolak
 * opts.approver → hanya approver
 */
function api_(token, opts, fn) {
  try {
    var user = requireSession_(token);
    if (opts.write && user.role === 'visitor') return { success: false, message: 'Akun Visitor hanya bisa melihat data.' };
    if (opts.approver && !user.isApprover) return { success: false, message: 'Hanya approver yang bisa melakukan aksi ini.' };
    return fn(user);
  } catch (e) {
    if (e && e.message === 'AUTH') return { success: false, authError: true, message: 'Sesi berakhir. Silakan login ulang.' };
    Logger.log('api_ error: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ========== MENU ADMIN (Spreadsheet) ==========
function hashUserPasswords() {
  var ui = requireUi_();
  var sheet = getSheet_('Users');
  if (!sheet || sheet.getLastRow() <= 1) { ui.alert('Error', 'Sheet Users kosong / tidak ditemukan.', ui.ButtonSet.OK); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var col = headers.indexOf('password');
  if (col < 0) col = 2;
  var range = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1);
  var vals = range.getValues();
  var n = 0;
  for (var i = 0; i < vals.length; i++) {
    var p = String(vals[i][0]).trim();
    if (p && p.indexOf('sha256$') !== 0) { vals[i][0] = hashPassword_(p); n++; }
  }
  range.setValues(vals);
  invalidateUsersCache();
  ui.alert('Selesai', n + ' password di-hash.\n\nUntuk mengganti password: ketik password baru (plaintext) di sheet Users, lalu jalankan menu ini lagi.', ui.ButtonSet.OK);
}

function refreshCaches() {
  var ui = requireUi_();
  invalidateUsersCache();
  invalidateConfig_();
  ui.alert('Selesai', 'Cache Users & Config dimuat ulang.', ui.ButtonSet.OK);
}
