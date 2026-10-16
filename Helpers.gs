/**
 * ============================================
 * Helpers.gs — Utilities, Cache, Config, ID, Tanggal
 * ============================================
 */

var CACHE_TTL_ = 3600;
var CACHE_TTL_SHORT_ = 300;
// Kolom yang selalu diisi server, tidak boleh dikirim/diubah dari browser
var SYSTEM_KEYS_ = ['record_id', 'timestamp', 'nama_operator', 'user_id'];

// ========== WAKTU & TANGGAL ==========
function tz_() { return Session.getScriptTimeZone(); }
function fmtDate_(d, pattern) { return Utilities.formatDate(d, tz_(), pattern); }

/**
 * Waktu operasional mengikuti zona waktu script (waktu lokal database).
 * Shift 1 = 06–14, Shift 2 = 14–22, Shift 3 = 22–06.
 * Shift 3 dicatat pada tanggal mulainya, jadi jam 00:00–05:59 masuk
 * Shift 3 tanggal kemarin (konsisten dengan getLastShiftData: Shift 1
 * tanggal D mengambil data Shift 3 tanggal D-1).
 */
function getOperationalNow_() {
  var now = new Date();
  var hour = parseInt(fmtDate_(now, 'H'), 10);
  var shift = (hour >= 6 && hour < 14) ? 1 : (hour >= 14 && hour < 22) ? 2 : 3;
  var opDate = hour < 6 ? new Date(now.getTime() - 86400000) : now;
  return { today: fmtDate_(opDate, 'yyyy-MM-dd'), calendarToday: fmtDate_(now, 'yyyy-MM-dd'), shift: shift };
}

function getServerInfo() {
  var op = getOperationalNow_();
  return {
    today: op.today, calendarToday: op.calendarToday, shift: op.shift,
    timezone: tz_(), serverTime: new Date().getTime(),
    pgUnit: getConfig_('pg_unit') || 'bar'
  };
}

/** Normalisasi nilai tanggal (Date / yyyy-MM-dd / dd/MM/yyyy / serial) → 'yyyy-MM-dd' atau '' */
function toYmd_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : fmtDate_(v, 'yyyy-MM-dd');
  if (typeof v === 'number') {
    if (v > 10000 && v < 100000) return fmtDate_(new Date(Math.round((v - 25569) * 86400000)), 'yyyy-MM-dd');
    return '';
  }
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function addDaysYmd_(ymd, n) {
  var p = ymd.split('-');
  var d = new Date(Date.UTC(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10) + n));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

/** 'yyyy-MM-dd' → Date jam 00:00 (atau 23:59:59.999) waktu lokal script */
function ymdToDate_(ymd, endOfDay) {
  var d = new Date(ymd + 'T00:00:00');
  if (endOfDay) d.setHours(23, 59, 59, 999);
  return d;
}

function isTrue_(v) { return v === true || v === 'true' || v === 'TRUE'; }

// ========== ID & LOCK ==========
/** ID unik: PREFIX-yyyyMMdd-HHmmss-XXXXXX (6 hex acak dari UUID) */
function genId_(prefix) {
  var rand = Utilities.getUuid().replace(/-/g, '').substring(0, 6).toUpperCase();
  return prefix + fmtDate_(new Date(), 'yyyyMMdd-HHmmss') + '-' + rand;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** Hanya bisa dipanggil dari Spreadsheet (menu). Dari web app getUi() melempar error. */
function requireUi_() {
  try { return SpreadsheetApp.getUi(); }
  catch (e) { throw new Error('Fungsi ini hanya bisa dijalankan dari menu Spreadsheet.'); }
}

// ========== CACHE SERVICE (versioned keys) ==========
// Setiap tulis data menaikkan versi namespace → key lama otomatis tidak terpakai.
function cacheGet_(key) {
  try {
    var val = CacheService.getScriptCache().get(key);
    return val === null ? null : JSON.parse(val);
  } catch (e) { return null; }
}
function cachePut_(key, value, ttl) {
  try {
    var json = JSON.stringify(value);
    if (json.length > 95000) return;
    CacheService.getScriptCache().put(key, json, ttl || CACHE_TTL_);
  } catch (e) { Logger.log('cachePut_ error: ' + e.message); }
}
function cacheKey_(parts) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(parts), Utilities.Charset.UTF_8);
  return 'k_' + Utilities.base64EncodeWebSafe(digest);
}
function ver_(ns) { return propGet_('ver_' + ns) || '0'; }
function bumpVer_(ns) { propSet_('ver_' + ns, String(new Date().getTime())); }

// ========== PROPERTIES SERVICE ==========
function propGet_(key) {
  try { return PropertiesService.getScriptProperties().getProperty(key); } catch (e) { return null; }
}
function propSet_(key, value) {
  try { PropertiesService.getScriptProperties().setProperty(key, value); } catch (e) {}
}
function propDel_(key) {
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch (e) {}
}
function propGetJson_(key) {
  var val = propGet_(key);
  if (!val) return null;
  try { return JSON.parse(val); } catch (e) { return null; }
}
function propSetJson_(key, value) {
  try { propSet_(key, JSON.stringify(value)); } catch (e) {}
}

// ========== CONFIG CACHE ==========
function getConfig_(key) {
  var cached = propGet_('cfg_' + key);
  if (cached !== null) return cached;
  var sheet = getSheet_('Config');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') propSet_('cfg_' + data[i][0], String(data[i][1]));
  }
  return propGet_('cfg_' + key);
}

function invalidateConfig_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    for (var key in all) {
      if (key.indexOf('cfg_') === 0) props.deleteProperty(key);
    }
  } catch (e) {}
}

// ========== SHEET HELPERS ==========
function getSheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

/** Ambil sheet; buat jika belum ada; tambahkan kolom yang belum ada di ujung kanan. */
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return { sheet: sheet, headers: headers.slice() };
  }
  return { sheet: sheet, headers: ensureColumns_(sheet, headers) };
}

function ensureColumns_(sheet, cols) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  var missing = cols.filter(function(c) { return headers.indexOf(c) < 0; });
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    headers = headers.concat(missing);
  }
  return headers;
}

function colIndex_(headers) {
  var ci = {};
  for (var i = 0; i < headers.length; i++) if (headers[i]) ci[headers[i]] = i;
  return ci;
}

function rowToObj_(headers, row) {
  var obj = {};
  for (var j = 0; j < headers.length; j++) {
    var k = headers[j];
    if (!k) continue;
    var v = row[j];
    if (/tanggal$/.test(k)) {
      v = (v instanceof Date || (typeof v === 'string' && v)) ? (toYmd_(v) || String(v)) : v;
    } else if (v instanceof Date) {
      v = fmtDate_(v, 'yyyy-MM-dd HH:mm:ss');
    }
    obj[k] = v;
  }
  return obj;
}

function rowFromObj_(headers, data) {
  return headers.map(function(h) { return (h && data.hasOwnProperty(h)) ? data[h] : ''; });
}

/** Cegah teks dieksekusi sebagai formula (=, +, -, @) dan batasi panjang */
function safeCell_(s) {
  s = String(s).substring(0, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function escHtml_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
