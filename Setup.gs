/**
 * ============================================
 * SETUP.gs — Setup Sheet, Migrasi & Backup
 * WWTP Operations Data Entry System
 * ============================================
 * Semua fungsi setup dijalankan dari menu Spreadsheet:
 *   WWTP Tools › Admin › Siapkan Sheet & Kolom Baru   (aman, tidak menghapus data)
 *   WWTP Tools › Admin › Setup Ulang Semua Sheet       (RESET, backup dulu)
 * Tidak bisa dipanggil dari web app (dicegah oleh requireUi_).
 */

var DEFAULT_CONFIG_ = [
  ['app_name', 'WWTP Operations System'],
  ['records_per_page', '15'],
  ['admin_email', ''],
  ['approver_ids', 'USR005,USR006,USR003'],
  ['pg_unit', 'bar'],
  ['backup_keep', '10']
];

/**
 * Aman dijalankan kapan saja: membuat sheet modul yang belum ada,
 * menambah kolom baru (Requests.module, Notifications.user_id),
 * dan menambah key Config yang belum ada. Tidak menghapus data.
 */
function prepareSheets() {
  var ui = requireUi_();
  var mods = getModules_();
  Object.keys(mods).forEach(function(k) { ensureModuleSheet_(mods[k]); });
  ensureSheet_('Requests', REQ_HEADERS_);
  ensureSheet_('Notifications', NTF_HEADERS_);
  var added = ensureConfigDefaults_();
  invalidateConfig_();
  invalidateUsersCache();
  ui.alert('Selesai', 'Sheet modul & kolom baru siap.' + (added.length ? '\nKey Config ditambahkan: ' + added.join(', ') : ''), ui.ButtonSet.OK);
}

function ensureConfigDefaults_() {
  var sh = ensureSheet_('Config', ['key', 'value']);
  var data = sh.sheet.getDataRange().getValues();
  var keys = data.map(function(r) { return String(r[0]); });
  var added = [];
  DEFAULT_CONFIG_.forEach(function(kv) {
    if (keys.indexOf(kv[0]) < 0) { sh.sheet.appendRow(kv); added.push(kv[0]); }
  });
  return added;
}

function setupAll() {
  var ui = requireUi_();
  var resp = ui.alert('Setup Ulang Semua Sheet',
    'Sheet Users, DataOps, DataQuality, Config, Requests, Notifications akan di-backup lalu DIBUAT ULANG (kosong).\n\nLanjutkan?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var pwds = setupUsersSheet_();
  setupDataOpsSheet_();
  setupDataQualitySheet_();
  setupConfigSheet_();
  mkSheet_('Requests', REQ_HEADERS_);
  mkSheet_('Notifications', NTF_HEADERS_);
  var mods = getModules_();
  Object.keys(mods).forEach(function(k) { ensureModuleSheet_(mods[k]); });
  invalidateUsersCache();
  invalidateConfig_();
  ui.alert('Setup selesai',
    'Password sementara (bagikan secara pribadi, lalu ganti di sheet Users + menu Amankan Password):\n\n' + pwds.join('\n'),
    ui.ButtonSet.OK);
}

function backupSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var ts = fmtDate_(new Date(), 'yyyyMMdd_HHmmss');
  var backupName = 'BACKUP_' + sheetName + '_' + ts;
  var backup = sheet.copyTo(ss);
  backup.setName(backupName);
  ss.setActiveSheet(backup);
  ss.moveActiveSheet(ss.getNumSheets());
  Logger.log('Backup created: ' + backupName + ' (' + (sheet.getLastRow() - 1) + ' records)');
  return true;
}

function mkSheet_(name, headers, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  backupSheet_(name);
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); } else { sheet.clear(); }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (data && data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  return sheet;
}

/** Users dibuat dengan password acak (disimpan sebagai hash). Return daftar 'Nama: password'. */
function setupUsersSheet_() {
  var seed = [
    ['USR001', 'Dede Rudiyana', 'operator'], ['USR002', 'Muhammad Lukman', 'operator'],
    ['USR003', 'Hari Nuralam', 'operator'], ['USR004', 'Tia Kurniawan', 'operator'],
    ['USR008', 'Wisnu Suryana', 'operator'], ['USR005', 'Rifki M.', 'admin'],
    ['USR006', 'Taryana', 'admin'], ['USR007', 'Visitor', 'visitor']
  ];
  var list = [], rows = [];
  seed.forEach(function(u) {
    var pwd = Utilities.getUuid().replace(/-/g, '').substring(0, 10);
    list.push(u[1] + ': ' + pwd);
    rows.push([u[0], u[1], hashPassword_(pwd), u[2], true]);
  });
  mkSheet_('Users', ['user_id', 'nama', 'password', 'role', 'active'], rows);
  return list;
}

function setupDataOpsSheet_() { mkSheet_('DataOps', getModule_('ops').keys); }

function setupDataQualitySheet_() {
  mkSheet_('DataQuality', [
    'qc_id','timestamp','tanggal','nama_operator','user_id',
    'clr_ph','clr_conductivity','clr_tss','clr_tds',
    'sm1_ph','sm1_conductivity','sm1_hardness','sm2_ph','sm2_conductivity','sm2_hardness',
    'sm3_ph','sm3_conductivity','sm3_hardness','sm4_ph','sm4_conductivity','sm4_hardness',
    'sm5_ph','sm5_conductivity','sm5_hardness','sm6_ph','sm6_conductivity','sm6_hardness',
    'pam_ph','pam_conductivity','pam_hardness',
    'bsb_ph','bsb_conductivity','bsb_tss','bsb_tds',
    'bs17_ph','bs17_conductivity','bs17_tss','bs17_tds',
    'sft_ph','sft_conductivity','sft_hardness'
  ]);
}

function setupConfigSheet_() {
  mkSheet_('Config', ['key', 'value'], DEFAULT_CONFIG_.concat([
    ['shift1_start', '6'], ['shift1_end', '14'], ['shift2_start', '14'], ['shift2_end', '22']
  ]));
}

// ============================================
// AUTO-BACKUP (ke Spreadsheet terpisah, hanya admin)
// - Semua sheet di-backup, kecuali sheet sementara (RPT_, BACKUP_, tmp_)
// - Per sheet disimpan N salinan terakhir (Config 'backup_keep', default 10)
// - Backup di hari yang sama menimpa salinan hari itu
// ============================================

var BACKUP_KEEP_DEFAULT_ = 10;
var BACKUP_SKIP_PREFIX_ = ['RPT_', 'BACKUP_', 'tmp_'];

function getBackupKeep_() {
  var n = parseInt(getConfig_('backup_keep'), 10);
  return n > 0 ? n : BACKUP_KEEP_DEFAULT_;
}

function getOrCreateBackupSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var savedId = String(getConfig_('backup_spreadsheet_id') || '').trim();
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) {}
  }
  // Dibuat di My Drive pemilik script (bukan di folder bersama), jadi tidak
  // mewarisi akses operator dari folder.
  var backupSS = SpreadsheetApp.create('WWTP_Backup_' + ss.getName());
  var defaultSheet = backupSS.getSheetByName('Sheet1');
  var cfg = ensureSheet_('Config', ['key', 'value']).sheet;
  var data = cfg.getDataRange().getValues(), found = false;
  for (var j = 0; j < data.length; j++) {
    if (data[j][0] === 'backup_spreadsheet_id') { cfg.getRange(j + 1, 2).setValue(backupSS.getId()); found = true; break; }
  }
  if (!found) cfg.appendRow(['backup_spreadsheet_id', backupSS.getId()]);
  invalidateConfig_();
  var infoSheet = backupSS.insertSheet('_INFO');
  infoSheet.getRange('A1:B1').setValues([['Key', 'Value']]).setFontWeight('bold');
  infoSheet.getRange('A2:B4').setValues([['Source', ss.getName()], ['Source ID', ss.getId()], ['Created', new Date().toISOString()]]);
  if (defaultSheet) { try { backupSS.deleteSheet(defaultSheet); } catch (e) {} }
  return backupSS;
}

/** Matikan link sharing & kembalikan daftar akun selain pemilik yang masih punya akses. */
function secureBackupFile_(backupSS) {
  var others = [];
  try {
    var file = DriveApp.getFileById(backupSS.getId());
    try { file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) {}
    var owner = file.getOwner() ? file.getOwner().getEmail() : '';
    file.getEditors().concat(file.getViewers()).forEach(function(u) {
      var em = u.getEmail();
      if (em && em !== owner && others.indexOf(em) < 0) others.push(em);
    });
  } catch (e) { Logger.log('secureBackupFile_: ' + e.message); }
  return others;
}

function isBackupSource_(name) {
  return !BACKUP_SKIP_PREFIX_.some(function(p) { return name.indexOf(p) === 0; });
}

/** Inti backup. Return ringkasan { copied, failed, deleted, keep, backupSS, others }. */
function runBackup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = fmtDate_(new Date(), 'yyyyMMdd');
  var keep = getBackupKeep_();
  var backupSS = getOrCreateBackupSS_();
  var others = secureBackupFile_(backupSS);
  var copied = [], failed = [];

  ss.getSheets().forEach(function(src) {
    var name = src.getName();
    if (!isBackupSource_(name)) return;
    var bkName = 'BK_' + name + '_' + today;
    try {
      var copy = src.copyTo(backupSS);
      var old = backupSS.getSheetByName(bkName);
      if (old) backupSS.deleteSheet(old); // timpa salinan hari ini setelah salinan baru berhasil
      copy.setName(bkName);
      copied.push(name);
    } catch (e) {
      failed.push(name + ' (' + e.message + ')');
    }
  });

  var deleted = pruneBackups_(backupSS, keep);
  Logger.log('Backup: ' + copied.length + ' sheet, gagal ' + failed.length + ', dihapus ' + deleted + ' → ' + backupSS.getName());
  return { copied: copied, failed: failed, deleted: deleted, keep: keep, backupSS: backupSS, others: others };
}

/** Per sheet sumber, simpan `keep` salinan terbaru; hapus sisanya. */
function pruneBackups_(backupSS, keep) {
  var groups = {};
  backupSS.getSheets().forEach(function(sh) {
    var m = sh.getName().match(/^BK_(.+)_(\d{8})$/);
    if (!m) return;
    (groups[m[1]] = groups[m[1]] || []).push({ sheet: sh, date: m[2] });
  });
  var deleted = 0;
  Object.keys(groups).forEach(function(k) {
    var list = groups[k].sort(function(a, b) { return a.date < b.date ? 1 : -1; });
    list.slice(keep).forEach(function(x) {
      if (backupSS.getSheets().length > 1) { backupSS.deleteSheet(x.sheet); deleted++; }
    });
  });
  return deleted;
}

/** Dipanggil trigger harian (tanpa dialog). */
function dailyAutoBackup() {
  runBackup_();
}

/** Menu: Backup Sekarang (dengan ringkasan). */
function backupNow() {
  var ui = requireUi_();
  var r = runBackup_();
  ui.alert(r.failed.length ? 'Backup selesai dengan error' : 'Backup selesai',
    r.copied.length + ' sheet di-backup ke ' + r.backupSS.getName() + '.\n' +
    'Disimpan ' + r.keep + ' salinan terakhir per sheet; ' + r.deleted + ' salinan lama dihapus.' +
    (r.failed.length ? '\n\nGagal:\n' + r.failed.join('\n') : '') + accessNote_(r.others),
    ui.ButtonSet.OK);
}

function accessNote_(others) {
  return others.length
    ? '\n\n⚠ Akun lain yang masih punya akses ke spreadsheet backup:\n' + others.join('\n') + '\nHapus lewat tombol Share bila bukan admin.'
    : '\n\nAkses: hanya pemilik (link sharing nonaktif).';
}

function setupDailyBackupTrigger() {
  requireUi_();
  removeBackupTriggers_();
  var backupSS = getOrCreateBackupSS_();
  var others = secureBackupFile_(backupSS);
  ScriptApp.newTrigger('dailyAutoBackup').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Auto-Backup Aktif',
    'Backup harian berhasil diaktifkan.\n\nJadwal: Setiap hari jam 05:00 (' + tz_() + ')\nTujuan: ' + backupSS.getName() +
    '\nSheet: semua sheet (kecuali RPT_, BACKUP_, tmp_)\nRetensi: ' + getBackupKeep_() + ' salinan terakhir per sheet\n\nURL backup:\n' + backupSS.getUrl() + accessNote_(others),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function removeDailyBackupTrigger() {
  var ui = requireUi_();
  removeBackupTriggers_();
  ui.alert('Auto-Backup Nonaktif', 'Trigger backup harian dihapus.', ui.ButtonSet.OK);
}

function removeBackupTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyAutoBackup') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function listBackups() {
  requireUi_();
  try {
    var backupSS = getOrCreateBackupSS_();
    var sheets = backupSS.getSheets();
    var backups = [];
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name.indexOf('BK_') === 0) backups.push(name + ' (' + sheets[i].getLastRow() + ' rows)');
    }
    backups.sort();
    SpreadsheetApp.getUi().alert('Daftar Backup (' + backups.length + ')',
      'Lokasi: ' + backupSS.getName() + '\nRetensi: ' + getBackupKeep_() + ' salinan per sheet\n\n' + (backups.length === 0 ? 'Belum ada backup.' : backups.join('\n')) + accessNote_(secureBackupFile_(backupSS)),
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK); }
}

function openBackupSpreadsheet() {
  requireUi_();
  try {
    var backupSS = getOrCreateBackupSS_();
    var html = '<script>window.open("' + backupSS.getUrl() + '");google.script.host.close();</script>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(200).setHeight(50), 'Membuka backup...');
  } catch (e) { SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK); }
}
