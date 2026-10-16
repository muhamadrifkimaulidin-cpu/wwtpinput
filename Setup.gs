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
  ['pg_unit', 'bar']
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
// AUTO-BACKUP HARIAN (ke Spreadsheet terpisah)
// ============================================

function getOrCreateBackupSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = ss.getSheetByName('Config');
  if (config && config.getLastRow() > 1) {
    var data = config.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'backup_spreadsheet_id') {
        var savedId = String(data[i][1]).trim();
        if (savedId) {
          try { var existing = SpreadsheetApp.openById(savedId); return existing; } catch (e) {}
        }
      }
    }
  }
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  var folder = folders.hasNext() ? folders.next() : DriveApp.getRootFolder();
  var backupSS = SpreadsheetApp.create('WWTP_Backup_' + ss.getName());
  var backupFile = DriveApp.getFileById(backupSS.getId());
  folder.addFile(backupFile);
  var root = DriveApp.getRootFolder();
  root.removeFile(backupFile);
  var defaultSheet = backupSS.getSheetByName('Sheet1');
  if (config) {
    var found = false;
    if (config.getLastRow() > 1) {
      var configData = config.getDataRange().getValues();
      for (var j = 0; j < configData.length; j++) {
        if (configData[j][0] === 'backup_spreadsheet_id') { config.getRange(j + 1, 2).setValue(backupSS.getId()); found = true; break; }
      }
    }
    if (!found) config.appendRow(['backup_spreadsheet_id', backupSS.getId()]);
  }
  var infoSheet = backupSS.insertSheet('_INFO');
  infoSheet.getRange('A1:B1').setValues([['Key', 'Value']]).setFontWeight('bold');
  infoSheet.getRange('A2:B4').setValues([['Source', ss.getName()],['Source ID', ss.getId()],['Created', new Date().toISOString()]]);
  if (defaultSheet) { try { backupSS.deleteSheet(defaultSheet); } catch (e) {} }
  return backupSS;
}

function dailyAutoBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var sheetsToBackup = ['DataOps', 'DataQuality', 'Requests', 'Notifications', 'DataProsesClarifier', 'DataProsesSoftener', 'DataMaintenance', 'DataPressureGauge'];
  var backupSS = getOrCreateBackupSS_();
  var count = 0;
  sheetsToBackup.forEach(function(name) {
    var src = ss.getSheetByName(name);
    if (!src || src.getLastRow() <= 1) return;
    var bkName = 'BK_' + name + '_' + today;
    if (backupSS.getSheetByName(bkName)) return;
    var copied = src.copyTo(backupSS);
    copied.setName(bkName);
    count++;
  });
  cleanOldBackups_(backupSS, 30);
  Logger.log('Daily backup: ' + count + ' sheets → ' + backupSS.getName() + ' (' + today + ')');
}

function cleanOldBackups_(backupSS, maxDays) {
  var now = new Date();
  var sheets = backupSS.getSheets();
  for (var i = sheets.length - 1; i >= 0; i--) {
    var name = sheets[i].getName();
    if (name.indexOf('BK_') !== 0) continue;
    var dateStr = name.substring(name.length - 8);
    if (!dateStr.match(/^\d{8}$/)) continue;
    var y = parseInt(dateStr.substring(0, 4)), m = parseInt(dateStr.substring(4, 6)) - 1, d = parseInt(dateStr.substring(6, 8));
    var bkDate = new Date(y, m, d);
    var diffDays = Math.floor((now - bkDate) / (1000 * 60 * 60 * 24));
    if (diffDays > maxDays) {
      if (backupSS.getSheets().length > 1) { backupSS.deleteSheet(sheets[i]); Logger.log('Deleted old backup: ' + name); }
    }
  }
}

function setupDailyBackupTrigger() {
  requireUi_();
  removeBackupTriggers_();
  var backupSS = getOrCreateBackupSS_();
  ScriptApp.newTrigger('dailyAutoBackup').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Auto-Backup Aktif',
    'Backup harian berhasil diaktifkan.\n\nJadwal: Setiap hari jam 05:00 WIB\nTujuan: ' + backupSS.getName() +
    '\nSheet: DataOps, DataQuality, Requests, Notifications, FIK, Softener, Maintenance, Pressure Gauge\nRetensi: 30 hari\n\nURL backup:\n' + backupSS.getUrl(),
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
    SpreadsheetApp.getUi().alert('Daftar Backup (' + backups.length + ')',
      'Lokasi: ' + backupSS.getName() + '\n\n' + (backups.length === 0 ? 'Belum ada backup.' : backups.join('\n')),
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
