/**
 * ============================================
 * SETUP.gs — Initial Setup & Sheet Creation
 * WWTP Operations Data Entry System
 * Auto-backup before any sheet reset
 * ============================================
 * Jalankan fungsi setupAll() SEKALI saat pertama kali setup
 */

function setupAll() {
  setupUsersSheet();
  setupDataOpsSheet();
  setupDataQualitySheet();
  setupConfigSheet();
  setupRequestsSheet();
  setupNotificationsSheet();
  SpreadsheetApp.getActive().toast('Setup selesai! Semua sheet di-backup & dibuat ulang.', 'Setup Complete', 5);
}

function backupSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var tz = Session.getScriptTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
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

function setupUsersSheet() {
  mkSheet_('Users', ['user_id','nama','password','role','active'], [
    ['USR001','Dede Rudiyana','operator123','operator',true],
    ['USR002','Muhammad Lukman','operator123','operator',true],
    ['USR003','Hari Nuralam','operator123','operator',true],
    ['USR004','Tia Kurniawan','operator123','operator',true],
    ['USR008','Wisnu Suryana','operator123','operator',true],
    ['USR005','Rifki M.','admin123','admin',true],
    ['USR006','Taryana','admin123','admin',true],
    ['USR007','Visitor','visitor123','visitor',true]
  ]);
}
function setupDataOpsSheet() {
  mkSheet_('DataOps', [
    'record_id','timestamp','tanggal','shift','nama_operator','user_id',
    'chk_mg17_18','chk_clarifier','chk_fe_removal','chk_bak_sirkulasi_besar',
    'clr_bak_sirkulasi_before','clr_bak_sirkulasi_after',
    'clr_sungai_before','clr_sungai_after','clr_mg17_18_before','clr_mg17_18_after',
    'fm_softener_lt3','fm_flotation','fm_mg16','fm_pam','fm_pam_shift',
    'fm_sm1','fm_sm2','fm_sm3','fm_sm4','fm_sm5','fm_sm6',
    'filter_press_batch','bongkar_lumpur',
    'lab_clarifier','lab_mg17_18','lab_bak_sirkulasi','lab_demin','lab_pam',
    'lab_sm1','lab_sm2','lab_sm3','lab_sm4','lab_sm5','lab_sm6','lab_tanggal',
    'chem_hydrofloc_kg','chem_polymer_g','pekerjaan_tambahan','catatan','fm_leaching'
  ]);
}
function setupDataQualitySheet() {
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
function setupConfigSheet() {
  mkSheet_('Config', ['key','value'], [
    ['app_name','WWTP Operations System'],['delete_password','wwtpmonokem'],
    ['records_per_page','15'],['shift1_start','6'],['shift1_end','14'],
    ['shift2_start','14'],['shift2_end','22'],
    ['admin_email','rifki@example.com']
  ]);
}
function setupRequestsSheet() {
  mkSheet_('Requests', ['request_id','timestamp','type','record_id','record_info','requester_name','requester_id','status','edit_data','approver','resolved_at','reason']);
}
function setupNotificationsSheet() {
  mkSheet_('Notifications', ['notif_id','timestamp','type','record_id','tanggal','shift','operator_name','field','before_val','after_val','diff','status','resolved_by','resolved_at','notes']);
}

// ============ NEW MODULES ============

function setupDataProsesClarifierSheet() {
  var waktu = ['07','12','15','19','23','03'];
  var suffixes = ['_rotameter','_flowmeter','_stroke_naoh','_stroke_hidroflok','_stroke_superflok','_ph','_visual_flokulasi','_visual_air','_batch_lumpur'];
  var cols = ['record_id','timestamp','tanggal','shift','nama_operator','user_id',
    'prep_naoh_flake_kg','prep_hidroflok_kg','prep_superflok_g',
    'prep_air_naoh_l','prep_air_hidroflok_l','prep_air_superflok_l',
    'prep_pompa_ok','prep_tangki_ok','prep_keterangan'];
  for (var w = 0; w < waktu.length; w++) {
    for (var s = 0; s < suffixes.length; s++) {
      cols.push('p' + waktu[w] + suffixes[s]);
    }
  }
  cols.push('catatan');
  mkSheet_('DataProsesClarifier', cols);
}

function setupDataProsesSoftenerSheet() {
  mkSheet_('DataProsesSoftener', [
    'record_id','timestamp','tanggal','shift','nama_operator','user_id',
    'regenerasi','garam_kg','catatan'
  ]);
}

function setupDataMaintenanceSheet() {
  mkSheet_('DataMaintenance', [
    'record_id','timestamp','tanggal','shift','nama_operator','user_id',
    'mnt_clr_tangki','mnt_clr_tangki_flokulan','mnt_clr_tangki_floagulan','mnt_clr_tangki_chemical',
    'mnt_clr_pompa_clarifier','mnt_clr_keseluruhan_proses','mnt_clr_pompa_siplai_leaching',
    'mnt_clr_filter_press_1','mnt_clr_filter_press_2','mnt_clr_kolam',
    'mnt_clr_fungsi_proses','mnt_clr_5r_area',
    'mnt_sft_tangki_air_baku','mnt_sft_tangki_sand_filter','mnt_sft_tangki_garam',
    'mnt_sft_tangki_softener','mnt_sft_tangki_hasil','mnt_sft_valve_piping',
    'mnt_sft_fungsi_proses','mnt_sft_5r_area',
    'mnt_catatan'
  ]);
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
  var sheetsToBackup = ['DataOps', 'DataQuality', 'Requests', 'Notifications'];
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
  removeDailyBackupTrigger();
  var backupSS = getOrCreateBackupSS_();
  ScriptApp.newTrigger('dailyAutoBackup').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('Auto-Backup Aktif',
    'Backup harian berhasil diaktifkan.\n\nJadwal: Setiap hari jam 05:00 WIB\nTujuan: ' + backupSS.getName() +
    '\nSheet: DataOps, DataQuality, Requests, Notifications\nRetensi: 30 hari\n\nURL backup:\n' + backupSS.getUrl(),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function removeDailyBackupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyAutoBackup') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function listBackups() {
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
  try {
    var backupSS = getOrCreateBackupSS_();
    var html = '<script>window.open("' + backupSS.getUrl() + '");google.script.host.close();</script>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(200).setHeight(50), 'Membuka backup...');
  } catch (e) { SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK); }
}
