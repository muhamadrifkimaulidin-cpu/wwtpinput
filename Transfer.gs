/**
 * ============================================
 * Transfer.gs — WWTP Tools v2.2
 * 1. Kirim data kualitas DataKirim → DataQuality, lalu kosongkan
 *    area input DataKirim (pindah, bukan salin)
 * 2. Cut-paste antar cell di DataKirim
 * 3. Hapus data input di DataKirim
 * ============================================
 */

// Area input DataKirim (hanya isi yang dihapus; format & formula tetap)
var DATAKIRIM_INPUT_RANGES_ = ['G10:K12', 'G14:K15', 'G20:K26', 'M10:R49', 'T10:Y19', 'T23:Y31', 'AA10:AE18', 'AA23:AE31'];

function clearDataKirimInputs_(sheet) {
  DATAKIRIM_INPUT_RANGES_.forEach(function(a) { sheet.getRange(a).clearContent(); });
}

function sendQualityData() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kirim = ss.getSheetByName('DataKirim');
    var quality = ss.getSheetByName('DataQuality');
    if (!kirim) { ui.alert('Error', 'Sheet DataKirim tidak ditemukan!', ui.ButtonSet.OK); return; }
    if (!quality) { ui.alert('Error', 'Sheet DataQuality tidak ditemukan!', ui.ButtonSet.OK); return; }
    if (quality.getLastColumn() < 1) { ui.alert('Error', 'Sheet DataQuality belum punya header!', ui.ButtonSet.OK); return; }

    var tanggal = kirim.getRange('H5').getValue();
    if (!tanggal) { ui.alert('Error', 'Tanggal di H5 belum diisi!', ui.ButtonSet.OK); return; }
    var tz = Session.getScriptTimeZone();
    var tanggalStr = (tanggal instanceof Date) ? Utilities.formatDate(tanggal, tz, 'yyyy-MM-dd') : String(tanggal).trim();
    var go = ui.alert('Kirim & pindahkan data?', 'Data kualitas tanggal ' + tanggalStr + ' akan dikirim ke DataQuality.\nSetelah berhasil, area input DataKirim dikosongkan (format & formula tetap).\n\nLanjutkan?', ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;

    var codes = kirim.getRange('C6:C41').getValues();
    var displayVals = kirim.getRange('D6:D41').getDisplayValues();
    var dataMap = {};
    for (var i = 0; i < codes.length; i++) {
      var code = String(codes[i][0]).trim();
      if (!code || code === '' || code === 'undefined') continue;
      var display = String(displayVals[i][0] || '').trim();
      if (display.length > 0 && display.charAt(0) === '#') continue;
      if (display === '') { dataMap[code] = ''; continue; }
      var num = parseFloat(display.replace(',', '.'));
      dataMap[code] = isNaN(num) ? display : num;
    }

    var qHeaders = quality.getRange(1, 1, 1, quality.getLastColumn()).getValues()[0];
    var colTanggal = qHeaders.indexOf('tanggal');
    if (colTanggal < 0) { ui.alert('Error', 'Kolom "tanggal" tidak ditemukan di DataQuality!', ui.ButtonSet.OK); return; }
    var lastRow = quality.getLastRow();
    if (lastRow > 1) {
      var tanggalCol = quality.getRange(2, colTanggal + 1, lastRow - 1, 1).getValues();
      for (var d = tanggalCol.length - 1; d >= 0; d--) {
        var rowDate = tanggalCol[d][0];
        var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') : String(rowDate).trim();
        if (rowDateStr === tanggalStr) {
          var resp = ui.alert('Data sudah ada', 'Data kualitas tanggal ' + tanggalStr + ' sudah ada.\nTimpa data lama?', ui.ButtonSet.YES_NO);
          if (resp === ui.Button.YES) { quality.deleteRow(d + 2); } else { return; }
        }
      }
    }

    var now = new Date();
    var qcId = genId_('QC-');
    var row = [];
    for (var h = 0; h < qHeaders.length; h++) {
      switch (qHeaders[h]) {
        case 'qc_id': row.push(qcId); break;
        case 'timestamp': row.push(now.toISOString()); break;
        case 'tanggal': row.push(tanggalStr); break;
        case 'nama_operator': row.push(Session.getActiveUser().getEmail() || 'Manual'); break;
        case 'user_id': row.push('MANUAL'); break;
        default: row.push(dataMap.hasOwnProperty(qHeaders[h]) ? dataMap[qHeaders[h]] : '');
      }
    }
    quality.appendRow(row);
    SpreadsheetApp.flush();
    // Pastikan baris benar-benar tersimpan sebelum mengosongkan input
    var colId = qHeaders.indexOf('qc_id');
    if (colId >= 0 && String(quality.getRange(quality.getLastRow(), colId + 1).getValue()) !== qcId) {
      ui.alert('Error', 'Data terkirim tidak terverifikasi. Input DataKirim TIDAK dikosongkan.', ui.ButtonSet.OK);
      return;
    }
    clearDataKirimInputs_(kirim);
    var filled = 0;
    for (var key in dataMap) { if (dataMap[key] !== '' && dataMap[key] !== 0) filled++; }
    ui.alert('Berhasil!', 'Data kualitas tanggal ' + tanggalStr + ' berhasil dipindahkan ke DataQuality.\nID: ' + qcId + '\nField terisi: ' + filled + ' dari ' + Object.keys(dataMap).length + '\n\nArea input DataKirim sudah dikosongkan.', ui.ButtonSet.OK);
  } catch (e) { ui.alert('Error', 'Gagal kirim data:\n' + e.message, ui.ButtonSet.OK); Logger.log('sendQualityData error: ' + e.message + '\n' + e.stack); }
}

function cutPasteDataKirim() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataKirim');
    if (!sheet) { ui.alert('Error', 'Sheet DataKirim tidak ditemukan!', ui.ButtonSet.OK); return; }
    var resp = ui.alert('Konfirmasi', 'Pindahkan data dari kolom M ke kolom T?\n(Data sumber akan dihapus setelah dipindah)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    var mapping = [['M14:R14','T10:Y10'],['M20:R20','T11:Y11'],['M26:R26','T12:Y12'],['M32:R32','T13:Y13'],['M38:R38','T14:Y14'],['M44:R44','T15:Y15'],['M50:R50','T16:Y16']];
    var rowHasData = function(vals) { return vals[0].some(function(v) { return v !== '' && v !== null; }); };
    // Cek baris tujuan yang masih berisi data dan akan tertimpa
    var conflicts = mapping.filter(function(m) {
      return rowHasData(sheet.getRange(m[0]).getValues()) && rowHasData(sheet.getRange(m[1]).getValues());
    }).map(function(m) { return m[1]; });
    if (conflicts.length) {
      var ow = ui.alert('Data tujuan masih terisi',
        'Baris tujuan berikut masih berisi data dan akan DITIMPA:\n• ' + conflicts.join('\n• ') + '\n\nTetap timpa?', ui.ButtonSet.YES_NO);
      if (ow !== ui.Button.YES) return;
    }
    var moved = 0;
    for (var i = 0; i < mapping.length; i++) {
      var srcRange = sheet.getRange(mapping[i][0]);
      var srcVals = srcRange.getValues();
      if (rowHasData(srcVals)) { sheet.getRange(mapping[i][1]).setValues(srcVals); srcRange.clearContent(); moved++; }
    }
    ui.alert('Berhasil!', moved + ' dari 7 baris berhasil dipindahkan.\n(' + (7 - moved) + ' baris kosong di-skip)', ui.ButtonSet.OK);
  } catch (e) { ui.alert('Error', 'Gagal pindah data:\n' + e.message, ui.ButtonSet.OK); }
}

function clearDataKirim() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataKirim');
    if (!sheet) { ui.alert('Error', 'Sheet DataKirim tidak ditemukan!', ui.ButtonSet.OK); return; }
    var resp = ui.alert('Konfirmasi', 'Hapus semua data input di DataKirim?\n(Hanya isi cell yang dihapus, format & formula tetap aman)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    clearDataKirimInputs_(sheet);
    ui.alert('Berhasil!', DATAKIRIM_INPUT_RANGES_.length + ' area data berhasil dihapus.', ui.ButtonSet.OK);
  } catch (e) { ui.alert('Error', 'Gagal hapus data:\n' + e.message, ui.ButtonSet.OK); }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('WWTP Tools')
    .addItem('Kirim & Pindahkan Data Kualitas', 'sendQualityData')
    .addItem('Pindahkan Data (Cut-Paste)', 'cutPasteDataKirim')
    .addSeparator()
    .addItem('Hapus Data Input', 'clearDataKirim')
    .addSeparator()
    .addSubMenu(ui.createMenu('Backup')
      .addItem('Backup Sekarang', 'backupNow')
      .addItem('Aktifkan Auto-Backup Harian', 'setupDailyBackupTrigger')
      .addItem('Nonaktifkan Auto-Backup', 'removeDailyBackupTrigger')
      .addSeparator()
      .addItem('Lihat Daftar Backup', 'listBackups')
      .addItem('Buka Spreadsheet Backup', 'openBackupSpreadsheet'))
    .addSubMenu(ui.createMenu('Admin')
      .addItem('Siapkan Sheet & Kolom Baru', 'prepareSheets')
      .addItem('Amankan Password Users (hash)', 'hashUserPasswords')
      .addItem('Muat Ulang Cache Users & Config', 'refreshCaches')
      .addSeparator()
      .addItem('Setup Ulang Semua Sheet (RESET)', 'setupAll'))
    .addToUi();
}
