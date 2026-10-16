/**
 * ============================================
 * CRUD.gs — Create, Read, Update, Delete
 * WWTP Operations Data Entry System
 * FIXED: ID collision, shift type comparison
 * ============================================
 */

var HEADERS_CACHE = null;

function getDataOpsHeaders_() {
  if (HEADERS_CACHE) return HEADERS_CACHE;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataOps');
  if (!sheet || sheet.getLastColumn() < 1) return [];
  HEADERS_CACHE = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return HEADERS_CACHE;
}

function getDataOpsSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataOps');
}

/**
 * CREATE — Insert new record
 * FIX: Timestamp+random ID instead of sequential counter
 */
function createRecord(data) {
  try {
    var sheet = getDataOpsSheet_();
    if (!sheet) return { success: false, message: 'Sheet DataOps tidak ditemukan!' };
    var headers = getDataOpsHeaders_();
    if (!headers || headers.length === 0) return { success: false, message: 'Headers DataOps kosong!' };
    var allData = sheet.getDataRange().getValues();

    var colTanggal = headers.indexOf('tanggal');
    var colShift = headers.indexOf('shift');

    for (var d = 1; d < allData.length; d++) {
      var rowDate = allData[d][colTanggal];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      var rowShift = String(allData[d][colShift]);
      if (String(rowDate) === String(data.tanggal) && rowShift === String(data.shift)) {
        var rowNama = String(allData[d][headers.indexOf('nama_operator')]).trim();
        return { success: false, message: 'Data tanggal ' + data.tanggal + ' Shift ' + data.shift + ' sudah diinput oleh ' + rowNama + '. Tidak bisa double input di shift yang sama.' };
      }
    }

    // FIX: Use timestamp + random to prevent ID collision
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var recordId = 'REC-' + Utilities.formatDate(now, tz, 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    data.record_id = recordId;
    data.timestamp = now.toISOString();

    var row = [];
    for (var i = 0; i < headers.length; i++) {
      row.push(data.hasOwnProperty(headers[i]) ? data[headers[i]] : '');
    }
    sheet.appendRow(row);

    var negDiffs = [];
    var clrFields = [
      {field:'Bak Sirkulasi', bKey:'clr_bak_sirkulasi_before', aKey:'clr_bak_sirkulasi_after'},
      {field:'Sungai', bKey:'clr_sungai_before', aKey:'clr_sungai_after'},
      {field:'MG 17-18', bKey:'clr_mg17_18_before', aKey:'clr_mg17_18_after'}
    ];
    clrFields.forEach(function(f) {
      var b = parseFloat(data[f.bKey]) || 0, a = parseFloat(data[f.aKey]) || 0;
      if (b > 0 && a > 0 && a < b) negDiffs.push({field:f.field, before:b, after:a, diff:Math.round((a-b)*100)/100});
    });
    if (negDiffs.length > 0) createNegativeDiffNotifs(recordId, data.tanggal, data.shift, data.nama_operator, negDiffs);

    var warning = negDiffs.length > 0 ? '\n⚠ ' + negDiffs.length + ' selisih negatif terdeteksi — notifikasi dikirim ke admin.' : '';
    invalidateDataCaches_();
    return { success: true, message: 'Data berhasil disimpan.' + warning, record_id: recordId };
  } catch (e) {
    Logger.log('createRecord error: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Gagal menyimpan: ' + e.message };
  }
}

/**
 * READ — Get records with filters & pagination
 */
function readRecords(options) {
  try {
    var p = (options && options.page) || 1;
    var sh = (options && options.shift) || '';
    var sd = (options && options.startDate) || '';
    var ed = (options && options.endDate) || '';
    var sq = (options && options.search) || '';
    var rCacheKey = 'records_' + p + '_' + sh + '_' + sd + '_' + ed + '_' + sq;
    var cached = cacheGet_(rCacheKey);
    if (cached) return cached;

    var sheet = getDataOpsSheet_();
    var headers = getDataOpsHeaders_();
    var allData = sheet.getDataRange().getValues();

    if (allData.length <= 1) return { success: true, records: [], total: 0, page: 1, totalPages: 0 };

    var page = p;
    var limit = (options && options.limit) || 15;
    var filterShift = options && options.shift;
    var filterStartDate = sd;
    var filterEndDate = ed;
    var filterOperator = options && options.operator;
    var searchQuery = sq ? String(sq).toLowerCase() : '';

    var tz = Session.getScriptTimeZone();
    var records = [];
    for (var i = allData.length - 1; i >= 1; i--) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = allData[i][j];
        if (val instanceof Date) {
          obj[headers[j]] = (headers[j] === 'tanggal')
            ? Utilities.formatDate(val, tz, 'yyyy-MM-dd')
            : Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm:ss');
        } else {
          obj[headers[j]] = val;
        }
      }

      if (filterShift && String(obj.shift) !== String(filterShift)) continue;
      if (filterStartDate || filterEndDate) {
        var recDate = obj.tanggal;
        if (recDate instanceof Date) recDate = Utilities.formatDate(recDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var rd = String(recDate);
        if (filterStartDate && rd < filterStartDate) continue;
        if (filterEndDate && rd > filterEndDate) continue;
      }
      if (filterOperator && String(obj.nama_operator).toLowerCase().indexOf(filterOperator.toLowerCase()) === -1) continue;
      if (searchQuery) {
        var searchStr = JSON.stringify(obj).toLowerCase();
        if (searchStr.indexOf(searchQuery) === -1) continue;
      }

      var tStr = String(obj.tanggal);
      if (tStr.match(/^\d{4}-\d{2}-\d{2}/)) {
        var parts = tStr.split('-');
        var months = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
        obj.tanggal_display = parseInt(parts[2]) + ' ' + months[parseInt(parts[1])] + ' ' + parts[0];
      } else {
        obj.tanggal_display = tStr;
      }
      records.push(obj);
    }

    var total = records.length;
    var totalPages = Math.ceil(total / limit);
    var start = (page - 1) * limit;
    var paged = records.slice(start, start + limit);

    var result = { success: true, records: paged, total: total, page: page, totalPages: totalPages };
    cachePut_(rCacheKey, result, CACHE_TTL_SHORT_);
    cacheTrackKey_('records', rCacheKey);
    return result;
  } catch (e) {
    return { success: false, message: 'Gagal memuat data: ' + e.message, records: [], total: 0 };
  }
}

function readRecordById(recordId) {
  try {
    var sheet = getDataOpsSheet_();
    var headers = getDataOpsHeaders_();
    var allData = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(recordId)) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var val = allData[i][j];
          if (val instanceof Date) {
            val = (headers[j] === 'tanggal')
              ? Utilities.formatDate(val, tz, 'yyyy-MM-dd')
              : Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm:ss');
          }
          obj[headers[j]] = val;
        }
        return { success: true, record: obj };
      }
    }
    return { success: false, message: 'Record tidak ditemukan.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

function updateRecord(recordId, data, password) {
  try {
    var updatePassword = getConfig_('delete_password') || '';
    if (String(password) !== updatePassword) return { success: false, message: 'Password salah! Update dibatalkan.' };
    var sheet = getDataOpsSheet_();
    var headers = getDataOpsHeaders_();
    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(recordId)) {
        data.record_id = recordId;
        data.timestamp = allData[i][1];
        var row = [];
        for (var j = 0; j < headers.length; j++) {
          row.push(data.hasOwnProperty(headers[j]) ? data[headers[j]] : allData[i][j]);
        }
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        invalidateDataCaches_();
        return { success: true, message: 'Data berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Record tidak ditemukan.' };
  } catch (e) { return { success: false, message: 'Gagal update: ' + e.message }; }
}

function deleteRecord(recordId, password) {
  try {
    var deletePassword = getConfig_('delete_password') || '';
    if (String(password) !== deletePassword) return { success: false, message: 'Password penghapusan salah!' };
    var sheet = getDataOpsSheet_();
    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(recordId)) {
        sheet.deleteRow(i + 1);
        invalidateDataCaches_();
        return { success: true, message: 'Data berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Record tidak ditemukan.' };
  } catch (e) { return { success: false, message: 'Gagal menghapus: ' + e.message }; }
}

// ============ QUALITY DATA ============

function getQualityDataForDashboard(options) {
  try {
    var sOpt = options && options.startDate ? options.startDate : 'auto';
    var eOpt = options && options.endDate ? options.endDate : 'auto';
    var qCacheKey = 'qual_' + sOpt + '_' + eOpt;
    var cached = cacheGet_(qCacheKey);
    if (cached) return cached;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataQuality');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: {} };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();

    var startDate = options && options.startDate ? new Date(options.startDate) : null;
    var endDate = options && options.endDate ? new Date(options.endDate) : null;
    if (!startDate) { startDate = new Date(); startDate.setDate(startDate.getDate() - 30); }
    if (!endDate) { endDate = new Date(); }
    startDate.setHours(0,0,0,0); endDate.setHours(23,59,59,999);

    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;

    var records = [];
    for (var i = 1; i < allData.length; i++) {
      var tanggal = allData[i][colIdx['tanggal']];
      var tanggalDate;
      if (tanggal instanceof Date) { tanggalDate = tanggal; }
      else if (typeof tanggal === 'string' && tanggal.match(/^\d{4}-\d{2}-\d{2}/)) { tanggalDate = new Date(tanggal.substring(0, 10) + 'T00:00:00'); }
      else if (typeof tanggal === 'string' && tanggal.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        var pts = tanggal.split('/');
        tanggalDate = new Date(pts[2] + '-' + ('0'+pts[1]).slice(-2) + '-' + ('0'+pts[0]).slice(-2) + 'T00:00:00');
      } else { continue; }
      if (tanggalDate < startDate || tanggalDate > endDate) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = allData[i][j];
      records.push(obj);
    }

    var pf = parseFloat;
    function avgF(arr, field) {
      var vals = [];
      arr.forEach(function(r) {
        var raw = r[field];
        if (raw === null || raw === undefined || raw === '') return;
        var v = (typeof raw === 'number') ? raw : pf(String(raw).replace(',', '.').trim());
        if (!isNaN(v) && v > 0) vals.push(v);
      });
      if (vals.length === 0) return 0;
      var s = 0; vals.forEach(function(v) { s += v; });
      return Math.round(s / vals.length * 100) / 100;
    }

    var clr = { ph: avgF(records,'clr_ph'), conductivity: avgF(records,'clr_conductivity'), tss: avgF(records,'clr_tss'), tds: avgF(records,'clr_tds') };
    var bsb = { ph: avgF(records,'bsb_ph'), conductivity: avgF(records,'bsb_conductivity'), tss: avgF(records,'bsb_tss'), tds: avgF(records,'bsb_tds') };
    var bs17 = { ph: avgF(records,'bs17_ph'), conductivity: avgF(records,'bs17_conductivity'), tss: avgF(records,'bs17_tss'), tds: avgF(records,'bs17_tds') };
    var pam = { ph: avgF(records,'pam_ph'), conductivity: avgF(records,'pam_conductivity'), hardness: avgF(records,'pam_hardness') };
    var sft = { ph: avgF(records,'sft_ph'), conductivity: avgF(records,'sft_conductivity'), hardness: avgF(records,'sft_hardness') };

    var sumur = {};
    for (var s = 1; s <= 6; s++) {
      var pfx = 'sm' + s + '_';
      sumur['sm' + s] = { ph: avgF(records, pfx+'ph'), conductivity: avgF(records, pfx+'conductivity'), hardness: avgF(records, pfx+'hardness') };
    }

    var qResult = { success: true, data: { clarifier: clr, bsb: bsb, bs17: bs17, pam: pam, softener: sft, sumur: sumur, count: records.length } };
    cachePut_(qCacheKey, qResult, CACHE_TTL_);
    cacheTrackKey_('qual', qCacheKey);
    return qResult;
  } catch (e) { return { success: false, message: e.message, data: {} }; }
}

// ============ REQUEST MANAGEMENT ============

var APPROVER_ID_ = 'USR005';

function getRequestsSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests'); }

function createRequest(type, recordId, recordInfo, editData, requesterName, requesterId) {
  try {
    var sheet = getRequestsSheet_();
    if (!sheet) return { success: false, message: 'Sheet Requests tidak ditemukan!' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var now = new Date();
    var reqId = 'REQ-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var row = [];
    for (var i = 0; i < headers.length; i++) {
      switch (headers[i]) {
        case 'request_id': row.push(reqId); break;
        case 'timestamp': row.push(now.toISOString()); break;
        case 'type': row.push(type); break;
        case 'record_id': row.push(recordId); break;
        case 'record_info': row.push(recordInfo); break;
        case 'requester_name': row.push(requesterName); break;
        case 'requester_id': row.push(requesterId); break;
        case 'status': row.push('pending'); break;
        case 'edit_data': row.push(editData ? JSON.stringify(editData) : ''); break;
        default: row.push('');
      }
    }
    sheet.appendRow(row);
    invalidateRequestsCaches_();
    notifyNewRequest_(type, reqId, recordInfo, requesterName);
    return { success: true, message: 'Request berhasil dikirim.\nID: ' + reqId, request_id: reqId };
  } catch (e) { return { success: false, message: 'Gagal membuat request: ' + e.message }; }
}

function getRequests(userId) {
  try {
    var rqCacheKey = 'req_' + userId;
    var cached = cacheGet_(rqCacheKey);
    if (cached) return cached;

    var sheet = getRequestsSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, requests: [] };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var isApprover = (userId === APPROVER_ID_);
    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;
    var requests = [];
    for (var i = allData.length - 1; i >= 1; i--) {
      var r = allData[i];
      if (!isApprover && String(r[colIdx['requester_id']]) !== String(userId)) continue;
      requests.push({ request_id: r[colIdx['request_id']], timestamp: r[colIdx['timestamp']], type: r[colIdx['type']], record_id: r[colIdx['record_id']], record_info: r[colIdx['record_info']], requester_name: r[colIdx['requester_name']], status: r[colIdx['status']], approver: r[colIdx['approver']] || '', resolved_at: r[colIdx['resolved_at']] || '', reason: r[colIdx['reason']] || '' });
      if (requests.length >= 50) break;
    }
    var rqResult = { success: true, requests: requests, isApprover: isApprover };
    cachePut_(rqCacheKey, rqResult, CACHE_TTL_SHORT_);
    cacheTrackKey_('req', rqCacheKey);
    return rqResult;
  } catch (e) { return { success: false, message: e.message, requests: [] }; }
}

function resolveRequest(requestId, action, approverId, approverName, reason) {
  try {
    if (approverId !== APPROVER_ID_) return { success: false, message: 'Hanya Rifki M. yang bisa approve/reject request.' };
    if (action !== 'approved' && action !== 'rejected') return { success: false, message: 'Action tidak valid.' };
    var sheet = getRequestsSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Request tidak ditemukan.' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;

    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][colIdx['request_id']]) !== requestId) continue;
      if (String(allData[i][colIdx['status']]) !== 'pending') return { success: false, message: 'Request sudah diproses sebelumnya.' };
      var rowNum = i + 1;
      var now = new Date();
      sheet.getRange(rowNum, colIdx['status'] + 1).setValue(action);
      sheet.getRange(rowNum, colIdx['approver'] + 1).setValue(approverName);
      sheet.getRange(rowNum, colIdx['resolved_at'] + 1).setValue(now.toISOString());
      if (reason) sheet.getRange(rowNum, colIdx['reason'] + 1).setValue(reason);

      if (action === 'approved') {
        var type = String(allData[i][colIdx['type']]);
        var recId = String(allData[i][colIdx['record_id']]);
        var pwd = getConfig_('delete_password') || 'wwtpmonokem';
        if (type === 'delete') {
          var delResult = deleteRecord(recId, pwd);
          if (!delResult.success) return { success: false, message: 'Approve OK tapi gagal hapus: ' + delResult.message };
        } else if (type === 'edit') {
          var editJson = String(allData[i][colIdx['edit_data']]);
          if (editJson) {
            var editData = JSON.parse(editJson);
            var updResult = updateRecord(recId, editData, pwd);
            if (!updResult.success) return { success: false, message: 'Approve OK tapi gagal update: ' + updResult.message };
          }
        }
      }
      var reqInfo = String(allData[i][colIdx['record_info']]);
      var reqName = String(allData[i][colIdx['requester_name']]);
      notifyRequestResolved_(requestId, action, reqInfo, reqName, approverName, reason);
      invalidateRequestsCaches_();
      return { success: true, message: 'Request ' + requestId + ' berhasil di-' + action + '.' };
    }
    return { success: false, message: 'Request ID tidak ditemukan.' };
  } catch (e) { return { success: false, message: 'Error: ' + e.message }; }
}

// ============ LAST SHIFT DATA ============
// FIX: Handle shift as both int and string

function getLastShiftData(tanggal, shift) {
  try {
    var lsCacheKey = 'lastshift_' + tanggal + '_' + shift;
    var cached = cacheGet_(lsCacheKey);
    if (cached) return cached;

    var sheet = getDataOpsSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: null };
    var headers = getDataOpsHeaders_();
    var allData = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;

    // FIX: Normalize shift to number for comparison
    var shiftNum = parseInt(shift) || 1;
    var prevTanggal, prevShift;
    if (shiftNum === 1) {
      var d = new Date(tanggal + 'T00:00:00'); d.setDate(d.getDate() - 1);
      prevTanggal = Utilities.formatDate(d, tz, 'yyyy-MM-dd'); prevShift = '3';
    } else if (shiftNum === 2) { prevTanggal = tanggal; prevShift = '1'; }
    else { prevTanggal = tanggal; prevShift = '2'; }

    var found = null;
    for (var i = allData.length - 1; i >= 1; i--) {
      var rowDate = allData[i][colIdx['tanggal']];
      var rowDateStr;
      if (rowDate instanceof Date) { rowDateStr = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd'); }
      else if (typeof rowDate === 'string' && rowDate.match(/^\d{4}-\d{2}-\d{2}/)) { rowDateStr = rowDate.substring(0, 10); }
      else if (typeof rowDate === 'string' && rowDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        var p = rowDate.split('/'); rowDateStr = p[2] + '-' + ('0'+p[1]).slice(-2) + '-' + ('0'+p[0]).slice(-2);
      } else continue;
      if (rowDateStr === prevTanggal && String(allData[i][colIdx['shift']]) === prevShift) {
        found = {};
        for (var j = 0; j < headers.length; j++) found[headers[j]] = allData[i][j];
        break;
      }
    }
    if (!found) return { success: true, data: null };
    var lsResult = { success: true, data: {
      clr_bak_sirkulasi: parseFloat(found.clr_bak_sirkulasi_after) || 0,
      clr_sungai: parseFloat(found.clr_sungai_after) || 0,
      clr_mg17_18: parseFloat(found.clr_mg17_18_after) || 0,
      fm_softener_lt3: parseFloat(found.fm_softener_lt3) || 0,
      fm_flotation: parseFloat(found.fm_flotation) || 0,
      fm_mg16: parseFloat(found.fm_mg16) || 0,
      fm_pam: parseFloat(found.fm_pam) || 0,
      fm_pam_shift: parseFloat(found.fm_pam_shift) || 0,
      fm_sm1: parseFloat(found.fm_sm1) || 0, fm_sm2: parseFloat(found.fm_sm2) || 0,
      fm_sm3: parseFloat(found.fm_sm3) || 0, fm_sm4: parseFloat(found.fm_sm4) || 0,
      fm_sm5: parseFloat(found.fm_sm5) || 0, fm_sm6: parseFloat(found.fm_sm6) || 0,
      fm_leaching: parseFloat(found.fm_leaching) || 0,
      prev_shift: prevShift, prev_tanggal: prevTanggal, prev_operator: found.nama_operator || ''
    }};
    cachePut_(lsCacheKey, lsResult, CACHE_TTL_SHORT_);
    cacheTrackKey_('lastshift', lsCacheKey);
    return lsResult;
  } catch (e) { return { success: false, message: e.message, data: null }; }
}

// ============ NEGATIVE DIFF NOTIFICATIONS ============

function createNegativeDiffNotifs(recordId, tanggal, shift, operatorName, diffs) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
    if (!sheet) return;
    var tz = Session.getScriptTimeZone(); var now = new Date();
    for (var i = 0; i < diffs.length; i++) {
      var dd = diffs[i];
      var notifId = 'NTF-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + i;
      sheet.appendRow([notifId, now.toISOString(), 'negative_diff', recordId, tanggal, shift, operatorName, dd.field, dd.before, dd.after, dd.diff, 'unread', '', '', '']);
    }
    notifyAnomaly_(tanggal, shift, operatorName, diffs);
    invalidateNotifsCaches_();
  } catch (e) { Logger.log('createNegativeDiffNotifs error: ' + e.message); }
}

function getNotifications(userId) {
  try {
    var ntCacheKey = 'ntf_' + userId;
    var cached = cacheGet_(ntCacheKey);
    if (cached) return cached;

    var isApprover = (userId === APPROVER_ID_);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, notifications: [], isApprover: isApprover };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;
    var notifs = [], unreadCount = 0;
    for (var i = allData.length - 1; i >= 1; i--) {
      var r = allData[i];
      if (!isApprover && String(r[colIdx['operator_name']]).trim().toLowerCase() !== userId.toLowerCase()) continue;
      var status = String(r[colIdx['status']]);
      if (status === 'unread') unreadCount++;
      notifs.push({ notif_id: r[colIdx['notif_id']], timestamp: r[colIdx['timestamp']], type: r[colIdx['type']], record_id: r[colIdx['record_id']], tanggal: r[colIdx['tanggal']], shift: r[colIdx['shift']], operator_name: r[colIdx['operator_name']], field: r[colIdx['field']], before_val: r[colIdx['before_val']], after_val: r[colIdx['after_val']], diff: r[colIdx['diff']], status: status, notes: r[colIdx['notes']] || '' });
      if (notifs.length >= 50) break;
    }
    var ntResult = { success: true, notifications: notifs, unreadCount: unreadCount, isApprover: isApprover };
    cachePut_(ntCacheKey, ntResult, CACHE_TTL_SHORT_);
    cacheTrackKey_('ntf', ntCacheKey);
    return ntResult;
  } catch (e) { return { success: false, message: e.message, notifications: [] }; }
}

function resolveNotification(notifId, approverName, notes) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Tidak ditemukan.' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var colIdx = {};
    for (var h = 0; h < headers.length; h++) colIdx[headers[h]] = h;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][colIdx['notif_id']]) === notifId) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, colIdx['status'] + 1).setValue('resolved');
        sheet.getRange(rowNum, colIdx['resolved_by'] + 1).setValue(approverName);
        sheet.getRange(rowNum, colIdx['resolved_at'] + 1).setValue(new Date().toISOString());
        if (notes) sheet.getRange(rowNum, colIdx['notes'] + 1).setValue(notes);
        invalidateNotifsCaches_();
        return { success: true, message: 'Notifikasi ditandai resolved.' };
      }
    }
    return { success: false, message: 'Notif ID tidak ditemukan.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============ EMAIL NOTIFICATIONS ============

function getAdminEmail_() { return getConfig_('admin_email') || null; }

function sendAdminEmail_(subject, htmlBody) {
  try {
    var email = getAdminEmail_(); if (!email) return;
    MailApp.sendEmail({ to: email, subject: '[WWTP] ' + subject,
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;"><div style="background:#16A34A;padding:12px 20px;color:#fff;font-weight:bold;font-size:16px;">WWTP Operations System</div><div style="padding:20px;border:1px solid #e0e0e0;border-top:none;">' + htmlBody + '</div><div style="padding:10px 20px;font-size:11px;color:#999;border-top:1px solid #eee;">Email otomatis — jangan reply.</div></div>'
    });
  } catch (e) { Logger.log('sendAdminEmail_ error: ' + e.message); }
}

function notifyAnomaly_(tanggal, shift, operator, diffs) {
  var rows = '';
  diffs.forEach(function(d) {
    rows += '<tr><td style="padding:6px 10px;border:1px solid #e0e0e0;">' + d.field + '</td><td style="padding:6px 10px;border:1px solid #e0e0e0;text-align:center;">' + d.before + '</td><td style="padding:6px 10px;border:1px solid #e0e0e0;text-align:center;">' + d.after + '</td><td style="padding:6px 10px;border:1px solid #e0e0e0;text-align:center;color:#DC2626;font-weight:bold;">' + d.diff + '</td></tr>';
  });
  sendAdminEmail_('Anomali — Selisih Negatif (' + tanggal + ' Shift ' + shift + ')',
    '<h2 style="color:#DC2626;margin:0 0 10px;">⚠ Selisih Negatif Terdeteksi</h2><p><strong>Tanggal:</strong> ' + tanggal + ' &nbsp;|&nbsp; <strong>Shift:</strong> ' + shift + ' &nbsp;|&nbsp; <strong>Operator:</strong> ' + operator + '</p><table style="border-collapse:collapse;width:100%;margin:12px 0;"><tr style="background:#f0f0f0;"><th style="padding:8px 10px;border:1px solid #ccc;text-align:left;">Field</th><th style="padding:8px 10px;border:1px solid #ccc;">Before</th><th style="padding:8px 10px;border:1px solid #ccc;">After</th><th style="padding:8px 10px;border:1px solid #ccc;">Selisih</th></tr>' + rows + '</table><p style="color:#666;font-size:13px;">Data tetap tersimpan. Silakan tinjau di menu Notifikasi Anomali.</p>');
}

function notifyNewRequest_(type, reqId, recordInfo, requesterName) {
  var typeLabel = type === 'edit' ? 'Perubahan Data' : 'Penghapusan Data';
  sendAdminEmail_('Request ' + typeLabel + ' — ' + requesterName,
    '<h2 style="color:#CA8A04;margin:0 0 10px;">🔔 Request Baru Masuk</h2><table style="border-collapse:collapse;margin:12px 0;"><tr><td style="padding:6px 10px;color:#666;">Tipe</td><td style="padding:6px 10px;font-weight:bold;">' + typeLabel + '</td></tr><tr><td style="padding:6px 10px;color:#666;">ID</td><td style="padding:6px 10px;font-family:monospace;">' + reqId + '</td></tr><tr><td style="padding:6px 10px;color:#666;">Record</td><td style="padding:6px 10px;">' + recordInfo + '</td></tr><tr><td style="padding:6px 10px;color:#666;">Pengaju</td><td style="padding:6px 10px;font-weight:bold;">' + requesterName + '</td></tr></table><p style="color:#666;font-size:13px;">Silakan approve/reject di menu Riwayat Request.</p>');
}

function notifyRequestResolved_(reqId, action, recordInfo, requesterName, approverName, reason) {
  var statusColor = action === 'approved' ? '#16A34A' : '#DC2626';
  var statusLabel = action === 'approved' ? 'DISETUJUI' : 'DITOLAK';
  var emoji = action === 'approved' ? '✅' : '❌';
  sendAdminEmail_('Request ' + statusLabel + ' — ' + reqId,
    '<h2 style="color:' + statusColor + ';margin:0 0 10px;">' + emoji + ' Request ' + statusLabel + '</h2><table style="border-collapse:collapse;margin:12px 0;"><tr><td style="padding:6px 10px;color:#666;">ID</td><td style="padding:6px 10px;font-family:monospace;">' + reqId + '</td></tr><tr><td style="padding:6px 10px;color:#666;">Record</td><td style="padding:6px 10px;">' + recordInfo + '</td></tr><tr><td style="padding:6px 10px;color:#666;">Pengaju</td><td style="padding:6px 10px;">' + requesterName + '</td></tr><tr><td style="padding:6px 10px;color:#666;">Oleh</td><td style="padding:6px 10px;font-weight:bold;">' + approverName + '</td></tr>' + (reason ? '<tr><td style="padding:6px 10px;color:#666;">Alasan</td><td style="padding:6px 10px;">' + reason + '</td></tr>' : '') + '</table>');
}
