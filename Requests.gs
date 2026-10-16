/**
 * ============================================
 * Requests.gs — Request Edit/Hapus + Approval
 * - Request baru pada record yang sama otomatis menggantikan
 *   request pending sebelumnya (status: superseded).
 * - Approver: Config 'approver_ids' (default Rifki M., Taryana, Hari Nuralam).
 * - Perubahan data diterapkan DULU, status ditulis setelah berhasil.
 * ============================================
 */

var REQ_HEADERS_ = ['request_id', 'timestamp', 'type', 'record_id', 'record_info', 'requester_name', 'requester_id',
  'status', 'edit_data', 'approver', 'resolved_at', 'reason', 'module'];

function reqModuleKey_(v) { var k = String(v || '').trim(); return k || 'ops'; }

function createRequest(token, req) {
  return api_(token, { write: true }, function(user) {
    req = req || {};
    var type = req.type;
    if (type !== 'edit' && type !== 'delete') return { success: false, message: 'Tipe request tidak valid.' };
    var mod = getModule_(reqModuleKey_(req.module));
    var recordId = String(req.recordId || '');

    var editData = null;
    if (type === 'edit') {
      editData = sanitizeEntry_(mod, req.editData);
      if (!Object.keys(editData).length) return { success: false, message: 'Tidak ada data yang diubah.' };
      if (editData.hasOwnProperty('tanggal') && !editData.tanggal) return { success: false, message: 'Tanggal tidak valid.' };
      if (editData.hasOwnProperty('shift') && !editData.shift) return { success: false, message: 'Shift tidak valid.' };
    }

    var found = findRecord_(mod, recordId);
    if (!found) return { success: false, message: 'Record tidak ditemukan.' };
    var rec = found.obj;
    var info = mod.title + ' — ' + rec.tanggal + ' — Shift ' + rec.shift + ' — ' + (rec.nama_operator || '');

    var out = withLock_(function() {
      var sh = ensureSheet_('Requests', REQ_HEADERS_);
      var ci = colIndex_(sh.headers);
      var reqId = genId_('REQ-');
      var now = new Date().toISOString();
      var superseded = [];

      var last = sh.sheet.getLastRow();
      if (last > 1) {
        var data = sh.sheet.getRange(2, 1, last - 1, sh.headers.length).getValues();
        for (var i = 0; i < data.length; i++) {
          var row = data[i];
          if (String(row[ci.status]) !== 'pending') continue;
          if (String(row[ci.record_id]) !== recordId) continue;
          if (reqModuleKey_(row[ci.module]) !== mod.key) continue;
          var rn = i + 2;
          sh.sheet.getRange(rn, ci.status + 1).setValue('superseded');
          sh.sheet.getRange(rn, ci.approver + 1).setValue('SYSTEM');
          sh.sheet.getRange(rn, ci.resolved_at + 1).setValue(now);
          sh.sheet.getRange(rn, ci.reason + 1).setValue('Digantikan oleh ' + reqId);
          superseded.push(String(row[ci.request_id]));
        }
      }

      sh.sheet.appendRow(rowFromObj_(sh.headers, {
        request_id: reqId, timestamp: now, type: type, record_id: recordId, record_info: info,
        requester_name: user.nama, requester_id: user.user_id, status: 'pending',
        edit_data: editData ? JSON.stringify(editData) : '', approver: '', resolved_at: '', reason: '', module: mod.key
      }));
      bumpVer_('req');
      bumpVer_('rec_' + mod.key);
      return { reqId: reqId, superseded: superseded };
    });

    notifyNewRequest_(type, out.reqId, info, user.nama, out.superseded);
    var msg = 'Request ' + (type === 'edit' ? 'perubahan' : 'penghapusan') + ' terkirim.\nID: ' + out.reqId;
    if (out.superseded.length) msg += '\nRequest sebelumnya (' + out.superseded.join(', ') + ') otomatis dibatalkan & digantikan.';
    return { success: true, message: msg, request_id: out.reqId, superseded: out.superseded };
  });
}

/**
 * Approver melihat semua request; user lain hanya request miliknya.
 * opts.status = 'pending' → hanya yang menunggu. opts.countOnly → hanya jumlah pending.
 */
function getRequests(token, opts) {
  return api_(token, {}, function(user) {
    opts = opts || {};
    var sheet = getSheet_('Requests');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, requests: [], pendingCount: 0, isApprover: user.isApprover };
    var all = sheet.getDataRange().getValues();
    var ci = colIndex_(all[0].map(String));
    var get = function(row, k) { return ci.hasOwnProperty(k) ? row[ci[k]] : ''; };

    var list = [], pendingCount = 0;
    for (var i = all.length - 1; i >= 1; i--) {
      var row = all[i];
      if (!user.isApprover && String(get(row, 'requester_id')) !== String(user.user_id)) continue;
      var status = String(get(row, 'status'));
      if (status === 'pending') pendingCount++;
      if (opts.countOnly) continue;
      if (opts.status === 'pending' && status !== 'pending') continue;
      if (list.length >= 100) continue;
      var ts = get(row, 'timestamp'), rs = get(row, 'resolved_at');
      list.push({
        request_id: String(get(row, 'request_id')),
        timestamp: ts instanceof Date ? ts.toISOString() : String(ts || ''),
        type: String(get(row, 'type')), module: reqModuleKey_(get(row, 'module')),
        record_id: String(get(row, 'record_id')), record_info: String(get(row, 'record_info')),
        requester_name: String(get(row, 'requester_name')), status: status,
        approver: String(get(row, 'approver') || ''), resolved_at: rs instanceof Date ? rs.toISOString() : String(rs || ''),
        reason: String(get(row, 'reason') || ''), _edit: String(get(row, 'edit_data') || '')
      });
    }
    if (opts.countOnly) return { success: true, pendingCount: pendingCount, isApprover: user.isApprover };

    // Ringkasan perubahan untuk request edit yang masih pending
    var recCache = {};
    list.forEach(function(r) {
      if (r.type === 'edit' && r.status === 'pending' && r._edit) {
        try { r.changes = diffEdit_(r.module, r.record_id, JSON.parse(r._edit), recCache); } catch (e) { r.changes = []; }
      }
      try { r.module_title = getModule_(r.module).title; } catch (e) { r.module_title = r.module; }
      delete r._edit;
    });
    return { success: true, requests: list, pendingCount: pendingCount, isApprover: user.isApprover };
  });
}

function diffEdit_(moduleKey, recordId, editData, recCache) {
  var mod = getModule_(moduleKey);
  if (!recCache[moduleKey]) {
    recCache[moduleKey] = {};
    var sheet = getSheet_(mod.sheet);
    if (sheet && sheet.getLastRow() > 1) {
      var all = sheet.getDataRange().getValues();
      var headers = all[0].map(String);
      for (var i = 1; i < all.length; i++) {
        var o = rowToObj_(headers, all[i]);
        recCache[moduleKey][String(o.record_id)] = o;
      }
    }
  }
  var cur = recCache[moduleKey][String(recordId)];
  if (!cur) return [{ key: '_', label: 'Record', from: '(sudah tidak ada)', to: '' }];
  var changes = [];
  Object.keys(editData).forEach(function(k) {
    var nv = editData[k], ov = cur[k];
    var same;
    if (typeof nv === 'boolean') same = isTrue_(ov) === nv;
    else if (typeof nv === 'number') same = parseFloat(ov) === nv;
    else same = String(ov === null || ov === undefined ? '' : ov).trim() === String(nv).trim();
    if (!same) changes.push({ key: k, label: mod.labels[k] || k, from: ov, to: nv });
  });
  return changes;
}

/** Map record_id → {request_id, type} untuk request pending di modul tertentu */
function getPendingMap_(moduleKey) {
  var map = {};
  var sheet = getSheet_('Requests');
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var all = sheet.getDataRange().getValues();
  var ci = colIndex_(all[0].map(String));
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][ci.status]) !== 'pending') continue;
    if (reqModuleKey_(ci.hasOwnProperty('module') ? all[i][ci.module] : '') !== moduleKey) continue;
    map[String(all[i][ci.record_id])] = { request_id: String(all[i][ci.request_id]), type: String(all[i][ci.type]) };
  }
  return map;
}

function resolveRequest(token, requestId, action, reason) {
  return api_(token, { approver: true }, function(user) {
    if (action !== 'approved' && action !== 'rejected') return { success: false, message: 'Action tidak valid.' };
    var res = withLock_(function() {
      var sheet = getSheet_('Requests');
      if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Request tidak ditemukan.' };
      var headers = ensureColumns_(sheet, REQ_HEADERS_);
      var all = sheet.getDataRange().getValues();
      var ci = colIndex_(headers);
      for (var i = 1; i < all.length; i++) {
        var row = all[i];
        if (String(row[ci.request_id]) !== String(requestId)) continue;
        var status = String(row[ci.status]);
        if (status !== 'pending') return { success: false, message: 'Request sudah diproses sebelumnya (status: ' + status + ').' };
        var mod = getModule_(reqModuleKey_(row[ci.module]));
        var recId = String(row[ci.record_id]);
        var type = String(row[ci.type]);

        if (action === 'approved') {
          var applied = type === 'delete'
            ? applyDelete_(mod, recId)
            : applyUpdate_(mod, recId, JSON.parse(String(row[ci.edit_data] || '{}')));
          if (!applied.success) return { success: false, message: 'Gagal menerapkan request: ' + applied.message + '\nStatus request tetap Menunggu.' };
        }

        var rn = i + 1, now = new Date().toISOString();
        sheet.getRange(rn, ci.status + 1).setValue(action);
        sheet.getRange(rn, ci.approver + 1).setValue(user.nama);
        sheet.getRange(rn, ci.resolved_at + 1).setValue(now);
        if (reason) sheet.getRange(rn, ci.reason + 1).setValue(safeCell_(reason));
        bumpVer_('req');
        bumpVer_('rec_' + mod.key);
        return { success: true, message: 'Request ' + requestId + ' ' + (action === 'approved' ? 'disetujui' : 'ditolak') + '.',
          info: String(row[ci.record_info]), requester: String(row[ci.requester_name]) };
      }
      return { success: false, message: 'Request ID tidak ditemukan.' };
    });
    if (res.success) notifyRequestResolved_(requestId, action, res.info, res.requester, user.nama, reason);
    return { success: res.success, message: res.message };
  });
}

/** Terapkan edit. Kolom sistem (record_id, timestamp, operator, user_id) tidak pernah berubah. */
function applyUpdate_(mod, recordId, editData) {
  editData = sanitizeEntry_(mod, editData);
  var found = findRecord_(mod, recordId);
  if (!found) return { success: false, message: 'Record tidak ditemukan.' };
  var ci = colIndex_(found.headers);
  var cur = found.all[found.index];

  if (mod.uniquePerShift) {
    var newTanggal = editData.tanggal || toYmd_(cur[ci.tanggal]);
    var newShift = editData.shift || cur[ci.shift];
    var dup = findShiftDuplicate_(found.all, newTanggal, newShift, recordId);
    if (dup) return { success: false, message: 'Sudah ada data ' + mod.title + ' lain di ' + newTanggal + ' Shift ' + newShift + ' (oleh ' + dup + ').' };
  }

  var row = found.headers.map(function(h, j) {
    return (h && SYSTEM_KEYS_.indexOf(h) < 0 && editData.hasOwnProperty(h)) ? editData[h] : cur[j];
  });
  found.sheet.getRange(found.index + 1, 1, 1, row.length).setValues([row]);
  bumpVer_('rec_' + mod.key);
  return { success: true };
}

function applyDelete_(mod, recordId) {
  var found = findRecord_(mod, recordId);
  if (!found) return { success: false, message: 'Record tidak ditemukan (mungkin sudah dihapus).' };
  found.sheet.deleteRow(found.index + 1);
  bumpVer_('rec_' + mod.key);
  return { success: true };
}
