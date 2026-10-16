/**
 * ============================================
 * Notifications.gs — Notifikasi Anomali + Email
 * - Operator melihat notifikasi miliknya (dicocokkan via user_id,
 *   fallback nama operator untuk baris lama).
 * - Approver melihat semua & bisa menandai resolved.
 * ============================================
 */

var NTF_HEADERS_ = ['notif_id', 'timestamp', 'type', 'record_id', 'tanggal', 'shift', 'operator_name', 'field',
  'before_val', 'after_val', 'diff', 'status', 'resolved_by', 'resolved_at', 'notes', 'user_id'];

/** Cek selisih negatif clarifier pada entri Operasional. Return jumlah anomali. */
function checkNegativeDiffs_(data) {
  var fields = [
    { field: 'Bak Sirkulasi', b: 'clr_bak_sirkulasi_before', a: 'clr_bak_sirkulasi_after' },
    { field: 'Sungai', b: 'clr_sungai_before', a: 'clr_sungai_after' },
    { field: 'MG 17-18', b: 'clr_mg17_18_before', a: 'clr_mg17_18_after' }
  ];
  var diffs = [];
  fields.forEach(function(f) {
    var b = parseFloat(data[f.b]) || 0, a = parseFloat(data[f.a]) || 0;
    if (b > 0 && a > 0 && a < b) diffs.push({ field: f.field, before: b, after: a, diff: Math.round((a - b) * 100) / 100 });
  });
  if (diffs.length) createNegativeDiffNotifs_(data, diffs);
  return diffs.length;
}

function createNegativeDiffNotifs_(data, diffs) {
  try {
    var sh = ensureSheet_('Notifications', NTF_HEADERS_);
    var now = new Date().toISOString();
    diffs.forEach(function(d) {
      sh.sheet.appendRow(rowFromObj_(sh.headers, {
        notif_id: genId_('NTF-'), timestamp: now, type: 'negative_diff', record_id: data.record_id,
        tanggal: data.tanggal, shift: data.shift, operator_name: data.nama_operator, field: d.field,
        before_val: d.before, after_val: d.after, diff: d.diff, status: 'unread',
        resolved_by: '', resolved_at: '', notes: '', user_id: data.user_id
      }));
    });
    bumpVer_('ntf');
    notifyAnomaly_(data.tanggal, data.shift, data.nama_operator, diffs);
  } catch (e) { Logger.log('createNegativeDiffNotifs_ error: ' + e.message); }
}

function getNotifications(token, opts) {
  return api_(token, {}, function(user) {
    opts = opts || {};
    var key = cacheKey_(['ntf', ver_('ntf'), user.user_id, user.isApprover, !!opts.countOnly]);
    var cached = cacheGet_(key);
    if (cached) return cached;

    var sheet = getSheet_('Notifications');
    var res = { success: true, notifications: [], unreadCount: 0, isApprover: user.isApprover };
    if (sheet && sheet.getLastRow() > 1) {
      var all = sheet.getDataRange().getValues();
      var ci = colIndex_(all[0].map(String));
      var get = function(row, k) { return ci.hasOwnProperty(k) ? row[ci[k]] : ''; };
      var myName = String(user.nama).trim().toLowerCase();
      for (var i = all.length - 1; i >= 1; i--) {
        var r = all[i];
        if (!user.isApprover) {
          var uid = String(get(r, 'user_id') || '').trim();
          var mine = uid ? uid === String(user.user_id) : String(get(r, 'operator_name')).trim().toLowerCase() === myName;
          if (!mine) continue;
        }
        var status = String(get(r, 'status'));
        if (status === 'unread') res.unreadCount++;
        if (opts.countOnly || res.notifications.length >= 100) continue;
        var ts = get(r, 'timestamp');
        res.notifications.push({
          notif_id: String(get(r, 'notif_id')), timestamp: ts instanceof Date ? ts.toISOString() : String(ts || ''),
          type: String(get(r, 'type')), record_id: String(get(r, 'record_id')),
          tanggal: toYmd_(get(r, 'tanggal')) || String(get(r, 'tanggal')), shift: String(get(r, 'shift')),
          operator_name: String(get(r, 'operator_name')), field: String(get(r, 'field')),
          before_val: get(r, 'before_val'), after_val: get(r, 'after_val'), diff: get(r, 'diff'),
          status: status, resolved_by: String(get(r, 'resolved_by') || ''), notes: String(get(r, 'notes') || '')
        });
      }
    }
    cachePut_(key, res, CACHE_TTL_SHORT_);
    return res;
  });
}

function resolveNotification(token, notifId, notes) {
  return api_(token, { approver: true }, function(user) {
    return withLock_(function() {
      var sheet = getSheet_('Notifications');
      if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Tidak ditemukan.' };
      var headers = ensureColumns_(sheet, NTF_HEADERS_);
      var ci = colIndex_(headers);
      var ids = sheet.getRange(2, ci.notif_id + 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) !== String(notifId)) continue;
        var rn = i + 2;
        sheet.getRange(rn, ci.status + 1).setValue('resolved');
        sheet.getRange(rn, ci.resolved_by + 1).setValue(user.nama);
        sheet.getRange(rn, ci.resolved_at + 1).setValue(new Date().toISOString());
        if (notes) sheet.getRange(rn, ci.notes + 1).setValue(safeCell_(notes));
        bumpVer_('ntf');
        return { success: true, message: 'Notifikasi ditandai resolved.' };
      }
      return { success: false, message: 'Notif ID tidak ditemukan.' };
    });
  });
}

// ============ EMAIL ============
// Penerima: Config 'admin_email' (boleh lebih dari satu, pisahkan dengan koma)
function sendAdminEmail_(subject, htmlBody) {
  try {
    var to = getConfig_('admin_email');
    if (!to) return;
    MailApp.sendEmail({ to: to, subject: '[WWTP] ' + subject,
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;"><div style="background:#000;padding:12px 20px;color:#3FDC6A;font-family:monospace;font-size:14px;">wwtp-ops/</div><div style="padding:20px;border:1px solid #C3CBC6;border-top:none;background:#EEF1EF;color:#0A0D0B;">' + htmlBody + '</div><div style="padding:10px 20px;font-size:11px;color:#66706A;">Email otomatis — jangan reply.</div></div>'
    });
  } catch (e) { Logger.log('sendAdminEmail_ error: ' + e.message); }
}

function emailRow_(label, value, bold) {
  return '<tr><td style="padding:6px 10px;color:#66706A;">' + escHtml_(label) + '</td><td style="padding:6px 10px;' + (bold ? 'font-weight:bold;' : '') + '">' + escHtml_(value) + '</td></tr>';
}

function notifyAnomaly_(tanggal, shift, operator, diffs) {
  var rows = '';
  diffs.forEach(function(d) {
    rows += '<tr><td style="padding:6px 10px;border:1px solid #C3CBC6;">' + escHtml_(d.field) + '</td><td style="padding:6px 10px;border:1px solid #C3CBC6;text-align:center;">' + d.before + '</td><td style="padding:6px 10px;border:1px solid #C3CBC6;text-align:center;">' + d.after + '</td><td style="padding:6px 10px;border:1px solid #C3CBC6;text-align:center;color:#C2261D;font-weight:bold;">' + d.diff + '</td></tr>';
  });
  sendAdminEmail_('Anomali — Selisih Negatif (' + tanggal + ' Shift ' + shift + ')',
    '<h2 style="color:#C2261D;margin:0 0 10px;">Selisih Negatif Terdeteksi</h2><p><strong>Tanggal:</strong> ' + escHtml_(tanggal) + ' &nbsp;|&nbsp; <strong>Shift:</strong> ' + escHtml_(shift) + ' &nbsp;|&nbsp; <strong>Operator:</strong> ' + escHtml_(operator) + '</p><table style="border-collapse:collapse;width:100%;margin:12px 0;"><tr><th style="padding:8px 10px;border:1px solid #C3CBC6;text-align:left;">Field</th><th style="padding:8px 10px;border:1px solid #C3CBC6;">Sebelum</th><th style="padding:8px 10px;border:1px solid #C3CBC6;">Sesudah</th><th style="padding:8px 10px;border:1px solid #C3CBC6;">Selisih</th></tr>' + rows + '</table><p style="color:#66706A;font-size:13px;">Data tetap tersimpan. Tinjau di menu Anomali.</p>');
}

function notifyNewRequest_(type, reqId, recordInfo, requesterName, superseded) {
  var typeLabel = type === 'edit' ? 'Perubahan Data' : 'Penghapusan Data';
  var extra = superseded && superseded.length ? emailRow_('Menggantikan', superseded.join(', ')) : '';
  sendAdminEmail_('Request ' + typeLabel + ' — ' + requesterName,
    '<h2 style="color:#0F8A35;margin:0 0 10px;">Request Baru Masuk</h2><table style="border-collapse:collapse;margin:12px 0;">' +
    emailRow_('Tipe', typeLabel, true) + emailRow_('ID', reqId) + emailRow_('Record', recordInfo) + emailRow_('Pengaju', requesterName, true) + extra +
    '</table><p style="color:#66706A;font-size:13px;">Approve/reject di menu Request.</p>');
}

function notifyRequestResolved_(reqId, action, recordInfo, requesterName, approverName, reason) {
  var label = action === 'approved' ? 'DISETUJUI' : 'DITOLAK';
  var color = action === 'approved' ? '#0F8A35' : '#C2261D';
  sendAdminEmail_('Request ' + label + ' — ' + reqId,
    '<h2 style="color:' + color + ';margin:0 0 10px;">Request ' + label + '</h2><table style="border-collapse:collapse;margin:12px 0;">' +
    emailRow_('ID', reqId) + emailRow_('Record', recordInfo) + emailRow_('Pengaju', requesterName) + emailRow_('Oleh', approverName, true) +
    (reason ? emailRow_('Alasan', reason) : '') + '</table>');
}
