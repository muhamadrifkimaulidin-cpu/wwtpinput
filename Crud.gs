/**
 * ============================================
 * CRUD.gs — Create & Read untuk semua modul
 * Edit/hapus hanya lewat Request + approval (Requests.gs)
 * ============================================
 */

/**
 * CREATE — simpan entri baru. Operator & user_id diambil dari sesi server.
 * Modul uniquePerShift (Operasional, Pressure Gauge): 1 entri per tanggal+shift.
 */
function createEntry(token, moduleKey, data) {
  return api_(token, { write: true }, function(user) {
    var mod = getModule_(moduleKey);
    var clean = sanitizeEntry_(mod, data);
    if (!clean.tanggal) return { success: false, message: 'Tanggal tidak valid.' };
    if (!clean.shift) return { success: false, message: 'Shift tidak valid.' };
    clean.nama_operator = user.nama;
    clean.user_id = user.user_id;

    var res = withLock_(function() {
      var info = ensureModuleSheet_(mod);
      if (mod.uniquePerShift) {
        var all = info.sheet.getDataRange().getValues();
        var dup = findShiftDuplicate_(all, clean.tanggal, clean.shift, null);
        if (dup) return { success: false, message: 'Data ' + mod.title + ' tanggal ' + clean.tanggal + ' Shift ' + clean.shift + ' sudah diinput oleh ' + dup + '. Tidak bisa double input di shift yang sama.' };
      }
      clean.record_id = genId_(mod.prefix);
      clean.timestamp = new Date().toISOString();
      info.sheet.appendRow(rowFromObj_(info.headers, clean));
      bumpVer_('rec_' + mod.key);
      return { success: true, message: 'Data ' + mod.title + ' berhasil disimpan.', record_id: clean.record_id };
    });

    if (res.success && mod.key === 'ops') {
      var n = checkNegativeDiffs_(clean);
      if (n > 0) res.message += '\n⚠ ' + n + ' selisih negatif terdeteksi — notifikasi dikirim ke approver.';
    }
    return res;
  });
}

/**
 * Hanya kolom yang dikenal modul; kolom sistem dibuang; teks diamankan.
 * Nilai kosong disimpan kosong (bukan 0).
 */
function sanitizeEntry_(mod, data) {
  var out = {};
  data = data || {};
  mod.keys.forEach(function(k) {
    if (SYSTEM_KEYS_.indexOf(k) >= 0 || !data.hasOwnProperty(k)) return;
    var v = data[k];
    if (/tanggal$/.test(k)) v = toYmd_(v);
    else if (k === 'shift') { v = parseInt(v, 10); if ([1, 2, 3].indexOf(v) < 0) v = ''; }
    else if (typeof v === 'boolean') { /* keep */ }
    else if (typeof v === 'number') { if (!isFinite(v)) v = ''; }
    else if (v === null || v === undefined) v = '';
    else v = safeCell_(String(v).trim());
    out[k] = v;
  });
  return out;
}

/** Cari entri lain dengan tanggal+shift sama. Return nama operator atau null. */
function findShiftDuplicate_(all, tanggal, shift, excludeId) {
  if (all.length <= 1) return null;
  var ci = colIndex_(all[0].map(String));
  for (var i = 1; i < all.length; i++) {
    if (excludeId && String(all[i][ci.record_id]) === String(excludeId)) continue;
    if (toYmd_(all[i][ci.tanggal]) === tanggal && String(all[i][ci.shift]) === String(shift)) {
      return String(all[i][ci.nama_operator] || '-').trim();
    }
  }
  return null;
}

/**
 * READ — daftar entri dengan filter, search, pagination.
 * Diurutkan tanggal terbaru → shift → waktu input. Ditandai jika ada request pending.
 */
function listEntries(token, moduleKey, options) {
  return api_(token, {}, function(user) {
    var mod = getModule_(moduleKey);
    var o = options || {};
    var page = Math.max(1, parseInt(o.page, 10) || 1);
    var limit = Math.min(50, Math.max(5, parseInt(o.limit, 10) || 15));
    var f = { shift: o.shift ? String(o.shift) : '', start: toYmd_(o.startDate), end: toYmd_(o.endDate), q: String(o.search || '').trim().toLowerCase() };

    var key = cacheKey_(['list', mod.key, ver_('rec_' + mod.key), ver_('req'), page, limit, f.shift, f.start, f.end, f.q]);
    var cached = cacheGet_(key);
    if (cached) return cached;

    var sheet = getSheet_(mod.sheet);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, records: [], total: 0, page: 1, totalPages: 0 };
    var all = sheet.getDataRange().getValues();
    var headers = all[0].map(String);
    var pending = getPendingMap_(mod.key);

    var recs = [];
    for (var i = 1; i < all.length; i++) {
      var r = rowToObj_(headers, all[i]);
      if (!r.record_id) continue;
      r.tanggal = toYmd_(r.tanggal) || String(r.tanggal);
      if (f.shift && String(r.shift) !== f.shift) continue;
      if (f.start && r.tanggal < f.start) continue;
      if (f.end && r.tanggal > f.end) continue;
      if (f.q && JSON.stringify(r).toLowerCase().indexOf(f.q) < 0) continue;
      if (pending[r.record_id]) r._pending = pending[r.record_id];
      recs.push(r);
    }
    recs.sort(function(a, b) {
      if (a.tanggal !== b.tanggal) return a.tanggal < b.tanggal ? 1 : -1;
      var sa = parseInt(a.shift, 10) || 0, sb = parseInt(b.shift, 10) || 0;
      if (sa !== sb) return sb - sa;
      return String(b.timestamp).localeCompare(String(a.timestamp));
    });

    var total = recs.length, totalPages = Math.ceil(total / limit);
    if (page > totalPages && totalPages > 0) page = totalPages;
    var start = (page - 1) * limit;
    var res = { success: true, records: recs.slice(start, start + limit), total: total, page: page, totalPages: totalPages };
    cachePut_(key, res, CACHE_TTL_SHORT_);
    return res;
  });
}

function getEntry(token, moduleKey, recordId) {
  return api_(token, {}, function(user) {
    var mod = getModule_(moduleKey);
    var found = findRecord_(mod, recordId);
    if (!found) return { success: false, message: 'Record tidak ditemukan.' };
    var rec = found.obj;
    var pending = getPendingMap_(mod.key)[rec.record_id];
    if (pending) rec._pending = pending;
    return { success: true, record: rec };
  });
}

/** Cari record di sheet modul. Return { sheet, all, headers, index, obj } atau null. */
function findRecord_(mod, recordId) {
  var sheet = getSheet_(mod.sheet);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var all = sheet.getDataRange().getValues();
  var headers = all[0].map(String);
  var col = headers.indexOf('record_id');
  if (col < 0) col = 0;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][col]) === String(recordId)) {
      return { sheet: sheet, all: all, headers: headers, index: i, obj: rowToObj_(headers, all[i]) };
    }
  }
  return null;
}

// ============ LAST SHIFT DATA (Operasional) ============
function getLastShiftData(token, tanggal, shift) {
  return api_(token, {}, function(user) {
    tanggal = toYmd_(tanggal);
    if (!tanggal) return { success: true, data: null };
    var shiftNum = parseInt(shift, 10) || 1;
    var key = cacheKey_(['lastshift', ver_('rec_ops'), tanggal, shiftNum]);
    var cached = cacheGet_(key);
    if (cached) return cached;

    var sheet = getSheet_('DataOps');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: null };
    var all = sheet.getDataRange().getValues();
    var headers = all[0].map(String);
    var ci = colIndex_(headers);

    var prevTanggal, prevShift;
    if (shiftNum === 1) { prevTanggal = addDaysYmd_(tanggal, -1); prevShift = '3'; }
    else if (shiftNum === 2) { prevTanggal = tanggal; prevShift = '1'; }
    else { prevTanggal = tanggal; prevShift = '2'; }

    var found = null;
    for (var i = all.length - 1; i >= 1; i--) {
      if (toYmd_(all[i][ci.tanggal]) === prevTanggal && String(all[i][ci.shift]) === prevShift) {
        found = rowToObj_(headers, all[i]);
        break;
      }
    }
    var res = { success: true, data: null };
    if (found) {
      var pf = function(v) { return parseFloat(v) || 0; };
      res.data = {
        clr_bak_sirkulasi: pf(found.clr_bak_sirkulasi_after), clr_sungai: pf(found.clr_sungai_after), clr_mg17_18: pf(found.clr_mg17_18_after),
        fm_softener_lt3: pf(found.fm_softener_lt3), fm_flotation: pf(found.fm_flotation), fm_mg16: pf(found.fm_mg16),
        fm_pam: pf(found.fm_pam), fm_pam_shift: pf(found.fm_pam_shift),
        fm_sm1: pf(found.fm_sm1), fm_sm2: pf(found.fm_sm2), fm_sm3: pf(found.fm_sm3),
        fm_sm4: pf(found.fm_sm4), fm_sm5: pf(found.fm_sm5), fm_sm6: pf(found.fm_sm6),
        fm_leaching: pf(found.fm_leaching),
        prev_shift: prevShift, prev_tanggal: prevTanggal, prev_operator: found.nama_operator || ''
      };
    }
    cachePut_(key, res, CACHE_TTL_SHORT_);
    return res;
  });
}
