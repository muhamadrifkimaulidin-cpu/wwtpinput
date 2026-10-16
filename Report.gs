/**
 * ============================================
 * Report.gs v7.0 — WWTP Report (Clean Data)
 * Buat sheet baru di spreadsheet dengan data
 * lengkap, tanpa garis tepi, mudah dianalisa
 * ============================================
 */

function generateReport(opts) {
  try {
    var startDate = opts.startDate;
    var endDate = opts.endDate;
    var generatedBy = opts.generatedBy || 'Unknown';
    var mode = opts.mode || 'periode';

    var ss = SpreadsheetApp.getActiveSpreadsheet(), tz = Session.getScriptTimeZone();
    var sDate = new Date(startDate); sDate.setHours(0,0,0,0);
    var eDate = new Date(endDate); eDate.setHours(23,59,59,999);

    // Validate dates
    if (isNaN(sDate.getTime()) || isNaN(eDate.getTime())) {
      return { success: false, message: 'Tanggal tidak valid: ' + startDate + ' — ' + endDate };
    }

    var ops = collectOps_(ss, sDate, eDate, tz);
    var qual = collectQual_(ss, sDate, eDate, tz);

    var sheetName = 'RPT_' + Utilities.formatDate(sDate,tz,'ddMM') + '-' + Utilities.formatDate(eDate,tz,'ddMMyy');
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
    var sh = ss.insertSheet(sheetName);

    var F = 'Arial', r = 1;

    // ===== HEADER INFO =====
    sh.getRange(r,1).setValue('LAPORAN OPERASIONAL WWTP').setFontFamily(F).setFontSize(12).setFontWeight('bold');
    r++;
    sh.getRange(r,1).setValue('Periode: ' + Utilities.formatDate(sDate,tz,'dd/MM/yyyy') + ' — ' + Utilities.formatDate(eDate,tz,'dd/MM/yyyy')).setFontFamily(F).setFontSize(9);
    sh.getRange(r,4).setValue('Oleh: ' + generatedBy).setFontFamily(F).setFontSize(9);
    sh.getRange(r,7).setValue('Tipe: ' + (mode==='harian'?'Per Hari':'Per Periode')).setFontFamily(F).setFontSize(9);
    sh.getRange(r,10).setValue('Generated: ' + Utilities.formatDate(new Date(),tz,'dd/MM/yyyy HH:mm')).setFontFamily(F).setFontSize(9);
    r += 2;

    // ===== SECTION A: DATA OPERASIONAL =====
    if (mode === 'harian') {
      r = writeOpsHarian_(sh, r, F, ops);
    } else {
      r = writeOpsPeriode_(sh, r, F, ops, tz);
    }

    r++; // spacer

    // ===== SECTION B: DATA KUALITAS AIR =====
    r = writeQuality_(sh, r, F, qual);

    // Auto-resize all columns
    for (var c = 1; c <= sh.getLastColumn(); c++) {
      try { sh.autoResizeColumn(c); } catch(e) {}
    }

    return { success: true, sheetName: sheetName };
  } catch (e) {
    Logger.log('generateReport: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Gagal: ' + e.message };
  }
}

// ============ SECTION LABEL ============

function label_(sh, r, text, F) {
  sh.getRange(r, 1).setValue(text).setFontFamily(F).setFontSize(10).setFontWeight('bold');
  return r + 1;
}

function writeHeader_(sh, r, headers, F) {
  sh.getRange(r, 1, 1, headers.length).setValues([headers])
    .setFontFamily(F).setFontSize(8).setFontWeight('bold')
    .setBackground('#f0f0f0');
  return r + 1;
}

function writeRows_(sh, r, data) {
  if (!data.length) return r;
  sh.getRange(r, 1, data.length, data[0].length).setValues(data).setFontSize(8);
  return r + data.length;
}

function writeSummary_(sh, r, vals, F) {
  sh.getRange(r, 1, 1, vals.length).setValues([vals])
    .setFontFamily(F).setFontSize(8).setFontWeight('bold').setBackground('#e8e8e8');
  return r + 1;
}

// ============ OPS HARIAN ============

function writeOpsHarian_(sh, r, F, ops) {
  var recs = ops.records;
  if (!recs.length) { sh.getRange(r,1).setValue('Tidak ada data operasional.'); return r+1; }
  var ck = function(v) { return (v===true||v==='true') ? 'Ya' : ''; };

  // Tabel 1: Semua data operasional per shift
  r = label_(sh, r, 'A. DATA OPERASIONAL (Per Shift)', F);

  var h = ['Tanggal','Shift','Operator',
    'Chk MG17-18','Chk Clarifier','Chk Fe Removal','Chk BSB',
    'Bak Sirk. Before','Bak Sirk. After','Bak Δ',
    'Sungai Before','Sungai After','Sungai Δ',
    'MG17-18 Before','MG17-18 After','MG Δ',
    'FM Softener','FM Flotation','FM MG16','FM PAM','FM PAM/shift',
    'FM Sm1','FM Sm2','FM Sm3','FM Sm4','FM Sm5','FM Sm6',
    'Hydrofloc (kg)','Polymer (g)',
    'Batch FP','Bongkar Lumpur',
    'Tgl Lab','Lab CLR','Lab MG','Lab BSir','Lab Dem','Lab PAM',
    'Lab Sm1','Lab Sm2','Lab Sm3','Lab Sm4','Lab Sm5','Lab Sm6',
    'Pekerjaan Tambahan','Catatan'];

  r = writeHeader_(sh, r, h, F);

  var data = recs.map(function(rc) {
    var bB=pf_(rc.clr_bak_sirkulasi_before), bA=pf_(rc.clr_bak_sirkulasi_after);
    var sB=pf_(rc.clr_sungai_before), sA=pf_(rc.clr_sungai_after);
    var mB=pf_(rc.clr_mg17_18_before), mA=pf_(rc.clr_mg17_18_after);
    return [
      rc._dd, rc.shift, rc.nama_operator||'',
      ck(rc.chk_mg17_18), ck(rc.chk_clarifier), ck(rc.chk_fe_removal), ck(rc.chk_bak_sirkulasi_besar),
      bB, bA, rd_(bA-bB),
      sB, sA, rd_(sA-sB),
      mB, mA, rd_(mA-mB),
      pf_(rc.fm_softener_lt3), pf_(rc.fm_flotation), pf_(rc.fm_mg16), pf_(rc.fm_pam), pf_(rc.fm_pam_shift),
      pf_(rc.fm_sm1), pf_(rc.fm_sm2), pf_(rc.fm_sm3), pf_(rc.fm_sm4), pf_(rc.fm_sm5), pf_(rc.fm_sm6),
      pf_(rc.chem_hydrofloc_kg), pf_(rc.chem_polymer_g),
      rc.filter_press_batch||'', rc.bongkar_lumpur||'',
      rc.lab_tanggal||'', ck(rc.lab_clarifier), ck(rc.lab_mg17_18), ck(rc.lab_bak_sirkulasi), ck(rc.lab_demin), ck(rc.lab_pam),
      ck(rc.lab_sm1), ck(rc.lab_sm2), ck(rc.lab_sm3), ck(rc.lab_sm4), ck(rc.lab_sm5), ck(rc.lab_sm6),
      rc.pekerjaan_tambahan||'', rc.catatan||''
    ];
  });

  r = writeRows_(sh, r, data);
  r++;
  return r;
}

// ============ OPS PERIODE ============

function writeOpsPeriode_(sh, r, F, ops, tz) {
  var recs = ops.records, days = Math.max(ops.totalDays, 1);
  var ck = function(v) { return (v===true||v==='true') ? 'Ya' : ''; };

  // A1: Ringkasan
  r = label_(sh, r, 'A1. RINGKASAN', F);
  r = writeHeader_(sh, r, ['Keterangan','Nilai'], F);
  r = writeRows_(sh, r, [
    ['Total Record', ops.totalRecords],
    ['Total Hari Operasi', ops.totalDays],
    ['Total Hydrofloc (kg)', rd_(ops.totalH)],
    ['Rata-rata Hydrofloc/Hari (kg)', rd_(ops.totalH/days)],
    ['Total Polymer (g)', rd_(ops.totalP)],
    ['Rata-rata Polymer/Hari (g)', rd_(ops.totalP/days)],
    ['Shift 1 Records', ops.shiftCount.s1],
    ['Shift 2 Records', ops.shiftCount.s2],
    ['Shift 3 Records', ops.shiftCount.s3],
    ['Jumlah Operator', ops.operatorCount]
  ]);
  r++;

  // A2: Clarifier per shift
  r = label_(sh, r, 'A2. DATA CLARIFIER (Per Shift)', F);
  r = writeHeader_(sh, r, ['Tanggal','Shift','Operator','Bak Before','Bak After','Bak Δ','Sungai Before','Sungai After','Sungai Δ','MG Before','MG After','MG Δ'], F);
  var clrD = recs.map(function(rc) {
    var bB=pf_(rc.clr_bak_sirkulasi_before),bA=pf_(rc.clr_bak_sirkulasi_after);
    var sB=pf_(rc.clr_sungai_before),sA=pf_(rc.clr_sungai_after);
    var mB=pf_(rc.clr_mg17_18_before),mA=pf_(rc.clr_mg17_18_after);
    return [rc._dd, rc.shift, rc.nama_operator||'', bB, bA, rd_(bA-bB), sB, sA, rd_(sA-sB), mB, mA, rd_(mA-mB)];
  });
  r = writeRows_(sh, r, clrD);
  r++;

  // A3: Clarifier daily
  r = label_(sh, r, 'A3. CLARIFIER HARIAN (m³)', F);
  var clrAgg = {};
  recs.forEach(function(rc) {
    var d = rc._ds; if (!clrAgg[d]) clrAgg[d] = {date:rc._dd, bak:0, sungai:0, mg:0};
    var bB=pf_(rc.clr_bak_sirkulasi_before),bA=pf_(rc.clr_bak_sirkulasi_after);
    var sB=pf_(rc.clr_sungai_before),sA=pf_(rc.clr_sungai_after);
    var mB=pf_(rc.clr_mg17_18_before),mA=pf_(rc.clr_mg17_18_after);
    if(bB>0&&bA>=bB) clrAgg[d].bak+=rd_(bA-bB);
    if(sB>0&&sA>=sB) clrAgg[d].sungai+=rd_(sA-sB);
    if(mB>0&&mA>=mB) clrAgg[d].mg+=rd_(mA-mB);
  });
  r = writeHeader_(sh, r, ['Tanggal','Bak Sirkulasi','Sungai','MG 17-18','Total'], F);
  var clrDay=[], clrT=[0,0,0,0];
  Object.keys(clrAgg).sort().forEach(function(d) {
    var c=clrAgg[d], tot=rd_(c.bak+c.sungai+c.mg);
    clrDay.push([c.date, c.bak, c.sungai, c.mg, tot]);
    clrT[0]+=c.bak; clrT[1]+=c.sungai; clrT[2]+=c.mg; clrT[3]+=tot;
  });
  r = writeRows_(sh, r, clrDay);
  var n = Math.max(clrDay.length, 1);
  r = writeSummary_(sh, r, ['TOTAL', rd_(clrT[0]), rd_(clrT[1]), rd_(clrT[2]), rd_(clrT[3])], F);
  r = writeSummary_(sh, r, ['RATA-RATA', rd_(clrT[0]/n), rd_(clrT[1]/n), rd_(clrT[2]/n), rd_(clrT[3]/n)], F);
  r++;

  // A4: Flow Meter per shift
  r = label_(sh, r, 'A4. FLOW METER (Per Shift)', F);
  r = writeHeader_(sh, r, ['Tanggal','Shift','Softener','Flotation','MG16','PAM','PAM/shift','Sm1','Sm2','Sm3','Sm4','Sm5','Sm6'], F);
  var fmD = recs.map(function(rc) {
    return [rc._dd, rc.shift, pf_(rc.fm_softener_lt3), pf_(rc.fm_flotation), pf_(rc.fm_mg16),
      pf_(rc.fm_pam), pf_(rc.fm_pam_shift),
      pf_(rc.fm_sm1), pf_(rc.fm_sm2), pf_(rc.fm_sm3), pf_(rc.fm_sm4), pf_(rc.fm_sm5), pf_(rc.fm_sm6)];
  });
  r = writeRows_(sh, r, fmD);
  r++;

  // A5: Chemical harian
  r = label_(sh, r, 'A5. CHEMICAL (Per Hari)', F);
  var chemAgg={}, chT=[0,0];
  recs.forEach(function(rc) { var d=rc._dd; if(!chemAgg[d]) chemAgg[d]={d:d,h:0,p:0}; chemAgg[d].h+=pf_(rc.chem_hydrofloc_kg); chemAgg[d].p+=pf_(rc.chem_polymer_g); });
  r = writeHeader_(sh, r, ['Tanggal','Hydrofloc (kg)','Polymer (g)'], F);
  var chemDay = [];
  sortDK_(chemAgg).forEach(function(d) { chemDay.push([d, rd_(chemAgg[d].h), rd_(chemAgg[d].p)]); chT[0]+=chemAgg[d].h; chT[1]+=chemAgg[d].p; });
  r = writeRows_(sh, r, chemDay);
  r = writeSummary_(sh, r, ['TOTAL', rd_(chT[0]), rd_(chT[1])], F);
  r++;

  // A6: Filter Press & Lab
  r = label_(sh, r, 'A6. FILTER PRESS & LAB SAMPLING', F);
  r = writeHeader_(sh, r, ['Tanggal','Shift','Tgl Lab','Batch FP','Bongkar','CLR','MG','BSir','Dem','PAM','Sm1','Sm2','Sm3','Sm4','Sm5','Sm6'], F);
  var fpD = recs.map(function(rc) {
    return [rc._dd, rc.shift, rc.lab_tanggal||'', rc.filter_press_batch||'', rc.bongkar_lumpur||'',
      ck(rc.lab_clarifier), ck(rc.lab_mg17_18), ck(rc.lab_bak_sirkulasi), ck(rc.lab_demin), ck(rc.lab_pam),
      ck(rc.lab_sm1), ck(rc.lab_sm2), ck(rc.lab_sm3), ck(rc.lab_sm4), ck(rc.lab_sm5), ck(rc.lab_sm6)];
  });
  r = writeRows_(sh, r, fpD);
  r++;

  // A7: Checklist compliance
  r = label_(sh, r, 'A7. CHECKLIST PROSES', F);
  var tot = Math.max(ops.totalRecords, 1);
  r = writeHeader_(sh, r, ['Item','Jumlah','Total Shift','%'], F);
  r = writeRows_(sh, r, ops.chkProcess.map(function(c) { return [c.label, c.count, ops.totalRecords, Math.round(c.count/tot*100)+'%']; }));
  r++;

  r = label_(sh, r, 'A8. LAB SAMPLING COMPLIANCE', F);
  r = writeHeader_(sh, r, ['Item','Jumlah','Total Shift','%'], F);
  r = writeRows_(sh, r, ops.chkLab.map(function(c) { return [c.label, c.count, ops.totalRecords, Math.round(c.count/tot*100)+'%']; }));
  r++;

  // A9: Shift coverage
  r = label_(sh, r, 'A9. COVERAGE SHIFT', F);
  var shMap = {};
  recs.forEach(function(rc) {
    var d=rc._dd; if(!shMap[d]) shMap[d]={d:d, s1:0, s2:0, s3:0, ops:{}};
    var s=parseInt(rc.shift); if(s===1)shMap[d].s1++;else if(s===2)shMap[d].s2++;else if(s===3)shMap[d].s3++;
    if(rc.nama_operator) shMap[d].ops[rc.nama_operator]=true;
  });
  r = writeHeader_(sh, r, ['Tanggal','Shift 1','Shift 2','Shift 3','Total','Operator'], F);
  var shD = [];
  sortDK_(shMap).forEach(function(d) { var v=shMap[d]; shD.push([d, v.s1, v.s2, v.s3, v.s1+v.s2+v.s3, Object.keys(v.ops).join(', ')]); });
  r = writeRows_(sh, r, shD);
  r++;

  return r;
}

// ============ QUALITY ============

function writeQuality_(sh, r, F, qual) {
  if (!qual || !qual.length) return r;

  var sources = [
    { pfx:'clr_', lbl:'Clarifier', k:['ph','conductivity','tss','tds'] },
    { pfx:'bsb_', lbl:'Bak Sirkulasi Besar', k:['ph','conductivity','tss','tds'] },
    { pfx:'bs17_', lbl:'Bak Sirkulasi 17-18', k:['ph','conductivity','tss','tds'] },
    { pfx:'sm1_', lbl:'Sumur 1', k:['ph','conductivity','hardness'] },
    { pfx:'sm2_', lbl:'Sumur 2', k:['ph','conductivity','hardness'] },
    { pfx:'sm3_', lbl:'Sumur 3', k:['ph','conductivity','hardness'] },
    { pfx:'sm4_', lbl:'Sumur 4', k:['ph','conductivity','hardness'] },
    { pfx:'sm5_', lbl:'Sumur 5', k:['ph','conductivity','hardness'] },
    { pfx:'sm6_', lbl:'Sumur 6', k:['ph','conductivity','hardness'] },
    { pfx:'pam_', lbl:'PAM', k:['ph','conductivity','hardness'] },
    { pfx:'sft_', lbl:'Softener', k:['ph','conductivity','hardness'] }
  ];

  var kl = function(k) {
    if(k==='ph') return 'pH';
    if(k==='conductivity') return 'Conductivity (µS/cm)';
    if(k==='tss') return 'TSS (ppm)';
    if(k==='tds') return 'TDS (ppm)';
    if(k==='hardness') return 'Hardness (ppm)';
    return k;
  };

  sources.forEach(function(src) {
    var hdrs = ['Tanggal'].concat(src.k.map(kl));
    var rows = [], sums = [], cnts = [];
    for (var j = 0; j < src.k.length; j++) { sums.push(0); cnts.push(0); }

    qual.forEach(function(q) {
      var vals = src.k.map(function(k) { return pv_(q[src.pfx + k]); });
      if (!vals.some(function(v) { return v !== '-'; })) return;
      var numVals = vals.map(function(v) { return v === '-' ? '' : parseFloat(v); });
      rows.push([q._dd].concat(numVals));
      for (var j = 0; j < src.k.length; j++) {
        if (vals[j] !== '-') { sums[j] += parseFloat(vals[j]); cnts[j]++; }
      }
    });

    if (!rows.length) return;
    r = label_(sh, r, 'B. MUTU ' + src.lbl.toUpperCase(), F);
    r = writeHeader_(sh, r, hdrs, F);
    r = writeRows_(sh, r, rows);
    r = writeSummary_(sh, r, ['RATA-RATA'].concat(sums.map(function(s,i) { return cnts[i]>0 ? rd_(s/cnts[i]) : ''; })), F);
    r++;
  });

  return r;
}

// ============ EXPORT TO DEVICE ============

function exportReportFile(opts) {
  try {
    var startDate = opts.startDate;
    var endDate = opts.endDate;
    var generatedBy = opts.generatedBy || 'Unknown';
    var format = opts.format || 'xlsx';
    var mode = opts.mode || 'periode';

    if (format === 'csv') return exportCsv_(startDate, endDate);

    var result = generateReport({startDate:startDate, endDate:endDate, generatedBy:generatedBy, mode:mode});
    if (!result.success) return result;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(result.sheetName);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    var gid = sheet.getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format='
      + (format === 'pdf' ? 'pdf' : 'xlsx') + '&gid=' + gid;
    if (format === 'pdf') url += '&portrait=false&size=A3&gridlines=true&fzr=true&fitw=true&sheetnames=false&printtitle=false&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3';

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      return { success: false, message: 'Export HTTP ' + resp.getResponseCode() };
    }

    var b64 = Utilities.base64Encode(resp.getBlob().getBytes());

    // Untuk export: hapus sheet setelah download
    // Untuk sheet mode: biarkan sheet tetap ada
    if (format !== 'sheet') {
      ss.deleteSheet(sheet);
    }

    var ext = format === 'pdf' ? '.pdf' : '.xlsx';
    var mime = format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    return { success: true, data: b64, filename: result.sheetName + ext, mimeType: mime };
  } catch (e) {
    Logger.log('exportReportFile: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Export gagal: ' + e.message };
  }
}

function exportCsv_(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), tz = Session.getScriptTimeZone();
    var sD = new Date(startDate); sD.setHours(0,0,0,0);
    var eD = new Date(endDate); eD.setHours(23,59,59,999);
    var sheet = ss.getSheetByName('DataOps');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Tidak ada data.' };
    var hdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var all = sheet.getDataRange().getValues();
    var ci = {}; for (var h = 0; h < hdr.length; h++) ci[hdr[h]] = h;
    var rows = [hdr.join(',')];
    for (var i = 1; i < all.length; i++) {
      var dt = all[i][ci['tanggal']];
      if (dt instanceof Date) { if (dt < sD || dt > eD) continue; }
      else { var ds = String(dt); if (!ds.match(/^\d{4}-\d{2}-\d{2}/)) continue; if (ds < Utilities.formatDate(sD,tz,'yyyy-MM-dd') || ds > Utilities.formatDate(eD,tz,'yyyy-MM-dd')) continue; }
      var cells = hdr.map(function(_,j) {
        var v = all[i][j]; if (v instanceof Date) v = Utilities.formatDate(v,tz,'yyyy-MM-dd HH:mm:ss');
        v = String(v===null||v===undefined?'':v);
        if (v.indexOf(',')>=0||v.indexOf('\n')>=0||v.indexOf('"')>=0) v='"'+v.replace(/"/g,'""')+'"';
        return v;
      });
      rows.push(cells.join(','));
    }
    var nm = 'RPT_' + Utilities.formatDate(sD,tz,'ddMM') + '-' + Utilities.formatDate(eD,tz,'ddMMyy');
    return { success: true, data: Utilities.base64Encode(rows.join('\n'), Utilities.Charset.UTF_8), filename: nm+'.csv', mimeType: 'text/csv' };
  } catch (e) { return { success: false, message: 'CSV gagal: ' + e.message }; }
}

// ============ DATA COLLECTORS ============

function collectOps_(ss, sDate, eDate, tz) {
  var sheet = ss.getSheetByName('DataOps');
  if (!sheet || sheet.getLastRow() <= 1) return { records:[], totalRecords:0, totalDays:0, totalH:0, totalP:0, shiftCount:{s1:0,s2:0,s3:0}, operatorCount:0, chkProcess:[], chkLab:[] };
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0], all = sheet.getDataRange().getValues();
  var ci = {}; for (var h = 0; h < headers.length; h++) ci[headers[h]] = h;
  var recs=[], days={}, ops={}, tH=0, tP=0, sc={s1:0,s2:0,s3:0};
  for (var i = 1; i < all.length; i++) {
    var dt = pd_(all[i][ci['tanggal']]); if (!dt || dt < sDate || dt > eDate) continue;
    var obj = {}; for (var j = 0; j < headers.length; j++) obj[headers[j]] = all[i][j];
    obj._ds = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    obj._dd = Utilities.formatDate(dt, tz, 'dd/MM/yyyy');
    recs.push(obj); days[obj._ds] = true;
    if (obj.nama_operator) ops[obj.nama_operator] = true;
    tH += pf_(obj.chem_hydrofloc_kg); tP += pf_(obj.chem_polymer_g);
    var s = parseInt(obj.shift); if(s===1)sc.s1++;else if(s===2)sc.s2++;else if(s===3)sc.s3++;
  }
  var pi=[{key:'chk_mg17_18',label:'MG 17-18'},{key:'chk_clarifier',label:'Clarifier'},{key:'chk_fe_removal',label:'Fe Removal'},{key:'chk_bak_sirkulasi_besar',label:'Bak Sirk. Besar'}];
  var li=[{key:'lab_clarifier',label:'Clarifier'},{key:'lab_mg17_18',label:'MG 17-18'},{key:'lab_bak_sirkulasi',label:'Bak Sirkulasi'},{key:'lab_demin',label:'Demin'},{key:'lab_pam',label:'PAM'},{key:'lab_sm1',label:'Sm 1'},{key:'lab_sm2',label:'Sm 2'},{key:'lab_sm3',label:'Sm 3'},{key:'lab_sm4',label:'Sm 4'},{key:'lab_sm5',label:'Sm 5'},{key:'lab_sm6',label:'Sm 6'}];
  function cc(items){return items.map(function(c){var n=0;recs.forEach(function(r){if(r[c.key]===true||r[c.key]==='true')n++;});return{label:c.label,count:n};});}
  return{records:recs,totalRecords:recs.length,totalDays:Object.keys(days).length,totalH:tH,totalP:tP,shiftCount:sc,operatorCount:Object.keys(ops).length,chkProcess:cc(pi),chkLab:cc(li)};
}

function collectQual_(ss, sDate, eDate, tz) {
  var sheet = ss.getSheetByName('DataQuality');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0], all = sheet.getDataRange().getValues();
  var ci = {}; for (var h = 0; h < headers.length; h++) ci[headers[h]] = h;
  var recs = [];
  for (var i = 1; i < all.length; i++) {
    var rawT = all[i][ci['tanggal']];
    var dt = pdr_(rawT);
    if (dt) { if (dt < sDate || dt > eDate) continue; }
    var obj = {}; for (var j = 0; j < headers.length; j++) obj[headers[j]] = all[i][j];
    obj._ds = dt ? Utilities.formatDate(dt, tz, 'yyyy-MM-dd') : String(rawT);
    obj._dd = dt ? Utilities.formatDate(dt, tz, 'dd/MM/yyyy') : String(rawT);
    recs.push(obj);
  }
  recs.sort(function(a,b) { return String(a._ds).localeCompare(String(b._ds)); });
  return recs;
}

// ============ UTILITY ============

function sortDK_(m) { return Object.keys(m).sort(function(a,b) { var pa=a.split('/'),pb=b.split('/'); return (pa[2]+pa[1]+pa[0]).localeCompare(pb[2]+pb[1]+pb[0]); }); }
function pd_(v) { if(v instanceof Date) return v; if(typeof v==='string') { if(v.match(/^\d{4}-\d{2}-\d{2}/)) return new Date(v.substring(0,10)+'T00:00:00'); if(v.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)){var p=v.split('/'); return new Date(p[2]+'-'+('0'+p[1]).slice(-2)+'-'+('0'+p[0]).slice(-2)+'T00:00:00');}} return null; }
function pf_(v) { return parseFloat(v) || 0; }
function pv_(v) { if(v===null||v===undefined||v==='') return '-'; var n=(typeof v==='number')?v:parseFloat(String(v).replace(',','.')); return isNaN(n)||n<=0?'-':rd_(n); }
function rd_(v) { return Math.round(v*100)/100; }
function pdr_(v) {
  if(!v) return null; if(v instanceof Date && !isNaN(v.getTime())) return v;
  if(typeof v==='number'){if(v>10000&&v<100000){var d=new Date((v-25569)*86400*1000);return isNaN(d.getTime())?null:d;}return null;}
  var s=String(v).trim(); if(!s) return null;
  if(s.match(/^\d{4}-\d{2}-\d{2}/)) return new Date(s.substring(0,10)+'T00:00:00');
  if(s.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)){var p=s.split('/');return new Date(p[2]+'-'+('0'+p[1]).slice(-2)+'-'+('0'+p[0]).slice(-2)+'T00:00:00');}
  if(s.match(/^\d{1,2}-\d{1,2}-\d{4}$/)){var p2=s.split('-');return new Date(p2[2]+'-'+('0'+p2[1]).slice(-2)+'-'+('0'+p2[0]).slice(-2)+'T00:00:00');}
  var attempt=new Date(s); return isNaN(attempt.getTime())?null:attempt;
}
