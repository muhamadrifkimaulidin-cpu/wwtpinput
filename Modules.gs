/**
 * ============================================
 * Modules.gs — Registry Modul Input
 * Satu definisi kolom + label untuk semua modul:
 * Operasional, FIK Clarifier, Softener, Maintenance, Pressure Gauge
 * ============================================
 */

var MODULE_DEFS_ = null;
var FIK_WAKTU_ = ['07', '12', '15', '19', '23', '03'];
var FIK_SUFFIX_ = [
  ['rotameter', 'Rotameter'], ['flowmeter', 'Flowmeter'],
  ['stroke_naoh', 'Stroke NaOH'], ['stroke_hidroflok', 'Stroke Hidroflok'], ['stroke_superflok', 'Stroke Superflok'],
  ['ph', 'pH'], ['visual_flokulasi', 'Visual Flokulasi'], ['visual_air', 'Visual Air'], ['batch_lumpur', 'Batch Lumpur']
];
var PG_SUMUR_ = [1, 2, 3, 4, 5, 6];

function baseCols_() {
  return [['record_id', 'Record ID'], ['timestamp', 'Timestamp'], ['tanggal', 'Tanggal'],
    ['shift', 'Shift'], ['nama_operator', 'Operator'], ['user_id', 'User ID']];
}

function getModules_() {
  if (MODULE_DEFS_) return MODULE_DEFS_;

  var ops = [
    ['chk_mg17_18', 'Chk MG 17-18'], ['chk_clarifier', 'Chk Clarifier'], ['chk_fe_removal', 'Chk Fe Removal'], ['chk_bak_sirkulasi_besar', 'Chk Bak Sirk. Besar'],
    ['clr_bak_sirkulasi_before', 'Bak Sirk. Sebelum'], ['clr_bak_sirkulasi_after', 'Bak Sirk. Sesudah'],
    ['clr_sungai_before', 'Sungai Sebelum'], ['clr_sungai_after', 'Sungai Sesudah'],
    ['clr_mg17_18_before', 'MG 17-18 Sebelum'], ['clr_mg17_18_after', 'MG 17-18 Sesudah'],
    ['fm_softener_lt3', 'FM Softener lt.3'], ['fm_flotation', 'FM Flotation'], ['fm_mg16', 'FM MG 16'],
    ['fm_pam', 'FM PAM Awal'], ['fm_pam_shift', 'FM PAM Akhir'],
    ['fm_sm1', 'FM Sm 1'], ['fm_sm2', 'FM Sm 2'], ['fm_sm3', 'FM Sm 3'], ['fm_sm4', 'FM Sm 4'], ['fm_sm5', 'FM Sm 5'], ['fm_sm6', 'FM Sm 6'],
    ['filter_press_batch', 'Batch Filter Press'], ['bongkar_lumpur', 'Bongkar Lumpur'],
    ['lab_clarifier', 'Lab Clarifier'], ['lab_mg17_18', 'Lab MG 17-18'], ['lab_bak_sirkulasi', 'Lab Bak Sirkulasi'],
    ['lab_demin', 'Lab Softener (Demin)'], ['lab_pam', 'Lab PAM'],
    ['lab_sm1', 'Lab Sm 1'], ['lab_sm2', 'Lab Sm 2'], ['lab_sm3', 'Lab Sm 3'], ['lab_sm4', 'Lab Sm 4'], ['lab_sm5', 'Lab Sm 5'], ['lab_sm6', 'Lab Sm 6'],
    ['lab_tanggal', 'Tgl Kirim Sample'],
    ['chem_hydrofloc_kg', 'Hydrofloc (kg)'], ['chem_polymer_g', 'Polymer (g)'],
    ['pekerjaan_tambahan', 'Pekerjaan Tambahan'], ['catatan', 'Catatan'], ['fm_leaching', 'FM Leaching (m³)']
  ];

  var fik = [
    ['prep_naoh_flake_kg', 'NaOH Flake (kg)'], ['prep_hidroflok_kg', 'Hidroflok (kg)'], ['prep_superflok_g', 'Superflok (g)'],
    ['prep_air_naoh_l', 'Air NaOH (L)'], ['prep_air_hidroflok_l', 'Air Hidroflok (L)'], ['prep_air_superflok_l', 'Air Superflok (L)'],
    ['prep_pompa_ok', 'Pompa OK'], ['prep_tangki_ok', 'Tangki/Pipa/Valve OK'], ['prep_keterangan', 'Keterangan Abnormal']
  ];
  FIK_WAKTU_.forEach(function(w) {
    FIK_SUFFIX_.forEach(function(s) { fik.push(['p' + w + '_' + s[0], w + ':00 ' + s[1]]); });
  });
  fik.push(['catatan', 'Catatan']);

  var sft = [['regenerasi', 'Regenerasi'], ['garam_kg', 'Garam (kg)'], ['catatan', 'Catatan']];

  var mnt = [
    ['mnt_clr_tangki', 'CLR Tangki'], ['mnt_clr_tangki_flokulan', 'CLR Tangki Flokulan'], ['mnt_clr_tangki_floagulan', 'CLR Tangki Floagulan'],
    ['mnt_clr_tangki_chemical', 'CLR Tangki Chemical'], ['mnt_clr_pompa_clarifier', 'CLR Pompa Clarifier'], ['mnt_clr_keseluruhan_proses', 'CLR Keseluruhan Proses'],
    ['mnt_clr_pompa_siplai_leaching', 'CLR Pompa Siplai Leaching'], ['mnt_clr_filter_press_1', 'CLR Filter Press 1'], ['mnt_clr_filter_press_2', 'CLR Filter Press 2'],
    ['mnt_clr_kolam', 'CLR Kolam'], ['mnt_clr_fungsi_proses', 'CLR Fungsi Proses'], ['mnt_clr_5r_area', 'CLR 5R Area'],
    ['mnt_sft_tangki_air_baku', 'SFT Tangki Air Baku'], ['mnt_sft_tangki_sand_filter', 'SFT Tangki Sand Filter'], ['mnt_sft_tangki_garam', 'SFT Tangki Garam'],
    ['mnt_sft_tangki_softener', 'SFT Tangki Softener'], ['mnt_sft_tangki_hasil', 'SFT Tangki Hasil'], ['mnt_sft_valve_piping', 'SFT Valve & Piping'],
    ['mnt_sft_fungsi_proses', 'SFT Fungsi Proses'], ['mnt_sft_5r_area', 'SFT 5R Area'],
    ['mnt_catatan', 'Catatan']
  ];

  var pg = [];
  PG_SUMUR_.forEach(function(n) {
    pg.push(['pg_sm' + n + '_pressure', 'Sm ' + n + ' Tekanan']);
    pg.push(['pg_sm' + n + '_kondisi', 'Sm ' + n + ' Kondisi']);
  });
  pg.push(['pg_catatan', 'Catatan']);

  MODULE_DEFS_ = {
    ops: mkModule_('ops', 'DataOps', 'REC-', 'Operasional', true, ops),
    fik: mkModule_('fik', 'DataProsesClarifier', 'FIK-', 'FIK Clarifier', false, fik),
    sft: mkModule_('sft', 'DataProsesSoftener', 'SFT-', 'Softener', false, sft),
    mnt: mkModule_('mnt', 'DataMaintenance', 'MNT-', 'Maintenance', false, mnt),
    pg: mkModule_('pg', 'DataPressureGauge', 'PG-', 'Pressure Gauge', true, pg)
  };
  return MODULE_DEFS_;
}

function mkModule_(key, sheet, prefix, title, uniquePerShift, cols) {
  var all = baseCols_().concat(cols);
  var keys = [], labels = {};
  all.forEach(function(c) { keys.push(c[0]); labels[c[0]] = c[1]; });
  return { key: key, sheet: sheet, prefix: prefix, title: title, uniquePerShift: uniquePerShift, keys: keys, labels: labels };
}

function getModule_(key) {
  var m = getModules_()[key];
  if (!m) throw new Error('Modul tidak dikenal: ' + key);
  return m;
}

function ensureModuleSheet_(mod) { return ensureSheet_(mod.sheet, mod.keys); }
