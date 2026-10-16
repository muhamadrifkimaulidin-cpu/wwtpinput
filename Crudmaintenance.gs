/**
 * ============================================
 * CRUDMaintenance.gs — CRUD Maintenance Checklist
 * Sheet: DataMaintenance
 * FIXED: Auto-create sheet, timestamp ID
 * ============================================
 */

function getMaintenanceHeaders_() {
  return [
    'record_id','timestamp','tanggal','shift','nama_operator','user_id',
    'mnt_clr_tangki','mnt_clr_tangki_flokulan','mnt_clr_tangki_floagulan','mnt_clr_tangki_chemical',
    'mnt_clr_pompa_clarifier','mnt_clr_keseluruhan_proses','mnt_clr_pompa_siplai_leaching',
    'mnt_clr_filter_press_1','mnt_clr_filter_press_2','mnt_clr_kolam',
    'mnt_clr_fungsi_proses','mnt_clr_5r_area',
    'mnt_sft_tangki_air_baku','mnt_sft_tangki_sand_filter','mnt_sft_tangki_garam',
    'mnt_sft_tangki_softener','mnt_sft_tangki_hasil','mnt_sft_valve_piping',
    'mnt_sft_fungsi_proses','mnt_sft_5r_area',
    'mnt_catatan'
  ];
}

/**
 * FIX: Uses ensureProsesSH_ (from CRUDProses.gs) to auto-create sheet
 */
function createMaintenance(data) {
  try {
    var info = ensureProsesSH_('DataMaintenance', getMaintenanceHeaders_());
    var recordId = genProsesId_('MNT-', info.sheet);
    data.record_id = recordId;
    data.timestamp = new Date().toISOString();
    info.sheet.appendRow(sanitizeProsesRow_(info.headers, data));
    return { success: true, message: 'Data Maintenance berhasil disimpan.', record_id: recordId };
  } catch (e) {
    Logger.log('createMaintenance error: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Gagal: ' + e.message };
  }
}

function readMaintenance(options) {
  return readProsesGeneric_('DataMaintenance', options);
}
