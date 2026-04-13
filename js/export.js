// export.js — Excel export via SheetJS (Community Edition, CDN)
// Exports: per-client XLSX (3 sheets) + all-clients XLSX

'use strict';

const SHEETJS_CDN = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';

let _xlsxLoaded = false;

// ── Load SheetJS lazily ───────────────────────────────────────
function loadXLSX() {
  if (_xlsxLoaded && typeof XLSX !== 'undefined') return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script  = document.createElement('script');
    script.src    = SHEETJS_CDN;
    script.onload = () => { _xlsxLoaded = true; resolve(); };
    script.onerror = () => reject(new Error('Nu s-a putut încărca librăria Excel. Verificați conexiunea.'));
    document.head.appendChild(script);
  });
}

// ── Save file to chosen export folder (File System Access API) ────
// Stores directory handle in IndexedDB so user picks folder only once.
// Creates subfolder per client automatically.

var _exportDirHandle = null; // cached FileSystemDirectoryHandle

/** Get stored export directory handle from IndexedDB */
async function _getExportDirHandle() {
  if (_exportDirHandle) return _exportDirHandle;
  try {
    var stored = await getByKey('settings', 'export_dir_handle');
    if (stored && stored.value) {
      _exportDirHandle = stored.value;
      return _exportDirHandle;
    }
  } catch (e) {}
  return null;
}

/** Prompt user to pick export root folder (first time or reset) */
async function pickExportFolder() {
  if (typeof window.showDirectoryPicker !== 'function') {
    showToast('Browserul nu suportă alegerea folderului. Folosește Chrome/Edge pe desktop.', 'warning');
    return null;
  }
  try {
    var handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    _exportDirHandle = handle;
    await put('settings', { key: 'export_dir_handle', value: handle });
    showToast('Folder export setat: ' + handle.name, 'success');
    return handle;
  } catch (e) {
    if (e.name === 'AbortError') return null; // user cancelled
    console.warn('[EXPORT] Directory picker failed:', e.message);
    return null;
  }
}

/** Verify we still have permission to write to the stored directory */
async function _verifyDirPermission(handle) {
  if (!handle) return false;
  try {
    var perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm === 'granted') return true;
    perm = await handle.requestPermission({ mode: 'readwrite' });
    return perm === 'granted';
  } catch (e) {
    return false;
  }
}

/** Write workbook to export folder, creating client subfolder if needed */
async function _writeFileWithPicker(wb, defaultName, clientName) {
  // Try stored directory handle first
  if (typeof window.showDirectoryPicker === 'function') {
    var dirHandle = await _getExportDirHandle();

    // First time: prompt to pick folder
    if (!dirHandle) {
      dirHandle = await pickExportFolder();
    }

    if (dirHandle) {
      // Verify permission (may need re-grant after browser restart)
      var hasPermission = await _verifyDirPermission(dirHandle);
      if (!hasPermission) {
        // Permission lost, re-prompt
        dirHandle = await pickExportFolder();
        hasPermission = dirHandle ? await _verifyDirPermission(dirHandle) : false;
      }

      if (dirHandle && hasPermission) {
        try {
          var targetDir = dirHandle;

          // If clientName provided, create/get client subfolder
          if (clientName) {
            var folderName = sanitizeFilename(clientName);
            targetDir = await dirHandle.getDirectoryHandle(folderName, { create: true });
          }

          // Write the file
          var fileHandle = await targetDir.getFileHandle(defaultName, { create: true });
          var writable = await fileHandle.createWritable();
          var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          await writable.write(new Uint8Array(buf));
          await writable.close();
          return true;
        } catch (e) {
          console.warn('[EXPORT] Directory write failed, falling back:', e.message);
          showToast('Eroare salvare în folder: ' + e.message, 'warning');
        }
      }
    }
  }
  // Fallback: standard download to Downloads folder
  XLSX.writeFile(wb, defaultName);
  return true;
}

// ── Export per client ─────────────────────────────────────────
function exportClientXLSX(client, interventions) {
  return loadXLSX().then(async () => {
    const wb = XLSX.utils.book_new();

    // Sort interventions descending by date
    const sorted = [...interventions].sort((a, b) => b.date.localeCompare(a.date));

    // --- Sheet 1: Intervenții (only show chemicals that were actually used) ---

    // Define all possible treatment columns with their data keys
    var treatCols = [
      { key: 'treat_cl_granule_gr',       label: 'Cl granule (gr)' },
      { key: 'treat_cl_tablete',          label: 'Cl tablete (buc)' },
      { key: 'treat_cl_tablete_export_gr',label: 'Cl tablete (gr)' },
      { key: 'treat_cl_lichid_bidoane',   label: 'Cl lichid (bid)' },
      { key: 'treat_ph_granule',          label: 'pH granule (kg)' },
      { key: 'treat_ph_lichid_bidoane',   label: 'pH lichid (bid)' },
      { key: 'treat_antialgic',           label: 'Antialgic (L)' },
      { key: 'treat_anticalcar',          label: 'Anticalcar (L)' },
      { key: 'treat_floculant',           label: 'Floculant (L)' },
      { key: 'treat_sare_saci',           label: 'Sare (saci)' },
      { key: 'treat_bicarbonat',          label: 'Bicarbonat (kg)' }
    ];

    // Also check dynamic stock products (treat_<product_id>)
    sorted.forEach(function(i) {
      Object.keys(i).forEach(function(k) {
        if (k.startsWith('treat_') && !treatCols.find(function(c) { return c.key === k; })) {
          var val = parseFloat(i[k]) || 0;
          if (val > 0) {
            treatCols.push({ key: k, label: k.replace('treat_', '').replace(/_/g, ' ') });
          }
        }
      });
    });

    // Filter: keep only treatment columns where at least one intervention has value > 0
    var usedTreatCols = treatCols.filter(function(col) {
      return sorted.some(function(i) { return (parseFloat(i[col.key]) || 0) > 0; });
    });

    // Build rows with only relevant columns
    var intRows = sorted.map(function(i) {
      var row = {
        'Data':              i.date,
        'Tehnician':         i.technician_name || '',
        'Clor măsurat':      i.measured_chlorine != null ? i.measured_chlorine : '',
        'pH măsurat':        i.measured_ph != null ? i.measured_ph : ''
      };
      // Add optional measurement columns only if any intervention has them
      if (sorted.some(function(x) { return x.measured_temp != null; }))
        row['Temperatură (°C)'] = i.measured_temp != null ? i.measured_temp : '';
      if (sorted.some(function(x) { return x.measured_hardness != null; }))
        row['Duritate'] = i.measured_hardness != null ? i.measured_hardness : '';
      if (sorted.some(function(x) { return x.measured_alkalinity != null; }))
        row['Alcalinitate'] = i.measured_alkalinity != null ? i.measured_alkalinity : '';
      if (sorted.some(function(x) { return x.measured_salinity != null; }))
        row['Salinitate'] = i.measured_salinity != null ? i.measured_salinity : '';

      // Add only used treatment columns
      usedTreatCols.forEach(function(col) {
        row[col.label] = i[col.key] || 0;
      });

      if (sorted.some(function(x) { return x.duration_minutes != null; }))
        row['Durată (min)'] = i.duration_minutes != null ? i.duration_minutes : '';
      row['Observații'] = i.observations || '';
      return row;
    });

    var ws1 = XLSX.utils.json_to_sheet(intRows);
    XLSX.utils.book_append_sheet(wb, ws1, 'Intervenții');

    // --- Sheet 2: Sumar (only used chemicals) ---
    var totals = calcTotals(sorted);
    var sumRows = [
      { 'Parametru': 'Total intervenții', 'Valoare': sorted.length, 'UM': 'buc' },
      { 'Parametru': 'Durată medie',      'Valoare': totals.avgDuration, 'UM': 'min' }
    ];
    var sumDefs = [
      { key: 'cl_granule_gr',       label: 'Cl granule total',  um: 'gr',      treatKey: 'treat_cl_granule_gr' },
      { key: 'cl_tablete',          label: 'Cl tablete total',  um: 'buc',     treatKey: 'treat_cl_tablete' },
      { key: 'cl_tablete_export_gr',label: 'Cl tablete total',  um: 'gr',      treatKey: 'treat_cl_tablete_export_gr' },
      { key: 'cl_lichid',           label: 'Cl lichid total',   um: 'bidoane', treatKey: 'treat_cl_lichid_bidoane' },
      { key: 'ph_granule',          label: 'pH granule total',  um: 'kg',      treatKey: 'treat_ph_granule' },
      { key: 'ph_lichid',           label: 'pH lichid total',   um: 'bidoane', treatKey: 'treat_ph_lichid_bidoane' },
      { key: 'antialgic',           label: 'Antialgic total',   um: 'L',       treatKey: 'treat_antialgic' },
      { key: 'anticalcar',          label: 'Anticalcar total',  um: 'L',       treatKey: 'treat_anticalcar' },
      { key: 'floculant',           label: 'Floculant total',   um: 'L',       treatKey: 'treat_floculant' },
      { key: 'sare',                label: 'Sare total',        um: 'saci',    treatKey: 'treat_sare_saci' },
      { key: 'bicarbonat',          label: 'Bicarbonat total',  um: 'kg',      treatKey: 'treat_bicarbonat' }
    ];
    sumDefs.forEach(function(sd) {
      if (sorted.some(function(i) { return (parseFloat(i[sd.treatKey]) || 0) > 0; })) {
        sumRows.push({ 'Parametru': sd.label, 'Valoare': totals[sd.key] || 0, 'UM': sd.um });
      }
    });
    var ws2 = XLSX.utils.json_to_sheet(sumRows);
    setColWidths(ws2, [28, 12, 12]);
    XLSX.utils.book_append_sheet(wb, ws2, 'Sumar');

    // --- Sheet 3: Măsurători (evoluție) ---
    const measRows = sorted.map(i => ({
      'Data':       i.date,
      'Clor':       i.measured_chlorine != null ? i.measured_chlorine : '',
      'pH':         i.measured_ph != null ? i.measured_ph : '',
      'Temp (°C)':  i.measured_temp != null ? i.measured_temp : '',
      'Duritate':   i.measured_hardness != null ? i.measured_hardness : '',
      'Alcalinitate': i.measured_alkalinity != null ? i.measured_alkalinity : ''
    }));
    const ws3 = XLSX.utils.json_to_sheet(measRows);
    setColWidths(ws3, [12,10,8,10,10,12]);
    XLSX.utils.book_append_sheet(wb, ws3, 'Măsurători');

    const filename = 'PoolMgr_' + sanitizeFilename(client.name) + '_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, filename, client.name);
    _uploadToDrive(wb, filename, null, client.name);
    return filename;
  });
}

// ── Export all clients ────────────────────────────────────────
function exportAllXLSX(clients, allInterventions) {
  return loadXLSX().then(async () => {
    const wb = XLSX.utils.book_new();

    // --- Sheet 1: Sumar General ---
    const generalRows = clients.map(c => {
      const ci = allInterventions.filter(i => i.client_id === c.client_id);
      const last = ci.length ? ci.sort((a, b) => b.date.localeCompare(a.date))[0] : null;
      const tot  = calcTotals(ci);
      return {
        'Client':               c.name,
        'Adresă':               c.address || '',
        'Volum (m³)':           c.pool_volume_mc || '',
        'Tip':                  c.pool_type || '',
        'Total intervenții':    ci.length,
        'Ultima intervenție':   last ? last.date : '',
        'Durată medie (min)':   tot.avgDuration,
        'Cl granule total (gr)': tot.cl_granule_gr,
        'Cl tablete total (buc)': tot.cl_tablete,
        'pH granule total (kg)': tot.ph_granule,
        'Antialgic total (L)':  tot.antialgic,
      };
    });
    const ws0 = XLSX.utils.json_to_sheet(generalRows);
    setColWidths(ws0, [22,28,12,12,16,16,16,18,18,16,16]);
    XLSX.utils.book_append_sheet(wb, ws0, 'Sumar General');

    // --- One sheet per client ---
    for (const c of clients) {
      const ci = allInterventions.filter(i => i.client_id === c.client_id);
      if (!ci.length) continue;
      const sorted = [...ci].sort((a, b) => b.date.localeCompare(a.date));
      const rows = sorted.map(i => ({
        'Data':             i.date,
        'Tehnician':        i.technician_name || '',
        'Clor':             i.measured_chlorine != null ? i.measured_chlorine : '',
        'pH':               i.measured_ph != null ? i.measured_ph : '',
        'Cl granule (gr)':  i.treat_cl_granule_gr || 0,
        'Cl tablete (buc)': i.treat_cl_tablete || 0,
        'pH granule (kg)':  i.treat_ph_granule || 0,
        'Antialgic (L)':    i.treat_antialgic || 0,
        'Durată (min)':     i.duration_minutes != null ? i.duration_minutes : '',
        'Observații':       i.observations || ''
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      setColWidths(ws, [12,18,8,8,14,14,14,12,12,30]);
      XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(c.name));
    }

    const filename = 'PoolMgr_Toate_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, filename, null);
    _uploadToDrive(wb, filename, null, null);
    return filename;
  });
}

// ── Export structured (one sheet per client, no summary) ──────
function exportStructuredXLSX(clients, allInterventions) {
  return loadXLSX().then(async () => {
    const wb = XLSX.utils.book_new();

    // Index sheet
    const indexRows = clients.map(c => ({
      'Client':             c.name,
      'Adresă':             c.address || '',
      'Volum (m³)':         c.pool_volume_mc || '',
      'Tip':                c.pool_type || '',
      'Total intervenții':  allInterventions.filter(i => i.client_id === c.client_id).length
    }));
    const wsIdx = XLSX.utils.json_to_sheet(indexRows);
    setColWidths(wsIdx, [24, 30, 12, 12, 18]);
    XLSX.utils.book_append_sheet(wb, wsIdx, 'Index');

    // One sheet per client (all intervention columns)
    for (const c of clients) {
      const ci = allInterventions.filter(i => i.client_id === c.client_id);
      if (!ci.length) continue;
      const sorted = [...ci].sort((a, b) => b.date.localeCompare(a.date));
      const rows = sorted.map(i => ({
        'Data':              i.date,
        'Tehnician':         i.technician_name || '',
        'Clor (FAC)':        i.measured_chlorine != null ? i.measured_chlorine : '',
        'Clor Total (TC)':   i.measured_tc != null ? i.measured_tc : '',
        'Clor Combinat (CC)':i.measured_tc != null && i.measured_chlorine != null
                               ? Math.round(Math.max(0, i.measured_tc - i.measured_chlorine) * 100) / 100 : '',
        'pH':                i.measured_ph != null ? i.measured_ph : '',
        'Temperatură (°C)':  i.measured_temp != null ? i.measured_temp : '',
        'Alcalinitate':      i.measured_alkalinity != null ? i.measured_alkalinity : '',
        'Duritate':          i.measured_hardness != null ? i.measured_hardness : '',
        'CYA':               i.measured_cya != null ? i.measured_cya : '',
        'Salinitate':        i.measured_salinity != null ? i.measured_salinity : '',
        'Cl granule (gr)':   i.treat_cl_granule_gr || 0,
        'Cl tablete (buc)':  i.treat_cl_tablete || 0,
        'Cl lichid (bid)':   i.treat_cl_lichid_bidoane || 0,
        'pH granule (kg)':   i.treat_ph_granule || 0,
        'pH lichid (bid)':   i.treat_ph_lichid_bidoane || 0,
        'Antialgic (L)':     i.treat_antialgic || 0,
        'Anticalcar (L)':    i.treat_anticalcar || 0,
        'Floculant (L)':     i.treat_floculant || 0,
        'Sare (saci)':       i.treat_sare_saci || 0,
        'Bicarbonat (kg)':   i.treat_bicarbonat || 0,
        'Durată (min)':      i.duration_minutes != null ? i.duration_minutes : '',
        'Observații':        i.observations || ''
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      setColWidths(ws, [12,18,10,12,14,8,14,12,10,8,10,14,14,14,14,14,12,12,12,12,14,12,30]);
      XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(c.name));
    }

    const filename = 'PoolMgr_Structurat_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, filename, null);
    _uploadToDrive(wb, filename, null, null);
    return filename;
  });
}

// ── Download import template ───────────────────────────────────
function downloadImportTemplate() {
  return loadXLSX().then(() => {
    const wb = XLSX.utils.book_new();
    const headers = [
      'client_name', 'data (YYYY-MM-DD)', 'clor_masurat', 'ph_masurat',
      'temperatura', 'alcalinitate', 'duritate', 'cya', 'clor_total',
      'cl_granule_gr', 'cl_tablete_buc', 'ph_granule_kg', 'antialgic_l',
      'observatii', 'durata_minute'
    ];
    const example1 = ['Andrei Ionescu', '2026-03-08', 1.2, 7.4, 28, 100, 250, '', '', 400, 2, 0.5, 0.75, 'Filtrare OK', 45];
    const example2 = ['Maria Popescu',  '2026-03-07', 0.8, 7.7, 26, 90,  220, 40, 1.2, 600, 3, 1.0, 0.5, '', 60];
    const ws = XLSX.utils.aoa_to_sheet([headers, example1, example2]);
    setColWidths(ws, [22,16,14,12,12,14,10,8,12,14,16,14,14,30,14]);
    XLSX.utils.book_append_sheet(wb, ws, 'Import');
    XLSX.writeFile(wb, 'PoolMgr_Template_Import.xlsx');
  });
}

// ── Import interventions from XLSX ────────────────────────────
async function importInterventionsXLSX(file) {
  if (!file) return;
  const xlsxInput = document.getElementById('import-xlsx-input');
  try {
    await loadXLSX();
    const data = await file.arrayBuffer();
    const wb   = XLSX.read(data, { type: 'array', cellDates: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!rows.length) { showToast('Fișierul este gol sau invalid.', 'error'); return; }

    let imported = 0, skipped = 0;
    const clientMap = {};
    (window.APP && APP.clients || []).forEach(c => {
      clientMap[c.name.toLowerCase().trim()] = c;
    });

    for (const row of rows) {
      const nameKey = String(row['client_name'] || '').toLowerCase().trim();
      const client  = clientMap[nameKey];
      if (!client) { skipped++; continue; }

      const dateRaw = row['data (YYYY-MM-DD)'];
      let dateStr = '';
      if (dateRaw instanceof Date) {
        dateStr = dateRaw.toISOString().split('T')[0];
      } else {
        dateStr = String(dateRaw || '').trim();
      }
      if (!dateStr) { skipped++; continue; }

      const num = v => { const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? null : n; };

      const intervention = {
        intervention_id:    'i_' + Date.now() + '_' + Math.random().toString(36).slice(2,8),
        client_id:          client.client_id,
        client_name:        client.name,
        technician_id:      (window.APP && APP.user) ? APP.user.technician_id : '',
        technician_name:    (window.APP && APP.user) ? APP.user.name : 'Import',
        date:               dateStr,
        created_at:         new Date().toISOString(),
        arrival_time:       null,
        departure_time:     null,
        measured_chlorine:  num(row['clor_masurat']),
        measured_ph:        num(row['ph_masurat']),
        measured_temp:      num(row['temperatura']),
        measured_alkalinity:num(row['alcalinitate']),
        measured_hardness:  num(row['duritate']),
        measured_cya:       num(row['cya']),
        measured_tc:        num(row['clor_total']),
        measured_salinity:  null,
        rec_cl_gr:          null, rec_cl_tab: null, rec_ph_kg: null, rec_anti_l: null,
        treat_cl_granule_gr:         num(row['cl_granule_gr'])  || 0,
        treat_cl_tablete:            num(row['cl_tablete_buc']) || 0,
        treat_ph_granule:            num(row['ph_granule_kg'])  || 0,
        treat_antialgic:             num(row['antialgic_l'])    || 0,
        treat_cl_lichid_bidoane:     0,
        treat_ph_lichid_bidoane:     0,
        treat_anticalcar:            0,
        treat_floculant:             0,
        treat_sare_saci:             0,
        treat_bicarbonat:            0,
        observations:       String(row['observatii'] || '').trim(),
        photos:             [],
        duration_minutes:   num(row['durata_minute']),
        geo_lat: null, geo_lng: null, geo_accuracy: null,
        synced: false
      };

      await saveIntervention(intervention);
      if (window.APP) APP.interventions.push(intervention);
      imported++;
    }

    if (xlsxInput) xlsxInput.value = '';
    if (imported > 0 && window.APP) { await loadData(); renderDashboard(); }
    showToast(`Import complet: ${imported} intervenții importate${skipped ? ', ' + skipped + ' rânduri ignorate (client negăsit)' : ''}.`, imported > 0 ? 'success' : 'error');
  } catch(e) {
    if (xlsxInput) xlsxInput.value = '';
    showToast('Eroare import: ' + e.message, 'error');
  }
}


// ── Import clients from XLSX ─────────────────────────────────
async function importClientsXLSX(file) {
  if (!file) return;
  const inp = document.getElementById('import-clients-input');
  try {
    await loadXLSX();
    const data = await file.arrayBuffer();
    const wb   = XLSX.read(data, { type: 'array', cellDates: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!rows.length) { showToast('Fișierul este gol sau invalid.', 'error'); return; }

    // Helper: strip non-breaking spaces (\u00a0) and trim
    const clean = v => String(v || '').replace(/\u00a0/g, ' ').trim();
    const fixPhone = v => { var s = clean(v); if (/^\d{9}$/.test(s) && s[0] === '7') s = '0' + s; return s; };

    let imported = 0, skipped = 0, idCounter = 0;
    const baseTime = Date.now();
    const existingMap = {};
    (window.APP && APP.clients || []).forEach(c => {
      existingMap[c.name.toLowerCase().replace(/\u00a0/g, ' ').trim()] = c;
    });

    const newClients = [];

    for (const row of rows) {
      const name = clean(row['nume'] || row['name'] || row['client_name'] || row['NUME']);
      if (!name) { skipped++; continue; }

      const nameKey = name.toLowerCase();
      if (existingMap[nameKey]) {
        const existing = existingMap[nameKey];
        const phone = fixPhone(row['telefon'] || row['phone'] || row['TELEFON']);
        const addr  = clean(row['adresa']  || row['address'] || row['ADRESA']);
        const vol   = parseFloat(clean(row['volum_mc'] || row['pool_volume_mc'] || row['VOLUM'])) || 0;
        const type  = clean(row['tip_piscina'] || row['pool_type'] || row['TIP']).toLowerCase();
        const notes = clean(row['observatii'] || row['notes'] || row['OBS']);
        if (phone) existing.phone = phone;
        if (addr)  existing.address = addr;
        if (vol)   existing.pool_volume_mc = vol;
        if (type === 'interior' || type === 'exterior') existing.pool_type = type;
        if (notes) existing.notes = notes;
        existing.updated_at = new Date().toISOString();
        await put('clients', existing);
        newClients.push(existing);
        imported++;
        continue;
      }

      idCounter++;
      const now = new Date().toISOString();
      const client = {
        client_id:      'c_' + (baseTime + idCounter) + '_' + Math.random().toString(36).slice(2, 8),
        name,
        phone:          fixPhone(row['telefon'] || row['phone'] || row['TELEFON']),
        address:        clean(row['adresa']  || row['address'] || row['ADRESA']),
        pool_volume_mc: parseFloat(clean(row['volum_mc'] || row['pool_volume_mc'] || row['VOLUM'])) || 0,
        pool_type:      (clean(row['tip_piscina'] || row['pool_type'] || row['TIP']).toLowerCase() === 'interior') ? 'interior' : 'exterior',
        notes:          clean(row['observatii'] || row['notes'] || row['OBS']),
        visit_frequency_days: parseInt(row['frecventa_zile'] || row['visit_frequency_days'] || 7) || 7,
        active:         true,
        created_at:     now,
        updated_at:     now,
        latitude:       null,
        longitude:      null,
        location_set:   false
      };

      await put('clients', client);
      existingMap[nameKey] = client;
      newClients.push(client);
      imported++;
    }

    // Batch push all to GAS (one request instead of N)
    if (newClients.length && typeof isSyncConfigured === 'function' && isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'clients', data: newClients })
      }).catch(err => console.warn('[SYNC] Client batch push failed:', err.message));
    }

    if (inp) inp.value = '';
    if (imported > 0 && window.APP) { await loadData(); renderDashboard(); }
    showToast('Import complet: ' + imported + ' clienți importați' + (skipped ? ', ' + skipped + ' rânduri ignorate' : '') + '.', imported > 0 ? 'success' : 'error');
  } catch(e) {
    if (inp) inp.value = '';
    showToast('Eroare import: ' + e.message, 'error');
  }
}

// ── Download client import template ──────────────────────────
async function downloadClientTemplate() {
  try { await loadXLSX(); } catch (e) {
    showToast('SheetJS nu este disponibil. Reconectați-vă la internet.', 'warning');
    return;
  }
  const wb = XLSX.utils.book_new();
  const headers = ['nume', 'telefon', 'adresa', 'volum_mc', 'tip_piscina', 'observatii'];
  const ex1 = ['Popescu Ion', '0712345678', 'Str. Exemplu 1, București', 50, 'exterior', 'Piscina 10x5m'];
  const ex2 = ['Ionescu Maria', '0723456789', 'Str. Exemplu 2, Cluj', 30, 'interior', ''];
  const ws = XLSX.utils.aoa_to_sheet([headers, ex1, ex2]);
  ws['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 36 }, { wch: 10 }, { wch: 14 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Clienti');
  XLSX.writeFile(wb, 'template-clienti.xlsx');
}

// ── Helpers ───────────────────────────────────────────────────
function calcTotals(interventions) {
  const durations = interventions.filter(i => i.duration_minutes != null).map(i => i.duration_minutes);
  return {
    cl_granule_gr:        sum(interventions, 'treat_cl_granule_gr'),
    cl_tablete:           sum(interventions, 'treat_cl_tablete'),
    cl_tablete_export_gr: sum(interventions, 'treat_cl_tablete_export_gr'),
    cl_lichid:            sum(interventions, 'treat_cl_lichid_bidoane'),
    ph_granule:           round2(sum(interventions, 'treat_ph_granule')),
    ph_lichid:            sum(interventions, 'treat_ph_lichid_bidoane'),
    antialgic:            round2(sum(interventions, 'treat_antialgic')),
    anticalcar:           round2(sum(interventions, 'treat_anticalcar')),
    floculant:            round2(sum(interventions, 'treat_floculant')),
    sare:                 sum(interventions, 'treat_sare_saci'),
    bicarbonat:           round2(sum(interventions, 'treat_bicarbonat')),
    avgDuration:          durations.length ? Math.round(durations.reduce((a, b) => a + b, 0) / durations.length) : 0
  };
}

function sum(arr, field) {
  return arr.reduce((acc, item) => acc + (parseFloat(item[field]) || 0), 0);
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function fmtDateExport(date) {
  return date.toISOString().split('T')[0].replace(/-/g, '');
}

function sanitizeFilename(name) {
  return (name || 'client').replace(/[^a-zA-Z0-9_\-\.]/g, '_').substring(0, 40);
}

function sanitizeSheetName(name) {
  return (name || 'Client').replace(/[\[\]\*\/\\\?:]/g, '_').substring(0, 31);
}


// ── Upload to Google Drive via GAS ────────────────────────────

// == Helper: format date as DD.MM.YYYY ==
function fmtDateDMY(dateStr) {
  if (!dateStr) return '';
  var s = String(dateStr).trim();
  // Already in DD.MM.YYYY format
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) return s;
  // YYYY-MM-DD format
  var parts = s.split('-');
  if (parts.length === 3 && parts[0].length === 4) return parts[2] + '.' + parts[1] + '.' + parts[0];
  // Try parsing as Date object string (e.g. "Mon Mar 09 2026 00:00:00 GMT+0200...")
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var dd = ('0' + d.getDate()).slice(-2);
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    return dd + '.' + mm + '.' + d.getFullYear();
  }
  return s;
}


// ══════════════════════════════════════════════════════════════════════
// ██ TEMPLATE-BASED EXPORT (ExcelJS) — V1 Chimicale + V2 Servicii ██
// ══════════════════════════════════════════════════════════════════════

var EXCELJS_CDN = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
var _exceljsLoaded = false;

/** Lazy-load ExcelJS from CDN */
function loadExcelJS() {
  if (_exceljsLoaded && typeof ExcelJS !== 'undefined') return Promise.resolve();
  return new Promise(function(resolve, reject) {
    var script  = document.createElement('script');
    script.src  = EXCELJS_CDN;
    script.onload = function() { _exceljsLoaded = true; resolve(); };
    script.onerror = function() { reject(new Error('Nu s-a putut incarca ExcelJS. Verificati conexiunea.')); };
    document.head.appendChild(script);
  });
}

/** Convert base64 string to ArrayBuffer */
function _b64toBuffer(b64) {
  var bin = atob(b64);
  var len = bin.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

/** Parse operations from intervention (handles array, JSON string, or empty) */
function _parseOps(operations) {
  if (Array.isArray(operations)) return operations;
  if (typeof operations === 'string' && operations.length > 0) {
    try { return JSON.parse(operations); } catch (e) { return []; }
  }
  return [];
}

/** Write ExcelJS workbook to export folder or download */
async function _writeExcelJSFile(wb, defaultName, clientName) {
  var buf = await wb.xlsx.writeBuffer();

  // Try stored directory handle first (File System Access API)
  if (typeof window.showDirectoryPicker === 'function') {
    var dirHandle = await _getExportDirHandle();
    if (!dirHandle) dirHandle = await pickExportFolder();

    if (dirHandle) {
      var hasPermission = await _verifyDirPermission(dirHandle);
      if (!hasPermission) {
        dirHandle = await pickExportFolder();
        hasPermission = dirHandle ? await _verifyDirPermission(dirHandle) : false;
      }

      if (dirHandle && hasPermission) {
        try {
          var targetDir = dirHandle;
          if (clientName) {
            var folderName = sanitizeFilename(clientName);
            targetDir = await dirHandle.getDirectoryHandle(folderName, { create: true });
          }
          var fileHandle = await targetDir.getFileHandle(defaultName, { create: true });
          var writable = await fileHandle.createWritable();
          await writable.write(new Uint8Array(buf));
          await writable.close();
          return true;
        } catch (e) {
          console.warn('[EXPORT] Directory write failed, falling back:', e.message);
          showToast('Eroare salvare in folder: ' + e.message, 'warning');
        }
      }
    }
  }

  // Fallback: standard download
  var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = defaultName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
  return true;
}

/** Upload ExcelJS workbook buffer to Google Drive via GAS */
function _uploadExcelJSToDrive(buf, fileName, clientName) {
  if (typeof isSyncConfigured !== 'function' || !isSyncConfigured()) return;
  try {
    var safeName = clientName ? clientName.trim().replace(/\s+/g, ' ') : '';
    // Convert buffer to base64
    var bytes = new Uint8Array(buf);
    var binary = '';
    for (var i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    var b64 = btoa(binary);

    fetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      redirect: 'follow',
      body: JSON.stringify({
        action: 'saveExportToDrive',
        fileName: fileName,
        data: b64,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        clientName: safeName
      })
    }).then(function(r) { return r.json(); })
      .then(function(res) {
        if (res.success) {
          showToast('Salvat in Drive: Export Interventii/' + (clientName ? clientName + '/' : '') + fileName, 'success', 4000);
        } else {
          console.warn('[DRIVE] Save failed:', res.error);
        }
      }).catch(function(e) {
        console.warn('[DRIVE] Upload failed:', e.message);
      });
  } catch (e) {
    console.warn('[DRIVE] Upload error:', e.message);
  }
}


// ── V2 Template: Fill Servicii Abonament sheet ─────────────────────
// Template structure: R1-R5 header, R6 labels, R7 values, R8 sep,
// R9-R10 table header, R11-R28 data (18 slots), R29 total, R30 sep,
// R31 total plata, R32 sep, R33 footer
//
// IMPORTANT: ExcelJS spliceRows does NOT update merge references!
// Instead we: save footer styles/content → clear everything → write data →
// write footer at new position → manually fix all merges.
//
// Operations map to columns B(2) through I(9) in the template.
// Extra operations (not in template) are added as new columns J, K, ...

/** Remove diacritics and normalize string for comparison */
function _normOp(str) {
  if (!str) return '';
  return String(str)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // strip diacritics
    .replace(/[\r\n]+/g, ' ')                           // newlines → space
    .replace(/\s+/g, ' ')                               // collapse whitespace
    .trim()
    .toLowerCase();
}

/** Read template R10 sub-headers (columns B onwards) and return normalized map */
function _readTemplateOpsHeaders(ws, headerRow, startCol, endCol) {
  var headers = [];
  for (var c = startCol; c <= endCol; c++) {
    var cell = ws.getRow(headerRow).getCell(c);
    var raw = cell.value ? String(cell.value) : '';
    headers.push({ col: c, raw: raw, norm: _normOp(raw) });
  }
  return headers;
}

/** Match an app operation name against template headers (diacritics-insensitive) */
function _findOpColumn(opName, templateHeaders) {
  var norm = _normOp(opName);
  for (var h = 0; h < templateHeaders.length; h++) {
    if (templateHeaders[h].norm === norm) return templateHeaders[h].col;
    // Partial match: one contains the other
    if (norm.length > 3 && templateHeaders[h].norm.length > 3) {
      if (templateHeaders[h].norm.indexOf(norm) >= 0 || norm.indexOf(templateHeaders[h].norm) >= 0) {
        return templateHeaders[h].col;
      }
    }
  }
  return -1; // no match
}

/** Helper: Parse cell reference like "A29" → { col: 1, row: 29 } */
function _parseCellRef(ref) {
  var match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  var letters = match[1];
  var row = parseInt(match[2]);
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col: col, row: row };
}

/** Build cell reference like { col: 1, row: 29 } → "A29" */
function _cellRef(col, row) {
  var letters = '';
  var c = col;
  while (c > 0) {
    var rem = (c - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor((c - 1) / 26);
  }
  return letters + row;
}

// ══════════════════════════════════════════════════════════════════
// BUILD FROM SCRATCH — V2 (Servicii Abonament)
// 27 rows, 9 cols (A-I), 12 data slots (R11-R22)
// ══════════════════════════════════════════════════════════════════
async function _buildV2(wb, client, sorted, prices) {

  // ── Colors ──
  var BLUE = 'FF3B6FA0', BDRK = 'FF1A3A5C', LTXT = 'FFD0DEF0', LBG = 'FFE8F1F8';
  var DTXT = 'FF1A1A2E', MTXT = 'FF4A4A5A', GBDR = 'FFD9D9D9', CREAM = 'FFFAF8F3';
  var GREEN = 'FF1A6B2A', GOLD_BDR = 'FF9A7020', GOLD_FILL = 'FFF5E6C8', WHITE = 'FFFFFFFF';

  // ── Constants ──
  var FIRST_DATA_ROW = 11, DATA_SLOTS = 12, ORIG_LAST_COL = 9;
  var FIRST_OP_COL = 2, LAST_OP_COL_DEFAULT = 9, HEADER_SUB_ROW = 10;
  var NR = Math.min(sorted.length, DATA_SLOTS);

  // ── Default V2 operation headers (B-I in R10) ──
  var DEFAULT_OPS = [
    'Verificare\nparametri\napa',
    'Curatare\npereti si\nvana',
    'Aspirare\nfund\npiscina',
    'Curatare\nlinie apa\nsi skimmer',
    'Tratament\nchimic\napa',
    'Spalare\nfiltru si\npompa',
    'Verificare\nechipamente\ntehnice',
    'Curatare\nzona\npiscina'
  ];

  // ── Build templateHeaders from DEFAULT_OPS ──
  var templateHeaders = [];
  for (var th = 0; th < DEFAULT_OPS.length; th++) {
    templateHeaders.push({ col: FIRST_OP_COL + th, raw: DEFAULT_OPS[th], norm: _normOp(DEFAULT_OPS[th]) });
  }

  // ── Date helpers ──
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = NR ? fmtDateDMY(sorted[0].date) : '';
  var lastDate  = NR ? fmtDateDMY(sorted[NR - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr  = 'AQS - ' + todayYMD;
  var pretIntv = parseFloat(client.pret_interventie) || prices.pret_interventie || 250;

  // ── 1. Collect all ops & map to columns ──
  var allOpsSet = {}, allOpsOrder = [];
  var anyHasOps = false;
  sorted.forEach(function(intv) {
    var ops = _parseOps(intv.operations);
    if (ops.length > 0) anyHasOps = true;
    ops.forEach(function(op) {
      if (op && !allOpsSet[op]) { allOpsSet[op] = true; allOpsOrder.push(op); }
    });
  });

  // FALLBACK: If NO intervention has operations data, default to all ops checked
  if (!anyHasOps && sorted.length > 0) {
    var defaultOpsFlat = DEFAULT_OPS.map(function(h) { return h.replace(/\n/g, ' ').trim(); });
    sorted.forEach(function(intv) {
      if (!intv.operations || (Array.isArray(intv.operations) && intv.operations.length === 0)) {
        intv._defaultOps = defaultOpsFlat;
      }
    });
  }

  var opToCol = {}, usedCols = {}, extraOps = [];
  allOpsOrder.forEach(function(op) {
    var col = _findOpColumn(op, templateHeaders);
    if (col >= 0 && !usedCols[col]) { opToCol[op] = col; usedCols[col] = true; }
    else { extraOps.push(op); }
  });
  var nextExtra = LAST_OP_COL_DEFAULT + 1;
  extraOps.forEach(function(op) { opToCol[op] = nextExtra++; });
  var LAST_COL = Math.max(ORIG_LAST_COL, nextExtra - 1);

  // ── 2. Create worksheet ──
  var ws = wb.addWorksheet('Servicii Abonament', {
    views: [{ state: 'frozen', ySplit: 10, zoomScale: 100 }],
    pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });

  // Column widths
  var colWidths = [{ width: 14 }];
  for (var cw = 2; cw <= LAST_COL; cw++) colWidths.push({ width: 12 });
  ws.columns = colWidths;

  // ── Reusable styles ──
  var thinBdr = { style: 'thin', color: { argb: GBDR } };
  var medBdr  = { style: 'medium', color: { argb: BLUE } };
  var goldThinBdr = { style: 'thin', color: { argb: GOLD_BDR } };
  var allThinBorders = { top: thinBdr, left: thinBdr, bottom: thinBdr, right: thinBdr };
  var fillLBG  = { type: 'pattern', pattern: 'solid', fgColor: { argb: LBG } };
  var fillWhite = { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } };
  var fillGold = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD_FILL } };
  var fillBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
  var fillCream = { type: 'pattern', pattern: 'solid', fgColor: { argb: CREAM } };
  var centerMiddle = { horizontal: 'center', vertical: 'middle' };
  var centerMiddleWrap = { horizontal: 'center', vertical: 'middle', wrapText: true };

  // Helper: apply thin borders + outer frame border to all cells in a row
  function _applyRowBorders(rowNum, firstCol, lastCol) {
    var row = ws.getRow(rowNum);
    for (var c = firstCol; c <= lastCol; c++) {
      var cell = row.getCell(c);
      var b = JSON.parse(JSON.stringify(allThinBorders));
      if (rowNum === 1) b.top = medBdr;
      if (rowNum === 27) b.bottom = medBdr;
      if (c === 1) b.left = medBdr;
      if (c === lastCol) b.right = medBdr;
      cell.border = b;
    }
  }

  // ── R1 (h=27.75): Company name ──
  ws.mergeCells(1, 1, 1, LAST_COL);
  var r1 = ws.getRow(1); r1.height = 27.75;
  var c1 = r1.getCell(1);
  c1.value = 'AQUATIS ENGINEERING';
  c1.font = { name: 'Arial', size: 16, bold: true, color: { argb: BDRK } };
  c1.fill = fillLBG;
  c1.alignment = centerMiddle;
  _applyRowBorders(1, 1, LAST_COL);
  r1.commit();

  // ── R2 (h=18): CUI ──
  ws.mergeCells(2, 1, 2, LAST_COL);
  var r2 = ws.getRow(2); r2.height = 18;
  var c2 = r2.getCell(1);
  c2.value = 'CUI: RO12345678 | J40/1234/2020';
  c2.font = { name: 'Arial', size: 9, color: { argb: MTXT } };
  c2.fill = fillLBG;
  c2.alignment = centerMiddle;
  _applyRowBorders(2, 1, LAST_COL);
  r2.commit();

  // ── R3 (h=15.75): Contact ──
  ws.mergeCells(3, 1, 3, LAST_COL);
  var r3 = ws.getRow(3); r3.height = 15.75;
  var c3 = r3.getCell(1);
  c3.value = 'Tel: 0700 000 000 | Email: office@aquatis.ro';
  c3.font = { name: 'Arial', size: 9, color: { argb: MTXT } };
  c3.fill = fillLBG;
  c3.alignment = centerMiddle;
  _applyRowBorders(3, 1, LAST_COL);
  r3.commit();

  // ── R4 (h=3.75): Gold separator ──
  ws.mergeCells(4, 1, 4, LAST_COL);
  var r4 = ws.getRow(4); r4.height = 3.75;
  var c4 = r4.getCell(1);
  c4.value = '';
  c4.fill = fillGold;
  for (var c4c = 1; c4c <= LAST_COL; c4c++) {
    var c4cell = ws.getRow(4).getCell(c4c);
    c4cell.border = { top: goldThinBdr, bottom: goldThinBdr, left: c4c === 1 ? medBdr : thinBdr, right: c4c === LAST_COL ? medBdr : thinBdr };
  }
  r4.commit();

  // ── R5 (h=21.75): Title ──
  ws.mergeCells(5, 1, 5, LAST_COL);
  var r5 = ws.getRow(5); r5.height = 21.75;
  var c5 = r5.getCell(1);
  c5.value = 'FISA SERVICII ABONAMENT INTRETINERE PISCINA';
  c5.font = { name: 'Arial', size: 11, bold: true, color: { argb: BDRK } };
  c5.fill = fillWhite;
  c5.alignment = centerMiddle;
  _applyRowBorders(5, 1, LAST_COL);
  r5.commit();

  // ── R6 (h=15.75): Labels row ──
  ws.mergeCells(6, 1, 6, 2); // A6:B6
  ws.mergeCells(6, 3, 6, 5); // C6:E6
  ws.mergeCells(6, 6, 6, 7); // F6:G6
  ws.mergeCells(6, 8, 6, Math.min(9, LAST_COL)); // H6:I6
  var r6 = ws.getRow(6); r6.height = 15.75;
  var labelFont = { name: 'Arial', size: 9, bold: true, color: { argb: BLUE } };
  var labels6 = [
    { col: 1, text: 'Client:' },
    { col: 3, text: 'Perioada:' },
    { col: 6, text: 'Nr. Document:' },
    { col: 8, text: 'Data:' }
  ];
  labels6.forEach(function(l) {
    var cell = r6.getCell(l.col);
    cell.value = l.text;
    cell.font = JSON.parse(JSON.stringify(labelFont));
    cell.fill = fillLBG;
    cell.alignment = centerMiddle;
  });
  // Fill remaining cells in R6 with style
  for (var c6 = 1; c6 <= LAST_COL; c6++) {
    var cell6 = r6.getCell(c6);
    if (!cell6.fill || !cell6.fill.fgColor) cell6.fill = fillLBG;
  }
  _applyRowBorders(6, 1, LAST_COL);
  // Extra cols beyond I for labels row
  if (LAST_COL > 9) {
    for (var ec6 = 10; ec6 <= LAST_COL; ec6++) {
      var eCell6 = r6.getCell(ec6);
      eCell6.fill = fillLBG;
    }
  }
  r6.commit();

  // ── R7 (h=16.5): Values row ──
  ws.mergeCells(7, 1, 7, 2); // A7:B7
  ws.mergeCells(7, 3, 7, 5); // C7:E7
  ws.mergeCells(7, 6, 7, 7); // F7:G7
  ws.mergeCells(7, 8, 7, Math.min(9, LAST_COL)); // H7:I7
  var r7 = ws.getRow(7); r7.height = 16.5;
  var valFont = { name: 'Arial', size: 9, color: { argb: DTXT } };
  var vals7 = [
    { col: 1, text: client.name || '' },
    { col: 3, text: period },
    { col: 6, text: docNr },
    { col: 8, text: todayStr }
  ];
  vals7.forEach(function(v) {
    var cell = r7.getCell(v.col);
    cell.value = v.text;
    cell.font = JSON.parse(JSON.stringify(valFont));
    cell.fill = fillWhite;
    cell.alignment = centerMiddle;
  });
  for (var c7 = 1; c7 <= LAST_COL; c7++) {
    var cell7 = r7.getCell(c7);
    if (!cell7.fill || !cell7.fill.fgColor) cell7.fill = fillWhite;
  }
  _applyRowBorders(7, 1, LAST_COL);
  r7.commit();

  // ── R8 (h=3.75): Gold separator ──
  ws.mergeCells(8, 1, 8, LAST_COL);
  var r8 = ws.getRow(8); r8.height = 3.75;
  var c8 = r8.getCell(1);
  c8.value = '';
  c8.fill = fillGold;
  for (var c8c = 1; c8c <= LAST_COL; c8c++) {
    ws.getRow(8).getCell(c8c).border = { top: goldThinBdr, bottom: goldThinBdr, left: c8c === 1 ? medBdr : thinBdr, right: c8c === LAST_COL ? medBdr : thinBdr };
  }
  r8.commit();

  // ── R9 (h=21.75): Header top row ──
  ws.mergeCells(9, 1, 10, 1); // A9:A10
  ws.mergeCells(9, 2, 9, LAST_COL); // B9:I9 (or extended)
  var r9 = ws.getRow(9); r9.height = 21.75;
  var hdrWhiteFont = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
  var cA9 = r9.getCell(1);
  cA9.value = 'Data\nInterventie';
  cA9.font = JSON.parse(JSON.stringify(hdrWhiteFont));
  cA9.fill = fillBlue;
  cA9.alignment = centerMiddleWrap;
  var cB9 = r9.getCell(2);
  cB9.value = 'Operatiuni efectuate';
  cB9.font = JSON.parse(JSON.stringify(hdrWhiteFont));
  cB9.fill = fillBlue;
  cB9.alignment = centerMiddleWrap;
  _applyRowBorders(9, 1, LAST_COL);
  // Override fill for all cells in R9
  for (var c9 = 1; c9 <= LAST_COL; c9++) {
    r9.getCell(c9).fill = fillBlue;
  }
  r9.commit();

  // ── R10 (h=45.75): Sub-headers ──
  var r10 = ws.getRow(10); r10.height = 45.75;
  var subFont = { name: 'Arial', size: 7, bold: true, color: { argb: WHITE } };
  // A10 is merged with A9
  var cA10 = r10.getCell(1);
  cA10.fill = fillBlue;
  cA10.alignment = centerMiddleWrap;
  // Write default ops in B10-I10
  for (var oi2 = 0; oi2 < DEFAULT_OPS.length; oi2++) {
    var opCol = FIRST_OP_COL + oi2;
    var opCell = r10.getCell(opCol);
    opCell.value = DEFAULT_OPS[oi2];
    opCell.font = JSON.parse(JSON.stringify(subFont));
    opCell.fill = fillGold;
    opCell.alignment = centerMiddleWrap;
    opCell.border = { top: goldThinBdr, bottom: goldThinBdr, left: goldThinBdr, right: goldThinBdr };
  }
  // Write extra ops headers beyond I
  for (var exi = 0; exi < extraOps.length; exi++) {
    var exCol = LAST_OP_COL_DEFAULT + 1 + exi;
    var exCell = r10.getCell(exCol);
    exCell.value = extraOps[exi].replace(/ /g, '\n');
    exCell.font = JSON.parse(JSON.stringify(subFont));
    exCell.fill = fillGold;
    exCell.alignment = centerMiddleWrap;
    exCell.border = { top: goldThinBdr, bottom: goldThinBdr, left: goldThinBdr, right: goldThinBdr };
  }
  _applyRowBorders(10, 1, LAST_COL);
  // Ensure A10 has blue fill (part of A9:A10 merge)
  cA10.fill = fillBlue;
  r10.commit();

  // ── R11-R22: Data rows (12 slots) ──
  var checkFont = { name: 'Arial', size: 11, bold: true, color: { argb: GREEN } };
  var dataFont  = { name: 'Arial', size: 8, color: { argb: MTXT } };
  var lastDataRow = NR > 0 ? (FIRST_DATA_ROW + NR - 1) : FIRST_DATA_ROW;

  for (var dr = 0; dr < DATA_SLOTS; dr++) {
    var rowNum = FIRST_DATA_ROW + dr;
    var row = ws.getRow(rowNum);
    row.height = 19.5;
    // Alternating fill: odd (11,13,15...)=CREAM, even (12,14,16...)=WHITE
    var rowFill = (rowNum % 2 === 1) ? fillCream : fillWhite;

    for (var dc = 1; dc <= LAST_COL; dc++) {
      var dCell = row.getCell(dc);
      dCell.value = '';
      dCell.fill = rowFill;
      dCell.font = JSON.parse(JSON.stringify(dataFont));
      dCell.alignment = centerMiddle;
    }

    if (dr < NR) {
      var entry = sorted[dr];
      // A: date
      row.getCell(1).value = fmtDateDMY(entry.date);
      row.getCell(1).font = JSON.parse(JSON.stringify(dataFont));
      row.getCell(1).alignment = centerMiddle;

      // Fill checkmarks
      var ops = _parseOps(entry.operations);
      if (ops.length === 0 && entry._defaultOps) ops = entry._defaultOps;
      for (var oi3 = 0; oi3 < ops.length; oi3++) {
        var col = opToCol[ops[oi3]];
        if (!col || col < FIRST_OP_COL) col = _findOpColumn(ops[oi3], templateHeaders);
        if (col && col >= FIRST_OP_COL) {
          var chkCell = row.getCell(col);
          chkCell.value = '\u2713';
          chkCell.font = JSON.parse(JSON.stringify(checkFont));
        }
      }
    }

    _applyRowBorders(rowNum, 1, LAST_COL);
    row.commit();
  }

  // ── R23 (h=24): Total interventii ──
  var ROW_TOTAL = 23;
  ws.mergeCells(ROW_TOTAL, 1, ROW_TOTAL, 6); // A23:F23
  ws.mergeCells(ROW_TOTAL, 7, ROW_TOTAL, LAST_COL); // G23:I23 (or extended)
  var r23 = ws.getRow(ROW_TOTAL); r23.height = 24;
  var totalFont = { name: 'Arial', size: 10, bold: true, color: { argb: BDRK } };
  var cA23 = r23.getCell(1);
  cA23.value = 'Total interventii efectuate';
  cA23.font = JSON.parse(JSON.stringify(totalFont));
  cA23.fill = fillLBG;
  cA23.alignment = centerMiddle;
  var cG23 = r23.getCell(7);
  cG23.font = JSON.parse(JSON.stringify(totalFont));
  cG23.fill = fillLBG;
  cG23.alignment = centerMiddle;
  cG23.numFmt = '0';
  if (NR > 0) {
    cG23.value = { formula: 'COUNTA(A' + FIRST_DATA_ROW + ':A' + lastDataRow + ')' };
  } else {
    cG23.value = 0;
  }
  _applyRowBorders(ROW_TOTAL, 1, LAST_COL);
  for (var c23 = 1; c23 <= LAST_COL; c23++) {
    var cell23 = r23.getCell(c23);
    if (!cell23.fill || !cell23.fill.fgColor) cell23.fill = fillLBG;
  }
  r23.commit();

  // ── R24 (h=3.75): Gold separator ──
  var ROW_SEP2 = 24;
  ws.mergeCells(ROW_SEP2, 1, ROW_SEP2, LAST_COL);
  var r24 = ws.getRow(ROW_SEP2); r24.height = 3.75;
  r24.getCell(1).value = '';
  r24.getCell(1).fill = fillGold;
  for (var c24 = 1; c24 <= LAST_COL; c24++) {
    ws.getRow(ROW_SEP2).getCell(c24).border = { top: goldThinBdr, bottom: goldThinBdr, left: c24 === 1 ? medBdr : thinBdr, right: c24 === LAST_COL ? medBdr : thinBdr };
  }
  r24.commit();

  // ── R25 (h=25.5): Total de plata ──
  var ROW_PAY = 25;
  ws.mergeCells(ROW_PAY, 1, ROW_PAY, 5); // A25:E25
  ws.mergeCells(ROW_PAY, 6, ROW_PAY, LAST_COL); // F25:I25 (or extended)
  var r25 = ws.getRow(ROW_PAY); r25.height = 25.5;
  var payFont = { name: 'Arial', size: 12, bold: true, color: { argb: BDRK } };
  var cA25 = r25.getCell(1);
  cA25.value = 'TOTAL DE PLATA';
  cA25.font = JSON.parse(JSON.stringify(payFont));
  cA25.fill = fillLBG;
  cA25.alignment = centerMiddle;
  var cF25 = r25.getCell(6);
  cF25.font = JSON.parse(JSON.stringify(payFont));
  cF25.fill = fillLBG;
  cF25.alignment = centerMiddle;
  cF25.numFmt = '#,##0 "Lei"';
  if (NR > 0) {
    cF25.value = { formula: 'IFERROR(COUNTA(A' + FIRST_DATA_ROW + ':A' + lastDataRow + ')*' + pretIntv + ',0)' };
  } else {
    cF25.value = 0;
  }
  _applyRowBorders(ROW_PAY, 1, LAST_COL);
  for (var c25 = 1; c25 <= LAST_COL; c25++) {
    var cell25 = r25.getCell(c25);
    if (!cell25.fill || !cell25.fill.fgColor) cell25.fill = fillLBG;
  }
  r25.commit();

  // ── R26 (h=4.5): Spacer ──
  var ROW_SPACER = 26;
  ws.mergeCells(ROW_SPACER, 1, ROW_SPACER, LAST_COL);
  var r26 = ws.getRow(ROW_SPACER); r26.height = 4.5;
  r26.getCell(1).value = '';
  r26.getCell(1).fill = fillWhite;
  _applyRowBorders(ROW_SPACER, 1, LAST_COL);
  r26.commit();

  // ── R27 (h=15.75): Footer ──
  var ROW_FOOTER = 27;
  ws.mergeCells(ROW_FOOTER, 1, ROW_FOOTER, 5); // A27:E27
  ws.mergeCells(ROW_FOOTER, 6, ROW_FOOTER, LAST_COL); // F27:I27 (or extended)
  var r27 = ws.getRow(ROW_FOOTER); r27.height = 15.75;
  var cA27 = r27.getCell(1);
  cA27.value = 'Toate preturile sunt exprimate in RON, inclusiv TVA.';
  cA27.font = { name: 'Arial', size: 8, italic: true, color: { argb: MTXT } };
  cA27.fill = fillWhite;
  cA27.alignment = { horizontal: 'left', vertical: 'middle' };
  var cF27 = r27.getCell(6);
  cF27.value = 'S.C. Aquatis Engineering S.R.L.';
  cF27.font = { name: 'Arial', size: 8, bold: true, color: { argb: BDRK } };
  cF27.fill = fillWhite;
  cF27.alignment = { horizontal: 'right', vertical: 'middle' };
  _applyRowBorders(ROW_FOOTER, 1, LAST_COL);
  r27.commit();

  // ── Strip diacritics on data rows only ──
  if (NR > 0) _stripAllDiacritics(ws, FIRST_DATA_ROW, lastDataRow, LAST_COL);

  return wb;
}


// ── V1 Template: Fill Raport Interventii (Chimicale) sheet ─────────
// New template (openpyxl/Python): R1 navy bar, R2 company info, R3 cyan bar,
// R4 title, R5 labels, R6 values (J5/K5 + J6/K6 NOT merged), R7 sep,
// R8-R9 table header, R10-R19 data (10 slots), R20 cantitate totala,
// R21 pret unitar, R22 total general, R23 footer
//
// Chemical columns C-J (3-10): Clor Rapid, Clor Lent, pH-, Antialgic,
// Floculant, Dedurizant, pH Lichid, Cl Lichid

var V1_CHEM_COLUMNS = [
  { col: 3,  label: 'Clor Rapid',  keys: ['treat_cl_granule_gr', 'treat_cl_granule'] },
  { col: 4,  label: 'Clor Lent',   keys: ['treat_cl_tablete', 'treat_cl_tablete_export_gr'] },
  { col: 5,  label: 'pH\u2212',    keys: ['treat_ph_granule', 'treat_ph_minus_gr'] },
  { col: 6,  label: 'Antialgic',   keys: ['treat_antialgic'] },
  { col: 7,  label: 'Floculant',   keys: ['treat_floculant'] },
  { col: 8,  label: 'Dedurizant',  keys: ['treat_bicarbonat'] },
  { col: 9,  label: 'pH Lichid',   keys: ['treat_ph_lichid_bidoane', 'treat_ph_minus_l'] },
  { col: 10, label: 'Cl Lichid',   keys: ['treat_cl_lichid_bidoane', 'treat_cl_lichid'] }
];

// Default prices for V1 template R21 (fallback values matching template)
var V1_DEFAULT_PRICES = {
  3: 57,     // Clor Rapid
  4: 56.4,   // Clor Lent
  5: 13,     // pH−
  6: 29,     // Antialgic
  7: 25,     // Floculant
  8: 32,     // Dedurizant
  9: 184,    // pH Lichid
  10: 180    // Cl Lichid
};

// Map from V1 column to product_id for price lookup
var V1_COL_PRICE_KEYS = {
  3: ['cl_granule'],
  4: ['cl_tablete'],
  5: ['ph_minus_gr'],
  6: ['antialgic'],
  7: ['floculant'],
  8: ['bicarbonat'],
  9: ['ph_minus_l'],
  10: ['cl_lichid']
};

async function _buildV1(wb, client, sorted, prices) {

  // ── Colors ──
  var NAVY = 'FF0D2D5A', DKBLUE = 'FF1D507F', LTBLUE = 'FFE8F3FB', EBLUE = 'FFE0EEF8';
  var VLTBLUE = 'FFF0F6FB', PALEBLUE = 'FFEDF4FB', WHITE = 'FFFFFFFF';

  // ── Constants ──
  var FIRST_DATA_ROW = 10, TEMPLATE_SLOTS = 10, LAST_COL = 11;
  var NR = Math.min(sorted.length, TEMPLATE_SLOTS);

  // ── Date helpers ──
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = NR ? fmtDateDMY(sorted[0].date) : '';
  var lastDate  = NR ? fmtDateDMY(sorted[NR - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr  = 'AQS - ' + todayYMD;

  // ── Create worksheet ──
  var ws = wb.addWorksheet('Raport Interventii', {
    views: [{ state: 'frozen', ySplit: 9, zoomScale: 115 }],
    pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });

  // Column widths: A=12, B=8, C-J=11, K=14
  ws.columns = [
    { width: 12 }, { width: 8 },
    { width: 11 }, { width: 11 }, { width: 11 }, { width: 11 },
    { width: 11 }, { width: 11 }, { width: 11 }, { width: 11 },
    { width: 14 }
  ];

  // ── Reusable styles ──
  var thinBdr = { style: 'thin' };
  var medBdr  = { style: 'medium' };
  var allThinBorders = { top: thinBdr, left: thinBdr, bottom: thinBdr, right: thinBdr };
  var fillNavy   = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
  var fillWhite  = { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } };
  var fillLtblue = { type: 'pattern', pattern: 'solid', fgColor: { argb: LTBLUE } };
  var fillEblue  = { type: 'pattern', pattern: 'solid', fgColor: { argb: EBLUE } };
  var fillVltblue = { type: 'pattern', pattern: 'solid', fgColor: { argb: VLTBLUE } };
  var fillPaleblue = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALEBLUE } };
  var fillDkblue = { type: 'pattern', pattern: 'solid', fgColor: { argb: DKBLUE } };
  var centerMiddle = { horizontal: 'center', vertical: 'middle' };
  var centerMiddleWrap = { horizontal: 'center', vertical: 'middle', wrapText: true };
  var fontDkblue9 = { name: 'Arial', size: 9, color: { argb: DKBLUE } };
  var fontDkblue9b = { name: 'Arial', size: 9, bold: true, color: { argb: DKBLUE } };

  // Helper: apply thin borders + outer frame (medium) for a row
  function _applyV1Borders(rowNum, fc, lc, extraBottom) {
    var row = ws.getRow(rowNum);
    for (var c = fc; c <= lc; c++) {
      var cell = row.getCell(c);
      var b = JSON.parse(JSON.stringify(allThinBorders));
      if (rowNum === 1) b.top = medBdr;
      if (rowNum === 23) b.bottom = medBdr;
      if (extraBottom) b.bottom = medBdr;
      if (c === 1) b.left = medBdr;
      if (c === lc) b.right = medBdr;
      cell.border = b;
    }
  }

  // ── R1 (h=6): Navy bar ──
  ws.mergeCells('A1:K1');
  var r1 = ws.getRow(1); r1.height = 6;
  r1.getCell(1).fill = fillNavy;
  _applyV1Borders(1, 1, LAST_COL);
  r1.commit();

  // ── R2 (h=42): Company info (richText) ──
  ws.mergeCells('A2:K2');
  var r2 = ws.getRow(2); r2.height = 42;
  var c2 = r2.getCell(1);
  c2.value = { richText: [
    { font: { name: 'Arial', size: 12, bold: true, color: { argb: DKBLUE } }, text: 'AQUATIS ENGINEERING S.R.L.' },
    { text: '\n' },
    { font: { name: 'Arial', size: 10, bold: true, color: { argb: DKBLUE } }, text: 'Str. Exemplu Nr. 10, Bucuresti | CUI: RO12345678' }
  ]};
  c2.fill = fillWhite;
  c2.alignment = centerMiddleWrap;
  _applyV1Borders(2, 1, LAST_COL);
  r2.commit();

  // ── R3 (h=4): Cyan separator ──
  ws.mergeCells('A3:K3');
  var r3 = ws.getRow(3); r3.height = 4;
  r3.getCell(1).fill = fillLtblue;
  _applyV1Borders(3, 1, LAST_COL);
  r3.commit();

  // ── R4 (h=24): Title ──
  ws.mergeCells('A4:K4');
  var r4 = ws.getRow(4); r4.height = 24;
  var c4 = r4.getCell(1);
  c4.value = 'RAPORT INTERVENTII PISCINA';
  c4.font = { name: 'Arial', size: 13, bold: true, color: { argb: DKBLUE } };
  c4.fill = fillWhite;
  c4.alignment = centerMiddle;
  _applyV1Borders(4, 1, LAST_COL);
  r4.commit();

  // ── R5 (h=15): Labels ──
  ws.mergeCells('A5:C5');
  ws.mergeCells('D5:F5');
  ws.mergeCells('G5:I5');
  ws.mergeCells('J5:K5');
  var r5 = ws.getRow(5); r5.height = 15;
  var labels5 = [
    { col: 1, text: 'Client' },
    { col: 4, text: 'Perioada' },
    { col: 7, text: 'Nr. Document' },
    { col: 10, text: 'Data' }
  ];
  labels5.forEach(function(l) {
    var cell = r5.getCell(l.col);
    cell.value = l.text;
    cell.font = JSON.parse(JSON.stringify(fontDkblue9b));
    cell.fill = fillEblue;
    cell.alignment = centerMiddle;
  });
  // Fill remaining cells with EBLUE
  for (var c5i = 1; c5i <= LAST_COL; c5i++) {
    var c5c = r5.getCell(c5i);
    if (!c5c.fill || !c5c.fill.fgColor) c5c.fill = fillEblue;
  }
  _applyV1Borders(5, 1, LAST_COL);
  r5.commit();

  // ── R6 (h=18): Values ──
  ws.mergeCells('A6:C6');
  ws.mergeCells('D6:F6');
  ws.mergeCells('G6:I6');
  ws.mergeCells('J6:K6');
  var r6 = ws.getRow(6); r6.height = 18;
  var vals6 = [
    { col: 1, text: client.name || '' },
    { col: 4, text: period },
    { col: 7, text: docNr },
    { col: 10, text: todayStr }
  ];
  vals6.forEach(function(v) {
    var cell = r6.getCell(v.col);
    cell.value = v.text;
    cell.font = JSON.parse(JSON.stringify(fontDkblue9));
    cell.fill = fillVltblue;
    cell.alignment = centerMiddle;
  });
  for (var c6i = 1; c6i <= LAST_COL; c6i++) {
    var c6c = r6.getCell(c6i);
    if (!c6c.fill || !c6c.fill.fgColor) c6c.fill = fillVltblue;
  }
  _applyV1Borders(6, 1, LAST_COL);
  r6.commit();

  // ── R7 (h=4): Separator ──
  ws.mergeCells('A7:K7');
  var r7 = ws.getRow(7); r7.height = 4;
  r7.getCell(1).fill = fillLtblue;
  _applyV1Borders(7, 1, LAST_COL, true); // bottom=medium
  r7.commit();

  // ── R8 (h=18): Header top ──
  ws.mergeCells('A8:A9');
  ws.mergeCells('B8:B9');
  ws.mergeCells('C8:J8');
  ws.mergeCells('K8:K9');
  var r8 = ws.getRow(8); r8.height = 18;
  var hdrWhiteFont = { name: 'Arial', size: 9, bold: true, color: { argb: WHITE } };
  var hdrCells8 = [
    { col: 1, text: 'Data' },
    { col: 2, text: 'Nr.\nIntv.' },
    { col: 3, text: 'Produse chimice utilizate' },
    { col: 11, text: 'Total\nPlata\n(RON)' }
  ];
  hdrCells8.forEach(function(h) {
    var cell = r8.getCell(h.col);
    cell.value = h.text;
    cell.font = JSON.parse(JSON.stringify(hdrWhiteFont));
    cell.fill = fillDkblue;
    cell.alignment = centerMiddleWrap;
  });
  for (var c8i = 1; c8i <= LAST_COL; c8i++) {
    var c8c = r8.getCell(c8i);
    if (!c8c.fill || !c8c.fill.fgColor) c8c.fill = fillDkblue;
  }
  _applyV1Borders(8, 1, LAST_COL);
  r8.commit();

  // ── R9 (h=30): Header sub (chemical column names) ──
  var r9 = ws.getRow(9); r9.height = 30;
  var subFont = { name: 'Arial', size: 8, bold: true, color: { argb: WHITE } };
  var chemHeaders = ['Clor\nRapid', 'Clor\nLent', 'pH\u2212', 'Antialgic', 'Floculant', 'Dedurizant', 'pH\nLichid', 'Cl\nLichid'];
  for (var ch = 0; ch < chemHeaders.length; ch++) {
    var chCell = r9.getCell(3 + ch);
    chCell.value = chemHeaders[ch];
    chCell.font = JSON.parse(JSON.stringify(subFont));
    chCell.fill = fillDkblue;
    chCell.alignment = centerMiddleWrap;
  }
  // A9, B9 are merged with R8; fill them with dkblue
  r9.getCell(1).fill = fillDkblue;
  r9.getCell(2).fill = fillDkblue;
  r9.getCell(11).fill = fillDkblue;
  _applyV1Borders(9, 1, LAST_COL, true); // bottom=medium
  r9.commit();

  // ── R10-R19: Data rows (10 slots) ──
  var dataFont = { name: 'Arial', size: 9, color: { argb: DKBLUE } };
  var lastDataRow = NR > 0 ? (FIRST_DATA_ROW + NR - 1) : FIRST_DATA_ROW;

  for (var dr = 0; dr < TEMPLATE_SLOTS; dr++) {
    var rowNum = FIRST_DATA_ROW + dr;
    var row = ws.getRow(rowNum);
    row.height = 18;
    // Alternating: even rows (10,12,14,16,18)=PALEBLUE, odd (11,13,15,17,19)=WHITE
    var rowFill = (rowNum % 2 === 0) ? fillPaleblue : fillWhite;

    for (var dc = 1; dc <= LAST_COL; dc++) {
      var dCell = row.getCell(dc);
      dCell.value = '';
      dCell.fill = rowFill;
      dCell.font = JSON.parse(JSON.stringify(dataFont));
      dCell.alignment = (dc === 1) ? { horizontal: 'left', vertical: 'middle' } : centerMiddle;
      if (dc >= 2 && dc <= LAST_COL) dCell.numFmt = '#,##0.00';
    }

    if (dr < NR) {
      var entry = sorted[dr];
      // A: date
      row.getCell(1).value = fmtDateDMY(entry.date);
      // B: count = 1
      row.getCell(2).value = 1;
      // C-J: chemical values
      V1_CHEM_COLUMNS.forEach(function(cc) {
        var val = _getChemValue(entry, cc.keys);
        var cell = row.getCell(cc.col);
        if (val > 0) {
          cell.value = (cc.col <= 5) ? Math.round(val) : (val % 1 === 0 ? val : Math.round(val * 100) / 100);
        } else {
          cell.value = '';
        }
      });
      // K: empty
    }

    _applyV1Borders(rowNum, 1, LAST_COL);
    row.commit();
  }

  // Fix last data row bottom border = medium (if data exists)
  if (NR > 0) {
    var ldr = ws.getRow(lastDataRow);
    for (var bc = 1; bc <= LAST_COL; bc++) {
      var bCell = ldr.getCell(bc);
      var bb = bCell.border ? JSON.parse(JSON.stringify(bCell.border)) : {};
      bb.bottom = medBdr;
      bCell.border = bb;
    }
    ldr.commit();
  }

  // ── R20 (h=18): Cantitate totala ──
  var totalsRow = 20;
  ws.mergeCells('A20:B20');
  var r20 = ws.getRow(totalsRow); r20.height = 18;
  var cA20 = r20.getCell(1);
  cA20.value = 'Cantitate totala';
  cA20.font = JSON.parse(JSON.stringify(fontDkblue9b));
  cA20.fill = fillEblue;
  cA20.alignment = centerMiddle;
  V1_CHEM_COLUMNS.forEach(function(cc) {
    var cell = r20.getCell(cc.col);
    cell.value = { formula: 'SUM(' + _excelCol(cc.col) + FIRST_DATA_ROW + ':' + _excelCol(cc.col) + lastDataRow + ')' };
    cell.font = JSON.parse(JSON.stringify(fontDkblue9b));
    cell.fill = fillEblue;
    cell.alignment = centerMiddle;
    cell.numFmt = '#,##0.00';
  });
  r20.getCell(11).fill = fillEblue;
  r20.getCell(11).font = JSON.parse(JSON.stringify(fontDkblue9b));
  _applyV1Borders(totalsRow, 1, LAST_COL);
  r20.commit();

  // ── R21 (h=18): Pret unitar ──
  var pretRow = 21;
  ws.mergeCells('A21:B21');
  var r21 = ws.getRow(pretRow); r21.height = 18;
  var cA21 = r21.getCell(1);
  cA21.value = 'Pret unitar (RON)';
  cA21.font = JSON.parse(JSON.stringify(fontDkblue9));
  cA21.fill = fillVltblue;
  cA21.alignment = centerMiddle;
  V1_CHEM_COLUMNS.forEach(function(cc) {
    var price = 0;
    var priceKeys = V1_COL_PRICE_KEYS[cc.col] || [];
    for (var pk = 0; pk < priceKeys.length; pk++) {
      if (prices[priceKeys[pk]] && prices[priceKeys[pk]] > 0) { price = prices[priceKeys[pk]]; break; }
    }
    if (!price) price = V1_DEFAULT_PRICES[cc.col] || 0;
    var cell = r21.getCell(cc.col);
    cell.value = price > 0 ? price : '';
    cell.font = JSON.parse(JSON.stringify(fontDkblue9));
    cell.fill = fillVltblue;
    cell.alignment = centerMiddle;
    cell.numFmt = '#,##0.00';
  });
  r21.getCell(11).fill = fillVltblue;
  r21.getCell(11).font = JSON.parse(JSON.stringify(fontDkblue9));
  _applyV1Borders(pretRow, 1, LAST_COL);
  r21.commit();

  // ── R22 (h=20): Total general ──
  var genRow = 22;
  ws.mergeCells('A22:B22');
  var r22 = ws.getRow(genRow); r22.height = 20;
  var fontGen = { name: 'Arial', size: 10, bold: true, color: { argb: DKBLUE } };
  var cA22 = r22.getCell(1);
  cA22.value = 'TOTAL GENERAL (RON)';
  cA22.font = JSON.parse(JSON.stringify(fontGen));
  cA22.fill = fillEblue;
  cA22.alignment = centerMiddle;
  V1_CHEM_COLUMNS.forEach(function(cc) {
    var cl = _excelCol(cc.col);
    var cell = r22.getCell(cc.col);
    cell.value = { formula: cl + totalsRow + '*' + cl + pretRow };
    cell.font = JSON.parse(JSON.stringify(fontGen));
    cell.fill = fillEblue;
    cell.alignment = centerMiddle;
    cell.numFmt = '#,##0.00';
  });
  // K22: SUM(C22:J22)
  var cK22 = r22.getCell(11);
  cK22.value = { formula: 'SUM(C' + genRow + ':J' + genRow + ')' };
  cK22.font = JSON.parse(JSON.stringify(fontGen));
  cK22.fill = fillEblue;
  cK22.alignment = centerMiddle;
  cK22.numFmt = '#,##0.00';
  _applyV1Borders(genRow, 1, LAST_COL, true); // bottom=medium
  r22.commit();

  // ── R23 (h=15): Footer ──
  ws.mergeCells('A23:G23');
  ws.mergeCells('H23:K23');
  var r23 = ws.getRow(23); r23.height = 15;
  var cA23 = r23.getCell(1);
  cA23.value = 'Toate preturile sunt exprimate in RON';
  cA23.font = { name: 'Arial', size: 8, italic: true, color: { argb: DKBLUE } };
  cA23.fill = fillWhite;
  cA23.alignment = { horizontal: 'left', vertical: 'middle' };
  var cH23 = r23.getCell(8);
  cH23.value = 'S.C. Aquatis Engineering S.R.L.';
  cH23.font = { name: 'Arial', size: 8, bold: true, color: { argb: DKBLUE } };
  cH23.fill = fillWhite;
  cH23.alignment = { horizontal: 'right', vertical: 'middle' };
  _applyV1Borders(23, 1, LAST_COL);
  r23.commit();

  // ── Strip diacritics on data rows only ──
  if (NR > 0) _stripAllDiacritics(ws, FIRST_DATA_ROW, lastDataRow, LAST_COL);

  return wb;
}

/** Get chemical value from intervention, checking multiple possible keys */
function _getChemValue(entry, keys) {
  for (var i = 0; i < keys.length; i++) {
    var val = parseFloat(entry[keys[i]]);
    if (val > 0) return val;
  }
  return 0;
}

/** Strip Romanian diacritics from a string */
function _stripDiacritics(s) {
  if (!s || typeof s !== 'string') return s;
  return s
    .replace(/[ăâ]/g, 'a').replace(/[ĂÂ]/g, 'A')
    .replace(/[îì]/g, 'i').replace(/[ÎÌ]/g, 'I')
    .replace(/[șş]/g, 's').replace(/[ȘŞ]/g, 'S')
    .replace(/[țţ]/g, 't').replace(/[ȚŢ]/g, 'T');
}

/** Strip diacritics from all text cells in a worksheet */
function _stripAllDiacritics(ws, startRow, lastRow, lastCol) {
  // Build set of merge slave cells (skip these to avoid duplicating text)
  var slaveCells = {};
  if (ws._merges) {
    Object.keys(ws._merges).forEach(function(key) {
      var m = ws._merges[key].model;
      for (var mr = m.top; mr <= m.bottom; mr++) {
        for (var mc = m.left; mc <= m.right; mc++) {
          if (mr !== m.top || mc !== m.left) slaveCells[mr + '_' + mc] = true;
        }
      }
    });
  }
  for (var r = startRow; r <= lastRow; r++) {
    var row = ws.getRow(r);
    var changed = false;
    for (var c = 1; c <= lastCol; c++) {
      if (slaveCells[r + '_' + c]) continue; // skip slave cells in merges
      var cell = row.getCell(c);
      if (cell.value && typeof cell.value === 'string') {
        var stripped = _stripDiacritics(cell.value);
        if (stripped !== cell.value) { cell.value = stripped; changed = true; }
      } else if (cell.value && cell.value.richText) {
        cell.value.richText.forEach(function(part) {
          if (part.text) part.text = _stripDiacritics(part.text);
        });
        changed = true;
      }
    }
    if (changed) row.commit();
  }
}

/** Convert 1-based column number to Excel letter (1=A, 2=B, ..., 26=Z, 27=AA) */
function _excelCol(num) {
  var s = '';
  while (num > 0) {
    var mod = (num - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

/** Set cell value preserving existing style */
function _setCellValue(ws, rowNum, colNum, value) {
  var row = ws.getRow(rowNum);
  var cell = row.getCell(colNum);
  cell.value = value;
  row.commit();
}

/** Set cell formula preserving existing style */
function _setCellFormula(ws, rowNum, colNum, formula) {
  var row = ws.getRow(rowNum);
  var cell = row.getCell(colNum);
  cell.value = { formula: formula };
  row.commit();
}

/** Capture styles from a template row (returns object keyed by 1-based col) */
function _captureRowStyles(ws, rowNum, numCols) {
  var styles = {};
  var row = ws.getRow(rowNum);
  for (var c = 1; c <= numCols; c++) {
    var cell = row.getCell(c);
    styles[c] = {
      font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : undefined,
      fill: cell.fill ? JSON.parse(JSON.stringify(cell.fill)) : undefined,
      border: cell.border ? JSON.parse(JSON.stringify(cell.border)) : undefined,
      alignment: cell.alignment ? JSON.parse(JSON.stringify(cell.alignment)) : undefined,
      numFmt: cell.numFmt || undefined
    };
  }
  return styles;
}

/** Set cell value and apply style */
function _setCellValueWithStyle(row, colNum, value, style) {
  var cell = row.getCell(colNum);
  cell.value = value;
  if (style) {
    if (style.font) cell.font = style.font;
    if (style.fill) cell.fill = style.fill;
    if (style.border) cell.border = style.border;
    if (style.alignment) cell.alignment = style.alignment;
    if (style.numFmt) cell.numFmt = style.numFmt;
  }
}

/** Copy worksheet from source to target workbook (row-by-row with styles, merges, dimensions) */
function _copyWorksheet(sourceWs, targetWb, sheetName) {
  var targetWs = targetWb.addWorksheet(sheetName);

  // 1. Copy column widths
  for (var ci = 1; ci <= sourceWs.columnCount; ci++) {
    var srcCol = sourceWs.getColumn(ci);
    targetWs.getColumn(ci).width = srcCol.width;
    if (srcCol.hidden) targetWs.getColumn(ci).hidden = true;
  }

  // 2. Build slave cells set from merges
  var slaves = {};
  if (sourceWs._merges) {
    Object.keys(sourceWs._merges).forEach(function(key) {
      var m = sourceWs._merges[key].model;
      for (var r = m.top; r <= m.bottom; r++) {
        for (var c = m.left; c <= m.right; c++) {
          if (r !== m.top || c !== m.left) slaves[r + ',' + c] = true;
        }
      }
    });
  }

  // 3. Copy all rows (values + styles)
  sourceWs.eachRow({ includeEmpty: true }, function(row, rn) {
    var tRow = targetWs.getRow(rn);
    tRow.height = row.height;
    if (row.hidden) tRow.hidden = true;
    row.eachCell({ includeEmpty: true }, function(cell, cn) {
      var tCell = tRow.getCell(cn);
      // Value: skip slaves to prevent duplicated text
      if (!slaves[rn + ',' + cn]) {
        if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
          tCell.value = { formula: cell.value.formula };
        } else {
          tCell.value = cell.value;
        }
      }
      // Styles: always copy
      if (cell.font) tCell.font = JSON.parse(JSON.stringify(cell.font));
      if (cell.fill) tCell.fill = JSON.parse(JSON.stringify(cell.fill));
      if (cell.border) tCell.border = JSON.parse(JSON.stringify(cell.border));
      if (cell.alignment) tCell.alignment = JSON.parse(JSON.stringify(cell.alignment));
      if (cell.numFmt) tCell.numFmt = cell.numFmt;
    });
    tRow.commit();
  });

  // 4. Copy merges using INTEGER COORDINATES (not string keys!)
  if (sourceWs._merges) {
    Object.keys(sourceWs._merges).forEach(function(key) {
      var m = sourceWs._merges[key].model;
      try {
        targetWs.mergeCells(m.top, m.left, m.bottom, m.right);
      } catch (e) {
        console.warn('[COPY] Merge failed:', key, e.message);
      }
    });
  }

  // 5. Re-apply ALL borders+fill+font via full style copy (ExcelJS clears
  //    slave cell styles during mergeCells). We re-copy ALL styles, not just borders.
  sourceWs.eachRow({ includeEmpty: true }, function(row, rn) {
    var tRow = targetWs.getRow(rn);
    row.eachCell({ includeEmpty: true }, function(cell, cn) {
      var tCell = tRow.getCell(cn);
      if (cell.border && Object.keys(cell.border).length > 0) {
        tCell.border = JSON.parse(JSON.stringify(cell.border));
      }
      if (cell.fill && cell.fill.type) {
        tCell.fill = JSON.parse(JSON.stringify(cell.fill));
      }
      if (cell.font) {
        tCell.font = JSON.parse(JSON.stringify(cell.font));
      }
      if (cell.alignment) {
        tCell.alignment = JSON.parse(JSON.stringify(cell.alignment));
      }
    });
    tRow.commit();
  });

  // 6. Outer frame fix: ExcelJS mergeCells() can clear master cell borders.
  //    Re-apply left:medium on col 1 and right:medium on last col by reading from source.
  var srcLastCol = sourceWs.columnCount || 11;
  sourceWs.eachRow({ includeEmpty: true }, function(row, rn) {
    var tRow = targetWs.getRow(rn);
    // Left border from source col 1
    var srcA = row.getCell(1);
    if (srcA.border && srcA.border.left && srcA.border.left.style === 'medium') {
      var tA = tRow.getCell(1);
      var bA = tA.border ? JSON.parse(JSON.stringify(tA.border)) : {};
      bA.left = { style: 'medium' };
      tA.border = bA;
    }
    // Right border from source last col
    var srcL = row.getCell(srcLastCol);
    if (srcL.border && srcL.border.right && srcL.border.right.style === 'medium') {
      var tL = tRow.getCell(srcLastCol);
      var bL = tL.border ? JSON.parse(JSON.stringify(tL.border)) : {};
      bL.right = { style: 'medium' };
      tL.border = bL;
    }
    tRow.commit();
  });

  // 7. Page setup
  if (sourceWs.pageSetup) {
    try { targetWs.pageSetup = JSON.parse(JSON.stringify(sourceWs.pageSetup)); } catch(e) {}
  }

  return targetWs;
}


// ── NEW: Export Deviz Chimicale (V1 template-based) ────────────────
function exportDevizChimicale(client, interventions) {
  return loadExcelJS().then(async function() {
    try {
      var sorted = interventions.slice().sort(function(a, b) { return String(a.date).localeCompare(String(b.date)); });
      var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};

      var wb = new ExcelJS.Workbook();
      await _buildV1(wb, client, sorted, prices);

      var ws = wb.getWorksheet(1);
      if (ws) ws.name = sanitizeSheetName(client.name || 'Chimicale');

      var fname = sanitizeFilename(client.name) + '_Chimicale_' + fmtDateExport(new Date()) + '.xlsx';
      await _writeExcelJSFile(wb, fname, client.name);
      return fname;
    } catch (e) {
      console.error('[EXPORT] exportDevizChimicale failed:', e.message, e.stack);
      if (typeof showToast === 'function') showToast('Eroare export chimicale: ' + e.message, 'error');
      throw e;
    }
  });
}

// ── NEW: Export Deviz Complet (V1 + V2 template-based) ─────────────
function exportDevizComplet(client, interventions) {
  return loadExcelJS().then(async function() {
    try {
      var sorted = interventions.slice().sort(function(a, b) { return String(a.date).localeCompare(String(b.date)); });
      var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};

      var wbV1 = new ExcelJS.Workbook();
      await _buildV1(wbV1, client, sorted, prices);

      var wbV2 = new ExcelJS.Workbook();
      await _buildV2(wbV2, client, sorted, prices);

      var wbFinal = new ExcelJS.Workbook();
      var nameChim = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Chim');
      var nameServ = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Serv');

      _copyWorksheet(wbV1.getWorksheet(1), wbFinal, nameChim);
      _copyWorksheet(wbV2.getWorksheet(1), wbFinal, nameServ);

      var fname = sanitizeFilename(client.name) + '_Deviz_' + fmtDateExport(new Date()) + '.xlsx';
      await _writeExcelJSFile(wbFinal, fname, client.name);
      return fname;
    } catch (e) {
      console.error('[EXPORT] exportDevizComplet failed:', e.message, e.stack);
      if (typeof showToast === 'function') showToast('Eroare export complet: ' + e.message, 'error');
      throw e;
    }
  });
}

// ── NEW: Export All Deviz Mixed (all clients, template-based) ──────
function exportAllDevizMixed(clients, allInterventions, filter) {
  return loadExcelJS().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};

    var wbFinal = new ExcelJS.Workbook();
    var sheetCount = 0;

    // Normalize interventions to object keyed by client_id
    var intvByClient = {};
    if (Array.isArray(allInterventions)) {
      allInterventions.forEach(function(i) {
        var cid = i.client_id;
        if (!intvByClient[cid]) intvByClient[cid] = [];
        intvByClient[cid].push(i);
      });
    } else {
      intvByClient = allInterventions || {};
    }

    for (var ci = 0; ci < clients.length; ci++) {
      var client = clients[ci];
      var cid = client.client_id;
      var clientIntv = (intvByClient[cid] || []).slice().sort(function(a, b) {
        return String(a.date).localeCompare(String(b.date));
      });

      // Apply filter
      if (filter) {
        if (filter.mode === 'date' && filter.fromDate) {
          clientIntv = clientIntv.filter(function(i) { return i.date >= filter.fromDate; });
        } else if (filter.mode === 'last' && filter.lastN) {
          clientIntv = clientIntv.slice(-filter.lastN);
        }
      }

      if (clientIntv.length === 0) continue;

      var baseName = sanitizeSheetName(client.name || 'Client');
      var devizType = parseInt(client.deviz_type) || 2;

      if (devizType === 2) {
        // V2 = complet (both sheets)
        // V1 (Chimicale)
        var wbV1 = new ExcelJS.Workbook();
        await _buildV1(wbV1, client, clientIntv, prices);
        var chemName = baseName.substring(0, 28) + '_Ch';
        if (wbFinal.worksheets.some(function(s) { return s.name === chemName; })) {
          chemName = chemName.substring(0, 24) + '_' + (sheetCount + 1);
        }
        _copyWorksheet(wbV1.getWorksheet(1), wbFinal, chemName);
        sheetCount++;

        // V2 (Servicii)
        var wbV2 = new ExcelJS.Workbook();
        await _buildV2(wbV2, client, clientIntv, prices);
        var opsName = baseName.substring(0, 28) + '_Sv';
        if (wbFinal.worksheets.some(function(s) { return s.name === opsName; })) {
          opsName = opsName.substring(0, 24) + '_' + (sheetCount + 1);
        }
        _copyWorksheet(wbV2.getWorksheet(1), wbFinal, opsName);
        sheetCount++;
      } else {
        // V1 = chimicale only
        var wbV1only = new ExcelJS.Workbook();
        await _buildV1(wbV1only, client, clientIntv, prices);
        var chemNameV = baseName.substring(0, 28) + '_Ch';
        if (wbFinal.worksheets.some(function(s) { return s.name === chemNameV; })) {
          chemNameV = chemNameV.substring(0, 24) + '_' + (sheetCount + 1);
        }
        _copyWorksheet(wbV1only.getWorksheet(1), wbFinal, chemNameV);
        sheetCount++;
      }
    }

    if (sheetCount === 0) {
      showToast('Nicio interventie de exportat.', 'warning');
      return;
    }

    var fname = 'DevizToti_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeExcelJSFile(wbFinal, fname);
    return fname;
  }).catch(function(e) {
    console.error('[EXPORT] exportAllDevizMixed failed:', e.message, e.stack);
    if (typeof showToast === 'function') showToast('Eroare export: ' + e.message, 'error');
    throw e;
  });
}

function _uploadToDrive(wb, fileName, mimeType, clientName) {
  if (typeof isSyncConfigured !== 'function' || !isSyncConfigured()) return;
  try {
    // Sanitize clientName for consistent Drive folder naming (avoid duplicate folders)
    var safeName = clientName ? clientName.trim().replace(/\s+/g, ' ') : '';
    var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
    fetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      redirect: 'follow',
      body: JSON.stringify({
        action: 'saveExportToDrive',
        fileName: fileName,
        data: wbout,
        mimeType: mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        clientName: safeName
      })
    }).then(function(r) { return r.json(); })
      .then(function(res) {
        if (res.success) {
          showToast('Salvat in Drive: Export Interventii/' + (clientName ? clientName + '/' : '') + fileName, 'success', 4000);
        } else {
          console.warn('[DRIVE] Save failed:', res.error);
        }
      }).catch(function(e) {
        console.warn('[DRIVE] Upload failed:', e.message);
      });
  } catch (e) {
    console.warn('[DRIVE] Upload error:', e.message);
  }
}

function setColWidths(ws, widths) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}

// ── Export Billing Deviz ──────────────────────────────────────
function exportBillingXLSX(client, interventions) {
  return loadXLSX().then(async function() {
    var wb = XLSX.utils.book_new();
    var sorted = interventions.slice().sort(function(a, b) { return a.date.localeCompare(b.date); });
    var since = client.last_billing_date || '';
    var today = new Date().toISOString().split('T')[0];
    var devizNr = 'D-' + today.replace(/-/g, '') + '-' + (client.client_id || '').slice(-4);

    // Sheet 1: Deviz
    var headerRows = [
      { 'A': 'DEVIZ SERVICII PISCINA' },
      { 'A': 'Client:', 'B': client.name, 'D': 'Nr. deviz:', 'E': devizNr },
      { 'A': 'Adresa:', 'B': client.address || '-', 'D': 'Data:', 'E': today },
      { 'A': 'Telefon:', 'B': client.phone || '-', 'D': 'Perioada:', 'E': (since || '-') + ' - ' + today },
      { 'A': '' }
    ];

    var dataRows = sorted.map(function(inv, idx) {
      return {
        'Nr.': idx + 1,
        'Data': inv.date,
        'Tehnician': inv.technician_name || '',
        'Cl granule (gr)': inv.treat_cl_granule_gr || 0,
        'Cl tablete (buc)': inv.treat_cl_tablete || 0,
        'pH granule (kg)': inv.treat_ph_granule || 0,
        'Antialgic (L)': inv.treat_antialgic || 0,
        'Anticalcar (L)': inv.treat_anticalcar || 0,
        'Floculant (L)': inv.treat_floculant || 0,
        'Sare (saci)': inv.treat_sare_saci || 0,
        'Bicarbonat (kg)': inv.treat_bicarbonat || 0,
        'Durata (min)': inv.duration_minutes || '',
        'Observatii': inv.observations || ''
      };
    });

    // Build header manually
    var ws1 = XLSX.utils.aoa_to_sheet([
      ['DEVIZ SERVICII PISCINA'],
      ['Client:', client.name, '', 'Nr. deviz:', devizNr],
      ['Adresa:', client.address || '-', '', 'Data:', today],
      ['Telefon:', client.phone || '-', '', 'Perioada:', (since || '-') + ' - ' + today],
      []
    ]);
    // Append data rows
    XLSX.utils.sheet_add_json(ws1, dataRows, { origin: 'A6' });

    // Add totals row
    var totals = calcTotals(sorted);
    var totalRow = sorted.length + 7; // header(5) + data header(1) + data rows
    XLSX.utils.sheet_add_aoa(ws1, [['', 'TOTAL:', sorted.length + ' interventii',
      totals.cl_granule_gr, totals.cl_tablete, totals.ph_granule,
      totals.antialgic, totals.anticalcar, totals.floculant,
      totals.sare, totals.bicarbonat,
      sorted.reduce(function(s, i) { return s + (i.duration_minutes || 0); }, 0) + ' min', ''
    ]], { origin: 'A' + totalRow });

    setColWidths(ws1, [6, 12, 16, 14, 14, 14, 12, 12, 12, 10, 14, 12, 28]);
    XLSX.utils.book_append_sheet(wb, ws1, 'Deviz');

    // Sheet 2: Detalii produse
    var sumRows = [
      { 'Produs': 'Cl granule', 'Cantitate': totals.cl_granule_gr, 'UM': 'gr' },
      { 'Produs': 'Cl tablete', 'Cantitate': totals.cl_tablete, 'UM': 'buc' },
      { 'Produs': 'pH granule', 'Cantitate': totals.ph_granule, 'UM': 'kg' },
      { 'Produs': 'Antialgic', 'Cantitate': totals.antialgic, 'UM': 'L' },
      { 'Produs': 'Anticalcar', 'Cantitate': totals.anticalcar, 'UM': 'L' },
      { 'Produs': 'Floculant', 'Cantitate': totals.floculant, 'UM': 'L' },
      { 'Produs': 'Sare', 'Cantitate': totals.sare, 'UM': 'saci' },
      { 'Produs': 'Bicarbonat', 'Cantitate': totals.bicarbonat, 'UM': 'kg' },
    ].filter(function(r) { return r.Cantitate > 0; });
    var ws2 = XLSX.utils.json_to_sheet(sumRows);
    setColWidths(ws2, [20, 12, 10]);
    XLSX.utils.book_append_sheet(wb, ws2, 'Produse');

    // Download
    var fname = 'Deviz_' + sanitizeFilename(client.name) + '_' + today.replace(/-/g, '') + '.xlsx';
    await _writeFileWithPicker(wb, fname, client.name);
    _uploadToDrive(wb, fname, null, client ? client.name : null);
    showToast('Deviz Excel descarcat: ' + fname, 'success');
  }).catch(function(e) {
    showToast('Eroare export: ' + e.message, 'error');
  });
}

