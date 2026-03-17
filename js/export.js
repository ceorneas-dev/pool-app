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

// == Dynamic chemical columns — built from stock products ==
// getChemColsFromStock(stockProducts) returns array of { key, label, productId }
function getChemColsFromStock(stockProducts) {
  if (!stockProducts || !stockProducts.length) return [];
  return stockProducts.filter(function(p) { return p.visible !== false; }).map(function(p) {
    return { key: 'treat_' + p.product_id, label: p.name, productId: p.product_id, unit: p.unit || '' };
  });
}

// ── Styled Deviz Constants (matching Python templates) ─────────────

// -- Border styles --
var S_MED    = { style: 'medium', color: { rgb: '000000' } };
var S_THIN_L = { style: 'thin',   color: { rgb: 'A8C8E8' } };
var S_THIN_M = { style: 'thin',   color: { rgb: '1E5FA8' } };
var S_THIN_N = { style: 'thin',   color: { rgb: '0D2D5A' } };
var S_DOT    = { style: 'dotted', color: { rgb: '4A7A99' } };

function _brd(top, bot, lft, rgt) {
  var b = {};
  if (top) b.top = top;
  if (bot) b.bottom = bot;
  if (lft) b.left = lft;
  if (rgt) b.right = rgt;
  return b;
}

// -- Fill colors --
var F_NAVY    = { fgColor: { rgb: '0D2D5A' } };
var F_MID     = { fgColor: { rgb: '1D507F' } };
var F_ACCENT  = { fgColor: { rgb: '4DB8E8' } };
var F_HEADER  = { fgColor: { rgb: '8DB4E2' } };
var F_HDRDARK = { fgColor: { rgb: '1F4E79' } };
var F_LIGHT1  = { fgColor: { rgb: 'E8F3FB' } };
var F_LIGHT2  = { fgColor: { rgb: 'EDF4FB' } };
var F_DATA_A  = { fgColor: { rgb: 'E0EEF8' } };
var F_DATA_BK = { fgColor: { rgb: 'F0F6FB' } };
var F_WHITE   = { fgColor: { rgb: 'FFFFFF' } };
// V2-specific fills
var F_HDRDARK2 = { fgColor: { rgb: '333F50' } };
var F_SUBHDR   = { fgColor: { rgb: '3560A0' } };
var F_DATA_E   = { fgColor: { rgb: 'D6E8F6' } };
var F_DATA_O   = { fgColor: { rgb: 'FFFFFF' } };
var F_TOT_DK   = { fgColor: { rgb: '223B6A' } };
var F_TOT_HDR  = { fgColor: { rgb: '333F50' } };

// -- Font presets --
function _fnt(name, sz, bold, color) {
  var f = { name: name || 'Arial', sz: sz || 10 };
  if (bold) f.bold = true;
  if (color) f.color = { rgb: color };
  return f;
}

// -- Company info --
var FIRMA_NUME    = 'S.C. AQUATIS ENGINEERING S.R.L.';
var FIRMA_ADRESA  = 'Str. Eufrosina Popescu 50, Sector 3';
var FIRMA_EMAIL   = 'office@aquatis.ro';
var FIRMA_WEB     = 'www.aquatis.ro';
var FIRMA_TELEFON = '0721.137.178';
var FIRMA_J       = 'J40/18144/2007';
var FIRMA_CUI     = 'RO22479681';
var FIRMA_IBAN    = 'RO77RNCB0074092331280001';

// -- Default prices per product_id (fallback if user hasn't set them) --
var CHEM_DEFAULT_PRICES = {
  cl_granule: 57, cl_tablete: 56.4, ph_minus_gr: 13, antialgic: 29,
  floculant: 25, bicarbonat: 32, ph_minus_l: 184, cl_lichid: 180,
  anticalcar: 25, sare: 15
};

// -- Page setup for landscape A4 --
function _setLandscapeA4(ws) {
  ws['!pageSetup'] = { paperSize: 9, orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 };
  ws['!margins'] = { left: 0.4, right: 0.4, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 };
  ws['!print'] = { gridLines: false };
}

// ── Styled Deviz Helpers (xlsx-js-style) ─────────────────────────

function _cellS(v, s) {
  if (v === null || v === undefined) v = '';
  var t = typeof v === 'number' ? 'n' : 's';
  return { t: t, v: v, s: s };
}

function _cellF(formula, s) {
  return { t: 'n', f: formula, s: s };
}

function _setRow(ws, rowIdx, values, style) {
  values.forEach(function(v, colIdx) {
    var ref = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    if (v !== null && v !== undefined) {
      var cell;
      if (typeof v === 'object' && v.hasOwnProperty('v')) {
        cell = v;
      } else if (typeof v === 'object' && v.hasOwnProperty('f')) {
        cell = v;
      } else {
        cell = { t: typeof v === 'number' ? 'n' : 's', v: v };
      }
      if (style && !cell.s) cell.s = style;
      ws[ref] = cell;
    } else if (style) {
      ws[ref] = { t: 's', v: '', s: style };
    }
  });
}

function _fillEmptyCells(ws, rowIdx, colCount, style) {
  for (var c = 0; c < colCount; c++) {
    var ref = XLSX.utils.encode_cell({ r: rowIdx, c: c });
    if (!ws[ref]) {
      ws[ref] = { t: 's', v: '', s: style };
    }
  }
}

function _setCell(ws, r, c, val, style) {
  var ref = XLSX.utils.encode_cell({ r: r, c: c });
  if (val !== null && val !== undefined && typeof val === 'object' && (val.hasOwnProperty('v') || val.hasOwnProperty('f'))) {
    ws[ref] = val;
  } else {
    var t = typeof val === 'number' ? 'n' : 's';
    ws[ref] = { t: t, v: (val === null || val === undefined) ? '' : val, s: style || {} };
  }
}

function _mergeFill(ws, merges, r, cStart, cEnd, val, style) {
  ws[XLSX.utils.encode_cell({ r: r, c: cStart })] = _cellS(val, style);
  for (var c = cStart + 1; c <= cEnd; c++) {
    ws[XLSX.utils.encode_cell({ r: r, c: c })] = _cellS('', style);
  }
  if (cEnd > cStart) merges.push({ s: { r: r, c: cStart }, e: { r: r, c: cEnd } });
}

// ── V1: Build Chimicale Sheet (dynamic chemicals from stock) ──────────
function _buildChimicaleSheet(client, sorted, prices, chemCols) {
  // chemCols = array of { key, label, productId, unit } from getChemColsFromStock()
  var numChem = chemCols.length;
  var COLS = 2 + numChem + 1; // A=Data, B=Cant, [chemicals], last=Total plată
  var lastCol = COLS - 1;
  var NR = sorted.length; // only rows with actual interventions
  var FR = 9;  // first data row (0-indexed) = row 10
  var ws = {};
  var merges = [];

  function colL(ci) { return XLSX.utils.encode_col(ci); }
  var lastColL = colL(lastCol);

  // Column widths
  var colWidths = [{ wch: 16 }, { wch: 8 }]; // A=date, B=cant
  for (var cw = 0; cw < numChem; cw++) colWidths.push({ wch: 10 });
  colWidths.push({ wch: 13 }); // Total plată
  ws['!cols'] = colWidths;

  // Row heights
  ws['!rows'] = [
    { hpt: 3.95 }, { hpt: 48 }, { hpt: 3 }, { hpt: 20.1 },
    { hpt: 18 }, { hpt: 15.95 }, { hpt: 3.95 }, { hpt: 26.1 }, { hpt: 32.1 }
  ];
  for (var dr = 0; dr < NR; dr++) ws['!rows'].push({ hpt: 18 });
  ws['!rows'].push({ hpt: 20.1 });  // Cantitate totală
  ws['!rows'].push({ hpt: 17.1 });  // Preț unitar
  ws['!rows'].push({ hpt: 21.95 }); // TOTAL GENERAL
  ws['!rows'].push({ hpt: 20.25 }); // Footer

  // Date helpers
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
  var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr = 'AQS - ' + todayYMD;

  // Dynamic merge ranges for header info (split into ~quarters)
  var labQ = Math.max(Math.floor(COLS / 4), 1);
  var lab1End = Math.min(labQ - 1, lastCol);
  var lab2End = Math.min(labQ * 2 - 1, lastCol);
  var lab3End = Math.min(labQ * 3 - 1, lastCol);

  // Company info merge ranges (split into ~thirds)
  var third = Math.max(Math.floor(COLS / 3), 1);
  var s1End = Math.min(third - 1, lastCol);
  var s2End = Math.min(third * 2 - 1, lastCol);

  // ═══ ROW 1 (idx 0): Navy bar ═══
  _mergeFill(ws, merges, 0, 0, lastCol, '', { fill: F_NAVY, border: _brd(S_MED, null, S_MED, S_MED) });

  // ═══ ROW 2 (idx 1): Company info ═══
  _mergeFill(ws, merges, 1, 0, s1End, FIRMA_NUME + '\n' + FIRMA_ADRESA,
    { fill: F_HEADER, font: _fnt('Arial', 11, true), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_MED, S_THIN_L) });
  _mergeFill(ws, merges, 1, s1End + 1, s2End, FIRMA_EMAIL + '\n' + FIRMA_WEB + '\n' + FIRMA_TELEFON,
    { fill: F_HEADER, font: _fnt('Arial', 9, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_THIN_L, S_THIN_L) });
  _mergeFill(ws, merges, 1, s2End + 1, lastCol, FIRMA_J + '\nCUI: ' + FIRMA_CUI + '\nIBAN: ' + FIRMA_IBAN,
    { fill: F_HEADER, font: _fnt('Arial', 8.5, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_THIN_L, S_MED) });

  // ═══ ROW 3 (idx 2): Accent bar ═══
  _mergeFill(ws, merges, 2, 0, lastCol, '', { fill: F_ACCENT, border: _brd(null, null, S_MED, S_MED) });

  // ═══ ROW 4 (idx 3): Title ═══
  _mergeFill(ws, merges, 3, 0, lastCol, 'RAPORT INTERVEN\u021AII \u2014 CHIMICALE PISCIN\u0102',
    { fill: F_MID, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(null, null, S_MED, S_MED) });

  // ═══ ROW 5 (idx 4): Labels ═══
  var sLbl5 = { fill: F_LIGHT1, font: _fnt('Arial', 8, true, '404040'), alignment: { horizontal: 'left', vertical: 'center' } };
  _mergeFill(ws, merges, 4, 0, lab1End, 'Client',
    Object.assign({}, sLbl5, { border: _brd(null, null, S_MED, S_THIN_L) }));
  _mergeFill(ws, merges, 4, lab1End + 1, lab2End, 'Perioada raportata',
    Object.assign({}, sLbl5, { border: _brd(null, null, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 4, lab2End + 1, lab3End, 'Nr. Document',
    Object.assign({}, sLbl5, { border: _brd(null, null, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 4, lab3End + 1, lastCol, 'Data emiterii',
    Object.assign({}, sLbl5, { border: _brd(null, null, S_THIN_L, S_MED) }));

  // ═══ ROW 6 (idx 5): Values ═══
  var sVal6 = { fill: F_WHITE, font: _fnt('Arial', 10, true, '0D2D5A'), alignment: { horizontal: 'left', vertical: 'center' } };
  _mergeFill(ws, merges, 5, 0, lab1End, client.name || '',
    Object.assign({}, sVal6, { border: _brd(null, S_DOT, S_MED, S_THIN_L) }));
  _mergeFill(ws, merges, 5, lab1End + 1, lab2End, period,
    Object.assign({}, sVal6, { border: _brd(null, S_DOT, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 5, lab2End + 1, lab3End, docNr,
    Object.assign({}, sVal6, { border: _brd(null, S_DOT, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 5, lab3End + 1, lastCol, todayStr,
    Object.assign({}, sVal6, { border: _brd(null, S_DOT, S_THIN_L, S_MED) }));

  // ═══ ROW 7 (idx 6): Separator ═══
  _mergeFill(ws, merges, 6, 0, lastCol, '', { fill: F_LIGHT1, border: _brd(null, null, S_MED, S_MED) });

  // ═══ ROW 8 (idx 7): Header row 1 ═══
  var brd89 = _brd(S_MED, S_THIN_N, S_MED, S_MED);
  // A8:A9 "Data interventie" merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 0 })] = _cellS('Data\ninterven\u021Bie',
    { fill: F_HDRDARK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: brd89 });
  ws[XLSX.utils.encode_cell({ r: 8, c: 0 })] = _cellS('',
    { fill: F_HDRDARK, border: _brd(null, S_THIN_N, S_MED, S_MED) });
  merges.push({ s: { r: 7, c: 0 }, e: { r: 8, c: 0 } });

  // B8:B9 "Cant." merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 1 })] = _cellS('Cant.',
    { fill: F_HDRDARK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: brd89 });
  ws[XLSX.utils.encode_cell({ r: 8, c: 1 })] = _cellS('',
    { fill: F_HDRDARK, border: _brd(null, S_THIN_N, S_MED, S_MED) });
  merges.push({ s: { r: 7, c: 1 }, e: { r: 8, c: 1 } });

  // C8 to secondLastCol: "CHIMICALE FOLOSITE" merged
  var chemLastCol = 2 + numChem - 1; // last chemical column
  ws[XLSX.utils.encode_cell({ r: 7, c: 2 })] = _cellS('CHIMICALE FOLOSITE',
    { fill: F_HDRDARK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_MED, S_MED, S_THIN_M) });
  for (var hc = 3; hc <= chemLastCol; hc++) {
    ws[XLSX.utils.encode_cell({ r: 7, c: hc })] = _cellS('', { fill: F_HDRDARK, border: _brd(S_MED, S_MED, null, null) });
  }
  merges.push({ s: { r: 7, c: 2 }, e: { r: 7, c: chemLastCol } });

  // Last col header: "Total plată" merged rows 8-9
  ws[XLSX.utils.encode_cell({ r: 7, c: lastCol })] = _cellS('Total\nplat\u0103',
    { fill: F_HDRDARK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: brd89 });
  ws[XLSX.utils.encode_cell({ r: 8, c: lastCol })] = _cellS('',
    { fill: F_HDRDARK, border: _brd(null, S_THIN_N, S_MED, S_MED) });
  merges.push({ s: { r: 7, c: lastCol }, e: { r: 8, c: lastCol } });

  // ═══ ROW 9 (idx 8): Sub-headers (chemical names) ═══
  for (var ci = 0; ci < numChem; ci++) {
    var bl = ci > 0 ? S_THIN_M : null;
    var br = ci < numChem - 1 ? S_THIN_M : null;
    ws[XLSX.utils.encode_cell({ r: 8, c: ci + 2 })] = _cellS(chemCols[ci].label,
      { fill: F_HEADER, font: _fnt('Arial', 8.5, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _brd(null, S_MED, bl, br) });
  }

  // ═══ DATA ROWS (idx 9+): alternating fills ═══
  var LR = FR + NR - 1; // last data row index
  for (var di = 0; di < NR; di++) {
    var rowIdx = FR + di;
    var isEven = (di % 2 === 0);
    var isFirst = (di === 0);
    var isLastRow = (di === NR - 1);
    var fillA = isEven ? F_DATA_A : F_WHITE;
    var fillBK = isEven ? F_DATA_BK : F_WHITE;
    var topB = isFirst ? S_MED : S_THIN_L;
    var botB = isLastRow ? S_MED : S_THIN_L;

    var entry = di < sorted.length ? sorted[di] : {};

    // A: date
    ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS(entry.date ? fmtDateDMY(entry.date) : '',
      { fill: fillA, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(topB, botB, S_MED, S_THIN_L) });

    // B: cant (always 1 per intervention)
    ws[XLSX.utils.encode_cell({ r: rowIdx, c: 1 })] = _cellS(entry.date ? 1 : '',
      { fill: fillBK, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_THIN_L, S_THIN_L) });

    // Chemical columns
    for (var cc = 0; cc < numChem; cc++) {
      var cVal = entry.date ? (parseFloat(entry[chemCols[cc].key]) || 0) : '';
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: cc + 2 })] = _cellS(cVal > 0 ? cVal : (entry.date ? '' : ''),
        { fill: fillBK, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_THIN_L, S_THIN_L) });
    }

    // Last col: Total plată (empty cell)
    ws[XLSX.utils.encode_cell({ r: rowIdx, c: lastCol })] = _cellS('',
      { fill: fillBK, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_THIN_L, S_MED) });
  }

  // ═══ Cantitate totală row ═══
  var totRow = FR + NR; // row after data
  var firstDataExcel = FR + 1; // Excel row number (1-indexed)
  var lastDataExcel = FR + NR;  // Excel row number (1-indexed)

  ws[XLSX.utils.encode_cell({ r: totRow, c: 0 })] = _cellS('Cantitate total\u0103',
    { fill: F_HEADER, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(S_MED, S_THIN_N, S_MED, S_THIN_N) });
  // B: SUM of Cant column
  ws[XLSX.utils.encode_cell({ r: totRow, c: 1 })] = _cellF('SUM(B' + firstDataExcel + ':B' + lastDataExcel + ')',
    { fill: F_HEADER, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_N, S_THIN_N, S_THIN_N) });
  for (var sc = 0; sc < numChem; sc++) {
    var colL20 = colL(sc + 2);
    ws[XLSX.utils.encode_cell({ r: totRow, c: sc + 2 })] = _cellF('SUM(' + colL20 + firstDataExcel + ':' + colL20 + lastDataExcel + ')',
      { fill: F_HEADER, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_N, S_THIN_N, S_THIN_N) });
  }
  // Last col in totRow: same blue as rest of row
  ws[XLSX.utils.encode_cell({ r: totRow, c: lastCol })] = _cellS('',
    { fill: F_HEADER, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_N, S_THIN_N, S_MED) });

  // ═══ Preț unitar row ═══
  var pretRow = totRow + 1;
  var totRowExcel = totRow + 1; // Excel row of cantitate totala
  ws[XLSX.utils.encode_cell({ r: pretRow, c: 0 })] = _cellS('Pre\u021B unitar',
    { fill: F_LIGHT2, font: _fnt('Arial', 8.5, false, '0D2D5A'), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(S_THIN_L, S_THIN_L, S_MED, S_THIN_L) });
  // B: preț per intervenție (per client)
  var pretIntv = parseFloat(client.pret_interventie) || prices.pret_interventie || 0;
  ws[XLSX.utils.encode_cell({ r: pretRow, c: 1 })] = _cellS(pretIntv > 0 ? pretIntv : '',
    { fill: F_LIGHT2, font: _fnt('Arial', 8.5, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L) });
  for (var pc = 0; pc < numChem; pc++) {
    var prc = prices[chemCols[pc].productId] || CHEM_DEFAULT_PRICES[chemCols[pc].productId] || 0;
    ws[XLSX.utils.encode_cell({ r: pretRow, c: pc + 2 })] = _cellS(prc > 0 ? prc : '',
      { fill: F_LIGHT2, font: _fnt('Arial', 8.5, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L) });
  }
  ws[XLSX.utils.encode_cell({ r: pretRow, c: lastCol })] = _cellS('',
    { fill: F_LIGHT2, border: _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_MED) });

  // ═══ TOTAL GENERAL row ═══
  var genRow = pretRow + 1;
  var pretRowExcel = pretRow + 1; // Excel row of pret unitar
  ws[XLSX.utils.encode_cell({ r: genRow, c: 0 })] = _cellS('TOTAL GENERAL',
    { fill: F_MID, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_MED, S_THIN_M) });
  // B: =cantitate_totala * pret_unitar
  ws[XLSX.utils.encode_cell({ r: genRow, c: 1 })] = _cellF('B' + totRowExcel + '*B' + pretRowExcel,
    { fill: F_MID, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_THIN_M, S_THIN_M) });
  for (var gc = 0; gc < numChem; gc++) {
    var gCol = colL(gc + 2);
    ws[XLSX.utils.encode_cell({ r: genRow, c: gc + 2 })] = _cellF(gCol + totRowExcel + '*' + gCol + pretRowExcel,
      { fill: F_MID, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_THIN_M, S_THIN_M) });
  }
  // Last col: SUM of all TOTAL GENERAL cells in row
  var genRowExcel = genRow + 1;
  ws[XLSX.utils.encode_cell({ r: genRow, c: lastCol })] = _cellF('SUM(B' + genRowExcel + ':' + colL(lastCol - 1) + genRowExcel + ')',
    { fill: F_MID, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_N, S_THIN_N, S_THIN_N, S_MED) });

  // ═══ Footer row ═══
  var footRow = genRow + 1;
  var footHalf = Math.floor(COLS / 2);
  _mergeFill(ws, merges, footRow, 0, footHalf - 1, 'Toate pre\u021Burile sunt exprimate \u00EEn RON',
    { fill: F_HEADER, font: _fnt('Arial', 7.5, false), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(null, S_MED, S_MED, null) });
  _mergeFill(ws, merges, footRow, footHalf, lastCol, 'S.C. Aquatis Engineering S.R.L.',
    { fill: F_HEADER, font: _fnt('Arial', 7.5, false), alignment: { horizontal: 'right', vertical: 'center' }, border: _brd(null, S_MED, null, S_MED) });

  ws['!ref'] = 'A1:' + lastColL + (footRow + 1);
  ws['!merges'] = merges;
  _setLandscapeA4(ws);
  return ws;
}

// ── V2: Build Servicii Sheet (exact match Python Template V2.py) ──────────
function _buildServiciiSheet(client, sorted, totalPlata, opsList) {
  // Build operations list: start with defaults, add any extras from interventions
  var defaultOps = [
    'Aspirare piscina', 'Curatare linie apa', 'Curatare skimmere',
    'Spalare filtru', 'Curatare prefiltru', 'Periere piscina',
    'Analiza apei', 'Tratament chimic'
  ];
  var allOps = (opsList && opsList.length) ? opsList.slice() : defaultOps.slice();
  sorted.forEach(function(intv) {
    (intv.operations || []).forEach(function(op) {
      if (op && allOps.indexOf(op) < 0) allOps.push(op);
    });
  });

  var numOps = allOps.length;
  var NR = sorted.length; // only rows with actual interventions
  var FR = 9;            // first data row 0-indexed
  var LR = 25;           // last data row 0-indexed
  var NCOLS = 1 + numOps; // A + ops columns
  var LC = NCOLS - 1;     // last col index (0-based)
  var ws = {};
  var merges = [];

  // Display labels for known operations
  var knownLabels = {
    'Aspirare piscina': 'Aspirare\npiscin\u0103',
    'Curatare linie apa': 'Cur\u0103\u021Bare\nlinie ap\u0103',
    'Curatare skimmere': 'Cur\u0103\u021Bare\nskimmere',
    'Spalare filtru': 'Sp\u0103lare\nfiltru',
    'Curatare prefiltru': 'Cur\u0103\u021Bare\nprefiltru',
    'Periere piscina': 'Periere\npiscin\u0103',
    'Analiza apei': 'Analiza\napei',
    'Tratament chimic': 'Tratament\nchimic',
    'Verificare automatizare': 'Verificare\nautomatizare'
  };

  // Column widths (matching Python: A=13, then specific per op)
  var defaultWidths = [13, 10.5, 11.5, 10.5, 9.5, 11, 9.5, 9.5, 11];
  var colWidths = [];
  for (var cw = 0; cw < NCOLS; cw++) {
    colWidths.push({ wch: cw < defaultWidths.length ? defaultWidths[cw] : 11 });
  }
  ws['!cols'] = colWidths;

  // Row heights
  ws['!rows'] = [
    { hpt: 3.95 },  { hpt: 52 },   { hpt: 3 },     { hpt: 20.1 },
    { hpt: 15.95 }, { hpt: 17.1 }, { hpt: 3.95 },  { hpt: 21.95 },
    { hpt: 45.95 }
  ];
  for (var dr = 0; dr < NR; dr++) ws['!rows'].push({ hpt: 20.1 });
  ws['!rows'].push({ hpt: 20.1 }); // row 26 (27)
  ws['!rows'].push({ hpt: 6 });    // row 27 (28)
  ws['!rows'].push({ hpt: 24 });   // row 28 (29)
  ws['!rows'].push({ hpt: 15 });   // row 29 (30)

  // Date helpers
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
  var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr = 'AQS - ' + todayYMD;
  function colLetter(ci) { return XLSX.utils.encode_col(ci); }
  var lastColLetter = colLetter(LC);

  // V2 uses S_THIN_K (thin black) for table borders (matching Python)
  var S_THIN_K = { style: 'thin', color: { rgb: '000000' } };

  // ═══ ROW 1 (idx 0): Navy bar ═══
  _mergeFill(ws, merges, 0, 0, LC, '', { fill: F_NAVY, border: _brd(S_MED, null, S_MED, S_MED) });

  // ═══ ROW 2 (idx 1): Company info ═══ (Python: A2:C2, D2:F2, G2:I2)
  // For 9 cols: thirds are 0-2, 3-5, 6-8. For dynamic cols, calculate proportionally
  var t = Math.max(Math.floor(NCOLS / 3), 1);
  var s1E = Math.min(t - 1, LC);
  var s2E = Math.min(t * 2 - 1, LC);
  _mergeFill(ws, merges, 1, 0, s1E, FIRMA_NUME + '\n' + FIRMA_ADRESA,
    { fill: F_HEADER, font: _fnt('Arial', 9, true), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_MED, S_THIN_L) });
  _mergeFill(ws, merges, 1, s1E + 1, s2E, '  ' + FIRMA_EMAIL + '\n  ' + FIRMA_WEB + '\n  ' + FIRMA_TELEFON,
    { fill: F_HEADER, font: _fnt('Arial', 9, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_THIN_L, S_THIN_L) });
  _mergeFill(ws, merges, 1, s2E + 1, LC, '  ' + FIRMA_J + '\n  CUI: ' + FIRMA_CUI + '\n  ' + FIRMA_IBAN,
    { fill: F_HEADER, font: _fnt('Arial', 8.5, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_THIN_L, S_MED) });

  // ═══ ROW 3 (idx 2): Accent bar ═══
  _mergeFill(ws, merges, 2, 0, LC, '', { fill: F_ACCENT, border: _brd(null, S_MED, S_MED, S_MED) });

  // ═══ ROW 4 (idx 3): Title ═══ (Python: FILL_NAVY, not FILL_MID)
  _mergeFill(ws, merges, 3, 0, LC, 'RAPORT SERVICII \u2014 ABONAMENT \u00CENTRE\u021AINERE PISCIN\u0102',
    { fill: F_NAVY, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(null, null, S_MED, S_MED) });

  // ═══ ROW 5 (idx 4): Labels ═══ (Python: A5:B5, C5:E5, F5:G5, H5 separate, I5 separate)
  var sLbl5v2 = { fill: F_LIGHT1, font: _fnt('Arial', 8, true, '2E5C8A'), alignment: { horizontal: 'left', vertical: 'center' } };
  // For 9 cols: A:B(0-1), C:E(2-4), F:G(5-6), H(7), I(8)
  // Dynamic: proportional
  var q5_1 = Math.max(Math.floor(NCOLS / 4.5), 1);  // ~2 cols
  var q5_2 = Math.max(Math.floor(NCOLS / 3), 2);     // ~3 cols
  var q5_3 = Math.max(Math.floor(NCOLS / 4.5), 1);   // ~2 cols
  var l5_1E = Math.min(q5_1, LC);                     // end of Client
  var l5_2E = Math.min(l5_1E + q5_2, LC);             // end of Perioada
  var l5_3E = Math.min(l5_2E + q5_3, LC);             // end of Nr. Doc
  // For default 9 cols: 0-1, 2-4, 5-6, 7, 8
  if (NCOLS === 9) { l5_1E = 1; l5_2E = 4; l5_3E = 6; }

  _mergeFill(ws, merges, 4, 0, l5_1E, 'Client',
    Object.assign({}, sLbl5v2, { border: _brd(null, null, S_MED, S_THIN_L) }));
  _mergeFill(ws, merges, 4, l5_1E + 1, l5_2E, 'Perioada raportare',
    Object.assign({}, sLbl5v2, { border: _brd(null, null, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 4, l5_2E + 1, l5_3E, 'Nr. Document',
    Object.assign({}, sLbl5v2, { border: _brd(null, null, S_THIN_L, S_THIN_L) }));
  // Data Emiterii — H5 separate, I5 NOT merged (keeps medium right)
  ws[XLSX.utils.encode_cell({ r: 4, c: l5_3E + 1 })] = _cellS('Data Emiterii',
    Object.assign({}, sLbl5v2, { border: _brd(null, null, S_THIN_L, null) }));
  if (LC > l5_3E + 1) {
    ws[XLSX.utils.encode_cell({ r: 4, c: LC })] = _cellS('', { fill: F_LIGHT1, border: _brd(null, null, null, S_MED) });
  }

  // ═══ ROW 6 (idx 5): Values ═══ (Python: A6:B6, C6:E6, F6:G6, H6 separate, I6 separate)
  var sVal6v2 = { fill: F_WHITE, font: _fnt('Arial', 10, true, '0D2D5A'), alignment: { horizontal: 'left', vertical: 'center' } };
  _mergeFill(ws, merges, 5, 0, l5_1E, client.name || '',
    Object.assign({}, sVal6v2, { border: _brd(null, S_DOT, S_MED, S_THIN_L) }));
  _mergeFill(ws, merges, 5, l5_1E + 1, l5_2E, period,
    Object.assign({}, sVal6v2, { border: _brd(null, S_DOT, S_THIN_L, S_THIN_L) }));
  _mergeFill(ws, merges, 5, l5_2E + 1, l5_3E, docNr,
    Object.assign({}, sVal6v2, { border: _brd(null, S_DOT, S_THIN_L, S_THIN_L) }));
  ws[XLSX.utils.encode_cell({ r: 5, c: l5_3E + 1 })] = _cellS(todayStr,
    Object.assign({}, sVal6v2, { border: _brd(null, S_DOT, S_THIN_L, null) }));
  if (LC > l5_3E + 1) {
    ws[XLSX.utils.encode_cell({ r: 5, c: LC })] = _cellS('', { fill: F_WHITE, border: _brd(null, S_DOT, null, S_MED) });
  }

  // ═══ ROW 7 (idx 6): Separator ═══
  _mergeFill(ws, merges, 6, 0, LC, '', { fill: F_LIGHT1, border: _brd(null, null, S_MED, S_MED) });

  // ═══ ROW 8 (idx 7): Header row 1 ═══ (Python: FILL_TOT_DK, S_THIN_K borders)
  // A8:A9 merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 0 })] = _cellS('Data\ninterven\u021Bie',
    { fill: F_TOT_DK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_K, S_MED, S_THIN_K) });
  ws[XLSX.utils.encode_cell({ r: 8, c: 0 })] = _cellS('',
    { fill: F_TOT_DK, border: _brd(null, S_THIN_K, S_MED, S_THIN_K) });
  merges.push({ s: { r: 7, c: 0 }, e: { r: 8, c: 0 } });

  // B8:H8 (or B8 to LC-1) "SERVICII INCLUSE ÎN ABONAMENT" merged (Python: B8:H8, I8 NOT merged)
  var svcEnd = Math.max(LC - 1, 1);
  ws[XLSX.utils.encode_cell({ r: 7, c: 1 })] = _cellS('SERVICII INCLUSE \u00CEN ABONAMENT',
    { fill: F_TOT_DK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_K, S_THIN_K, S_THIN_K) });
  for (var hc = 2; hc <= svcEnd; hc++) {
    ws[XLSX.utils.encode_cell({ r: 7, c: hc })] = _cellS('', { fill: F_TOT_DK, border: _brd(S_MED, S_THIN_K, null, null) });
  }
  if (svcEnd > 1) merges.push({ s: { r: 7, c: 1 }, e: { r: 7, c: svcEnd } });
  // Last col (I8) NOT merged — keeps medium right border
  if (LC > svcEnd) {
    ws[XLSX.utils.encode_cell({ r: 7, c: LC })] = _cellS('', { fill: F_TOT_DK, border: _brd(S_MED, S_THIN_K, null, S_MED) });
  }

  // ═══ ROW 9 (idx 8): Sub-headers ═══ (Python: FILL_SUBHDR)
  for (var si = 0; si < numOps; si++) {
    var opName = allOps[si];
    var label = knownLabels[opName] || opName.replace(/ /g, '\n');
    var brR = (si + 1 === LC) ? S_MED : S_THIN_K;
    ws[XLSX.utils.encode_cell({ r: 8, c: si + 1 })] = _cellS(label,
      { fill: F_SUBHDR, font: _fnt('Arial', 8, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _brd(S_THIN_K, null, S_THIN_K, brR) });
  }

  // ═══ ROWS 10-26 (idx 9-25): Data rows (alternating) ═══
  for (var di = 0; di < NR; di++) {
    var rowIdx = FR + di;
    var isEven = (di % 2 === 0);
    var isFirst = (di === 0);
    var isLast = (di === NR - 1);
    var fillRow = isEven ? F_DATA_E : F_DATA_O;
    var topB = isFirst ? S_MED : S_THIN_L;
    var botB = isLast ? S_MED : S_THIN_L;

    var entry = di < sorted.length ? sorted[di] : {};

    // A: date
    ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS(entry.date ? fmtDateDMY(entry.date) : '',
      { fill: fillRow, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_MED, S_THIN_L) });

    // B onwards: check services by exact match
    var ops = entry.operations || [];
    for (var sc = 0; sc < numOps; sc++) {
      var matched = entry.date ? (ops.indexOf(allOps[sc]) >= 0) : false;
      var brR2 = (sc + 1 === LC) ? S_MED : S_THIN_L;
      if (matched) {
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: sc + 1 })] = _cellS('\u2713',
          { fill: fillRow, font: _fnt('Arial', 12, true, '1A6B2A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_THIN_L, brR2) });
      } else {
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: sc + 1 })] = _cellS('',
          { fill: fillRow, font: _fnt('Arial', 9, false, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(topB, botB, S_THIN_L, brR2) });
      }
    }
  }

  // ═══ Total intervenții efectuate row ═══
  var totRow2 = FR + NR; // row right after data
  var firstDataExcel2 = FR + 1; // Excel row (1-indexed)
  var lastDataExcel2 = FR + NR;

  // A: label "Total intervenții"
  ws[XLSX.utils.encode_cell({ r: totRow2, c: 0 })] = _cellS('Total interven\u021Bii',
    { fill: F_SUBHDR, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'right', vertical: 'center' }, border: _brd(null, S_THIN_N, S_MED, S_THIN_N) });
  // B: count value
  ws[XLSX.utils.encode_cell({ r: totRow2, c: 1 })] = _cellF('COUNTA(A' + firstDataExcel2 + ':A' + lastDataExcel2 + ')',
    { fill: F_SUBHDR, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(null, S_THIN_N, S_THIN_N, S_THIN_N) });
  // Fill remaining cols with same style
  for (var tc = 2; tc <= LC; tc++) {
    ws[XLSX.utils.encode_cell({ r: totRow2, c: tc })] = _cellS('',
      { fill: F_SUBHDR, font: _fnt('Arial', 9, true, 'FFFFFF'), border: _brd(null, S_THIN_N, S_THIN_N, tc === LC ? S_MED : S_THIN_N) });
  }

  // ═══ Separator row ═══
  var sepRow2 = totRow2 + 1;
  _mergeFill(ws, merges, sepRow2, 0, LC, '', { fill: F_LIGHT1, border: _brd(null, null, S_MED, S_MED) });

  // ═══ TOTAL DE PLATĂ row ═══
  var payRow2 = sepRow2 + 1;
  var payLblEnd = Math.max(Math.floor(NCOLS * 2 / 3) - 1, 0);
  if (NCOLS === 9) payLblEnd = 5;
  var payValStart = payLblEnd + 1;
  _mergeFill(ws, merges, payRow2, 0, payLblEnd, 'TOTAL DE PLAT\u0102 (RON)',
    { fill: F_TOT_DK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_MED, S_THIN_M) });
  _mergeFill(ws, merges, payRow2, payValStart, LC, totalPlata || '',
    { fill: F_TOT_DK, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_MED, S_MED, S_MED) });

  // ═══ Footer row ═══
  var footRow2 = payRow2 + 1;
  var footEnd1 = Math.max(Math.floor(NCOLS * 5 / 9) - 1, 0);
  if (NCOLS === 9) footEnd1 = 4;
  _mergeFill(ws, merges, footRow2, 0, footEnd1, '  Document generat de S.C. Aquatis Engineering S.R.L.',
    { fill: F_HEADER, font: _fnt('Arial', 7.5, false), alignment: { horizontal: 'left', vertical: 'center' }, border: _brd(null, S_MED, S_MED, null) });
  _mergeFill(ws, merges, footRow2, footEnd1 + 1, LC, FIRMA_WEB + '  |  ' + FIRMA_TELEFON + '  ',
    { fill: F_HEADER, font: _fnt('Arial', 7.5, false), alignment: { horizontal: 'right', vertical: 'center' }, border: _brd(null, S_MED, null, S_MED) });

  ws['!ref'] = 'A1:' + lastColLetter + (footRow2 + 1);
  ws['!merges'] = merges;
  _setLandscapeA4(ws);
  return ws;
}

// ── Export Deviz Chimicale (V1 only) ───────────────────────────────
function exportDevizChimicale(client, interventions) {
  return loadXLSX().then(async function() {
    var sorted = interventions.slice().sort(function(a, b) { return String(a.date).localeCompare(String(b.date)); });
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var stockProducts = (typeof getAllStock === 'function') ? await getAllStock() : [];
    var chemCols = getChemColsFromStock(stockProducts);

    var wb = XLSX.utils.book_new();
    var ws = _buildChimicaleSheet(client, sorted, prices, chemCols);
    var sheetName = sanitizeSheetName(client.name || 'Chimicale');
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    var fname = sanitizeFilename(client.name) + '_Chimicale_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, fname, client.name);
    _uploadToDrive(wb, fname, null, client.name);
    return fname;
  });
}

// ── Export Deviz Complet (V1 + V2 in same workbook) ────────────────
function exportDevizComplet(client, interventions) {
  return loadXLSX().then(async function() {
    var sorted = interventions.slice().sort(function(a, b) { return String(a.date).localeCompare(String(b.date)); });
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var opsList = (typeof getOperations === 'function') ? await getOperations() : null;
    var stockProducts = (typeof getAllStock === 'function') ? await getAllStock() : [];
    var chemCols = getChemColsFromStock(stockProducts);

    var wb = XLSX.utils.book_new();

    // Sheet 1: Chimicale (V1)
    var ws1 = _buildChimicaleSheet(client, sorted, prices, chemCols);
    var name1 = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Chim');
    XLSX.utils.book_append_sheet(wb, ws1, name1);

    // Sheet 2: Servicii (V2)
    var ws2 = _buildServiciiSheet(client, sorted, '', opsList);
    var name2 = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Serv');
    XLSX.utils.book_append_sheet(wb, ws2, name2);

    var fname = sanitizeFilename(client.name) + '_Deviz_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, fname, client.name);
    _uploadToDrive(wb, fname, null, client.name);
    return fname;
  });
}

// ── Export All Deviz Mixed (all clients) ───────────────────────────
function exportAllDevizMixed(clients, allInterventions, filter) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var opsList = (typeof getOperations === 'function') ? await getOperations() : null;
    var stockProducts = (typeof getAllStock === 'function') ? await getAllStock() : [];
    var chemCols = getChemColsFromStock(stockProducts);
    var wb = XLSX.utils.book_new();
    var sheetCount = 0;

    // allInterventions can be array or object — normalize to object keyed by client_id
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

    clients.forEach(function(client) {
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

      if (clientIntv.length === 0) return;

      var baseName = sanitizeSheetName(client.name || 'Client');
      var devizType = parseInt(client.deviz_type) || 2;

      if (devizType === 2) {
        // V2 = complet (both sheets)
        var ws1 = _buildChimicaleSheet(client, clientIntv, prices, chemCols);
        var chemName = baseName.substring(0, 28) + '_Ch';
        if (wb.SheetNames.indexOf(chemName) >= 0) chemName = chemName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws1, chemName);
        sheetCount++;

        var ws2 = _buildServiciiSheet(client, clientIntv, '', opsList);
        var opsName = baseName.substring(0, 28) + '_Sv';
        if (wb.SheetNames.indexOf(opsName) >= 0) opsName = opsName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws2, opsName);
        sheetCount++;
      } else {
        // V1 = chimicale only
        var ws1v = _buildChimicaleSheet(client, clientIntv, prices, chemCols);
        var chemNameV = baseName.substring(0, 28) + '_Ch';
        if (wb.SheetNames.indexOf(chemNameV) >= 0) chemNameV = chemNameV.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws1v, chemNameV);
        sheetCount++;
      }
    });

    if (sheetCount === 0) {
      showToast('Nicio interventie de exportat.', 'warning');
      return;
    }

    var fname = 'DevizToti_' + fmtDateExport(new Date()) + '.xlsx';
    await _writeFileWithPicker(wb, fname);
    _uploadToDrive(wb, fname, null, null);
    return fname;
  });
}

function _uploadToDrive(wb, fileName, mimeType, clientName) {
  if (typeof isSyncConfigured !== 'function' || !isSyncConfigured()) return;
  try {
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
        clientName: clientName || ''
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

