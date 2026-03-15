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

// ── Export per client ─────────────────────────────────────────
function exportClientXLSX(client, interventions) {
  return loadXLSX().then(() => {
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
    XLSX.writeFile(wb, filename);
    _uploadToDrive(wb, filename, null, client.name);
    return filename;
  });
}

// ── Export all clients ────────────────────────────────────────
function exportAllXLSX(clients, allInterventions) {
  return loadXLSX().then(() => {
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
    XLSX.writeFile(wb, filename);
    _uploadToDrive(wb, filename, null, null);
    return filename;
  });
}

// ── Export structured (one sheet per client, no summary) ──────
function exportStructuredXLSX(clients, allInterventions) {
  return loadXLSX().then(() => {
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
    XLSX.writeFile(wb, filename);
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
  var parts = String(dateStr).split('-');
  if (parts.length === 3) return parts[2] + '.' + parts[1] + '.' + parts[0];
  return dateStr;
}

// == All possible chemical columns ==
var ALL_CHEM_COLS = [
  { key: 'treat_cl_granule_gr',       label: 'Clor Rapid',  priceKey: 'clor_rapid' },
  { key: 'treat_cl_tablete_export_gr',label: 'Clor Lent',   priceKey: 'clor_lent' },
  { key: 'treat_ph_granule',          label: 'pH-',         priceKey: 'ph_minus' },
  { key: 'treat_antialgic',           label: 'Antialgic',   priceKey: 'antialgic' },
  { key: 'treat_floculant',           label: 'Floculant',   priceKey: 'floculant' },
  { key: 'treat_bicarbonat',          label: 'Dedurizant',  priceKey: 'dedurizant' },
  { key: 'treat_ph_lichid_bidoane',   label: 'Ph Lichid',   priceKey: 'ph_lichid' },
  { key: 'treat_cl_lichid_bidoane',   label: 'Cl Lichid',   priceKey: 'cl_lichid' }
];

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

// -- Default prices --
var DEFAULT_PRICES = {
  clor_rapid: 57, clor_lent: 56.4, ph_minus: 13, antialgic: 29,
  floculant: 25, dedurizant: 32, ph_lichid: 184, cl_lichid: 180
};

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

// ── V1: Build Chimicale Sheet ──────────────────────────────────────
function _buildChimicaleSheet(client, sorted, prices) {
  var NR = 10; // fixed 10 data rows
  var COLS = 11; // A-K
  var ws = {};
  var merges = [];

  // Column widths A-K
  ws['!cols'] = [
    { wch: 16 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 8 },
    { wch: 10 }, { wch: 8 }, { wch: 11 }, { wch: 10 }, { wch: 8 }, { wch: 13 }
  ];

  // Row heights (0-indexed)
  ws['!rows'] = [
    { hpt: 3.95 },  // row 0 (1)
    { hpt: 48 },     // row 1 (2)
    { hpt: 3 },      // row 2 (3)
    { hpt: 20.1 },   // row 3 (4)
    { hpt: 18 },     // row 4 (5)
    { hpt: 15.95 },  // row 5 (6)
    { hpt: 3.95 },   // row 6 (7)
    { hpt: 26.1 },   // row 7 (8)
    { hpt: 32.1 }    // row 8 (9)
  ];
  // rows 9-18: data rows (height 18)
  for (var dr = 0; dr < NR; dr++) ws['!rows'].push({ hpt: 18 });
  // row 19 (20): 20.1, row 20 (21): 17.1, row 21 (22): 21.95, row 22 (23): 20.25
  ws['!rows'].push({ hpt: 20.1 });
  ws['!rows'].push({ hpt: 17.1 });
  ws['!rows'].push({ hpt: 21.95 });
  ws['!rows'].push({ hpt: 20.25 });

  // Date helpers
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
  var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr = 'D-' + todayYMD + '-' + (client.client_id || '').slice(-4);

  // Chemical column labels (fixed C-J = cols 2-9)
  var chemLabels = ['Clor\nRapid', 'Clor\nLent', 'pH\u2212', 'Anti-\nalgi\u0107', 'Flocu-\nlant', 'Deduri-\nzant', 'pH\nLichid', 'Cl\nLichid'];
  var chemKeys = ['treat_cl_granule_gr', 'treat_cl_tablete_export_gr', 'treat_ph_granule', 'treat_antialgic', 'treat_floculant', 'treat_bicarbonat', 'treat_ph_lichid_bidoane', 'treat_cl_lichid_bidoane'];
  var priceKeys = ['clor_rapid', 'clor_lent', 'ph_minus', 'antialgic', 'floculant', 'dedurizant', 'ph_lichid', 'cl_lichid'];

  // ─── ROW 0 (row 1): Navy bar ───
  var sNavyBar = { fill: F_NAVY, font: _fnt('Arial', 1, false, 'FFFFFF') };
  _mergeFill(ws, merges, 0, 0, 10, '', sNavyBar);

  // ─── ROW 1 (row 2): Company info — 3 sections ───
  var bdrCompany = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L);
  var sComp1 = { fill: F_HEADER, font: _fnt('Arial', 11, true), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_MED, S_THIN_L) };
  var sComp2 = { fill: F_HEADER, font: _fnt('Arial', 9, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_THIN_L, S_THIN_L) };
  var sComp3 = { fill: F_HEADER, font: _fnt('Arial', 8.5, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_MED, S_THIN_L, S_MED) };

  _mergeFill(ws, merges, 1, 0, 3, FIRMA_NUME + '\n' + FIRMA_ADRESA, sComp1);   // A2:D2
  _mergeFill(ws, merges, 1, 4, 6, FIRMA_EMAIL + '\n' + FIRMA_WEB + '\n' + FIRMA_TELEFON, sComp2); // E2:G2
  _mergeFill(ws, merges, 1, 7, 10, FIRMA_J + '\nCUI: ' + FIRMA_CUI + '\nIBAN: ' + FIRMA_IBAN, sComp3); // H2:K2

  // ─── ROW 2 (row 3): Accent bar ───
  var sAccentBar = { fill: F_ACCENT, font: _fnt('Arial', 1, false, '4DB8E8') };
  _mergeFill(ws, merges, 2, 0, 10, '', sAccentBar);

  // ─── ROW 3 (row 4): Title ───
  var sTitle = { fill: F_MID, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_M, S_MED, S_MED) };
  _mergeFill(ws, merges, 3, 0, 10, 'RAPORT INTERVEN\u021AII \u2014 CHIMICALE PISCIN\u0102', sTitle);

  // ─── ROW 4 (row 5): Labels ───
  var sLabel = { fill: F_LIGHT1, font: _fnt('Arial', 8, true, '404040'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_L, S_DOT, S_THIN_L, S_THIN_L) };
  _mergeFill(ws, merges, 4, 0, 2, 'Client', sLabel);         // A5:C5
  _mergeFill(ws, merges, 4, 3, 5, 'Perioada raportata', sLabel); // D5:F5
  _mergeFill(ws, merges, 4, 6, 8, 'Nr. Document', sLabel);    // G5:I5
  _mergeFill(ws, merges, 4, 9, 10, 'Data emiterii', sLabel);  // J5:K5

  // ─── ROW 5 (row 6): Values ───
  var sValue = { fill: F_WHITE, font: _fnt('Arial', 10, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_DOT, S_THIN_L, S_THIN_L, S_THIN_L) };
  _mergeFill(ws, merges, 5, 0, 2, client.name || '', sValue);
  _mergeFill(ws, merges, 5, 3, 5, period, sValue);
  _mergeFill(ws, merges, 5, 6, 8, docNr, sValue);
  _mergeFill(ws, merges, 5, 9, 10, todayStr, sValue);

  // ─── ROW 6 (row 7): Separator ───
  var sSep = { fill: F_LIGHT1, font: _fnt('Arial', 1, false, 'E8F3FB') };
  _mergeFill(ws, merges, 6, 0, 10, '', sSep);

  // ─── ROW 7 (row 8): Header row 1 ───
  var bdrHdr = _brd(S_MED, S_THIN_N, S_THIN_N, S_THIN_N);
  var sHdr1 = { fill: F_HDRDARK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrHdr };
  var sHdrChem = { fill: F_HEADER, font: _fnt('Arial', 10, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrHdr };
  var sHdrTotal = { fill: F_HDRDARK, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrHdr };

  // A8:A9 "Data interventie" merged
  _setCell(ws, 7, 0, '', sHdr1);
  ws[XLSX.utils.encode_cell({ r: 7, c: 0 })] = _cellS('Data\ninterven\u021Bie', sHdr1);
  _setCell(ws, 8, 0, '', sHdr1);
  merges.push({ s: { r: 7, c: 0 }, e: { r: 8, c: 0 } });

  // B8:B9 "Cant." merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 1 })] = _cellS('Cant.\n(l/kg)', sHdr1);
  _setCell(ws, 8, 1, '', sHdr1);
  merges.push({ s: { r: 7, c: 1 }, e: { r: 8, c: 1 } });

  // C8:J8 "CHIMICALE FOLOSITE" merged
  _mergeFill(ws, merges, 7, 2, 9, 'CHIMICALE FOLOSITE', sHdrChem);

  // K8:K9 "Total plata" merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 10 })] = _cellS('Total\nplat\u0103', sHdrTotal);
  _setCell(ws, 8, 10, '', sHdrTotal);
  merges.push({ s: { r: 7, c: 10 }, e: { r: 8, c: 10 } });

  // ─── ROW 8 (row 9): Sub-headers (chemical names) ───
  var bdrSub = _brd(S_THIN_N, S_MED, S_THIN_N, S_THIN_N);
  var sSubHdr = { fill: F_HEADER, font: _fnt('Arial', 8, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrSub };
  for (var ci = 0; ci < 8; ci++) {
    ws[XLSX.utils.encode_cell({ r: 8, c: ci + 2 })] = _cellS(chemLabels[ci], sSubHdr);
  }

  // ─── ROWS 9-18 (rows 10-19): Data rows ───
  var bdrDataL = _brd(S_THIN_L, S_THIN_L, S_MED, S_THIN_L);
  var bdrDataM = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L);
  var bdrDataR = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_MED);

  for (var di = 0; di < NR; di++) {
    var rowIdx = 9 + di;
    var fillD = (di % 2 === 0) ? F_DATA_A : F_WHITE;
    var sDataL = { fill: fillD, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataL };
    var sDataM = { fill: fillD, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataM };
    var sDataR = { fill: fillD, font: _fnt('Arial', 9, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataR };

    if (di < sorted.length) {
      var intv = sorted[di];
      // A: date
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS(fmtDateDMY(intv.date), sDataL);
      // B: cant (sum of all chems)
      var totalCant = 0;
      for (var ck = 0; ck < chemKeys.length; ck++) {
        totalCant += (parseFloat(intv[chemKeys[ck]]) || 0);
      }
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 1 })] = _cellS(totalCant > 0 ? totalCant : '', sDataM);
      // C-J: chemical values
      for (var cc = 0; cc < 8; cc++) {
        var val = parseFloat(intv[chemKeys[cc]]) || 0;
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: cc + 2 })] = _cellS(val > 0 ? val : '', sDataM);
      }
      // K: total plata for this row (sum of qty * price for each chem)
      var rowTotal = 0;
      for (var cp = 0; cp < 8; cp++) {
        var qty = parseFloat(intv[chemKeys[cp]]) || 0;
        var price = prices[priceKeys[cp]] || DEFAULT_PRICES[priceKeys[cp]] || 0;
        rowTotal += qty * price;
      }
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 10 })] = _cellS('', sDataR);
    } else {
      // Empty data row
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS('', sDataL);
      for (var ec = 1; ec < 10; ec++) {
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: ec })] = _cellS('', sDataM);
      }
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 10 })] = _cellS('', sDataR);
    }
  }

  // ─── ROW 19 (row 20): Cantitate totala ───
  var bdrTotL = _brd(S_MED, S_THIN_M, S_MED, S_THIN_M);
  var bdrTotM = _brd(S_MED, S_THIN_M, S_THIN_M, S_THIN_M);
  var bdrTotR = _brd(S_MED, S_THIN_M, S_THIN_M, S_MED);
  var sTotLabel = { fill: F_DATA_BK, font: _fnt('Arial', 9, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTotL };
  var sTotVal   = { fill: F_DATA_BK, font: _fnt('Arial', 9, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTotM };
  var sTotR     = { fill: F_DATA_BK, font: _fnt('Arial', 9, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTotR };

  _mergeFill(ws, merges, 19, 0, 1, 'Cantitate total\u0103', sTotLabel);
  // SUM formulas for C20:J20 (cols 2-9, rows 10-19 = 0-indexed 9-18)
  for (var sc = 0; sc < 8; sc++) {
    var colLetter = String.fromCharCode(67 + sc); // C=67
    var formula = 'SUM(' + colLetter + '10:' + colLetter + '19)';
    ws[XLSX.utils.encode_cell({ r: 19, c: sc + 2 })] = _cellF(formula, sTotVal);
  }
  ws[XLSX.utils.encode_cell({ r: 19, c: 10 })] = _cellS('', sTotR);

  // ─── ROW 20 (row 21): Pret unitar ───
  var sPretLabel = { fill: F_LIGHT1, font: _fnt('Arial', 9, true, '404040'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_MED, S_THIN_M) };
  var sPretVal   = { fill: F_LIGHT1, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_THIN_M, S_THIN_M) };
  var sPretR     = { fill: F_LIGHT1, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_M, S_THIN_M, S_THIN_M, S_MED) };

  _mergeFill(ws, merges, 20, 0, 1, 'Pre\u021B unitar', sPretLabel);
  for (var pc = 0; pc < 8; pc++) {
    var prc = prices[priceKeys[pc]] || DEFAULT_PRICES[priceKeys[pc]] || 0;
    ws[XLSX.utils.encode_cell({ r: 20, c: pc + 2 })] = _cellS(prc, sPretVal);
  }
  ws[XLSX.utils.encode_cell({ r: 20, c: 10 })] = _cellS('', sPretR);

  // ─── ROW 21 (row 22): TOTAL GENERAL ───
  var bdrGenL = _brd(S_MED, S_MED, S_MED, S_THIN_N);
  var bdrGenM = _brd(S_MED, S_MED, S_THIN_N, S_THIN_N);
  var bdrGenR = _brd(S_MED, S_MED, S_THIN_N, S_MED);
  var sGenLabel = { fill: F_HDRDARK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrGenL };
  var sGenVal   = { fill: F_MID, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrGenM };
  var sGenTot   = { fill: F_NAVY, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrGenR };

  _mergeFill(ws, merges, 21, 0, 1, 'TOTAL GENERAL', sGenLabel);
  // Formulas: =C20*C21 for each chem column
  for (var gc = 0; gc < 8; gc++) {
    var gColLetter = String.fromCharCode(67 + gc);
    var gFormula = gColLetter + '20*' + gColLetter + '21';
    ws[XLSX.utils.encode_cell({ r: 21, c: gc + 2 })] = _cellF(gFormula, sGenVal);
  }
  // K22 = SUM(C22:J22)
  ws[XLSX.utils.encode_cell({ r: 21, c: 10 })] = _cellF('SUM(C22:J22)', sGenTot);

  // ─── ROW 22 (row 23): Footer ───
  var bdrFoot = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L);
  var sFootL = { fill: F_LIGHT1, font: _fnt('Arial', 8.5, false, '555555'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrFoot };
  var sFootR = { fill: F_LIGHT1, font: _fnt('Arial', 8.5, false, '555555'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrFoot };

  _mergeFill(ws, merges, 22, 0, 4, 'Pre\u021Burile sunt exprimate \u00EEn lei/kg.\nCantit\u0103\u021Bile sunt exprimate \u00EEn grame.', sFootL);
  _mergeFill(ws, merges, 22, 5, 10, 'Generat automat \u2014 Pool Manager\n' + FIRMA_NUME, sFootR);

  // Set sheet ref and merges
  ws['!ref'] = 'A1:K23';
  ws['!merges'] = merges;

  return ws;
}

// ── V2: Build Servicii Sheet ───────────────────────────────────────
function _buildServiciiSheet(client, sorted, totalPlata, opsList) {
  // Build operations list: start with defaults, add any extras from interventions
  var defaultOps = [
    'Aspirare piscina', 'Curatare linie apa', 'Curatare skimmere',
    'Spalare filtru', 'Curatare prefiltru', 'Periere piscina',
    'Analiza apei', 'Tratament chimic'
  ];
  // Use provided opsList or defaults
  var allOps = (opsList && opsList.length) ? opsList.slice() : defaultOps.slice();
  // Scan interventions for any operations not yet in the list
  sorted.forEach(function(intv) {
    (intv.operations || []).forEach(function(op) {
      if (op && allOps.indexOf(op) < 0) allOps.push(op);
    });
  });

  var numOps = allOps.length;
  var NR = 17; // fixed 17 data rows
  var COLS = 1 + numOps; // A=Data + operation columns
  var lastCol = COLS - 1;
  var ws = {};
  var merges = [];

  // Display labels for known operations (wrap on /)
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

  // Column widths
  var colWidths = [{ wch: 13 }]; // A = date
  for (var cw = 0; cw < numOps; cw++) colWidths.push({ wch: 11 });
  ws['!cols'] = colWidths;

  // Row heights (0-indexed)
  ws['!rows'] = [
    { hpt: 3.95 },   // row 0 (1)
    { hpt: 52 },      // row 1 (2)
    { hpt: 3 },       // row 2 (3)
    { hpt: 20.1 },    // row 3 (4)
    { hpt: 15.95 },   // row 4 (5)
    { hpt: 17.1 },    // row 5 (6)
    { hpt: 3.95 },    // row 6 (7)
    { hpt: 21.95 },   // row 7 (8)
    { hpt: 45.95 }    // row 8 (9)
  ];
  // rows 9-25: data rows (height 20.1)
  for (var dr = 0; dr < NR; dr++) ws['!rows'].push({ hpt: 20.1 });
  // row 26 (27): 20.1, row 27 (28): 6, row 28 (29): 24, row 29 (30): 15
  ws['!rows'].push({ hpt: 20.1 });
  ws['!rows'].push({ hpt: 6 });
  ws['!rows'].push({ hpt: 24 });
  ws['!rows'].push({ hpt: 15 });

  // ─── ROW 0 (row 1): Navy bar ───
  var sNavyBar = { fill: F_NAVY, font: _fnt('Arial', 1, false, 'FFFFFF') };
  // Date helpers
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
  var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
  var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;
  var docNr = 'D-' + todayYMD + '-' + (client.client_id || '').slice(-4);

  // Helper to get column letter (0=A, 1=B, ..., 25=Z, 26=AA)
  function colLetter(ci) { return XLSX.utils.encode_col(ci); }
  var lastColLetter = colLetter(lastCol);

  // ─── ROW 0 (row 1): Navy bar ───
  var sNavyBar = { fill: F_NAVY, font: _fnt('Arial', 1, false, 'FFFFFF') };
  _mergeFill(ws, merges, 0, 0, lastCol, '', sNavyBar);

  // ─── ROW 1 (row 2): Company info — 3 sections ───
  var third = Math.max(Math.floor(COLS / 3), 1);
  var s1End = Math.min(third - 1, lastCol);
  var s2End = Math.min(third * 2 - 1, lastCol);
  var s3End = lastCol;

  var sComp1 = { fill: F_HEADER, font: _fnt('Arial', 9, true), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_MED, S_THIN_L) };
  var sComp2 = { fill: F_HEADER, font: _fnt('Arial', 9, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_THIN_L, S_THIN_L) };
  var sComp3 = { fill: F_HEADER, font: _fnt('Arial', 8.5, false), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: _brd(S_MED, S_THIN_L, S_THIN_L, S_MED) };

  _mergeFill(ws, merges, 1, 0, s1End, FIRMA_NUME + '\n' + FIRMA_ADRESA, sComp1);
  _mergeFill(ws, merges, 1, s1End + 1, s2End, '  ' + FIRMA_EMAIL + '\n  ' + FIRMA_WEB + '\n  ' + FIRMA_TELEFON, sComp2);
  _mergeFill(ws, merges, 1, s2End + 1, s3End, '  ' + FIRMA_J + '\n  CUI: ' + FIRMA_CUI + '\n  ' + FIRMA_IBAN, sComp3);

  // ─── ROW 2 (row 3): Accent bar ───
  var sAccentBar = { fill: F_ACCENT, font: _fnt('Arial', 1, false, '4DB8E8') };
  _mergeFill(ws, merges, 2, 0, lastCol, '', sAccentBar);

  // ─── ROW 3 (row 4): Title ───
  var sTitleV2 = { fill: F_MID, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_MED, S_THIN_M, S_MED, S_MED) };
  _mergeFill(ws, merges, 3, 0, lastCol, 'RAPORT SERVICII \u2014 ABONAMENT \u00CENTRE\u021AINERE PISCIN\u0102', sTitleV2);

  // ─── ROW 4 (row 5): Labels ───
  var sLabelV2 = { fill: F_LIGHT1, font: _fnt('Arial', 8, true, '404040'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_THIN_L, S_DOT, S_THIN_L, S_THIN_L) };
  var labQ = Math.max(Math.floor(COLS / 4), 1);
  var lab1End = Math.min(labQ - 1, lastCol);
  var lab2End = Math.min(labQ * 2 - 1, lastCol);
  var lab3End = Math.min(labQ * 3 - 1, lastCol);
  _mergeFill(ws, merges, 4, 0, lab1End, 'Client', sLabelV2);
  _mergeFill(ws, merges, 4, lab1End + 1, lab2End, 'Perioada raportare', sLabelV2);
  _mergeFill(ws, merges, 4, lab2End + 1, lab3End, 'Nr. Document', sLabelV2);
  _mergeFill(ws, merges, 4, lab3End + 1, lastCol, 'Data Emiterii', sLabelV2);

  // ─── ROW 5 (row 6): Values ───
  var sValueV2 = { fill: F_WHITE, font: _fnt('Arial', 10, true, '0D2D5A'), alignment: { horizontal: 'center', vertical: 'center' }, border: _brd(S_DOT, S_THIN_L, S_THIN_L, S_THIN_L) };
  _mergeFill(ws, merges, 5, 0, lab1End, client.name || '', sValueV2);
  _mergeFill(ws, merges, 5, lab1End + 1, lab2End, period, sValueV2);
  _mergeFill(ws, merges, 5, lab2End + 1, lab3End, docNr, sValueV2);
  _mergeFill(ws, merges, 5, lab3End + 1, lastCol, todayStr, sValueV2);

  // ─── ROW 6 (row 7): Separator ───
  var sSepV2 = { fill: F_LIGHT1, font: _fnt('Arial', 1, false, 'E8F3FB') };
  _mergeFill(ws, merges, 6, 0, lastCol, '', sSepV2);

  // ─── ROW 7 (row 8): Header row 1 ───
  var bdrHdr2 = _brd(S_MED, S_THIN_N, S_THIN_N, S_THIN_N);
  var sHdr1V2 = { fill: F_HDRDARK2, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrHdr2 };
  var sHdrSvc = { fill: F_SUBHDR, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrHdr2 };

  // A8:A9 "Data interventie" merged
  ws[XLSX.utils.encode_cell({ r: 7, c: 0 })] = _cellS('Data\ninterven\u021Bie', sHdr1V2);
  _setCell(ws, 8, 0, '', sHdr1V2);
  merges.push({ s: { r: 7, c: 0 }, e: { r: 8, c: 0 } });

  // B8:lastCol-1 "SERVICII INCLUSE IN ABONAMENT" merged
  var svcMergeEnd = Math.max(lastCol - 1, 1);
  _mergeFill(ws, merges, 7, 1, svcMergeEnd, 'SERVICII INCLUSE \u00CEN ABONAMENT', sHdrSvc);
  // Last col header cell
  if (lastCol > svcMergeEnd) {
    ws[XLSX.utils.encode_cell({ r: 7, c: lastCol })] = _cellS('', sHdr1V2);
  }

  // ─── ROW 8 (row 9): Sub-headers (service names) ───
  var bdrSub2 = _brd(S_THIN_N, S_MED, S_THIN_N, S_THIN_N);
  var sSubHdrV2 = { fill: F_SUBHDR, font: _fnt('Arial', 8, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdrSub2 };
  for (var si = 0; si < numOps; si++) {
    var opName = allOps[si];
    var label = knownLabels[opName] || opName.replace(/ /g, '\n');
    ws[XLSX.utils.encode_cell({ r: 8, c: si + 1 })] = _cellS(label, sSubHdrV2);
  }

  // ─── ROWS 9-25 (rows 10-26): Data rows ───
  var bdrDataL2 = _brd(S_THIN_L, S_THIN_L, S_MED, S_THIN_L);
  var bdrDataM2 = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L);
  var bdrDataR2 = _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_MED);

  for (var di = 0; di < NR; di++) {
    var rowIdx = 9 + di;
    var fillD = (di % 2 === 0) ? F_DATA_E : F_DATA_O;
    var sDateCell = { fill: fillD, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataL2 };
    var sSvcCell  = { fill: fillD, font: _fnt('Arial', 11, true, '1B6B3A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataM2 };
    var sSvcEmpty = { fill: fillD, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataM2 };
    var sSvcCellR = { fill: fillD, font: _fnt('Arial', 11, true, '1B6B3A'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataR2 };
    var sLastCell = { fill: fillD, font: _fnt('Arial', 9, false, '333333'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrDataR2 };

    if (di < sorted.length) {
      var intv = sorted[di];
      // A: date
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS(fmtDateDMY(intv.date), sDateCell);
      // B onwards: check services by exact match
      var ops = intv.operations || [];
      for (var sc = 0; sc < numOps; sc++) {
        var matched = ops.indexOf(allOps[sc]) >= 0;
        var isLast = (sc === numOps - 1);
        if (matched) {
          ws[XLSX.utils.encode_cell({ r: rowIdx, c: sc + 1 })] = _cellS('\u2713', isLast ? sSvcCellR : sSvcCell);
        } else {
          ws[XLSX.utils.encode_cell({ r: rowIdx, c: sc + 1 })] = _cellS('', isLast ? sLastCell : sSvcEmpty);
        }
      }
    } else {
      // Empty data row
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = _cellS('', sDateCell);
      for (var ec = 1; ec < numOps; ec++) {
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: ec })] = _cellS('', sSvcEmpty);
      }
      ws[XLSX.utils.encode_cell({ r: rowIdx, c: lastCol })] = _cellS('', sLastCell);
    }
  }

  // ─── ROW 26 (row 27): Total interventii efectuate ───
  var bdrTot2L = _brd(S_MED, S_MED, S_MED, S_THIN_N);
  var bdrTot2M = _brd(S_MED, S_MED, S_THIN_N, S_THIN_N);
  var bdrTot2R = _brd(S_MED, S_MED, S_THIN_N, S_MED);
  var sTotLabelV2 = { fill: F_TOT_HDR, font: _fnt('Arial', 9, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTot2L };
  var sTotValV2   = { fill: F_TOT_DK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTot2M };
  var sTotRV2     = { fill: F_TOT_DK, font: _fnt('Arial', 10, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrTot2R };

  ws[XLSX.utils.encode_cell({ r: 26, c: 0 })] = _cellS('Total interven\u021Bii efectuate', sTotLabelV2);
  // COUNTA formulas for each ops column (rows 10-26)
  for (var tc = 0; tc < numOps; tc++) {
    var tCol = colLetter(tc + 1);
    var tFormula = 'COUNTA(' + tCol + '10:' + tCol + '26)';
    var isLast = (tc === numOps - 1);
    ws[XLSX.utils.encode_cell({ r: 26, c: tc + 1 })] = _cellF(tFormula, isLast ? sTotRV2 : sTotValV2);
  }

  // ─── ROW 27 (row 28): Separator ───
  var sSep2 = { fill: F_LIGHT1, font: _fnt('Arial', 1, false, 'E8F3FB') };
  _mergeFill(ws, merges, 27, 0, lastCol, '', sSep2);

  // ─── ROW 28 (row 29): TOTAL DE PLATA ───
  var payLabelEnd = Math.max(lastCol - 2, 0);
  var payValStart = payLabelEnd + 1;
  var bdrPayL = _brd(S_MED, S_MED, S_MED, S_THIN_N);
  var bdrPayR = _brd(S_MED, S_MED, S_THIN_N, S_MED);
  var sPayLabel = { fill: F_HDRDARK2, font: _fnt('Arial', 11, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrPayL };
  var sPayVal   = { fill: F_NAVY, font: _fnt('Arial', 12, true, 'FFFFFF'), alignment: { horizontal: 'center', vertical: 'center' }, border: bdrPayR };

  _mergeFill(ws, merges, 28, 0, payLabelEnd, 'TOTAL DE PLAT\u0102', sPayLabel);
  _mergeFill(ws, merges, 28, payValStart, lastCol, totalPlata || '', sPayVal);

  // ─── ROW 29 (row 30): Footer ───
  var sFootV2 = { fill: F_LIGHT1, font: _fnt('Arial', 8.5, false, '555555'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _brd(S_THIN_L, S_THIN_L, S_THIN_L, S_THIN_L) };
  var footHalf = Math.floor(COLS / 2);
  _mergeFill(ws, merges, 29, 0, footHalf - 1, 'Generat automat \u2014 Pool Manager', sFootV2);
  _mergeFill(ws, merges, 29, footHalf, lastCol, FIRMA_NUME, sFootV2);

  // Set sheet ref and merges
  ws['!ref'] = 'A1:' + lastColLetter + '30';
  ws['!merges'] = merges;

  return ws;
}

// ── Export Deviz Chimicale (V1 only) ───────────────────────────────
function exportDevizChimicale(client, interventions) {
  return loadXLSX().then(async function() {
    var sorted = interventions.slice().sort(function(a, b) { return a.date.localeCompare(b.date); });
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : DEFAULT_PRICES;

    var wb = XLSX.utils.book_new();
    var ws = _buildChimicaleSheet(client, sorted, prices);
    var sheetName = sanitizeSheetName(client.name || 'Chimicale');
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    var fname = sanitizeFilename(client.name) + '_Chimicale_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname, null, client.name);
    return fname;
  });
}

// ── Export Deviz Complet (V1 + V2 in same workbook) ────────────────
function exportDevizComplet(client, interventions) {
  return loadXLSX().then(async function() {
    var sorted = interventions.slice().sort(function(a, b) { return a.date.localeCompare(b.date); });
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : DEFAULT_PRICES;
    var opsList = (typeof getOperations === 'function') ? await getOperations() : null;

    var wb = XLSX.utils.book_new();

    // Sheet 1: Chimicale (V1)
    var ws1 = _buildChimicaleSheet(client, sorted, prices);
    var name1 = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Chim');
    XLSX.utils.book_append_sheet(wb, ws1, name1);

    // Sheet 2: Servicii (V2)
    var ws2 = _buildServiciiSheet(client, sorted, '', opsList);
    var name2 = sanitizeSheetName((client.name || 'Client').substring(0, 25) + '_Serv');
    XLSX.utils.book_append_sheet(wb, ws2, name2);

    var fname = sanitizeFilename(client.name) + '_Deviz_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname, null, client.name);
    return fname;
  });
}

// ── Export All Deviz Mixed (all clients) ───────────────────────────
function exportAllDevizMixed(clients, allInterventions, filter) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : DEFAULT_PRICES;
    var opsList = (typeof getOperations === 'function') ? await getOperations() : null;
    var wb = XLSX.utils.book_new();
    var sheetCount = 0;

    clients.forEach(function(client) {
      var cid = client.client_id;
      var clientIntv = (allInterventions[cid] || []).slice().sort(function(a, b) { return a.date.localeCompare(b.date); });

      // Apply date filter if provided
      if (filter && filter.startDate) {
        clientIntv = clientIntv.filter(function(i) { return i.date >= filter.startDate; });
      }
      if (filter && filter.endDate) {
        clientIntv = clientIntv.filter(function(i) { return i.date <= filter.endDate; });
      }

      if (clientIntv.length === 0) return;

      var baseName = sanitizeSheetName(client.name || 'Client');

      // Determine deviz type from client or filter
      var devizType = (filter && filter.devizType) || client.deviz_type || 'V1';

      if (devizType === 'V2' || devizType === 'complet') {
        // Add Servicii sheet
        var ws2 = _buildServiciiSheet(client, clientIntv, '', opsList);
        var opsName = baseName.substring(0, 28) + '_Sv';
        if (wb.SheetNames.indexOf(opsName) >= 0) opsName = opsName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws2, opsName);
        sheetCount++;
      }

      if (devizType === 'V1' || devizType === 'complet') {
        // Add Chimicale sheet
        var ws1 = _buildChimicaleSheet(client, clientIntv, prices);
        var chemName = baseName.substring(0, 28) + '_Ch';
        if (wb.SheetNames.indexOf(chemName) >= 0) chemName = chemName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws1, chemName);
        sheetCount++;
      }
    });

    if (sheetCount === 0) {
      showToast('Nicio interventie de exportat.', 'warning');
      return;
    }

    var fname = 'DevizToti_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
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
  return loadXLSX().then(function() {
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
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname, null, client ? client.name : null);
    showToast('Deviz Excel descarcat: ' + fname, 'success');
  }).catch(function(e) {
    showToast('Eroare export: ' + e.message, 'error');
  });
}

