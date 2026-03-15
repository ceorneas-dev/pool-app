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
  { key: 'treat_cl_lichid_bidoane',   label: 'Cl Lichid',   priceKey: 'cl_lichid' },
  { key: 'treat_sare_saci',           label: 'Sare',        priceKey: 'sare' }
];

// ── Styled Deviz Helpers (xlsx-js-style) ─────────────────────────

var _BORDER_THIN = {
  top:    { style: 'thin', color: { rgb: 'B0B0B0' } },
  bottom: { style: 'thin', color: { rgb: 'B0B0B0' } },
  left:   { style: 'thin', color: { rgb: 'B0B0B0' } },
  right:  { style: 'thin', color: { rgb: 'B0B0B0' } }
};

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

// == Build styled V1 chimicale sheet ==
function _buildChimicaleSheet(client, sorted, prices) {
  var usedCols = ALL_CHEM_COLS.filter(function(c) {
    return sorted.some(function(i) { return (parseFloat(i[c.key]) || 0) > 0; });
  });

  var numChem = usedCols.length;
  var totalCols = 2 + numChem + 2; // A=Data, B=Cant, chemicals, empty sep, Total plata
  var lastCol = totalCols - 1;
  var lastChemCol = 2 + numChem - 1; // last chemical column index
  var sepCol = 2 + numChem; // empty separator column index

  var ws = {};
  var merges = [];
  var r = 0; // current row

  // Today's date
  var today = new Date();
  var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
  var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');

  // Period from interventions
  var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
  var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
  var period = firstDate + ' - ' + lastDate;

  // Doc number
  var docNr = 'D-' + todayYMD + '-' + (client.client_id || '').slice(-4);

  // Styles
  var sNavy = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { color: { rgb: 'FFFFFF' }, sz: 1 }, border: _BORDER_THIN };
  var sAccent = { fill: { fgColor: { rgb: '4DB8E8' } }, font: { sz: 1 }, border: _BORDER_THIN };
  var sTitle = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sLabelRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { bold: true, sz: 8 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sValueRow = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sSepRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 1 }, border: _BORDER_THIN };
  var sHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
  var sSubHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 8.5, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
  var sCompany1 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sCompany2 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
  var sCompany3 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 8.5 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };

  // === ROW 0: Dark navy banner ===
  _fillEmptyCells(ws, r, totalCols, sNavy);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } });
  r++;

  // === ROW 1: Company info (3 sections) ===
  var thirdW = Math.floor(totalCols / 3);
  var sec1End = thirdW - 1;
  var sec2End = thirdW * 2 - 1;
  var sec3End = lastCol;

  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('S.C. AQUATIS ENGINEERING S.R.L.', sCompany1);
  for (var ci = 1; ci <= sec1End; ci++) ws[XLSX.utils.encode_cell({ r: r, c: ci })] = _cellS('', sCompany1);
  merges.push({ s: { r: r, c: 0 }, e: { r: r, c: sec1End } });

  ws[XLSX.utils.encode_cell({ r: r, c: sec1End + 1 })] = _cellS('office@aquatis.ro\nwww.aquatis.ro', sCompany2);
  for (var ci2 = sec1End + 2; ci2 <= sec2End; ci2++) ws[XLSX.utils.encode_cell({ r: r, c: ci2 })] = _cellS('', sCompany2);
  merges.push({ s: { r: r, c: sec1End + 1 }, e: { r: r, c: sec2End } });

  ws[XLSX.utils.encode_cell({ r: r, c: sec2End + 1 })] = _cellS('J40/18144/2007\nCUI: RO22479695', sCompany3);
  for (var ci3 = sec2End + 2; ci3 <= sec3End; ci3++) ws[XLSX.utils.encode_cell({ r: r, c: ci3 })] = _cellS('', sCompany3);
  merges.push({ s: { r: r, c: sec2End + 1 }, e: { r: r, c: sec3End } });
  r++;

  // === ROW 2: Accent line ===
  _fillEmptyCells(ws, r, totalCols, sAccent);
  merges.push({ s: { r: r, c: 0 }, e: { r: r, c: lastCol } });
  r++;

  // === ROW 3: Title ===
  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('RAPORT INTERVEN\u021AII \u2014 CHIMICALE FOLOSITE', sTitle);
  _fillEmptyCells(ws, r, totalCols, sTitle);
  merges.push({ s: { r: r, c: 0 }, e: { r: r, c: lastCol } });
  r++;

  // === ROW 4: Labels row ===
  var labQ = Math.floor(totalCols / 4);
  var labGroups = [
    { label: 'Client', start: 0, end: labQ - 1 },
    { label: 'Luna / Perioada', start: labQ, end: labQ * 2 - 1 },
    { label: 'Nr. Document', start: labQ * 2, end: labQ * 3 - 1 },
    { label: 'Data emiterii', start: labQ * 3, end: lastCol }
  ];
  labGroups.forEach(function(g) {
    ws[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(g.label, sLabelRow);
    for (var lc = g.start + 1; lc <= g.end; lc++) ws[XLSX.utils.encode_cell({ r: r, c: lc })] = _cellS('', sLabelRow);
    merges.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
  });
  r++;

  // === ROW 5: Values row ===
  var valValues = [client.name || '', period, docNr, todayStr];
  labGroups.forEach(function(g, gi) {
    ws[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(valValues[gi], sValueRow);
    for (var vc = g.start + 1; vc <= g.end; vc++) ws[XLSX.utils.encode_cell({ r: r, c: vc })] = _cellS('', sValueRow);
    merges.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
  });
  r++;

  // === ROW 6: Separator ===
  _fillEmptyCells(ws, r, totalCols, sSepRow);
  merges.push({ s: { r: r, c: 0 }, e: { r: r, c: lastCol } });
  r++;

  // === ROW 7: Header row 1 ===
  var headerRow1 = r;
  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Data\ninterven\u021Bie', sHeader);
  ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS('Cant.\n(l/kg)', sHeader);

  if (numChem > 0) {
    ws[XLSX.utils.encode_cell({ r: r, c: 2 })] = _cellS('CHIMICALE FOLOSITE', { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN });
    for (var hc = 3; hc <= lastChemCol; hc++) {
      ws[XLSX.utils.encode_cell({ r: r, c: hc })] = _cellS('', sHeader);
    }
    if (numChem > 1) {
      merges.push({ s: { r: r, c: 2 }, e: { r: r, c: lastChemCol } });
    }
  }

  // Separator col empty
  ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', sHeader);

  // Total plata header (merged with row 8)
  ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellS('Total plat\u0103\n(RON)', sHeader);
  r++;

  // === ROW 8: Sub-headers ===
  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('', sSubHeader);
  ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS('', sSubHeader);

  // Merge A7:A8, B7:B8
  merges.push({ s: { r: headerRow1, c: 0 }, e: { r: r, c: 0 } });
  merges.push({ s: { r: headerRow1, c: 1 }, e: { r: r, c: 1 } });

  // Chemical sub-headers
  usedCols.forEach(function(c, ci2) {
    ws[XLSX.utils.encode_cell({ r: r, c: 2 + ci2 })] = _cellS(c.label, sSubHeader);
  });

  ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', sSubHeader);

  // Merge lastCol row7:row8
  ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellS('', sHeader);
  merges.push({ s: { r: headerRow1, c: lastCol }, e: { r: r, c: lastCol } });
  r++;

  // === DATA ROWS ===
  var dataStartRow = r;
  sorted.forEach(function(inv, idx) {
    var isEven = idx % 2 === 0;
    var bgColor = isEven ? 'E0EEF8' : 'F0F6FB';
    var sDataCell = { fill: { fgColor: { rgb: bgColor } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sDateCell = { fill: { fgColor: { rgb: bgColor } }, font: { sz: 9 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };

    ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS(fmtDateDMY(inv.date), sDateCell);
    ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS(1, sDataCell);

    usedCols.forEach(function(c, ci3) {
      var val = parseFloat(inv[c.key]) || 0;
      ws[XLSX.utils.encode_cell({ r: r, c: 2 + ci3 })] = _cellS(val > 0 ? val : '', sDataCell);
    });

    ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', { fill: { fgColor: { rgb: bgColor } }, border: _BORDER_THIN });
    ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellS('', sDataCell);
    r++;
  });
  var dataEndRow = r - 1;

  // === Empty separator row ===
  _fillEmptyCells(ws, r, totalCols, { border: _BORDER_THIN });
  r++;

  // === Cantitate totala row ===
  var totRowIdx = r;
  var sTotRow = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sTotLabel = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };

  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Cantitate total\u0103', sTotLabel);
  ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS(sorted.length, sTotRow);

  usedCols.forEach(function(c, ci4) {
    var colLetter = XLSX.utils.encode_cell({ r: 0, c: 2 + ci4 }).replace(/[0-9]/g, '');
    var formula = 'SUM(' + colLetter + (dataStartRow + 1) + ':' + colLetter + (dataEndRow + 1) + ')';
    ws[XLSX.utils.encode_cell({ r: r, c: 2 + ci4 })] = _cellF(formula, sTotRow);
  });

  ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', sTotRow);
  ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellS('', sTotRow);
  r++;

  // === Pret unitar row ===
  var priceRowIdx = r;
  var sPriceRow = { fill: { fgColor: { rgb: 'EDF4FB' } }, font: { sz: 8.5 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sPriceLabel = { fill: { fgColor: { rgb: 'EDF4FB' } }, font: { sz: 8.5 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };

  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Pre\u021B unitar (RON)', sPriceLabel);
  ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS(prices.pret_interventie || 0, sPriceRow);

  usedCols.forEach(function(c, ci5) {
    ws[XLSX.utils.encode_cell({ r: r, c: 2 + ci5 })] = _cellS(prices[c.priceKey] || 0, sPriceRow);
  });

  ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', sPriceRow);
  ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellS('', sPriceRow);
  r++;

  // === TOTAL GENERAL row ===
  var totalGenRowIdx = r;
  var sTotalGen = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
  var sTotalGenLabel = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
  var sTotalGenFinal = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };

  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('TOTAL GENERAL (RON)', sTotalGenLabel);

  // B col: formula = B_tot * B_price
  var bTotRef = XLSX.utils.encode_cell({ r: totRowIdx, c: 1 });
  var bPriceRef = XLSX.utils.encode_cell({ r: priceRowIdx, c: 1 });
  ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellF(bTotRef + '*' + bPriceRef, sTotalGen);

  var genFormulaRefs = [XLSX.utils.encode_cell({ r: r, c: 1 })];

  usedCols.forEach(function(c, ci6) {
    var colIdx = 2 + ci6;
    var totRef = XLSX.utils.encode_cell({ r: totRowIdx, c: colIdx });
    var priRef = XLSX.utils.encode_cell({ r: priceRowIdx, c: colIdx });
    ws[XLSX.utils.encode_cell({ r: r, c: colIdx })] = _cellF(totRef + '*' + priRef, sTotalGen);
    genFormulaRefs.push(XLSX.utils.encode_cell({ r: r, c: colIdx }));
  });

  ws[XLSX.utils.encode_cell({ r: r, c: sepCol })] = _cellS('', sTotalGen);

  // Last col: SUM of all total general values
  var sumFormula = genFormulaRefs.join('+');
  ws[XLSX.utils.encode_cell({ r: r, c: lastCol })] = _cellF(sumFormula, sTotalGenFinal);
  r++;

  // === Footer row ===
  var sFooterL = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
  var sFooterR = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'right', vertical: 'center' }, border: _BORDER_THIN };

  var halfCols = Math.floor(totalCols / 2);
  ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Toate pre\u021Burile sunt exprimate \u00een RON, f\u0103r\u0103 TVA', sFooterL);
  for (var fc = 1; fc < halfCols; fc++) ws[XLSX.utils.encode_cell({ r: r, c: fc })] = _cellS('', sFooterL);
  merges.push({ s: { r: r, c: 0 }, e: { r: r, c: halfCols - 1 } });

  ws[XLSX.utils.encode_cell({ r: r, c: halfCols })] = _cellS('S.C. Aquatis Engineering S.R.L.', sFooterR);
  for (var fc2 = halfCols + 1; fc2 <= lastCol; fc2++) ws[XLSX.utils.encode_cell({ r: r, c: fc2 })] = _cellS('', sFooterR);
  merges.push({ s: { r: r, c: halfCols }, e: { r: r, c: lastCol } });
  r++;

  // Set sheet range
  ws['!ref'] = XLSX.utils.encode_cell({ r: 0, c: 0 }) + ':' + XLSX.utils.encode_cell({ r: r - 1, c: lastCol });
  ws['!merges'] = merges;

  // Column widths
  var cols = [{ wch: 16 }, { wch: 8 }];
  for (var wc = 0; wc < numChem; wc++) cols.push({ wch: 10 });
  cols.push({ wch: 3 }); // separator
  cols.push({ wch: 13 }); // total plata
  ws['!cols'] = cols;

  return { ws: ws, totalPlata: 0, usedCols: usedCols, dataStartRow: dataStartRow, dataEndRow: dataEndRow, totRowIdx: totRowIdx, priceRowIdx: priceRowIdx, totalGenRowIdx: totalGenRowIdx };
}

// == Export Format 1: Deviz Chimicale (V1) — Styled ==
function exportDevizChimicale(client, interventions) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var wb = XLSX.utils.book_new();
    var sorted = interventions.slice().sort(function(a, b) { return a.date.localeCompare(b.date); });

    var result = _buildChimicaleSheet(client, sorted, prices);
    XLSX.utils.book_append_sheet(wb, result.ws, 'Raport Interventii');

    var fname = 'Deviz_' + sanitizeFilename(client.name) + '_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname, null, client ? client.name : null);
    return fname;
  });
}

// == Export Format 2: Deviz Complet (V2 — Chimicale + Servicii Abonament) — Styled ==
function exportDevizComplet(client, interventions) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var wb = XLSX.utils.book_new();
    var sorted = interventions.slice().sort(function(a, b) { return a.date.localeCompare(b.date); });

    // Sheet 1: Chimicale (V1 styled)
    var chemResult = _buildChimicaleSheet(client, sorted, prices);
    XLSX.utils.book_append_sheet(wb, chemResult.ws, 'Raport Interventii');

    // Sheet 2: Servicii Abonament (V2 styled)
    var opsList = (typeof getOperations === 'function') ? await getOperations() : [
      'Aspirare piscina', 'Curatare linie apa', 'Curatare skimmere',
      'Spalare filtru', 'Curatare prefiltru', 'Periere piscina',
      'Analiza apei', 'Tratament chimic', 'Verificare automatizare'
    ];

    var numOps = opsList.length;
    var v2TotalCols = 1 + numOps; // A=Data + operation columns
    var v2LastCol = v2TotalCols - 1;

    var ws2 = {};
    var merges2 = [];
    var r = 0;

    // Today
    var today = new Date();
    var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
    var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
    var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
    var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
    var period = firstDate + ' - ' + lastDate;
    var docNr = 'D-' + todayYMD + '-' + (client.client_id || '').slice(-4);

    // Styles
    var sNavy = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { color: { rgb: 'FFFFFF' }, sz: 1 }, border: _BORDER_THIN };
    var sAccent = { fill: { fgColor: { rgb: '4DB8E8' } }, font: { sz: 1 }, border: _BORDER_THIN };
    var sTitle = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sLabelRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { bold: true, sz: 8 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sValueRow = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sSepRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 1 }, border: _BORDER_THIN };
    var sHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
    var sSubHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 8, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
    var sCompany1 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sCompany2 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
    var sCompany3 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 8.5 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };

    // === ROW 0: Dark navy banner ===
    _fillEmptyCells(ws2, r, v2TotalCols, sNavy);
    merges2.push({ s: { r: 0, c: 0 }, e: { r: 0, c: v2LastCol } });
    r++;

    // === ROW 1: Company info ===
    var thirdW = Math.floor(v2TotalCols / 3);
    var sec1End = Math.max(thirdW - 1, 0);
    var sec2End = Math.max(thirdW * 2 - 1, sec1End + 1);
    var sec3End = v2LastCol;

    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('S.C. AQUATIS ENGINEERING S.R.L.', sCompany1);
    for (var ci = 1; ci <= sec1End; ci++) ws2[XLSX.utils.encode_cell({ r: r, c: ci })] = _cellS('', sCompany1);
    if (sec1End > 0) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: sec1End } });

    ws2[XLSX.utils.encode_cell({ r: r, c: sec1End + 1 })] = _cellS('office@aquatis.ro\nwww.aquatis.ro', sCompany2);
    for (var ci2 = sec1End + 2; ci2 <= sec2End; ci2++) ws2[XLSX.utils.encode_cell({ r: r, c: ci2 })] = _cellS('', sCompany2);
    if (sec2End > sec1End + 1) merges2.push({ s: { r: r, c: sec1End + 1 }, e: { r: r, c: sec2End } });

    ws2[XLSX.utils.encode_cell({ r: r, c: sec2End + 1 })] = _cellS('J40/18144/2007\nCUI: RO22479695', sCompany3);
    for (var ci3 = sec2End + 2; ci3 <= sec3End; ci3++) ws2[XLSX.utils.encode_cell({ r: r, c: ci3 })] = _cellS('', sCompany3);
    if (sec3End > sec2End + 1) merges2.push({ s: { r: r, c: sec2End + 1 }, e: { r: r, c: sec3End } });
    r++;

    // === ROW 2: Accent line ===
    _fillEmptyCells(ws2, r, v2TotalCols, sAccent);
    merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
    r++;

    // === ROW 3: Title ===
    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('RAPORT SERVICII \u2014 ABONAMENT \u00cENTRE\u021aINERE', sTitle);
    _fillEmptyCells(ws2, r, v2TotalCols, sTitle);
    merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
    r++;

    // === ROW 4: Labels ===
    var labQ2 = Math.floor(v2TotalCols / 4);
    if (labQ2 < 1) labQ2 = 1;
    var labGroups2 = [
      { label: 'Client', start: 0, end: Math.min(labQ2 - 1, v2LastCol) },
      { label: 'Luna / Perioada', start: Math.min(labQ2, v2LastCol), end: Math.min(labQ2 * 2 - 1, v2LastCol) },
      { label: 'Nr. Document', start: Math.min(labQ2 * 2, v2LastCol), end: Math.min(labQ2 * 3 - 1, v2LastCol) },
      { label: 'Data emiterii', start: Math.min(labQ2 * 3, v2LastCol), end: v2LastCol }
    ];
    labGroups2.forEach(function(g) {
      if (g.start > v2LastCol) return;
      ws2[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(g.label, sLabelRow);
      for (var lc = g.start + 1; lc <= g.end; lc++) ws2[XLSX.utils.encode_cell({ r: r, c: lc })] = _cellS('', sLabelRow);
      if (g.end > g.start) merges2.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
    });
    r++;

    // === ROW 5: Values ===
    var valValues2 = [client.name || '', period, docNr, todayStr];
    labGroups2.forEach(function(g, gi) {
      if (g.start > v2LastCol) return;
      ws2[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(valValues2[gi], sValueRow);
      for (var vc = g.start + 1; vc <= g.end; vc++) ws2[XLSX.utils.encode_cell({ r: r, c: vc })] = _cellS('', sValueRow);
      if (g.end > g.start) merges2.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
    });
    r++;

    // === ROW 6: Separator ===
    _fillEmptyCells(ws2, r, v2TotalCols, sSepRow);
    merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
    r++;

    // === ROW 7: Header row 1 ===
    var v2HeaderRow1 = r;
    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Data\ninterven\u021Bie', sHeader);

    if (numOps > 0) {
      ws2[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS('SERVICII INCLUSE \u00cEN ABONAMENT', { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN });
      for (var hc = 2; hc <= v2LastCol; hc++) ws2[XLSX.utils.encode_cell({ r: r, c: hc })] = _cellS('', sHeader);
      if (numOps > 1) merges2.push({ s: { r: r, c: 1 }, e: { r: r, c: v2LastCol } });
    }
    r++;

    // === ROW 8: Sub-headers (operation names) ===
    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('', sSubHeader);
    // Merge A7:A8
    merges2.push({ s: { r: v2HeaderRow1, c: 0 }, e: { r: r, c: 0 } });

    // Operation name wrapping: split on space to add \n
    var opNameWrap = {
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
    opsList.forEach(function(op, oi) {
      var label = opNameWrap[op] || op.replace(/ /g, '\n');
      ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellS(label, sSubHeader);
    });
    r++;

    // === DATA ROWS ===
    var v2DataStart = r;
    sorted.forEach(function(inv, idx) {
      var isEven = idx % 2 === 0;
      var bgColor = isEven ? 'E0EEF8' : 'F0F6FB';
      var sDataCell = { fill: { fgColor: { rgb: bgColor } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
      var sCheckCell = { fill: { fgColor: { rgb: bgColor } }, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };

      ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS(fmtDateDMY(inv.date), sDataCell);

      var ops = inv.operations || [];
      opsList.forEach(function(op, oi) {
        var hasOp = ops.indexOf(op) >= 0;
        ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellS(hasOp ? '\u2713' : '', hasOp ? sCheckCell : sDataCell);
      });
      r++;
    });
    var v2DataEnd = r - 1;

    // === Empty filler rows (visual space, up to ~3 extra) ===
    var fillerRows = Math.max(0, 3 - sorted.length);
    for (var fi = 0; fi < fillerRows; fi++) {
      _fillEmptyCells(ws2, r, v2TotalCols, { border: _BORDER_THIN });
      r++;
    }

    // === Total interventii efectuate row ===
    var sTotalDark = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
    var sTotalDarkL = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };

    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Total interven\u021Bii efectuate', sTotalDarkL);
    opsList.forEach(function(op, oi) {
      var colLetter = XLSX.utils.encode_cell({ r: 0, c: 1 + oi }).replace(/[0-9]/g, '');
      var formula = 'COUNTIF(' + colLetter + (v2DataStart + 1) + ':' + colLetter + (v2DataEnd + 1) + ',"\u2713")';
      ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellF(formula, sTotalDark);
    });
    r++;

    // === Separator row ===
    _fillEmptyCells(ws2, r, v2TotalCols, sSepRow);
    merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
    r++;

    // === TOTAL DE PLATA row ===
    // Calculate total from chemResult — get the grand total from the V1 sheet
    // We need the actual total: recalculate
    var totalPlata = (prices.pret_interventie || 0) * sorted.length;
    var usedChemCols = chemResult.usedCols || [];
    usedChemCols.forEach(function(c) {
      var s = 0;
      sorted.forEach(function(inv) { s += parseFloat(inv[c.key]) || 0; });
      totalPlata += s * (prices[c.priceKey] || 0);
    });

    var sTotalPay = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
    var sTotalPayVal = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };

    var payLabelEnd = Math.max(v2LastCol - 2, 0);
    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('TOTAL DE PLAT\u0102 (RON)', sTotalPay);
    for (var pc = 1; pc <= payLabelEnd; pc++) ws2[XLSX.utils.encode_cell({ r: r, c: pc })] = _cellS('', sTotalPay);
    if (payLabelEnd > 0) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: payLabelEnd } });

    var payValStart = payLabelEnd + 1;
    ws2[XLSX.utils.encode_cell({ r: r, c: payValStart })] = _cellS(totalPlata + ' lei', sTotalPayVal);
    for (var pv = payValStart + 1; pv <= v2LastCol; pv++) ws2[XLSX.utils.encode_cell({ r: r, c: pv })] = _cellS('', sTotalPayVal);
    if (v2LastCol > payValStart) merges2.push({ s: { r: r, c: payValStart }, e: { r: r, c: v2LastCol } });
    r++;

    // === Footer row ===
    var sFooterL = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
    var sFooterR = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'right', vertical: 'center' }, border: _BORDER_THIN };

    var halfCols2 = Math.floor(v2TotalCols / 2);
    ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Document generat de S.C. Aquatis Engineering S.R.L.', sFooterL);
    for (var fc = 1; fc < halfCols2; fc++) ws2[XLSX.utils.encode_cell({ r: r, c: fc })] = _cellS('', sFooterL);
    if (halfCols2 > 1) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: halfCols2 - 1 } });

    ws2[XLSX.utils.encode_cell({ r: r, c: halfCols2 })] = _cellS('www.aquatis.ro  |  0721.137.100', sFooterR);
    for (var fc2 = halfCols2 + 1; fc2 <= v2LastCol; fc2++) ws2[XLSX.utils.encode_cell({ r: r, c: fc2 })] = _cellS('', sFooterR);
    if (v2LastCol > halfCols2) merges2.push({ s: { r: r, c: halfCols2 }, e: { r: r, c: v2LastCol } });
    r++;

    // Set sheet range
    ws2['!ref'] = XLSX.utils.encode_cell({ r: 0, c: 0 }) + ':' + XLSX.utils.encode_cell({ r: r - 1, c: v2LastCol });
    ws2['!merges'] = merges2;

    // Column widths
    var v2Cols = [{ wch: 16 }, { wch: 13 }];
    for (var wc = 1; wc < numOps; wc++) v2Cols.push({ wch: 10 });
    ws2['!cols'] = v2Cols;

    XLSX.utils.book_append_sheet(wb, ws2, 'Servicii Abonament');

    var fname = 'DevizComplet_' + sanitizeFilename(client.name) + '_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname, null, client ? client.name : null);
    return fname;
  });
}

// == Export ALL clients — Mixed V1/V2 per client.deviz_type — Styled ==
function exportAllDevizMixed(clients, allInterventions, filter) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var opsList = (typeof getOperations === 'function') ? await getOperations() : [
      'Aspirare piscina', 'Curatare linie apa', 'Curatare skimmere',
      'Spalare filtru', 'Curatare prefiltru', 'Periere piscina',
      'Analiza apei', 'Tratament chimic', 'Verificare automatizare'
    ];
    var wb = XLSX.utils.book_new();
    var sheetCount = 0;

    for (var cx = 0; cx < clients.length; cx++) {
      var client = clients[cx];
      var ci = allInterventions.filter(function(i) { return i.client_id === client.client_id; });
      if (!ci.length) continue;

      var sorted = ci.slice().sort(function(a, b) { return b.date.localeCompare(a.date); });
      // Apply filter
      if (filter && filter.mode === 'last') {
        sorted = sorted.slice(0, filter.lastN || 4);
      } else if (filter && filter.mode === 'date' && filter.fromDate) {
        sorted = sorted.filter(function(i) { return i.date >= filter.fromDate; });
      }
      // Re-sort ascending for display
      sorted.sort(function(a, b) { return a.date.localeCompare(b.date); });
      if (!sorted.length) continue;

      var devizType = parseInt(client.deviz_type) || 1;
      var baseName = sanitizeSheetName(client.name);

      // V1 sheet (chimicale) — always present
      var chemResult = _buildChimicaleSheet(client, sorted, prices);
      var chimName = devizType === 2 ? baseName.substring(0, 27) + '_Chim' : baseName;
      if (wb.SheetNames.indexOf(chimName) >= 0) chimName = chimName.substring(0, 24) + '_' + (sheetCount + 1);
      XLSX.utils.book_append_sheet(wb, chemResult.ws, chimName);

      if (devizType === 2) {
        // Build V2 operations sheet (inline — same structure as exportDevizComplet's sheet 2)
        var numOps = opsList.length;
        var v2TotalCols = 1 + numOps;
        var v2LastCol = v2TotalCols - 1;
        var ws2 = {};
        var merges2 = [];
        var r = 0;

        var today = new Date();
        var todayStr = ('0' + today.getDate()).slice(-2) + '.' + ('0' + (today.getMonth() + 1)).slice(-2) + '.' + today.getFullYear();
        var todayYMD = today.toISOString().split('T')[0].replace(/-/g, '');
        var firstDate = sorted.length ? fmtDateDMY(sorted[0].date) : '';
        var lastDate = sorted.length ? fmtDateDMY(sorted[sorted.length - 1].date) : '';
        var period = firstDate + ' - ' + lastDate;
        var docNr = 'D-' + todayYMD + '-' + (client.client_id || '').slice(-4);

        var sNavy = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { color: { rgb: 'FFFFFF' }, sz: 1 }, border: _BORDER_THIN };
        var sAccent = { fill: { fgColor: { rgb: '4DB8E8' } }, font: { sz: 1 }, border: _BORDER_THIN };
        var sTitle = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var sLabelRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { bold: true, sz: 8 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var sValueRow = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var sSepRow = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 1 }, border: _BORDER_THIN };
        var sHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
        var sSubHeader = { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 8, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
        var sCompany1 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var sCompany2 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };
        var sCompany3 = { fill: { fgColor: { rgb: 'CDE3F5' } }, font: { sz: 8.5 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: _BORDER_THIN };

        // Row 0: navy
        _fillEmptyCells(ws2, r, v2TotalCols, sNavy);
        merges2.push({ s: { r: 0, c: 0 }, e: { r: 0, c: v2LastCol } });
        r++;

        // Row 1: company
        var thirdW = Math.max(Math.floor(v2TotalCols / 3), 1);
        var sec1End = Math.min(thirdW - 1, v2LastCol);
        var sec2End = Math.min(thirdW * 2 - 1, v2LastCol);
        var sec3End = v2LastCol;

        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('S.C. AQUATIS ENGINEERING S.R.L.', sCompany1);
        for (var ci = 1; ci <= sec1End; ci++) ws2[XLSX.utils.encode_cell({ r: r, c: ci })] = _cellS('', sCompany1);
        if (sec1End > 0) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: sec1End } });
        ws2[XLSX.utils.encode_cell({ r: r, c: sec1End + 1 })] = _cellS('office@aquatis.ro\nwww.aquatis.ro', sCompany2);
        for (var ci2 = sec1End + 2; ci2 <= sec2End; ci2++) ws2[XLSX.utils.encode_cell({ r: r, c: ci2 })] = _cellS('', sCompany2);
        if (sec2End > sec1End + 1) merges2.push({ s: { r: r, c: sec1End + 1 }, e: { r: r, c: sec2End } });
        ws2[XLSX.utils.encode_cell({ r: r, c: sec2End + 1 })] = _cellS('J40/18144/2007\nCUI: RO22479695', sCompany3);
        for (var ci3 = sec2End + 2; ci3 <= sec3End; ci3++) ws2[XLSX.utils.encode_cell({ r: r, c: ci3 })] = _cellS('', sCompany3);
        if (sec3End > sec2End + 1) merges2.push({ s: { r: r, c: sec2End + 1 }, e: { r: r, c: sec3End } });
        r++;

        // Row 2: accent
        _fillEmptyCells(ws2, r, v2TotalCols, sAccent);
        merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
        r++;

        // Row 3: title
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('RAPORT SERVICII \u2014 ABONAMENT \u00cENTRE\u021aINERE', sTitle);
        _fillEmptyCells(ws2, r, v2TotalCols, sTitle);
        merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
        r++;

        // Row 4: labels
        var labQ2 = Math.max(Math.floor(v2TotalCols / 4), 1);
        var labGroups2 = [
          { label: 'Client', start: 0, end: Math.min(labQ2 - 1, v2LastCol) },
          { label: 'Luna / Perioada', start: Math.min(labQ2, v2LastCol), end: Math.min(labQ2 * 2 - 1, v2LastCol) },
          { label: 'Nr. Document', start: Math.min(labQ2 * 2, v2LastCol), end: Math.min(labQ2 * 3 - 1, v2LastCol) },
          { label: 'Data emiterii', start: Math.min(labQ2 * 3, v2LastCol), end: v2LastCol }
        ];
        var valValues2 = [client.name || '', period, docNr, todayStr];
        labGroups2.forEach(function(g, gi) {
          if (g.start > v2LastCol) return;
          ws2[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(g.label, sLabelRow);
          for (var lc = g.start + 1; lc <= g.end; lc++) ws2[XLSX.utils.encode_cell({ r: r, c: lc })] = _cellS('', sLabelRow);
          if (g.end > g.start) merges2.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
        });
        r++;

        // Row 5: values
        labGroups2.forEach(function(g, gi) {
          if (g.start > v2LastCol) return;
          ws2[XLSX.utils.encode_cell({ r: r, c: g.start })] = _cellS(valValues2[gi], sValueRow);
          for (var vc = g.start + 1; vc <= g.end; vc++) ws2[XLSX.utils.encode_cell({ r: r, c: vc })] = _cellS('', sValueRow);
          if (g.end > g.start) merges2.push({ s: { r: r, c: g.start }, e: { r: r, c: g.end } });
        });
        r++;

        // Row 6: separator
        _fillEmptyCells(ws2, r, v2TotalCols, sSepRow);
        merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
        r++;

        // Row 7: header
        var v2HeaderRow1 = r;
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Data\ninterven\u021Bie', sHeader);
        if (numOps > 0) {
          ws2[XLSX.utils.encode_cell({ r: r, c: 1 })] = _cellS('SERVICII INCLUSE \u00cEN ABONAMENT', { fill: { fgColor: { rgb: '1D507F' } }, font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN });
          for (var hc = 2; hc <= v2LastCol; hc++) ws2[XLSX.utils.encode_cell({ r: r, c: hc })] = _cellS('', sHeader);
          if (numOps > 1) merges2.push({ s: { r: r, c: 1 }, e: { r: r, c: v2LastCol } });
        }
        r++;

        // Row 8: sub-headers
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('', sSubHeader);
        merges2.push({ s: { r: v2HeaderRow1, c: 0 }, e: { r: r, c: 0 } });
        var opNameWrap = {
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
        opsList.forEach(function(op, oi) {
          var label = opNameWrap[op] || op.replace(/ /g, '\n');
          ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellS(label, sSubHeader);
        });
        r++;

        // Data rows
        var v2DataStart = r;
        sorted.forEach(function(inv, idx) {
          var isEven = idx % 2 === 0;
          var bgColor = isEven ? 'E0EEF8' : 'F0F6FB';
          var sDataCell = { fill: { fgColor: { rgb: bgColor } }, font: { sz: 9 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
          var sCheckCell = { fill: { fgColor: { rgb: bgColor } }, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
          ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS(fmtDateDMY(inv.date), sDataCell);
          var ops = inv.operations || [];
          opsList.forEach(function(op, oi) {
            var hasOp = ops.indexOf(op) >= 0;
            ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellS(hasOp ? '\u2713' : '', hasOp ? sCheckCell : sDataCell);
          });
          r++;
        });
        var v2DataEnd = r - 1;

        // Total row
        var sTotalDark = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var sTotalDarkL = { fill: { fgColor: { rgb: '0D2D5A' } }, font: { bold: true, sz: 9, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Total interven\u021Bii efectuate', sTotalDarkL);
        opsList.forEach(function(op, oi) {
          var colLetter = XLSX.utils.encode_cell({ r: 0, c: 1 + oi }).replace(/[0-9]/g, '');
          var formula = 'COUNTIF(' + colLetter + (v2DataStart + 1) + ':' + colLetter + (v2DataEnd + 1) + ',"\u2713")';
          ws2[XLSX.utils.encode_cell({ r: r, c: 1 + oi })] = _cellF(formula, sTotalDark);
        });
        r++;

        // Separator
        _fillEmptyCells(ws2, r, v2TotalCols, sSepRow);
        merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: v2LastCol } });
        r++;

        // Total de plata
        var totalPlata = (prices.pret_interventie || 0) * sorted.length;
        var usedChemCols = chemResult.usedCols || [];
        usedChemCols.forEach(function(c) {
          var s = 0;
          sorted.forEach(function(inv) { s += parseFloat(inv[c.key]) || 0; });
          totalPlata += s * (prices[c.priceKey] || 0);
        });

        var sTotalPay = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
        var sTotalPayVal = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: _BORDER_THIN };
        var payLabelEnd = Math.max(v2LastCol - 2, 0);
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('TOTAL DE PLAT\u0102 (RON)', sTotalPay);
        for (var pc = 1; pc <= payLabelEnd; pc++) ws2[XLSX.utils.encode_cell({ r: r, c: pc })] = _cellS('', sTotalPay);
        if (payLabelEnd > 0) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: payLabelEnd } });
        var payValStart = payLabelEnd + 1;
        ws2[XLSX.utils.encode_cell({ r: r, c: payValStart })] = _cellS(totalPlata + ' lei', sTotalPayVal);
        for (var pv = payValStart + 1; pv <= v2LastCol; pv++) ws2[XLSX.utils.encode_cell({ r: r, c: pv })] = _cellS('', sTotalPayVal);
        if (v2LastCol > payValStart) merges2.push({ s: { r: r, c: payValStart }, e: { r: r, c: v2LastCol } });
        r++;

        // Footer
        var sFooterL = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'left', vertical: 'center' }, border: _BORDER_THIN };
        var sFooterR = { fill: { fgColor: { rgb: 'E8F3FB' } }, font: { sz: 7.5 }, alignment: { horizontal: 'right', vertical: 'center' }, border: _BORDER_THIN };
        var halfCols2 = Math.floor(v2TotalCols / 2);
        ws2[XLSX.utils.encode_cell({ r: r, c: 0 })] = _cellS('Document generat de S.C. Aquatis Engineering S.R.L.', sFooterL);
        for (var fc = 1; fc < halfCols2; fc++) ws2[XLSX.utils.encode_cell({ r: r, c: fc })] = _cellS('', sFooterL);
        if (halfCols2 > 1) merges2.push({ s: { r: r, c: 0 }, e: { r: r, c: halfCols2 - 1 } });
        ws2[XLSX.utils.encode_cell({ r: r, c: halfCols2 })] = _cellS('www.aquatis.ro  |  0721.137.100', sFooterR);
        for (var fc2 = halfCols2 + 1; fc2 <= v2LastCol; fc2++) ws2[XLSX.utils.encode_cell({ r: r, c: fc2 })] = _cellS('', sFooterR);
        if (v2LastCol > halfCols2) merges2.push({ s: { r: r, c: halfCols2 }, e: { r: r, c: v2LastCol } });
        r++;

        ws2['!ref'] = XLSX.utils.encode_cell({ r: 0, c: 0 }) + ':' + XLSX.utils.encode_cell({ r: r - 1, c: v2LastCol });
        ws2['!merges'] = merges2;
        var v2Cols = [{ wch: 16 }, { wch: 13 }];
        for (var wc = 1; wc < numOps; wc++) v2Cols.push({ wch: 10 });
        ws2['!cols'] = v2Cols;

        var opsName = baseName.substring(0, 28) + '_Ops';
        if (wb.SheetNames.indexOf(opsName) >= 0) opsName = opsName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws2, opsName);
      }

      sheetCount++;
    }

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
        mimeType: mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
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

