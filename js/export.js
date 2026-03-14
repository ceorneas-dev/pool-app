// export.js — Excel export via SheetJS (Community Edition, CDN)
// Exports: per-client XLSX (3 sheets) + all-clients XLSX

'use strict';

const SHEETJS_CDN = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';

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
    _uploadToDrive(wb, filename);
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
    _uploadToDrive(wb, filename);
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
    _uploadToDrive(wb, filename);
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

// == Build chimicale sheet — only used chemicals ==
function _buildChimicaleSheet(client, sorted, prices) {
  // Filter to only chemicals that have at least one non-zero value
  var usedCols = ALL_CHEM_COLS.filter(function(c) {
    return sorted.some(function(i) { return (parseFloat(i[c.key]) || 0) > 0; });
  });

  // Need: col0=Data, col1=Cant, col2..N=chemicals, colN+1=empty, colN+2=TOTAL PLATA
  var totalCols = 2 + usedCols.length + 2; // data+cant + chemicals + empty + total

  var data = [];

  // Row 0: client name at [0], title starting at [2], TOTAL PLATA at last col
  var h0 = new Array(totalCols).fill('');
  h0[0] = client.name || '';
  h0[2] = 'C H I M I C A L E  FOLOSITE';
  h0[totalCols - 1] = 'TOTAL PLATA';
  data.push(h0);

  // Row 1: column headers
  var h1 = ['Data Interventie', 'Cant'];
  usedCols.forEach(function(c) { h1.push(c.label); });
  h1.push('', '');  // empty + empty (under TOTAL PLATA)
  data.push(h1);

  // Intervention data rows
  sorted.forEach(function(i) {
    var row = [fmtDateDMY(i.date), 1];
    usedCols.forEach(function(c) {
      var v = parseFloat(i[c.key]) || 0;
      row.push(v > 0 ? v : '');
    });
    row.push('', '');
    data.push(row);
  });

  // Empty separator
  data.push(new Array(totalCols).fill(''));

  // Cantitate totala
  var totRow = ['Cantitate totala', sorted.length];
  usedCols.forEach(function(c) {
    var s = 0;
    sorted.forEach(function(i) { s += parseFloat(i[c.key]) || 0; });
    totRow.push(s > 0 ? s : '');
  });
  totRow.push('', '');
  data.push(totRow);

  // Calculate total plata
  var totalPlata = (prices.pret_interventie || 0) * sorted.length;
  usedCols.forEach(function(c) {
    var s = 0;
    sorted.forEach(function(i) { s += parseFloat(i[c.key]) || 0; });
    totalPlata += s * (prices[c.priceKey] || 0);
  });

  // Pret unitar — show price for ALL chemicals (even ones not used, just matching the used columns)
  var priceRow = ['Pret unitar', prices.pret_interventie || 0];
  usedCols.forEach(function(c) { priceRow.push(prices[c.priceKey] || 0); });
  priceRow.push('', totalPlata);
  data.push(priceRow);

  // PRET TOTAL — only where quantity > 0
  var ptRow = ['PRET TOTAL', (prices.pret_interventie || 0) * sorted.length];
  usedCols.forEach(function(c) {
    var s = 0;
    sorted.forEach(function(i) { s += parseFloat(i[c.key]) || 0; });
    ptRow.push(s > 0 ? Math.round(s * (prices[c.priceKey] || 0) * 100) / 100 : '');
  });
  ptRow.push('', '');
  data.push(ptRow);

  // Column widths
  var colWidths = [{wch:18}, {wch:6}];
  usedCols.forEach(function() { colWidths.push({wch:12}); });
  colWidths.push({wch:3}, {wch:14});

  return { data: data, totalPlata: totalPlata, colWidths: colWidths };
}

// == Export Format 1: Deviz Chimicale (V1) ==
function exportDevizChimicale(client, interventions) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var wb = XLSX.utils.book_new();
    var sorted = interventions.slice().sort(function(a,b) { return a.date.localeCompare(b.date); });

    var result = _buildChimicaleSheet(client, sorted, prices);
    var ws = XLSX.utils.aoa_to_sheet(result.data);
    ws['!cols'] = result.colWidths;
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    var fname = 'Deviz_' + sanitizeFilename(client.name) + '_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname);
    return fname;
  });
}

// == Export Format 2: Deviz Complet (V2 — chimicale + operatiuni) ==
function exportDevizComplet(client, interventions) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var wb = XLSX.utils.book_new();
    var sorted = interventions.slice().sort(function(a,b) { return a.date.localeCompare(b.date); });

    // Sheet 1: Chimicale
    var result = _buildChimicaleSheet(client, sorted, prices);
    var ws1 = XLSX.utils.aoa_to_sheet(result.data);
    ws1['!cols'] = result.colWidths;
    XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');

    // Sheet 2: Operatiuni Efectuate
    // Get operations list from settings (same as app uses)
    var opsList = (typeof getOperations === 'function') ? await getOperations() : [
      'Aspirare piscina','Curatare linie apa','Curatare skimmere',
      'Spalare filtru','Curatare prefiltru','Periere piscina',
      'Analiza apei','Tratament chimic','Verificare automatizare'
    ];

    var opsData = [];

    // Row 0: headers spanning all columns
    var oh0 = ['Data interventie', 'Servicii incluse in abonament'];
    for (var oi = 1; oi < opsList.length; oi++) oh0.push('');
    opsData.push(oh0);

    // Row 1: individual operation names
    var oh1 = [''];
    opsList.forEach(function(op) { oh1.push(op); });
    opsData.push(oh1);

    // Data rows — P where operation was checked in the app
    sorted.forEach(function(i) {
      var row = [fmtDateDMY(i.date)];
      var ops = i.operations || [];
      opsList.forEach(function(op) {
        row.push(ops.indexOf(op) >= 0 ? 'P' : '');
      });
      opsData.push(row);
    });

    // Empty separator
    opsData.push(new Array(oh1.length).fill(''));

    // TOTAL de plata row
    var totalRow = ['TOTAL de plata'];
    for (var ti = 0; ti < opsList.length - 2; ti++) totalRow.push('');
    totalRow.push(result.totalPlata + ' lei');
    totalRow.push('');
    opsData.push(totalRow);

    var ws2 = XLSX.utils.aoa_to_sheet(opsData);
    var opsCols = [{wch:18}];
    opsList.forEach(function() { opsCols.push({wch:18}); });
    ws2['!cols'] = opsCols;
    XLSX.utils.book_append_sheet(wb, ws2, 'Operatiuni');

    var fname = 'DevizComplet_' + sanitizeFilename(client.name) + '_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname);
    return fname;
  });
}

// == Export ALL clients — Mixed V1/V2 per client.deviz_type ==
function exportAllDevizMixed(clients, allInterventions) {
  return loadXLSX().then(async function() {
    var prices = (typeof getExportPrices === 'function') ? await getExportPrices() : {};
    var opsList = (typeof getOperations === 'function') ? await getOperations() : [
      'Aspirare piscina','Curatare linie apa','Curatare skimmere',
      'Spalare filtru','Curatare prefiltru','Periere piscina',
      'Analiza apei','Tratament chimic','Verificare automatizare'
    ];
    var wb = XLSX.utils.book_new();
    var sheetCount = 0;

    clients.forEach(function(client) {
      var ci = allInterventions.filter(function(i) { return i.client_id === client.client_id; });
      if (!ci.length) return;

      var sorted = ci.slice().sort(function(a,b) { return a.date.localeCompare(b.date); });
      var devizType = parseInt(client.deviz_type) || 1;

      // Sheet: Chimicale (always present for both V1 and V2)
      var result = _buildChimicaleSheet(client, sorted, prices);
      var ws1 = XLSX.utils.aoa_to_sheet(result.data);
      ws1['!cols'] = result.colWidths;

      var baseName = sanitizeSheetName(client.name);

      if (devizType === 2) {
        // V2: two sheets — Chimicale + Operatiuni
        var chimName = baseName.substring(0, 27) + '_Chim';
        if (wb.SheetNames.indexOf(chimName) >= 0) chimName = chimName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws1, chimName);

        // Build operations sheet
        var opsData = [];
        var oh0 = ['Data interventie', 'Servicii incluse in abonament'];
        for (var oi = 1; oi < opsList.length; oi++) oh0.push('');
        opsData.push(oh0);

        var oh1 = [''];
        opsList.forEach(function(op) { oh1.push(op); });
        opsData.push(oh1);

        sorted.forEach(function(i) {
          var row = [fmtDateDMY(i.date)];
          var ops = i.operations || [];
          opsList.forEach(function(op) {
            row.push(ops.indexOf(op) >= 0 ? 'P' : '');
          });
          opsData.push(row);
        });

        opsData.push(new Array(oh1.length).fill(''));

        var totalRow = ['TOTAL de plata'];
        for (var ti = 0; ti < opsList.length - 2; ti++) totalRow.push('');
        totalRow.push(result.totalPlata + ' lei');
        totalRow.push('');
        opsData.push(totalRow);

        var ws2 = XLSX.utils.aoa_to_sheet(opsData);
        var opsCols = [{wch:18}];
        opsList.forEach(function() { opsCols.push({wch:18}); });
        ws2['!cols'] = opsCols;

        var opsName = baseName.substring(0, 28) + '_Ops';
        if (wb.SheetNames.indexOf(opsName) >= 0) opsName = opsName.substring(0, 24) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws2, opsName);
      } else {
        // V1: single sheet — Chimicale only
        var sheetName = baseName;
        if (wb.SheetNames.indexOf(sheetName) >= 0) sheetName = sheetName.substring(0, 28) + '_' + (sheetCount + 1);
        XLSX.utils.book_append_sheet(wb, ws1, sheetName);
      }

      sheetCount++;
    });

    if (sheetCount === 0) {
      showToast('Nicio interventie de exportat.', 'warning');
      return;
    }

    var fname = 'DevizToti_' + fmtDateExport(new Date()) + '.xlsx';
    XLSX.writeFile(wb, fname);
    _uploadToDrive(wb, fname);
    return fname;
  });
}

function _uploadToDrive(wb, fileName, mimeType) {
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
          showToast('Salvat in Drive: Export Interventii/' + fileName, 'success', 4000);
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
    _uploadToDrive(wb, fname);
    showToast('Deviz Excel descarcat: ' + fname, 'success');
  }).catch(function(e) {
    showToast('Eroare export: ' + e.message, 'error');
  });
}

