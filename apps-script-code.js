// apps-script-code.js — Google Apps Script backend for Pool Manager PWA
// Deploy ca Web App: Execute as ME, Access: Anyone
// Copiați TOTUL în Google Apps Script și rulați setupSheetStructure() o singură dată.

// ── IMPORTANT: Actualizați SPREADSHEET_ID cu ID-ul foii dvs ──
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// ── Sheet column definitions ──────────────────────────────────
const CLIENTS_COLS = [
  'client_id','name','phone','address','pool_volume_mc','pool_type',
  'active','notes','created_at','updated_at',
  'latitude','longitude','location_set',
  'deviz_type','pret_interventie','billing_interval_interventions',
  'visit_frequency_days','last_billing_date'
];

const INTERVENTIONS_COLS = [
  'intervention_id','client_id','client_name','technician_id','technician_name',
  'date','created_at','synced_at',
  'measured_chlorine','measured_ph','measured_temp','measured_hardness','measured_alkalinity','measured_salinity',
  'measured_tc','measured_cya',
  'rec_cl_gr','rec_cl_tab','rec_ph_kg','rec_anti_l',
  'treat_cl_granule_gr','treat_cl_tablete','treat_cl_tablete_export_gr','treat_cl_lichid_bidoane',
  'treat_ph_granule','treat_ph_lichid_bidoane',
  'treat_antialgic','treat_anticalcar','treat_floculant','treat_sare_saci','treat_bicarbonat',
  'observations','operations',
  'geo_lat','geo_lng','geo_accuracy',
  'arrival_time','departure_time','duration_minutes',
  'audio_file_url',
  'treatments_json',
  // New columns MUST be appended at the very end: getOrCreateSheet adds missing
  // headers at the sheet's end, and handlePushInterventions writes by column
  // position, so the array order must match the physical sheet order.
  'paid'
];

const TECHNICIANS_COLS = [
  'technician_id','name','username','password','role','active','last_sync'
];

const RULES_COLS = [
  'rule_id','pool_vol_min','pool_vol_max','cl_min','cl_max','ph_min','ph_max',
  'rec_cl_gr','rec_cl_tab','rec_ph_kg','rec_anti'
];

const SYNC_LOG_COLS = [
  'sync_id','technician_id','timestamp','records_pushed','records_pulled','status','error_message'
];

// Evidenta Checklist - un singur rand (suprascris la fiecare salvare)
const CHECKLIST_COLS = ['updated_at','title','items_json'];

// Jurnal audit (activitate generală) — un rând per acțiune (append), vizibil doar adminului în UI
const AUDIT_LOG_COLS = [
  'timestamp','technician_id','technician_name','log_action','client_id','client_name','details'
];

// ── HTTP Handlers ─────────────────────────────────────────────
function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';

  try {
    let result;
    if (action === 'ping') {
      result = { status: 'ok', sheets: 7, timestamp: new Date().toISOString() };
    } else if (action === 'pull') {
      result = handlePull(params);
    } else if (action === 'stats') {
      result = handleStats();
    } else if (action === 'getChecklist') {
      result = handleGetChecklist();
    } else if (action === 'getAuditLog') {
      result = handleGetAuditLog();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    let result;

    if (action === 'login') {
      result = handleLogin(body);
    } else if (action === 'push') {
      result = handlePush(body);
    } else if (action === 'saveChecklist') {
      result = handleSaveChecklist(body);
    } else if (action === 'saveExportToDrive') {
      result = handleSaveExportToDrive(body);
    } else if (action === 'sendEmail') {
      result = handleSendEmail(body);
    } else if (action === 'logAudit') {
      result = handleLogAudit(body);
    } else if (action === 'saveConfig') {
      result = handleSaveConfig(body);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ── GET: pull ─────────────────────────────────────────────────
function handlePull(params) {
  const type   = params.type || 'all';
  const result = {};

  if (type === 'all' || type === 'clients') {
    result.clients = sheetToObjects('clients', CLIENTS_COLS);
  }
  if (type === 'all' || type === 'technicians') {
    result.technicians = sheetToObjects('technicians', TECHNICIANS_COLS);
  }
  if (type === 'all' || type === 'rules') {
    result.treatment_rules = sheetToObjects('treatment_rules', RULES_COLS);
  }
  if (type === 'all' || type === 'interventions') {
    const techId = params.tech_id || '';
    const all    = sheetToObjects('interventions', INTERVENTIONS_COLS);
    result.interventions = techId ? all.filter(i => i.technician_id === techId) : all;
  }
  if (type === 'all' || type === 'config') {
    result.config = handleGetConfig();
  }

  return result;
}

// ── App config (permisiuni globale tehnicieni etc.) ───────────
// Sheet key/value: astfel o permisiune setată de admin ajunge pe toate telefoanele.
const APP_CONFIG_COLS = ['key', 'value'];

function handleGetConfig() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('app_config');
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const out = {};
  rows.slice(1).forEach(r => { if (r[0]) out[String(r[0])] = String(r[1]); });
  return out;
}

function handleSaveConfig(body) {
  const entries = body.entries || {};
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'app_config', APP_CONFIG_COLS);
  const rows  = sheet.getDataRange().getValues();
  const idx = {};
  rows.slice(1).forEach((r, i) => { if (r[0]) idx[String(r[0])] = i + 2; });
  Object.keys(entries).forEach(k => {
    const v = String(entries[k]);
    if (idx[k]) sheet.getRange(idx[k], 2).setValue(v);
    else sheet.appendRow([k, v]);
  });
  return { success: true };
}

// ── GET: stats ────────────────────────────────────────────────
function handleStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    clients:       getRowCount(ss, 'clients'),
    interventions: getRowCount(ss, 'interventions'),
    technicians:   getRowCount(ss, 'technicians'),
    rules:         getRowCount(ss, 'treatment_rules')
  };
}

// ── POST: login ───────────────────────────────────────────────
function handleLogin(body) {
  const { username, password } = body;
  if (!username || !password) {
    return { success: false, error: 'Username și parola sunt obligatorii' };
  }
  const techs = sheetToObjects('technicians', TECHNICIANS_COLS);
  const tech  = techs.find(t => t.username === username && t.password === password && t.active !== 'false');
  if (!tech) {
    return { success: false, error: 'Utilizator sau parolă incorectă' };
  }
  return {
    success: true,
    user: {
      technician_id: tech.technician_id,
      name:          tech.name,
      username:      tech.username,
      role:          tech.role
    }
  };
}

// ── POST: push interventions ──────────────────────────────────
function handlePush(body) {
  const type = body.type || 'interventions';
  if (type === 'clients') return handlePushClients(body);
  if (type === 'technicians') return handlePushTechnicians(body);
  if (type === 'delete_intervention') return handleDeleteIntervention(body);
  if (type === 'delete_interventions_bulk') return handleDeleteInterventionsBulk(body);
  if (type === 'clear_interventions') return handleClearInterventions();
  return handlePushInterventions(body);
}

/** Clear ALL interventions from the sheet (keeps header row) */
function handleClearInterventions() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'interventions', INTERVENTIONS_COLS);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return { success: true, cleared: lastRow - 1 };
}

/** Delete a single intervention by intervention_id */
function handleDeleteIntervention(body) {
  var iid = (body.data && body.data.intervention_id) || '';
  if (!iid) return { success: false, error: 'Missing intervention_id' };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, 'interventions', INTERVENTIONS_COLS);
  var rows  = sheet.getDataRange().getValues();
  var idCol = INTERVENTIONS_COLS.indexOf('intervention_id');
  if (idCol < 0) return { success: false, error: 'Column intervention_id not found' };

  for (var r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][idCol]) === String(iid)) {
      sheet.deleteRow(r + 1);
      return { success: true, deleted: iid };
    }
  }
  return { success: true, deleted: null, message: 'Not found on server' };
}

function handlePushInterventions(body) {
  const { data, tech_id } = body;
  if (!data || !Array.isArray(data)) {
    return { success: false, error: 'Lipsesc datele' };
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'interventions', INTERVENTIONS_COLS);
  const rows  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString();

  // Map field name → physical column index from the ACTUAL sheet header row, then
  // read and write every value BY NAME. Writing by array position corrupts data
  // whenever the sheet's physical column order drifts from INTERVENTIONS_COLS
  // (e.g. a column added mid-array in a newer version was appended at the end of
  // an older sheet) — a value then lands under the wrong header. Header-based
  // mapping is order-independent and matches how sheetToObjects reads on pull.
  const header = (rows.length && rows[0].length) ? rows[0].map(String) : INTERVENTIONS_COLS.slice();
  const width  = header.length;
  const colIdx = {};
  header.forEach((h, idx) => { colIdx[h] = idx; });

  const idCol      = colIdx['intervention_id'];
  const dateCol    = colIdx['date'];
  const clientCol  = colIdx['client_id'];
  const techCol    = colIdx['technician_id'];
  const createdCol = colIdx['created_at'];

  // Build index: intervention_id → row number (1-based)
  const existingRows = {};
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][idCol] || '').trim();
    if (id) existingRows[id] = i + 1;
  }

  // Also build index: client_id+tech_id+date → {row, created_at} for server-side dedup
  const dedupIndex = {};
  for (let i = 1; i < rows.length; i++) {
    const cid = String(rows[i][clientCol] || '').trim();
    const tid = String(rows[i][techCol] || '').trim();
    const dt  = _cellToString(rows[i][dateCol]).substring(0, 10);
    const key = cid + '_' + tid + '_' + dt;
    if (cid && dt) {
      dedupIndex[key] = { rowNum: i + 1, created_at: String(rows[i][createdCol] || ''), intervention_id: String(rows[i][idCol] || '').trim() };
    }
  }

  let pushed = 0, updated = 0;
  const rowsToDelete = []; // row numbers to delete (old duplicates)

  data.forEach(item => {
    const rowNum = existingRows[item.intervention_id];
    // Start from the existing row (update) so any extra/unknown columns are kept,
    // or a blank row (insert); then place each known field at its header index.
    const row = rowNum ? rows[rowNum - 1].slice() : [];
    while (row.length < width) row.push('');
    INTERVENTIONS_COLS.forEach(col => {
      const idx = colIdx[col];
      if (idx === undefined) return; // header not present (shouldn't happen after getOrCreateSheet)
      const val = (col === 'synced_at') ? now : item[col];
      row[idx] = (val !== undefined && val !== null) ? String(val) : '';
    });
    if (rowNum) {
      // UPSERT: update existing row
      sheet.getRange(rowNum, 1, 1, width).setValues([row]);
      updated++;
    } else {
      // Server-side dedup: check if another intervention exists for same client+date
      const dk = String(item.client_id) + '_' + String(item.technician_id || '') + '_' + String(item.date || '').substring(0, 10);
      const existing = dedupIndex[dk];
      if (existing && existing.intervention_id !== item.intervention_id) {
        const newCreated = item.created_at || '';
        if (newCreated >= existing.created_at) {
          rowsToDelete.push(existing.rowNum);
        } else {
          return; // Old is newer → skip this older duplicate
        }
      }
      sheet.appendRow(row);
      pushed++;
      dedupIndex[dk] = { rowNum: -1, created_at: item.created_at || '', intervention_id: item.intervention_id };
    }
  });

  // Delete old duplicate rows (bottom-to-top to preserve indices)
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(r => { try { sheet.deleteRow(r); } catch(e) {} });
  }

  logSync(ss, tech_id, pushed + updated, 0, 'success', '');
  return { success: true, pushed, updated, deduped: rowsToDelete.length };
}

/** Delete multiple interventions by IDs (bulk sync of deleted_intervention_ids) */
function handleDeleteInterventionsBulk(body) {
  const ids = body.data && body.data.ids;
  if (!ids || !Array.isArray(ids) || !ids.length) return { success: true, deleted: 0 };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('interventions');
  if (!sheet) return { success: true, deleted: 0 };

  const data    = sheet.getDataRange().getValues();
  const idCol   = INTERVENTIONS_COLS.indexOf('intervention_id');
  const idsSet  = {};
  ids.forEach(function(id) { idsSet[String(id)] = true; });

  // Delete from bottom to top to preserve row indices
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (idsSet[String(data[i][idCol]).trim()]) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { success: true, deleted };
}

function handlePushClients(body) {
  const { data } = body;
  if (!data || !Array.isArray(data)) {
    return { success: false, error: 'Lipsesc datele' };
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'clients', CLIENTS_COLS);
  const rows  = sheet.getDataRange().getValues();

  // Map field name → physical column index from the real header, then read/write
  // by NAME (order-independent). Writing by array position corrupts data when the
  // sheet's column order drifts from CLIENTS_COLS (e.g. billing/deviz columns
  // added in a newer version were appended to the end of an older sheet).
  const header = (rows.length && rows[0].length) ? rows[0].map(String) : CLIENTS_COLS.slice();
  const width  = header.length;
  const colIdx = {};
  header.forEach((h, idx) => { colIdx[h] = idx; });
  const idCol = colIdx['client_id'];

  const existingRows = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol]) existingRows[rows[i][idCol]] = i + 1;
  }

  let saved = 0, updated = 0;

  data.forEach(client => {
    const rowNum = existingRows[client.client_id];
    const row = rowNum ? rows[rowNum - 1].slice() : [];
    while (row.length < width) row.push('');
    CLIENTS_COLS.forEach(col => {
      const idx = colIdx[col];
      if (idx === undefined) return;
      const val = client[col];
      row[idx] = (val !== undefined && val !== null) ? String(val) : '';
    });
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, width).setValues([row]);
      updated++;
    } else {
      sheet.appendRow(row);
      saved++;
    }
  });

  return { success: true, saved, updated };
}

// ── POST: pushTechnicians ─────────────────────────────────────
function handlePushTechnicians(body) {
  const { data } = body;
  if (!data || !Array.isArray(data)) {
    return { success: false, error: 'Lipsesc datele' };
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'technicians', TECHNICIANS_COLS);
  const rows  = sheet.getDataRange().getValues();

  // Map field name → physical column index from the real header; read/write by NAME.
  const header = (rows.length && rows[0].length) ? rows[0].map(String) : TECHNICIANS_COLS.slice();
  const width  = header.length;
  const colIdx = {};
  header.forEach((h, idx) => { colIdx[h] = idx; });
  const idCol = colIdx['technician_id'];

  const existingRows = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol]) existingRows[rows[i][idCol]] = i + 1;
  }

  let saved = 0, updated = 0, deleted = 0;

  // Process deletions first (reverse order to keep row indices valid)
  const toDelete = [];
  data.forEach(tech => {
    if (tech._deleted && existingRows[tech.technician_id]) {
      toDelete.push(existingRows[tech.technician_id]);
    }
  });
  toDelete.sort((a, b) => b - a).forEach(rowNum => {
    sheet.deleteRow(rowNum);
    deleted++;
  });

  // Re-read rows after deletions (header unchanged, colIdx still valid)
  const rows2 = sheet.getDataRange().getValues();
  const existingRows2 = {};
  for (let i = 1; i < rows2.length; i++) {
    if (rows2[i][idCol]) existingRows2[rows2[i][idCol]] = i + 1;
  }

  data.forEach(tech => {
    if (tech._deleted) return; // already handled above
    const rowNum = existingRows2[tech.technician_id];
    const row = rowNum ? rows2[rowNum - 1].slice() : [];
    while (row.length < width) row.push('');
    TECHNICIANS_COLS.forEach(col => {
      const idx = colIdx[col];
      if (idx === undefined) return;
      const val = tech[col];
      row[idx] = (val !== undefined && val !== null) ? String(val) : '';
    });
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, width).setValues([row]);
      updated++;
    } else {
      sheet.appendRow(row);
      saved++;
    }
  });

  return { success: true, saved, updated, deleted };
}

// ── GET: getChecklist ────────────────────────────────────────────────
// Returneaza titlul si itemii listei de evidenta (un singur rand in sheet).
function handleGetChecklist() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('checklist');
  if (!sheet) return { success: true, data: null };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, data: null };

  // CHECKLIST_COLS: [updated_at, title, items_json]
  const [updated_at, title, items_json] = data[1].map(v => String(v || ''));
  return { success: true, data: { updated_at, title, items_json } };
}

// ── POST: saveChecklist ──────────────────────────────────────────────────
// Suprascrie randul unic din sheet cu starea curenta a listei.
function handleSaveChecklist(body) {
  const { title, items_json, updated_at } = body;
  if (items_json === undefined) return { success: false, error: 'Lipsesc datele' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'checklist', CHECKLIST_COLS);
  const data  = sheet.getDataRange().getValues();

  const row = [updated_at || new Date().toISOString(), title || '', items_json || '[]'];
  if (data.length <= 1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(2, 1, 1, CHECKLIST_COLS.length).setValues([row]);
  }
  return { success: true };
}

// ── POST: logAudit ─────────────────────────────────────────────
// Adaugă o intrare în jurnalul de activitate (client/intervenție/locație adăugat(ă)/schimbat(ă)/șters(ă)).
function handleLogAudit(body) {
  const { technician_id, technician_name, log_action, client_id, client_name, details, timestamp } = body;
  if (!log_action) return { success: false, error: 'Lipsește log_action' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'audit_log', AUDIT_LOG_COLS);
  sheet.appendRow([
    timestamp || new Date().toISOString(),
    technician_id   || '',
    technician_name || '',
    log_action,
    client_id   || '',
    client_name || '',
    details     || ''
  ]);
  return { success: true };
}

// ── GET: getAuditLog ───────────────────────────────────────────
// Returnează ultimele 300 intrări din jurnal, cele mai noi primele.
function handleGetAuditLog() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('audit_log');
  if (!sheet) return { entries: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { entries: [] };

  const entries = data.slice(1).map(row => ({
    timestamp:        _cellToString(row[0]),
    technician_id:    String(row[1]),
    technician_name:  String(row[2]),
    log_action:       String(row[3]),
    client_id:        String(row[4]),
    client_name:      String(row[5]),
    details:          String(row[6] || '')
  }));

  entries.sort((a, b) => (a.timestamp < b.timestamp ? 1 : -1)); // newest first
  return { entries: entries.slice(0, 300) };
}

/** Save exported file to a Google Drive folder (default "Export Interventii", or body.rootFolder) */
function handleSaveExportToDrive(body) {
  var fileName = body.fileName || 'export.xlsx';
  var base64Data = body.data || '';
  var mimeType = body.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  if (!base64Data) return { error: 'No file data provided' };

  // Find or create root folder (default "Export Interventii", overridable e.g. for backups)
  var folderName = body.rootFolder || 'Export Interventii';
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // If clientName provided, find or create client subfolder
  var clientName = body.clientName || '';
  if (clientName) {
    var subFolders = folder.getFoldersByName(clientName);
    if (subFolders.hasNext()) {
      folder = subFolders.next();
    } else {
      folder = folder.createFolder(clientName);
    }
  }

  // Create file from base64
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = folder.createFile(blob);

  return {
    success: true,
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function handleSendEmail(body) {
  var to = body.to || '';
  var subject = body.subject || 'Pool Manager - Notificare';
  var emailBody = body.body || '';

  // Fallback: no address given → send to the account that owns this script.
  if (!to) {
    try { to = Session.getEffectiveUser().getEmail(); } catch (e) {}
  }
  if (!to) return { error: 'No email address provided' };

  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: emailBody
    });
    return { success: true };
  } catch (e) {
    return { error: 'Email send failed: ' + e.message };
  }
}

// ── Sheet helpers ─────────────────────────────────────────────
/** Convert a value to string, handling Date objects → YYYY-MM-DD */
function _cellToString(val) {
  if (val === undefined || val === null || val === '') return '';
  // Google Sheets converts date strings to Date objects — normalize them back
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    var h = val.getHours(), min = val.getMinutes(), sec = val.getSeconds();
    // Time-only value: Sheets uses base date 1899-12-30 for time fields
    if (y <= 1900) return ('0'+h).slice(-2) + ':' + ('0'+min).slice(-2);
    // Date-only (no time component)
    if (h === 0 && min === 0 && sec === 0) return y + '-' + m + '-' + d;
    // Date + time
    return y + '-' + m + '-' + d + 'T' + ('0'+h).slice(-2) + ':' + ('0'+min).slice(-2) + ':' + ('0'+sec).slice(-2);
  }
  return String(val);
}

function sheetToObjects(sheetName, cols) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(String);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = _cellToString(row[i]); });
    return obj;
  });
}

function getRowCount(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  const count = sheet.getLastRow() - 1;
  return Math.max(0, count);
}

function getOrCreateSheet(ss, name, cols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(cols);
    sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    var lastCol = sheet.getLastColumn() || 0;
    var currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    var missing = [];
    cols.forEach(function(col) {
      if (currentHeaders.indexOf(col) < 0) missing.push(col);
    });
    if (missing.length > 0) {
      var startCol = lastCol + 1;
      for (var m = 0; m < missing.length; m++) {
        sheet.getRange(1, startCol + m).setValue(missing[m]).setFontWeight('bold');
      }
    }
  }
  return sheet;
}

function logSync(ss, techId, pushed, pulled, status, errorMsg) {
  try {
    const sheet = getOrCreateSheet(ss, 'sync_log', SYNC_LOG_COLS);
    sheet.appendRow([
      'sl_' + Date.now(),
      techId || '',
      new Date().toISOString(),
      pushed,
      pulled,
      status,
      errorMsg
    ]);
  } catch (e) {
    // Log errors are non-critical
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Cleanup: remove duplicate interventions (keep newest per client+date) ──
// Run this ONCE from Apps Script editor to clean existing duplicates
function cleanDuplicateInterventions() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('interventions');
  if (!sheet) { Logger.log('No interventions sheet'); return; }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { Logger.log('No interventions to clean'); return; }

  // Read column positions from the ACTUAL header row (not INTERVENTIONS_COLS
  // array order) so a sheet whose columns drifted isn't grouped by the wrong
  // columns — which would delete non-duplicate rows.
  var header = data[0].map(String);
  var clientIdCol  = header.indexOf('client_id');
  var dateCol      = header.indexOf('date');
  var createdAtCol = header.indexOf('created_at');

  // Group by client_id + date, keep the newest (by created_at)
  var best = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][clientIdCol]) + '_' + String(data[i][dateCol]);
    var ca  = String(data[i][createdAtCol] || '');
    if (!best[key] || ca > best[key].created_at) {
      best[key] = { rowIndex: i, created_at: ca };
    }
  }

  var keepRows = {};
  Object.keys(best).forEach(function(k) { keepRows[best[k].rowIndex] = true; });

  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (!keepRows[i]) rowsToDelete.push(i + 1);
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var d = 0; d < rowsToDelete.length; d++) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  Logger.log('Cleaned ' + rowsToDelete.length + ' duplicate interventions. Kept ' + Object.keys(best).length + ' unique entries.');
}

// ── Setup (run once!) ─────────────────────────────────────────
function setupSheetStructure() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Create all sheets with headers and colors
  const sheets = [
    { name: 'clients',          cols: CLIENTS_COLS,       color: '#1e40af' }, // blue
    { name: 'interventions',    cols: INTERVENTIONS_COLS,  color: '#15803d' }, // green
    { name: 'treatment_rules',  cols: RULES_COLS,          color: '#b45309' }, // yellow/amber
    { name: 'technicians',      cols: TECHNICIANS_COLS,    color: '#b91c1c' }, // red
    { name: 'sync_log',         cols: SYNC_LOG_COLS,       color: '#475569' }, // gray
    { name: 'checklist',        cols: CHECKLIST_COLS,        color: '#b45309' },  // amber - evidenta checklist
    { name: 'audit_log',        cols: AUDIT_LOG_COLS,        color: '#374151' },  // gri inchis - jurnal audit
    { name: 'app_config',       cols: APP_CONFIG_COLS,       color: '#0891b2' }   // cyan - config global (permisiuni)
  ];

  sheets.forEach(({ name, cols, color }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // Set headers
    const headerRange = sheet.getRange(1, 1, 1, cols.length);
    headerRange.setValues([cols]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground(color);
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Auto-resize columns
    sheet.autoResizeColumns(1, cols.length);
  });

  // Create default admin user
  const techSheet = ss.getSheetByName('technicians');
  const existing  = techSheet.getDataRange().getValues();
  if (existing.length <= 1) {
    techSheet.appendRow(['t_001', 'Administrator', 'admin', 'admin123', 'admin', 'true', '']);
    techSheet.appendRow(['t_002', 'Tehnician Demo', 'dan', 'dan123', 'technician', 'true', '']);
  }

  Logger.log('✅ Structura Google Sheets creată cu succes!');
  Logger.log('URL-ul Web App: ' + ScriptApp.getService().getUrl());
}
