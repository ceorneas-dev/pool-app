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
  'rec_cl_gr','rec_cl_tab','rec_ph_kg','rec_anti_l',
  'treat_cl_granule_gr','treat_cl_tablete','treat_cl_tablete_export_gr','treat_cl_lichid_bidoane',
  'treat_ph_granule','treat_ph_lichid_bidoane',
  'treat_antialgic','treat_anticalcar','treat_floculant','treat_sare_saci','treat_bicarbonat',
  'observations',
  'geo_lat','geo_lng','geo_accuracy',
  'arrival_time','departure_time','duration_minutes'
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

const LOCATIONS_COLS = [
  'technician_id','name','lat','lng','accuracy','timestamp'
];

// Istoric GPS — un rând per trimitere (append), păstrat 30 zile
const LOCATIONS_HIST_COLS = [
  'timestamp','technician_id','name','lat','lng','accuracy'
];

// Un rând per pereche (tech+admin) — actualizat la fiecare cerere
const AUDIO_CALLS_COLS = [
  'tech_id','admin_id','admin_name','channel','status','updated_at'
];

// Programul intervențiilor — un rând per intrare planificată
const PROGRAM_COLS = [
  'id','date','time','technician_id','technician_name','client_name','address','notes'
];

// Evidenta Checklist - un singur rand (suprascris la fiecare salvare)
const CHECKLIST_COLS = ['updated_at','title','items_json'];

// ── HTTP Handlers ─────────────────────────────────────────────
function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';

  try {
    let result;
    if (action === 'ping') {
      result = { status: 'ok', sheets: 5, timestamp: new Date().toISOString() };
    } else if (action === 'pull') {
      result = handlePull(params);
    } else if (action === 'stats') {
      result = handleStats();
    } else if (action === 'getLocations') {
      result = handleGetLocations();
    } else if (action === 'getLocationHistory') {
      result = handleGetLocationHistory(params);
    } else if (action === 'getCalendar') {
      result = handleGetCalendar(params);
    } else if (action === 'getChecklist') {
      result = handleGetChecklist();
    } else if (action === 'getAudioCall') {
      result = handleGetAudioCall(params);
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
    } else if (action === 'saveLocation') {
      result = handleSaveLocation(body);
    } else if (action === 'saveCalendarEntries') {
      result = handleSaveCalendarEntries(body);
    } else if (action === 'deleteCalendarEntry') {
      result = handleDeleteCalendarEntry(body);
    } else if (action === 'saveChecklist') {
      result = handleSaveChecklist(body);
    } else if (action === 'requestAudioCall') {
      result = handleRequestAudioCall(body);
    } else if (action === 'updateAudioCall') {
      result = handleUpdateAudioCall(body);
    } else if (action === 'saveExportToDrive') {
      result = handleSaveExportToDrive(body);
    } else if (action === 'sendEmail') {
      result = handleSendEmail(body);
    } else if (body._type === 'location') {
      // OwnTracks HTTP mode — trimite direct fără câmpul "action"
      result = handleOwnTracksLocation(body);
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

  return result;
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
  return handlePushInterventions(body);
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

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet    = getOrCreateSheet(ss, 'interventions', INTERVENTIONS_COLS);
  const existing = sheetToObjects('interventions', INTERVENTIONS_COLS).map(i => i.intervention_id);

  let pushed = 0, duplicates = 0;
  const now = new Date().toISOString();

  data.forEach(item => {
    if (existing.includes(item.intervention_id)) {
      duplicates++;
      return;
    }
    const row = INTERVENTIONS_COLS.map(col => {
      if (col === 'synced_at') return now;
      const val = item[col];
      return val !== undefined && val !== null ? String(val) : '';
    });
    sheet.appendRow(row);
    pushed++;
  });

  logSync(ss, tech_id, pushed, 0, 'success', '');
  return { success: true, pushed, duplicates };
}

function handlePushClients(body) {
  const { data } = body;
  if (!data || !Array.isArray(data)) {
    return { success: false, error: 'Lipsesc datele' };
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'clients', CLIENTS_COLS);
  const rows  = sheet.getDataRange().getValues();
  const idCol = CLIENTS_COLS.indexOf('client_id');

  const existingRows = {};
  rows.slice(1).forEach((row, i) => {
    if (row[idCol]) existingRows[row[idCol]] = i + 2;
  });

  let saved = 0, updated = 0;

  data.forEach(client => {
    const row = CLIENTS_COLS.map(col => {
      const val = client[col];
      return val !== undefined && val !== null ? String(val) : '';
    });
    const rowNum = existingRows[client.client_id];
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
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
  const idCol = TECHNICIANS_COLS.indexOf('technician_id');

  const existingRows = {};
  rows.slice(1).forEach((row, i) => {
    if (row[idCol]) existingRows[row[idCol]] = i + 2;
  });

  let saved = 0, updated = 0;

  data.forEach(tech => {
    const row = TECHNICIANS_COLS.map(col => {
      const val = tech[col];
      return val !== undefined && val !== null ? String(val) : '';
    });

    const rowNum = existingRows[tech.technician_id];
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      sheet.appendRow(row);
      saved++;
    }
  });

  return { success: true, saved, updated };
}

// ── POST: saveLocation ────────────────────────────────────────
// Salvează sau actualizează poziția unui tehnician (un rând per tehnician).
function handleSaveLocation(body) {
  const { technician_id, name, lat, lng, accuracy, timestamp } = body;
  if (!technician_id || lat === undefined || lng === undefined) {
    return { success: false, error: 'Date GPS incomplete' };
  }
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'locatii', LOCATIONS_COLS);
  const data  = sheet.getDataRange().getValues();

  // Caută rândul existent pentru acest tehnician (actualizare, nu append)
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(technician_id)) {
      sheet.getRange(i + 1, 1, 1, LOCATIONS_COLS.length).setValues([
        [technician_id, name || '', lat, lng, accuracy || 0, timestamp]
      ]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([technician_id, name || '', lat, lng, accuracy || 0, timestamp]);
  }

  // ── Append în istoricul GPS (un rând per trimitere, 30 zile retenție) ──
  const histSheet = getOrCreateSheet(ss, 'locatii_istoric', LOCATIONS_HIST_COLS);
  histSheet.appendRow([timestamp, technician_id, name || '', lat, lng, accuracy || 0]);

  // Curățare o dată pe zi (șterge rânduri mai vechi de 30 zile)
  _cleanupGpsHistory(ss);

  return { success: true };
}

// ── GET: getLocationHistory ────────────────────────────────────
// Returnează pozițiile unui tehnician pentru o dată specifică.
function handleGetLocationHistory(params) {
  const techId = params.tech_id;
  const date   = params.date;   // format YYYY-MM-DD
  if (!techId || !date) return { error: 'Lipsesc tech_id sau date', positions: [] };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('locatii_istoric');
  if (!sheet) return { positions: [], tech_name: '' };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { positions: [], tech_name: '' };

  // LOCATIONS_HIST_COLS: [timestamp, technician_id, name, lat, lng, accuracy]
  const positions = [];
  let techName = '';
  for (let i = 1; i < data.length; i++) {
    const [ts, tid, tname, lat, lng, acc] = data[i];
    if (String(tid) !== String(techId)) continue;
    const tsStr = String(ts);
    if (!tsStr.startsWith(date)) continue;    // filtru pe dată
    if (!techName && tname) techName = String(tname);
    positions.push({
      timestamp: tsStr,
      lat:  parseFloat(lat),
      lng:  parseFloat(lng),
      accuracy: parseFloat(acc) || 0,
      time: new Date(ts).toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' })
    });
  }

  // Sortare cronologică
  positions.sort((a, b) => (a.timestamp < b.timestamp ? -1 : 1));
  return { positions, tech_name: techName };
}

// ── Helper: curățare istoric GPS mai vechi de 30 zile ─────────
function _cleanupGpsHistory(ss) {
  try {
    const props    = PropertiesService.getScriptProperties();
    const today    = new Date().toISOString().slice(0, 10);
    if (props.getProperty('last_gps_cleanup') === today) return; // o dată pe zi

    const sheet  = ss.getSheetByName('locatii_istoric');
    if (!sheet) return;

    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 60);
    const cutoffStr = cutoff.toISOString().slice(0, 10); // YYYY-MM-DD

    const data = sheet.getDataRange().getValues();
    // Parcurgem de jos în sus ca să nu deranjăm indicii la ștergere
    for (let i = data.length - 1; i >= 1; i--) {
      const ts = String(data[i][0]);
      if (ts && ts.slice(0, 10) < cutoffStr) sheet.deleteRow(i + 1);
    }
    props.setProperty('last_gps_cleanup', today);
  } catch (e) {
    Logger.log('Cleanup GPS history error: ' + e.message);
  }
}

// ── GET: getCalendar ──────────────────────────────────────────
// Returnează intrările din programul intervențiilor pentru intervalul dat.
function handleGetCalendar(params) {
  const dateFrom = params.date_from || '';
  const dateTo   = params.date_to   || '';
  const techId   = params.tech_id   || '';

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('program');
  if (!sheet) return { entries: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { entries: [] };

  // PROGRAM_COLS: [id, date, time, technician_id, technician_name, client_name, address, notes]
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i].map(v => String(v || '').trim());
    const [id, date, time, tid, tname, cname, addr, notes] = row;
    if (!id) continue;
    if (dateFrom && date < dateFrom) continue;
    if (dateTo   && date > dateTo)   continue;
    if (techId   && tid !== techId)  continue;
    entries.push({ id, date, time, technician_id: tid, technician_name: tname, client_name: cname, address: addr, notes });
  }

  // Sortare cronologică (dată + oră)
  entries.sort((a, b) => (a.date + a.time < b.date + b.time ? -1 : 1));
  return { entries };
}

// ── POST: saveCalendarEntries ─────────────────────────────────
// Adaugă sau suprascrie intrări în programul intervențiilor.
function handleSaveCalendarEntries(body) {
  const { entries } = body;
  if (!entries || !Array.isArray(entries)) return { success: false, error: 'Lipsesc datele' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'program', PROGRAM_COLS);
  const data  = sheet.getDataRange().getValues();

  // Construiește index id → număr rând (1-based)
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (id) existing[id] = i + 1;
  }

  let saved = 0;
  entries.forEach(entry => {
    const row = PROGRAM_COLS.map(col => entry[col] !== undefined ? String(entry[col]) : '');
    const rowNum = existing[entry.id];
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, PROGRAM_COLS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    saved++;
  });

  return { success: true, saved };
}

// ── POST: deleteCalendarEntry ─────────────────────────────────
// Șterge o intrare din programul intervențiilor după ID.
function handleDeleteCalendarEntry(body) {
  const { id } = body;
  if (!id) return { success: false, error: 'Lipsește ID-ul' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('program');
  if (!sheet) return { success: false, error: 'Sheet program inexistent' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Intrare negăsită' };
}

// ── GET: getLocations ─────────────────────────────────────────
// Returnează ultima poziție a fiecărui tehnician (rânduri din sheet locatii).
function handleGetLocations() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('locatii');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => ({
    technician_id: String(row[0]),
    name:          String(row[1]),
    lat:           parseFloat(row[2]),
    lng:           parseFloat(row[3]),
    accuracy:      parseInt(row[4]) || 0,
    timestamp:     String(row[5])
  })).filter(r => r.lat && r.lng); // exclude rânduri goale
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

// ── Audio Call Signaling ──────────────────────────────────────
// Folosit de Agora.io pentru a coordona cine ascultă pe cine.
// Foaia audio_calls: un rând per tech_id (suprascris la fiecare cerere).

/** Admin creează o cerere de ascultare pentru un tehnician. */
function handleRequestAudioCall(body) {
  const { tech_id, admin_id, admin_name, channel } = body;
  if (!tech_id || !channel) return { success: false, error: 'Date incomplete' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, 'audio_calls', AUDIO_CALLS_COLS);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString();

  // Caută rândul existent pentru acest tech_id
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(tech_id)) {
      sheet.getRange(i + 1, 1, 1, AUDIO_CALLS_COLS.length)
           .setValues([[tech_id, admin_id || '', admin_name || '', channel, 'pending', now]]);
      return { success: true };
    }
  }
  sheet.appendRow([tech_id, admin_id || '', admin_name || '', channel, 'pending', now]);
  return { success: true };
}

/** Tehnician verifică dacă există o cerere de ascultare pentru el. */
function handleGetAudioCall(params) {
  const tech_id = params.tech_id || '';
  if (!tech_id) return { status: 'none' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('audio_calls');
  if (!sheet) return { status: 'none' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(tech_id)) {
      const status = String(data[i][4] || 'none');
      // Ignoră cereri mai vechi de 2 minute (evită "cereri fantomă")
      const updatedAt = new Date(String(data[i][5] || 0)).getTime();
      if (Date.now() - updatedAt > 2 * 60 * 1000 && status === 'pending') {
        return { status: 'none' };
      }
      return {
        status,
        admin_id:   String(data[i][1]),
        admin_name: String(data[i][2]),
        channel:    String(data[i][3])
      };
    }
  }
  return { status: 'none' };
}

/** Actualizează statusul unui apel (accepted / ended). */
function handleUpdateAudioCall(body) {
  const { channel, status } = body;
  if (!channel) return { success: false };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('audio_calls');
  if (!sheet) return { success: false };

  const data = sheet.getDataRange().getValues();
  const now  = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(channel)) {
      sheet.getRange(i + 1, 5, 1, 2).setValues([[status, now]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Canal negăsit' };
}

// ── POST: OwnTracks HTTP mode ─────────────────────────────────
// OwnTracks trimite format propriu: { "_type":"location", "lat":45.9, "lon":24.9,
// "acc":15, "tst":1709123456, "tid":"DAN" }
// tid = tracker ID configurat în OwnTracks = username-ul tehnicianului (ex: "dan", "admin")
// Răspuns OwnTracks: { result:[] } (obligatoriu — altfel reîncercă)
/** Save exported file to Google Drive folder "Export Interventii" */
function handleSaveExportToDrive(body) {
  var fileName = body.fileName || 'export.xlsx';
  var base64Data = body.data || '';
  var mimeType = body.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  if (!base64Data) return { error: 'No file data provided' };

  // Find or create "Export Interventii" folder
  var folderName = 'Export Interventii';
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

function handleOwnTracksLocation(body) {
  const { _type, lat, lon, acc, tst, tid } = body;
  if (_type !== 'location' || !lat || !lon) return { result: [] };

  const timestamp = tst ? new Date(tst * 1000).toISOString() : new Date().toISOString();
  const trackerId = (tid || '').toLowerCase();

  // Caută tehnicianul după username (tid = username configurat în OwnTracks)
  const techs  = sheetToObjects('technicians', TECHNICIANS_COLS);
  const tech   = techs.find(t =>
    t.username.toLowerCase() === trackerId ||
    t.name.toLowerCase().replace(/\s/g, '') === trackerId
  );
  const techId   = tech ? tech.technician_id : 'ot_' + trackerId;
  const techName = tech ? tech.name : (tid || 'OwnTracks_' + trackerId);

  handleSaveLocation({
    technician_id: techId,
    name:          techName,
    lat:           lat,
    lng:           lon,    // OwnTracks folosește "lon", nu "lng"
    accuracy:      acc || 0,
    timestamp:     timestamp
  });

  return { result: [] }; // Răspuns obligatoriu pentru OwnTracks
}

// ── Sheet helpers ─────────────────────────────────────────────
function sheetToObjects(sheetName, cols) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(String);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
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
    { name: 'locatii',          cols: LOCATIONS_COLS,       color: '#7c3aed' }, // violet — GPS current
    { name: 'locatii_istoric',  cols: LOCATIONS_HIST_COLS,  color: '#a855f7' }, // violet clar — GPS history
    { name: 'audio_calls',      cols: AUDIO_CALLS_COLS,     color: '#0f766e' }, // teal — audio
    { name: 'program',          cols: PROGRAM_COLS,          color: '#0369a1' },  // albastru — calendar interventii
    { name: 'checklist',        cols: CHECKLIST_COLS,        color: '#b45309' }   // amber - evidenta checklist
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
