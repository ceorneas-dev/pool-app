// db.js — IndexedDB wrapper for Pool Manager PWA
// 7 stores: clients, interventions, treatment_rules, technicians, sync_queue, settings, stock

'use strict';

const DB_NAME    = 'poolmgmt';
const DB_VERSION = 3;  // v3: adds stock store + visit_frequency_days on clients

let _db = null;

// ── Open / Init ──────────────────────────────────────────────
function openDB() {
  if (_db) return Promise.resolve(_db);

  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);

    req.onupgradeneeded = e => {
      const db         = e.target.result;
      const oldVersion = e.oldVersion;

      // Drop all existing stores on major upgrades (ensures clean schema)
      if (oldVersion > 0 && oldVersion < 3) {
        Array.from(db.objectStoreNames).forEach(name => db.deleteObjectStore(name));
      }

      // ── Create stores ──────────────────────────────────────
      const cs = db.createObjectStore('clients', { keyPath: 'client_id' });
      cs.createIndex('name',   'name',   { unique: false });
      cs.createIndex('active', 'active', { unique: false });

      const is = db.createObjectStore('interventions', { keyPath: 'intervention_id' });
      is.createIndex('client_id',     'client_id',     { unique: false });
      is.createIndex('technician_id', 'technician_id', { unique: false });
      is.createIndex('date',          'date',          { unique: false });
      is.createIndex('synced',        'synced',        { unique: false });

      db.createObjectStore('treatment_rules', { keyPath: 'rule_id', autoIncrement: true });

      const ts = db.createObjectStore('technicians', { keyPath: 'technician_id' });
      ts.createIndex('username', 'username', { unique: true });

      const sq = db.createObjectStore('sync_queue', { keyPath: 'queue_id', autoIncrement: true });
      sq.createIndex('status',     'status',     { unique: false });
      sq.createIndex('created_at', 'created_at', { unique: false });

      db.createObjectStore('settings', { keyPath: 'key' });

      // NEW in v3: chemical stock inventory
      db.createObjectStore('stock', { keyPath: 'product_id' });
    };

    req.onsuccess = e => {
      _db = e.target.result;
      resolve(_db);
    };

    req.onerror = () => reject(req.error);
  });
}

// ── Generic CRUD ─────────────────────────────────────────────
function getAll(storeName) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  }));
}

function getByKey(storeName, key) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  }));
}

function getByIndex(storeName, indexName, value) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx    = db.transaction(storeName, 'readonly');
    const store = tx.objectStore(storeName);
    const idx   = store.index(indexName);
    const req   = idx.getAll(value);
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  }));
}

function getByIndexFirst(storeName, indexName, value) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).index(indexName).get(value);
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  }));
}

function put(storeName, data) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readwrite');
    const req = tx.objectStore(storeName).put(data);
    tx.oncomplete = () => resolve(req.result);
    tx.onerror    = () => reject(tx.error);
    tx.onabort    = () => reject(tx.error || new Error('Transaction aborted'));
    req.onerror   = () => reject(req.error);
  }));
}

function putMany(storeName, items) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx    = db.transaction(storeName, 'readwrite');
    const store = tx.objectStore(storeName);
    let count = 0;
    for (const item of items) {
      store.put(item);
      count++;
    }
    tx.oncomplete = () => resolve(count);
    tx.onerror    = () => reject(tx.error);
  }));
}

function clearStore(storeName) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).clear();
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
    tx.onabort    = () => reject(tx.error || new Error('Transaction aborted'));
  }));
}

function count(storeName) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).count();
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  }));
}

function deleteRecord(storeName, key) {
  return openDB().then(db => new Promise((resolve, reject) => {
    const tx  = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
    tx.onabort    = () => reject(tx.error || new Error('Transaction aborted'));
  }));
}

// ── Domain Helpers ───────────────────────────────────────────
function getActiveClients() {
  return getByIndex('clients', 'active', true).catch(() =>
    // Fallback if 'active' index missing (old DB schema)
    getAll('clients').then(all => all.filter(c => c.active !== false))
  ).then(clients =>
    // Extra safety: filter out entries without valid client_id
    clients.filter(c => c.client_id && String(c.client_id).trim() !== '')
  );
}

function getClientInterventions(clientId) {
  return getByIndex('interventions', 'client_id', clientId);
}

function getUnsyncedInterventions() {
  // Note: boolean false is not a valid IDB key, so we filter in JS
  return getAll('interventions').then(list => list.filter(i => !i.synced));
}

function saveIntervention(intervention) {
  return put('interventions', intervention);
}

function markSynced(interventionId) {
  return getByKey('interventions', interventionId).then(rec => {
    if (!rec) return;
    rec.synced = true;
    return put('interventions', rec);
  });
}

// ── Settings ────────────────────────────────────────────────
function getSetting(key) {
  return getByKey('settings', key).then(rec => rec ? rec.value : null);
}

function setSetting(key, value) {
  return put('settings', { key, value });
}

// ── Session ─────────────────────────────────────────────────
function getSession() {
  return getSetting('session');
}

function setSession(user) {
  return setSetting('session', user);
}

function clearSession() {
  return deleteRecord('settings', 'session');
}

// ── Counts ───────────────────────────────────────────────────
function getPendingSyncCount() {
  return getUnsyncedInterventions().then(list => list.length);
}

// ── Stock Helpers (v3) ───────────────────────────────────────
function getAllStock() {
  return getAll('stock');
}

function updateStockProduct(product) {
  return put('stock', product);
}

// ── Seed Demo Data ───────────────────────────────────────────
function seedDemoData() {
  const now   = new Date().toISOString();

  const technicians = [
    {
      technician_id: 't_001',
      name: 'Administrator',
      username: 'admin',
      password: 'admin123',
      role: 'admin',
      active: true,
      last_sync: null
    },
    {
      technician_id: 't_002',
      name: 'Dan Popescu',
      username: 'dan',
      password: 'dan123',
      role: 'technician',
      active: true,
      last_sync: null
    }
  ];

  const clients = [
    {
      client_id: 'c_001',
      name: 'Andrei Ionescu',
      phone: '0722 111 222',
      address: 'Str. Florilor 12, Sector 3, București',
      pool_volume_mc: 50,
      pool_type: 'exterior',
      active: true,
      visit_frequency_days: 7,
      notes: 'Piscină cu saltwater system. Acces prin poartă laterală.',
      created_at: now, updated_at: now,
      latitude: 44.4268, longitude: 26.1025, location_set: true
    },
    {
      client_id: 'c_002',
      name: 'Maria Dumitrescu',
      phone: '0744 333 444',
      address: 'Bd. Unirii 45, Sector 4, București',
      pool_volume_mc: 80,
      pool_type: 'exterior',
      active: true,
      visit_frequency_days: 14,
      notes: 'Piscină mare, pompă Hayward. Client de 3 ani.',
      created_at: now, updated_at: now,
      latitude: 44.4194, longitude: 26.1057, location_set: true
    },
    {
      client_id: 'c_003',
      name: 'Vasile Gheorghe',
      phone: '0733 555 666',
      address: 'Str. Mihai Eminescu 7, Otopeni',
      pool_volume_mc: 35,
      pool_type: 'interior',
      active: true,
      visit_frequency_days: 14,
      notes: 'Piscină interioară. Atenție la ventilație.',
      created_at: now, updated_at: now,
      latitude: 44.5456, longitude: 26.0768, location_set: true
    },
    {
      client_id: 'c_004',
      name: 'Elena Popa',
      phone: '0755 777 888',
      address: 'Aleea Trandafirilor 3, Voluntari',
      pool_volume_mc: 120,
      pool_type: 'exterior',
      active: true,
      visit_frequency_days: 10,
      notes: 'Vilă cu 2 piscine. Aceasta este cea principală (120 m³).',
      created_at: now, updated_at: now,
      latitude: null, longitude: null, location_set: false
    }
  ];

  // Past dates helper
  const d = daysAgo => {
    const dt = new Date();
    dt.setDate(dt.getDate() - daysAgo);
    return dt.toISOString().split('T')[0];
  };

  const interventions = [
    {
      intervention_id: 'i_demo_001',
      client_id: 'c_001', client_name: 'Andrei Ionescu',
      technician_id: 't_002', technician_name: 'Dan Popescu',
      date: d(5), created_at: now,
      measured_chlorine: 0.8, measured_ph: 7.4, measured_temp: 28,
      measured_hardness: 250, measured_alkalinity: 110, measured_salinity: null,
      rec_cl_gr: 600, rec_cl_tab: 3, rec_ph_kg: 0.6, rec_anti_l: 0.75,
      treat_cl_granule_gr: 600, treat_cl_tablete: 3, treat_cl_tablete_export_gr: 750,
      treat_cl_lichid_bidoane: 0, treat_ph_granule: 0.6, treat_ph_lichid_bidoane: 0,
      treat_antialgic: 0.75, treat_anticalcar: 0, treat_floculant: 0,
      treat_sare_saci: 0, treat_bicarbonat: 0,
      observations: 'Apă limpede, filtrare OK. Adăugat antialgic preventiv.',
      photos: [], synced: false,
      geo_lat: 44.4268, geo_lng: 26.1025, geo_accuracy: 15,
      arrival_time: d(5) + 'T09:15:00.000Z',
      departure_time: d(5) + 'T10:05:00.000Z',
      duration_minutes: 50
    },
    {
      intervention_id: 'i_demo_002',
      client_id: 'c_002', client_name: 'Maria Dumitrescu',
      technician_id: 't_002', technician_name: 'Dan Popescu',
      date: d(8), created_at: now,
      measured_chlorine: 0.2, measured_ph: 7.8, measured_temp: 26,
      measured_hardness: 180, measured_alkalinity: 95, measured_salinity: null,
      rec_cl_gr: 1200, rec_cl_tab: 5, rec_ph_kg: 1.6, rec_anti_l: 1.0,
      treat_cl_granule_gr: 1200, treat_cl_tablete: 5, treat_cl_tablete_export_gr: 1250,
      treat_cl_lichid_bidoane: 0, treat_ph_granule: 1.6, treat_ph_lichid_bidoane: 0,
      treat_antialgic: 1.0, treat_anticalcar: 0.5, treat_floculant: 0,
      treat_sare_saci: 0, treat_bicarbonat: 0,
      observations: 'Clor scăzut după weekendul ploios. pH ridicat — dozare pH-.',
      photos: [], synced: false,
      geo_lat: 44.4194, geo_lng: 26.1057, geo_accuracy: 20,
      arrival_time: d(8) + 'T14:00:00.000Z',
      departure_time: d(8) + 'T15:10:00.000Z',
      duration_minutes: 70
    },
    {
      intervention_id: 'i_demo_003',
      client_id: 'c_003', client_name: 'Vasile Gheorghe',
      technician_id: 't_001', technician_name: 'Administrator',
      date: d(15), created_at: now,
      measured_chlorine: 1.5, measured_ph: 7.2, measured_temp: 30,
      measured_hardness: 300, measured_alkalinity: 130, measured_salinity: null,
      rec_cl_gr: 0, rec_cl_tab: 0, rec_ph_kg: 0, rec_anti_l: 0.5,
      treat_cl_granule_gr: 0, treat_cl_tablete: 0, treat_cl_tablete_export_gr: 0,
      treat_cl_lichid_bidoane: 0, treat_ph_granule: 0, treat_ph_lichid_bidoane: 0,
      treat_antialgic: 0.5, treat_anticalcar: 1.0, treat_floculant: 0,
      treat_sare_saci: 0, treat_bicarbonat: 0,
      observations: 'Parametri în normă. Anticalcar pentru depuneri pe pereți.',
      photos: [], synced: true,
      geo_lat: 44.5456, geo_lng: 26.0768, geo_accuracy: 12,
      arrival_time: d(15) + 'T10:30:00.000Z',
      departure_time: d(15) + 'T11:15:00.000Z',
      duration_minutes: 45
    },
    {
      intervention_id: 'i_demo_004',
      client_id: 'c_001', client_name: 'Andrei Ionescu',
      technician_id: 't_002', technician_name: 'Dan Popescu',
      date: d(35), created_at: now,
      measured_chlorine: 0.5, measured_ph: 7.5, measured_temp: 25,
      measured_hardness: 230, measured_alkalinity: 100, measured_salinity: null,
      rec_cl_gr: 600, rec_cl_tab: 3, rec_ph_kg: 0.4, rec_anti_l: 0.75,
      treat_cl_granule_gr: 600, treat_cl_tablete: 3, treat_cl_tablete_export_gr: 750,
      treat_cl_lichid_bidoane: 0, treat_ph_granule: 0.4, treat_ph_lichid_bidoane: 0,
      treat_antialgic: 0.75, treat_anticalcar: 0, treat_floculant: 0.5,
      treat_sare_saci: 0, treat_bicarbonat: 0,
      observations: 'Apă ușor tulbure — floculant adăugat.',
      photos: [], synced: true,
      geo_lat: 44.4268, geo_lng: 26.1025, geo_accuracy: 18,
      arrival_time: d(35) + 'T09:00:00.000Z',
      departure_time: d(35) + 'T10:00:00.000Z',
      duration_minutes: 60
    },
    {
      intervention_id: 'i_demo_005',
      client_id: 'c_004', client_name: 'Elena Popa',
      technician_id: 't_002', technician_name: 'Dan Popescu',
      date: d(3), created_at: now,
      measured_chlorine: 0.4, measured_ph: 7.7, measured_temp: 27,
      measured_hardness: 200, measured_alkalinity: 108, measured_salinity: null,
      rec_cl_gr: 1300, rec_cl_tab: 5, rec_ph_kg: 2.5, rec_anti_l: 2.0,
      treat_cl_granule_gr: 1300, treat_cl_tablete: 5, treat_cl_tablete_export_gr: 1250,
      treat_cl_lichid_bidoane: 0, treat_ph_granule: 2.5, treat_ph_lichid_bidoane: 0,
      treat_antialgic: 2.0, treat_anticalcar: 0, treat_floculant: 0,
      treat_sare_saci: 0, treat_bicarbonat: 0,
      observations: 'Prima vizită. Documentat situația inițială.',
      photos: [], synced: false,
      geo_lat: null, geo_lng: null, geo_accuracy: null,
      arrival_time: d(3) + 'T11:00:00.000Z',
      departure_time: d(3) + 'T12:20:00.000Z',
      duration_minutes: 80
    }
  ];

  // Default stock products (v3+: includes step and visible fields)
  const stock = [
    { product_id: 'cl_granule',  name: 'Cl Granule',      unit: 'gr',      quantity: 5000,  alert_threshold: 500,  step: 50,   visible: true },
    { product_id: 'cl_tablete',  name: 'Cl Tablete',      unit: 'buc',     quantity: 50,    alert_threshold: 5,    step: 1,    visible: true },
    { product_id: 'cl_lichid',   name: 'Cl Lichid',       unit: 'bidoane', quantity: 20,    alert_threshold: 2,    step: 1,    visible: true },
    { product_id: 'ph_minus_gr', name: 'pH Granule (−)',  unit: 'kg',      quantity: 3,     alert_threshold: 0.3,  step: 0.1,  visible: true },
    { product_id: 'ph_minus_l',  name: 'pH Lichid (−)',   unit: 'bidoane', quantity: 10,    alert_threshold: 1,    step: 1,    visible: true },
    { product_id: 'antialgic',   name: 'Antialgic',       unit: 'L',       quantity: 15,    alert_threshold: 1,    step: 0.25, visible: true },
    { product_id: 'anticalcar',  name: 'Anticalcar',      unit: 'L',       quantity: 8,     alert_threshold: 1,    step: 0.25, visible: true },
    { product_id: 'floculant',   name: 'Floculant',       unit: 'L',       quantity: 5,     alert_threshold: 0.5,  step: 0.25, visible: true },
    { product_id: 'sare',        name: 'Sare piscină',    unit: 'saci',    quantity: 50,    alert_threshold: 5,    step: 1,    visible: true },
    { product_id: 'bicarbonat',  name: 'Bicarbonat sodiu',unit: 'kg',      quantity: 10,    alert_threshold: 1,    step: 0.5,  visible: true }
  ];

  return Promise.all([
    putMany('technicians',  technicians),
    putMany('clients',      clients),
    putMany('interventions', interventions),
    putMany('stock',        stock)
  ]);
}
