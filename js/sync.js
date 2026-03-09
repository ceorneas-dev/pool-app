// sync.js — API sync layer (singurul fișier care se schimbă la migrare backend)
// Configurare: setează SYNC_CONFIG.API_URL cu URL-ul Google Apps Script

'use strict';

const SYNC_CONFIG = {
  API_URL:          '',       // completat de user din Settings
  SYNC_INTERVAL_MS: 900000   // 15 minute
};

let _syncTimer  = null;
let _syncActive = false;

// ── Config helpers ───────────────────────────────────────────
function isSyncConfigured() {
  return SYNC_CONFIG.API_URL && SYNC_CONFIG.API_URL.trim().length > 0;
}

function loadSyncConfig() {
  return getSetting('api_url').then(url => {
    if (url) SYNC_CONFIG.API_URL = url;
  });
}

// ── Init / Start / Stop ──────────────────────────────────────
function initSync() {
  return loadSyncConfig().then(() => {
    if (!isSyncConfigured()) return;
    stopSync();
    doSync();
    _syncTimer = setInterval(doSync, SYNC_CONFIG.SYNC_INTERVAL_MS);
    console.log('[SYNC] Sync initialized, interval:', SYNC_CONFIG.SYNC_INTERVAL_MS / 60000, 'min');
  });
}

function stopSync() {
  if (_syncTimer) {
    clearInterval(_syncTimer);
    _syncTimer = null;
  }
}

function forceSync() {
  return doSync();
}

// ── Main sync cycle ──────────────────────────────────────────
function doSync() {
  if (_syncActive) {
    console.log('[SYNC] Sync already running, skipping');
    return Promise.resolve();
  }
  if (!isSyncConfigured()) {
    console.log('[SYNC] API not configured, skipping');
    return Promise.resolve();
  }
  if (!navigator.onLine) {
    console.log('[SYNC] Offline, skipping');
    return Promise.resolve();
  }

  _syncActive = true;
  console.log('[SYNC] Starting sync cycle...');

  return pushInterventions()
    .then(() => pullData())
    .then(() => {
      setSetting('last_sync', new Date().toISOString());
      console.log('[SYNC] Sync cycle complete');
      if (typeof window.onSyncComplete === 'function') window.onSyncComplete();
    })
    .catch(err => {
      console.error('[SYNC] Sync error:', err.message);
      if (typeof window.onSyncError === 'function') window.onSyncError(err);
    })
    .finally(() => {
      _syncActive = false;
    });
}

// ── Push interventions to server ─────────────────────────────
function pushInterventions() {
  return getUnsyncedInterventions().then(list => {
    if (!list.length) {
      console.log('[SYNC] No unsynced interventions to push');
      return;
    }
    console.log('[SYNC] Pushing', list.length, 'interventions');

    const session = null; // loaded below
    return getSession().then(user => {
      const payload = {
        action: 'push',
        type:   'interventions',
        tech_id: user ? user.technician_id : '',
        data: list.map(i => ({
          intervention_id:     i.intervention_id,
          client_id:           i.client_id,
          client_name:         i.client_name,
          technician_id:       i.technician_id,
          technician_name:     i.technician_name,
          date:                i.date,
          created_at:          i.created_at,
          measured_chlorine:   i.measured_chlorine,
          measured_ph:         i.measured_ph,
          measured_temp:       i.measured_temp,
          measured_hardness:   i.measured_hardness,
          measured_alkalinity: i.measured_alkalinity,
          measured_salinity:   i.measured_salinity,
          rec_cl_gr:           i.rec_cl_gr,
          rec_cl_tab:          i.rec_cl_tab,
          rec_ph_kg:           i.rec_ph_kg,
          rec_anti_l:          i.rec_anti_l,
          treat_cl_granule_gr:      i.treat_cl_granule_gr,
          treat_cl_tablete:         i.treat_cl_tablete,
          treat_cl_tablete_export_gr: i.treat_cl_tablete_export_gr,
          treat_cl_lichid_bidoane:  i.treat_cl_lichid_bidoane,
          treat_ph_granule:         i.treat_ph_granule,
          treat_ph_lichid_bidoane:  i.treat_ph_lichid_bidoane,
          treat_antialgic:          i.treat_antialgic,
          treat_anticalcar:         i.treat_anticalcar,
          treat_floculant:          i.treat_floculant,
          treat_sare_saci:          i.treat_sare_saci,
          treat_bicarbonat:         i.treat_bicarbonat,
          observations:        i.observations,
          // GPS + time fields
          geo_lat:             i.geo_lat,
          geo_lng:             i.geo_lng,
          geo_accuracy:        i.geo_accuracy,
          arrival_time:        i.arrival_time,
          departure_time:      i.departure_time,
          duration_minutes:    i.duration_minutes
        }))
      };

      return apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body:   JSON.stringify(payload)
      }).then(data => {
        console.log('[SYNC] Push result:', data.pushed, 'pushed,', data.duplicates, 'duplicates');
        return Promise.all(list.map(i => markSynced(i.intervention_id)));
      });
    });
  });
}

// ── Pull data from server ────────────────────────────────────
function pullData() {
  console.log('[SYNC] Pulling data...');
  const url = SYNC_CONFIG.API_URL + '?action=pull&type=all';
  return apiFetch(url).then(data => {
    const tasks = [];

    if (data.clients && data.clients.length) {
      const parsed = data.clients.map(c => ({
        client_id:      c.client_id,
        name:           c.name,
        phone:          c.phone || '',
        address:        c.address || '',
        pool_volume_mc: parseFloat(c.pool_volume_mc) || 0,
        pool_type:      c.pool_type || 'exterior',
        active:         c.active === true || c.active === 'true',
        notes:          c.notes || '',
        created_at:     c.created_at || new Date().toISOString(),
        updated_at:     c.updated_at || new Date().toISOString(),
        latitude:       c.latitude  ? parseFloat(c.latitude)  : null,
        longitude:      c.longitude ? parseFloat(c.longitude) : null,
        location_set:   c.location_set === true || c.location_set === 'true'
      }));
      tasks.push(clearStore('clients').then(() => putMany('clients', parsed)));
      console.log('[SYNC] Pulled', parsed.length, 'clients');
    }

    if (data.technicians && data.technicians.length) {
      const parsed = data.technicians.map(t => ({
        technician_id: t.technician_id,
        name:          t.name,
        username:      t.username,
        password:      t.password,
        role:          t.role || 'technician',
        active:        t.active === true || t.active === 'true',
        last_sync:     t.last_sync || null
      }));
      tasks.push(clearStore('technicians').then(() => putMany('technicians', parsed)));
      console.log('[SYNC] Pulled', parsed.length, 'technicians');
    }

    if (data.treatment_rules && data.treatment_rules.length) {
      const parsed = data.treatment_rules.map(r => ({
        rule_id:      parseInt(r.rule_id),
        pool_vol_min: parseFloat(r.pool_vol_min),
        pool_vol_max: parseFloat(r.pool_vol_max),
        cl_min:       parseFloat(r.cl_min),
        cl_max:       parseFloat(r.cl_max),
        ph_min:       parseFloat(r.ph_min),
        ph_max:       parseFloat(r.ph_max),
        rec_cl_gr:    parseFloat(r.rec_cl_gr),
        rec_cl_tab:   parseInt(r.rec_cl_tab),
        rec_ph_kg:    parseFloat(r.rec_ph_kg),
        rec_anti:     parseFloat(r.rec_anti)
      }));
      tasks.push(clearStore('treatment_rules').then(() => putMany('treatment_rules', parsed)));
      console.log('[SYNC] Pulled', parsed.length, 'treatment rules');
    }

    return Promise.all(tasks);
  });
}

// ── Login via API ────────────────────────────────────────────
function apiLogin(username, password) {
  if (!isSyncConfigured()) return Promise.reject(new Error('API not configured'));
  return apiFetch(SYNC_CONFIG.API_URL, {
    method: 'POST',
    body: JSON.stringify({ action: 'login', username, password })
  }).then(data => {
    if (!data.success) throw new Error(data.error || 'Login failed');
    return data.user;
  });
}

// ── Sync state helpers ───────────────────────────────────────
function getLastSyncTime() {
  return getSetting('last_sync');
}

// ── Internal fetch wrapper ───────────────────────────────────
function apiFetch(url, options) {
  const opts = Object.assign({
    method:  'GET',
    headers: { 'Content-Type': 'application/json' },
    mode:    'cors'
  }, options);

  return fetch(url, opts)
    .then(res => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    });
}
