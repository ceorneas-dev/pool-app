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

  return pushTechnicians()
    .then(() => pushInterventions())
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

// Push all local technicians to server on each sync cycle
function pushTechnicians() {
  return getAll('technicians').then(function(techs) {
    if (!techs || !techs.length) return;
    return apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'technicians', data: techs })
    }).then(function() {
      console.log('[SYNC] Pushed', techs.length, 'technicians');
    }).catch(function(err) {
      console.warn('[SYNC] Technician push failed:', err.message);
    });
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
        phone:          (function(p){ var s = String(p || ''); if (/^\d{9}$/.test(s) && s[0] === '7') s = '0' + s; return s; })(c.phone),
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
      // Merge: upsert remote technicians one-by-one (avoids unique-index abort)
      const techMerge = (async function() {
        let ok = 0;
        for (const t of parsed) {
          try { await put('technicians', t); ok++; } catch(e) {
            console.warn('[SYNC] Tech put failed for', t.username, ':', e.message);
          }
        }
        console.log('[SYNC] Pulled', ok, '/', parsed.length, 'technicians (merged)');
        // Update backup
        try {
          const all = await getAll('technicians');
          await setSetting('technicians_backup', JSON.stringify(all));
        } catch(_) {}
      })();
      tasks.push(techMerge);
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

    // Pull interventions from server and merge with local
    if (data.interventions && data.interventions.length) {
      const mergeInterventions = (async function() {
        const localAll = await getAll('interventions');
        const localMap = {};
        localAll.forEach(function(i) { localMap[i.intervention_id] = i; });

        let added = 0, updated = 0;
        for (const ri of data.interventions) {
          const parsed = {
            intervention_id:     ri.intervention_id,
            client_id:           ri.client_id,
            client_name:         ri.client_name || '',
            technician_id:       ri.technician_id,
            technician_name:     ri.technician_name || '',
            date:                ri.date,
            created_at:          ri.created_at || '',
            measured_chlorine:   ri.measured_chlorine !== '' ? parseFloat(ri.measured_chlorine) || null : null,
            measured_ph:         ri.measured_ph !== '' ? parseFloat(ri.measured_ph) || null : null,
            measured_temp:       ri.measured_temp !== '' ? parseFloat(ri.measured_temp) || null : null,
            measured_hardness:   ri.measured_hardness !== '' ? parseFloat(ri.measured_hardness) || null : null,
            measured_alkalinity: ri.measured_alkalinity !== '' ? parseFloat(ri.measured_alkalinity) || null : null,
            measured_salinity:   ri.measured_salinity !== '' ? parseFloat(ri.measured_salinity) || null : null,
            measured_tc:         ri.measured_tc !== '' ? parseFloat(ri.measured_tc) || null : null,
            measured_cya:        ri.measured_cya !== '' ? parseFloat(ri.measured_cya) || null : null,
            rec_cl_gr:           ri.rec_cl_gr !== '' ? parseFloat(ri.rec_cl_gr) || null : null,
            rec_cl_tab:          ri.rec_cl_tab !== '' ? parseFloat(ri.rec_cl_tab) || null : null,
            rec_ph_kg:           ri.rec_ph_kg !== '' ? parseFloat(ri.rec_ph_kg) || null : null,
            rec_anti_l:          ri.rec_anti_l !== '' ? parseFloat(ri.rec_anti_l) || null : null,
            treat_cl_granule_gr:      parseFloat(ri.treat_cl_granule_gr) || 0,
            treat_cl_tablete:         parseFloat(ri.treat_cl_tablete) || 0,
            treat_cl_tablete_export_gr: parseFloat(ri.treat_cl_tablete_export_gr) || 0,
            treat_cl_lichid_bidoane:  parseFloat(ri.treat_cl_lichid_bidoane) || 0,
            treat_ph_granule:         parseFloat(ri.treat_ph_granule) || 0,
            treat_ph_lichid_bidoane:  parseFloat(ri.treat_ph_lichid_bidoane) || 0,
            treat_antialgic:          parseFloat(ri.treat_antialgic) || 0,
            treat_anticalcar:         parseFloat(ri.treat_anticalcar) || 0,
            treat_floculant:          parseFloat(ri.treat_floculant) || 0,
            treat_sare_saci:          parseFloat(ri.treat_sare_saci) || 0,
            treat_bicarbonat:         parseFloat(ri.treat_bicarbonat) || 0,
            observations:        ri.observations || '',
            geo_lat:             ri.geo_lat !== '' ? parseFloat(ri.geo_lat) || null : null,
            geo_lng:             ri.geo_lng !== '' ? parseFloat(ri.geo_lng) || null : null,
            geo_accuracy:        ri.geo_accuracy !== '' ? parseFloat(ri.geo_accuracy) || null : null,
            arrival_time:        ri.arrival_time || null,
            departure_time:      ri.departure_time || null,
            duration_minutes:    ri.duration_minutes !== '' ? parseFloat(ri.duration_minutes) || null : null,
            synced:              true
          };

          const local = localMap[parsed.intervention_id];
          if (!local) {
            // New from server
            try { await put('interventions', parsed); added++; } catch(e) {
              console.warn('[SYNC] Intervention put failed:', parsed.intervention_id, e.message);
            }
          } else if (local.synced === false) {
            // Local has unsynced changes — keep local version
          } else {
            // Both synced — update with server version
            try { await put('interventions', parsed); updated++; } catch(e) {
              console.warn('[SYNC] Intervention update failed:', parsed.intervention_id, e.message);
            }
          }
        }
        console.log('[SYNC] Pulled interventions: ' + added + ' added, ' + updated + ' updated (server total: ' + data.interventions.length + ')');
      })();
      tasks.push(mergeInterventions);
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
    headers: { 'Content-Type': 'text/plain' },
    redirect: 'follow'
  }, options);

  return fetch(url, opts)
    .then(res => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    });
}
