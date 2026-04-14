// sync.js — API sync layer (singurul fișier care se schimbă la migrare backend)
// Configurare: setează SYNC_CONFIG.API_URL cu URL-ul Google Apps Script

'use strict';

const SYNC_CONFIG = {
  API_URL:          'https://script.google.com/macros/s/AKfycbwvKuTXZBl3613nSOId-3fuiOdt3SawvUth1ZZUjNHkyjZh5SuJg_jCckbYLYG2_ODYIg/exec',
  SYNC_INTERVAL_MS: 900000   // 15 minute
};

/** Parse numeric value preserving 0 (unlike parseFloat(x)||null which loses zero) */
function _parseNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = parseFloat(v);
  return isNaN(n) ? null : n;
}

/** Normalize date to YYYY-MM-DD format (handles Date objects, timestamps, various string formats) */
function _normalizeDate(val) {
  if (!val) return '';
  var s = String(val).trim();
  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Try to parse as Date
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    var y = d.getFullYear();
    var m = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);
    return y + '-' + m + '-' + day;
  }
  return s; // fallback: return as-is
}

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
    .then(() => pushClients())
    .then(() => pushDeletedInterventions())
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

// Push all local clients to server (syncs deviz_type, pret_interventie, etc.)
function pushClients() {
  return getAll('clients').then(function(clients) {
    if (!clients || !clients.length) return;
    return apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'clients', data: clients })
    }).then(function() {
      console.log('[SYNC] Pushed', clients.length, 'clients');
    }).catch(function(err) {
      console.warn('[SYNC] Client push failed:', err.message);
    });
  });
}

// Push all local technicians to server on each sync cycle (including pending deletions)
function pushTechnicians() {
  return Promise.all([getAll('technicians'), getSetting('deleted_technician_ids')]).then(function(results) {
    var techs = results[0] || [];
    var deletedIds = results[1] || [];
    // Build payload: active techs + deletion markers
    var payload = techs.slice();
    deletedIds.forEach(function(id) {
      payload.push({ technician_id: id, _deleted: true });
    });
    if (!payload.length) return;
    return apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'technicians', data: payload })
    }).then(function() {
      console.log('[SYNC] Pushed', techs.length, 'technicians +', deletedIds.length, 'deletions');
      // Clear pending flags after successful push — allow pull to update on next cycle
      setSetting('techs_pending_push', false);
      if (deletedIds.length) setSetting('deleted_technician_ids', []);
    }).catch(function(err) {
      console.warn('[SYNC] Technician push failed:', err.message);
    });
  });
}

// ── Push deleted intervention IDs to server (bulk) ──────────
function pushDeletedInterventions() {
  return getSetting('deleted_intervention_ids').then(function(ids) {
    if (!ids || !Array.isArray(ids) || !ids.length) return;
    console.log('[SYNC] Pushing', ids.length, 'deleted intervention IDs to server');
    return apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'delete_interventions_bulk', data: { ids: ids } })
    }).then(function(res) {
      console.log('[SYNC] Server deleted', res.deleted, 'interventions — clearing local deleted_intervention_ids');
      // Clear the list after successful server deletion to prevent re-pushing
      return setSetting('deleted_intervention_ids', []);
    }).catch(function(err) {
      console.warn('[SYNC] Bulk delete push failed:', err.message);
    });
  }).catch(function() { /* no deleted IDs */ });
}

// ── Push interventions to server ─────────────────────────────
function pushInterventions() {
  return getUnsyncedInterventions().then(list => {
    if (!list.length) {
      console.log('[SYNC] No unsynced interventions to push');
      return;
    }
    console.log('[SYNC] Pushing', list.length, 'interventions');

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
          measured_tc:         i.measured_tc,
          measured_cya:        i.measured_cya,
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
          operations:          Array.isArray(i.operations) ? JSON.stringify(i.operations) : (i.operations || ''),
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
        console.log('[SYNC] Push result:', data.pushed, 'pushed,', data.updated, 'updated,', data.deduped || 0, 'deduped');
        return Promise.all(list.map(i => markSynced(i.intervention_id)));
      }).catch(function(err) {
        console.warn('[SYNC] Intervention push failed:', err.message);
        // Don't throw — allow pullData to still run
      });
    });
  });
}

// ── Pull data from server ────────────────────────────────────
function pullData() {
  console.log('[SYNC] Pulling data...');
  const url = SYNC_CONFIG.API_URL + '?action=pull&type=all';
  return apiFetch(url).then(async data => {
    const tasks = [];
    console.log('[SYNC] Pull response: clients=' + (data.clients ? data.clients.length : 0) +
      ', interventions=' + (data.interventions ? data.interventions.length : 0) +
      ', technicians=' + (data.technicians ? data.technicians.length : 0));

    if (data.clients && data.clients.length) {
      // Read existing local clients to preserve local-only fields
      const localClients = await getAll('clients').catch(() => []);
      const localMap = {};
      localClients.forEach(function(lc) { localMap[lc.client_id] = lc; });

      const parsed = data.clients.map(c => {
        var local = localMap[c.client_id] || {};
        return {
          client_id:      c.client_id,
          name:           c.name,
          phone:          (function(p){ var s = String(p || ''); if (/^\d{9}$/.test(s) && s[0] === '7') s = '0' + s; return s; })(c.phone),
          address:        c.address || '',
          pool_volume_mc: parseFloat(c.pool_volume_mc) || 0,
          pool_type:      c.pool_type || 'exterior',
          active:         c.active === true || c.active === 'true' || c.active === 'TRUE' || c.active === '1' || c.active === 1,
          notes:          c.notes || '',
          created_at:     c.created_at || new Date().toISOString(),
          updated_at:     c.updated_at || new Date().toISOString(),
          latitude:       c.latitude  ? parseFloat(c.latitude)  : null,
          longitude:      c.longitude ? parseFloat(c.longitude) : null,
          location_set:   c.location_set === true || c.location_set === 'true',
          // Fields synced via GAS (prefer remote, fallback to local)
          deviz_type:     parseInt(c.deviz_type || local.deviz_type) || 2,
          pret_interventie: parseFloat(c.pret_interventie || local.pret_interventie) || 0,
          billing_interval_interventions: parseInt(c.billing_interval_interventions || local.billing_interval_interventions) || 4,
          visit_frequency_days: parseInt(c.visit_frequency_days || local.visit_frequency_days) || 7,
          last_billing_date: c.last_billing_date || local.last_billing_date || null
        };
      });
      // Robust client replacement: delete each old client individually, then add new ones
      const clientReplace = (async function() {
        // Step 1: delete ALL existing local clients one by one
        for (var ci = 0; ci < localClients.length; ci++) {
          try { await deleteRecord('clients', localClients[ci].client_id); } catch(e) {}
        }
        // Step 2: also try clearStore as backup
        try { await clearStore('clients'); } catch(e) {}
        // Step 3: add fresh clients from server
        var added = 0;
        for (var pi = 0; pi < parsed.length; pi++) {
          try { await put('clients', parsed[pi]); added++; } catch(e) {
            console.warn('[SYNC] Client put failed:', parsed[pi].client_id, e.message);
          }
        }
        console.log('[SYNC] Pulled', added, '/', parsed.length, 'clients (replaced ' + localClients.length + ' old)');
      })();
      tasks.push(clientReplace);
    }

    if (data.technicians && data.technicians.length) {
      // Technicians sync strategy:
      // - If there are pending local changes (techs_pending_push), skip pull to avoid overwrite
      // - Otherwise, pull from server (GAS is source of truth after push completes)
      const techMerge = (async function() {
        var pendingPush = await getSetting('techs_pending_push');
        var deletedTechIds = (await getSetting('deleted_technician_ids')) || [];
        if (pendingPush) {
          console.log('[SYNC] Technicians: pending local changes — skipping pull');
          return;
        }
        // Replace local technicians with server data
        console.log('[SYNC] Technicians: pulling from server...');
        const parsed = data.technicians.map(t => ({
          technician_id: t.technician_id,
          name:          t.name,
          username:      t.username,
          password:      t.password,
          role:          t.role || 'technician',
          active:        t.active === true || t.active === 'true',
          last_sync:     t.last_sync || null
        }));
        // Clear local techs and replace with server data
        try { await clearStore('technicians'); } catch(_) {}
        const usedUsernames = new Set();
        let ok = 0;
        for (const t of parsed) {
          // Skip locally-deleted technicians
          if (deletedTechIds.indexOf(t.technician_id) !== -1) continue;
          if (!t.username || !String(t.username).trim()) {
            t.username = 'user_' + t.technician_id;
          }
          if (usedUsernames.has(t.username)) {
            t.username = t.username + '_' + t.technician_id;
          }
          usedUsernames.add(t.username);
          try { await put('technicians', t); ok++; } catch(e) {
            console.warn('[SYNC] Tech put failed for', t.username, ':', e.message);
          }
        }
        console.log('[SYNC] Pulled', ok, '/', parsed.length, 'technicians from server');
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
    // Strategy: server is source of truth; local unsynced are preserved
    if (data.interventions) {
      const mergeInterventions = (async function() {
        const localAll = await getAll('interventions');
        const localMap = {};
        localAll.forEach(function(i) { localMap[i.intervention_id] = i; });

        // Load deleted IDs to skip re-adding
        var deletedIds = await getSetting('deleted_intervention_ids').catch(function() { return []; }) || [];
        if (!Array.isArray(deletedIds)) deletedIds = [];
        var deletedSet = {};
        deletedIds.forEach(function(id) { deletedSet[id] = true; });

        // Build server ID set (NO dedup — server is source of truth)
        var serverIdSet = {};
        data.interventions.forEach(function(ri) { serverIdSet[ri.intervention_id] = true; });

        // Step 1: Remove local synced interventions that server no longer has
        // (they were explicitly deleted on another device via pushDeletedInterventions)
        var removedFromLocal = 0;
        for (var li = 0; li < localAll.length; li++) {
          var localIntv = localAll[li];
          if (localIntv.synced !== false && !serverIdSet[localIntv.intervention_id]) {
            try { await deleteRecord('interventions', localIntv.intervention_id); removedFromLocal++; } catch(e) {}
          }
        }

        // Step 2: Parse and upsert ALL server interventions (no dedup — trust server)
        let added = 0, updated = 0;

        for (var si = 0; si < data.interventions.length; si++) {
          var ri = data.interventions[si];
          // Skip if locally deleted on THIS device
          if (deletedSet[ri.intervention_id]) continue;

          var parsed = {
            intervention_id:     ri.intervention_id,
            client_id:           ri.client_id,
            client_name:         ri.client_name || '',
            technician_id:       ri.technician_id,
            technician_name:     ri.technician_name || '',
            date:                _normalizeDate(ri.date),
            created_at:          ri.created_at || '',
            measured_chlorine:   _parseNum(ri.measured_chlorine),
            measured_ph:         _parseNum(ri.measured_ph),
            measured_temp:       _parseNum(ri.measured_temp),
            measured_hardness:   _parseNum(ri.measured_hardness),
            measured_alkalinity: _parseNum(ri.measured_alkalinity),
            measured_salinity:   _parseNum(ri.measured_salinity),
            measured_tc:         _parseNum(ri.measured_tc),
            measured_cya:        _parseNum(ri.measured_cya),
            rec_cl_gr:           _parseNum(ri.rec_cl_gr),
            rec_cl_tab:          _parseNum(ri.rec_cl_tab),
            rec_ph_kg:           _parseNum(ri.rec_ph_kg),
            rec_anti_l:          _parseNum(ri.rec_anti_l),
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
            operations:          (function(v) { if (!v) return []; try { return JSON.parse(v); } catch(e) { return []; } })(ri.operations),
            geo_lat:             _parseNum(ri.geo_lat),
            geo_lng:             _parseNum(ri.geo_lng),
            geo_accuracy:        _parseNum(ri.geo_accuracy),
            arrival_time:        ri.arrival_time || null,
            departure_time:      ri.departure_time || null,
            duration_minutes:    ri.duration_minutes !== '' && ri.duration_minutes != null ? Math.round(parseFloat(ri.duration_minutes)) || null : null,
            synced:              true
          };

          var local = localMap[parsed.intervention_id];
          if (!local) {
            // New from server
            try { await put('interventions', parsed); added++; } catch(e) {
              console.warn('[SYNC] Intervention put failed:', parsed.intervention_id, e.message);
            }
          } else if (local.synced === false) {
            // Local has unsynced changes — keep local version (will be pushed next cycle)
          } else {
            // Both synced — update with server version (server is truth)
            try { await put('interventions', parsed); updated++; } catch(e) {
              console.warn('[SYNC] Intervention update failed:', parsed.intervention_id, e.message);
            }
          }
        }
        console.log('[SYNC] Pulled interventions: ' + added + ' added, ' + updated + ' updated, ' + removedFromLocal + ' removed (server total: ' + data.interventions.length + ')');
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
