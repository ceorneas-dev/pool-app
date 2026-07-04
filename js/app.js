// app.js — Pool Manager PWA — Main application logic
// Features: Login + PIN, Dashboard, Intervention form, Client GPS location, Toast

'use strict';

// ── Global Error Handling + Error Log ────────────────────────
// Catches unhandled errors and stores last 50 in localStorage for debugging.
var _ERROR_LOG_KEY = 'pool_error_log';
var _ERROR_LOG_MAX = 50;

function _logError(source, message, extra) {
  try {
    var log = JSON.parse(localStorage.getItem(_ERROR_LOG_KEY) || '[]');
    log.unshift({
      ts: new Date().toISOString(),
      src: source,
      msg: String(message).substring(0, 300),
      extra: extra ? String(extra).substring(0, 200) : undefined
    });
    if (log.length > _ERROR_LOG_MAX) log = log.slice(0, _ERROR_LOG_MAX);
    localStorage.setItem(_ERROR_LOG_KEY, JSON.stringify(log));
  } catch (_) {}
}

function getErrorLog() {
  try { return JSON.parse(localStorage.getItem(_ERROR_LOG_KEY) || '[]'); } catch (_) { return []; }
}

function clearErrorLog() {
  try { localStorage.removeItem(_ERROR_LOG_KEY); } catch (_) {}
}

window.onerror = function(msg, src, line, col, err) {
  var loc = (src || '').split('/').pop() + ':' + line + ':' + col;
  console.error('[ERROR]', loc, msg);
  _logError('onerror', msg, loc);
  // Show toast if app is loaded (not during startup)
  if (typeof showToast === 'function' && APP && APP.currentScreen !== 'login') {
    showToast('Eroare: ' + String(msg).substring(0, 80), 'error', 5000);
  }
};

window.addEventListener('unhandledrejection', function(e) {
  var msg = e.reason ? (e.reason.message || String(e.reason)) : 'Promise rejected';
  console.error('[UNHANDLED]', msg);
  _logError('promise', msg);
  if (typeof showToast === 'function' && APP && APP.currentScreen !== 'login') {
    showToast('Eroare: ' + String(msg).substring(0, 80), 'error', 5000);
  }
});

// ── Helpers ───────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const $q = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);
const uid = () => 'i_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);

/**
 * Deduplicate interventions: keep only the newest per client_id + date.
 * "Newest" = latest created_at. Returns a new array (does not mutate input).
 */
function _deduplicateInterventions(interventions) {
  // Sort newest first so the first encountered per key wins
  var sorted = interventions.slice().sort(function(a, b) {
    return (b.created_at || '').localeCompare(a.created_at || '');
  });
  var seen = {};
  var result = [];
  for (var i = 0; i < sorted.length; i++) {
    // Normalize date for key — ensures "2026-03-18" matches even if stored differently
    var rawDate = String(sorted[i].date || '');
    var normDate = /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate : (function(d) { var p = new Date(d); return isNaN(p.getTime()) ? d : p.getFullYear() + '-' + ('0'+(p.getMonth()+1)).slice(-2) + '-' + ('0'+p.getDate()).slice(-2); })(rawDate);
    var key = String(sorted[i].client_id) + '_' + String(sorted[i].technician_id || '') + '_' + normDate;
    if (!seen[key]) {
      seen[key] = true;
      result.push(sorted[i]);
    }
  }
  return result;
}

// ── Global State ─────────────────────────────────────────────
const APP = {
  currentScreen:  'login',
  user:            null,
  selectedClient:  null,
  clients:         [],
  interventions:   [],
  pendingSync:     0,
  clGranUnit:      'gr',       // 'gr' or 'kg'
  currentPhotos:   [],
  currentPosition: null,       // GPS: {lat, lng, accuracy} — one-shot fix for distance badge / client location
  pinBuffer:       '',         // PIN input buffer
  installPrompt:   null,       // beforeinstallprompt event
  alertShown:      false,      // toast de alertă intervenții (1x per sesiune)
  alertThreshold:  4,          // prag configurabil (default 4)
  dashboardTab:    'all',      // 'all' | 'due'
  clientFormMode:  'add',      // 'add' | 'edit'
  wizardStep:      1,          // 1 | 2 | 3 — pasul curent al wizard-ului intervenție
  _stockProducts:  [],         // cache produse stoc (actualizat la deschidere formular)
  _billingClientId: null       // client_id pentru care se afișează butonul "Marchează facturat"
};

// ── Init ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Show version immediately
  const vb = document.getElementById('app-version-badge');
  if (vb) vb.textContent = 'v' + APP_VERSION;

  setupConnectivityIndicator();
  setupInstallPrompt();
  initApp();
});

const APP_VERSION = 222;

async function initApp() {
  await openDB();

  // ── Version tracking (no longer clears clients on update) ──
  try {
    var lastVer = await getSetting('app_version');
    if (parseInt(lastVer) !== APP_VERSION) {
      await setSetting('app_version', APP_VERSION);
      // v209: cleanup old flags that blocked technician sync
      await setSetting('techs_local_auth', null);
      await setSetting('techs_pending_push', false);
    }
  } catch(e) { console.warn('[INIT] Version check error:', e.message); }

  // Load configurable alert threshold
  const savedThr = await getSetting('alert_threshold');
  if (savedThr) APP.alertThreshold = parseInt(savedThr) || 4;

  const session = await getSession();

  if (session) {
    APP.user = session;
    // Check if PIN is set for this user
    const pinKey = await getSetting('pin_' + session.username);
    if (pinKey) {
      showPinScreen(session);
    } else {
      await postLogin();
    }
  } else {
    // Fără sesiune activă — încearcă auto-login din credențiale salvate
    const saved = getSavedCredentials();
    if (saved) {
      // Tentativă silențioasă — dacă eșuează (parolă schimbată) → afișează form
      await doLogin(saved.username, saved.password, true /* silent */);
    } else {
      showScreen('login');
      initLoginScreen();
    }
  }

  // Seed demo if DB empty — with backup recovery
  const techCount = await count('technicians');
  if (techCount === 0) {
    let restored = false;
    try {
      const backup = await getSetting('technicians_backup');
      if (backup) {
        const techs = JSON.parse(backup);
        if (techs && techs.length) {
          for (const t of techs) { try { await put('technicians', t); } catch(_) {} }
          restored = (await count('technicians')) > 0;
        }
      }
    } catch(_) {}
    if (!restored) {
      await seedDemoData();
      showToast('Date demo încărcate. Login: admin / admin123', 'info', 5000);
    }
  }
  // Always keep technicians backup fresh
  try {
    const allTechs = await getAll('technicians');
    if (allTechs.length) await setSetting('technicians_backup', JSON.stringify(allTechs));
  } catch(_) {}

  // Register service worker with auto-update detection
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
      .then(reg => {
        // Check for updates every 5 minutes
        setInterval(() => { reg.update().catch(() => {}); }, 300000);
        // When a new SW is installed, notify user to refresh
        reg.addEventListener('updatefound', () => {
          var newSW = reg.installing;
          if (newSW) {
            newSW.addEventListener('statechange', () => {
              if (newSW.state === 'activated' && navigator.serviceWorker.controller) {
                showToast('Versiune nouă disponibilă! Se reîncarcă...', 'info', 2000);
                setTimeout(() => window.location.reload(), 2000);
              }
            });
          }
        });
      })
      .catch(err => console.warn('[SW] Registration failed:', err));
  }

  setupNotifications();
  initSync();
  seedMissingStockProducts().catch(() => {});
}

// ── Screen Navigation ────────────────────────────────────────
/** Toggle nav menu overlay */
function toggleNavMenu() {
  const overlay = $('nav-menu-overlay');
  if (overlay) overlay.classList.toggle('open');
}

/** Open settings from header button — scroll to and open settings details */
function _updateExportFolderLabel() {
  var el = $('export-folder-name');
  if (!el) return;
  if (typeof _getExportDirHandle === 'function') {
    _getExportDirHandle().then(function(h) {
      el.textContent = h ? '📁 ' + h.name : 'Nesetat';
    }).catch(function() { el.textContent = 'Nesetat'; });
  } else {
    el.textContent = '';
  }
}

function openSettingsFromHeader() {
  const details = $('settings-section');
  if (details) {
    details.open = true;
    details.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function showScreen(name) {
  $$('.screen').forEach(s => s.classList.remove('active'));
  const el = $('screen-' + name);
  if (el) el.classList.add('active');
  APP.currentScreen = name;

  // Keyboard: focus search input when going to dashboard (mobile UX)
  if (name === 'dashboard') {
    setTimeout(() => {
      const s = $('search-input');
      if (s && window.innerWidth <= 640) s.focus();
    }, 350);
  }

  // Checklist: admin-only — redirect technician back to dashboard
  if (name === 'checklist') {
    if (!isAdmin()) { showScreen('dashboard'); return; }
    loadChecklistScreen();
  }

  // Info page: reset search, load stored content, show Edit button for admin
  if (name === 'info') {
    const infoSearch = $('info-search');
    if (infoSearch) { infoSearch.value = ''; filterInfoSections(''); }
    loadInfoContent(); // async — injects stored guide content if any
    const editBtn = $('btn-info-edit');
    if (editBtn) editBtn.style.display = (APP.user && APP.user.role === 'admin') ? '' : 'none';
  }

  // Keyboard: blur any active input when leaving a screen
  if (name !== 'dashboard' && name !== 'intervention') {
    if (document.activeElement && document.activeElement.tagName !== 'BODY') {
      document.activeElement.blur();
    }
  }
}

// ── Toast Notifications ───────────────────────────────────────
function showToast(msg, type = 'success', duration = 3000) {
  const container = $('toast-container');
  if (!container) return;

  const icons = { success: '✓', warning: '⚠', error: '✕', info: 'ℹ' };
  const toast = document.createElement('div');
  toast.className = 'toast ' + type;
  toast.innerHTML = `<span class="toast-icon">${icons[type] || 'ℹ'}</span><span>${escHtml(String(msg))}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ── Connectivity ─────────────────────────────────────────────
function setupConnectivityIndicator() {
  function update() {
    const badge = $('conn-badge');
    if (!badge) return;
    if (navigator.onLine) {
      badge.textContent = '\u{1F7E2}';
      badge.className = 'conn-dot online';
      badge.title = 'Online';
    } else {
      badge.textContent = '\u{1F534}';
      badge.className = 'conn-dot offline';
      badge.title = 'Offline';
    }
  }
  window.addEventListener('online',  () => { update(); forceSync().catch(() => {}); });
  window.addEventListener('offline', () => { update(); showToast('Conexiune pierdută. Datele se salvează local.', 'warning'); });
  update();

  // Sync callbacks
  window.onSyncComplete = () => {
    updateSyncBadge();
    // Refresh whichever screen is active
    if (APP.currentScreen === 'dashboard') loadData().then(renderDashboard);
    else if (APP.currentScreen === 'checklist') { if (typeof loadChecklistScreen === 'function') loadChecklistScreen(); }
  };
  window.onSyncError = (err) => {
    showToast('Eroare sincronizare: ' + (err && err.message || 'necunoscută'), 'error');
  };
}

function setupInstallPrompt() {
  window.addEventListener('beforeinstallprompt', e => {
    e.preventDefault();
    APP.installPrompt = e;
  });
}

// ── Saved Credentials (auto-login pe același dispozitiv) ─────
// Stocăm în localStorage — persistă chiar dacă IndexedDB e golit.
// Utilizat NUMAI în rețele interne / context business (nu financiar).
function getSavedCredentials() {
  try { return JSON.parse(localStorage.getItem('pool_creds') || 'null'); } catch { return null; }
}
function saveCredentials(u, p) {
  try { localStorage.setItem('pool_creds', JSON.stringify({ username: u, password: p })); } catch {}
}
function clearSavedCredentials() {
  try { localStorage.removeItem('pool_creds'); } catch {}
}

// ── Login Screen ─────────────────────────────────────────────
function initLoginScreen() {
  const form = $('login-form');
  if (!form) return;

  // Listener submit — atașat o singură dată (flag pe element)
  if (!form._loginListenerAdded) {
    form._loginListenerAdded = true;
    form.addEventListener('submit', async e => {
      e.preventDefault();
      const username = $('login-username').value.trim();
      const password = $('login-password').value;
      if (!username || !password) return;
      await doLogin(username, password);
    });
  }

  // Pre-completează din credențiale salvate (la fiecare afișare)
  const saved = getSavedCredentials();
  const switchEl = $('login-switch-user');
  if (saved) {
    const uEl = $('login-username');
    const pEl = $('login-password');
    if (uEl) uEl.value = saved.username;
    if (pEl) pEl.value = saved.password;
    if (switchEl) switchEl.style.display = '';
    // Focusează direct butonul — apasă Enter sau click
    setTimeout(() => { const btn = $('login-btn'); if (btn) btn.focus(); }, 120);
  } else {
    if (switchEl) switchEl.style.display = 'none';
    setTimeout(() => { const uEl = $('login-username'); if (uEl) uEl.focus(); }, 120);
  }
}

function switchLoginUser() {
  clearSavedCredentials();
  const uEl = $('login-username');
  const pEl = $('login-password');
  if (uEl) { uEl.value = ''; uEl.focus(); }
  if (pEl) pEl.value = '';
  const switchEl = $('login-switch-user');
  if (switchEl) switchEl.style.display = 'none';
}

async function doLogin(username, password, silent = false) {
  const btn = $('login-btn');
  if (!silent && btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>'; }

  try {
    let user = null;

    // Try API login if configured
    if (isSyncConfigured()) {
      try {
        user = await apiLogin(username, password);
      } catch {
        // fall through to local
      }
    }

    // Local login fallback
    if (!user) {
      let tech = null;
      try {
        // Try index lookup (fastest) — case-insensitive fallback below
        tech = await getByIndexFirst('technicians', 'username', username);
        if (!tech) tech = await getByIndexFirst('technicians', 'username', username.toLowerCase());
      } catch {
        // Index might not exist in old DB — scan all technicians
        const all = await getAll('technicians');
        tech = all.find(t => (t.username || '').toLowerCase() === username.toLowerCase()) || null;
      }
      if (tech && tech.password === password && tech.active !== false) {
        user = { technician_id: tech.technician_id, name: tech.name, role: tech.role, username: tech.username };
      }
    }

    if (!user) {
      if (silent) {
        // Credențiale salvate invalide (parolă schimbată?) → afișează form pre-completat
        showScreen('login');
        initLoginScreen();
        return;
      }
      showToast('Utilizator sau parolă incorectă.', 'error');
      if (btn) { btn.disabled = false; btn.textContent = 'Intră în cont'; }
      return;
    }

    // Salvăm credențialele pentru auto-login la viitoarele deschideri
    saveCredentials(username, password);

    APP.user = user;
    await setSession(user);
    await postLogin();
  } catch (err) {
    if (silent) {
      // Eroare silențioasă → afișează form
      showScreen('login');
      initLoginScreen();
      return;
    }
    showToast('Eroare la autentificare: ' + err.message, 'error');
    if (btn) { btn.disabled = false; btn.textContent = 'Intră în cont'; }
  }
}

async function postLogin() {
  APP.alertShown = false;  // reset per sesiune

  await loadData();
  renderDashboard();
  showScreen('dashboard');
  updateSyncBadge();
  // QR deeplink: ?client=ID
  setTimeout(checkClientDeeplink, 200);
}

// ── PIN Screen ───────────────────────────────────────────────
function showPinScreen(user) {
  showScreen('pin');
  APP.pinBuffer = '';
  renderPinDots();
  $('pin-username-label').textContent = 'Bine ai venit, ' + user.name;

  // Setup keypad (use onclick to avoid listener accumulation)
  $$('.pin-key').forEach(btn => {
    btn.onclick = function() {
      const val = btn.dataset.val;
      if (val === 'del') {
        APP.pinBuffer = APP.pinBuffer.slice(0, -1);
      } else if (APP.pinBuffer.length < 4) {
        APP.pinBuffer += val;
        if (APP.pinBuffer.length === 4) checkPin(user);
      }
      renderPinDots();
    };
  });

  const switchBtn = $('pin-switch-user');
  if (switchBtn) {
    switchBtn.onclick = async function() {
      await clearSession();
      APP.user = null;
      showScreen('login');
      initLoginScreen();
    };
  }
}

function renderPinDots() {
  $$('.pin-dot').forEach((dot, i) => {
    dot.classList.toggle('filled', i < APP.pinBuffer.length);
  });
}

async function checkPin(user) {
  const storedHash = await getSetting('pin_' + user.username);
  const inputHash  = btoa(APP.pinBuffer + user.username);
  if (inputHash === storedHash) {
    APP.user = user;
    await postLogin();
  } else {
    showToast('PIN incorect.', 'error');
    APP.pinBuffer = '';
    renderPinDots();
  }
}

async function setUserPin(pin) {
  if (!APP.user) return;
  const hash = btoa(pin + APP.user.username);
  await setSetting('pin_' + APP.user.username, hash);
  showToast('PIN setat cu succes.', 'success');
}

// ── Load Data ────────────────────────────────────────────────
async function loadData() {
  const [clients, interventions] = await Promise.all([
    getActiveClients(),
    getAll('interventions')
  ]);
  APP.clients       = clients;

  // Normalize operations field: ensure it's always an array (fixes old sync data stored as string)
  for (var ni = 0; ni < interventions.length; ni++) {
    var ops = interventions[ni].operations;
    if (typeof ops === 'string' && ops.length > 0) {
      try { interventions[ni].operations = JSON.parse(ops); } catch(e) { interventions[ni].operations = []; }
    } else if (!Array.isArray(ops)) {
      interventions[ni].operations = [];
    }
    // Normalize date to YYYY-MM-DD (fixes sort order when dates come as Date objects from GAS)
    var dt = interventions[ni].date;
    if (dt && !/^\d{4}-\d{2}-\d{2}$/.test(String(dt))) {
      var dp = new Date(dt);
      if (!isNaN(dp.getTime())) {
        interventions[ni].date = dp.getFullYear() + '-' + ('0' + (dp.getMonth() + 1)).slice(-2) + '-' + ('0' + dp.getDate()).slice(-2);
      }
    }
  }

  // Filter orphaned interventions (whose client doesn't exist locally)
  // Only remove from local display — do NOT track for server deletion!
  // The server is the source of truth; orphans may belong to clients
  // that are active on another device or just not synced yet.
  var clientIdSet = {};
  clients.forEach(function(c) { clientIdSet[String(c.client_id)] = true; });
  var orphaned = interventions.filter(function(i) { return !clientIdSet[String(i.client_id)]; });
  if (orphaned.length > 0) {
    APP.interventions = interventions.filter(function(i) { return clientIdSet[String(i.client_id)]; });
  } else {
    APP.interventions = interventions;
  }

  // Deduplicate: keep only the newest intervention per client+date
  // No dedup here — server is the single source of truth.
  // All interventions from server are kept as-is.
  APP.pendingSync   = APP.interventions.filter(i => !i.synced).length;

  // Auto-sync if no clients locally but sync is configured (once per session)
  if (APP.clients.length === 0 && isSyncConfigured() && !APP._autoSyncAttempted) {
    APP._autoSyncAttempted = true; // prevent infinite loop
    try { await forceSync(); } catch(e) {}
    var freshClients = await getActiveClients();
    if (freshClients.length > 0) {
      APP.clients = freshClients;
      var freshInt = await getAll('interventions');
      APP.interventions = freshInt;
      APP.pendingSync = freshInt.filter(i => !i.synced).length;
    }
  }
}

// ── Dashboard ────────────────────────────────────────────────
function isAdmin() {
  return APP.user && APP.user.role === 'admin';
}

function renderDashboard() {
  if (!APP.user) return;

  // Version badge
  const vBadge = $('app-version-badge');
  if (vBadge) vBadge.textContent = 'v' + APP_VERSION;

  // Apply role class on <body> — drives all .admin-only visibility via CSS
  document.body.classList.toggle('role-admin',      isAdmin());
  document.body.classList.toggle('role-technician', !isAdmin());

  // Aplica permisiuni granulare pentru tehnicieni (poate unhide admin-only)
  applyTechPermissions();

  // Show export folder name if set
  _updateExportFolderLabel();

  // User info + role badge
  const userEl = $('footer-user-name');
  if (userEl) userEl.textContent = APP.user.name;
  const roleEl = $('footer-user-role');
  if (roleEl) {
    roleEl.textContent = isAdmin() ? 'Admin' : 'Tehnician';
    roleEl.className   = 'role-badge ' + (isAdmin() ? 'role-badge-admin' : 'role-badge-tech');
  }

  // Stats
  const today = new Date().toISOString().split('T')[0];
  const todayCount = APP.interventions.filter(i => i.date === today).length;
  const el_total   = $('stat-total-clients');
  const el_today   = $('stat-today');
  const el_pending = $('stat-pending');
  if (el_total)   el_total.textContent   = APP.clients.length;
  if (el_today)   el_today.textContent   = todayCount;
  if (el_pending) el_pending.textContent = APP.pendingSync;

  // Billing count (admin only)
  if (isAdmin()) {
    var billingCount = _getBillableClients().length;
    var elBilling = $('stat-billing-count');
    if (elBilling) {
      elBilling.textContent = billingCount;
      elBilling.style.color = billingCount > 0 ? 'var(--danger)' : '';
    }
    var billingCard = $('stat-billing-card');
    if (billingCard) {
      var lbl = billingCard.querySelector('.stat-label');
      if (lbl) lbl.style.color = billingCount > 0 ? 'var(--danger)' : '';
    }
  }

  updateSyncBadge();
  renderClientList('');
  renderAdminStats();


  // Search
  const searchInput = $('search-input');
  if (searchInput) {
    searchInput.value = '';
    searchInput.oninput = e => renderClientList(e.target.value);
    searchInput.onkeydown = function(e) {
      if (e.key === 'Enter') { e.preventDefault(); searchInput.blur(); }
    };
  }

  // Dismiss keyboard on swipe/scroll (mobile UX — use ontouchmove to avoid accumulation)
  const dashboard = $('screen-dashboard');
  if (dashboard) {
    dashboard.ontouchmove = function() {
      if (document.activeElement && document.activeElement.tagName === 'INPUT') {
        document.activeElement.blur();
      }
    };
  }

  // Logout
  const logoutBtn = $('btn-logout-hidden');
  if (logoutBtn) {
    logoutBtn.onclick = async () => {
      APP.alertShown = false;
      await clearSession();      // Sesiunea curentă = ștearsă
      // Credențialele rămân salvate → la revenire form pre-completat + focus pe buton
      APP.user = null;
      APP.clients = [];
      APP.interventions = [];
      document.body.classList.remove('role-admin', 'role-technician');
      showScreen('login');
      initLoginScreen();
    };
  }

  // Sync badge — visible to all (info), clickable only for admin
  const syncBadge = $('sync-badge');
  if (syncBadge) {
    if (isAdmin()) {
      syncBadge.style.cursor = 'pointer';
      syncBadge.title  = 'Click pentru sincronizare manuală';
      syncBadge.onclick = async () => {
        if (!isSyncConfigured()) {
          showToast('API URL nu este configurat. Mergi la Setări.', 'error');
          return;
        }
        showToast('Sincronizare în curs...', 'info');
        try {
          await forceSync();
          await loadData();
          renderDashboard();
          updateSyncBadge();
          showToast('Sincronizare completă!', 'success');
        } catch (e) {
          showToast('Eroare la sincronizare: ' + e.message, 'error');
        }
      };
    } else {
      syncBadge.style.cursor = 'default';
      syncBadge.title  = '';
      syncBadge.onclick = null;
    }
  }

  // Export all button (admin only — also hidden via CSS .admin-only)
  const exportAllBtn = $('btn-export-all');
  if (exportAllBtn) {
    exportAllBtn.onclick = isAdmin() ? () => showExportModal(null) : null;
  }

  // Settings save
  const settingsBtn = $('btn-settings-save');
  if (settingsBtn) {
    settingsBtn.onclick = async () => {
      const url = $('settings-api-url');
      if (url) {
        var urlVal = url.value.trim();
        // If empty, keep the hardcoded default from SYNC_CONFIG
        if (urlVal) {
          SYNC_CONFIG.API_URL = urlVal;
          await setSetting('api_url', urlVal);
        }
        initSync();
      }
      // PIN setting
      const pin = $('settings-pin');
      if (pin && pin.value.length === 4 && /^\d{4}$/.test(pin.value)) {
        await setUserPin(pin.value);
        pin.value = '';
      }
      // Alert threshold setting
      const thrInput = $('settings-alert-threshold');
      if (thrInput && thrInput.value) {
        const v = parseInt(thrInput.value);
        if (v >= 1 && v <= 50) {
          APP.alertThreshold = v;
          await setSetting('alert_threshold', v);
          APP.alertShown = false; // permite re-evaluarea cu noul prag
        }
      }
      // WhatsApp notification (CallMeBot)
      const waPhoneEl = $('settings-wa-phone');
      const waKeyEl = $('settings-wa-apikey');
      if (waPhoneEl) await setSetting('wa_phone', waPhoneEl.value.trim());
      if (waKeyEl) await setSetting('wa_apikey', waKeyEl.value.trim());
      showToast('Setări salvate.', 'success');
      // Close settings section after saving
      const settingsDetails = $('settings-section');
      if (settingsDetails) settingsDetails.open = false;
    };
  }

  // Load settings into UI
  getSetting('api_url').then(url => {
    const urlInput = $('settings-api-url');
    if (urlInput) urlInput.value = url || SYNC_CONFIG.API_URL || '';
  });
  getSetting('alert_threshold').then(thr => {
    const thrInput = $('settings-alert-threshold');
    if (thrInput) thrInput.value = thr || APP.alertThreshold;
  });
  getSetting('wa_phone').then(val => {
    const el = $('settings-wa-phone');
    if (el && val) el.value = val;
  });
  getSetting('wa_apikey').then(val => {
    const el = $('settings-wa-apikey');
    if (el && val) el.value = val;
  });
  // Permisiuni tehnicieni
  getSetting('perm_tech_add_client').then(val => {
    const el = $('settings-perm-tech-add-client');
    if (el) el.checked = val === 'true' || val === true;
  });
}

// Salveaza permisiuni tehnicieni si aplica imediat pe pagina
async function savePermSettings() {
  const addEl = $('settings-perm-tech-add-client');
  const permAdd = !!(addEl && addEl.checked);
  await setSetting('perm_tech_add_client', permAdd ? 'true' : 'false');
  await applyTechPermissions();
  showToast('Permisiuni salvate.', 'success');
}

// Aplica (pentru utilizatorul curent) permisiunile acordate tehnicienilor.
// Admin: toate butoanele raman vizibile (.admin-only + role-admin).
// Tehnician: daca permisiunea e activa, scoatem clasa .admin-only de pe butonul specific.
async function applyTechPermissions() {
  // Buton controlat: Add Client tab
  const addBtn = document.querySelector('.tab-btn[onclick*="showAddClientModal"]');

  // Reset (re-adauga admin-only inainte de a evalua)
  if (addBtn) addBtn.classList.add('admin-only');

  if (isAdmin()) return; // admin vede tot oricum

  const permAdd = await getSetting('perm_tech_add_client');
  if ((permAdd === 'true' || permAdd === true) && addBtn) addBtn.classList.remove('admin-only');
}

async function renderClientList(searchTerm) {
  const list = $('client-list');
  if (!list) return;

  const term = (searchTerm || '').toLowerCase().trim();
  let filtered = APP.clients.filter(c =>
    !term ||
    c.name.toLowerCase().includes(term) ||
    (c.phone && c.phone.includes(term)) ||
    (c.address && c.address.toLowerCase().includes(term))
  );

  // Compute urgency for each client
  filtered = filtered.map(c => ({ client: c, urgency: getUrgencyLevel(c) }));

  // Tab filter: 'due' shows only overdue/never/soon
  if (APP.dashboardTab === 'due') {
    filtered = filtered.filter(f => f.urgency !== 'ok');
  }

  // Sort alphabetically by client name
  filtered.sort((a, b) => (a.client.name || '').localeCompare(b.client.name || '', 'ro'));

  // Update "De vizitat" tab badge count
  const dueCount = APP.clients.filter(c => getUrgencyLevel(c) !== 'ok').length;
  const dueBtnEl = $('tab-due-btn');
  if (dueBtnEl) dueBtnEl.textContent = dueCount > 0 ? `🔴 De vizitat (${dueCount})` : '🔴 De vizitat';

  if (!filtered.length) {
    list.innerHTML = '<li class="empty-state"><div class="empty-icon">🔍</div><p>Niciun client găsit</p></li>';
    return;
  }

  // Fetch all unread counts in parallel
  const counts = await Promise.all(filtered.map(f => getUnreportedCount(f.client.client_id)));
  const thr    = APP.alertThreshold;

  // Billing alert removed — De Facturat stat card handles this now
  APP.alertShown = true;

  list.innerHTML = filtered.map(({ client, urgency }, idx) => {
    const cnt       = counts[idx];
    const lastVisit = getLastVisitInfo(client.client_id);
    const distInfo  = getDistanceBadge(client);
    const hasNav    = client.location_set && client.latitude && client.longitude;

    const alertBadge = cnt >= thr
      ? `<span class="alert-badge danger">⚠ ${cnt} noi</span>`
      : cnt >= 2
        ? `<span class="alert-badge warn">⚡ ${cnt} noi</span>`
        : '';

    const admin = isAdmin();

    const resetBtn = (admin && cnt > 0)
      ? `<button class="btn-reset-counter"
           onclick="event.stopPropagation(); resetInterventionCounter('${client.client_id}')"
           title="Resetează contorizarea">↺ Reset</button>`
      : '';

    // Urgency badge
    const urgencyLabels = { overdue: '🔴 Vizită depășită', never: '⚫ Nicio vizită', soon: '🟡 Curând', ok: '' };
    const urgencyBadge = urgency !== 'ok'
      ? `<span class="urgency-badge urgency-${urgency}">${urgencyLabels[urgency]}</span>` : '';

    // Contact buttons
    const phone = client.phone ? client.phone.replace(/\D/g, '') : '';
    const callBtn = client.phone
      ? `<a href="tel:${client.phone}" class="btn-contact" onclick="event.stopPropagation()" title="Sună">📞</a>` : '';
    const waBtn = phone.length >= 9
      ? `<a href="https://wa.me/4${phone.slice(-9)}" target="_blank" rel="noopener" class="btn-contact" onclick="event.stopPropagation()" title="WhatsApp">💬</a>` : '';

    return `<li class="client-card urgency-${urgency}">
      <div class="client-card-main" onclick="openClientIntervention('${client.client_id}')">
        <div class="client-info">
          <div class="client-name">${escHtml(client.name)}</div>
          <div class="client-meta">
            <span class="client-volume">🌊 ${client.pool_volume_mc} m³ · ${client.pool_type}</span>
            ${client.phone ? `<span class="client-phone">📞 ${escHtml(client.phone)}</span>` : ''}
          </div>
          <div class="client-meta" style="margin-top:4px">
            ${lastVisit.badge}
            ${distInfo}
            ${urgencyBadge}
            ${alertBadge}
            ${resetBtn}
          </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end">
          ${callBtn}${waBtn}
          ${hasNav ? `<button class="btn-navigate" onclick="event.stopPropagation(); navigateToClient('${client.client_id}')" title="Navighează">🧭</button>` : ''}
        </div>
      </div>
      <div class="client-actions">
        <button class="client-action-btn" onclick="openClientIntervention('${client.client_id}')">➕ Intervenție nouă</button>
        <button class="client-action-btn" onclick="event.stopPropagation(); openVoiceIntervention('${client.client_id}')" style="color:var(--blue-600)" title="Intervenție rapidă — notă vocală">🎙️ Rapid</button>
        <button class="client-action-btn" onclick="showClientDetails('${client.client_id}')">ℹ️ Info</button>
        ${admin ? `<button class="client-action-btn" onclick="showEditClientModal('${client.client_id}')">✏️ Editează</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="showQRCode('${client.client_id}')">📱 QR</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="showExportModal('${client.client_id}')">📥 Export</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="event.stopPropagation(); setClientLocation('${client.client_id}')" style="color:var(--emerald-600)">📍 ${client.location_set ? 'Relocare' : 'Locație'}</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="event.stopPropagation(); deleteClient('${client.client_id}')" style="color:var(--danger)">🗑️ Șterge</button>` : ''}
      </div>
    </li>`;
  }).join('');
}

function getLastVisitInfo(clientId) {
  const ci = APP.interventions.filter(i => i.client_id === clientId);
  if (!ci.length) return { badge: '<span class="last-visit-badge none">Nicio vizită</span>', days: null };

  const latest = ci.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')))[0];
  if (!latest || !latest.date) return { badge: '<span class="last-visit-badge none">Nicio vizită</span>', days: null };
  const days = Math.floor((Date.now() - Date.parse(latest.date)) / 86400000);
  if (isNaN(days)) return { badge: '<span class="last-visit-badge none">Dată necunoscută</span>', days: null };
  let cls = 'good', label = 'Ultima vizită: ' + days + ' zile';
  if (days > 30) cls = 'overdue';
  else if (days > 14) cls = 'warn';
  if (days === 0) label = 'Ultima vizită: azi';
  else if (days === 1) label = 'Ultima vizită: ieri';

  return { badge: `<span class="last-visit-badge ${cls}">${label}</span>`, days };
}

function getDistanceBadge(client) {
  if (!APP.currentPosition || !client.location_set || !client.latitude || !client.longitude) return '';
  const dist = haversineDistance(APP.currentPosition.lat, APP.currentPosition.lng, client.latitude, client.longitude);
  const label = dist < 1 ? Math.round(dist * 1000) + ' m' : dist.toFixed(1) + ' km';
  return `<span class="distance-badge">📍 ~${label}</span>`;
}

function navigateToClient(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client || !client.latitude) return;
  const url = `https://www.google.com/maps/dir/?api=1&destination=${client.latitude},${client.longitude}`;
  window.open(url, '_blank');
}

/** Set client GPS location from current device position (any logged-in user). */
async function setClientLocation(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;
  if (!navigator.geolocation) {
    showToast('GPS nu este disponibil pe acest dispozitiv.', 'error');
    return;
  }
  const wasSet = !!client.location_set;
  showToast('Se obține locația...', 'info', 3000);
  navigator.geolocation.getCurrentPosition(
    async (pos) => {
      client.latitude     = pos.coords.latitude;
      client.longitude    = pos.coords.longitude;
      client.location_set = true;
      client.updated_at   = new Date().toISOString();
      await put('clients', client);
      // Push to GAS
      if (isSyncConfigured()) {
        apiFetch(SYNC_CONFIG.API_URL, {
          method: 'POST',
          body: JSON.stringify({ action: 'push', type: 'clients', data: [client] })
        }).catch(err => console.warn('[SYNC] Client loc push failed:', err.message));
      }
      logLocationAudit(wasSet ? 'update_location' : 'set_location', client);
      showToast('📍 Locația salvată pentru ' + client.name, 'success');
      renderClientList($('search-input') ? $('search-input').value : '');
      // Refresh the client-info modal's location row, if it's open for this client
      const locStatusEl = $('client-detail-gps-status');
      if (locStatusEl) locStatusEl.textContent = '✅ Setată';
      const updBtn = $('client-detail-gps-update-btn');
      if (updBtn) updBtn.textContent = '📍 Actualizează';
      const delBtn = $('client-detail-gps-delete-btn');
      if (delBtn) delBtn.style.display = '';
    },
    (err) => {
      showToast('Eroare GPS: ' + err.message, 'error');
    },
    { enableHighAccuracy: true, timeout: 10000 }
  );
}

/** Delete (clear) a client's saved GPS location — any logged-in user. */
async function deleteClientLocation(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client || !client.location_set) return;
  if (!confirm('Ștergi locația GPS salvată pentru ' + client.name + '?')) return;

  client.latitude     = null;
  client.longitude    = null;
  client.location_set = false;
  client.updated_at   = new Date().toISOString();
  await put('clients', client);
  // Push to GAS
  if (isSyncConfigured()) {
    apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'clients', data: [client] })
    }).catch(err => console.warn('[SYNC] Client loc delete push failed:', err.message));
  }
  logLocationAudit('delete_location', client);
  showToast('🗑️ Locația a fost ștearsă pentru ' + client.name, 'success');
  renderClientList($('search-input') ? $('search-input').value : '');
  // Refresh the client-info modal's location row, if it's open for this client
  const locStatusEl = $('client-detail-gps-status');
  if (locStatusEl) locStatusEl.textContent = '❌ Nesetată';
  const updBtn = $('client-detail-gps-update-btn');
  if (updBtn) updBtn.textContent = '📍 Adaugă';
  const delBtn = $('client-detail-gps-delete-btn');
  if (delBtn) delBtn.style.display = 'none';
}

/** Fire-and-forget audit log entry for a client GPS location change (set/update/delete). */
function logLocationAudit(actionType, client) {
  if (!APP.user || !isSyncConfigured()) return;
  apiFetch(SYNC_CONFIG.API_URL, {
    method: 'POST',
    body: JSON.stringify({
      action:           'logAudit',
      technician_id:    APP.user.technician_id,
      technician_name:  APP.user.name,
      log_action:       actionType,
      client_id:        client.client_id,
      client_name:      client.name,
      timestamp:        new Date().toISOString()
    })
  }).catch(err => console.warn('[AUDIT] log failed:', err.message));
}

const AUDIT_ACTION_LABELS = {
  set_location:    '📍 Locație adăugată',
  update_location: '📍 Locație actualizată',
  delete_location: '🗑️ Locație ștearsă'
};

/** Admin: show the GPS-location audit log (who / when / what changed). */
async function showAuditLogModal() {
  const modal = $('modal-audit-log');
  const body  = $('audit-log-modal-body');
  if (!modal || !body) {
    showToast('Aplicația s-a actualizat parțial. Închide-o complet și redeschide-o, apoi încearcă din nou.', 'warning', 6000);
    return;
  }
  modal.classList.add('open');

  if (!isSyncConfigured()) {
    body.innerHTML = '<p style="color:var(--text-secondary);font-size:.85rem">Configurați API URL în Setări pentru a vedea jurnalul.</p>';
    return;
  }
  body.innerHTML = '<p style="color:var(--text-secondary);font-size:.85rem">Se încarcă...</p>';
  try {
    const data = await apiFetch(SYNC_CONFIG.API_URL + '?action=getAuditLog', { cache: 'no-store' });
    const entries = data.entries || [];
    if (!entries.length) {
      body.innerHTML = '<p style="color:var(--text-secondary);font-size:.85rem">Nicio intrare în jurnal.</p>';
      return;
    }
    body.innerHTML = '<div style="display:flex;flex-direction:column;gap:8px">' + entries.map(e => {
      const label = AUDIT_ACTION_LABELS[e.log_action] || e.log_action;
      const dt = new Date(e.timestamp);
      const dtLabel = isNaN(dt.getTime()) ? e.timestamp : dt.toLocaleString('ro-RO', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      return '<div style="padding:8px 10px;border:1px solid var(--slate-200);border-radius:8px;font-size:.82rem">' +
        '<div style="font-weight:600">' + escHtml(label) + ' — ' + escHtml(e.client_name || e.client_id || '') + '</div>' +
        '<div style="color:var(--text-secondary);margin-top:2px">' + escHtml(e.technician_name || '') + ' · ' + dtLabel + '</div>' +
        '</div>';
    }).join('') + '</div>';
  } catch (err) {
    body.innerHTML = '<p style="color:var(--danger);font-size:.85rem">Eroare la încărcarea jurnalului: ' + escHtml(err.message) + '</p>';
  }
}

function closeAuditLogModal() {
  const modal = $('modal-audit-log');
  if (modal) modal.classList.remove('open');
}

// ════════════════════════════════════════════════════════════════
// VOICE-NOTE QUICK INTERVENTION — record audio, log a minimal
// intervention automatically, upload the recording to Drive so the
// treatment can be filled in manually later from the audio.
// ════════════════════════════════════════════════════════════════

var _voiceRecorder     = null;
var _voiceChunks       = [];
var _voiceStream       = null;
var _voiceTimerInt     = null;
var _voiceStartTs      = 0;
var _voiceBlob         = null;
var _voiceClientId     = null;

/** Open the quick voice-note modal for a client. */
function openVoiceIntervention(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;
  const modal = $('modal-voice-intervention');
  if (!modal) {
    showToast('Aplicația s-a actualizat parțial. Închide-o complet și redeschide-o, apoi încearcă din nou.', 'warning', 6000);
    return;
  }
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia || typeof MediaRecorder === 'undefined') {
    showToast('Înregistrarea audio nu este suportată pe acest dispozitiv/browser.', 'error');
    return;
  }
  _voiceClientId = clientId;
  _voiceBlob = null;
  $('voice-client-name').textContent = client.name;
  $('voice-rec-idle').style.display = '';
  $('voice-rec-active').style.display = 'none';
  $('voice-rec-preview').style.display = 'none';
  modal.classList.add('open');
}

function closeVoiceIntervention() {
  _stopVoiceStream();
  if (_voiceRecorder && _voiceRecorder.state === 'recording') {
    try { _voiceRecorder.stop(); } catch (e) {}
  }
  if (_voiceTimerInt) { clearInterval(_voiceTimerInt); _voiceTimerInt = null; }
  _voiceBlob = null;
  $('modal-voice-intervention').classList.remove('open');
}

function _stopVoiceStream() {
  if (_voiceStream) {
    _voiceStream.getTracks().forEach(function(t) { t.stop(); });
    _voiceStream = null;
  }
}

/** Pick the best-supported audio mime type for MediaRecorder on this browser. */
function _pickVoiceMimeType() {
  var candidates = ['audio/webm;codecs=opus', 'audio/webm', 'audio/mp4', 'audio/ogg'];
  for (var i = 0; i < candidates.length; i++) {
    if (MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(candidates[i])) return candidates[i];
  }
  return '';
}

async function startVoiceRecording() {
  $('voice-rec-preview').style.display = 'none';
  $('voice-rec-idle').style.display = 'none';
  $('voice-rec-active').style.display = '';

  try {
    _voiceStream = await navigator.mediaDevices.getUserMedia({ audio: true });
  } catch (e) {
    showToast('Nu am putut accesa microfonul: ' + e.message, 'error');
    $('voice-rec-active').style.display = 'none';
    $('voice-rec-idle').style.display = '';
    return;
  }

  var mimeType = _pickVoiceMimeType();
  _voiceChunks = [];
  _voiceRecorder = mimeType ? new MediaRecorder(_voiceStream, { mimeType: mimeType }) : new MediaRecorder(_voiceStream);
  _voiceRecorder.ondataavailable = function(e) { if (e.data && e.data.size > 0) _voiceChunks.push(e.data); };
  _voiceRecorder.onstop = function() {
    _voiceBlob = new Blob(_voiceChunks, { type: _voiceRecorder.mimeType || mimeType || 'audio/webm' });
    _stopVoiceStream();
    var audioEl = $('voice-rec-audio');
    if (audioEl) audioEl.src = URL.createObjectURL(_voiceBlob);
    $('voice-rec-active').style.display = 'none';
    $('voice-rec-preview').style.display = '';
  };
  _voiceRecorder.start();

  _voiceStartTs = Date.now();
  _updateVoiceTimer();
  _voiceTimerInt = setInterval(_updateVoiceTimer, 500);
}

function _updateVoiceTimer() {
  var el = $('voice-rec-timer');
  if (!el) return;
  var sec = Math.floor((Date.now() - _voiceStartTs) / 1000);
  el.textContent = Math.floor(sec / 60) + ':' + String(sec % 60).padStart(2, '0');
}

function stopVoiceRecording() {
  if (_voiceTimerInt) { clearInterval(_voiceTimerInt); _voiceTimerInt = null; }
  if (_voiceRecorder && _voiceRecorder.state === 'recording') {
    _voiceRecorder.stop();
  }
}

/** Save the recorded note as a minimal intervention, then upload the audio to Drive in background. */
async function saveVoiceIntervention() {
  const clientId = _voiceClientId;
  const audioBlob = _voiceBlob;
  const client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client || !audioBlob || !APP.user) return;

  closeVoiceIntervention();
  showToast('🎙️ Se salvează nota vocală...', 'info', 4000);

  // Upload the recording to Drive FIRST so the link is baked into the very first save —
  // patching it in afterward raced with the normal push/pull sync cycle and could get lost.
  const audioFileUrl = await _uploadVoiceAudio(audioBlob, client);

  const today = new Date().toISOString().split('T')[0];
  const intervention = {
    intervention_id:  uid(),
    client_id:        client.client_id,
    client_name:      client.name,
    technician_id:    APP.user.technician_id,
    technician_name:  APP.user.name,
    date:             today,
    created_at:       new Date().toISOString(),
    measured_chlorine: null,
    measured_ph:       null,
    observations:     '🎙️ Notă vocală — completați clorul/pH-ul și tratamentul din înregistrarea audio.',
    operations:       [],
    photos:           [],
    synced:           false,
    audio_file_url:   audioFileUrl
  };

  await saveIntervention(intervention);
  APP.interventions.push(intervention);
  APP.pendingSync++;

  showToast('🎙️ Intervenție rapidă salvată pentru ' + client.name, 'success');
  forceSync().catch(function() {});
  await loadData();
  renderDashboard();
}

/** Upload the audio recording to Google Drive ("Export Interventii/<client>"). Returns the file URL, or null. */
async function _uploadVoiceAudio(audioBlob, client) {
  if (!isSyncConfigured()) return null;
  try {
    var ext = audioBlob.type.indexOf('mp4') !== -1 ? 'm4a' : (audioBlob.type.indexOf('ogg') !== -1 ? 'ogg' : 'webm');
    var fname = 'Audio_' + sanitizeFilename(client.name) + '_' + new Date().toISOString().slice(0, 10) + '_' + Date.now() + '.' + ext;
    var b64 = await _blobToBase64(audioBlob);

    const res = await apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({
        action:   'saveExportToDrive',
        fileName: fname,
        data:     b64,
        mimeType: audioBlob.type || 'audio/webm',
        clientName: client.name
      })
    });

    if (res && res.success && res.fileUrl) return res.fileUrl;
    console.warn('[VOICE] Drive upload did not return a file URL:', res && res.error);
    return null;
  } catch (e) {
    console.warn('[VOICE] Audio upload failed:', e.message);
    showToast('Nota a fost salvată, dar încărcarea audio în Drive a eșuat.', 'warning');
    return null;
  }
}

/** Convert a Blob to a base64 string (no data: prefix). */
function _blobToBase64(blob) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onloadend = function() {
      var result = reader.result || '';
      var comma = result.indexOf(',');
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function openClientIntervention(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;
  APP.selectedClient = client;
  // Dismiss keyboard when leaving search
  const si = $('search-input');
  if (si) si.blur();
  renderIntervention(client);
  showScreen('intervention');
}

function updateSyncBadge() {
  const badge = $('sync-badge');
  if (!badge) return;
  APP.pendingSync = APP.interventions.filter(i => !i.synced).length;
  if (APP.pendingSync > 0) {
    badge.textContent = '⬆ ' + APP.pendingSync + ' nesincronizat' + (APP.pendingSync > 1 ? 'e' : 'ă');
    badge.classList.add('visible');
  } else {
    badge.classList.remove('visible');
  }
}

// ── Alert counter helpers ─────────────────────────────────────
async function getUnreportedCount(clientId) {
  const total    = APP.interventions.filter(i => i.client_id === clientId).length;
  const reported = await getSetting('reported_count_' + clientId);
  return Math.max(0, total - (parseInt(reported) || 0));
}

async function resetInterventionCounter(clientId) {
  const total = APP.interventions.filter(i => i.client_id === clientId).length;
  await setSetting('reported_count_' + clientId, total);
  showToast('Contorizare resetată.', 'success');
  APP.alertShown = false;
  await loadData();
  renderDashboard();
}


// -- Operations & Prices for export --
var DEFAULT_OPERATIONS = [
  'Aspirare piscina',
  'Curatare linie apa',
  'Curatare skimmere',
  'Spalare filtru',
  'Curatare prefiltru',
  'Periere piscina',
  'Analiza apei',
  'Tratament chimic',
  'Verificare automatizare'
];


/** Get operations list from storage (falls back to DEFAULT_OPERATIONS) */
async function getOperations() {
  try {
    var stored = await getByKey('settings', 'operations_list');
    if (stored && Array.isArray(stored.value) && stored.value.length > 0) return stored.value;
  } catch (e) {}
  return DEFAULT_OPERATIONS.slice();
}

/** Save operations list to storage */
async function saveOperationsList(arr) {
  await put('settings', { key: 'operations_list', value: arr });
}

/** Render operations list in Settings panel */
async function renderOpsSettings() {
  var list = $('ops-settings-list');
  if (!list) return;
  var ops = await getOperations();
  if (!ops.length) {
    list.innerHTML = '<p style="font-size:.8rem;color:var(--slate-400);padding:4px 0">Nicio operatiune.</p>';
    return;
  }
  list.innerHTML = ops.map(function(op, i) {
    return '<div class="obs-tmpl-setting-row">' +
      '<span class="obs-tmpl-setting-text">' + escHtml(op) + '</span>' +
      '<button class="obs-tmpl-del-btn" onclick="deleteOperation(' + i + ')" title="Sterge">&#128465;</button>' +
      '</div>';
  }).join('');
}

/** Add a new operation */
async function addOperation() {
  var input = $('ops-new-input');
  var text = input ? input.value.trim() : '';
  if (!text) { showToast('Scrie numele operatiunii.', 'warning'); return; }
  var ops = await getOperations();
  if (ops.indexOf(text) !== -1) { showToast('Operatiunea exista deja.', 'warning'); return; }
  ops.push(text);
  await saveOperationsList(ops);
  if (input) input.value = '';
  renderOpsSettings();
  showToast('Operatiune adaugata.', 'success');
}

/** Delete an operation by index */
async function deleteOperation(index) {
  var ops = await getOperations();
  ops.splice(index, 1);
  await saveOperationsList(ops);
  renderOpsSettings();
  showToast('Operatiune stearsa.', 'success');
}

var DEFAULT_PRICES = {
  pret_interventie: 250
};

async function getExportPrices() {
  try {
    var saved = await getSetting('export_prices');
    if (saved) return Object.assign({}, DEFAULT_PRICES, JSON.parse(saved));
  } catch(_) {}
  return Object.assign({}, DEFAULT_PRICES);
}

async function saveExportPrices(prices) {
  await setSetting('export_prices', JSON.stringify(prices));
}

// ── Intervention Screen ───────────────────────────────────────
async function renderIntervention(client) {
  // Header
  const nameEl = $('intervention-client-name');
  const dateEl = $('intervention-date');
  if (nameEl) {
    nameEl.textContent = client.name;
    // Add info icon button if not already present
    var infoBtn = $('client-info-btn');
    if (!infoBtn) {
      infoBtn = document.createElement('button');
      infoBtn.id = 'client-info-btn';
      infoBtn.className = 'client-info-btn';
      infoBtn.title = 'Info client';
      infoBtn.innerHTML = 'ℹ️';
      nameEl.parentNode.insertBefore(infoBtn, nameEl.nextSibling);
    }
    infoBtn.onclick = function() { showClientDetails(client.client_id); };
  }
  if (dateEl) {
    var today = new Date();
    var todayStr = today.toISOString().split('T')[0];
    dateEl.textContent = today.toLocaleDateString('ro-RO', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    dateEl.title = 'Click pentru a schimba data';
    var picker = $('intervention-date-picker');
    if (picker) picker.value = todayStr;
    APP._interventionDate = todayStr;
  }

  // Render operations checklist in step 2
  var opsContainer = $('ops-checklist');
  if (opsContainer) {
    getOperations().then(function(opsList) {
      var opsHtml = '';
      opsList.forEach(function(op, idx) {
        opsHtml += '<label class="ops-check-item"><input type="checkbox" id="op-' + idx + '" class="ops-checkbox" value="' + escHtml(op) + '"><span>' + escHtml(op) + '</span></label>';
      });
      opsContainer.innerHTML = opsHtml;
      // If editing, re-check the operations
      if (APP._editingIntervention && APP._editingIntervention.operations) {
        var editOps = APP._editingIntervention.operations;
        document.querySelectorAll('.ops-checkbox').forEach(function(chk) {
          if (editOps.indexOf(chk.value) >= 0) chk.checked = true;
        });
      }
    });
  }

  APP.currentPhotos  = [];
  APP.currentPosition = null;

  resetInterventionForm();
  updateRecommendation();
  setupPhotoCapture();

  // GPS capture (non-blocking)
  updateGpsIndicator('locating');
  getCurrentPosition().then(pos => {
    APP.currentPosition = pos;
    updateGpsIndicator(pos ? 'located' : 'no-gps');

    // Offer to set client location if not set
    if (pos && !client.location_set) {
      showSetLocationPrompt(client, pos);
    }

    // Update distance badges in dashboard list (in background)
    if (pos && APP.currentScreen === 'dashboard') {
      renderClientList($('search-input') ? $('search-input').value : '');
    }
  });

  // Back button
  const backBtn = $('btn-back');
  if (backBtn) backBtn.onclick = () => { APP._editingIntervention = null; showScreen('dashboard'); };

  // Save button — managed by switchP2Tab()

  // Show CYA input only for exterior pools
  const cyaWrap = $('measure-cya-wrap');
  if (cyaWrap) cyaWrap.style.display = (client.pool_type === 'exterior') ? '' : 'none';

  // Recommendation auto-update
  const measuredInputs = ['m-chlorine', 'm-ph', 'm-tc', 'm-cya', 'm-alkalinity', 'm-hardness'];
  measuredInputs.forEach(id => {
    const el = $(id);
    if (el) el.oninput = updateRecommendation;
  });

  // Previous interventions
  renderPreviousInterventions(client);

  // Observation template chips
  renderObsTemplates();

  // Wizard: reset to step 1 + render dynamic treatment steppers
  goWizardStep(1);
  renderTreatmentSteppers().then(function() {
    // If editing an existing intervention, prefill all fields
    if (APP._editingIntervention) {
      _prefillInterventionForm(APP._editingIntervention);
    }
  }).catch(e => console.warn('[STEPPER] Error:', e));
}

function updateGpsIndicator(state) {
  const el = $('gps-indicator');
  if (!el) return;
  if (state === 'locating') {
    el.textContent = '📍 Se localizează...';
    el.className = 'gps-indicator locating';
  } else if (state === 'located') {
    const acc = APP.currentPosition ? Math.round(APP.currentPosition.accuracy) : '?';
    el.textContent = `📍 Localizat (±${acc}m)`;
    el.className = 'gps-indicator located';
  } else {
    el.textContent = '📍 Fără GPS';
    el.className = 'gps-indicator no-gps';
  }
}

function showSetLocationPrompt(client, pos) {
  // Non-blocking: show toast with option to set
  showToast(`Setați locația clientului ${client.name}?`, 'info', 8000);
  // Optionally we could add a "Da" button in the toast — for now just auto-set
  client.latitude    = pos.lat;
  client.longitude   = pos.lng;
  client.location_set = true;
  put('clients', client).then(() => {
    APP.clients = APP.clients.map(c => c.client_id === client.client_id ? client : c);
  });
}

function _prefillInterventionForm(intv) {
  // Date
  if (intv.date) {
    APP._interventionDate = intv.date;
    var dateEl = $('intervention-date');
    if (dateEl) {
      var d = new Date(intv.date + 'T12:00:00');
      dateEl.textContent = d.toLocaleDateString('ro-RO', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    }
    var picker = $('intervention-date-picker');
    if (picker) picker.value = intv.date;
  }

  // Measured values
  var measuredFields = {
    'm-chlorine': intv.measured_chlorine,
    'm-ph': intv.measured_ph,
    'm-temp': intv.measured_temp,
    'm-hardness': intv.measured_hardness,
    'm-alkalinity': intv.measured_alkalinity,
    'm-salinity': intv.measured_salinity,
    'm-tc': intv.measured_tc,
    'm-cya': intv.measured_cya
  };
  Object.keys(measuredFields).forEach(function(id) {
    var el = $(id);
    if (el && measuredFields[id] != null) el.value = measuredFields[id];
  });

  // Observations
  var obs = $('observations');
  if (obs && intv.observations) obs.value = intv.observations;

  // Treatment stepper values - use setTimeout to wait for renderTreatmentSteppers
  setTimeout(function() {
    Object.keys(intv).forEach(function(key) {
      if (key.startsWith('treat_') && intv[key]) {
        var inputId = 't-' + key.substring(6); // treat_xxx -> t-xxx
        var el = $(inputId);
        if (el) el.value = intv[key];
      }
    });
  }, 500);

  // Photos
  if (intv.photos && intv.photos.length) {
    APP.currentPhotos = [].concat(intv.photos);
    renderPhotoGrid();
  }

  // Update recommendation with pre-filled values
  updateRecommendation();

  // Show edit indicator
  showToast('Editezi interventie din ' + fmtDate(intv.date), 'info');
}

function resetInterventionForm() {
  // Measured values
  ['m-chlorine','m-ph','m-temp','m-hardness','m-alkalinity','m-salinity','m-tc','m-cya'].forEach(id => {
    const el = $(id);
    if (el) { el.value = ''; el.classList.remove('error'); }
  });

  // Recommendation display
  ['rec-cl-granule','rec-cl-tab','rec-ph-kg','rec-anti'].forEach(id => {
    const el = $(id);
    if (el) el.textContent = '—';
  });

  // Treatment steppers — reset all dynamic inputs
  $$('#treatment-steppers-container input[type="number"]').forEach(el => { el.value = '0'; });

  // Observations + chips
  const obs = $('observations');
  if (obs) obs.value = '';
  $$('.obs-chip').forEach(el => el.classList.remove('active'));

  // Photos
  APP.currentPhotos = [];
  renderPhotoGrid();
}

function updateRecommendation() {
  const vol = APP.selectedClient ? APP.selectedClient.pool_volume_mc : 0;
  const cl  = parseFloat($('m-chlorine')  ? $('m-chlorine').value  : '') || null;
  const ph  = parseFloat($('m-ph')        ? $('m-ph').value        : '') || null;
  const tc  = parseFloat($('m-tc')        ? $('m-tc').value        : '') || null;
  const ta  = parseFloat($('m-alkalinity')? $('m-alkalinity').value: '') || null;
  const ch  = parseFloat($('m-hardness')  ? $('m-hardness').value  : '') || null;
  const cya = parseFloat($('m-cya')       ? $('m-cya').value       : '') || null;

  // CC = Total Chlorine − FAC (clamp to 0)
  const cc = (tc !== null && cl !== null) ? Math.round(Math.max(0, tc - cl) * 100) / 100 : null;

  // ── Status badges ──────────────────────────────────────────
  const badgesEl = $('rec-status-badges');
  if (badgesEl) {
    const params = [
      { key: 'fac', label: 'Clor (FAC)', val: cl,  unit: 'ppm' },
      { key: 'ph',  label: 'pH',         val: ph,  unit: ''    },
      { key: 'ta',  label: 'Alcalinitate', val: ta, unit: 'ppm'},
      { key: 'ch',  label: 'Duritate',   val: ch,  unit: 'ppm' },
      { key: 'cc',  label: 'CC',         val: cc,  unit: 'ppm' },
    ];
    if (cya !== null) params.push({ key: 'cya', label: 'CYA', val: cya, unit: 'ppm' });
    const filled = params.filter(p => p.val !== null);
    if (filled.length) {
      badgesEl.style.display = '';
      badgesEl.innerHTML = filled.map(p => {
        const st = getParameterStatus(p.key, p.val);
        if (!st) return '';
        const valStr = p.val + (p.unit ? '\u00a0' + p.unit : '');
        return `<span class="status-badge status-${st.status}">${escHtml(p.label)}: ${valStr} <em>${st.label}</em></span>`;
      }).join('');
    } else {
      badgesEl.style.display = 'none';
    }
  }

  // ── pH efficiency + CC analysis ────────────────────────────
  const analysisEl = $('rec-analysis');
  if (analysisEl) {
    const parts = [];
    if (ph !== null) {
      const eff = getPhEfficiency(ph);
      const cls = eff >= 55 ? 'eff-tag-ok' : eff >= 33 ? 'eff-tag-warn' : 'eff-tag-bad';
      parts.push(`<span class="ph-eff-tag ${cls}">pH ${ph} → clor <strong>${eff}%</strong> eficient</span>`);
    }
    if (cc !== null) {
      const cls = cc <= 0.2 ? 'eff-tag-ok' : cc <= 0.5 ? 'eff-tag-warn' : 'eff-tag-bad';
      parts.push(`<span class="ph-eff-tag ${cls}">CC = <strong>${cc}\u00a0ppm</strong></span>`);
    }
    if (parts.length) { analysisEl.style.display = ''; analysisEl.innerHTML = parts.join(''); }
    else { analysisEl.style.display = 'none'; }
  }

  // ── Breakpoint chlorination alert ──────────────────────────
  const bpEl = $('rec-breakpoint');
  if (bpEl) {
    if (cc !== null && cc > 0.5) {
      const dose = Math.round(cc * 10 * 100) / 100;
      bpEl.style.display = '';
      bpEl.innerHTML = `⚡ <strong>Breakpoint necesar!</strong> CC = ${cc}\u00a0ppm → adaugă <strong>${dose}\u00a0ppm clor nestabilizat</strong> (Ca(OCl)₂ sau NaOCl)`;
    } else { bpEl.style.display = 'none'; }
  }

  // ── CYA-adjusted FAC minimum ───────────────────────────────
  const cyaEl = $('rec-cya-min');
  if (cyaEl) {
    if (cya !== null && cya > 0) {
      const facMin = Math.round(cya * 0.075 * 100) / 100;
      cyaEl.style.display = '';
      const danger = cya > 100 ? ' <span class="status-badge status-danger">Diluție obligatorie!</span>' : '';
      cyaEl.innerHTML = `💡 CYA = ${cya}\u00a0ppm → FAC minim necesar: <strong>${facMin}\u00a0ppm</strong>${danger}`;
    } else { cyaEl.style.display = 'none'; }
  }

  // ── Dose recommendations (existing logic) ─────────────────
  if (!vol || cl === null || ph === null) {
    ['rec-cl-granule','rec-cl-tab','rec-ph-kg','rec-anti'].forEach(id => {
      const el = $(id); if (el) el.textContent = '—';
    });
    const extHide2 = $('rec-extrapolation');
    if (extHide2) extHide2.style.display = 'none';
    updateSaveButton();
    return;
  }

  const rec = getRecommendation(vol, cl, ph);
  if (!rec) {
    ['rec-cl-granule','rec-cl-tab','rec-ph-kg','rec-anti'].forEach(id => {
      const el = $(id); if (el) el.textContent = 'N/A';
    });
    const extHide = $('rec-extrapolation');
    if (extHide) extHide.style.display = 'none';
    updateSaveButton();
    return;
  }

  const elGr = $('rec-cl-granule');
  const elTab = $('rec-cl-tab');
  const elPh  = $('rec-ph-kg');
  const elAnt = $('rec-anti');
  if (elGr)  elGr.textContent  = rec.cl_granule_gr + ' gr';
  if (elTab) elTab.textContent = rec.cl_tablete + ' buc';
  if (elPh)  elPh.textContent  = rec.ph_kg + ' kg';
  if (elAnt) elAnt.textContent = rec.antialgic_l + ' L';

  // Show extrapolation warning if values were outside rule ranges
  const extEl = $('rec-extrapolation');
  if (extEl) {
    if (rec._extrapolated) {
      const parts = [];
      if (vol < 30 || vol > 200) parts.push('volum ' + vol + 'm³ (reguli: 30-200)');
      if (rec._phClamped) parts.push('pH ' + ph + ' (reguli: 7.0-8.5)');
      extEl.style.display = '';
      extEl.innerHTML = '⚠️ <em>Valori extrapolate</em> — ' + parts.join(', ') + '. Dozele sunt estimate.';
    } else {
      extEl.style.display = 'none';
    }
  }

  updateSaveButton();
}

function updateSaveButton() {
  const btn = $('btn-save');
  if (!btn) return;
  const cl = $('m-chlorine') ? $('m-chlorine').value : '';
  const ph = $('m-ph')       ? $('m-ph').value       : '';
  btn.disabled = !cl || !ph;
}

// ── Cl Granule unit toggle ────────────────────────────────────
function toggleClGranUnit(unit) {
  APP.clGranUnit = unit;
  const unitGr = $('unit-gr');
  const unitKg = $('unit-kg');
  if (unitGr) unitGr.classList.toggle('active', unit === 'gr');
  if (unitKg) unitKg.classList.toggle('active', unit === 'kg');
  updateTabConvHint();
}

function getClGranInGrams() {
  const raw = parseFloat($('t-cl-granule') ? $('t-cl-granule').value : '0') || 0;
  return APP.clGranUnit === 'kg' ? raw * 1000 : raw;
}

function updateTabConvHint() {
  const hint = $('tab-conv-hint');
  if (!hint) return;
  const tabCount = parseInt($('t-cl-tablete') ? $('t-cl-tablete').value : '0') || 0;
  if (tabCount > 0) {
    hint.textContent = tabCount + ' tablete = ' + (tabCount * GRAMS_PER_TABLET) + ' gr Cl granule';
  } else {
    hint.textContent = '1 tabletă = ' + GRAMS_PER_TABLET + ' gr Cl granule';
  }
}

// ── Steppers ──────────────────────────────────────────────────
// delta = exact amount to add/subtract (already matches step size)
function stepperChange(inputId, delta) {
  const el = $(inputId);
  if (!el) return;
  const min = parseFloat(el.min) || 0;
  let val = (parseFloat(el.value) || 0) + delta;
  if (val < min) val = min;
  val = Math.round(val * 100) / 100;
  el.value = val;
  if (inputId === 't-cl-tablete') updateTabConvHint();
}

// ── Validation ────────────────────────────────────────────────
function validateInterventionForm() {
  let valid = true;
  const required = [
    { id: 'm-chlorine', label: 'Clor măsurat' },
    { id: 'm-ph',       label: 'pH măsurat' }
  ];

  required.forEach(field => {
    const el = $(field.id);
    if (!el) return;
    const val = el.value.trim();
    if (!val || isNaN(parseFloat(val))) {
      el.classList.add('error');
      valid = false;
    } else {
      el.classList.remove('error');
    }
  });

  if (!valid) {
    showToast('Completați clorul și pH-ul măsurate.', 'error');
    // Scroll to first error
    const firstError = $q('.measure-item input.error');
    if (firstError) firstError.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  return valid;
}

// ── Save Intervention ─────────────────────────────────────────
function showConfirmModal() {
  if (!validateInterventionForm()) return;
  const modal = $('modal-confirm');
  if (modal) modal.classList.add('open');
}

function closeConfirmModal() {
  const modal = $('modal-confirm');
  if (modal) modal.classList.remove('open');
}

async function doSaveIntervention() {
  closeConfirmModal();

  const client = APP.selectedClient;
  if (!client || !APP.user) return;

  const departureTime  = new Date().toISOString();

  const vol = client.pool_volume_mc;
  const cl  = parseFloat($('m-chlorine').value) || null;
  const ph  = parseFloat($('m-ph').value)       || null;
  const rec = (cl !== null && ph !== null) ? getRecommendation(vol, cl, ph) : null;

  const intervention = {
    intervention_id:  uid(),
    client_id:        client.client_id,
    client_name:      client.name,
    technician_id:    APP.user.technician_id,
    technician_name:  APP.user.name,
    date:             APP._interventionDate || new Date().toISOString().split('T')[0],
    created_at:       departureTime,

    measured_chlorine:   cl,
    measured_ph:         ph,
    measured_temp:       parseFloat($('m-temp')        ? $('m-temp').value        : '') || null,
    measured_hardness:   parseFloat($('m-hardness')    ? $('m-hardness').value    : '') || null,
    measured_alkalinity: parseFloat($('m-alkalinity')  ? $('m-alkalinity').value  : '') || null,
    measured_salinity:   parseFloat($('m-salinity')    ? $('m-salinity').value    : '') || null,
    measured_tc:         parseFloat($('m-tc')          ? $('m-tc').value          : '') || null,
    measured_cya:        parseFloat($('m-cya')         ? $('m-cya').value         : '') || null,

    rec_cl_gr:    rec ? rec.cl_granule_gr : null,
    rec_cl_tab:   rec ? rec.cl_tablete    : null,
    rec_ph_kg:    rec ? rec.ph_kg         : null,
    rec_anti_l:   rec ? rec.antialgic_l   : null,

    observations: $('observations') ? $('observations').value.trim() : '',
    operations: (function() {
      var ops = [];
      var checkboxes = document.querySelectorAll('.ops-checkbox');
      checkboxes.forEach(function(chk) {
        if (chk.checked) ops.push(chk.value);
      });
      return ops;
    })(),
    photos:       [...APP.currentPhotos],
    synced:       false,
    // Preserve the linked voice-note recording (if this intervention started as a quick voice note)
    audio_file_url: APP._editingIntervention ? (APP._editingIntervention.audio_file_url || null) : null
  };

  // Dynamic treatment fields from stock products
  const products = APP._stockProducts.length ? APP._stockProducts : await getAllStock();
  products.forEach(p => {
    const el = $('t-' + p.product_id);
    intervention['treat_' + p.product_id] = el ? (parseFloat(el.value) || 0) : 0;
  });

  try {
    // If same client already has interventions on this date, remove ALL old ones
    var existingOld = APP.interventions.filter(function(i) {
      return String(i.client_id) === String(intervention.client_id)
        && i.date === intervention.date
        && String(i.technician_id) === String(intervention.technician_id);
    });
    for (var oi = 0; oi < existingOld.length; oi++) {
      var oldIntv = existingOld[oi];
      try { await deleteRecord('interventions', oldIntv.intervention_id); } catch(e) {}
      await _trackDeletedIntervention(oldIntv.intervention_id);
      if (isSyncConfigured()) {
        try {
          await apiFetch(SYNC_CONFIG.API_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'push', type: 'delete_intervention', data: { intervention_id: oldIntv.intervention_id } })
          });
        } catch(e) { console.warn('[SYNC] Duplicate delete push failed:', e.message); }
      }
    }
    if (existingOld.length > 0) {
      var oldIds = {};
      existingOld.forEach(function(o) { oldIds[o.intervention_id] = true; });
      APP.interventions = APP.interventions.filter(function(i) { return !oldIds[i.intervention_id]; });
    }

    await saveIntervention(intervention);
    APP.interventions.push(intervention);
    APP.pendingSync++;
    APP.lastIntervention = intervention;   // for share report

    // Deduct consumed products from stock
    deductStockForIntervention(intervention).catch(e => console.warn('[STOCK] Deduction error:', e));

    // Check billing notification
    checkBillingAlert(client);

    // Show success screen
    const clientEl = $('success-client-name');
    if (clientEl) clientEl.textContent = client.name;

    if (isAdmin()) {
      // Admin: show duration + share buttons
      const durEl = $('success-duration');
      if (durEl) {
        durEl.textContent = durationMin !== null ? '⏱ Durată: ' + durationMin + ' min' : '';
        durEl.style.display = durationMin !== null ? '' : 'none';
      }
      const waBtn = $('btn-share-wa');
      if (waBtn) waBtn.style.display = client.phone ? '' : 'none';
      const hint = $('success-share-hint');
      if (hint) { hint.style.display = 'none'; hint.textContent = ''; }
      showScreen('success');
      showToast('Intervenție salvată cu succes!', 'success');
      APP._editingIntervention = null;
      // Auto-return to dashboard after 1s
      setTimeout(async () => {
        await loadData();
        renderDashboard();
        showScreen('dashboard');
      }, 1000);
    } else {
      // Tehnicien: ecran simplu, auto-dismiss după 1s
      const durEl = $('success-duration');
      if (durEl) durEl.style.display = 'none';
      const waBtn = $('btn-share-wa');
      if (waBtn) waBtn.style.display = 'none';
      const copyBtn = $('btn-share-copy');
      if (copyBtn) copyBtn.style.display = 'none';
      showScreen('success');
      showToast('✓ Intervenție salvată!', 'success');
      APP._editingIntervention = null;
      setTimeout(async () => {
        await loadData();
        renderDashboard();
        showScreen('dashboard');
        // Restore share buttons for next session
        if (waBtn) waBtn.style.display = '';
        if (copyBtn) copyBtn.style.display = '';
      }, 1000);
    }

    // Setup success back button
    const successBtn = $('btn-success-back');
    if (successBtn) {
      successBtn.onclick = async () => {
        await loadData();
        renderDashboard();
        showScreen('dashboard');
      };
    }

    // Trigger sync
    forceSync().catch(() => {});
    updateSyncBadge();
    showLocalNotification('Intervenție salvată', client.name + ' — ' + intervention.date);
  } catch (err) {
    showToast('Eroare la salvare: ' + err.message, 'error');
  }
}

// ── Previous Interventions ────────────────────────────────────
function renderPreviousInterventions(client) {
  const container = $('prev-interventions');
  if (!container) return;

  const ci = APP.interventions.filter(i => i.client_id === client.client_id && i.date)
    .map(function(i) {
      var raw = String(i.date || '');
      if (raw && !/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        var dp = new Date(raw);
        if (!isNaN(dp.getTime())) i.date = dp.getFullYear() + '-' + ('0'+(dp.getMonth()+1)).slice(-2) + '-' + ('0'+dp.getDate()).slice(-2);
      }
      return i;
    })
    .sort((a, b) => {
      var cmp = String(b.date || '').localeCompare(String(a.date || ''));
      if (cmp !== 0) return cmp;
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    })
    .slice(0, 5);

  if (!ci.length) {
    container.innerHTML = '<p style="padding:12px;color:var(--slate-400);font-size:.85rem">Nicio intervenție anterioară.</p>';
    return;
  }

  container.innerHTML = ci.map(i => {
    const dur = i.duration_minutes != null ? `<span class="prev-int-duration">⏱ ${Math.round(i.duration_minutes)} min</span>` : '';
    return `<div class="prev-intervention" style="cursor:pointer" onclick="showInterventionDetails('${i.intervention_id}')">
      <div class="prev-int-header">
        <span class="prev-int-date">${fmtDate(i.date)}</span>
        ${dur}
      </div>
      <div class="prev-int-tech">👤 ${escHtml(i.technician_name || '')}</div>
      <div class="prev-int-measures">
        <span class="prev-measure">Cl: <strong>${i.measured_chlorine ?? '—'}</strong></span>
        <span class="prev-measure">pH: <strong>${i.measured_ph ?? '—'}</strong></span>
        <span class="prev-measure">T°: <strong>${i.measured_temp ?? '—'}</strong></span>
      </div>
      ${i.observations ? `<div style="margin-top:6px;font-size:.78rem;color:var(--slate-500)">${escHtml(i.observations.substring(0,80))}${i.observations.length > 80 ? '...' : ''}</div>` : ''}
    </div>`;
  }).join('');
}

// ── Client Details Modal ──────────────────────────────────────
async function showClientDetails(clientId) {
  try {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;

  const modal = $('modal-client');
  const body  = $('modal-client-body');
  if (!modal || !body) return;

  // Re-fetch fresh from IndexedDB (not APP.interventions cache)
  const allFromDb = await getAll('interventions');
  // Update in-memory cache too
  APP.interventions = allFromDb;
  APP.pendingSync = allFromDb.filter(i => !i.synced).length;

  const hasLocation = client.location_set && client.latitude;

  // Filter, normalize dates, and sort descending (newest first)
  const ci = allFromDb.filter(i => String(i.client_id) === String(clientId) && i.date)
    .map(function(i) {
      // Normalize date to YYYY-MM-DD (fixes GAS Date objects stored as "Tue Mar 18 2026...")
      var raw = String(i.date || '');
      if (raw && !/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        var dp = new Date(raw);
        if (!isNaN(dp.getTime())) i.date = dp.getFullYear() + '-' + ('0'+(dp.getMonth()+1)).slice(-2) + '-' + ('0'+dp.getDate()).slice(-2);
      }
      return i;
    })
    .sort((a, b) => {
      var cmp = String(b.date || '').localeCompare(String(a.date || ''));
      if (cmp !== 0) return cmp;
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    });

  body.innerHTML = `
    <div class="client-detail-section">
      <h4>Informații</h4>
      <div class="client-detail-row"><span class="detail-label">Volum piscină</span><span class="detail-value">${client.pool_volume_mc} m³</span></div>
      <div class="client-detail-row"><span class="detail-label">Tip</span><span class="detail-value">${client.pool_type}</span></div>
      ${client.phone ? `<div class="client-detail-row"><span class="detail-label">Telefon</span><span class="detail-value">${escHtml(client.phone)}</span></div>` : ''}
      ${client.address ? `<div class="client-detail-row"><span class="detail-label">Adresă</span><span class="detail-value">${escHtml(client.address)}</span></div>` : ''}
      <div class="client-detail-row" style="flex-direction:column;align-items:flex-start;gap:6px">
        <span class="detail-label">Locație GPS <span id="client-detail-gps-status">${hasLocation ? '✅ Setată' : '❌ Nesetată'}</span></span>
        <span class="detail-value" style="display:flex;flex-wrap:wrap;gap:8px;width:100%">
          <button id="client-detail-gps-update-btn" class="client-action-btn" style="flex:0 0 auto;padding:4px 10px;font-size:.78rem" onclick="event.stopPropagation(); setClientLocation('${clientId}')">📍 ${hasLocation ? 'Actualizează' : 'Adaugă'}</button>
          <button id="client-detail-gps-delete-btn" class="client-action-btn" style="flex:0 0 auto;padding:4px 10px;font-size:.78rem;color:var(--danger);display:${hasLocation ? '' : 'none'}" onclick="event.stopPropagation(); deleteClientLocation('${clientId}')">🗑️ Șterge</button>
        </span>
      </div>
      ${client.notes ? `<div class="client-detail-row"><span class="detail-label">Note</span><span class="detail-value">${escHtml(client.notes)}</span></div>` : ''}
    </div>
    <div class="client-detail-section" id="history-section">
      <h4>Istoric intervenții (${ci.length})</h4>
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">
        <label style="font-size:.8rem;color:var(--text-secondary)">Din data:</label>
        <input type="date" id="history-date-filter" onchange="filterHistoryByDate('${clientId}')" style="font-size:.8rem;padding:4px 8px;border:1px solid var(--slate-300);border-radius:6px;background:var(--bg-primary);color:var(--text-primary)">
        <button onclick="document.getElementById('history-date-filter').value='';filterHistoryByDate('${clientId}')" style="font-size:.7rem;padding:3px 8px;border:1px solid var(--slate-300);border-radius:6px;background:var(--bg-secondary);color:var(--text-secondary);cursor:pointer">Toate</button>
      </div>
      <div id="history-list"></div>
    </div>
    ${ci.length >= 2 ? `
    <div class="client-detail-section">
      <h4>Evoluție Cl / pH (ultimele 10)</h4>
      <div class="chart-container">
        <div class="chart-legend">
          <span style="color:#3b82f6;font-weight:600">▬ Cl (mg/L)</span>
          &nbsp;&nbsp;
          <span style="color:#10b981;font-weight:600">▬ pH</span>
        </div>
        <canvas id="params-chart" width="320" height="160" style="width:100%;height:160px"></canvas>
      </div>
    </div>` : ''}
  `;

  $('modal-client-title').textContent = client.name;
  modal.classList.add('open');

  // Render history list
  _renderHistoryList(clientId, ci);

  if (ci.length >= 2) {
    requestAnimationFrame(() => drawParamsChart(clientId));
  }

  // Billing: show "Marchează facturat" button if threshold configured + reached
  APP._billingClientId = clientId;
  const billBtn = $('btn-mark-billed');
  if (billBtn && isAdmin()) {
    const interval = client.billing_interval_interventions;
    if (interval && interval > 0) {
      const since = client.last_billing_date || '1970-01-01';
      const countSince = APP.interventions.filter(i =>
        i.client_id === clientId && i.date > since
      ).length;
      billBtn.style.display = countSince >= interval ? '' : 'none';
      billBtn.textContent = `💰 Marchează facturat (${countSince}/${interval})`;
    } else {
      billBtn.style.display = 'none';
    }
  }

  } catch(e) {
    console.error('[showClientDetails] Error:', e.message);
    showToast('Eroare la deschidere info: ' + e.message, 'error');
  }
}

function closeClientModal() {
  const modal = $('modal-client');
  if (modal) modal.classList.remove('open');
}


// ── Prices Settings UI (dynamic, based on stock products) ────
async function openPricesSettings() {
  var prices = await getExportPrices();
  var stockProducts = await getAllStock();
  var modal = $('modal-prices');
  if (!modal) return;

  // Fixed field: preț intervenție
  var html = '<div style="margin-bottom:12px"><label style="font-size:.78rem;font-weight:700;color:var(--text-secondary)">Pre\u021B interven\u021Bie (RON)</label>';
  html += '<input type="number" id="price-pret_interventie" class="form-input" style="width:100%" step="0.5" value="' + (prices.pret_interventie || 250) + '"></div>';

  // Dynamic fields from stock products
  html += '<p style="font-size:.78rem;font-weight:700;color:var(--text-secondary);margin:8px 0 6px">Pre\u021Buri chimicale:</p>';
  html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">';
  stockProducts.forEach(function(p) {
    var label = escHtml(p.name) + ' (RON/' + escHtml(p.unit || 'buc') + ')';
    html += '<div><label style="font-size:.72rem;font-weight:600;color:var(--text-secondary)">' + label + '</label>';
    html += '<input type="number" id="price-' + p.product_id + '" class="form-input" style="width:100%" step="0.1" value="' + (prices[p.product_id] || 0) + '"></div>';
  });
  html += '</div>';
  if (!stockProducts.length) {
    html += '<p style="font-size:.82rem;color:var(--slate-400)">Niciun produs \u00EEn stoc. Ad\u0103uga\u021Bi produse din Set\u0103ri \u2192 Stoc.</p>';
  }

  $('modal-prices-body').innerHTML = html;
  // Store product IDs for save
  modal._stockProductIds = stockProducts.map(function(p) { return p.product_id; });
  modal.classList.add('open');
}

async function savePricesSettings() {
  var prices = {};
  // Fixed field
  var pretEl = $('price-pret_interventie');
  prices.pret_interventie = pretEl ? (parseFloat(pretEl.value) || 250) : 250;

  // Dynamic product prices
  var modal = $('modal-prices');
  var ids = (modal && modal._stockProductIds) || [];
  ids.forEach(function(pid) {
    var el = $('price-' + pid);
    if (el) prices[pid] = parseFloat(el.value) || 0;
  });

  await saveExportPrices(prices);
  if (modal) modal.classList.remove('open');
  showToast('Preturi salvate!', 'success');
}

/** Show export format choice dialog. */
function showExportChoice() {
  return new Promise(function(resolve) {
    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay open';
    overlay.style.zIndex = '300';
    overlay.innerHTML = '<div class="modal-sheet" style="max-width:340px;margin:auto;border-radius:16px">' +
      '<div class="modal-handle"></div>' +
      '<div class="modal-title">Alege formatul export</div>' +
      '<div style="display:flex;flex-direction:column;gap:8px;padding:0 16px 16px">' +
      '<button class="btn-primary" style="padding:12px" data-ch="standard">Raport Standard</button>' +
      '<button class="btn-primary" style="padding:12px;background:var(--blue-600)" data-ch="chimicale">Deviz Chimicale</button>' +
      '<button class="btn-primary" style="padding:12px;background:#16a34a" data-ch="complet">Deviz Complet (+ Operatiuni)</button>' +
      '<button class="btn-modal-cancel" data-ch="">Anuleaza</button>' +
      '</div></div>';
    overlay.addEventListener('click', function(e) {
      var ch = e.target.dataset.ch;
      if (ch !== undefined || e.target === overlay) {
        overlay.remove();
        resolve(ch || '');
      }
    });
    document.body.appendChild(overlay);
  });
}


/** Show export filter dialog — choose interval only */
function showExportFilter(client, allInterventions) {
  return new Promise(function(resolve) {
    var sorted = allInterventions.slice().sort(function(a,b) { return b.date.localeCompare(a.date); });
    var defaultCount = Math.min(4, sorted.length);
    var defaultFrom = sorted.length >= 4 ? sorted[3].date : (sorted.length ? sorted[sorted.length - 1].date : '');

    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay open';
    overlay.style.zIndex = '300';
    overlay.innerHTML = '<div class="modal-sheet" style="max-width:400px;margin:auto;border-radius:16px">' +
      '<div class="modal-handle"></div>' +
      '<div class="modal-title">Export Deviz ' + escHtml(client.name) + '</div>' +
      '<div style="padding:0 16px 16px">' +
        '<p style="font-size:.82rem;color:var(--text-secondary);margin:0 0 4px">' + sorted.length + ' interventii disponibile</p>' +
        '<p style="font-size:.78rem;color:var(--text-secondary);margin:0 0 12px">Format: <strong>' + (parseInt(client.deviz_type) === 2 ? 'V2 (Chimicale + Operatiuni)' : 'V1 (Chimicale)') + '</strong></p>' +

        '<div style="font-size:.78rem;font-weight:600;color:var(--text-secondary);margin:0 0 6px;text-transform:uppercase">Interval</div>' +
        '<div style="display:flex;flex-direction:column;gap:8px;margin-bottom:16px">' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-filter" value="last" checked style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">Ultimele</span>' +
            '<input type="number" id="exp-last-n" value="' + defaultCount + '" min="1" max="1000" style="width:56px;padding:5px;border:1px solid var(--slate-200);border-radius:6px;text-align:center;font-size:.9rem">' +
            '<span style="font-size:.88rem">interventii</span>' +
          '</label>' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-filter" value="date" style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">De la data:</span>' +
            '<input type="date" id="exp-from-date" value="' + defaultFrom + '" style="padding:5px;border:1px solid var(--slate-200);border-radius:6px;font-size:.88rem;flex:1">' +
          '</label>' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-filter" value="all" style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">Toate interventiile</span>' +
          '</label>' +
        '</div>' +

        '<div style="display:flex;gap:8px">' +
          '<button class="btn-modal-cancel" style="flex:1" data-action="cancel">Anuleaza</button>' +
          '<button class="btn-modal-confirm" style="flex:1" data-action="export">Exporta</button>' +
        '</div>' +
      '</div></div>';

    var lastN = overlay.querySelector('#exp-last-n');
    var fromDate = overlay.querySelector('#exp-from-date');
    if (lastN) lastN.onfocus = function() { overlay.querySelector('input[value="last"]').checked = true; };
    if (fromDate) fromDate.onfocus = function() { overlay.querySelector('input[value="date"]').checked = true; };

    overlay.addEventListener('click', function(e) {
      var action = e.target.dataset.action;
      if (action === 'cancel' || e.target === overlay) {
        overlay.remove();
        resolve(null);
        return;
      }
      if (action === 'export') {
        var mode = overlay.querySelector('input[name="exp-filter"]:checked').value;
        var filtered;
        if (mode === 'last') {
          var n = parseInt(lastN.value) || 4;
          filtered = sorted.slice(0, n);
        } else if (mode === 'date') {
          var from = fromDate.value;
          filtered = sorted.filter(function(i) { return i.date >= from; });
        } else {
          filtered = sorted;
        }
        overlay.remove();
        resolve(filtered);
      }
    });

    document.body.appendChild(overlay);
  });
}

// ── Export Modal ──────────────────────────────────────────────
function showExportModal(clientId) {
  // Per-client export: go directly to filter+format dialog
  if (clientId) {
    _exportClientDirect(clientId);
    return;
  }

  // All-clients export: show format choice dialog
  _exportAllDirect();
}

async function _exportAllDirect() {
  try {
    await loadData();
    var totalInt = APP.interventions.length;
    if (!totalInt) { showToast('Nicio interventie de exportat.', 'warning'); return; }

    // Show interval filter dialog (same as per-client, but for all)
    var filterResult = await _showAllExportFilter(totalInt);
    if (!filterResult) return;

    showToast('Generare Excel...', 'info');
    // Apply filter to each client's interventions
    await exportAllDevizMixed(APP.clients, APP.interventions, filterResult);
    showToast('Export complet!', 'success');
  } catch(e) {
    showToast('Eroare export: ' + e.message, 'error');
  }
}

function _showAllExportFilter(totalCount) {
  return new Promise(function(resolve) {
    var today = new Date().toISOString().split('T')[0];
    // Default from date: 3 months ago
    var d = new Date();
    d.setMonth(d.getMonth() - 3);
    var defaultFrom = d.toISOString().split('T')[0];

    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay open';
    overlay.style.zIndex = '300';
    overlay.innerHTML = '<div class="modal-sheet" style="max-width:400px;margin:auto;border-radius:16px">' +
      '<div class="modal-handle"></div>' +
      '<div class="modal-title">Export Toti Clientii</div>' +
      '<div style="padding:0 16px 16px">' +
        '<p style="font-size:.82rem;color:var(--text-secondary);margin:0 0 12px">' + totalCount + ' interventii totale. Formatul deviz este cel setat pe fiecare client.</p>' +

        '<div style="font-size:.78rem;font-weight:600;color:var(--text-secondary);margin:0 0 6px;text-transform:uppercase">Interval</div>' +
        '<div style="display:flex;flex-direction:column;gap:8px;margin-bottom:16px">' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-all-filter" value="last" checked style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">Ultimele</span>' +
            '<input type="number" id="exp-all-last-n" value="4" min="1" max="999" style="width:56px;padding:5px;border:1px solid var(--slate-200);border-radius:6px;text-align:center;font-size:.9rem">' +
            '<span style="font-size:.88rem">interventii / client</span>' +
          '</label>' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-all-filter" value="date" style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">De la data:</span>' +
            '<input type="date" id="exp-all-from-date" value="' + defaultFrom + '" style="padding:5px;border:1px solid var(--slate-200);border-radius:6px;font-size:.88rem;flex:1">' +
          '</label>' +
          '<label style="display:flex;align-items:center;gap:8px;cursor:pointer">' +
            '<input type="radio" name="exp-all-filter" value="all" style="accent-color:var(--primary)">' +
            '<span style="font-size:.88rem">Toate interventiile</span>' +
          '</label>' +
        '</div>' +

        '<div style="display:flex;gap:8px">' +
          '<button class="btn-modal-cancel" style="flex:1" data-action="cancel">Anuleaza</button>' +
          '<button class="btn-modal-confirm" style="flex:1" data-action="export">Exporta</button>' +
        '</div>' +
      '</div></div>';

    var lastN = overlay.querySelector('#exp-all-last-n');
    var fromDate = overlay.querySelector('#exp-all-from-date');
    if (lastN) lastN.onfocus = function() { overlay.querySelector('input[value="last"]').checked = true; };
    if (fromDate) fromDate.onfocus = function() { overlay.querySelector('input[value="date"]').checked = true; };

    overlay.addEventListener('click', function(e) {
      var action = e.target.dataset.action;
      if (action === 'cancel' || e.target === overlay) {
        overlay.remove();
        resolve(null);
        return;
      }
      if (action === 'export') {
        var mode = overlay.querySelector('input[name="exp-all-filter"]:checked').value;
        overlay.remove();
        resolve({ mode: mode, lastN: parseInt(lastN.value) || 4, fromDate: fromDate.value });
      }
    });

    document.body.appendChild(overlay);
  });
}

async function _exportClientDirect(clientId) {
  try {
    await loadData();
    var client = APP.clients.find(function(c) { return c.client_id === clientId; });
    if (!client) { showToast('Client negasit.', 'error'); return; }
    var allCi = APP.interventions.filter(function(i) { return i.client_id === clientId; });
    if (!allCi.length) { showToast('Nicio interventie pentru acest client.', 'warning'); return; }

    var filtered = await showExportFilter(client, allCi);
    if (!filtered) return;
    if (!filtered.length) {
      showToast('Nicio interventie in intervalul selectat.', 'warning');
      return;
    }

    showToast('Generare Excel...', 'info');
    var devizType = parseInt(client.deviz_type) || 2;
    if (devizType === 2) {
      await exportDevizComplet(client, filtered);
    } else {
      await exportDevizChimicale(client, filtered);
    }
    showToast('Export complet!', 'success');
  } catch(e) {
    if (e.message) showToast('Eroare export: ' + e.message, 'error');
  }
}

function closeExportModal() {
  const modal = $('modal-export');
  if (modal) modal.classList.remove('open');
}

// ── Photo Capture ─────────────────────────────────────────────
function setupPhotoCapture() {
  const addBtn   = $('btn-add-photo');
  const fileInput = $('photo-input');
  if (!addBtn || !fileInput) return;

  addBtn.onclick = () => {
    if (APP.currentPhotos.length >= 4) {
      showToast('Maximum 4 fotografii per intervenție.', 'warning');
      return;
    }
    fileInput.click();
  };

  fileInput.onchange = e => {
    const files = Array.from(e.target.files);
    files.forEach(file => {
      if (APP.currentPhotos.length >= 4) return;
      const reader = new FileReader();
      reader.onload = re => {
        resizeImage(re.result, 800, dataUrl => {
          APP.currentPhotos.push(dataUrl);
          renderPhotoGrid();
        });
      };
      reader.readAsDataURL(file);
    });
    fileInput.value = '';
  };
}

function resizeImage(dataUrl, maxSize, callback) {
  const img = new Image();
  img.onload = () => {
    let w = img.width, h = img.height;
    if (w > maxSize || h > maxSize) {
      if (w > h) { h = Math.round(h * maxSize / w); w = maxSize; }
      else       { w = Math.round(w * maxSize / h); h = maxSize; }
    }
    const canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = h;
    canvas.getContext('2d').drawImage(img, 0, 0, w, h);
    callback(canvas.toDataURL('image/jpeg', 0.72));
  };
  img.src = dataUrl;
}

function renderPhotoGrid() {
  const grid = $('photo-grid');
  if (!grid) return;
  grid.innerHTML = APP.currentPhotos.map((dataUrl, idx) => `
    <div class="photo-thumb">
      <img src="${dataUrl}" alt="Foto ${idx + 1}">
      <button class="photo-remove" onclick="removePhoto(${idx})" title="Șterge">✕</button>
    </div>
  `).join('');

  const addBtn = $('btn-add-photo');
  if (addBtn) addBtn.style.display = APP.currentPhotos.length >= 4 ? 'none' : '';

  // Photo count indicator
  let indicator = $('photo-count-indicator');
  if (!indicator) {
    indicator = document.createElement('div');
    indicator.id = 'photo-count-indicator';
    indicator.className = 'photo-count-indicator';
    const parent = grid.parentElement;
    if (parent) parent.appendChild(indicator);
  }
  const n = APP.currentPhotos.length;
  if (n > 0) {
    indicator.innerHTML = '<span class="photo-check">✓</span> ' + n + ' foto' + (n > 1 ? 'grafii' : 'grafie') + ' adăugat' + (n > 1 ? 'e' : 'ă');
    indicator.style.display = '';
  } else {
    indicator.style.display = 'none';
  }
}

function removePhoto(idx) {
  APP.currentPhotos.splice(idx, 1);
  renderPhotoGrid();
}

// ── GPS Helpers ───────────────────────────────────────────────
function getCurrentPosition() {
  return new Promise(resolve => {
    if (!navigator.geolocation) { resolve(null); return; }
    navigator.geolocation.getCurrentPosition(
      pos => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude, accuracy: pos.coords.accuracy }),
      err => { console.warn('[GEO] Error:', err.message); resolve(null); },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 }
    );
  });
}

function haversineDistance(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) ** 2 +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLng/2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ── Notifications ─────────────────────────────────────────────
function setupNotifications() {
  if (!('Notification' in window)) return;
  if (Notification.permission === 'default') {
    Notification.requestPermission();
  }
}

function showLocalNotification(title, body) {
  if (!('Notification' in window) || Notification.permission !== 'granted') return;
  try { new Notification(title, { body, icon: './icons/icon-192.png' }); } catch {}
}

// ── Utility ───────────────────────────────────────────────────
function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fmtDate(isoDate) {
  if (!isoDate) return '';
  // Handle both YYYY-MM-DD and full Date strings
  var d;
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
    d = new Date(isoDate + 'T12:00:00');
  } else {
    d = new Date(isoDate);
  }
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('ro-RO', { day: '2-digit', month: 'long', year: 'numeric' });
}

// ════════════════════════════════════════════════════════════════
// FEATURE 1 — Dark Mode
// ════════════════════════════════════════════════════════════════
function toggleDarkMode() {
  const isDark = document.body.classList.toggle('dark-mode');
  localStorage.setItem('darkMode', isDark ? '1' : '');
  const btn = $('btn-dark-mode');
  if (btn) btn.textContent = isDark ? '☀️ Mod Normal' : '🌙 Toggle Dark Mode';
}

// ════════════════════════════════════════════════════════════════
// FEATURE 2 — Dashboard Tabs
// ════════════════════════════════════════════════════════════════
function switchTab(tab) {
  APP.dashboardTab = tab;
  // Mark the all-tab button active when tab='all', due-tab button when tab='due'
  const allBtn = $q('.tab-btn:not(#tab-due-btn):not([onclick*="showAddClientModal"])');
  const dueBtn = $('tab-due-btn');
  if (allBtn) allBtn.classList.toggle('active', tab === 'all');
  if (dueBtn) dueBtn.classList.toggle('active', tab === 'due');
  renderClientList($('search-input') ? $('search-input').value : '');
}

// ── Manual Sync ──────────────────────────────────────────────
async function manualSync() {
  var btn = $('btn-manual-sync');
  var icon = $('sync-icon');
  var text = null; // sync-text removed, button is icon-only
  if (btn) btn.disabled = true;
  if (icon) icon.style.animation = 'spin 1s linear infinite';
  if (text) text.textContent = 'Se sincronizeaza...';

  try {
    await forceSync();
    await loadData();
    renderDashboard();
    showToast('Sincronizare completa!', 'success');
  } catch (e) {
    showToast('Eroare sincronizare: ' + e.message, 'error');
  }

  if (btn) btn.disabled = false;
  if (icon) icon.style.animation = '';
  if (text) text.textContent = 'Sincronizeaza';
}

// ════════════════════════════════════════════════════════════════
// FEATURE 3 — Urgency Level
// ════════════════════════════════════════════════════════════════
function getUrgencyLevel(client) {
  const freq = client.visit_frequency_days;
  if (!freq) return 'ok';
  const ci = APP.interventions.filter(i => i.client_id === client.client_id);
  if (!ci.length) return 'never';
  const lastDate = ci.sort((a, b) => b.date.localeCompare(a.date))[0].date;
  const days = Math.floor((Date.now() - Date.parse(lastDate)) / 86400000);
  if (days > freq)          return 'overdue';
  if (days > freq * 0.8)    return 'soon';
  return 'ok';
}

// ════════════════════════════════════════════════════════════════
// FEATURE 4 — Observation Templates
// ════════════════════════════════════════════════════════════════
const OBS_TEMPLATES = [
  'Apă limpede, filtrare OK.',
  'Alge depistate pe pereți — antialgic adăugat.',
  'Pompă curățată și verificată.',
  'pH stabilizat după tratament.',
  'Saltwater system OK.',
  'Filtru spălat contracurent.',
  'Prima vizită — situație inițială documentată.',
  'Clor scăzut după weekend ploios.',
];

/** Get observation templates from storage (falls back to built-in defaults) */
async function getObsTemplates() {
  try {
    const stored = await getByKey('settings', 'obs_templates');
    if (stored && Array.isArray(stored.value) && stored.value.length > 0) return stored.value;
  } catch (e) {}
  return [...OBS_TEMPLATES];
}

/** Persist observation templates to storage */
async function saveObsTemplates(arr) {
  await put('settings', { key: 'obs_templates', value: arr });
}


/** Toggle observation suggestions visibility */
function toggleObsSuggestions() {
  var container = $('obs-templates-container');
  var arrow = $('obs-toggle-arrow');
  if (!container) return;
  container.classList.toggle('open');
  if (arrow) {
    arrow.textContent = container.classList.contains('open') ? '▼ sugestii' : '▶ sugestii';
  }
}

async function renderObsTemplates() {
  const container = $('obs-templates-container');
  if (!container) return;
  const templates = await getObsTemplates();
  // IMPORTANT: use data-obs-text attribute to avoid quote conflicts in onclick HTML
  container.innerHTML = templates.map(t =>
    `<button type="button" class="obs-chip" data-obs-text="${escHtml(t)}" onclick="toggleObsChip(this)">${escHtml(t)}</button>`
  ).join('');
}

/** Render obs template list inside Settings panel */
async function renderObsTemplatesSettings() {
  const list = $('obs-templates-settings-list');
  if (!list) return;
  const templates = await getObsTemplates();
  if (!templates.length) {
    list.innerHTML = '<p style="font-size:.8rem;color:var(--slate-400);padding:4px 0">Nicio sugestie. Adaugă una mai jos.</p>';
    return;
  }
  list.innerHTML = templates.map((t, i) =>
    `<div class="obs-tmpl-setting-row">
      <span class="obs-tmpl-setting-text">${escHtml(t)}</span>
      <button class="obs-tmpl-del-btn" onclick="deleteObsTemplate(${i})" title="Șterge">🗑</button>
    </div>`
  ).join('');
}

/** Add a new obs template */
async function addObsTemplate() {
  const input = $('obs-template-new-input');
  const text = input ? input.value.trim() : '';
  if (!text) { showToast('Scrie textul sugestiei.', 'warning'); return; }
  const templates = await getObsTemplates();
  if (templates.includes(text)) { showToast('Sugestia există deja.', 'warning'); return; }
  templates.push(text);
  await saveObsTemplates(templates);
  if (input) input.value = '';
  renderObsTemplatesSettings();
  renderObsTemplates();
  showToast('Sugestie adăugată.', 'success');
}

/** Delete an obs template by index */
async function deleteObsTemplate(index) {
  const templates = await getObsTemplates();
  templates.splice(index, 1);
  await saveObsTemplates(templates);
  renderObsTemplatesSettings();
  renderObsTemplates();
  showToast('Sugestie ștearsă.', 'success');
}

function toggleObsChip(btn) {
  const text = btn.dataset.obsText;
  if (!text) return;
  const ta = $('observations');
  if (!ta) return;
  const isActive = btn.classList.toggle('active');
  if (isActive) {
    const sep = ta.value.trim() ? '. ' : '';
    ta.value = ta.value.trimEnd() + sep + text;
  }
}

// ════════════════════════════════════════════════════════════════
// FEATURE 5 — Backup / Restore DB
// ════════════════════════════════════════════════════════════════
async function exportBackupJSON() {
  const stores = ['clients', 'interventions', 'technicians', 'stock', 'settings'];
  const backup = { version: 3, date: new Date().toISOString(), data: {} };
  for (const s of stores) {
    try { backup.data[s] = await getAll(s); } catch { backup.data[s] = []; }
  }
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `pool-backup-${new Date().toISOString().split('T')[0]}.json`;
  a.click();
  showToast('Backup descărcat.', 'success');
}

async function importBackupJSON(file) {
  if (!file) return;
  try {
    const text = await file.text();
    const backup = JSON.parse(text);
    if (!backup.data) throw new Error('Format invalid');
    for (const [store, items] of Object.entries(backup.data)) {
      try {
        await clearStore(store);
        if (items && items.length) await putMany(store, items);
      } catch (e) { console.warn('[RESTORE] Skipped store', store, e); }
    }
    showToast('Backup restaurat. Se reîncarcă...', 'success');
    setTimeout(() => location.reload(), 1500);
  } catch (e) {
    showToast('Eroare la restaurare: ' + e.message, 'error');
  }
  // Reset file input
  const fi = $('restore-file-input');
  if (fi) fi.value = '';
}

// ════════════════════════════════════════════════════════════════
// FEATURE 6 — Add / Edit Clients
// ════════════════════════════════════════════════════════════════
function showAddClientModal() {
  APP.clientFormMode = 'add';
  $('client-form-title').textContent = 'Adaugă client';
  ['cf-name','cf-phone','cf-address','cf-notes','cf-billing-interval'].forEach(id => { const el = $(id); if (el) el.value = ''; });
  const vol = $('cf-pool-vol');   if (vol) vol.value = '';
  const freq = $('cf-visit-freq'); if (freq) freq.value = '7';
  const type = $('cf-pool-type'); if (type) type.value = 'exterior';
  $('modal-client-form').classList.add('open');
}

function showEditClientModal(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;
  APP.clientFormMode = 'edit';
  APP._editingClientId = clientId;
  $('client-form-title').textContent = 'Editează client';
  const set = (id, val) => { const el = $(id); if (el) el.value = val ?? ''; };
  set('cf-name',       client.name);
  set('cf-phone',      client.phone);
  set('cf-address',    client.address);
  set('cf-pool-vol',   client.pool_volume_mc);
  set('cf-notes',      client.notes);
  set('cf-visit-freq', client.visit_frequency_days || 14);
  set('cf-billing-interval', client.billing_interval_interventions || '');
  var pretEl = $('cf-pret-interventie');
  if (pretEl) pretEl.value = client.pret_interventie || '';
  var devizSel = $('cf-deviz-type');
  if (devizSel) devizSel.value = String(client.deviz_type || 2);
  const type = $('cf-pool-type');
  if (type) type.value = client.pool_type || 'exterior';
  $('modal-client-form').classList.add('open');
}

async function doSaveClientForm() {
  const name = $('cf-name') ? $('cf-name').value.trim() : '';
  if (!name) { showToast('Numele este obligatoriu.', 'error'); return; }

  const now = new Date().toISOString();
  const isEdit = APP.clientFormMode === 'edit';
  const existing = isEdit ? APP.clients.find(c => c.client_id === APP._editingClientId) : null;

  const billingRaw = parseInt($('cf-billing-interval') ? $('cf-billing-interval').value : '0') || 0;
  const data = {
    client_id:           isEdit ? APP._editingClientId : ('c_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
    name,
    phone:               $('cf-phone')    ? $('cf-phone').value.trim()     : '',
    address:             $('cf-address')  ? $('cf-address').value.trim()   : '',
    pool_volume_mc:      parseFloat($('cf-pool-vol') ? $('cf-pool-vol').value : '0') || 0,
    pool_type:           $('cf-pool-type') ? $('cf-pool-type').value       : 'exterior',
    notes:               $('cf-notes')    ? $('cf-notes').value.trim()     : '',
    visit_frequency_days: parseInt($('cf-visit-freq') ? $('cf-visit-freq').value : '7') || 7,
    billing_interval_interventions: billingRaw > 0 ? billingRaw : 4,
    pret_interventie: parseFloat($('cf-pret-interventie') ? $('cf-pret-interventie').value : '0') || 0,
    deviz_type: parseInt($('cf-deviz-type') ? $('cf-deviz-type').value : '2') || 2,
    last_billing_date:   isEdit && existing ? (existing.last_billing_date || null) : null,
    active:              true,
    created_at:          isEdit ? (existing ? existing.created_at : now) : now,
    updated_at:          now,
    // Preserve GPS data if editing
    latitude:    isEdit && existing ? existing.latitude    : null,
    longitude:   isEdit && existing ? existing.longitude   : null,
    location_set: isEdit && existing ? existing.location_set : false
  };

  try {
    await put('clients', data);
    // Push to GAS immediately if configured
    if (isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'clients', data: [data] })
      }).catch(err => console.warn('[SYNC] Client push failed:', err.message));
    }
    await loadData();
    renderDashboard();
    closeClientFormModal();
    showToast(isEdit ? 'Client actualizat.' : 'Client adăugat.', 'success');
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

function closeClientFormModal() {
  const modal = $('modal-client-form');
  if (modal) modal.classList.remove('open');
}

/** Șterge un client (admin only) — marchez inactiv, nu șterge fizic din bază. */
async function deleteClient(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;
  if (!confirm('Ștergi clientul "' + client.name + '"?\nClientul va fi dezactivat (nu șters definitiv).')) return;
  try {
    client.active = false;
    client.updated_at = new Date().toISOString();
    await put('clients', client);
    // Push to GAS if configured
    if (isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'clients', data: [client] })
      }).catch(err => console.warn('[SYNC] Client deactivate push failed:', err.message));
    }
    await loadData();
    renderDashboard();
    showToast('Clientul "' + client.name + '" a fost dezactivat.', 'success');
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

// ════════════════════════════════════════════════════════════════
// FEATURE 7 — Technician Manager
// ════════════════════════════════════════════════════════════════
async function showTechManager() {
  const modal = $('modal-tech-manager');
  const body  = $('tech-manager-body');
  if (!modal || !body) return;

  hideTechForm();
  const techs = await getAll('technicians');
  APP._techList = techs;

  body.innerHTML = techs.length ? techs.map(t => `
    <div class="tech-row">
      <div>
        <strong>${escHtml(t.name)}</strong>
        <span class="tech-role-badge ${t.role === 'admin' ? 'badge-admin' : 'badge-tech'}">${t.role}</span>
        <div style="font-size:.78rem;color:var(--slate-500)">@${escHtml(t.username)}</div>
      </div>
      <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
        <span style="font-size:.8rem;color:${t.active !== false ? 'var(--success)' : 'var(--danger)'}">${t.active !== false ? '● Activ' : '● Inactiv'}</span>
        <button class="client-action-btn" onclick="toggleTechActive('${t.technician_id}')">${t.active !== false ? 'Dezactivează' : 'Activează'}</button>
        <button class="client-action-btn" onclick="showTechForm('${t.technician_id}')">✏️ Editează</button>
        <button class="client-action-btn" style="color:var(--danger)" onclick="deleteTech('${t.technician_id}','${escHtml(t.name)}')">🗑️ Șterge</button>
      </div>
    </div>
  `).join('') : '<p style="padding:12px;color:var(--slate-400)">Niciun tehnician.</p>';

  modal.classList.add('open');
}

function showTechForm(techId) {
  const section = $('tech-form-section');
  if (!section) return;
  section.style.display = '';
  $('tf-id').value       = techId || '';
  $('tf-name').value     = '';
  $('tf-username').value = '';
  $('tf-password').value = '';
  $('tf-role').value     = 'technician';

  if (techId) {
    const t = APP._techList ? APP._techList.find(t => t.technician_id === techId) : null;
    if (t) {
      $('tf-name').value     = t.name     || '';
      $('tf-username').value = t.username || '';
      $('tf-role').value     = t.role     || 'technician';
    }
  }
}

function hideTechForm() {
  const section = $('tech-form-section');
  if (section) section.style.display = 'none';
}

async function doSaveTech() {
  const name     = $('tf-name')     ? $('tf-name').value.trim()     : '';
  const username = $('tf-username') ? $('tf-username').value.trim() : '';
  const password = $('tf-password') ? $('tf-password').value        : '';
  const role     = $('tf-role')     ? $('tf-role').value            : 'technician';
  const existingId = $('tf-id')     ? $('tf-id').value              : '';

  if (!name || !username) { showToast('Numele și username-ul sunt obligatorii.', 'error'); return; }
  if (!existingId && !password) { showToast('Parola este obligatorie pentru cont nou.', 'error'); return; }

  const data = {
    technician_id: existingId || ('t_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)),
    name, username, role, active: true
  };
  if (password) data.password = password;
  else {
    // Keep existing password
    try {
      const existing = await getByKey('technicians', existingId);
      if (existing) data.password = existing.password;
    } catch {}
  }

  try {
    await put('technicians', data);
    showToast('Salvat local: ' + data.username, 'success');
    // Backup all techs to settings for persistence
    try { const all = await getAll('technicians'); await setSetting('technicians_backup', JSON.stringify(all)); } catch(_) {}
    // Refresh list immediately (before GAS push)
    showTechManager();
    // Push to GAS in background (don't block UI)
    if (isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'technicians', data: [data] })
      }).then(function(resp) {
        if (resp && resp.success) {
          showToast('Sincronizat cu serverul ✓', 'success');
        } else {
          console.warn('[SYNC] Technician push response:', resp);
        }
      }).catch(function(pushErr) {
        console.warn('[SYNC] Technician push failed:', pushErr.message);
        showToast('Push la server eșuat. Se va reîncerca automat.', 'warning', 4000);
      });
    }
  } catch (e) {
    showToast('Eroare salvare: ' + e.message, 'error');
  }
}

async function toggleTechActive(techId) {
  try {
    const tech = await getByKey('technicians', techId);
    if (!tech) return;
    tech.active = tech.active === false ? true : false;
    await put('technicians', tech);
    try { const all = await getAll('technicians'); await setSetting('technicians_backup', JSON.stringify(all)); } catch(_) {}
    showTechManager();
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

async function deleteTech(techId, techName) {
  if (!confirm('Sigur vrei să ștergi tehnicianul "' + techName + '"?\n\nAceastă acțiune este ireversibilă.')) return;
  try {
    await deleteRecord('technicians', techId);
    // Track deleted tech IDs so sync pull doesn't re-insert them
    var deletedIds = (await getSetting('deleted_technician_ids')) || [];
    if (deletedIds.indexOf(techId) === -1) deletedIds.push(techId);
    await setSetting('deleted_technician_ids', deletedIds);
    // Push deletion to GAS if configured
    if (isSyncConfigured()) {
      try {
        await apiFetch(SYNC_CONFIG.API_URL, {
          method: 'POST',
          body: JSON.stringify({ action: 'push', type: 'technicians', data: [{ technician_id: techId, _deleted: true }] })
        });
        // GAS deletion successful — remove from local tracking list
        deletedIds = deletedIds.filter(function(id) { return id !== techId; });
        await setSetting('deleted_technician_ids', deletedIds);
      } catch (e) { console.warn('[SYNC] Tech delete push failed:', e.message); }
    }
    try { const all = await getAll('technicians'); await setSetting('technicians_backup', JSON.stringify(all)); } catch(_) {}
    showToast('Tehnicianul "' + techName + '" a fost șters.', 'success');
    showTechManager();
  } catch (e) {
    showToast('Eroare la ștergere: ' + e.message, 'error');
  }
}

function closeTechManager() {
  const modal = $('modal-tech-manager');
  if (modal) modal.classList.remove('open');
}

// ════════════════════════════════════════════════════════════════
// FEATURE 8 — Admin Stats
// ════════════════════════════════════════════════════════════════
function renderAdminStats() {
  const container = $('admin-stats');
  if (!container || !isAdmin()) { if (container) container.innerHTML = ''; return; }

  const now     = Date.now();
  const ms30    = 30 * 24 * 60 * 60 * 1000;
  const recent  = APP.interventions.filter(i => (now - Date.parse(i.date + 'T12:00:00')) <= ms30);

  // Per technician
  const byTech = {};
  recent.forEach(i => {
    byTech[i.technician_name] = (byTech[i.technician_name] || 0) + 1;
  });

  // Total Cl granule consumed (last 30 days)
  const totalCl = recent.reduce((s, i) => s + (i.treat_cl_granule_gr || 0), 0);

  // Average duration
  const withDur = recent.filter(i => i.duration_minutes != null);
  const avgDur  = withDur.length ? Math.round(withDur.reduce((s, i) => s + i.duration_minutes, 0) / withDur.length) : null;

  // Due clients
  const dueCount = APP.clients.filter(c => {
    const u = getUrgencyLevel(c);
    return u === 'overdue' || u === 'never';
  }).length;

  container.innerHTML = `
    <div style="padding:12px 14px 4px;font-size:.85rem;font-weight:700;color:var(--slate-600)">📊 Statistici admin (30 zile)</div>
    <div class="admin-stats-grid">
      <div class="admin-stat-item">
        <div class="admin-stat-value">${recent.length}</div>
        <div class="admin-stat-label">Intervenții totale</div>
      </div>
      <div class="admin-stat-item">
        <div class="admin-stat-value">${(totalCl / 1000).toFixed(1)} kg</div>
        <div class="admin-stat-label">Cl granule consumat</div>
      </div>
      <div class="admin-stat-item">
        <div class="admin-stat-value">${avgDur !== null ? avgDur + ' min' : '—'}</div>
        <div class="admin-stat-label">Durată medie</div>
      </div>
      <div class="admin-stat-item">
        <div class="admin-stat-value" style="color:${dueCount > 0 ? 'var(--danger)' : 'var(--success)'}">${dueCount}</div>
        <div class="admin-stat-label">Clienți de vizitat</div>
      </div>
    </div>
    ${Object.keys(byTech).length ? `
    <div style="padding:0 14px 10px;font-size:.8rem;color:var(--slate-600)">
      ${Object.entries(byTech).sort((a,b)=>b[1]-a[1]).map(([name,n])=>`<span style="margin-right:12px">👤 ${escHtml(name)}: <strong>${n}</strong></span>`).join('')}
    </div>` : ''}
  `;
}

// ════════════════════════════════════════════════════════════════
// FEATURE 9 — Cl/pH Chart (pure canvas)
// ════════════════════════════════════════════════════════════════
function drawParamsChart(clientId) {
  const canvas = $('params-chart');
  if (!canvas) return;

  const W = canvas.offsetWidth || 320;
  const H = 160;
  canvas.width  = W;
  canvas.height = H;

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  const data = APP.interventions
    .filter(i => i.client_id === clientId && i.measured_chlorine != null && i.measured_ph != null)
    .sort((a, b) => a.date.localeCompare(b.date))
    .slice(-10);

  if (data.length < 2) return;

  const PAD = { top: 12, right: 10, bottom: 24, left: 30 };
  const cW = W - PAD.left - PAD.right;
  const cH = H - PAD.top  - PAD.bottom;

  // Scales
  const clMin = 0, clMax = 5;
  const phMin = 6, phMax = 9;

  function xPos(idx) { return PAD.left + (idx / (data.length - 1)) * cW; }
  function clY(v)    { return PAD.top + cH - ((v - clMin) / (clMax - clMin)) * cH; }
  function phY(v)    { return PAD.top + cH - ((v - phMin) / (phMax - phMin)) * cH; }

  // Grid lines
  ctx.strokeStyle = '#e2e8f0';
  ctx.lineWidth = 1;
  [0, 0.25, 0.5, 0.75, 1].forEach(f => {
    const y = PAD.top + f * cH;
    ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(PAD.left + cW, y); ctx.stroke();
  });

  // Y-axis labels
  ctx.fillStyle = '#94a3b8';
  ctx.font = '9px sans-serif';
  ctx.textAlign = 'right';
  [[0,'0'],[2.5,'2.5'],[5,'5']].forEach(([v,l]) => {
    ctx.fillText(l, PAD.left - 4, clY(v) + 3);
  });

  // X-axis date labels
  ctx.textAlign = 'center';
  data.forEach((d, i) => {
    if (i % Math.ceil(data.length / 4) === 0 || i === data.length - 1) {
      const label = d.date.slice(5); // MM-DD
      ctx.fillText(label, xPos(i), H - 6);
    }
  });

  // Draw Cl line (blue)
  ctx.strokeStyle = '#3b82f6';
  ctx.lineWidth = 2;
  ctx.beginPath();
  data.forEach((d, i) => {
    const y = clY(Math.min(clMax, Math.max(clMin, d.measured_chlorine)));
    i === 0 ? ctx.moveTo(xPos(i), y) : ctx.lineTo(xPos(i), y);
  });
  ctx.stroke();

  // Draw pH line (green) — mapped to separate scale but same canvas
  ctx.strokeStyle = '#10b981';
  ctx.lineWidth = 2;
  ctx.beginPath();
  data.forEach((d, i) => {
    const y = phY(Math.min(phMax, Math.max(phMin, d.measured_ph)));
    i === 0 ? ctx.moveTo(xPos(i), y) : ctx.lineTo(xPos(i), y);
  });
  ctx.stroke();

  // Optimal reference lines (dashed)
  ctx.setLineDash([4, 3]);
  ctx.lineWidth = 1;
  // Cl optimal 1.0 – 3.0
  ctx.strokeStyle = 'rgba(59,130,246,0.4)';
  [1, 3].forEach(v => {
    ctx.beginPath(); ctx.moveTo(PAD.left, clY(v)); ctx.lineTo(PAD.left + cW, clY(v)); ctx.stroke();
  });
  // pH optimal 7.2 – 7.6
  ctx.strokeStyle = 'rgba(16,185,129,0.4)';
  [7.2, 7.6].forEach(v => {
    ctx.beginPath(); ctx.moveTo(PAD.left, phY(v)); ctx.lineTo(PAD.left + cW, phY(v)); ctx.stroke();
  });
  ctx.setLineDash([]);

  // Dots
  data.forEach((d, i) => {
    const cx  = xPos(i);
    const clv = Math.min(clMax, Math.max(clMin, d.measured_chlorine));
    const phv = Math.min(phMax, Math.max(phMin, d.measured_ph));
    ctx.fillStyle = '#3b82f6';
    ctx.beginPath(); ctx.arc(cx, clY(clv), 3, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = '#10b981';
    ctx.beginPath(); ctx.arc(cx, phY(phv), 3, 0, Math.PI * 2); ctx.fill();
  });
}

// ════════════════════════════════════════════════════════════════
// FEATURE 10 — Stock Management
// ════════════════════════════════════════════════════════════════
async function showStockModal() {
  const modal = $('modal-stock');
  const body  = $('stock-modal-body');
  if (!modal || !body) return;

  hideProductForm();
  const stock = await getAllStock();
  const isAdm = isAdmin();

  body.innerHTML = stock.map(p => {
    const low = (p.quantity || 0) <= (p.alert_threshold || 0);
    const visIcon = p.visible !== false ? '👁' : '👁‍🗨';
    return `
    <div class="stock-product-row" id="srow-${p.product_id}">
      <div style="flex:1">
        <div style="font-weight:600;font-size:.9rem;opacity:${p.visible !== false ? 1 : 0.5}">${escHtml(p.name)}</div>
        <div style="font-size:.75rem;color:var(--slate-500)">${p.unit} · pas: ${p.step || 1} · prag: ${p.alert_threshold || 0}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;justify-content:flex-end">
        <input type="number" class="stock-qty-input" id="stock-qty-${p.product_id}" value="${p.quantity || 0}" min="0" step="any" inputmode="decimal">
        <span style="font-size:.8rem;color:var(--slate-500)">${p.unit}</span>
        ${low ? `<span class="stock-low-badge">⚠</span>` : ''}
        ${isAdm ? `
        <button class="product-icon-btn" title="${p.visible !== false ? 'Ascunde din formular' : 'Afișează în formular'}" onclick="toggleProductVisible('${p.product_id}')">${visIcon}</button>
        <button class="product-icon-btn" title="Editează" onclick="showEditProductForm('${p.product_id}')">✏️</button>
        <button class="product-icon-btn product-icon-del" title="Șterge" onclick="deleteProduct('${p.product_id}')">🗑</button>
        ` : ''}
      </div>
    </div>`;
  }).join('');

  modal.classList.add('open');
}

async function saveStock() {
  const stock = await getAllStock();
  try {
    await Promise.all(stock.map(p => {
      const input = $(`stock-qty-${p.product_id}`);
      if (input) p.quantity = parseFloat(input.value) || 0;
      return updateStockProduct(p);
    }));
    showToast('Stoc actualizat.', 'success');
    closeStockModal();
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

async function deductStockForIntervention(intervention) {
  const stock = await getAllStock();
  for (const p of stock) {
    const used = intervention['treat_' + p.product_id] || 0;
    if (used > 0) {
      p.quantity = Math.max(0, (p.quantity || 0) - used);
      await updateStockProduct(p);
      if (p.quantity <= (p.alert_threshold || 0)) {
        showToast(`⚠ Stoc scăzut: ${p.name} (${p.quantity.toFixed(1)} ${p.unit})`, 'warning', 6000);
      }
    }
  }
}

function closeStockModal() {
  const modal = $('modal-stock');
  if (modal) modal.classList.remove('open');
}

// ════════════════════════════════════════════════════════════════
// FEATURE 11 — QR Code per client
// ════════════════════════════════════════════════════════════════
function showQRCode(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;

  const url    = location.origin + location.pathname + '?client=' + encodeURIComponent(clientId);
  const modal  = $('modal-qr');
  const canvas = $('qr-canvas');
  const nameEl = $('qr-client-name');
  const urlEl  = $('qr-url-text');
  const copyBtn = $('qr-copy-btn');

  if (!modal || !canvas) return;

  if (nameEl) {
    nameEl.textContent = client.name;
    // Add info icon button if not already present
    var infoBtn = $('client-info-btn');
    if (!infoBtn) {
      infoBtn = document.createElement('button');
      infoBtn.id = 'client-info-btn';
      infoBtn.className = 'client-info-btn';
      infoBtn.title = 'Info client';
      infoBtn.innerHTML = 'ℹ️';
      nameEl.parentNode.insertBefore(infoBtn, nameEl.nextSibling);
    }
    infoBtn.onclick = function() { showClientDetails(client.client_id); };
  }
  if (urlEl)  urlEl.textContent  = url;
  canvas.innerHTML = '';

  // Lazy-load QRCode.js from CDN if not already loaded
  if (typeof QRCode === 'undefined') {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js';
    script.onload = () => _renderQR(canvas, url);
    document.head.appendChild(script);
  } else {
    _renderQR(canvas, url);
  }

  if (copyBtn) {
    copyBtn.onclick = () => {
      navigator.clipboard.writeText(url).then(() => showToast('Link copiat!', 'success'))
        .catch(() => { prompt('Copiați URL-ul:', url); });
    };
  }

  modal.classList.add('open');
}

function _renderQR(container, text) {
  try {
    new QRCode(container, { text, width: 200, height: 200, correctLevel: QRCode.CorrectLevel.M });
  } catch (e) {
    container.textContent = 'Eroare QR: ' + e.message;
  }
}

function closeQRModal() {
  const modal = $('modal-qr');
  if (modal) modal.classList.remove('open');
}

// ════════════════════════════════════════════════════════════════
// FEATURE 12 — Share Raport Intervenție
// ════════════════════════════════════════════════════════════════

/**
 * Generates a formatted text report from the last saved intervention.
 * Uses WhatsApp-style bold (*text*) for good formatting.
 */
function generateInterventionReport(intervention, client) {
  if (!intervention) return '';

  const date = fmtDate(intervention.date);

  // Cl status
  const cl = intervention.measured_chlorine;
  const ph = intervention.measured_ph;
  const clOk = cl != null && cl >= 1 && cl <= 3;
  const phOk = ph != null && ph >= 7.2 && ph <= 7.6;

  // Build measured section
  const measured = [
    cl != null  ? `• Clor: ${cl} mg/L ${clOk ? '✅' : '⚠️'}` : null,
    ph != null  ? `• pH: ${ph} ${phOk ? '✅' : '⚠️'}` : null,
    intervention.measured_temp     != null ? `• Temperatură: ${intervention.measured_temp}°C` : null,
    intervention.measured_hardness != null ? `• Duritate: ${intervention.measured_hardness}` : null,
    intervention.measured_alkalinity != null ? `• Alcalinitate: ${intervention.measured_alkalinity}` : null,
    intervention.measured_salinity != null ? `• Salinitate: ${intervention.measured_salinity} g/L` : null,
  ].filter(Boolean).join('\n');

  // Build treatment section — only non-zero values
  const treatments = [
    (intervention.treat_cl_granule_gr || 0) > 0
      ? `• Cl Granule: ${intervention.treat_cl_granule_gr} gr` : null,
    (intervention.treat_cl_tablete || 0) > 0
      ? `• Cl Tablete: ${intervention.treat_cl_tablete} buc` : null,
    (intervention.treat_cl_lichid_bidoane || 0) > 0
      ? `• Cl Lichid: ${intervention.treat_cl_lichid_bidoane} bidoane` : null,
    (intervention.treat_ph_granule || 0) > 0
      ? `• pH Granule: ${intervention.treat_ph_granule} kg` : null,
    (intervention.treat_ph_lichid_bidoane || 0) > 0
      ? `• pH Lichid: ${intervention.treat_ph_lichid_bidoane} bidoane` : null,
    (intervention.treat_antialgic || 0) > 0
      ? `• Antialgic: ${intervention.treat_antialgic} L` : null,
    (intervention.treat_anticalcar || 0) > 0
      ? `• Anticalcar: ${intervention.treat_anticalcar} L` : null,
    (intervention.treat_floculant || 0) > 0
      ? `• Floculant: ${intervention.treat_floculant} L` : null,
    (intervention.treat_sare_saci || 0) > 0
      ? `• Sare: ${intervention.treat_sare_saci} saci` : null,
    (intervention.treat_bicarbonat || 0) > 0
      ? `• Bicarbonat: ${intervention.treat_bicarbonat} kg` : null,
  ].filter(Boolean);

  const treatmentBlock = treatments.length
    ? `\n🧪 *Tratament aplicat:*\n${treatments.join('\n')}`
    : '\n🧪 *Tratament:* fără produse adăugate';

  const durationBlock = intervention.duration_minutes != null
    ? `\n⏱ *Durată intervenție:* ${Math.round(intervention.duration_minutes)} min` : '';

  const obsBlock = intervention.observations
    ? `\n\n📝 *Observații:*\n${intervention.observations}` : '';

  return [
    `🏊 *Raport intervenție piscină*`,
    ``,
    `📅 *Data:* ${date}`,
    `👤 *Client:* ${client ? client.name : (intervention.client_name || '')}`,
    `👨‍🔧 *Tehnician:* ${intervention.technician_name || ''}`,
    ``,
    `📊 *Valori măsurate:*`,
    measured || '—',
    treatmentBlock,
    durationBlock,
    obsBlock,
    ``,
    `_Pool Manager App_`
  ].join('\n');
}

/**
 * Shares the intervention report via:
 * - 'whatsapp': Opens WhatsApp with client's phone pre-filled
 * - 'copy': Copies to clipboard, shows toast confirmation
 * - 'native': Uses Web Share API (Android Chrome)
 */
async function shareIntervention(method) {
  const intervention = APP.lastIntervention;
  const client = intervention ? APP.clients.find(c => c.client_id === intervention.client_id) : null;

  if (!intervention) {
    showToast('Nicio intervenție de partajat.', 'warning');
    return;
  }

  const text = generateInterventionReport(intervention, client);
  const hint = $('success-share-hint');

  if (method === 'whatsapp') {
    // Try Web Share API first (Android Chrome native share sheet)
    if (navigator.share) {
      try {
        await navigator.share({ text });
        return;
      } catch (e) {
        if (e.name === 'AbortError') return; // user cancelled
        // fall through to WhatsApp link
      }
    }
    // WhatsApp deeplink with client phone or generic
    const phone = client && client.phone
      ? '4' + client.phone.replace(/\D/g, '').slice(-9)
      : '';
    const waUrl = phone
      ? `https://wa.me/${phone}?text=${encodeURIComponent(text)}`
      : `https://wa.me/?text=${encodeURIComponent(text)}`;
    window.open(waUrl, '_blank', 'noopener');

  } else if (method === 'copy') {
    try {
      await navigator.clipboard.writeText(text);
      showToast('Raport copiat în clipboard!', 'success');
      if (hint) {
        hint.textContent = '✓ Copiat! Poți lipi în orice aplicație (WhatsApp, SMS, email...).';
        hint.style.display = '';
        setTimeout(() => { if (hint) hint.style.display = 'none'; }, 4000);
      }
    } catch {
      // Fallback: prompt with text selected
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'position:fixed;opacity:0;top:0;left:0';
      document.body.appendChild(ta);
      ta.focus(); ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      showToast('Raport copiat!', 'success');
    }
  }
}

// Check URL ?client=ID deeplink after login
function checkClientDeeplink() {
  const params = new URLSearchParams(location.search);
  const clientId = params.get('client');
  if (!clientId) return;
  const client = APP.clients.find(c => c.client_id === clientId);
  if (client) {
    // Clean URL without reload
    history.replaceState(null, '', location.pathname);
    openClientIntervention(client.client_id);
  }
}

// ── Info page search ─────────────────────────────────────────
function filterInfoSections(term) {
  const sections = $$('#screen-info .info-section');
  const noResults = $('info-no-results');
  const q = (term || '').trim().toLowerCase();
  let anyVisible = false;

  sections.forEach(sec => {
    // Remove previous highlights
    sec.querySelectorAll('mark').forEach(m => {
      m.replaceWith(document.createTextNode(m.textContent));
    });
    sec.normalize();

    if (!q) { sec.style.display = ''; return; }

    const text = sec.textContent.toLowerCase();
    // Also check data-title for keyword matching
    const title = (sec.dataset.title || '').toLowerCase();
    if (text.includes(q) || title.includes(q)) {
      sec.style.display = '';
      anyVisible = true;
      // Highlight matches in text nodes
      highlightInElement(sec, q);
    } else {
      sec.style.display = 'none';
    }
  });

  if (!q) { if (noResults) noResults.style.display = 'none'; return; }
  if (noResults) noResults.style.display = anyVisible ? 'none' : '';
}

function highlightInElement(el, q) {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
  const nodes = [];
  let node;
  while ((node = walker.nextNode())) nodes.push(node);
  nodes.forEach(n => {
    const idx = n.textContent.toLowerCase().indexOf(q);
    if (idx < 0 || n.parentNode.tagName === 'MARK') return;
    const before = n.textContent.slice(0, idx);
    const match  = n.textContent.slice(idx, idx + q.length);
    const after  = n.textContent.slice(idx + q.length);
    const frag   = document.createDocumentFragment();
    if (before) frag.appendChild(document.createTextNode(before));
    const mark = document.createElement('mark');
    mark.textContent = match;
    frag.appendChild(mark);
    if (after) frag.appendChild(document.createTextNode(after));
    n.parentNode.replaceChild(frag, n);
  });
}

// ── Info Screen — Edit Mode ──────────────────────────────────────

/** Load stored guide content from IndexedDB and inject into DOM sections.
 *  Also captures defaults the first time (before any injection). */
async function loadInfoContent() {
  if (!_infoDefaultHTML) {
    _infoDefaultHTML = {};
    $$('#screen-info .info-section').forEach((sec, i) => {
      const c = sec.querySelector('.form-section');
      if (c) _infoDefaultHTML[i] = c.innerHTML;
    });
  }
  const stored = await getSetting('info_sections');
  if (!stored) return;
  $$('#screen-info .info-section').forEach((sec, i) => {
    if (stored[i]) {
      const c = sec.querySelector('.form-section');
      if (c) c.innerHTML = stored[i];
    }
  });
}

/** Enter edit mode: make all .form-section divs contenteditable. */
function enterInfoEditMode() {
  _infoEditMode = true;
  _infoPreEditHTML = {};
  $$('#screen-info .info-section').forEach((sec, i) => {
    const c = sec.querySelector('.form-section');
    if (c) { _infoPreEditHTML[i] = c.innerHTML; c.contentEditable = 'true'; }
  });
  $('screen-info').classList.add('info-edit-mode');
  $('btn-info-edit').style.display = 'none';
  $('info-edit-actions').style.display = 'flex';
  // Disable search during editing to avoid mark-element conflicts
  const s = $('info-search');
  if (s) { s.value = ''; filterInfoSections(''); s.disabled = true; }
}

/** Save all section HTML to IndexedDB, exit edit mode. */
async function saveInfoContent() {
  const data = {};
  $$('#screen-info .info-section').forEach((sec, i) => {
    const c = sec.querySelector('.form-section');
    if (c) {
      // Strip <mark> highlights before saving
      const clone = c.cloneNode(true);
      clone.querySelectorAll('mark').forEach(m =>
        m.replaceWith(document.createTextNode(m.textContent)));
      data[i] = clone.innerHTML;
    }
  });
  await setSetting('info_sections', data);
  _exitInfoEditMode();
  showToast('Ghid salvat cu succes.', 'success');
}

/** Cancel edit: restore pre-edit snapshot, exit edit mode. */
function cancelInfoEditMode() {
  $$('#screen-info .info-section').forEach((sec, i) => {
    const c = sec.querySelector('.form-section');
    if (c && _infoPreEditHTML[i] !== undefined) c.innerHTML = _infoPreEditHTML[i];
  });
  _exitInfoEditMode();
}

/** Reset guide to original HTML defaults, clear stored overrides. */
async function resetInfoContent() {
  if (!confirm('Resetezi ghidul la conținutul implicit?\nModificările salvate se pierd definitiv.')) return;
  await setSetting('info_sections', null);
  if (_infoDefaultHTML) {
    $$('#screen-info .info-section').forEach((sec, i) => {
      const c = sec.querySelector('.form-section');
      if (c && _infoDefaultHTML[i] !== undefined) c.innerHTML = _infoDefaultHTML[i];
    });
  }
  _exitInfoEditMode();
  showToast('Ghid resetat la conținutul implicit.', 'success');
}

/** Internal: exit edit mode UI — remove contenteditable, restore buttons, re-enable search. */
function _exitInfoEditMode() {
  _infoEditMode = false;
  $$('#screen-info .info-section .form-section').forEach(c => c.removeAttribute('contenteditable'));
  $('screen-info').classList.remove('info-edit-mode');
  $('info-edit-actions').style.display = 'none';
  const editBtn = $('btn-info-edit');
  if (editBtn) editBtn.style.display = (APP.user && APP.user.role === 'admin') ? '' : 'none';
  const s = $('info-search'); if (s) s.disabled = false;
}

// ════════════════════════════════════════════════════════════════
// FEATURE A — Manager Produse Dinamic
// ════════════════════════════════════════════════════════════════

/**
 * Adds step + visible fields to existing stock products that don't have them.
 * Called at initApp(). Does NOT reset quantities.
 */
async function seedMissingStockProducts() {
  const stock = await getAllStock();
  const defaults = {
    cl_granule:  { step: 50,   visible: true },
    cl_tablete:  { step: 1,    visible: true },
    cl_lichid:   { step: 1,    visible: true },
    ph_minus_gr: { step: 0.1,  visible: true },
    ph_minus_l:  { step: 1,    visible: true },
    antialgic:   { step: 0.25, visible: true },
    anticalcar:  { step: 0.25, visible: true },
    floculant:   { step: 0.25, visible: true },
    sare:        { step: 1,    visible: true },
    bicarbonat:  { step: 0.5,  visible: true }
  };
  for (const p of stock) {
    let changed = false;
    if (p.step === undefined) {
      p.step = (defaults[p.product_id] || {}).step || 1;
      changed = true;
    }
    if (p.visible === undefined) {
      p.visible = true;
      changed = true;
    }
    if (changed) await updateStockProduct(p);
  }
}

/** Renders dynamic stepper rows for all visible stock products */
async function renderTreatmentSteppers() {
  const container = $('treatment-steppers-container');
  if (!container) return;

  const products = await getAllStock();
  APP._stockProducts = products;
  const visible = products.filter(p => p.visible !== false);

  if (!visible.length) {
    container.innerHTML = '<p style="padding:12px;color:var(--slate-400);font-size:.85rem">Niciun produs activ. Adaugă produse din Setări → Stoc.</p>';
    return;
  }

  container.innerHTML =
    `<div class="stepper-table-header">
      <span>Produs</span><span>Cantitate</span>
    </div>` +
    visible.map(p => {
      const isDecimal = (p.step || 1) < 1;
      return `
      <div class="stepper-row">
        <div class="stepper-label">${escHtml(p.name)} <small>${escHtml(p.unit)}</small></div>
        <div class="stepper-controls">
          <button class="stepper-btn" onclick="stepperChange('t-${p.product_id}',${-(p.step || 1)})">−</button>
          <input type="number" id="t-${p.product_id}" class="stepper-input"
                 value="0" min="0" step="${p.step || 1}"
                 inputmode="${isDecimal ? 'decimal' : 'numeric'}"
                 data-step="${p.step || 1}"
                 data-unit="${escHtml(p.unit)}"
                 data-label="${escHtml(p.name)}"
                 onclick="openDrumPicker(this)">
          <button class="stepper-btn" onclick="stepperChange('t-${p.product_id}',${p.step || 1})">+</button>
        </div>
      </div>`;
    }).join('');

  // Admin link to manage products
  if (isAdmin()) {
    container.insertAdjacentHTML('beforeend',
      `<div class="treat-manage-link admin-only">
        <button type="button" onclick="showStockModal()" class="btn-treat-manage">⚙ Gestionează produse</button>
      </div>`
    );
  }
}

/** Show add product form (blank) */
function showAddProductForm() {
  const form = $('product-form');
  if (!form) return;
  $('pf-id').value        = '';
  $('pf-name').value      = '';
  $('pf-unit').value      = 'kg';
  $('pf-step').value      = '1';
  $('pf-threshold').value = '0';
  $('pf-visible').checked = true;
  form.style.display = '';
  $('pf-name').focus();
}

/** Populate and show form for editing an existing product */
async function showEditProductForm(productId) {
  const form = $('product-form');
  if (!form) return;
  const p = await getByKey('stock', productId);
  if (!p) return;
  $('pf-id').value        = p.product_id;
  $('pf-name').value      = p.name;
  $('pf-unit').value      = p.unit;
  $('pf-step').value      = p.step || 1;
  $('pf-threshold').value = p.alert_threshold || 0;
  $('pf-visible').checked = p.visible !== false;
  form.style.display = '';
  $('pf-name').focus();
}

/** Hide product form */
function hideProductForm() {
  const form = $('product-form');
  if (form) form.style.display = 'none';
}

/** Save (add or edit) a product */
async function doSaveProduct() {
  const name = $('pf-name') ? $('pf-name').value.trim() : '';
  if (!name) { showToast('Denumirea produsului este obligatorie.', 'error'); return; }

  const existingId = $('pf-id') ? $('pf-id').value : '';
  const unit       = $('pf-unit')      ? $('pf-unit').value      : 'kg';
  const step       = parseFloat($('pf-step')?.value)      || 1;
  const threshold  = parseFloat($('pf-threshold')?.value) || 0;
  const visible    = $('pf-visible')   ? $('pf-visible').checked  : true;

  // Preserve existing quantity if editing
  let quantity = 0;
  if (existingId) {
    const existing = await getByKey('stock', existingId);
    if (existing) quantity = existing.quantity || 0;
  }

  const productId = existingId || ('prod_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));

  await updateStockProduct({ product_id: productId, name, unit, step, alert_threshold: threshold, visible, quantity });
  showToast(existingId ? 'Produs actualizat.' : 'Produs adăugat.', 'success');
  hideProductForm();
  showStockModal(); // re-render stock list
  renderTreatmentSteppers().catch(() => {}); // refresh treatment form if open
}

/** Delete a product (with confirm) */
async function deleteProduct(productId) {
  if (!confirm('Ștergi produsul? Această acțiune nu poate fi anulată.')) return;
  await deleteRecord('stock', productId);
  showToast('Produs șters.', 'success');
  showStockModal();
  renderTreatmentSteppers().catch(() => {}); // refresh treatment form if open
}

/** Toggle visibility of a product in the intervention form */
async function toggleProductVisible(productId) {
  const p = await getByKey('stock', productId);
  if (!p) return;
  p.visible = p.visible === false ? true : false;
  await updateStockProduct(p);
  showStockModal();
  renderTreatmentSteppers().catch(() => {}); // refresh treatment form if open
}

// ════════════════════════════════════════════════════════════════
// FEATURE B — Wizard pe 3 Pași
// ════════════════════════════════════════════════════════════════

/** Navigate to a wizard step (1, 2, or 3) */
function goWizardStep(step) {
  APP.wizardStep = step;

  // Show/hide step panels
  [1, 2].forEach(s => {
    const el = $('wiz-step-' + s);
    if (el) el.classList.toggle('active', s === step);
  });

  // Update progress dots
  $$('#wizard-progress .wiz-dot').forEach(dot => {
    const dotStep = parseInt(dot.dataset.step);
    dot.classList.toggle('active', dotStep <= step);
  });

  // Save bar: visible only on step 2
  const saveBar = $('save-bar');
  if (saveBar) saveBar.style.display = step === 2 ? '' : 'none';
  // When entering step 2, default to treatment tab button state
  if (step === 2) switchP2Tab('treatment');

  // Scroll to top of intervention screen
  const screen = $('screen-intervention');
  if (screen) screen.scrollTop = 0;
}

/** Go to next step with validation on step 1 */
function nextWizardStep() {
  if (APP.wizardStep === 1) {
    // Validate chlorine + pH before proceeding
    const cl = $('m-chlorine');
    const ph = $('m-ph');
    let valid = true;
    [cl, ph].forEach(el => {
      if (!el) return;
      const val = el.value.trim();
      if (!val || isNaN(parseFloat(val))) {
        el.classList.add('error');
        valid = false;
      } else {
        el.classList.remove('error');
      }
    });
    if (!valid) {
      showToast('Completați clorul și pH-ul măsurate.', 'error');
      return;
    }
    goWizardStep(2);
  }
}

/** Go to previous step */
function prevWizardStep() {
  if (APP.wizardStep > 1) goWizardStep(APP.wizardStep - 1);
}

// ── Swipe on step 2 tabs (Tratament <-> Note & Foto) ─────
(function setupP2Swipe() {
  let _p2TouchX = 0;
  let _p2TouchY = 0;
  document.addEventListener('touchstart', e => {
    const screen = document.getElementById('screen-intervention');
    if (!screen || !screen.classList.contains('active')) return;
    if (typeof APP === 'undefined' || APP.wizardStep !== 2) return;
    _p2TouchX = e.touches[0].clientX;
    _p2TouchY = e.touches[0].clientY;
  }, { passive: true });
  document.addEventListener('touchend', e => {
    const screen = document.getElementById('screen-intervention');
    if (!screen || !screen.classList.contains('active')) return;
    if (typeof APP === 'undefined' || APP.wizardStep !== 2) return;
    const dx = e.changedTouches[0].clientX - _p2TouchX;
    const dy = e.changedTouches[0].clientY - _p2TouchY;
    if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.5) {
      const treatActive = document.getElementById('tab-treatment');
      if (dx < 0 && treatActive && treatActive.classList.contains('active')) {
        switchP2Tab('notes');   // swipe left = go to Notes & Foto
      } else if (dx > 0 && treatActive && !treatActive.classList.contains('active')) {
        switchP2Tab('treatment'); // swipe right = go to Treatment
      }
    }
  }, { passive: true });
})();

/** Switch tab on page 2 (Tratament / Note & Foto) */
function switchP2Tab(tab) {
  ['treatment', 'notes'].forEach(t => {
    const btn   = $('tab-' + t);
    const panel = $('panel-' + t);
    if (btn)   btn.classList.toggle('active',   t === tab);
    if (panel) panel.classList.toggle('active', t === tab);
  });
  // Update save button based on active tab
  var saveBtn = $('btn-save');
  if (saveBtn) {
    if (tab === 'treatment') {
      saveBtn.textContent = '➡ Spre Finalizare';
      saveBtn.disabled = false;
      saveBtn.onclick = function() { switchP2Tab('notes'); };
    } else {
      saveBtn.textContent = '💾 Salvează Intervenția';
      saveBtn.onclick = showConfirmModal;
    }
  }
}

/** Toggle collapsible section (used for "Ultimele intervenții") */
function toggleSection(titleEl) {
  const body = titleEl.nextElementSibling;
  if (!body) return;
  const isHidden = body.style.display === 'none';
  body.style.display = isHidden ? '' : 'none';
  const span = titleEl.querySelector('span') || titleEl;
  span.textContent = span.textContent.replace(/^[▶▼]\s*/, (isHidden ? '▼ ' : '▶ '));
}

// ── History List Rendering ────────────────────────────────────
function _renderHistoryList(clientId, allInterventions) {
  var container = $('history-list');
  if (!container) return;

  var dateFilter = $('history-date-filter');
  var fromDate = dateFilter ? dateFilter.value : '';
  var filtered = allInterventions;
  if (fromDate) {
    filtered = allInterventions.filter(function(i) { return i.date >= fromDate; });
  }

  if (filtered.length === 0) {
    container.innerHTML = '<p style="color:var(--slate-400);font-size:.85rem">Nicio intervenție' + (fromDate ? ' din această perioadă' : '') + '.</p>';
    return;
  }

  var html = '';
  filtered.forEach(function(i) {
    var chems = [];
    if (i.treat_cl_granule_gr > 0) chems.push('Cl.gr: ' + i.treat_cl_granule_gr + 'g');
    if (i.treat_cl_tablete > 0) chems.push('Cl.tab: ' + i.treat_cl_tablete);
    if (i.treat_ph_granule > 0) chems.push('pH: ' + i.treat_ph_granule + 'kg');
    if (i.treat_antialgic > 0) chems.push('Anti: ' + i.treat_antialgic + 'L');
    if (i.treat_floculant > 0) chems.push('Floc: ' + i.treat_floculant + 'L');
    if (i.treat_bicarbonat > 0) chems.push('Dedur: ' + i.treat_bicarbonat + 'kg');
    if (i.treat_ph_lichid_bidoane > 0) chems.push('pH.L: ' + i.treat_ph_lichid_bidoane);
    if (i.treat_cl_lichid_bidoane > 0) chems.push('Cl.L: ' + i.treat_cl_lichid_bidoane);
    if (i.treat_sare_saci > 0) chems.push('Sare: ' + i.treat_sare_saci);

    var opsArr = Array.isArray(i.operations) ? i.operations : (typeof i.operations === 'string' && i.operations.length > 0 ? (function() { try { return JSON.parse(i.operations); } catch(e) { return []; } })() : []);
    var ops = opsArr.join(', ');

    html += '<div class="prev-intervention" style="position:relative;cursor:pointer" onclick="showInterventionDetails(\'' + i.intervention_id + '\')">';
    html += '<div class="prev-int-header">';
    html += '<span class="prev-int-date">' + fmtDate(i.date) + '</span>';
    if (i.duration_minutes != null) {
      html += '<span class="prev-int-duration">⏱ ' + Math.round(i.duration_minutes) + ' min</span>';
    }
    // Edit: all users; Delete: admin only (stopPropagation to avoid opening details)
    html += '<span style="display:flex;gap:4px;margin-left:auto" onclick="event.stopPropagation()">';
    html += '<button onclick="editIntervention(\'' + i.intervention_id + '\',\'' + clientId + '\')" style="background:var(--blue-100);border:none;border-radius:6px;padding:3px 8px;font-size:.75rem;color:var(--blue-700);cursor:pointer">✏️</button>';
    if (isAdmin()) {
      html += '<button onclick="deleteIntervention(\'' + i.intervention_id + '\',\'' + clientId + '\')" style="background:var(--red-100,#fee2e2);border:none;border-radius:6px;padding:3px 8px;font-size:.75rem;color:var(--danger);cursor:pointer">🗑️</button>';
    }
    html += '</span>';
    html += '</div>';
    html += '<div class="prev-int-tech">👤 ' + escHtml(i.technician_name || '') + '</div>';
    html += '<div class="prev-int-measures">';
    html += '<span class="prev-measure">Cl: <strong>' + (i.measured_chlorine != null ? i.measured_chlorine : '—') + '</strong></span>';
    html += '<span class="prev-measure">pH: <strong>' + (i.measured_ph != null ? i.measured_ph : '—') + '</strong></span>';
    html += '</div>';
    if (chems.length) {
      html += '<div class="prev-int-measures" style="margin-top:2px"><span class="prev-measure" style="font-size:.75rem;color:var(--text-secondary)">' + chems.join(' · ') + '</span></div>';
    }
    if (ops) {
      html += '<div style="font-size:.72rem;color:var(--emerald-600);margin-top:2px">✓ ' + escHtml(ops) + '</div>';
    }
    if (i.observations) {
      html += '<div style="font-size:.75rem;color:var(--text-secondary);margin-top:2px;font-style:italic">"' + escHtml(i.observations) + '"</div>';
    }
    if (i.audio_file_url) {
      html += '<div style="margin-top:4px" onclick="event.stopPropagation()">';
      html += '<a href="' + escHtml(i.audio_file_url) + '" target="_blank" rel="noopener" style="font-size:.75rem;color:var(--blue-600);text-decoration:none;background:var(--blue-50,#eff6ff);padding:3px 8px;border-radius:6px;display:inline-block">🎙️ Ascultă înregistrarea</a>';
      if (i.measured_chlorine == null && i.measured_ph == null) {
        html += ' <span style="font-size:.7rem;color:var(--amber-600)">⚠ de completat</span>';
      }
      html += '</div>';
    }
    // Show photos if available
    if (i.photos && i.photos.length > 0) {
      html += '<div style="display:flex;gap:4px;margin-top:4px;flex-wrap:wrap" onclick="event.stopPropagation()">';
      i.photos.forEach(function(photoUrl, pi) {
        html += '<img src="' + photoUrl + '" alt="Foto ' + (pi+1) + '" style="width:48px;height:48px;object-fit:cover;border-radius:6px;cursor:pointer;border:1px solid var(--slate-300)" onclick="window.open(this.src)">';
      });
      html += '</div>';
    }
    if (!i.synced) {
      html += '<span style="font-size:.65rem;color:var(--amber-600);display:block;margin-top:2px">⚠ Nesincronizat</span>';
    }
    html += '</div>';
  });

  container.innerHTML = html;
}

function filterHistoryByDate(clientId) {
  var ci = APP.interventions.filter(function(i) { return i.client_id === clientId && i.date; })
    .map(function(i) {
      var raw = String(i.date || '');
      if (raw && !/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        var dp = new Date(raw);
        if (!isNaN(dp.getTime())) i.date = dp.getFullYear() + '-' + ('0'+(dp.getMonth()+1)).slice(-2) + '-' + ('0'+dp.getDate()).slice(-2);
      }
      return i;
    })
    .sort(function(a, b) {
      var cmp = String(b.date || '').localeCompare(String(a.date || ''));
      if (cmp !== 0) return cmp;
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    });
  _renderHistoryList(clientId, ci);
}

// ── Delete / Edit Intervention ────────────────────────────────
async function deleteIntervention(interventionId, clientId) {
  if (!isAdmin()) return;
  if (!confirm('Sigur vrei sa stergi aceasta interventie?')) return;

  try {
    await deleteRecord('interventions', interventionId);
    APP.interventions = APP.interventions.filter(function(i) { return i.intervention_id !== interventionId; });

    // Track deleted ID so pull won't re-add it
    await _trackDeletedIntervention(interventionId);

    // If synced, notify GAS
    if (isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'delete_intervention', data: { intervention_id: interventionId } })
      }).catch(function(e) { console.warn('[SYNC] Delete push failed:', e.message); });
    }

    showToast('Interventie stearsa.', 'success');
    // Refresh the details modal
    showClientDetails(clientId);
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

/** Track deleted intervention IDs so sync pull ignores them */
async function _trackDeletedIntervention(interventionId) {
  var deleted = await getSetting('deleted_intervention_ids').catch(function() { return null; }) || [];
  if (!Array.isArray(deleted)) deleted = [];
  if (deleted.indexOf(interventionId) < 0) deleted.push(interventionId);
  // Keep max 500 entries to avoid bloat
  if (deleted.length > 500) deleted = deleted.slice(-500);
  await setSetting('deleted_intervention_ids', deleted);
}

/** Check if an intervention was locally deleted */
async function _isDeletedIntervention(interventionId) {
  var deleted = await getSetting('deleted_intervention_ids').catch(function() { return null; }) || [];
  return Array.isArray(deleted) && deleted.indexOf(interventionId) >= 0;
}

/** Delete ALL interventions (local + GAS). Call from console: deleteAllInterventions() */
async function deleteAllInterventions() {
  if (!confirm('ATENȚIE: Sigur vrei să ștergi TOATE intervențiile? Acțiunea este ireversibilă!')) return;
  try {
    // Track all existing IDs as deleted so pull won't re-add them
    var allIds = APP.interventions.map(function(i) { return i.intervention_id; });
    var deleted = await getSetting('deleted_intervention_ids').catch(function() { return null; }) || [];
    if (!Array.isArray(deleted)) deleted = [];
    allIds.forEach(function(id) { if (deleted.indexOf(id) < 0) deleted.push(id); });
    await setSetting('deleted_intervention_ids', deleted);

    await clearStore('interventions');
    APP.interventions = [];
    showToast('Toate intervențiile au fost șterse local.', 'success');

    // Also clear on GAS
    if (isSyncConfigured()) {
      apiFetch(SYNC_CONFIG.API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'push', type: 'clear_interventions', data: {} })
      }).then(function() {
        showToast('Intervențiile au fost șterse și pe server.', 'success');
      }).catch(function(e) {
        showToast('Șterse local, eroare server: ' + e.message, 'warning');
      });
    }

    renderDashboard();
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

function editIntervention(interventionId, clientId) {
  var intervention = APP.interventions.find(function(i) { return i.intervention_id === interventionId; });
  if (!intervention) { showToast('Interventie negasita.', 'error'); return; }

  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;

  // Close history modal
  var modal = $('modal-client');
  if (modal) modal.classList.remove('open');

  // Open intervention screen with pre-filled data
  APP._editingIntervention = intervention;
  openClientIntervention(clientId);
}

// ════════════════════════════════════════════════════════════════
// BILLING LIST SCREEN — De Facturat
// ════════════════════════════════════════════════════════════════

/** Get all clients that reached billing threshold */
function _getBillableClients() {
  if (!isAdmin()) return [];
  return APP.clients.filter(function(client) {
    var interval = client.billing_interval_interventions;
    if (!interval || interval <= 0) return false;
    var since = client.last_billing_date || '1970-01-01';
    var count = APP.interventions.filter(function(i) {
      return i.client_id === client.client_id && String(i.date || '') > since;
    }).length;
    return count >= interval;
  }).map(function(client) {
    var since = client.last_billing_date || '1970-01-01';
    var billable = APP.interventions.filter(function(i) {
      return i.client_id === client.client_id && String(i.date || '') > since;
    }).sort(function(a, b) { return String(a.date || '').localeCompare(String(b.date || '')); });
    return { client: client, interventions: billable, count: billable.length };
  });
}

/** Show billing list screen */
function showBillingListScreen() {
  if (!isAdmin()) return;
  showScreen('billing-list');
  renderBillingList();
}

// ── Azi card: arata interventiile din ziua curenta ─────────────
function showTodayInterventions() {
  const today = new Date().toISOString().split('T')[0];
  const items = APP.interventions
    .filter(i => i.date === today)
    .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));

  const title = $('modal-today-title');
  const body  = $('modal-today-body');
  if (!body) return;

  if (title) title.textContent = `Intervenții astăzi (${items.length})`;

  if (!items.length) {
    body.innerHTML = '<div style="text-align:center;padding:30px 16px;color:var(--text-secondary)">' +
      '<div style="font-size:2.5rem;margin-bottom:10px">📋</div>' +
      '<p style="font-size:.95rem;font-weight:600">Nicio intervenție astăzi</p>' +
      '<p style="font-size:.8rem">Intervențiile salvate azi vor apărea aici.</p></div>';
  } else {
    let html = '';
    items.forEach(i => {
      const client = APP.clients.find(c => c.client_id === i.client_id);
      const cname  = client ? client.name : 'Client șters';
      const tname  = i.technician_name || '';
      const time   = i.created_at ? new Date(i.created_at).toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' }) : '';
      const ops    = Array.isArray(i.operations) ? i.operations.filter(Boolean).join(', ') : '';
      const obs    = i.observations || '';
      const fac    = (i.chlorine != null ? 'Cl:' + i.chlorine : '');
      const ph     = (i.ph != null ? ' pH:' + i.ph : '');

      html += '<div class="billing-list-card" style="cursor:pointer" onclick="closeTodayModal();openClientModalById(\'' + i.client_id + '\')">' +
        '<div class="billing-list-info" style="flex:1">' +
          '<div style="font-weight:700;font-size:.95rem">' + escHtml(cname) + '</div>' +
          '<div style="font-size:.78rem;color:var(--text-secondary);margin-top:2px">' +
            (time ? '🕒 ' + time + ' · ' : '') + escHtml(tname) +
          '</div>' +
          (fac || ph ? '<div style="font-size:.78rem;color:var(--text-secondary);margin-top:2px">' + escHtml(fac + ph) + '</div>' : '') +
          (ops   ? '<div style="font-size:.78rem;color:var(--text-secondary);margin-top:2px">⚙ ' + escHtml(ops) + '</div>' : '') +
          (obs   ? '<div style="font-size:.78rem;color:var(--text-secondary);margin-top:2px;font-style:italic">💬 ' + escHtml(obs) + '</div>' : '') +
        '</div>' +
      '</div>';
    });
    body.innerHTML = html;
  }

  const modal = $('modal-today');
  if (modal) modal.classList.add('open');
}

function closeTodayModal() {
  const modal = $('modal-today');
  if (modal) modal.classList.remove('open');
}

function openClientModalById(clientId) {
  if (typeof showClientDetails === 'function') showClientDetails(clientId);
}

// ── Detalii intervenție (modal) ───────────────────────────────
function showInterventionDetails(interventionId) {
  const i = APP.interventions.find(x => x.intervention_id === interventionId);
  if (!i) { showToast('Intervenția nu a fost găsită.', 'error'); return; }

  const body  = $('modal-intv-details-body');
  const title = $('modal-intv-details-title');
  if (!body) return;

  const client = APP.clients.find(c => c.client_id === i.client_id);
  const cname  = client ? client.name : (i.client_name || 'Client șters');
  if (title) title.textContent = cname + ' — ' + fmtDate(i.date);

  const row = (lbl, val) =>
    '<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--slate-200);font-size:.88rem">' +
    '<span style="color:var(--text-secondary)">' + lbl + '</span>' +
    '<span style="font-weight:600;text-align:right">' + val + '</span>' +
    '</div>';

  const num = v => (v != null && v !== '' ? v : '—');
  const sect = (ttl, inner) =>
    '<div style="margin-bottom:14px">' +
    '<div style="font-size:.78rem;font-weight:700;color:var(--blue-700);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px;padding-bottom:4px;border-bottom:2px solid var(--blue-200)">' + ttl + '</div>' +
    inner + '</div>';

  let html = '';

  // Info general
  let general = '';
  general += row('👤 Tehnician', escHtml(i.technician_name || '—'));
  general += row('📅 Data', fmtDate(i.date));
  if (i.created_at) {
    const t = new Date(i.created_at);
    general += row('🕒 Ora', t.toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' }));
  }
  if (i.duration_minutes != null) general += row('⏱ Durata', Math.round(i.duration_minutes) + ' min');
  if (i.arrival_time)   general += row('➡ Sosire',  new Date(i.arrival_time).toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' }));
  if (i.departure_time) general += row('⬅ Plecare', new Date(i.departure_time).toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' }));
  if (i.geo_lat && i.geo_lng) {
    general += row('📍 GPS',
      '<a href="https://www.google.com/maps?q=' + i.geo_lat + ',' + i.geo_lng + '" target="_blank" style="color:var(--blue-600)">' +
      Number(i.geo_lat).toFixed(5) + ', ' + Number(i.geo_lng).toFixed(5) + '</a>');
  }
  if (i.audio_file_url) {
    general += row('🎙️ Notă vocală', '<a href="' + escHtml(i.audio_file_url) + '" target="_blank" rel="noopener" style="color:var(--blue-600)">Ascultă înregistrarea</a>');
  }
  html += sect('Informații generale', general);

  // Valori măsurate
  let meas = '';
  meas += row('Clor liber (Cl)',    num(i.measured_chlorine) + (i.measured_chlorine != null ? ' mg/L' : ''));
  meas += row('pH',                 num(i.measured_ph));
  if (i.measured_tc != null)         meas += row('Clor total (TC)', i.measured_tc + ' mg/L');
  if (i.measured_cya != null)        meas += row('CYA (acid cianuric)', i.measured_cya + ' mg/L');
  if (i.measured_temp != null)       meas += row('Temperatură', i.measured_temp + ' °C');
  if (i.measured_hardness != null)   meas += row('Duritate', i.measured_hardness);
  if (i.measured_alkalinity != null) meas += row('Alcalinitate', i.measured_alkalinity);
  if (i.measured_salinity != null)   meas += row('Salinitate', i.measured_salinity);
  html += sect('Valori măsurate', meas);

  // Operațiuni efectuate
  const opsArr = Array.isArray(i.operations) ? i.operations :
    (typeof i.operations === 'string' && i.operations ? (function(){ try { return JSON.parse(i.operations); } catch { return []; } })() : []);
  if (opsArr.length) {
    let opsHtml = '<div style="display:flex;flex-wrap:wrap;gap:6px;padding:4px 0">';
    opsArr.forEach(op => {
      opsHtml += '<span style="background:var(--emerald-100);color:var(--emerald-700);padding:4px 10px;border-radius:12px;font-size:.82rem;font-weight:600">✓ ' + escHtml(op) + '</span>';
    });
    opsHtml += '</div>';
    html += sect('Operațiuni efectuate', opsHtml);
  }

  // Tratament efectuat
  const treatments = [];
  if (i.treat_cl_granule_gr > 0)     treatments.push(['Clor granule',        i.treat_cl_granule_gr + ' g']);
  if (i.treat_cl_tablete > 0)        treatments.push(['Clor tablete',        i.treat_cl_tablete + ' buc']);
  if (i.treat_cl_lichid_bidoane > 0) treatments.push(['Clor lichid',         i.treat_cl_lichid_bidoane + ' bid']);
  if (i.treat_ph_granule > 0)        treatments.push(['pH minus granule',    i.treat_ph_granule + ' kg']);
  if (i.treat_ph_lichid_bidoane > 0) treatments.push(['pH minus lichid',     i.treat_ph_lichid_bidoane + ' bid']);
  if (i.treat_antialgic > 0)         treatments.push(['Antialgic',           i.treat_antialgic + ' L']);
  if (i.treat_anticalcar > 0)        treatments.push(['Anticalcar',          i.treat_anticalcar + ' L']);
  if (i.treat_floculant > 0)         treatments.push(['Floculant',           i.treat_floculant + ' L']);
  if (i.treat_bicarbonat > 0)        treatments.push(['Bicarbonat',          i.treat_bicarbonat + ' kg']);
  if (i.treat_sare_saci > 0)         treatments.push(['Sare',                i.treat_sare_saci + ' saci']);

  if (treatments.length) {
    let tHtml = '';
    treatments.forEach(t => { tHtml += row(t[0], t[1]); });
    html += sect('Tratament efectuat', tHtml);
  }

  // Observații
  if (i.observations && i.observations.trim()) {
    html += sect('Observații',
      '<div style="background:var(--slate-50);padding:10px 12px;border-radius:8px;font-size:.88rem;font-style:italic;color:var(--slate-700);white-space:pre-wrap">' +
      escHtml(i.observations) + '</div>');
  }

  // Poze
  if (Array.isArray(i.photos) && i.photos.length) {
    let pHtml = '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px;padding:4px 0">';
    i.photos.forEach((p, idx) => {
      pHtml += '<img src="' + p + '" alt="Foto ' + (idx+1) + '" style="width:100%;aspect-ratio:1;object-fit:cover;border-radius:8px;cursor:pointer;border:1px solid var(--slate-300)" onclick="window.open(this.src)">';
    });
    pHtml += '</div>';
    html += sect('Fotografii (' + i.photos.length + ')', pHtml);
  }

  // Status sync
  if (i.synced === false) {
    html += '<div style="margin-top:8px;padding:8px 12px;background:var(--amber-100);color:var(--amber-700);border-radius:8px;font-size:.82rem">⚠ Nesincronizat — va fi trimis la următoarea sincronizare.</div>';
  }

  body.innerHTML = html;

  const modal = $('modal-intv-details');
  if (modal) modal.classList.add('open');
}

function closeInterventionDetails() {
  const modal = $('modal-intv-details');
  if (modal) modal.classList.remove('open');
}

/** Render the billing list */
function renderBillingList() {
  var container = $('billing-list-content');
  if (!container) return;

  var items = _getBillableClients();
  if (!items.length) {
    container.innerHTML = '<div style="text-align:center;padding:40px 20px;color:var(--text-secondary)">' +
      '<div style="font-size:2.5rem;margin-bottom:12px">&#9989;</div>' +
      '<p style="font-size:1rem;font-weight:600">Niciun client de facturat</p>' +
      '<p style="font-size:.85rem">Toti clientii sunt la zi cu facturarea.</p></div>';
    return;
  }

  var html = '<div style="margin-bottom:12px;font-size:.85rem;color:var(--text-secondary)">' +
    items.length + ' client(i) de facturat</div>';

  items.forEach(function(item) {
    var c = item.client;
    var since = c.last_billing_date ? fmtDate(c.last_billing_date) : 'prima interventie';
    var interval = c.billing_interval_interventions || 4;

    html += '<div class="billing-list-card">' +
      '<div class="billing-list-info">' +
        '<div style="font-weight:700;font-size:.95rem">' + escHtml(c.name) + '</div>' +
        '<div style="font-size:.8rem;color:var(--text-secondary)">' +
          item.count + ' interventii (prag: ' + interval + ') &middot; din ' + since +
        '</div>' +
      '</div>' +
      '<div class="billing-list-actions" style="display:flex;gap:6px;flex-wrap:wrap">' +
        '<button class="billing-list-btn export" onclick="exportBillingClient(\'' + c.client_id + '\')" title="Deviz Excel" style="font-size:1.1rem">&#128230;</button>' +
        '<button class="billing-list-btn" onclick="exportBillingPdf(\'' + c.client_id + '\')" title="Deviz PDF" style="font-size:1.1rem">&#128196;</button>' +
        '<button class="billing-list-btn" onclick="sendBillingWhatsApp(\'' + c.client_id + '\')" title="Trimite WhatsApp" style="font-size:1.1rem;color:#25D366">&#128172;</button>' +
        '<button class="billing-list-btn reset" onclick="resetBillingClient(\'' + c.client_id + '\')" title="Marcheaza facturat" style="font-size:1.1rem">&#8634;</button>' +
      '</div>' +
    '</div>';
  });

  container.innerHTML = html;
}

/** Export one client's billing deviz */
async function exportBillingClient(clientId) {
  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;
  var since = client.last_billing_date || '1970-01-01';
  var billable = APP.interventions.filter(function(i) {
    return i.client_id === clientId && i.date > since;
  }).sort(function(a, b) { return a.date.localeCompare(b.date); });

  if (!billable.length) { showToast('Nicio interventie de exportat.', 'warning'); return; }

  showToast('Generare deviz ' + client.name + '...', 'info');
  try {
    var devizType = parseInt(client.deviz_type) || 2;
    if (devizType === 2) {
      await exportDevizComplet(client, billable);
    } else {
      await exportDevizChimicale(client, billable);
    }
    showToast('Export complet: ' + client.name, 'success');
  } catch (e) {
    showToast('Eroare export: ' + e.message, 'error');
  }
}

/** Reset one client's billing (mark as billed) */
async function resetBillingClient(clientId) {
  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;

  client.last_billing_date = new Date().toISOString().split('T')[0];
  client.updated_at = new Date().toISOString();
  await put('clients', client);
  APP.clients = APP.clients.map(function(c) { return c.client_id === clientId ? client : c; });

  if (isSyncConfigured()) {
    apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'clients', data: [client] })
    }).catch(function(e) { console.warn('[SYNC] Billing push failed:', e.message); });
  }

  showToast(client.name + ' marcat ca facturat.', 'success');
  renderBillingList();
  var elBilling = $('stat-billing-count');
  if (elBilling) elBilling.textContent = _getBillableClients().length;
}

/** Mark client as billed from client details modal */
async function markClientBilled() {
  var clientId = APP._billingClientId;
  if (!clientId) return;
  await resetBillingClient(clientId);
  var billBtn = $('btn-mark-billed');
  if (billBtn) billBtn.style.display = 'none';
}

/** Export one client's billing as PDF (print) */
async function exportBillingPdf(clientId) {
  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;
  var since = client.last_billing_date || '1970-01-01';
  var billable = APP.interventions.filter(function(i) {
    return i.client_id === clientId && i.date > since;
  }).sort(function(a, b) { return a.date.localeCompare(b.date); });
  if (!billable.length) { showToast('Nicio interventie de exportat.', 'warning'); return; }

  showToast('Generare PDF ' + client.name + '...', 'info');
  try {
    var devizType = parseInt(client.deviz_type) || 2;
    if (devizType === 2) {
      await exportDevizComplet(client, billable);
    } else {
      await exportDevizChimicale(client, billable);
    }
    showToast('PDF generat: ' + client.name, 'success');
  } catch (e) {
    showToast('Eroare PDF: ' + e.message, 'error');
  }
}

/** Send WhatsApp notification for a specific billable client */
function sendBillingWhatsApp(clientId) {
  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;
  var since = client.last_billing_date || '1970-01-01';
  var count = APP.interventions.filter(function(i) {
    return i.client_id === clientId && String(i.date || '') > since;
  }).length;

  Promise.all([getSetting('wa_phone'), getSetting('wa_apikey')]).then(function(vals) {
    var phone = vals[0], apikey = vals[1];
    if (!phone || !apikey) {
      showToast('WhatsApp neconfigurat! Mergi la Settings.', 'warning');
      return;
    }
    var msg = '*Facturare: ' + client.name + '*\n'
      + count + ' interventii nefacturate\n'
      + 'Tel: ' + (client.phone || '-') + '\n'
      + 'Adresa: ' + (client.address || '-') + '\n'
      + 'Data: ' + new Date().toLocaleDateString('ro-RO') + '\n'
      + '_Generat de Pool Manager_';

    var url = 'https://api.callmebot.com/whatsapp.php'
      + '?phone=' + encodeURIComponent(phone)
      + '&text=' + encodeURIComponent(msg)
      + '&apikey=' + encodeURIComponent(apikey);

    fetch(url, { mode: 'no-cors' }).then(function() {
      showToast('WhatsApp trimis pentru ' + client.name + '!', 'success');
    }).catch(function(e) {
      showToast('WhatsApp nereusit: ' + e.message, 'warning');
    });
  });
}

/** Export all billing clients */
async function exportAllBilling() {
  var items = _getBillableClients();
  if (!items.length) { showToast('Niciun client de facturat.', 'warning'); return; }

  showToast('Export ' + items.length + ' clienti...', 'info');
  var errors = 0;
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    try {
      var devizType = parseInt(item.client.deviz_type) || 2;
      if (devizType === 2) {
        await exportDevizComplet(item.client, item.interventions);
      } else {
        await exportDevizChimicale(item.client, item.interventions);
      }
    } catch (e) {
      errors++;
      console.warn('[BILLING] Export error for', item.client.name, e.message);
    }
  }
  showToast('Export complet: ' + (items.length - errors) + ' clienti.' + (errors ? ' (' + errors + ' erori)' : ''), errors ? 'warning' : 'success');
}

/** Reset all billing clients */
async function resetAllBilling() {
  var items = _getBillableClients();
  if (!items.length) { showToast('Niciun client de facturat.', 'warning'); return; }

  var today = new Date().toISOString().split('T')[0];
  var now = new Date().toISOString();
  var updated = [];

  for (var idx = 0; idx < items.length; idx++) {
    var client = items[idx].client;
    client.last_billing_date = today;
    client.updated_at = now;
    await put('clients', client);
    updated.push(client);
  }

  APP.clients = APP.clients.map(function(c) {
    var u = updated.find(function(x) { return x.client_id === c.client_id; });
    return u || c;
  });

  if (isSyncConfigured()) {
    apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'clients', data: updated })
    }).catch(function(e) { console.warn('[SYNC] Billing reset push failed:', e.message); });
  }

  showToast(updated.length + ' clienti marcati ca facturati.', 'success');
  renderBillingList();
  var elBilling = $('stat-billing-count');
  if (elBilling) elBilling.textContent = 0;
}

// ════════════════════════════════════════════════════════════════
// FEATURE C — Notificare Facturare per Client
// ════════════════════════════════════════════════════════════════

/** Check if billing threshold is reached — admin-only modal */
function checkBillingAlert(client) {
  const interval = parseInt(client.billing_interval_interventions) || 0;
  if (!interval || interval <= 0) {
    return;
  }

  const since = client.last_billing_date || '1970-01-01';
  // Normalize dates before comparing to handle GAS Date objects
  const billable = APP.interventions.filter(function(i) {
    if (i.client_id !== client.client_id) return false;
    var raw = String(i.date || '');
    // Normalize to YYYY-MM-DD
    if (raw && !/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
      var dp = new Date(raw);
      if (!isNaN(dp.getTime())) raw = dp.getFullYear() + '-' + ('0'+(dp.getMonth()+1)).slice(-2) + '-' + ('0'+dp.getDate()).slice(-2);
    }
    return raw > since;
  }).sort((a, b) => String(a.date).localeCompare(String(b.date)));

  if (billable.length >= interval) {
    // Send WhatsApp notification via CallMeBot (works for all roles, no modal)
    _sendBillingWhatsApp(client, billable.length);
  }
}

/** Send billing WhatsApp notification via CallMeBot — once per billing cycle */
function _sendBillingWhatsApp(client, count) {
  var sentKey = 'billing_wa_sent_' + client.client_id;
  var cycleMarker = (client.last_billing_date || 'initial') + '_' + count;
  getSetting(sentKey).then(function(alreadySent) {
    if (alreadySent === cycleMarker) return;
    return Promise.all([getSetting('wa_phone'), getSetting('wa_apikey')]).then(function(vals) {
      var phone = vals[0], apikey = vals[1];
      if (!phone || !apikey) {
        console.warn('[WA] WhatsApp neconfigurat. Mergi la Settings.');
        showToast('WhatsApp neconfigurat! Mergi la Settings.', 'warning');
        return;
      }
      var msg = '*Facturare: ' + client.name + '*\n'
        + count + ' interventii nefacturate\n'
        + 'Tel: ' + (client.phone || '-') + '\n'
        + 'Adresa: ' + (client.address || '-') + '\n'
        + 'Data: ' + new Date().toLocaleDateString('ro-RO') + '\n'
        + '_Generat de Pool Manager_';

      var url = 'https://api.callmebot.com/whatsapp.php'
        + '?phone=' + encodeURIComponent(phone)
        + '&text=' + encodeURIComponent(msg)
        + '&apikey=' + encodeURIComponent(apikey);

      fetch(url, { mode: 'no-cors' }).then(function() {
        showToast('WhatsApp trimis!', 'success');
        setSetting(sentKey, cycleMarker);
      }).catch(function(e) {
        console.warn('[WA] Error:', e.message);
        showToast('WhatsApp nereusit: ' + e.message, 'warning');
      });
    });
  });
}

/** Show billing modal with deviz actions */
function showBillingModal(client, interventions) {
  var modal = $('modal-billing');
  if (!modal) return;

  APP._billingClientId = client.client_id;
  APP._billingInterventions = interventions;
  APP._billingClient = client;

  var title = $('billing-modal-title');
  if (title) title.innerHTML = '&#128176; Facturare: ' + escHtml(client.name);

  var body = $('billing-modal-body');
  if (!body) return;

  var since = client.last_billing_date || null;
  var sinceLabel = since ? fmtDate(since) : 'prima interven\u021bie';
  var today = new Date().toLocaleDateString('ro-RO', { day: '2-digit', month: '2-digit', year: 'numeric' });

  var html = '<div class="billing-summary">';
  html += '<strong>' + interventions.length + ' interven\u021bii</strong> din ' + sinceLabel + ' p\u00e2n\u0103 azi (' + today + ')';
  html += '</div>';

  // Mini table
  html += '<table class="billing-table"><thead><tr><th>Nr.</th><th>Data</th><th>Tehnician</th><th>Produse</th></tr></thead><tbody>';
  interventions.forEach(function(inv, idx) {
    html += '<tr>';
    html += '<td>' + (idx + 1) + '</td>';
    html += '<td>' + escHtml(inv.date) + '</td>';
    html += '<td>' + escHtml(inv.technician_name || '') + '</td>';
    html += '<td style="font-size:.75rem">' + escHtml(_fmtTreatShort(inv)) + '</td>';
    html += '</tr>';
  });
  html += '</tbody></table>';

  // Action buttons
  html += '<div class="billing-actions">';
  html += '<button class="billing-action-btn" onclick="generateBillingExcel()"><span style="font-size:1.3rem">\ud83d\udcca</span><div><strong>Deviz Excel</strong><small>Desc\u0103rca\u021bi fi\u0219ier .xlsx</small></div></button>';
  html += '<button class="billing-action-btn" onclick="generateBillingPdf()"><span style="font-size:1.3rem">\ud83d\udda8\ufe0f</span><div><strong>Deviz PDF</strong><small>Deschide pentru print</small></div></button>';
  if (client.phone) {
    html += '<button class="billing-action-btn" onclick="shareBillingWhatsApp()"><span style="font-size:1.3rem">\ud83d\udcac</span><div><strong>Trimite WhatsApp</strong><small>Rezumat pe WhatsApp</small></div></button>';
  }
  html += '</div>';

  // Bottom actions
  html += '<div style="display:flex;gap:8px;margin-top:14px">';
  html += '<button class="btn-primary" style="flex:1" onclick="billingMarkAndClose()">\u2713 Marcheaz\u0103 facturat</button>';
  html += '<button style="flex:0 0 auto;padding:8px 18px;border-radius:10px;background:var(--slate-200);color:var(--slate-600);font-weight:600" onclick="closeBillingModal()">Mai t\u00e2rziu</button>';
  html += '</div>';

  body.innerHTML = html;
  modal.classList.add('open');
}

function closeBillingModal() {
  var modal = $('modal-billing');
  if (modal) modal.classList.remove('open');
}

async function billingMarkAndClose() {
  var clientId = APP._billingClientId;
  if (!clientId) return;
  var client = APP.clients.find(function(c) { return c.client_id === clientId; });
  if (!client) return;

  client.last_billing_date = new Date().toISOString().split('T')[0];
  client.updated_at = new Date().toISOString();
  await put('clients', client);
  APP.clients = APP.clients.map(function(c) { return c.client_id === clientId ? client : c; });

  if (isSyncConfigured()) {
    apiFetch(SYNC_CONFIG.API_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'push', type: 'clients', data: [client] })
    }).catch(function(e) { console.warn('[SYNC] Billing push failed:', e.message); });
  }

  closeBillingModal();
  showToast('\u2713 ' + client.name + ' marcat ca facturat.', 'success');

  var billBtn = $('btn-mark-billed');
  if (billBtn) billBtn.style.display = 'none';
}

/** Short treatment summary for billing table */
function _fmtTreatShort(inv) {
  var parts = [];
  if (inv.treat_cl_granule_gr > 0) parts.push('Cl:' + inv.treat_cl_granule_gr + 'g');
  if (inv.treat_cl_tablete > 0) parts.push('ClTab:' + inv.treat_cl_tablete);
  if (inv.treat_ph_granule > 0) parts.push('pH:' + inv.treat_ph_granule + 'kg');
  if (inv.treat_antialgic > 0) parts.push('Anti:' + inv.treat_antialgic + 'L');
  if (inv.treat_anticalcar > 0) parts.push('Antical:' + inv.treat_anticalcar + 'L');
  if (inv.treat_floculant > 0) parts.push('Floc:' + inv.treat_floculant + 'L');
  if (inv.treat_sare_saci > 0) parts.push('Sare:' + inv.treat_sare_saci);
  if (inv.treat_bicarbonat > 0) parts.push('Bicarb:' + inv.treat_bicarbonat + 'kg');
  // Dynamic stock products
  if (typeof APP !== 'undefined' && APP._stockProducts) {
    APP._stockProducts.forEach(function(p) {
      var val = inv['treat_' + p.product_id];
      if (val > 0 && !parts.some(function(x) { return x.indexOf(p.name.slice(0,6)) === 0; })) {
        parts.push(p.name.slice(0,10) + ':' + val + (p.unit || ''));
      }
    });
  }
  return parts.join(', ') || '\u2014';
}

/** Generate billing Excel */
function generateBillingExcel() {
  var client = APP._billingClient;
  var interventions = APP._billingInterventions;
  if (!client || !interventions) return;
  var devizType = parseInt(client.deviz_type) || 2;
  if (devizType === 2) {
    exportDevizComplet(client, interventions);
  } else {
    exportDevizChimicale(client, interventions);
  }
}

/** Generate billing PDF */
function generateBillingPdf() {
  var client = APP._billingClient;
  var interventions = APP._billingInterventions;
  if (!client || !interventions) return;

  var printHtml = _buildBillingPrintHtml(client, interventions);
  var w = window.open('', '_blank');
  if (!w) { showToast('Popup blocat. Permite popups pentru acest site.', 'error'); return; }
  w.document.write(printHtml);
  w.document.close();
  setTimeout(function() { w.print(); }, 400);
}

/** Build billing PDF HTML (A4 print-ready) */
function _buildBillingPrintHtml(client, interventions) {
  var since = client.last_billing_date || '';
  var today = new Date().toISOString().split('T')[0];
  var devizNr = 'D-' + today.replace(/-/g, '') + '-' + (client.client_id || '').slice(-4);
  var totals = calcTotals(interventions);
  var totalMin = interventions.reduce(function(s, i) { return s + (i.duration_minutes || 0); }, 0);

  var rows = '';
  interventions.forEach(function(inv, idx) {
    rows += '<tr>';
    rows += '<td style="text-align:center">' + (idx + 1) + '</td>';
    rows += '<td>' + escHtml(inv.date) + '</td>';
    rows += '<td>' + escHtml(inv.technician_name || '') + '</td>';
    rows += '<td style="font-size:11px">' + escHtml(_fmtTreatFull(inv)) + '</td>';
    rows += '<td style="text-align:center">' + (inv.duration_minutes != null ? Math.round(inv.duration_minutes) : '-') + '</td>';
    rows += '<td style="font-size:10px">' + escHtml(inv.observations || '') + '</td>';
    rows += '</tr>';
  });

  // Totals row for products
  var prodSummary = [];
  if (totals.cl_granule_gr) prodSummary.push('Cl granule: ' + totals.cl_granule_gr + ' gr');
  if (totals.cl_tablete) prodSummary.push('Cl tablete: ' + totals.cl_tablete + ' buc');
  if (totals.ph_granule) prodSummary.push('pH granule: ' + totals.ph_granule + ' kg');
  if (totals.antialgic) prodSummary.push('Antialgic: ' + totals.antialgic + ' L');
  if (totals.anticalcar) prodSummary.push('Anticalcar: ' + totals.anticalcar + ' L');
  if (totals.floculant) prodSummary.push('Floculant: ' + totals.floculant + ' L');
  if (totals.sare) prodSummary.push('Sare: ' + totals.sare + ' saci');
  if (totals.bicarbonat) prodSummary.push('Bicarbonat: ' + totals.bicarbonat + ' kg');

  return '<!DOCTYPE html><html lang="ro"><head><meta charset="UTF-8"><title>Deviz ' + escHtml(client.name) + '</title>'
    + '<style>'
    + '* { box-sizing: border-box; margin: 0; padding: 0; }'
    + 'body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #111; padding: 30px; max-width: 210mm; margin: 0 auto; }'
    + 'h1 { font-size: 18px; color: #1d4ed8; margin-bottom: 4px; }'
    + '.header { display: flex; justify-content: space-between; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 2px solid #1d4ed8; }'
    + '.header-left { line-height: 1.6; }'
    + '.header-right { text-align: right; line-height: 1.6; }'
    + '.label { color: #64748b; font-size: 11px; }'
    + 'table { width: 100%; border-collapse: collapse; margin-top: 16px; }'
    + 'th { background: #1d4ed8; color: #fff; padding: 7px 8px; font-size: 12px; text-align: left; }'
    + 'td { padding: 6px 8px; border-bottom: 1px solid #e2e8f0; font-size: 12px; vertical-align: top; }'
    + 'tr:nth-child(even) td { background: #f8fafc; }'
    + '.totals { margin-top: 16px; padding: 12px; background: #eff6ff; border-radius: 6px; font-size: 12px; line-height: 1.7; }'
    + '.totals strong { color: #1d4ed8; }'
    + '.footer { margin-top: 30px; font-size: 10px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 8px; }'
    + '@media print { body { padding: 15px; } .footer { position: fixed; bottom: 10px; left: 0; right: 0; } }'
    + '</style></head><body>'
    + '<h1>DEVIZ SERVICII PISCIN\u0102</h1>'
    + '<div class="header">'
    + '<div class="header-left">'
    + '<div><span class="label">Client:</span> <strong>' + escHtml(client.name) + '</strong></div>'
    + (client.address ? '<div><span class="label">Adres\u0103:</span> ' + escHtml(client.address) + '</div>' : '')
    + (client.phone ? '<div><span class="label">Telefon:</span> ' + escHtml(client.phone) + '</div>' : '')
    + '<div><span class="label">Volum piscin\u0103:</span> ' + (client.pool_volume_mc || '-') + ' m\u00b3 (' + (client.pool_type || '-') + ')</div>'
    + '</div>'
    + '<div class="header-right">'
    + '<div><span class="label">Nr. deviz:</span> <strong>' + devizNr + '</strong></div>'
    + '<div><span class="label">Data:</span> ' + today + '</div>'
    + '<div><span class="label">Perioada:</span> ' + (since || '-') + ' \u2013 ' + today + '</div>'
    + '</div>'
    + '</div>'
    + '<table><thead><tr><th>Nr.</th><th>Data</th><th>Tehnician</th><th>Produse utilizate</th><th>Durata</th><th>Observa\u021bii</th></tr></thead>'
    + '<tbody>' + rows + '</tbody></table>'
    + '<div class="totals">'
    + '<strong>Total: ' + interventions.length + ' interven\u021bii</strong> \u00b7 Durat\u0103 total\u0103: ' + totalMin + ' min<br>'
    + (prodSummary.length ? '<strong>Produse consumate:</strong> ' + prodSummary.join(' \u00b7 ') : '')
    + '</div>'
    + '<div class="footer">Generat de Pool Manager \u00b7 ' + today + '</div>'
    + '</body></html>';
}

/** Full treatment summary for PDF */
function _fmtTreatFull(inv) {
  var parts = [];
  if (inv.treat_cl_granule_gr > 0) parts.push('Cl granule: ' + inv.treat_cl_granule_gr + 'g');
  if (inv.treat_cl_tablete > 0) parts.push('Cl tablete: ' + inv.treat_cl_tablete + ' buc');
  if (inv.treat_cl_lichid_bidoane > 0) parts.push('Cl lichid: ' + inv.treat_cl_lichid_bidoane + ' bid');
  if (inv.treat_ph_granule > 0) parts.push('pH: ' + inv.treat_ph_granule + 'kg');
  if (inv.treat_ph_lichid_bidoane > 0) parts.push('pH lichid: ' + inv.treat_ph_lichid_bidoane + ' bid');
  if (inv.treat_antialgic > 0) parts.push('Antialgic: ' + inv.treat_antialgic + 'L');
  if (inv.treat_anticalcar > 0) parts.push('Anticalcar: ' + inv.treat_anticalcar + 'L');
  if (inv.treat_floculant > 0) parts.push('Floculant: ' + inv.treat_floculant + 'L');
  if (inv.treat_sare_saci > 0) parts.push('Sare: ' + inv.treat_sare_saci + ' saci');
  if (inv.treat_bicarbonat > 0) parts.push('Bicarbonat: ' + inv.treat_bicarbonat + 'kg');
  return parts.join(', ') || '\u2014';
}

/** Share billing summary via WhatsApp */
function shareBillingWhatsApp() {
  var client = APP._billingClient;
  var interventions = APP._billingInterventions;
  if (!client || !interventions) return;

  var totals = calcTotals(interventions);
  var since = client.last_billing_date || '';
  var today = new Date().toISOString().split('T')[0];

  var text = '*Rezumat servicii piscin\u0103*\n\n';
  text += '*Client:* ' + client.name + '\n';
  text += '*Perioada:* ' + (since || '-') + ' \u2013 ' + today + '\n';
  text += '*Total interven\u021bii:* ' + interventions.length + '\n\n';
  text += '*Produse consumate:*\n';
  if (totals.cl_granule_gr) text += '\u2022 Cl granule: ' + totals.cl_granule_gr + ' gr\n';
  if (totals.cl_tablete) text += '\u2022 Cl tablete: ' + totals.cl_tablete + ' buc\n';
  if (totals.ph_granule) text += '\u2022 pH granule: ' + totals.ph_granule + ' kg\n';
  if (totals.antialgic) text += '\u2022 Antialgic: ' + totals.antialgic + ' L\n';
  if (totals.anticalcar) text += '\u2022 Anticalcar: ' + totals.anticalcar + ' L\n';
  if (totals.floculant) text += '\u2022 Floculant: ' + totals.floculant + ' L\n';
  if (totals.sare) text += '\u2022 Sare: ' + totals.sare + ' saci\n';
  if (totals.bicarbonat) text += '\u2022 Bicarbonat: ' + totals.bicarbonat + ' kg\n';
  text += '\n_Pool Manager_';

  var phone = client.phone ? '4' + client.phone.replace(/\D/g, '').slice(-9) : '';
  var url = phone
    ? 'https://wa.me/' + phone + '?text=' + encodeURIComponent(text)
    : 'https://wa.me/?text=' + encodeURIComponent(text);
  window.open(url, '_blank');
}

// ════════════════════════════════════════════════════════════════
// DRUM PICKER — popover inline lângă input
// ════════════════════════════════════════════════════════════════

const DRUM_ITEM_H = 44; // px per item
const DRUM_PAD_H  = 132; // 3 items padding top/bottom so first/last can center

let _drumInput = null;
let _drumJustClosed = null; // input that just closed — prevents same-button toggle re-open

// ── Info Edit Mode ──────────────────────────────────────────────
let _infoEditMode   = false;
let _infoDefaultHTML = null;  // captured once before any injection — used for Reset
let _infoPreEditHTML = {};    // snapshot before entering edit — used for Cancel

function openDrumPicker(inputEl) {
  // Toggle: if drum already open for this input, close it
  const _popup = $('drum-popup');
  if (_drumInput === inputEl && _popup && _popup.style.display !== 'none') {
    confirmDrumPicker();
    return;
  }
  // Prevent toggling the same input open again immediately after close
  if (inputEl === _drumJustClosed) return;

  // Dismiss keyboard immediately (important on mobile)
  inputEl.blur();

  _drumInput = inputEl;
  const step   = parseFloat(inputEl.dataset.step || inputEl.step) || 1;
  const curVal = parseFloat(inputEl.value) || (parseFloat(inputEl.min) || 0);
  const unit   = inputEl.dataset.unit  || '';
  const label  = inputEl.dataset.label || inputEl.dataset.label || '';

  // Build value list: start from min, generate plenty of values
  // We intentionally IGNORE inputEl.max so the user can scroll beyond normal limits
  const minVal   = parseFloat(inputEl.min) || 0;
  const maxCount = Math.max(100, Math.ceil((curVal - minVal) / step) + 40);

  const values = [];
  for (let i = 0; i <= maxCount; i++) {
    values.push(Math.round((minVal + i * step) * 10000) / 10000);
  }

  // Render items inside viewport
  const viewport = $('drum-popup-viewport');
  const dec = step < 0.1 ? 2 : step < 1 ? 1 : 0;
  viewport.innerHTML =
    `<div style="height:${DRUM_PAD_H}px;flex-shrink:0"></div>` +
    values.map((v, i) => {
      const disp = Number.isInteger(v) ? String(v) : v.toFixed(dec);
      return `<div class="drum-popup-item" data-index="${i}" data-value="${v}" onclick="_drumItemClick(${i})">${disp}${unit ? '<small class="drum-unit"> ' + unit + '</small>' : ''}</div>`;
    }).join('') +
    `<div style="height:${DRUM_PAD_H}px;flex-shrink:0"></div>`;

  // Set label
  const lbl = $('drum-popup-label');
  if (lbl) lbl.textContent = label || '';
  if (lbl) lbl.style.display = label ? '' : 'none';

  // Position popup near input
  const popup = $('drum-popup');
  const rect  = inputEl.getBoundingClientRect();
  const popupW = 200;
  const popupH = label ? 365 : 340; // label adds ~25px

  // Horizontal: center on input, clamp to viewport
  let left = rect.left + rect.width / 2 - popupW / 2;
  left = Math.max(8, Math.min(left, window.innerWidth - popupW - 8));
  popup.style.left = left + 'px';

  // Vertical: prefer below, fallback above
  const spaceBelow = window.innerHeight - rect.bottom - 8;
  if (spaceBelow >= popupH) {
    popup.style.top  = (rect.bottom + 6) + 'px';
    popup.style.bottom = 'auto';
  } else {
    popup.style.top  = Math.max(8, rect.top - popupH - 6) + 'px';
    popup.style.bottom = 'auto';
  }

  popup.style.display = 'block';

  // Scroll to current value
  const idx = values.findIndex(v => Math.abs(v - curVal) < step * 0.5);
  requestAnimationFrame(() => {
    viewport.scrollTop = (idx >= 0 ? idx : 0) * DRUM_ITEM_H;
    _updateDrumHighlight();
  });
}

function onDrumScroll() {
  const viewport = $('drum-popup-viewport');
  if (!viewport) return;

  // Update highlight immediately so selected item always appears in center zone
  _updateDrumHighlight();

  // After scroll settles: snap to nearest item
  clearTimeout(viewport._t);
  viewport._t = setTimeout(() => {
    const idx = Math.round(viewport.scrollTop / DRUM_ITEM_H);
    viewport.scrollTop = idx * DRUM_ITEM_H;
    _updateDrumHighlight();
  }, 120);
}

function _updateDrumHighlight() {
  const viewport = $('drum-popup-viewport');
  if (!viewport) return;
  const idx = Math.round(viewport.scrollTop / DRUM_ITEM_H);
  $$('#drum-popup-viewport .drum-popup-item').forEach((el, i) => {
    el.classList.toggle('selected', i === idx);
  });
}

function confirmDrumPicker() {
  if (!_drumInput) { closeDrumPicker(); return; }
  const viewport = $('drum-popup-viewport');
  const idx = Math.round(viewport.scrollTop / DRUM_ITEM_H);
  const items = $$('#drum-popup-viewport .drum-popup-item');
  if (items[idx]) {
    _drumInput.value = parseFloat(items[idx].dataset.value) ?? 0;
    // Trigger input event so any listeners update
    _drumInput.dispatchEvent(new Event('input', { bubbles: true }));
  }
  closeDrumPicker();
}

function closeDrumPicker() {
  const popup = $('drum-popup');
  if (popup) popup.style.display = 'none';
  _drumJustClosed = _drumInput; // remember which input closed (to prevent same-button toggle re-open)
  _drumInput = null;
  setTimeout(() => { _drumJustClosed = null; }, 150);
}

function _drumItemClick(idx) {
  // Click on a drum item: scroll to it, select value and confirm
  const viewport = $('drum-popup-viewport');
  if (!viewport || !_drumInput) return;
  viewport.scrollTop = idx * DRUM_ITEM_H;
  _updateDrumHighlight();
  // Small delay to let user see the selection, then confirm
  setTimeout(() => confirmDrumPicker(), 120);
}

// Click outside popup → confirm and close
document.addEventListener('click', function _drumOutside(e) {
  const popup = $('drum-popup');
  if (!popup || popup.style.display === 'none') return;
  if (popup.contains(e.target)) return;
  if (_drumInput && _drumInput.contains(e.target)) return;
  confirmDrumPicker();
}, true);

/** Update intervention date from picker. */
function updateInterventionDate(val) {
  if (!val) return;
  APP._interventionDate = val;
  var d = new Date(val + 'T12:00:00');
  var dateEl = $('intervention-date');
  if (dateEl) {
    dateEl.textContent = d.toLocaleDateString('ro-RO', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    // Highlight if not today
    var today = new Date().toISOString().split('T')[0];
    if (val !== today) {
      dateEl.style.background = 'rgba(251,191,36,0.3)';
      dateEl.style.padding = '2px 8px';
      dateEl.style.borderRadius = '6px';
    } else {
      dateEl.style.background = '';
      dateEl.style.padding = '';
      dateEl.style.borderRadius = '';
    }
  }
}
// ─────────────────────────────────────────────────────────────
// ── EVIDENȚĂ CHECKLIST ────────────────────────────────────────
// ─────────────────────────────────────────────────────────────

let _checklistItems = [];
let _checklistTitle = '';

/** Încarcă lista salvată în IndexedDB și randează ecranul. */
async function loadChecklistScreen() {
  const title = await getSetting('checklist_title');
  const items = await getSetting('checklist_items');

  _checklistTitle = title || '';
  try { _checklistItems = items ? JSON.parse(items) : []; }
  catch { _checklistItems = []; }

  renderChecklist();
  // Fetch from GAS in background — re-renders if remote is newer
  _fetchChecklistFromGas();
}

/** Randează tabelul checklist în tbody. */
function renderChecklist() {
  const titleEl = $('checklist-title');
  if (titleEl) titleEl.textContent = _checklistTitle || 'Nicio listă importată';

  const tbody = $('checklist-tbody');
  if (!tbody) return;

  const isAdmin = APP.user && APP.user.role === 'admin';

  if (_checklistItems.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="cl-empty">Importați un fișier Excel folosind butonul 📥 din header.</td></tr>';
    _updateChecklistCounter();
    return;
  }

  tbody.innerHTML = _checklistItems.map(item => {
    const id = escHtml(item.id);
    return `<tr class="cl-row${item.checked ? ' cl-checked' : ''}" data-id="${id}">
      <td class="cl-cell-f" onclick="toggleChecklistF('${id}')">
        <span class="cl-f-btn${item.f_marked ? ' cl-f-active' : ''}">${item.f_marked ? 'F' : '○'}</span>
      </td>
      <td class="cl-cell-name">${escHtml(item.name)}</td>
      <td class="cl-cell-value">${escHtml(item.value)}</td>
      <td class="cl-cell-check">
        <label class="cl-chk-wrap">
          <input type="checkbox" ${item.checked ? 'checked' : ''}
                 onchange="toggleChecklistItem('${id}', this.checked)">
          <span class="cl-chkmark"></span>
        </label>
      </td>
      <td class="cl-cell-del${isAdmin ? '' : ' admin-only'}">
        ${isAdmin ? `<button class="cl-del-btn" onclick="deleteChecklistItem('${id}')" title="Șterge rândul">✕</button>` : ''}
      </td>
    </tr>`;
  }).join('');

  _updateChecklistCounter();
}

/** Actualizează contorul "X din Y bifate". */
function _updateChecklistCounter() {
  const el = $('checklist-counter');
  if (!el) return;
  const total   = _checklistItems.length;
  const checked = _checklistItems.filter(i => i.checked).length;
  el.textContent = total > 0 ? `✅ ${checked} din ${total} bifate` : '';
}

/** Toggle stare bifat pe un rând. Actualizează UI fără re-render complet. */
async function toggleChecklistItem(id, checked) {
  const item = _checklistItems.find(i => i.id === id);
  if (!item) return;
  item.checked = checked;
  const row = document.querySelector(`.cl-row[data-id="${id}"]`);
  if (row) row.classList.toggle('cl-checked', checked);
  _updateChecklistCounter();
  await _saveChecklist();
}

/** Toggle marcaj "F" pe un rând. */
async function toggleChecklistF(id) {
  const item = _checklistItems.find(i => i.id === id);
  if (!item) return;
  item.f_marked = !item.f_marked;
  const btn = document.querySelector(`.cl-row[data-id="${id}"] .cl-f-btn`);
  if (btn) {
    btn.textContent = item.f_marked ? 'F' : '○';
    btn.classList.toggle('cl-f-active', item.f_marked);
  }
  await _saveChecklist();
}

async function _saveChecklist() {
  const updatedAt = new Date().toISOString();
  await setSetting('checklist_title', _checklistTitle);
  await setSetting('checklist_items', JSON.stringify(_checklistItems));
  await setSetting('checklist_updated_at', updatedAt);
  if (isSyncConfigured()) {
    _syncChecklistToGas(updatedAt).catch(err =>
      console.warn('[CHECKLIST] GAS sync failed:', err.message)
    );
  }
}

/** Trimite starea curenta a checklistului la Google Sheets. */
async function _syncChecklistToGas(updatedAt) {
  return apiFetch(SYNC_CONFIG.API_URL, {
    method: 'POST',
    body: JSON.stringify({
      action:     'saveChecklist',
      title:      _checklistTitle,
      items_json: JSON.stringify(_checklistItems),
      updated_at: updatedAt
    })
  }).then(data => {
    if (!data.success) console.warn('[CHECKLIST] GAS sync error:', data.error);
  });
}

/** Preia datele checklistului din GAS si le aplica daca sunt mai noi. */
async function _fetchChecklistFromGas() {
  if (!isSyncConfigured()) return;
  try {
    const data = await apiFetch(SYNC_CONFIG.API_URL + '?action=getChecklist');
    if (!data.success || !data.data || !data.data.items_json) return;

    const remote = data.data;
    const localUpdatedAt  = (await getSetting('checklist_updated_at')) || '';
    const remoteUpdatedAt = remote.updated_at || '';

    if (remoteUpdatedAt > localUpdatedAt) {
      _checklistTitle = remote.title || '';
      try { _checklistItems = JSON.parse(remote.items_json) || []; }
      catch { _checklistItems = []; }
      await setSetting('checklist_title', _checklistTitle);
      await setSetting('checklist_items', remote.items_json);
      await setSetting('checklist_updated_at', remoteUpdatedAt);
      renderChecklist();
    }
  } catch (err) {
    console.warn('[CHECKLIST] Fetch from GAS failed:', err.message);
  }
}

/** Importă fișier Excel și înlocuiește lista curentă. */
async function onChecklistFileImport(file) {
  if (!file) return;
  const inp = $('checklist-import-input');
  if (inp) inp.value = '';

  try { await loadXLSX(); } catch (e) {
    showToast('SheetJS nu este disponibil. Reconectați-vă la internet.', 'warning');
    return;
  }
  showToast('Se procesează fișierul...', 'info', 4000);

  try {
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array', raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Găsește rândul de header (cel care conține "NUME")
    let headerRow = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      if (rows[i].some(c => String(c).trim().toUpperCase() === 'NUME')) {
        headerRow = i;
        break;
      }
    }
    if (headerRow === -1) {
      showToast('Nu am găsit coloana "NUME" în fișier. Verificați formatul.', 'error');
      return;
    }

    // Mapează coloanele după header
    const hdr    = rows[headerRow].map(c => String(c).trim().toUpperCase());
    const colF   = hdr.indexOf('F');
    const colN   = hdr.indexOf('NUME');
    const colV   = hdr.findIndex(h => h === 'VALOARE' || h === 'VAL' || h === 'SUMA' || h === 'SUMA (LEI)');

    // Titlu: primul rând nevid înainte de header, sau numele fișierului
    let title = file.name.replace(/\.[^.]+$/, '');
    for (let i = 0; i < headerRow; i++) {
      const t = rows[i].map(c => String(c).trim()).filter(Boolean).join(' ').trim();
      if (t.length > 3) { title = t; break; }
    }

    // Parsează rândurile de date
    const now   = Date.now();
    const items = [];
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row  = rows[i];
      const name = String(row[colN] || '').trim();
      if (!name) continue;
      const val  = colV >= 0 ? String(row[colV] || '').trim() : '';
      const fVal = colF >= 0 ? String(row[colF] || '').trim().toUpperCase() : '';
      items.push({
        id:       'cl_' + now + '_' + i,
        row_order: i - headerRow,
        f_marked: fVal === 'F',
        name,
        value:    val,
        checked:  false
      });
    }

    if (items.length === 0) {
      showToast('Nicio intrare validă în fișier.', 'error');
      return;
    }

    _checklistTitle = title;
    _checklistItems = items;
    await _saveChecklist();
    showToast(`Import reușit: ${items.length} rânduri din "${title}".`, 'success', 5000);
    renderChecklist();
  } catch (err) {
    showToast('Eroare la procesare: ' + err.message, 'error');
  }
}

/** Descarcă template Excel pentru import evidență. */
async function downloadChecklistTemplate() {
  try { await loadXLSX(); } catch (e) {
    showToast('SheetJS nu este disponibil. Reconectați-vă la internet.', 'warning');
    return;
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['INCASARI ' + new Date().toLocaleDateString('ro-RO')],
    ['F', 'NUME', 'VALOARE', 'OBS'],
    ['',  'Adrian Driga',   '2475 lei',  'Achitat'],
    ['F', 'Barbulescu',     '2728 lei',  'Achitat'],
    ['',  'Bogdan Azur',    '3320 lei',  ''],
  ]);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }]; // merge titlu
  ws['!cols']   = [{ wch: 4 }, { wch: 26 }, { wch: 16 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Incasari');
  XLSX.writeFile(wb, 'template-evidenta.xlsx');
}

/** Șterge un singur rând din checklist (admin only). */
async function deleteChecklistItem(id) {
  _checklistItems = _checklistItems.filter(i => i.id !== id);
  await _saveChecklist();
  renderChecklist();
}

/** Resetează toate bifele și marcajele F, dar păstrează lista. */
async function resetChecklist() {
  if (!_checklistItems.length) return;
  if (!confirm('Resetezi toate bifele și marcajele F? Lista de nume rămâne neschimbată.')) return;
  _checklistItems.forEach(i => { i.checked = false; i.f_marked = false; });
  await _saveChecklist();
  renderChecklist();
  showToast('Lista a fost resetată.', 'success');
}

/** Șterge lista curentă după confirmare. */
async function clearChecklist() {
  if (!confirm('Ștergeți toată lista curentă? Starea bifatelor se va pierde.')) return;
  _checklistItems = [];
  _checklistTitle = '';
  await _saveChecklist();
  renderChecklist();
  showToast('Lista a fost ștearsă.', 'success');
}

// ── Swipe-back gesture on all non-dashboard screens ─────────────────
(function() {
  var _swipeStartX = 0, _swipeStartY = 0, _swipeTracking = false;
  var SWIPEABLE = ['info', 'checklist', 'intervention', 'success', 'billing-list'];

  document.addEventListener('touchstart', function(e) {
    if (SWIPEABLE.indexOf(APP.currentScreen) < 0) return;
    var t = e.touches[0];
    if (t.clientX <= 40) {
      _swipeStartX = t.clientX;
      _swipeStartY = t.clientY;
      _swipeTracking = true;
    }
  }, { passive: true });

  document.addEventListener('touchend', function(e) {
    if (!_swipeTracking) return;
    _swipeTracking = false;
    var t = e.changedTouches[0];
    var dx = t.clientX - _swipeStartX;
    var dy = Math.abs(t.clientY - _swipeStartY);
    if (dx > 80 && dy < dx * 0.6) {
      // Cleanup for intervention screen
      if (APP.currentScreen === 'intervention') APP._editingIntervention = null;
      showScreen('dashboard');
    }
  }, { passive: true });
})();

