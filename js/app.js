// app.js — Pool Manager PWA — Main application logic
// Features: Login + PIN, Dashboard, Intervention form, GPS, Toast, Time tracking

'use strict';

// ── Helpers ───────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const $q = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);
const uid = () => 'i_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);

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
  currentPosition: null,       // GPS: {lat, lng, accuracy}
  arrivalTime:     null,       // ISO string when form opened
  pinBuffer:       '',         // PIN input buffer
  installPrompt:   null,       // beforeinstallprompt event
  alertShown:      false,      // toast de alertă intervenții (1x per sesiune)
  alertThreshold:  4,          // prag configurabil (default 4)
  dashboardTab:    'all',      // 'all' | 'due'
  clientFormMode:  'add',      // 'add' | 'edit'
  wizardStep:      1,          // 1 | 2 | 3 — pasul curent al wizard-ului intervenție
  _stockProducts:  [],         // cache produse stoc (actualizat la deschidere formular)
  _billingClientId: null,      // client_id pentru care se afișează butonul "Marchează facturat"
  gpsStart:        7,          // ora de start tracking automat (0–23)
  gpsEnd:          18          // ora de stop tracking automat (0–23)
};

// ── Init ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupConnectivityIndicator();
  setupInstallPrompt();
  initApp();
});

async function initApp() {
  await openDB();

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

  // Seed demo if DB empty
  const techCount = await count('technicians');
  if (techCount === 0) {
    await seedDemoData();
    showToast('Date demo încărcate. Login: admin / admin123', 'info', 5000);
  }

  // Register service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
      .then(reg => console.log('[SW] Registered:', reg.scope))
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

  // Map: stop refresh when leaving, start when entering
  if (APP.currentScreen === 'map' && name !== 'map') {
    if (_mapRefreshInterval) { clearInterval(_mapRefreshInterval); _mapRefreshInterval = null; }
  }
  if (name === 'map') {
    loadMapScreen(); // async — loads Leaflet lazily, starts 60s refresh
  }

  // Calendar: load current week
  if (name === 'calendar') {
    loadCalendarScreen();
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
  toast.innerHTML = `<span class="toast-icon">${icons[type] || 'ℹ'}</span><span>${msg}</span>`;
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
      badge.textContent = '🟢 Online';
      badge.className = 'conn-badge online';
    } else {
      badge.textContent = '🔴 Offline';
      badge.className = 'conn-badge offline';
    }
  }
  window.addEventListener('online',  () => { update(); forceSync().catch(() => {}); });
  window.addEventListener('offline', () => { update(); showToast('Conexiune pierdută. Datele se salvează local.', 'warning'); });
  update();

  // Sync callbacks
  window.onSyncComplete = () => {
    updateSyncBadge();
    if (APP.currentScreen === 'dashboard') loadData().then(renderDashboard);
  };
  window.onSyncError = () => {};
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
        // Try index lookup (fastest)
        tech = await getByIndexFirst('technicians', 'username', username);
      } catch {
        // Index might not exist in old DB — scan all technicians
        const all = await getAll('technicians');
        tech = all.find(t => t.username === username) || null;
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

  // Încarcă orele de program GPS în APP (înainte de renderDashboard)
  const gpsStart = await getSetting('gps_schedule_start');
  const gpsEnd   = await getSetting('gps_schedule_end');
  APP.gpsStart = parseInt(gpsStart ?? '7',  10);
  APP.gpsEnd   = parseInt(gpsEnd   ?? '18', 10);

  await loadData();
  renderDashboard();
  showScreen('dashboard');
  updateSyncBadge();
  // Start GPS tracking (sends position every configured interval to GAS if API configured)
  startLocationTracking();
  // QR deeplink: ?client=ID
  setTimeout(checkClientDeeplink, 200);
}

// ── PIN Screen ───────────────────────────────────────────────
function showPinScreen(user) {
  showScreen('pin');
  APP.pinBuffer = '';
  renderPinDots();
  $('pin-username-label').textContent = 'Bine ai venit, ' + user.name;

  // Setup keypad
  $$('.pin-key').forEach(btn => {
    btn.addEventListener('click', () => {
      const val = btn.dataset.val;
      if (val === 'del') {
        APP.pinBuffer = APP.pinBuffer.slice(0, -1);
      } else if (APP.pinBuffer.length < 4) {
        APP.pinBuffer += val;
        if (APP.pinBuffer.length === 4) checkPin(user);
      }
      renderPinDots();
    });
  });

  const switchBtn = $('pin-switch-user');
  if (switchBtn) {
    switchBtn.addEventListener('click', async () => {
      await clearSession();
      APP.user = null;
      showScreen('login');
      initLoginScreen();
    });
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
  APP.interventions = interventions;
  APP.pendingSync   = interventions.filter(i => !i.synced).length;
}

// ── Dashboard ────────────────────────────────────────────────
function isAdmin() {
  return APP.user && APP.user.role === 'admin';
}

function renderDashboard() {
  if (!APP.user) return;

  // Apply role class on <body> — drives all .admin-only visibility via CSS
  document.body.classList.toggle('role-admin',      isAdmin());
  document.body.classList.toggle('role-technician', !isAdmin());

  // Actualizează butonul GPS din footer
  updateGpsToggleBtn();

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

  updateSyncBadge();
  renderClientList('');
  renderAdminStats();

  // Search
  const searchInput = $('search-input');
  if (searchInput) {
    searchInput.value = '';
    searchInput.oninput = e => renderClientList(e.target.value);
  }

  // Logout
  const logoutBtn = $('btn-logout');
  if (logoutBtn) {
    logoutBtn.onclick = async () => {
      APP.alertShown = false;
      stopLocationTracking();    // Oprește GPS tracking
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
        SYNC_CONFIG.API_URL = url.value.trim();
        await setSetting('api_url', SYNC_CONFIG.API_URL);
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
      // GPS interval
      const gpsIntervalEl = $('settings-gps-interval');
      if (gpsIntervalEl) {
        await setSetting('gps_interval', gpsIntervalEl.value);
        // Repornește tracking-ul cu noul interval (dacă era activ)
        if (_locationInterval) { stopLocationTracking(); startLocationTracking(); }
      }
      // GPS program (ore start / end)
      const gpsStartEl = $('settings-gps-start');
      const gpsEndEl   = $('settings-gps-end');
      if (gpsStartEl && gpsEndEl) {
        const newStart = parseInt(gpsStartEl.value, 10);
        const newEnd   = parseInt(gpsEndEl.value,   10);
        await setSetting('gps_schedule_start', String(newStart));
        await setSetting('gps_schedule_end',   String(newEnd));
        APP.gpsStart = newStart;
        APP.gpsEnd   = newEnd;
        updateGpsToggleBtn(); // reflectă noile ore imediat
      }
      showToast('Setări salvate.', 'success');
    };
  }

  // Load settings into UI
  getSetting('api_url').then(url => {
    const urlInput = $('settings-api-url');
    if (urlInput && url) urlInput.value = url;
  });
  getSetting('alert_threshold').then(thr => {
    const thrInput = $('settings-alert-threshold');
    if (thrInput) thrInput.value = thr || APP.alertThreshold;
  });
  getSetting('gps_interval').then(val => {
    const el = $('settings-gps-interval');
    if (el && val) el.value = val;
  });
  // Populare selects ore GPS (0–23) și setare valori salvate
  ['settings-gps-start', 'settings-gps-end'].forEach(id => {
    const el = $(id);
    if (!el) return;
    el.innerHTML = Array.from({ length: 24 }, (_, h) =>
      `<option value="${h}">${String(h).padStart(2,'0')}:00</option>`
    ).join('');
  });
  getSetting('gps_schedule_start').then(val => {
    const el = $('settings-gps-start');
    if (el) el.value = val ?? '7';
  });
  getSetting('gps_schedule_end').then(val => {
    const el = $('settings-gps-end');
    if (el) el.value = val ?? '18';
  });
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

  // Sort by urgency: overdue → never → soon → ok
  const urgencyOrder = { overdue: 0, never: 1, soon: 2, ok: 3 };
  filtered.sort((a, b) => (urgencyOrder[a.urgency] || 3) - (urgencyOrder[b.urgency] || 3));

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

  // Toast warning — o singură dată per sesiune
  if (!APP.alertShown) {
    const overLimit = filtered.filter((f, i) => counts[i] >= thr);
    if (overLimit.length) {
      const names = overLimit.map(f => f.client.name).join(', ');
      showToast(`⚠ ${overLimit.length} client(i) cu ${thr}+ intervenții neraportate: ${names}`, 'warning', 8000);
      APP.alertShown = true;
    }
  }

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
        ${admin ? `<button class="client-action-btn" onclick="showClientDetails('${client.client_id}')">📋 Istoric</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="showEditClientModal('${client.client_id}')">✏️ Editează</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="showQRCode('${client.client_id}')">📱 QR</button>` : ''}
        ${admin ? `<button class="client-action-btn" onclick="showExportModal('${client.client_id}')">📥 Export</button>` : ''}
      </div>
    </li>`;
  }).join('');
}

function getLastVisitInfo(clientId) {
  const ci = APP.interventions.filter(i => i.client_id === clientId);
  if (!ci.length) return { badge: '<span class="last-visit-badge none">Nicio vizită</span>', days: null };

  const latest = ci.sort((a, b) => b.date.localeCompare(a.date))[0];
  const days = Math.floor((Date.now() - Date.parse(latest.date)) / 86400000);
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

// ── Intervention Screen ───────────────────────────────────────
async function renderIntervention(client) {
  // Header
  const nameEl = $('intervention-client-name');
  const dateEl = $('intervention-date');
  if (nameEl) nameEl.textContent = client.name;
  if (dateEl) dateEl.textContent = new Date().toLocaleDateString('ro-RO', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  // Track arrival time
  APP.arrivalTime    = new Date().toISOString();
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
  if (backBtn) backBtn.onclick = () => showScreen('dashboard');

  // Save button
  const saveBtn = $('btn-save');
  if (saveBtn) saveBtn.onclick = showConfirmModal;

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
  renderTreatmentSteppers().catch(e => console.warn('[STEPPER] Error:', e));
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

function resetInterventionForm() {
  // Measured values
  ['m-chlorine','m-ph','m-temp','m-hardness','m-alkalinity','m-salinity'].forEach(id => {
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
    updateSaveButton();
    return;
  }

  const rec = getRecommendation(vol, cl, ph);
  if (!rec) {
    ['rec-cl-granule','rec-cl-tab','rec-ph-kg','rec-anti'].forEach(id => {
      const el = $(id); if (el) el.textContent = 'N/A';
    });
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
  const durationMin    = APP.arrivalTime
    ? Math.round((Date.parse(departureTime) - Date.parse(APP.arrivalTime)) / 60000)
    : null;

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
    date:             new Date().toISOString().split('T')[0],
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
    photos:       [...APP.currentPhotos],
    synced:       false,

    // GPS
    geo_lat:      APP.currentPosition ? APP.currentPosition.lat : null,
    geo_lng:      APP.currentPosition ? APP.currentPosition.lng : null,
    geo_accuracy: APP.currentPosition ? APP.currentPosition.accuracy : null,

    // Time tracking
    arrival_time:     APP.arrivalTime,
    departure_time:   departureTime,
    duration_minutes: durationMin
  };

  // Dynamic treatment fields from stock products
  const products = APP._stockProducts.length ? APP._stockProducts : await getAllStock();
  products.forEach(p => {
    const el = $('t-' + p.product_id);
    intervention['treat_' + p.product_id] = el ? (parseFloat(el.value) || 0) : 0;
  });

  try {
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

  const ci = APP.interventions
    .filter(i => i.client_id === client.client_id)
    .sort((a, b) => b.date.localeCompare(a.date))
    .slice(0, 5);

  if (!ci.length) {
    container.innerHTML = '<p style="padding:12px;color:var(--slate-400);font-size:.85rem">Nicio intervenție anterioară.</p>';
    return;
  }

  container.innerHTML = ci.map(i => {
    const dur = i.duration_minutes != null ? `<span class="prev-int-duration">⏱ ${i.duration_minutes} min</span>` : '';
    return `<div class="prev-intervention">
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
function showClientDetails(clientId) {
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;

  const modal = $('modal-client');
  const body  = $('modal-client-body');
  if (!modal || !body) return;

  const hasLocation = client.location_set && client.latitude;
  const ci = APP.interventions.filter(i => i.client_id === clientId)
               .sort((a, b) => b.date.localeCompare(a.date));

  body.innerHTML = `
    <div class="client-detail-section">
      <h4>Informații</h4>
      <div class="client-detail-row"><span class="detail-label">Volum piscină</span><span class="detail-value">${client.pool_volume_mc} m³</span></div>
      <div class="client-detail-row"><span class="detail-label">Tip</span><span class="detail-value">${client.pool_type}</span></div>
      ${client.phone ? `<div class="client-detail-row"><span class="detail-label">Telefon</span><span class="detail-value">${escHtml(client.phone)}</span></div>` : ''}
      ${client.address ? `<div class="client-detail-row"><span class="detail-label">Adresă</span><span class="detail-value">${escHtml(client.address)}</span></div>` : ''}
      <div class="client-detail-row"><span class="detail-label">Locație GPS</span><span class="detail-value">${hasLocation ? '✅ Setată' : '❌ Nesetată'}</span></div>
      ${client.notes ? `<div class="client-detail-row"><span class="detail-label">Note</span><span class="detail-value">${escHtml(client.notes)}</span></div>` : ''}
    </div>
    <div class="client-detail-section">
      <h4>Ultimele ${Math.min(ci.length, 5)} intervenții</h4>
      ${ci.slice(0, 5).map(i => `
        <div class="prev-intervention">
          <div class="prev-int-header">
            <span class="prev-int-date">${fmtDate(i.date)}</span>
            ${i.duration_minutes != null ? `<span class="prev-int-duration">⏱ ${i.duration_minutes} min</span>` : ''}
          </div>
          <div class="prev-int-tech">👤 ${escHtml(i.technician_name || '')}</div>
          <div class="prev-int-measures">
            <span class="prev-measure">Cl: <strong>${i.measured_chlorine ?? '—'}</strong></span>
            <span class="prev-measure">pH: <strong>${i.measured_ph ?? '—'}</strong></span>
            <span class="prev-measure">Cl.gr: <strong>${i.treat_cl_granule_gr || 0}gr</strong></span>
            <span class="prev-measure">pH.gr: <strong>${i.treat_ph_granule || 0}kg</strong></span>
          </div>
        </div>
      `).join('')}
      ${ci.length === 0 ? '<p style="color:var(--slate-400);font-size:.85rem">Nicio intervenție.</p>' : ''}
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
}

function closeClientModal() {
  const modal = $('modal-client');
  if (modal) modal.classList.remove('open');
}

// ── Export Modal ──────────────────────────────────────────────
function showExportModal(clientId) {
  const modal = $('modal-export');
  if (!modal) return;

  const client = clientId ? APP.clients.find(c => c.client_id === clientId) : null;

  const exportClientBtn = $('export-client-btn');
  const exportAllBtn    = $('export-all-btn');

  if (exportClientBtn) {
    if (client) {
      exportClientBtn.style.display = '';
      exportClientBtn.querySelector('.export-option-text h4').textContent = 'Export ' + client.name;
      exportClientBtn.onclick = async () => {
        modal.classList.remove('open');
        showToast('Generare Excel...', 'info');
        try {
          const ci = APP.interventions.filter(i => i.client_id === client.client_id);
          await exportClientXLSX(client, ci);
          showToast('Export complet!', 'success');
        } catch(e) { showToast('Eroare export: ' + e.message, 'error'); }
      };
    } else {
      exportClientBtn.style.display = 'none';
    }
  }

  if (exportAllBtn) {
    exportAllBtn.onclick = async () => {
      modal.classList.remove('open');
      showToast('Generare Excel complet...', 'info');
      try {
        await exportAllXLSX(APP.clients, APP.interventions);
        showToast('Export complet!', 'success');
      } catch(e) { showToast('Eroare export: ' + e.message, 'error'); }
    };
  }

  const exportStructuredBtn = $('export-structured-btn');
  if (exportStructuredBtn) {
    exportStructuredBtn.onclick = async () => {
      modal.classList.remove('open');
      showToast('Generare Excel structurat...', 'info');
      try {
        await exportStructuredXLSX(APP.clients, APP.interventions);
        showToast('Export structurat complet!', 'success');
      } catch(e) { showToast('Eroare export: ' + e.message, 'error'); }
    };
  }

  modal.classList.add('open');
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
  const d = new Date(isoDate + 'T12:00:00');
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
  const freq = $('cf-visit-freq'); if (freq) freq.value = '14';
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
    visit_frequency_days: parseInt($('cf-visit-freq') ? $('cf-visit-freq').value : '14') || 14,
    billing_interval_interventions: billingRaw > 0 ? billingRaw : null,
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

// ════════════════════════════════════════════════════════════════
// FEATURE 7 — Technician Manager
// ════════════════════════════════════════════════════════════════
async function showTechManager() {
  const modal = $('modal-tech-manager');
  const body  = $('tech-manager-body');
  if (!modal || !body) return;

  hideTechForm();
  const techs = await getAll('technicians');

  body.innerHTML = techs.length ? techs.map(t => `
    <div class="tech-row">
      <div>
        <strong>${escHtml(t.name)}</strong>
        <span class="tech-role-badge ${t.role === 'admin' ? 'badge-admin' : 'badge-tech'}">${t.role}</span>
        <div style="font-size:.78rem;color:var(--slate-500)">@${escHtml(t.username)}</div>
      </div>
      <div style="display:flex;gap:6px;align-items:center">
        <span style="font-size:.8rem;color:${t.active !== false ? 'var(--success)' : 'var(--danger)'}">${t.active !== false ? '● Activ' : '● Inactiv'}</span>
        <button class="client-action-btn" onclick="toggleTechActive('${t.technician_id}')">${t.active !== false ? 'Dezactivează' : 'Activează'}</button>
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
  const username = $('tf-username') ? $('tf-username').value.trim().toLowerCase() : '';
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
    showToast(existingId ? 'Tehnician actualizat.' : 'Tehnician adăugat.', 'success');
    showTechManager();
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
  }
}

async function toggleTechActive(techId) {
  try {
    const tech = await getByKey('technicians', techId);
    if (!tech) return;
    tech.active = tech.active === false ? true : false;
    await put('technicians', tech);
    showTechManager();
  } catch (e) {
    showToast('Eroare: ' + e.message, 'error');
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

  if (nameEl) nameEl.textContent = client.name;
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
    ? `\n⏱ *Durată intervenție:* ${intervention.duration_minutes} min` : '';

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
  await remove('stock', productId);
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

/** Switch tab on page 2 (Tratament / Note & Foto) */
function switchP2Tab(tab) {
  ['treatment', 'notes'].forEach(t => {
    const btn   = $('tab-' + t);
    const panel = $('panel-' + t);
    if (btn)   btn.classList.toggle('active',   t === tab);
    if (panel) panel.classList.toggle('active', t === tab);
  });
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

// ════════════════════════════════════════════════════════════════
// FEATURE C — Notificare Facturare per Client
// ════════════════════════════════════════════════════════════════

/** Check if billing threshold is reached and show toast */
function checkBillingAlert(client) {
  const interval = client.billing_interval_interventions;
  if (!interval || interval <= 0) return;

  const since = client.last_billing_date || '1970-01-01';
  const countSince = APP.interventions.filter(i =>
    i.client_id === client.client_id && i.date > since
  ).length; // includes the one just added (already in APP.interventions)

  if (countSince >= interval) {
    showToast(`💰 ${client.name} — ${countSince} intervenții de la ultima facturare!`, 'warning', 8000);
  }
}

/** Mark a client as billed (reset last_billing_date to today) */
async function markClientBilled() {
  const clientId = APP._billingClientId;
  if (!clientId) return;
  const client = APP.clients.find(c => c.client_id === clientId);
  if (!client) return;

  client.last_billing_date = new Date().toISOString().split('T')[0];
  client.updated_at = new Date().toISOString();
  await put('clients', client);
  APP.clients = APP.clients.map(c => c.client_id === clientId ? client : c);

  showToast(`✓ ${client.name} marcat ca facturat.`, 'success');

  // Update billing badge visibility
  const billBtn = $('btn-mark-billed');
  if (billBtn) billBtn.style.display = 'none';
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

// ════════════════════════════════════════════════════════════════
// GPS LOCATION TRACKING — trimite poziția la GAS la fiecare 5 min
// ════════════════════════════════════════════════════════════════

let _locationInterval       = null;
let _bgLocationWatcherId    = null; // Capacitor BackgroundGeolocation watcher ID

/**
 * Pornește tracking GPS.
 * — În APK Capacitor: folosește BackgroundGeolocation (funcționează în background)
 * — În browser/PWA:   folosește navigator.geolocation la fiecare 5 min (foreground)
 */
async function startLocationTracking() {
  if (!isSyncConfigured()) return;
  stopLocationTracking();

  // ── Capacitor native mode (APK) ──────────────────────────────
  if (window.Capacitor?.isNativePlatform?.()) {
    const BGL = window.Capacitor.Plugins?.BackgroundGeolocation;
    if (BGL) {
      try {
        _bgLocationWatcherId = await BGL.addWatcher(
          { backgroundMessage: 'Pool Manager urmărește locația ta',
            backgroundTitle:   'Pool Manager GPS Activ',
            requestPermissions: true,
            stale: false,
            distanceFilter: 100 },
          (location, error) => {
            if (error || !location) return;
            _sendLocationData(location.latitude, location.longitude, location.accuracy);
          }
        );
        return; // native tracker activ, nu mai folosim fallback
      } catch { /* plugin indisponibil — cade la fallback */ }
    }
  }

  // ── Browser / PWA fallback: trimite la intervalul configurat (foreground only) ──
  if (!navigator.geolocation) return;
  const gpsIntervalSec = parseInt(await getSetting('gps_interval') || '300', 10);
  sendCurrentLocation();
  _locationInterval = setInterval(sendCurrentLocation, gpsIntervalSec * 1000);
}

/** Oprește tracking-ul GPS (ambele moduri). */
function stopLocationTracking() {
  if (_locationInterval) { clearInterval(_locationInterval); _locationInterval = null; }
  if (_bgLocationWatcherId) {
    const BGL = window.Capacitor?.Plugins?.BackgroundGeolocation;
    if (BGL) BGL.removeWatcher({ id: _bgLocationWatcherId }).catch(() => {});
    _bgLocationWatcherId = null;
  }
}

// ── GPS Schedule & Manual Override ───────────────────────────────

/**
 * Verifică dacă ora curentă este în intervalul de program configurat (ex. 07:00–18:00).
 * Folosește APP.gpsStart și APP.gpsEnd (încărcate din setări la login).
 */
function isWithinGpsHours() {
  const h = new Date().getHours();
  return h >= APP.gpsStart && h < APP.gpsEnd;
}

/**
 * Returnează true dacă GPS-ul trebuie să trimită date acum:
 * — override='on'  → mereu activ
 * — override='off' → mereu inactiv
 * — null           → urmează programul (isWithinGpsHours)
 */
async function shouldSendGps() {
  const ov = await getSetting('gps_manual_override');
  if (ov === 'on')  return true;
  if (ov === 'off') return false;
  return isWithinGpsHours();
}

/**
 * Buton GPS din footer: comutare inteligentă On/Off.
 *   — GPS activ + în program  → oprire manuală (override='off')
 *   — GPS oprit + în program  → reactivare (șterge override)
 *   — GPS activ + afara prog. → oprire (șterge override='on')
 *   — GPS oprit + afara prog. → pornire manuală (override='on')
 */
async function toggleGpsOverride() {
  const active   = await shouldSendGps();
  const inHours  = isWithinGpsHours();
  const startStr = APP.gpsStart + ':00';
  const endStr   = APP.gpsEnd   + ':00';

  if (active) {
    // GPS trimite acum → utilizatorul vrea să îl oprească
    await setSetting('gps_manual_override', 'off');
    showToast('GPS dezactivat manual.', 'warn', 3500);
  } else if (inHours) {
    // În program, dar era oprit manual → reactivare la normal
    await setSetting('gps_manual_override', null);
    showToast('GPS reactivat (program normal).', 'success', 3000);
  } else {
    // În afara programului → pornire forțată
    await setSetting('gps_manual_override', 'on');
    showToast(`GPS pornit manual (în afara programului ${startStr}–${endStr}).`, 'success', 4000);
  }
  updateGpsToggleBtn();
}

/**
 * Actualizează aspectul butonului GPS din footer în funcție de starea curentă.
 * 🟢 GPS Activ    — program activ, fără override
 * 🟡 GPS Manual   — override='on' în afara programului
 * 🔴 GPS Oprit    — override='off' (oprit manual în program)
 * ⭕ GPS Inactiv  — în afara programului, fără override
 */
async function updateGpsToggleBtn() {
  const btn = $('btn-gps-status');
  if (!btn) return;
  const ov     = await getSetting('gps_manual_override');
  const active = ov === 'on' || (ov !== 'off' && isWithinGpsHours());

  if (ov === 'off') {
    btn.textContent = '🔴 GPS';
    btn.title = 'GPS oprit manual — apasă pentru reactivare';
    btn.dataset.state = 'off';
  } else if (ov === 'on') {
    btn.textContent = '🟡 GPS';
    btn.title = `GPS pornit manual (în afara programului ${APP.gpsStart}:00–${APP.gpsEnd}:00)`;
    btn.dataset.state = 'manual';
  } else if (active) {
    btn.textContent = '🟢 GPS';
    btn.title = `GPS activ conform programului (${APP.gpsStart}:00–${APP.gpsEnd}:00)`;
    btn.dataset.state = 'on';
  } else {
    btn.textContent = '⭕ GPS';
    btn.title = `GPS inactiv (în afara programului ${APP.gpsStart}:00–${APP.gpsEnd}:00) — apasă pentru pornire manuală`;
    btn.dataset.state = 'idle';
  }
}

/**
 * Trimite datele GPS la GAS — apelat din ambele căi (nativă și browser).
 */
async function _sendLocationData(lat, lng, accuracy) {
  if (!APP.user || !isSyncConfigured()) return;
  // Verificare program + override manual
  if (!await shouldSendGps()) return;
  try {
    await fetch(SYNC_CONFIG.API_URL, {
      method: 'POST', mode: 'no-cors',
      body: JSON.stringify({
        action: 'saveLocation',
        technician_id: APP.user.technician_id,
        name:          APP.user.name,
        lat, lng,
        accuracy:  Math.round(accuracy || 0),
        timestamp: new Date().toISOString()
      })
    });
  } catch { /* offline — silently skip */ }
}

/** Browser fallback — obține poziția și apelează _sendLocationData. */
function sendCurrentLocation() {
  if (!APP.user || !isSyncConfigured()) return;
  navigator.geolocation.getCurrentPosition(
    pos => _sendLocationData(pos.coords.latitude, pos.coords.longitude, pos.coords.accuracy),
    () => { /* GPS indisponibil / permisiune refuzată — ignorăm silențios */ },
    { timeout: 10000, enableHighAccuracy: false, maximumAge: 120000 }
  );
}

// ════════════════════════════════════════════════════════════════
// MAP SCREEN — Leaflet.js + OpenStreetMap (complet gratuit)
// ════════════════════════════════════════════════════════════════

let _leafletLoaded    = false;
let _leafletMap       = null;
let _historyLayer     = null;   // Leaflet LayerGroup pentru traseul istoric
let _mapMarkers       = [];
let _mapRefreshInterval = null;

/**
 * Inițializează / reîncarcă ecranul hartă.
 * Apelat din showScreen('map').
 */
async function loadMapScreen() {
  const noApi    = $('map-no-api');
  const mapDiv   = $('map-container');

  if (!isSyncConfigured()) {
    if (noApi)  noApi.style.display  = '';
    if (mapDiv) mapDiv.style.display = 'none';
    return;
  }
  if (noApi)  noApi.style.display  = 'none';
  if (mapDiv) mapDiv.style.display = '';

  // Leaflet se încarcă o singură dată (lazy CDN)
  if (!_leafletLoaded) {
    await loadLeaflet();
    _leafletLoaded = true;
  }

  // Inițializează harta o singură dată (centru România)
  if (!_leafletMap && mapDiv) {
    _leafletMap = L.map('map-container').setView([45.9432, 24.9668], 7);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© <a href="https://openstreetmap.org">OpenStreetMap</a>',
      maxZoom: 19
    }).addTo(_leafletMap);
  }

  // Recalculează dimensiunea (ecranul era ascuns la init)
  setTimeout(() => { if (_leafletMap) _leafletMap.invalidateSize(); }, 150);

  await refreshMapLocations();
  _mapRefreshInterval = setInterval(refreshMapLocations, 60 * 1000); // refresh la 60s

  // ── Populare dropdown tehnicieni pentru istoricul GPS (admin only) ──
  if (isAdmin()) {
    const histTech = $('map-hist-tech');
    if (histTech) {
      getAll('technicians').then(techs => {
        const options = techs
          .filter(t => t.role === 'technician' || t.role !== 'admin')
          .map(t => `<option value="${escHtml(t.technician_id)}">${escHtml(t.name)}</option>`)
          .join('');
        histTech.innerHTML = '<option value="">Alege tehnician...</option>' + options;
      });
    }
    const dateEl = $('map-hist-date');
    if (dateEl && !dateEl.value) {
      dateEl.value = new Date().toISOString().slice(0, 10); // azi implicit
    }
  }
}

/** Obține locațiile de la GAS și actualizează markerii pe hartă. */
async function refreshMapLocations() {
  if (!isSyncConfigured() || !_leafletMap) return;
  try {
    const resp      = await fetch(SYNC_CONFIG.API_URL + '?action=getLocations', { cache: 'no-store' });
    const locations = await resp.json();
    if (Array.isArray(locations)) renderMapMarkers(locations);
    const upd = $('map-last-update');
    if (upd) upd.textContent = 'actualizat ' + new Date().toLocaleTimeString('ro-RO', { hour: '2-digit', minute: '2-digit' });
  } catch { /* offline */ }
}

/** Desenează markerii pe hartă. Culorile indică vârsta poziției. */
function renderMapMarkers(locations) {
  if (!_leafletMap) return;
  _mapMarkers.forEach(m => _leafletMap.removeLayer(m));
  _mapMarkers = [];
  if (!locations.length) return;

  locations.forEach(loc => {
    const ageMs   = Date.now() - new Date(loc.timestamp).getTime();
    const minsAgo = Math.round(ageMs / 60000);
    // Verde = sub 10 min, Portocaliu = sub 1h, Gri = mai vechi
    const color   = ageMs < 10 * 60000 ? '#16a34a' : ageMs < 60 * 60000 ? '#d97706' : '#94a3b8';
    const initials = loc.name.split(' ').map(w => w[0] || '').join('').slice(0, 2).toUpperCase();
    const timeLabel = minsAgo < 1 ? 'acum' : minsAgo < 60 ? minsAgo + 'min' : Math.round(minsAgo / 60) + 'h';

    const icon = L.divIcon({
      className: '',
      html: `<div class="map-marker" style="background:${color}">` +
            `<span class="map-marker-init">${initials}</span>` +
            `<span class="map-marker-time">${timeLabel}</span></div>`,
      iconSize:   [44, 52],
      iconAnchor: [22, 52]
    });

    const popup =
      `<div class="map-popup-inner">` +
      `<b>${escHtml(loc.name)}</b><br>` +
      `<span style="color:#64748b;font-size:12px">` +
        (minsAgo < 1 ? 'Locație curentă' : `acum ${minsAgo} min`) +
      `</span><br>` +
      `<span style="color:#94a3b8;font-size:11px">±${loc.accuracy} m</span>` +
      `</div>`;

    const marker = L.marker([loc.lat, loc.lng], { icon }).addTo(_leafletMap).bindPopup(popup, { maxWidth: 220 });
    _mapMarkers.push(marker);
  });

  // Zoom automat să cuprindă toți markerii vizibili
  if (_mapMarkers.length > 0) {
    const bounds = L.latLngBounds(locations.map(l => [l.lat, l.lng]));
    _leafletMap.fitBounds(bounds, { padding: [50, 50], maxZoom: 14 });
  }
}

/** Încarcă Leaflet CSS + JS din CDN, o singură dată. */
function loadLeaflet() {
  return new Promise(resolve => {
    if (window.L) { resolve(); return; }
    const link  = document.createElement('link');
    link.rel    = 'stylesheet';
    link.href   = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
    document.head.appendChild(link);
    const script   = document.createElement('script');
    script.src     = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
    script.onload  = resolve;
    script.onerror = resolve; // nu bloca app-ul dacă CDN offline
    document.head.appendChild(script);
  });
}

// ════════════════════════════════════════════════════════════════
// GPS HISTORY — Traseu istoric pe hartă
// ════════════════════════════════════════════════════════════════

/**
 * Preia traseul unui tehnician pentru o dată selectată și îl desenează pe hartă.
 */
async function loadAndShowHistory() {
  const techId  = $('map-hist-tech')?.value;
  const dateVal = $('map-hist-date')?.value;

  if (!techId)  { showToast('Alege un tehnician.', 'warn', 3000); return; }
  if (!dateVal) { showToast('Alege o dată.', 'warn', 3000); return; }
  if (!isSyncConfigured()) { showToast('Configurați API URL în Setări.', 'warn'); return; }
  if (!_leafletMap) { showToast('Harta nu este inițializată.', 'warn'); return; }

  showToast('Se încarcă traseul...', 'info', 2000);

  try {
    const url  = SYNC_CONFIG.API_URL +
      '?action=getLocationHistory&tech_id=' + encodeURIComponent(techId) +
      '&date=' + encodeURIComponent(dateVal);
    const resp = await fetch(url, { cache: 'no-store' });
    const data = await resp.json();

    if (!data.positions || data.positions.length === 0) {
      showToast('Nicio poziție găsită pentru această dată.', 'warn', 4000);
      return;
    }
    renderHistoryTrail(data.positions, data.tech_name || 'Tehnician');
  } catch {
    showToast('Eroare la încărcarea istoricului GPS.', 'error');
  }
}

/**
 * Desenează traseul pe hartă: polyline albastru + markeri start/stop/intermediari.
 */
function renderHistoryTrail(positions, techName) {
  if (!_leafletMap || !window.L) return;
  clearHistoryTrail();

  const latlngs = positions.map(p => [p.lat, p.lng]);

  _historyLayer = L.layerGroup();

  // Linie traseu
  L.polyline(latlngs, { color: '#3b82f6', weight: 3, opacity: 0.75, dashArray: null })
    .addTo(_historyLayer);

  // Marker START — verde
  L.circleMarker(latlngs[0], {
    radius: 9, color: '#fff', weight: 2,
    fillColor: '#16a34a', fillOpacity: 1
  }).bindPopup(`<b>🟢 Start</b><br>${positions[0].time}`).addTo(_historyLayer);

  // Marker STOP — roșu
  const last = positions[positions.length - 1];
  L.circleMarker([last.lat, last.lng], {
    radius: 9, color: '#fff', weight: 2,
    fillColor: '#dc2626', fillOpacity: 1
  }).bindPopup(`<b>🔴 Sfârșit</b><br>${last.time}`).addTo(_historyLayer);

  // Markeri intermediari — mici, albaștri
  positions.slice(1, -1).forEach(p => {
    L.circleMarker([p.lat, p.lng], {
      radius: 4, color: '#3b82f6', weight: 1,
      fillColor: '#93c5fd', fillOpacity: 0.85
    }).bindPopup(`${p.time}<br><span style="font-size:11px;color:#64748b">±${Math.round(p.accuracy)} m</span>`)
      .addTo(_historyLayer);
  });

  _historyLayer.addTo(_leafletMap);

  // Zoom să cuprindă tot traseul
  try { _leafletMap.fitBounds(L.latLngBounds(latlngs), { padding: [40, 40], maxZoom: 16 }); }
  catch { /* bounds invalide dacă un singur punct */ }

  // Info bar sub controale
  const info = $('map-hist-info');
  if (info) {
    const dur = positions.length > 1
      ? _formatDuration(positions[0].timestamp, last.timestamp)
      : '';
    info.style.display = '';
    info.innerHTML =
      `<b>${escHtml(techName)}</b> — ${positions.length} poziții` +
      (dur ? ` &nbsp;·&nbsp; ${dur}` : '') +
      ` &nbsp;·&nbsp; ${positions[0].time} → ${last.time}`;
  }
}

/** Elimină traseul de pe hartă și resetează info bar. */
function clearHistoryTrail() {
  if (_historyLayer && _leafletMap) {
    _leafletMap.removeLayer(_historyLayer);
    _historyLayer = null;
  }
  const info = $('map-hist-info');
  if (info) info.style.display = 'none';
}

/** Formatează durata dintre două timestamps ISO ca "Xh Ymin". */
function _formatDuration(tsStart, tsEnd) {
  const diffMin = Math.round((new Date(tsEnd) - new Date(tsStart)) / 60000);
  if (diffMin < 1)  return '< 1 min';
  if (diffMin < 60) return diffMin + ' min';
  const h = Math.floor(diffMin / 60);
  const m = diffMin % 60;
  return h + 'h' + (m > 0 ? ' ' + m + 'min' : '');
}

// ─────────────────────────────────────────────────────────────
// ── CALENDAR INTERVENȚII ──────────────────────────────────────
// ─────────────────────────────────────────────────────────────

let _calWeekOffset = 0; // 0 = săptămâna curentă, -1/+1 = prev/next

/** Returnează {start, end, label} pentru săptămâna Luni–Duminică cu offset-ul dat. */
function getWeekBounds(offset) {
  const now  = new Date();
  const day  = now.getDay();                 // 0=Dum, 1=Lun…
  const diff = (day === 0 ? -6 : 1 - day);  // zile până la Luni acestei săptămâni
  const mon  = new Date(now);
  mon.setHours(0, 0, 0, 0);
  mon.setDate(now.getDate() + diff + offset * 7);
  const sun = new Date(mon);
  sun.setDate(mon.getDate() + 6);

  const toISO = d => d.toISOString().slice(0, 10);
  const mo    = ['ian','feb','mar','apr','mai','iun','iul','aug','sep','oct','nov','dec'];
  const monLbl = mon.getDate() + ' ' + mo[mon.getMonth()];
  const sunLbl = sun.getDate() + ' ' + mo[sun.getMonth()] + ' ' + sun.getFullYear();
  return { start: toISO(mon), end: toISO(sun), label: monLbl + ' – ' + sunLbl };
}

/** Încarcă și randează calendarul pentru săptămâna curentă (cu offset). */
async function loadCalendarScreen() {
  const noApi   = $('cal-no-api');
  const loading = $('cal-loading');
  const content = $('cal-content');

  // Dacă API nu e configurat — mesaj explicativ
  if (!isSyncConfigured()) {
    if (noApi)   noApi.style.display   = '';
    if (loading) loading.style.display = 'none';
    if (content) content.innerHTML     = '';
    return;
  }
  if (noApi) noApi.style.display = 'none';

  const bounds = getWeekBounds(_calWeekOffset);
  const label  = $('cal-week-label');
  if (label) label.textContent = bounds.label;

  if (loading) loading.style.display = '';
  if (content) content.innerHTML     = '';

  try {
    let url = SYNC_CONFIG.API_URL + '?action=getCalendar&date_from=' + bounds.start + '&date_to=' + bounds.end;
    // Tehnicianul vede doar propriile intervenții
    if (APP.user && APP.user.role !== 'admin') {
      url += '&tech_id=' + encodeURIComponent(APP.user.technician_id);
    }
    const resp = await fetch(url, { cache: 'no-store' });
    const data = await resp.json();
    if (loading) loading.style.display = 'none';
    renderCalendar(data.entries || [], bounds);
  } catch (err) {
    if (loading) loading.style.display = 'none';
    if (content) content.innerHTML =
      '<p style="text-align:center;padding:40px 16px;color:var(--slate-400)">⚠️ Eroare la încărcarea programului.<br><small>' + escHtml(err.message) + '</small></p>';
  }
}

/** Navighează cu o săptămână înainte sau înapoi. */
function changeCalendarWeek(delta) {
  _calWeekOffset += delta;
  loadCalendarScreen();
}

/** Revine la săptămâna curentă. */
function jumpCalendarToToday() {
  _calWeekOffset = 0;
  loadCalendarScreen();
}

/** Randează lista de intervenții zi cu zi în #cal-content. */
function renderCalendar(entries, bounds) {
  const content = $('cal-content');
  if (!content) return;

  const today    = new Date().toISOString().slice(0, 10);
  const dayNames = ['Duminică','Luni','Marți','Miercuri','Joi','Vineri','Sâmbătă'];
  const mo       = ['ianuarie','februarie','martie','aprilie','mai','iunie',
                    'iulie','august','septembrie','octombrie','noiembrie','decembrie'];
  const isAdmin  = APP.user && APP.user.role === 'admin';

  // Mapare technician_id → indice culoare (0–4, ciclic)
  const techColors = {};
  let   colorIdx   = 0;
  entries.forEach(e => {
    if (!techColors.hasOwnProperty(e.technician_id)) {
      techColors[e.technician_id] = colorIdx++ % 5;
    }
  });

  // Grupare după dată
  const byDate = {};
  entries.forEach(e => {
    if (!byDate[e.date]) byDate[e.date] = [];
    byDate[e.date].push(e);
  });

  let html = '';
  const d = new Date(bounds.start + 'T00:00:00');

  for (let i = 0; i < 7; i++) {
    const dateStr  = d.toISOString().slice(0, 10);
    const isToday  = dateStr === today;
    const dayLabel = dayNames[d.getDay()] + ', ' + d.getDate() + ' ' + mo[d.getMonth()];

    html += `<div class="cal-day-group">
      <div class="cal-day-header${isToday ? ' is-today' : ''}">
        ${escHtml(dayLabel)}${isToday ? ' <span class="cal-today-badge">Azi</span>' : ''}
      </div>`;

    const dayEntries = byDate[dateStr] || [];
    if (dayEntries.length === 0) {
      html += `<div class="cal-day-empty">Nicio intervenție planificată</div>`;
    } else {
      dayEntries.forEach(e => {
        const tci = techColors[e.technician_id] !== undefined ? techColors[e.technician_id] : 0;
        html += `<div class="cal-entry" data-tech-color="${tci}">
          <div class="cal-entry-time">${escHtml(e.time || '—')}</div>
          <div class="cal-entry-body">
            ${isAdmin ? `<div class="cal-entry-tech">👤 ${escHtml(e.technician_name || '')}</div>` : ''}
            <div class="cal-entry-client">${escHtml(e.client_name || '—')}</div>
            ${e.address ? `<div class="cal-entry-addr">📍 ${escHtml(e.address)}</div>` : ''}
            ${e.notes   ? `<div class="cal-entry-notes">${escHtml(e.notes)}</div>` : ''}
          </div>
          ${isAdmin ? `<button class="cal-entry-delete" onclick="deleteCalendarEntry('${e.id.replace(/'/g,"\\'")}')">✕</button>` : ''}
        </div>`;
      });
    }
    html += '</div>';
    d.setDate(d.getDate() + 1);
  }

  // Dacă nu există deloc intervenții pentru săptămână
  if (!entries.length) {
    html += `<div style="text-align:center;padding:32px 16px;color:var(--slate-400);font-size:.9rem">
      📭 Nu există intervenții programate în această săptămână.
      ${isAdmin ? '<br><small style="font-size:.8rem">Importați un fișier Excel cu butonul 📥 din header.</small>' : ''}
    </div>`;
  }

  content.innerHTML = html;
}

/** Șterge o intrare din calendar (admin only). */
async function deleteCalendarEntry(id) {
  if (!id) return;
  if (!confirm('Ștergi această intervenție din program?')) return;
  try {
    const resp = await fetch(SYNC_CONFIG.API_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ action: 'deleteCalendarEntry', id })
    });
    const data = await resp.json();
    if (data.success) {
      showToast('Intervenție ștearsă din program.', 'success');
      loadCalendarScreen();
    } else {
      showToast('Eroare la ștergere: ' + (data.error || ''), 'error');
    }
  } catch (err) {
    showToast('Eroare de rețea.', 'error');
  }
}

/** Descarcă template Excel pentru import calendar. */
function downloadCalendarTemplate() {
  if (typeof XLSX === 'undefined') {
    showToast('SheetJS nu este disponibil. Reconectați-vă la internet.', 'warning');
    return;
  }
  const headers = ['data (YYYY-MM-DD)', 'ora (HH:MM)', 'technician_name', 'client_name', 'adresa', 'observatii'];
  const ex1     = [new Date().toISOString().slice(0,10), '09:00', 'Nume Tehnician', 'Client ABC', 'Str. Exemplu 1, București', 'Verificare clor'];
  const ex2     = [new Date().toISOString().slice(0,10), '11:30', 'Nume Tehnician', 'Client DEF', 'Str. Exemplu 2, Cluj', ''];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ex1, ex2]);
  // Stil header (lărgime coloane)
  ws['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 20 }, { wch: 24 }, { wch: 36 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Program');
  XLSX.writeFile(wb, 'template-program-' + new Date().toISOString().slice(0,10) + '.xlsx');
}

/** Procesează fișierul Excel importat și trimite intrările la GAS. */
async function onCalendarFileImport(file) {
  if (!file) return;
  // Reset file input
  const inp = $('cal-import-input');
  if (inp) inp.value = '';

  if (typeof XLSX === 'undefined') {
    showToast('SheetJS nu este disponibil. Reconectați-vă la internet.', 'warning');
    return;
  }
  showToast('Se procesează fișierul...', 'info', 5000);

  try {
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if (rows.length < 2) {
      showToast('Fișierul este gol sau are format greșit.', 'error');
      return;
    }

    // Construiește map: technician_name.lower → obiect tehnician (din IndexedDB)
    const techs      = await getAll('technicians');
    const techByName = {};
    techs.forEach(t => { techByName[t.name.toLowerCase().trim()] = t; });

    const entries = [];
    const skipped = [];
    const now     = Date.now();

    for (let i = 1; i < rows.length; i++) {
      const cells = rows[i].map(v => String(v || '').trim());
      const [date, time, tname, cname, addr, notes] = cells;
      if (!date && !tname && !cname) continue; // rând complet gol — skip silențios

      // Validare dată
      if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        skipped.push(`Rândul ${i+1}: format dată invalid "${date}" (trebuie YYYY-MM-DD)`);
        continue;
      }
      // Validare technician
      if (!tname) {
        skipped.push(`Rândul ${i+1}: lipsește numele tehnicianului`);
        continue;
      }
      const tech     = techByName[tname.toLowerCase().trim()];
      const techId   = tech ? tech.technician_id : 'unknown_' + tname.toLowerCase().replace(/\s+/g, '_');
      const techName = tech ? tech.name : tname;
      if (!tech) {
        skipped.push(`Rândul ${i+1}: tehnician "${tname}" nu a fost găsit (va fi marcat "unknown")`);
      }

      entries.push({
        id:              'p_' + now + '_' + i,
        date,
        time:            time || '',
        technician_id:   techId,
        technician_name: techName,
        client_name:     cname || '',
        address:         addr  || '',
        notes:           notes || ''
      });
    }

    if (entries.length === 0) {
      showToast('Nicio intrare validă în fișier.', 'error');
      if (skipped.length) showToast(skipped.slice(0, 3).join('\n'), 'warning', 6000);
      return;
    }

    if (!isSyncConfigured()) {
      showToast('API-ul nu este configurat — intrările nu pot fi salvate.', 'error');
      return;
    }

    const resp = await fetch(SYNC_CONFIG.API_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ action: 'saveCalendarEntries', entries })
    });
    const data = await resp.json();

    if (data.success) {
      let msg = `Import reușit: ${data.saved} intrări salvate.`;
      if (skipped.length) msg += ` (${skipped.length} omise)`;
      showToast(msg, 'success', 5000);
      if (skipped.length) {
        setTimeout(() => showToast('Omise: ' + skipped.slice(0, 3).join(' | '), 'warning', 7000), 1200);
      }
      loadCalendarScreen();
    } else {
      showToast('Eroare la salvare: ' + (data.error || ''), 'error');
    }
  } catch (err) {
    showToast('Eroare la procesarea fișierului: ' + err.message, 'error');
  }
}

/** Exportă calendarul curent ca PDF via browser print. */
function exportCalendarPdf() {
  const content = $('cal-content');
  const label   = $('cal-week-label');
  if (!content || !content.innerHTML.trim()) {
    showToast('Nu există date de exportat.', 'warning');
    return;
  }

  const title    = 'Program Intervenții — ' + (label ? label.textContent : '');
  const printInner = _buildCalendarPrintHtml();

  const printHtml = `<!DOCTYPE html>
<html lang="ro">
<head>
  <meta charset="UTF-8">
  <title>${escHtml(title)}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #111; padding: 20px; }
    h2 { color: #1d4ed8; border-bottom: 2px solid #1d4ed8; padding-bottom: 8px; margin-bottom: 20px; font-size: 17px; }
    .day-group { margin-bottom: 18px; page-break-inside: avoid; }
    .day-header { font-weight: 700; font-size: 13px; color: #374151; padding: 5px 0; border-bottom: 1.5px solid #d1d5db; margin-bottom: 8px; text-transform: uppercase; letter-spacing: .04em; }
    .day-header.today { color: #1d4ed8; }
    .cal-entry { display: flex; gap: 10px; margin-bottom: 7px; padding: 7px 10px; border-left: 3px solid #1d4ed8; background: #eff6ff; border-radius: 5px; }
    .e-time { font-weight: 700; color: #1d4ed8; min-width: 42px; flex-shrink: 0; font-size: 13px; padding-top: 1px; }
    .e-body { flex: 1; min-width: 0; }
    .e-tech { color: #4b5563; font-size: 11px; margin-bottom: 2px; }
    .e-client { font-weight: 600; font-size: 13px; }
    .e-addr { color: #6b7280; font-size: 11px; }
    .e-notes { color: #6b7280; font-size: 11px; font-style: italic; }
    .empty { color: #9ca3af; font-style: italic; font-size: 12px; padding: 4px 0 10px; }
    .footer { margin-top: 24px; font-size: 11px; color: #9ca3af; border-top: 1px solid #e5e7eb; padding-top: 8px; }
    @media print { body { padding: 8mm; } @page { margin: 10mm; } }
  </style>
</head>
<body>
  <h2>${escHtml(title)}</h2>
  ${printInner}
  <div class="footer">Generat de Pool Manager · ${new Date().toLocaleDateString('ro-RO', { day:'2-digit', month:'long', year:'numeric' })}</div>
</body>
</html>`;

  const w = window.open('', '_blank');
  if (!w) {
    showToast('Activați popup-urile pentru a exporta PDF.', 'warning');
    return;
  }
  w.document.write(printHtml);
  w.document.close();
  w.focus();
  setTimeout(() => { w.print(); w.close(); }, 400);
}

/** Construiește HTML-ul intern al calendarului pentru print. */
function _buildCalendarPrintHtml() {
  const content = $('cal-content');
  if (!content) return '';
  const today = new Date().toISOString().slice(0, 10);
  let html = '';

  content.querySelectorAll('.cal-day-group').forEach(dg => {
    const hdr    = dg.querySelector('.cal-day-header');
    const isEmpty= !!dg.querySelector('.cal-day-empty');
    const entEls = dg.querySelectorAll('.cal-entry');

    // Determină data din header pentru colorare "azi"
    const isToday = hdr && hdr.classList.contains('is-today');

    html += '<div class="day-group">';
    html += `<div class="day-header${isToday ? ' today' : ''}">${hdr ? escHtml(hdr.textContent.replace('Azi','').trim()) : ''}</div>`;

    if (isEmpty) {
      html += '<div class="empty">Nicio intervenție planificată</div>';
    } else {
      entEls.forEach(ent => {
        const time   = ent.querySelector('.cal-entry-time');
        const tech   = ent.querySelector('.cal-entry-tech');
        const client = ent.querySelector('.cal-entry-client');
        const addr   = ent.querySelector('.cal-entry-addr');
        const notes  = ent.querySelector('.cal-entry-notes');
        html += '<div class="cal-entry">';
        html += `<div class="e-time">${escHtml(time   ? time.textContent.trim()   : '—')}</div>`;
        html += '<div class="e-body">';
        if (tech)   html += `<div class="e-tech">${escHtml(tech.textContent.trim())}</div>`;
        if (client) html += `<div class="e-client">${escHtml(client.textContent.trim())}</div>`;
        if (addr)   html += `<div class="e-addr">${escHtml(addr.textContent.trim())}</div>`;
        if (notes)  html += `<div class="e-notes">${escHtml(notes.textContent.trim())}</div>`;
        html += '</div></div>';
      });
    }
    html += '</div>';
  });
  return html;
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
    if (data.success) console.log('[CHECKLIST] Synced to GAS');
    else console.warn('[CHECKLIST] GAS sync error:', data.error);
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
      console.log('[CHECKLIST] Updated from GAS (remote newer)');
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

  if (typeof XLSX === 'undefined') {
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
function downloadChecklistTemplate() {
  if (typeof XLSX === 'undefined') {
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

