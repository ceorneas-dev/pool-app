// sw.js — Service Worker v50 for Pool Manager PWA
// Strategy: cache-first for app shell, network-first for API

'use strict';

const CACHE_NAME   = 'pool-mgmt-v89';
const APP_SHELL    = [
  './',
  './index.html',
  './manifest.json',
  './css/styles.css',
  './js/rules.js',
  './js/db.js',
  './js/sync.js',
  './js/export.js',
  './js/app.js',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

// CDN resources (opaque — cached with care)
const CDN_RESOURCES = [
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
];

// ── Install ───────────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      // Cache app shell (fail silently if any resource missing)
      return cache.addAll(APP_SHELL).catch(err => {
        console.warn('[SW] Some app shell resources failed to cache:', err);
      });
    }).then(() => {
      // Opaque CDN cache (ignore errors)
      return caches.open(CACHE_NAME).then(cache =>
        Promise.allSettled(
          CDN_RESOURCES.map(url =>
            fetch(url, { mode: 'no-cors' }).then(res => cache.put(url, res))
          )
        )
      );
    }).then(() => {
      console.log('[SW] v38 installed');
      return self.skipWaiting();
    })
  );
});

// ── Activate ──────────────────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => {
          console.log('[SW] Deleting old cache:', key);
          return caches.delete(key);
        })
      )
    ).then(() => {
      console.log('[SW] v38 activated');
      return self.clients.claim();
    })
  );
});

// ── Fetch ─────────────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  // Skip non-GET requests
  if (request.method !== 'GET') return;

  // Skip Chrome extension requests
  if (url.protocol === 'chrome-extension:') return;

  // API calls (Google Apps Script) → network-first, no cache
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleapis.com')) {
    event.respondWith(
      fetch(request).catch(() => new Response(JSON.stringify({ error: 'Offline' }), {
        headers: { 'Content-Type': 'application/json' }
      }))
    );
    return;
  }

  // CDN resources → cache-first with network fallback
  if (url.hostname.includes('jsdelivr.net') || url.hostname.includes('cdnjs.')) {
    event.respondWith(
      caches.match(request).then(cached => {
        if (cached) return cached;
        return fetch(request, { mode: 'no-cors' }).then(response => {
          if (response) {
            caches.open(CACHE_NAME).then(cache => cache.put(request, response.clone()));
          }
          return response;
        });
      })
    );
    return;
  }

  // App shell → cache-first with network fallback
  event.respondWith(
    caches.match(request).then(cached => {
      if (cached) {
        // Background revalidate (stale-while-revalidate)
        const revalidate = fetch(request).then(response => {
          if (response && response.status === 200 && response.type === 'basic') {
            caches.open(CACHE_NAME).then(cache => cache.put(request, response.clone()));
          }
          return response;
        }).catch(() => {});
        return cached;
      }

      return fetch(request).then(response => {
        if (!response || response.status !== 200 || response.type !== 'basic') {
          return response;
        }
        const cloned = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(request, cloned));
        return response;
      }).catch(() => {
        // Offline fallback for navigation requests
        if (request.mode === 'navigate') {
          return caches.match('./index.html');
        }
      });
    })
  );
});

// ── Background sync (if supported) ───────────────────────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-interventions') {
    console.log('[SW] Background sync triggered');
    // The actual sync is handled by sync.js in the page context
  }
});

// ── Push notifications ────────────────────────────────────────
self.addEventListener('push', event => {
  if (!event.data) return;
  const data = event.data.json();
  event.waitUntil(
    self.registration.showNotification(data.title || 'Pool Manager', {
      body:  data.body || '',
      icon:  './icons/icon-192.png',
      badge: './icons/icon-192.png'
    })
  );
});
