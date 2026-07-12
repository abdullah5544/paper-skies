/* Paper Skies service worker */
const VERSION = 'ps-v3';
const SHELL = ['./', 'index.html', 'manifest.webmanifest',
               'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // never touch API POSTs
  const url = new URL(req.url);
  if (url.hostname.includes('script.google')) return;     // never cache backend

  if (url.origin === location.origin) {
    // app shell: network-first so updates land immediately, cache as offline fallback
    e.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put(req, copy));
        return res;
      }).catch(() => caches.match(req).then(r => r || caches.match('./')))
    );
  } else {
    // CDN (three.js, fonts, NASA textures): cache-first with background refresh
    e.respondWith(
      caches.match(req).then(hit => {
        const refresh = fetch(req).then(res => {
          if (res.ok) caches.open(VERSION).then(c => c.put(req, res.clone()));
          return res;
        }).catch(() => hit);
        return hit || refresh;
      })
    );
  }
});
