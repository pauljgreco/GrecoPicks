// Greco Picks service worker.
// - The app itself (page, scripts, settings): NETWORK FIRST. With internet, opening the app always loads the
//   newest version; the saved copy is only used when offline (or the connection is too slow to answer).
// - Images: saved copy first (they rarely change).
// - Picks and scores always come live from the Google script (never cached here).
// Push notifications: OneSignal's push code is loaded into this same worker (one worker, one scope).
try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) { /* offline: app still works */ }
const CACHE = 'nfl-picks-v7';
const FILES = ['./', './index.html', './config.js', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './icon-180.png'];
const SLOW_MS = 5000;   // after this long with no answer, show the saved copy (and keep loading the new one)

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(FILES.map(f => c.add(new Request(f, { cache: 'reload' })).catch(() => {})))));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
    tellPagesToReload();   // a new version of the app just took over: reload open pages onto it
  })());
});

async function tellPagesToReload() {
  const pages = await self.clients.matchAll({ type: 'window' });
  pages.forEach(p => p.postMessage('nflpicks-updated'));
}

// Saved copies are stored without the ?k=… / ?recap=… part of the address, so one copy serves every link.
const keyFor = req => { const u = new URL(req.url); u.search = ''; return u.toString(); };

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  const isApp = e.request.mode === 'navigate' || /(\/|\.html|\.js|\.webmanifest)$/.test(url.pathname);
  if (/\/sw\.js$/.test(url.pathname)) return;   // never intercept the worker itself
  e.respondWith(isApp ? networkFirst(e) : cacheFirst(e));
});

async function networkFirst(e) {
  const cache = await caches.open(CACHE);
  const key = keyFor(e.request);
  const net = fetch(e.request, { cache: 'no-cache' }).then(async r => {
    if (r && r.ok) await cache.put(key, r.clone());
    return r;
  });
  // Wait for the internet first; fall back to the saved copy only if offline or very slow.
  const slow = new Promise(res => setTimeout(res, SLOW_MS, 'slow'));
  try {
    const first = await Promise.race([net, slow]);
    if (first !== 'slow' && first && first.ok) return first;
  } catch (err) { /* offline */ }
  const cached = await cache.match(key) || await cache.match(e.request, { ignoreSearch: true });
  if (cached) {
    // Showed the saved copy because the connection was slow: when the new one arrives, reload onto it if it changed.
    e.waitUntil(net.then(async r => {
      if (!r || !r.ok) return;
      const [a, b] = await Promise.all([cached.clone().text(), r.clone().text()]);
      if (a !== b) tellPagesToReload();
    }).catch(() => {}));
    return cached;
  }
  return net;
}

async function cacheFirst(e) {
  const cache = await caches.open(CACHE);
  const cached = await cache.match(e.request, { ignoreSearch: true });
  const fresh = fetch(e.request).then(async r => { if (r && r.ok) await cache.put(e.request, r.clone()); return r; }).catch(() => cached);
  return cached || fresh;
}
