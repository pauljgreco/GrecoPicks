// Keeps the app's own files on the phone so it opens instantly.
// Picks and scores always come live from the Google script (never cached here).
// When a newer version of the app is published, the open page is told to reload once.
// Push notifications are handled by OneSignal's own worker in push/ (separate scope).
const CACHE = 'nfl-picks-v4';
const FILES = ['./', './index.html', './config.js', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './icon-180.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(FILES.map(f => c.add(f).catch(() => {})))));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

async function tellPagesToReload() {
  const pages = await self.clients.matchAll({ type: 'window' });
  pages.forEach(p => p.postMessage('nflpicks-updated'));
}

// Serve the saved copy right away, then fetch a fresh one. If the page or its
// settings changed, save the new copy and ask the open app to reload.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  const isText = e.request.mode === 'navigate' || /(\/|\.html|\.js|\.webmanifest)$/.test(url.pathname);
  e.respondWith(caches.open(CACHE).then(async cache => {
    const cached = await cache.match(e.request, { ignoreSearch: true });
    const fresh = fetch(e.request, { cache: 'no-cache' }).then(async r => {
      if (!r || !r.ok) return r;
      if (cached && isText) {
        const [oldText, newText] = await Promise.all([cached.clone().text(), r.clone().text()]);
        await cache.put(e.request, r.clone());
        if (oldText !== newText) tellPagesToReload();
      } else {
        await cache.put(e.request, r.clone());
      }
      return r;
    }).catch(() => cached);
    return cached || fresh;
  }));
});
