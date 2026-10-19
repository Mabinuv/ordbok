/* MJ Ordbok service worker.
   - The page itself (navigations) is NETWORK-FIRST: online you always get the
     newest deploy; offline (or after a 3s stall) the last good copy is served.
   - Google Fonts are cached so text renders correctly offline.
   - Everything else — notably the Apps Script API (GET and POST) — is never
     touched: the app has its own local-first data cache and outbox. */
const CACHE = 'ordbok-shell-v1';
const SHELL = ['/', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/icon-512.png', '/icons/apple-touch-icon.png'];
const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function networkFirst(req, timeoutMs) {
  return new Promise(resolve => {
    let settled = false;
    const fromCache = () => caches.match(req).then(hit => hit || caches.match('/'));
    const timer = setTimeout(() => { if (settled) return; fromCache().then(hit => { if (hit && !settled) { settled = true; resolve(hit); } }); }, timeoutMs);
    fetch(req).then(res => {
      clearTimeout(timer);
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req.mode === 'navigate' ? '/' : req, copy)); }
      if (!settled) { settled = true; resolve(res); }
    }).catch(() => {
      clearTimeout(timer);
      if (settled) return;
      fromCache().then(hit => { settled = true; resolve(hit || Response.error()); });
    });
  });
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (req.mode === 'navigate' && url.origin === self.location.origin) {
    e.respondWith(networkFirst(req, 3000));
    return;
  }
  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.open(CACHE).then(c => c.match(req).then(hit => {
        const net = fetch(req).then(res => { if (res && (res.ok || res.type === 'opaque')) c.put(req, res.clone()); return res; }).catch(() => hit);
        return hit || net;
      }))
    );
    return;
  }
  if (url.origin === self.location.origin && (url.pathname.startsWith('/icons/') || url.pathname === '/manifest.webmanifest')) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
  }
});
