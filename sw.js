// Minimal service worker: cache the app shell so the lock screen still
// opens offline. Live trip data always requires a network round-trip to
// the Apps Script backend, so we deliberately do NOT cache API responses.
const SHELL_CACHE = 'peru-tour-shell-v2';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/tabler-icons/3.46.0/tabler-icons.min.css',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== SHELL_CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = req.url;
  // Never intercept calls to the Apps Script backend or the live FX-rate
  // API — always go live for both (the app itself handles caching/fallback
  // for the FX rates in localStorage).
  if (url.includes('script.google.com') || url.includes('er-api.com')) return;

  // Network-first for the HTML shell itself. This file changes with every
  // deploy (including config like the Apps Script URL) — cache-first here
  // means a device can get stuck forever serving a stale, possibly-broken
  // config after a deploy, with no obvious way to tell. Only fall back to
  // the cached copy when actually offline.
  if (req.mode === 'navigate' || url.endsWith('/index.html') || url.endsWith('/')) {
    event.respondWith(
      fetch(req).then((res) => {
        caches.open(SHELL_CACHE).then((c) => c.put(req, res.clone()));
        return res;
      }).catch(() => caches.match(req))
    );
    return;
  }

  // Cache-first for static assets (icons, manifest, icon font) — these
  // rarely change and this keeps things fast / usable offline.
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).catch(() => cached))
  );
});
