/* Service Worker — Work Order Map PWA — Enables offline use after first load */

const CACHE_NAME = 'wo-map-v2.5';
const CACHE_FILES = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './manifest.json',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CACHE_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // Always hit the network for API and admin routes (need fresh data)
  const url = new URL(e.request.url);
  if (url.pathname.startsWith('/api/') || url.pathname.startsWith('/admin')) {
    e.respondWith(fetch(e.request));
    return;
  }
  // Cache-first for all other assets (offline support)
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request))
  );
});
