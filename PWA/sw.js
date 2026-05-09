const CACHE = 'tkd-v1';

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(['/tkd-academy/']))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // Network first — always load fresh GAS content
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
