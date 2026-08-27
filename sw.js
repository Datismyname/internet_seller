const CACHE_NAME = 'net-seller-v1';
const urlsToCache = [
  './', './index.html', './manifest.json',
  './icon-72x72.png', './icon-96x96.png', './icon-128x128.png',
  './icon-144x144.png', './icon-152x152.png', './icon-192x192.png',
  './icon-384x384.png', './icon-512x512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then((c) => c.addAll(urlsToCache)));
  self.skipWaiting();
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((names) => Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))));
  self.clients.claim();
});
self.addEventListener('fetch', (e) => {
  e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
