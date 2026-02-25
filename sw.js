const CACHE_NAME = 'docgen-v8';
const ASSETS = ['./', 'index.html', 'style.css', 'app.js', 'manifest.json', 'icon-192.svg'];

self.addEventListener('install', (e) => {
    e.waitUntil(caches.open(CACHE_NAME).then((c) => c.addAll(ASSETS)));
    self.skipWaiting();
});

self.addEventListener('activate', (e) => {
    e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))));
    self.clients.claim();
});

self.addEventListener('fetch', (e) => {
    // Don't cache API calls
    if (e.request.url.includes('api.anthropic.com')) return;
    e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
