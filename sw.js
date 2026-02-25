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
    if (e.request.url.includes('api.anthropic.com')) return;
    const url = new URL(e.request.url);
    const isMainAsset = ['index.html', 'style.css', 'app.js', 'sw.js'].some(a => url.pathname.endsWith(a)) || url.pathname.endsWith('/');
    if (isMainAsset) {
        // Network first: always get fresh files, update cache, fallback to cache
        e.respondWith(fetch(e.request).then(r => {
            const clone = r.clone();
            caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
            return r;
        }).catch(() => caches.match(e.request)));
    } else {
        e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
    }
});
