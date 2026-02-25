const CACHE_NAME = 'docgen-v10';

self.addEventListener('install', () => {
    self.skipWaiting();
});

self.addEventListener('activate', (e) => {
    // Delete ALL old caches unconditionally
    e.waitUntil(
        caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k))))
    );
    self.clients.claim();
});

// No fetch handler → every request goes straight to the network.
// This eliminates stale-cache issues while keeping the SW registered for PWA.
