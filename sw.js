const CACHE_NAME = 'engajavr-v1';
const ASSETS = [
  '/engajavr/',
  '/engajavr/index.html',
  '/engajavr/manifest.json',
  '/engajavr/icon-192.png',
  '/engajavr/icon-512.png'
];

// Install: cache assets
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

// Activate: clean old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch: network first, fallback to cache
self.addEventListener('fetch', function(e) {
  // Não interceptar chamadas ao GAS (externas)
  if (e.request.url.indexOf('script.google.com') > -1) return;
  if (e.request.url.indexOf('googleapis.com') > -1) return;

  e.respondWith(
    fetch(e.request)
      .then(function(response) {
        // Atualizar cache com resposta fresca
        if (response && response.status === 200 && e.request.method === 'GET') {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(e.request, clone);
          });
        }
        return response;
      })
      .catch(function() {
        // Offline: servir do cache
        return caches.match(e.request).then(function(cached) {
          return cached || caches.match('/engajavr/index.html');
        });
      })
  );
});
