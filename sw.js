// Service Worker do EngajaVR
// Usa caminhos relativos ao escopo do SW para funcionar em qualquer subpath
// (GitHub Pages em /engajavr/ e Cloudflare Workers na raiz /)

const CACHE_VERSION = 'v3';
const CACHE_NAME = 'engajavr-' + CACHE_VERSION;

// Calcula o base path a partir de onde o sw.js foi registrado
// Em GitHub Pages: /engajavr/   Em Cloudflare: /
const BASE = self.registration.scope;

const ASSETS = [
  BASE,
  BASE + 'index.html',
  BASE + 'manifest.json',
  BASE + 'icon-192.png',
  BASE + 'icon-512.png'
];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys()
      .then(function(keys) {
        return Promise.all(
          keys.filter(function(k) { return k !== CACHE_NAME; })
              .map(function(k) { return caches.delete(k); })
        );
      })
      .then(function() { return self.clients.claim(); })
      .then(function() {
        return self.clients.matchAll({ type: 'window' }).then(function(clients) {
          clients.forEach(function(client) {
            client.postMessage({ type: 'SW_UPDATED', version: CACHE_VERSION });
          });
        });
      })
  );
});

self.addEventListener('fetch', function(e) {
  if (e.request.url.indexOf('script.google.com') > -1) return;
  if (e.request.url.indexOf('googleapis.com') > -1) return;

  e.respondWith(
    fetch(e.request)
      .then(function(response) {
        if (response && response.status === 200 && e.request.method === 'GET') {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(e.request, clone);
          });
        }
        return response;
      })
      .catch(function() {
        return caches.match(e.request).then(function(cached) {
          return cached || caches.match(BASE + 'index.html');
        });
      })
  );
});
