// Service Worker - Gestione CdC ITIS Riva
const CACHE_NAME = 'cdc-riva-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json'
];

// Installazione: pre-cache dei file base dell'app
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    }).then(function() { return self.skipWaiting(); })
  );
});

// Attivazione: pulizia delle cache vecchie
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    }).then(function() { return self.clients.claim(); })
  );
});

// Fetch:
// - le chiamate all'API (script.google.com) passano SEMPRE dalla rete, mai cache
// - i file dell'app: prima la rete, se offline si usa la cache
self.addEventListener('fetch', function(event) {
  const url = event.request.url;

  // Le chiamate all'API non vanno mai messe in cache
  if (url.indexOf('script.google.com') !== -1) {
    return; // lascia gestire al browser normalmente
  }

  event.respondWith(
    fetch(event.request)
      .then(function(res) {
        // aggiorno la cache con la versione fresca
        const copia = res.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(event.request, copia);
        });
        return res;
      })
      .catch(function() {
        // offline: provo a servire dalla cache
        return caches.match(event.request).then(function(cached) {
          return cached || caches.match('./index.html');
        });
      })
  );
});
