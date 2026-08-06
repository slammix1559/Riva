// Service Worker - Gestione CdC ITIS Riva

// ═══════════════════════════════════════════════════════════════
// NOTIFICHE PUSH (Firebase Cloud Messaging) — gestisce le notifiche
// quando l'app NON è aperta (schermo spento, browser chiuso, altra
// scheda). Le notifiche a schermo aperto sono gestite direttamente
// in Index.html tramite messaging.onMessage().
// ═══════════════════════════════════════════════════════════════
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyC1rsmLIVUr9drlaD9tNK1Iftj-V1taHY8",
  authDomain: "rivapp-ed917.firebaseapp.com",
  projectId: "rivapp-ed917",
  storageBucket: "rivapp-ed917.firebasestorage.app",
  messagingSenderId: "553925987080",
  appId: "1:553925987080:web:b9c80e97abb6c24c9ffabd"
});

const _messaging = firebase.messaging();

// Notifica di sistema quando l'app è in background/chiusa.
_messaging.onBackgroundMessage(function(payload) {
  const n = payload.notification || {};
  const title = n.title || 'RivApp';
  const options = {
    body: n.body || '',
    icon: 'icon-512.png',
    badge: 'icon-512.png'
  };
  self.registration.showNotification(title, options);
});

// Tocco sulla notifica: apre (o porta in primo piano) l'app.
self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.link) ||
              'https://slammix1559.github.io/Riva/';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(list) {
      for (let i = 0; i < list.length; i++) {
        if (list[i].url.indexOf(url) >= 0 && 'focus' in list[i]) return list[i].focus();
      }
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});

// ═══════════════════════════════════════════════════════════════
// CACHE / OFFLINE (logica originale, invariata)
// ═══════════════════════════════════════════════════════════════
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
