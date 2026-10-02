/* Service Worker do DocSeg Digital
   Objetivo: deixar o "corpo" do app (HTML, ícones, manifesto) disponível
   mesmo sem internet, para abrir e preencher formulários offline.
   O envio por e-mail (Google Apps Script) continua exigindo internet —
   isso já é tratado separadamente pela fila de reenvio do próprio app.
   Também recebe as notificações push (lembrete diário da PT) quando o
   app está fechado ou em segundo plano. */

importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "BPgYVvJ6OVui3WzovHWNkyhHjg-32cyBnwZmBmqAJq8jcrP__KugYElozOusKgZgG43qUiYvAye99HdB4gvpNUg",
  authDomain: "docseg-app.firebaseapp.com",
  projectId: "docseg-app",
  storageBucket: "docseg-app.firebasestorage.app",
  messagingSenderId: "258167322153",
  appId: "1:258167322153:web:f8e651e39f0ecffcef4e38"
});

const messaging = firebase.messaging();

// Mostra a notificação quando ela chega com o app fechado/minimizado.
messaging.onBackgroundMessage((payload) => {
  const title = (payload.notification && payload.notification.title) || 'DocSeg Digital';
  const options = {
    body: (payload.notification && payload.notification.body) || '',
    icon: './icon-192.png',
    badge: './icon-192.png'
  };
  self.registration.showNotification(title, options);
});

// Ao tocar na notificação, abre (ou foca) o app.
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
      for (const client of list) {
        if ('focus' in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow('./');
    })
  );
});

const CACHE_NAME = 'docseg-shell-v4';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch(() => {}) // não trava a instalação se algum arquivo não existir ainda
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Só cuida de GET no mesmo domínio (a própria página/ícones/manifest).
  // POSTs para o Google Apps Script e chamadas a CDNs externos passam direto,
  // sem interferência — a fila de reenvio do app já cuida do caso offline ali.
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    caches.match(req).then((cached) => {
      const networkFetch = fetch(req)
        .then((res) => {
          if (res && res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          }
          return res;
        })
        .catch(() => cached);
      return cached || networkFetch;
    })
  );
});
