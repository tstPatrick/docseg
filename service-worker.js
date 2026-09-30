/* Service Worker do DocSeg Digital
   Objetivo: deixar o "corpo" do app (HTML, ícones, manifesto) disponível
   mesmo sem internet, para abrir e preencher formulários offline.
   O envio por e-mail (Google Apps Script) continua exigindo internet —
   isso já é tratado separadamente pela fila de reenvio do próprio app. */

const CACHE_NAME = 'docseg-shell-v1';
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
