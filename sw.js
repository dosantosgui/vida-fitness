// Service worker: guarda o app no celular para funcionar sem internet.
const VERSAO = 'vf-v5';
const ARQUIVOS = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './vendor/xlsx.mini.min.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-512.png',
  './icons/apple-180.png'
];

self.addEventListener('install', (e) => {
  // cache: 'reload' busca sempre do servidor (sem pegar cópia velha do navegador).
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(ARQUIVOS.map((u) => new Request(u, { cache: 'reload' })))));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((nomes) => Promise.all(nomes.filter((n) => n !== VERSAO).map((n) => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (e) => {
  if (e.data === 'atualizar') self.skipWaiting();
});

// Primeiro o que está guardado; sem internet, funciona igual.
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then((r) => r || fetch(e.request))
  );
});

// Tocar na notificação abre o app.
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((cs) => {
    const c = cs[0]; if (c) return c.focus(); return self.clients.openWindow('./');
  }));
});
