// Guarda la app y los datos en el celular para que abra sin conexion.
// Estrategia: primero la red (para tener lo mas reciente); si no hay red, la copia guardada.
const CACHE = 'inversiones-v1';
const BASE = ['./', 'index.html', 'manifest.webmanifest', 'icono-192.png', 'icono-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const pedido = e.request;
  if (pedido.method !== 'GET' || new URL(pedido.url).origin !== location.origin) return;
  e.respondWith(
    fetch(pedido).then(resp => {
      if (resp.ok) { const copia = resp.clone(); caches.open(CACHE).then(c => c.put(pedido, copia)); }
      return resp;
    }).catch(() => caches.match(pedido, {ignoreSearch: true}).then(r => r || caches.match('index.html')))
  );
});
