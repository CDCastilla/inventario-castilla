// Service worker: guarda la app para que abra sin conexión.
// Al publicar una versión nueva, sube VERSION (y el ?v= de index.html).
const VERSION = 'inv-12';
const ARCHIVOS = [
  './', 'index.html', 'css/app.css?v=5', 'js/app.js?v=12', 'js/gs1.js', 'js/logica.js',
  'manifest.webmanifest', 'iconos/icono-192.png', 'iconos/icono-512.png', 'iconos/monograma-beige.svg'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // el servidor (Apps Script) siempre va a la red
  if (url.searchParams.has('comprobar')) { e.respondWith(new Response(VERSION)); return; }
  e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
