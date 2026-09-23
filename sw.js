/* Service worker: permite instalar la app y que abra sin conexión.
   Estrategia: la red manda, la caché es el plan B. Así siempre ves la última
   versión publicada, pero la app arranca aunque estés sin cobertura.

   Al cambiar este número se descarta toda la caché anterior. Súbelo cuando
   publiques una versión nueva y quieras forzar el refresco en los equipos. */
const CACHE = 'monitor-nyse-v2';

/* Solo lo imprescindible. Los iconos y el manifiesto entran solos en la caché
   la primera vez que el navegador los pide: incluirlos aquí haría que un
   nombre de archivo distinto abortara la instalación entera. */
const SHELL = ['./', './index.html'];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    // Uno a uno y tolerando fallos: addAll aborta todo si un archivo falta.
    await Promise.all(SHELL.map(u =>
      c.add(new Request(u, {cache: 'reload'})).catch(() => {})
    ));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const ks = await caches.keys();
    await Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  // Las cotizaciones nunca pasan por aquí: un precio viejo es peor que ninguno.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

  // El HTML se pide siempre a la red saltándose la caché HTTP del navegador.
  // GitHub Pages lo sirve con max-age=600, y sin esto el navegador devolvía
  // la versión anterior hasta diez minutos después de publicar.
  const esDocumento = req.mode === 'navigate' || req.destination === 'document'
                   || new URL(req.url).pathname.endsWith('.html');

  e.respondWith((async () => {
    try {
      const res = await fetch(esDocumento ? new Request(req, {cache: 'no-store'}) : req);
      if (res && res.ok) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      }
      return res;
    } catch (err) {
      const hit = await caches.match(req);
      return hit || await caches.match('./index.html') || Response.error();
    }
  })());
});

/* La página puede pedir que el service worker nuevo tome el control ya mismo. */
self.addEventListener('message', e => {
  if (e.data === 'skipWaiting') self.skipWaiting();
});
