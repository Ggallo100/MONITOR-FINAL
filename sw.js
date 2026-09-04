/* Service worker mínimo: permite instalar la app y que abra sin conexión.
   Estrategia: la red manda, la caché es el plan B. Así siempre ves la última
   versión publicada, pero la app arranca aunque estés sin cobertura. */
const CACHE = 'monitor-nyse-v2';
const SHELL = ['./', './index.html', './marketsurge.html', './manifest.webmanifest', './icon192.png', './icon512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  // Las cotizaciones nunca se cachean: un precio viejo es peor que ninguno.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
        return res;
      })
      .catch(() => caches.match(req).then(m => m || caches.match('./index.html')))
  );
});
