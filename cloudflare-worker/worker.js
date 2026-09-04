/**
 * Proxy de Yahoo Finance para MarketSurge Web
 * ---------------------------------------------------------------------------
 * El navegador no puede llamar a Yahoo Finance directamente por la política
 * CORS. Este Worker hace de intermediario: recibe la URL de Yahoo, la pide
 * desde el servidor de Cloudflare y devuelve la respuesta con las cabeceras
 * que el navegador necesita.
 *
 * No es un proxy abierto: solo deja pasar los cinco endpoints públicos de
 * Yahoo que usa la aplicación. Cualquier otra dirección se rechaza, así que
 * nadie puede usar tu Worker para navegar por internet a tu costa.
 *
 * Plan gratuito de Cloudflare: 100.000 peticiones al día. Una sesión normal
 * de análisis gasta unas cientos.
 *
 * Uso:  https://TU-WORKER.workers.dev/?url=<URL de Yahoo codificada>
 * Salud: https://TU-WORKER.workers.dev/         → responde "ok" y la versión
 */

const UPSTREAM_HOSTS = new Set([
  'query1.finance.yahoo.com',
  'query2.finance.yahoo.com'
]);

/* Rutas permitidas y cuántos segundos se guarda cada respuesta en caché.
   Las cotizaciones caducan rápido; los estados financieros cambian una vez
   por trimestre, así que se pueden guardar horas y ahorran peticiones. */
const ROUTES = [
  { re: /^\/v8\/finance\/chart\/[^/]+$/,                                  ttl: 60 },
  { re: /^\/v1\/finance\/search$/,                                        ttl: 600 },
  { re: /^\/ws\/fundamentals-timeseries\/v1\/finance\/timeseries\/[^/]+$/, ttl: 21600 },
  { re: /^\/v10\/finance\/quoteSummary\/[^/]+$/,                          ttl: 3600 },
  { re: /^\/v7\/finance\/quote$/,                                         ttl: 60 }
];

const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
           '(KHTML, like Gecko) Chrome/122.0 Safari/537.36';

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get('Origin') || '';
    const cors = corsHeaders(origin, env);

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (request.method !== 'GET' && request.method !== 'HEAD')
      return json({ error: 'solo se admiten peticiones GET' }, 405, cors);

    const url = new URL(request.url);
    const target = url.searchParams.get('url');

    // Sin parámetro: página de salud, para comprobar que el Worker vive.
    if (!target)
      return json({ ok: true, service: 'yahoo-proxy', version: 1,
                    uso: 'añade ?url=<URL de Yahoo codificada>' }, 200, cors);

    // ---- validación -------------------------------------------------------
    let u;
    try { u = new URL(target); }
    catch { return json({ error: 'la url no es válida' }, 400, cors); }

    if (u.protocol !== 'https:' || !UPSTREAM_HOSTS.has(u.hostname))
      return json({ error: 'solo se permite query1/query2.finance.yahoo.com' }, 403, cors);

    const route = ROUTES.find(r => r.re.test(u.pathname));
    if (!route)
      return json({ error: 'ruta de Yahoo no permitida: ' + u.pathname }, 403, cors);

    // ---- caché ------------------------------------------------------------
    const cache = caches.default;
    const cacheKey = new Request(u.toString(), { method: 'GET' });
    const hit = await cache.match(cacheKey);
    if (hit) return withHeaders(hit, cors, 'HIT');

    // ---- petición a Yahoo -------------------------------------------------
    let upstream;
    try {
      upstream = await fetch(u.toString(), {
        headers: { 'User-Agent': UA, 'Accept': 'application/json', 'Accept-Language': 'en-US,en;q=0.9' },
        signal: AbortSignal.timeout(20000),
        cf: { cacheTtl: route.ttl, cacheEverything: true }
      });
    } catch (e) {
      return json({ error: 'Yahoo no respondió: ' + (e && e.message || e) }, 502, cors);
    }

    const body = await upstream.text();
    if (!upstream.ok)
      return json({ error: 'Yahoo devolvió ' + upstream.status, detalle: body.slice(0, 300) },
                  upstream.status === 404 ? 404 : 502, cors);

    const res = new Response(body, {
      status: 200,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'public, max-age=' + route.ttl
      }
    });
    ctx.waitUntil(cache.put(cacheKey, res.clone()));
    return withHeaders(res, cors, 'MISS');
  }
};

/* --------------------------------------------------------------------------
   Auxiliares
   -------------------------------------------------------------------------- */
function corsHeaders(origin, env) {
  // Por defecto, cualquier origen. Para restringirlo a tu web, define la
  // variable ALLOWED_ORIGINS en el Worker con los orígenes separados por comas,
  // por ejemplo: https://ggallo100.github.io,http://localhost:8000
  const allow = (env && env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
  const ok = !allow.length || (origin && allow.includes(origin));
  return {
    'Access-Control-Allow-Origin': allow.length ? (ok ? origin : 'null') : '*',
    'Access-Control-Allow-Methods': 'GET, HEAD, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Accept',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}
function withHeaders(res, cors, cacheState) {
  const out = new Response(res.body, res);
  for (const [k, v] of Object.entries(cors)) out.headers.set(k, v);
  out.headers.set('X-Proxy-Cache', cacheState);
  return out;
}
function json(obj, status, cors) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, cors)
  });
}
