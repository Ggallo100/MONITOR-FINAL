# Proxy de Yahoo Finance en Cloudflare Workers

Intermediario para que MarketSurge Web pueda leer Yahoo Finance desde el
navegador. Resuelve el error «No se pudo cargar» y devuelve los fundamentales
(EPS, ventas, P/E, ROE) que los proxies públicos no alcanzan.

- **Coste:** gratis. El plan Free da 100.000 peticiones al día; una sesión de
  análisis normal gasta unos cientos.
- **Tarjeta de crédito:** no hace falta.
- **Tiempo:** unos tres minutos.

## Opción A — desde el panel de Cloudflare (sin instalar nada)

1. Crea una cuenta en <https://dash.cloudflare.com/sign-up>. Confirma el correo.
2. En el menú lateral entra en **Compute (Workers)** → **Create** →
   **Start with Hello World!** → **Get started**.
3. Ponle un nombre, por ejemplo `yahoo-proxy`, y pulsa **Deploy**.
4. Pulsa **Edit code**. Borra todo lo que hay en el editor, pega el contenido
   completo de [`worker.js`](./worker.js) y pulsa **Deploy** otra vez.
5. Copia la dirección que aparece arriba, del tipo
   `https://yahoo-proxy.tu-usuario.workers.dev`.
6. Ábrela en una pestaña para comprobarla: debe responder
   `{"ok":true,"service":"yahoo-proxy",...}`.
7. En MarketSurge Web abre **Ajustes ⚙**, pega esa dirección en
   **«Tu propio proxy de Cloudflare»**, pulsa **Probar** y luego **Guardar**.

La app completa sola lo que falte de la dirección: da igual si la pegas con o
sin barra final.

## Opción B — por línea de comandos

```bash
npm install -g wrangler
wrangler login
cd cloudflare-worker
wrangler deploy
```

Al terminar imprime la dirección del Worker. Pégala en Ajustes igual que arriba.

## Qué hace y qué no

- Solo deja pasar cinco rutas públicas de `query1/query2.finance.yahoo.com`:
  gráficos, series fundamentales, búsqueda, resumen de cotización y cotización.
  Cualquier otra dirección devuelve 403, así que **nadie puede usarlo como
  proxy abierto** para navegar a tu costa.
- Guarda en caché las respuestas: 1 minuto las cotizaciones, 6 horas los
  estados financieros. Eso reduce mucho el consumo de tu cuota.
- Solo admite GET; responde a las peticiones OPTIONS que hace el navegador.
- No guarda datos personales ni registra lo que consultas.

### Restringir el acceso a tu web (opcional)

Por defecto acepta peticiones desde cualquier origen. Para que solo funcione
desde tu sitio, en el panel del Worker ve a **Settings → Variables** y añade la
variable `ALLOWED_ORIGINS` con el valor:

```
https://ggallo100.github.io
```

Si abres la app como archivo local (`file://`), déjalo sin definir: el navegador
envía el origen `null` y la petición se bloquearía.

## Comprobaciones

El código está probado contra un entorno simulado de Workers: rechaza hosts y
rutas ajenas, rechaza `http://`, rechaza POST, propaga los errores de Yahoo,
sirve desde caché en la segunda llamada y aplica la restricción de origen.
