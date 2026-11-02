"""
servidor_web.py
===============
Inicia la versión web multiusuario del Sistema POA - DPDI con el servidor de producción Waitress.

    python servidor_web.py                 # http://0.0.0.0:8080
    python servidor_web.py --port 9000

Variables de entorno opcionales:
    POA_DATA_DIR        carpeta de datos (base de datos, logo, clave de sesión). Por defecto ./data
    POA_DB_PATH         ruta del archivo SQLite (por defecto POA_DATA_DIR/poa_dpdi_web.sqlite3)
    POA_ADMIN_PASSWORD  clave inicial del usuario «admin» (si no se define, se genera y se guarda en
                        POA_DATA_DIR/credenciales_iniciales.txt)
    POA_HTTPS=1         marca la cookie de sesión como segura (use detrás de un proxy HTTPS)
    POA_SESION_HORAS    duración de la sesión (por defecto 8)
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from web.app import VERSION_APP, create_app  # noqa: E402


def main() -> None:
    ap = argparse.ArgumentParser(description="Servidor web del Sistema POA - DPDI")
    ap.add_argument("--host", default=os.environ.get("POA_HOST", "0.0.0.0"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("POA_PORT", "8080")))
    ap.add_argument("--threads", type=int, default=int(os.environ.get("POA_THREADS", "8")))
    args = ap.parse_args()
    app = create_app()
    from waitress import serve
    print(f"Sistema POA - DPDI (web) versión {VERSION_APP} en http://{args.host}:{args.port}  · datos: {app.config['DATA_DIR']}", flush=True)
    opciones = {"host": args.host, "port": args.port, "threads": args.threads,
                "url_scheme": "https" if os.environ.get("POA_HTTPS") == "1" else "http"}
    if os.environ.get("POA_PROXY"):  # IP del proxy inverso HTTPS (IIS, Nginx, Apache) que antecede al servidor
        opciones.update(trusted_proxy=os.environ["POA_PROXY"],
                        trusted_proxy_headers="x-forwarded-for x-forwarded-proto")
    serve(app, **opciones)


if __name__ == "__main__":
    main()
