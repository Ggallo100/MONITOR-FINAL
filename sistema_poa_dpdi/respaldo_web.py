"""
respaldo_web.py
===============
Copia de seguridad de la base de datos de la versión web (consistente aunque haya usuarios conectados).
Conserva los últimos N respaldos. Programe su ejecución diaria (Programador de tareas de Windows o cron).

    python respaldo_web.py                  # usa POA_DATA_DIR o ./data
    python respaldo_web.py --conservar 60
"""
from __future__ import annotations

import argparse
import os
import sqlite3
from datetime import datetime
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser(description="Respaldo de la base de datos del Sistema POA - DPDI (web)")
    datos = Path(os.environ.get("POA_DATA_DIR") or Path(__file__).resolve().parent / "data")
    ap.add_argument("--db", default=os.environ.get("POA_DB_PATH") or str(datos / "poa_dpdi_web.sqlite3"))
    ap.add_argument("--destino", default=str(datos / "respaldos"))
    ap.add_argument("--conservar", type=int, default=30)
    args = ap.parse_args()
    destino = Path(args.destino)
    destino.mkdir(parents=True, exist_ok=True)
    archivo = destino / f"poa_dpdi_{datetime.now():%Y%m%d_%H%M%S}.sqlite3"
    origen = sqlite3.connect(args.db)
    copia = sqlite3.connect(archivo)
    origen.backup(copia)
    copia.close()
    origen.close()
    respaldos = sorted(destino.glob("poa_dpdi_*.sqlite3"))
    for viejo in respaldos[:-args.conservar]:
        viejo.unlink()
    print(f"Respaldo creado: {archivo} ({archivo.stat().st_size / 1024:.0f} KB). Conservados: "
          f"{min(len(respaldos), args.conservar)}")


if __name__ == "__main__":
    main()
