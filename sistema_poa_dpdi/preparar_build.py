"""
preparar_build.py
=================
Paso previo de build_exe.bat: cierra el Sistema POA si está abierto y libera dist\\Sistema_POA_DPDI.exe
(lo borra o, si Windows lo mantiene bloqueado, lo renombra) para que PyInstaller pueda generar el nuevo.
Se hace en Python para no depender de las redirecciones a «nul» del símbolo del sistema.
"""
from __future__ import annotations

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

NOMBRE = "Sistema_POA_DPDI.exe"
DIST = Path(__file__).resolve().parent / "dist"
EXE = DIST / NOMBRE


def cerrar_sistema() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        r = subprocess.run(["taskkill", "/f", "/im", NOMBRE], capture_output=True, text=True)
        if r.returncode == 0:
            print("    Se cerró el Sistema POA que estaba abierto.")
            time.sleep(2)
    except OSError:
        pass


def limpiar_anteriores() -> None:
    for viejo in DIST.glob("Sistema_POA_DPDI_anterior_*.exe"):
        try:
            viejo.unlink()
        except OSError:
            pass


def liberar_exe() -> bool:
    if not EXE.exists():
        return True
    for _ in range(5):
        try:
            EXE.unlink()
            print("    Ejecutable anterior eliminado.")
            return True
        except OSError:
            try:
                destino = DIST / f"Sistema_POA_DPDI_anterior_{datetime.now():%Y%m%d_%H%M%S}.exe"
                EXE.rename(destino)
                print(f"    El ejecutable anterior estaba en uso; se renombró a {destino.name}.")
                return True
            except OSError:
                time.sleep(2)
    return False


def main() -> int:
    print("[0/5] Preparando la carpeta dist ...")
    cerrar_sistema()
    if DIST.exists():
        limpiar_anteriores()
    if not liberar_exe():
        print()
        print(f"*** No se puede reemplazar {EXE} ***")
        print("Cierre el Sistema POA y cualquier ventana que lo esté usando (Explorador de archivos con vista previa,")
        print("antivirus). Si persiste, reinicie el equipo o borre manualmente la carpeta «dist» y vuelva a intentar.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
