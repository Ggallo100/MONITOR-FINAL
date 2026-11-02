"""
administrar_usuarios.py
=======================
Herramienta del área de Tecnología para la versión web (se ejecuta en el servidor, no desde el navegador).

    python administrar_usuarios.py listar
    python administrar_usuarios.py restablecer USUARIO      # genera una clave temporal y desbloquea la cuenta
    python administrar_usuarios.py desbloquear USUARIO
    python administrar_usuarios.py crear-admin USUARIO      # crea (o reactiva) un Administrador DPDI

Usa la misma base que el servidor (POA_DATA_DIR / POA_DB_PATH). La clave temporal se debe cambiar al ingresar.
"""
from __future__ import annotations

import argparse
import os
import secrets
import sqlite3
import sys
from pathlib import Path

from werkzeug.security import generate_password_hash


def conectar() -> sqlite3.Connection:
    datos = Path(os.environ.get("POA_DATA_DIR") or Path(__file__).resolve().parent / "data")
    ruta = Path(os.environ.get("POA_DB_PATH") or datos / "poa_dpdi_web.sqlite3")
    if not ruta.exists():
        sys.exit(f"No existe la base de datos {ruta}. Inicie primero el servidor o defina POA_DATA_DIR.")
    con = sqlite3.connect(ruta)
    con.row_factory = sqlite3.Row
    return con


def main() -> None:
    ap = argparse.ArgumentParser(description="Administración de usuarios del Sistema POA - DPDI (web)")
    ap.add_argument("accion", choices=["listar", "restablecer", "desbloquear", "crear-admin"])
    ap.add_argument("usuario", nargs="?")
    args = ap.parse_args()
    con = conectar()
    if args.accion == "listar":
        for r in con.execute("SELECT usuario, nombre, rol, activo, bloqueado_hasta, ultimo_acceso FROM usuarios "
                             "ORDER BY usuario"):
            print(f"{r['usuario']:<28} {r['rol']:<12} {'activo' if r['activo'] else 'INACTIVO':<9} "
                  f"bloqueado hasta: {r['bloqueado_hasta'] or '-':<20} último acceso: {r['ultimo_acceso'] or '-'}")
        return
    if not args.usuario:
        sys.exit("Indique el usuario.")
    usuario = args.usuario.strip().lower()
    fila = con.execute("SELECT id FROM usuarios WHERE usuario = ?", (usuario,)).fetchone()
    if args.accion == "desbloquear":
        if not fila:
            sys.exit(f"El usuario {usuario} no existe.")
        con.execute("UPDATE usuarios SET intentos_fallidos = 0, bloqueado_hasta = '' WHERE id = ?", (fila["id"],))
        con.commit()
        print(f"Cuenta {usuario} desbloqueada.")
        return
    clave = secrets.token_urlsafe(9) + "7"
    if args.accion == "restablecer":
        if not fila:
            sys.exit(f"El usuario {usuario} no existe.")
        con.execute("UPDATE usuarios SET password_hash = ?, debe_cambiar_clave = 1, intentos_fallidos = 0, "
                    "bloqueado_hasta = '', activo = 1 WHERE id = ?", (generate_password_hash(clave), fila["id"]))
    else:
        if fila:
            con.execute("UPDATE usuarios SET rol = 'ADMIN', activo = 1, password_hash = ?, debe_cambiar_clave = 1, "
                        "intentos_fallidos = 0, bloqueado_hasta = '' WHERE id = ?",
                        (generate_password_hash(clave), fila["id"]))
        else:
            con.execute("INSERT INTO usuarios (usuario, nombre, rol, password_hash, debe_cambiar_clave) "
                        "VALUES (?, ?, 'ADMIN', ?, 1)", (usuario, "Administrador DPDI", generate_password_hash(clave)))
    con.commit()
    print(f"Usuario: {usuario}\nClave temporal: {clave}\nDebe cambiarla al ingresar.")


if __name__ == "__main__":
    main()
