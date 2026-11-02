"""
Versión web multiusuario del Sistema de Formulación Presupuestal y Operativa POA - DPDI.

Un único servidor guarda la base de datos; cada usuario ingresa desde su navegador con su propia clave.
Roles:
  ADMIN        Administrador DPDI: todas las unidades, parámetros, catálogos, usuarios, bitácora y consolidado.
  CONSULTA     Consulta DPDI: ve todas las unidades y descarga reportes, sin modificar.
  RESPONSABLE  Responsable de unidad orgánica: formula solo las unidades asignadas, mientras su ficha esté en
               «Borrador» u «Observada».
Las reglas de negocio, importaciones y exportaciones son las mismas del sistema de escritorio (controllers.py,
exporter.py).
"""
from __future__ import annotations

import os
import re
import secrets
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

from flask import (Flask, abort, flash, g, jsonify, redirect, render_template, request, send_file, session,
                   url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

BASE = Path(__file__).resolve().parent.parent
VERSION_APP = "2.4"
# Sufijo ?v= de CSS/JS: al actualizar el programa los navegadores descargan los archivos nuevos (no usan la caché).
VERSION_ESTATICOS = str(int(max(f.stat().st_mtime for f in (Path(__file__).parent / "static").rglob("*") if f.is_file())))
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from controllers import (DETAIL_TABLE, FLAT_TABLE, MASTER_TABLE, POAController, calc_detalle, calc_fegr,  # noqa: E402
                         calc_fing, code_display, control_meta, fmt_money, parse_pct, s, to_float)
from database import MCOLS, MESES, Database  # noqa: E402
from exporter import (export_bd_consolidada, export_libro_oficial, export_pdf,  # noqa: E402
                      export_plantilla_parametros)
from web import schemas  # noqa: E402

ROLES = {"ADMIN": "Administrador DPDI", "CONSULTA": "Consulta DPDI", "RESPONSABLE": "Responsable de unidad orgánica"}
ESTADOS_EDITABLES = ("Borrador", "Observada", "", None)
MAX_INTENTOS = 5
BLOQUEO_MIN = 15
# Paleta del portal institucional (www.autonoma.pe): naranja de marca, títulos gris oscuro y botones naranjas.
TEMA_DEFECTO = {"tema_primario": "#FF6600", "tema_secundario": "#333333", "tema_acento": "#FF6600", "tema_fuente": ""}
# Paleta de la primera versión web (azul/rojo). Si quedó guardada en la base, se reemplaza por la del portal UA.
TEMA_ANTERIOR = {"tema_primario": "#1F3864", "tema_secundario": "#2E5597", "tema_acento": "#C8102E"}
MESES_CORTOS = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"]
KINDS = ("INI", "OO", "AP", "INV")
SLUG_KIND = {"iniciativas": "INI", "otras-operativas": "OO", "actividades-propias": "AP", "inversion": "INV"}


# =============================================================================
# Aplicación
# =============================================================================
def create_app(data_dir: Optional[str] = None, db_path: Optional[str] = None, testing: bool = False) -> Flask:
    app = Flask(__name__, template_folder="templates", static_folder="static")
    data = Path(data_dir or os.environ.get("POA_DATA_DIR") or (BASE / "data"))
    data.mkdir(parents=True, exist_ok=True)
    app.config.update(
        DATA_DIR=data,
        DB_PATH=str(db_path or os.environ.get("POA_DB_PATH") or (data / "poa_dpdi_web.sqlite3")),
        SECRET_KEY=os.environ.get("POA_SECRET_KEY") or _secret_key(data),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=os.environ.get("POA_HTTPS", "0") == "1",
        PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.environ.get("POA_SESION_HORAS", "8"))),
        MAX_CONTENT_LENGTH=25 * 1024 * 1024,
        TESTING=testing,
        JSON_AS_ASCII=False,
    )
    app.json.ensure_ascii = False
    app.json.sort_keys = False

    # Inicializa la base (esquema, catálogos oficiales, migraciones) y el administrador inicial.
    db = Database(app.config["DB_PATH"])
    _crear_admin_inicial(db, data)
    _migrar_tema(db)
    db.close()

    app.teardown_appcontext(_close_db)
    app.before_request(_antes_de_cada_peticion)
    app.context_processor(_contexto_plantillas)
    app.register_error_handler(ValueError, _error_valor)
    app.register_error_handler(PermissionError, _error_permiso)
    app.register_error_handler(413, lambda e: _respuesta_error("El archivo supera el límite de 25 MB.", 413))
    _rutas(app)
    return app


def _secret_key(data: Path) -> str:
    f = data / "secret_key.txt"
    if not f.exists():
        f.write_text(secrets.token_hex(32), encoding="utf-8")
    return f.read_text(encoding="utf-8").strip()


def _migrar_tema(db: Database) -> None:
    """Una sola vez: reemplaza los colores por defecto de la versión anterior por la paleta del portal UA.
    Los colores que la DPDI haya personalizado (distintos de los antiguos por defecto) se respetan."""
    if db.meta("tema_version") == "ua":
        return
    for k, viejo in TEMA_ANTERIOR.items():
        if (db.meta(k) or "").upper() == viejo:
            db.set_meta(k, TEMA_DEFECTO[k])
    db.set_meta("tema_version", "ua")


def _crear_admin_inicial(db: Database, data: Path) -> None:
    if db.scalar("SELECT COUNT(*) FROM usuarios"):
        return
    clave = os.environ.get("POA_ADMIN_PASSWORD") or secrets.token_urlsafe(9)
    db.insert("usuarios", {"usuario": "admin", "nombre": "Administrador DPDI", "rol": "ADMIN",
                           "password_hash": generate_password_hash(clave), "debe_cambiar_clave": 1})
    aviso = (f"Usuario inicial: admin\nClave inicial: {clave}\n"
             "Cámbiela al ingresar. Borre este archivo después de leerlo.\n")
    (data / "credenciales_iniciales.txt").write_text(aviso, encoding="utf-8")
    print("=" * 70 + "\n" + aviso + "=" * 70, flush=True)


# =============================================================================
# Contexto por petición
# =============================================================================
def db() -> Database:
    if "db" not in g:
        from flask import current_app
        g.db = Database(current_app.config["DB_PATH"], seed=False, init=False)
    return g.db


def ctrl() -> POAController:
    if "ctrl" not in g:
        g.ctrl = POAController(db())
    return g.ctrl


def _close_db(_exc=None):
    d = g.pop("db", None)
    if d is not None:
        d.close()


def usuario() -> Optional[dict]:
    if "usuario" not in g:
        g.usuario = None
        uid = session.get("user_id")
        if uid:
            u = db().get("usuarios", uid)
            if u and u["activo"]:
                g.usuario = u
    return g.usuario


def es_admin() -> bool:
    u = usuario()
    return bool(u and u["rol"] == "ADMIN")


def unidades_permitidas() -> list[int]:
    u = usuario()
    if not u:
        return []
    if u["rol"] in ("ADMIN", "CONSULTA"):
        return [r["id"] for r in db().query("SELECT id FROM unidades_organicas ORDER BY codigo_area")]
    return [r["unidad_id"] for r in db().query(
        "SELECT uu.unidad_id FROM usuario_unidades uu JOIN unidades_organicas u ON u.id = uu.unidad_id "
        "WHERE uu.usuario_id = ? ORDER BY u.codigo_area", (u["id"],))]


def unidad_actual() -> Optional[int]:
    permitidas = unidades_permitidas()
    uid = session.get("uid")
    if uid not in permitidas:
        uid = permitidas[0] if permitidas else None
        session["uid"] = uid
    return uid


def puede_editar(uid: Optional[int]) -> bool:
    u = usuario()
    if not u or uid is None:
        return False
    if u["rol"] == "ADMIN":
        return True
    if u["rol"] != "RESPONSABLE" or uid not in unidades_permitidas():
        return False
    estado = db().scalar("SELECT estado_ficha FROM unidades_organicas WHERE id=?", (uid,))
    return estado in ESTADOS_EDITABLES


def exigir_edicion(uid: Optional[int]) -> None:
    if uid not in unidades_permitidas():
        raise PermissionError("No tiene acceso a esta unidad orgánica.")
    if not puede_editar(uid):
        raise PermissionError("La ficha no se puede modificar: es de solo lectura para su rol o ya fue enviada a "
                              "la DPDI (estado «Enviada» o «Aprobada»).")


def ip() -> str:
    return (request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0]).strip()


def log(accion: str, detalle: str = "") -> None:
    u = usuario()
    db().log(u["usuario"] if u else "-", accion, detalle, ip())


def _respuesta_error(msg: str, code: int):
    if request.path.startswith("/api/") or request.path.startswith(("/cargas/", "/descargas/")):
        return jsonify(error=msg), code
    flash(msg, "error")
    return redirect(request.referrer or url_for("inicio")), 302


def _error_valor(e):
    return _respuesta_error(str(e), 400)


def _error_permiso(e):
    return _respuesta_error(str(e) or "Acceso denegado.", 403)


LIBRES = {"login", "static", "marca_logo", "salud"}


def _antes_de_cada_peticion():
    session.permanent = True
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(24)
    if request.endpoint in LIBRES or request.endpoint is None:
        if request.method == "POST" and request.endpoint == "login":
            _verificar_csrf()
        return None
    u = usuario()
    if not u:
        if request.path.startswith("/api/"):
            return jsonify(error="Sesión expirada. Vuelva a ingresar."), 401
        return redirect(url_for("login", next=request.path))
    if request.method not in ("GET", "HEAD", "OPTIONS"):
        _verificar_csrf()
    if u["debe_cambiar_clave"] and request.endpoint not in ("cambiar_clave", "logout"):
        if request.path.startswith("/api/"):
            return jsonify(error="Debe cambiar su clave antes de continuar."), 403
        return redirect(url_for("cambiar_clave"))
    return None


def _verificar_csrf():
    token = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token")
    if not token or not secrets.compare_digest(token, session.get("csrf", "")):
        abort(400, "Token de seguridad inválido. Recargue la página.")


def tema() -> dict:
    t = dict(TEMA_DEFECTO)
    for k in t:
        v = db().meta(k)
        if v:
            t[k] = v
    return t


def _logo_path() -> Optional[Path]:
    from flask import current_app
    d = Path(current_app.config["DATA_DIR"]) / "marca"
    for ext in ("svg", "png", "jpg", "jpeg", "webp"):
        p = d / f"logo.{ext}"
        if p.exists():
            return p
    return None


def _contexto_plantillas():
    u = usuario() if session.get("user_id") else None
    prm = db().parametros()
    cierre = ""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s(prm.get("fecha_limite")))
    if m:
        cierre = f"{int(m.group(3))} {MESES_CORTOS[int(m.group(2)) - 1]}"
    ctx = {"csrf": session.get("csrf", ""), "usuario": u, "roles": ROLES, "tema": tema(),
           "hay_logo": _logo_path() is not None, "anio": s(prm.get("anio_formulacion")), "cierre": cierre,
           "institucion": s(prm.get("institucion")), "unidad": None, "unidades": [],
           "editable": False, "es_admin": bool(u and u["rol"] == "ADMIN"), "v_est": VERSION_ESTATICOS}
    if u:
        c = ctrl()
        ctx["anio"] = c.anio
        uid = unidad_actual()
        ctx["unidades"] = [{"id": x["id"], "label": c.unidad_display(x)}
                           for x in c.unidades() if x["id"] in unidades_permitidas()]
        ctx["unidades"].sort(key=lambda x: x["label"])
        if uid:
            un = c.unidad(uid)
            ctx["unidad"] = {**un, "display": c.unidad_display(un)}
            ctx["editable"] = puede_editar(uid)
    return ctx


# =============================================================================
# Utilidades de serialización
# =============================================================================
def _limpio(r: dict) -> dict:
    out = {}
    for k, v in r.items():
        if k in ("_detalle",):
            continue
        if k == "_c":
            continue
        out[k] = v
    return out


def _fila_master(kind: str, r: dict) -> dict:
    c = ctrl()
    out = _limpio(r)
    if kind == "INI":
        out["eje_nombre"] = r["eje"]
        out["eje"] = c.eje_codigo(r["eje"]) or r["eje"]
        out["_oe"] = c.oe_display(r["oe"]) if r["oe"] else ""
        out["prioridad"] = s(r["prioridad"])
    elif kind == "INV":
        out["prioridad"] = s(r["prioridad"])
    elif kind == "OO":
        cat = next((x["codigo"] for x in c.db.catalogo("CAT_OO") if x["descripcion"] == r["categoria"]), "")
        out["categoria_codigo"] = cat
    if kind == "AP":
        t = to_float(r.get("tarifa"))
        out["_tarifa_neta"] = r.get("_tarifa_neta")
        out["tarifa"] = t
    return out


def _fila_detalle(r: dict) -> dict:
    cc = r["_c"]
    out = _limpio(r)
    out.update(_cu=cc["costo_unit_soles"], _total=cc["total"], _prog=cc["programado"], _control=cc["control"],
               meses=[r[m] for m in MCOLS])
    for i, v in enumerate(cc["meses"]):
        out[f"_m{i}"] = v or None
    return out


def _fila_flat(kind: str, r: dict) -> dict:
    cc = r["_c"]
    out = _limpio(r)
    out["meses"] = [r[m] for m in MCOLS]
    if kind == "FI":
        out.update(_soles=cc["tarifa_soles"], _bruto=cc["bruto"], _desc=cc["descuentos"], _neto=cc["neto"],
                   _cobranza=fmt_money(cc["cobranza_total"]), _siguiente=fmt_money(cc["por_cobrar_siguiente"]),
                   _incobrable=fmt_money(cc["incobrable"]))
    else:
        out.update(_soles=cc["costo_unit_soles"], _coste=cc["coste_total"], _total=cc["total"],
                   _control=cc["control"], _n=cc["n_pagos"])
    for i, v in enumerate(cc["meses"]):
        out[f"_m{i}"] = v or None
    return out


def _con_meses(data: dict) -> dict:
    data = dict(data)
    meses = data.pop("meses", None)
    if isinstance(meses, list):
        for i, m in enumerate(MCOLS):
            data[m] = meses[i] if i < len(meses) else None
    return {k: v for k, v in data.items() if not k.startswith("_")}


def _master_de_unidad(kind: str, mid: int) -> dict:
    m = db().get(MASTER_TABLE[kind], mid)
    if not m or m["unidad_id"] not in unidades_permitidas():
        raise PermissionError("Registro no encontrado o sin acceso.")
    return m


def _detalle_de_unidad(kind: str, did: int) -> tuple[dict, dict]:
    d = db().get(DETAIL_TABLE[kind], did)
    if not d:
        raise PermissionError("Registro no encontrado.")
    return d, _master_de_unidad(kind, d["master_id"])


def _flat_de_unidad(kind: str, rid: int) -> dict:
    r = db().get(FLAT_TABLE[kind], rid)
    if not r or r["unidad_id"] not in unidades_permitidas():
        raise PermissionError("Registro no encontrado o sin acceso.")
    return r


def _uid_param() -> int:
    uid = request.args.get("u", type=int) or unidad_actual()
    if uid not in unidades_permitidas():
        raise PermissionError("No tiene acceso a esta unidad orgánica.")
    return uid


def _techo_resumen(uid: int, kind: str) -> dict:
    cons = ctrl().consumo_techo(uid, kind)
    return {k: cons[k] for k in ("ingresos", "egresos", "techo", "uso", "saldo", "estado")}


def _guardar_upload(campo: str = "archivo", extensiones=(".xlsx", ".xlsm", ".csv")) -> Path:
    f = request.files.get(campo)
    if not f or not f.filename:
        raise ValueError("Seleccione un archivo.")
    nombre = secure_filename(f.filename) or "archivo.xlsx"
    if not nombre.lower().endswith(extensiones):
        raise ValueError("Formato no permitido. Use: " + ", ".join(extensiones))
    tmp = Path(tempfile.mkdtemp(prefix="poa_")) / nombre
    f.save(tmp)
    return tmp


def _enviar(path: Path, nombre: str):
    """Envía el archivo generado y elimina la carpeta temporal (no quedan copias en el servidor)."""
    import io
    import shutil
    datos = io.BytesIO(path.read_bytes())
    shutil.rmtree(path.parent, ignore_errors=True)
    return send_file(datos, as_attachment=True, download_name=nombre)


def _importar_parametros_auto(path: Path) -> dict:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    hojas = set(wb.sheetnames)
    wb.close()
    if {"Datos_Generales", "Unidades_Techos"} & hojas:
        return {"Archivo reconocido": "Plantilla de Parámetros y Catálogos", **ctrl().import_plantilla_parametros(path)}
    if "Parámetros" in hojas:
        return {"Archivo reconocido": "Parámetros de una Ficha Excel", **ctrl().import_parametros(path)}
    raise ValueError("El archivo no es la plantilla de Parámetros y Catálogos (hojas Datos_Generales, Unidades_Techos…) "
                     "ni una Ficha_Formulacion_POA (hoja «Parámetros»).")


def _procesar_upload(fn, extensiones=(".xlsx", ".xlsm", ".csv")):
    """Guarda el archivo subido en una carpeta temporal, lo procesa y la elimina."""
    import shutil
    path = _guardar_upload(extensiones=extensiones)
    try:
        return fn(path)
    finally:
        shutil.rmtree(path.parent, ignore_errors=True)


# =============================================================================
# Rutas
# =============================================================================
def _rutas(app: Flask) -> None:  # noqa: C901 — registro de rutas
    # ----------------------------------------------------------------- acceso
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "POST":
            nombre = request.form.get("usuario", "").strip()
            clave = request.form.get("clave", "")
            u = db().query_one("SELECT * FROM usuarios WHERE usuario = ?", (nombre,))
            ahora = datetime.now()
            if u and u["bloqueado_hasta"] and u["bloqueado_hasta"] > ahora.isoformat(timespec="seconds"):
                flash(f"Cuenta bloqueada temporalmente por intentos fallidos. Intente después de "
                      f"{u['bloqueado_hasta'][11:16]}.", "error")
                return render_template("login.html", usuario_previo=nombre), 429
            if not u or not u["activo"] or not check_password_hash(u["password_hash"], clave):
                if u:
                    n = u["intentos_fallidos"] + 1
                    hasta = (ahora + timedelta(minutes=BLOQUEO_MIN)).isoformat(timespec="seconds") \
                        if n >= MAX_INTENTOS else ""
                    db().update("usuarios", u["id"], {"intentos_fallidos": 0 if hasta else n, "bloqueado_hasta": hasta})
                db().log(nombre or "-", "Ingreso fallido", "", ip())
                flash("Usuario o clave incorrectos.", "error")
                return render_template("login.html", usuario_previo=nombre), 401
            session.clear()
            session["user_id"] = u["id"]
            session["csrf"] = secrets.token_urlsafe(24)
            session.permanent = True
            db().update("usuarios", u["id"], {"intentos_fallidos": 0, "bloqueado_hasta": "",
                                              "ultimo_acceso": ahora.isoformat(sep=" ", timespec="seconds")})
            g.pop("usuario", None)
            log("Ingreso al sistema")
            destino = request.args.get("next") or url_for("inicio")
            return redirect(destino if destino.startswith("/") and not destino.startswith("//") else url_for("inicio"))
        return render_template("login.html", usuario_previo="")

    @app.post("/logout")
    def logout():
        log("Salida del sistema")
        session.clear()
        return redirect(url_for("login"))

    @app.route("/cambiar-clave", methods=["GET", "POST"])
    def cambiar_clave():
        u = usuario()
        if request.method == "POST":
            actual, nueva, rep = (request.form.get(k, "") for k in ("actual", "nueva", "repetir"))
            if not check_password_hash(u["password_hash"], actual):
                flash("La clave actual no es correcta.", "error")
            elif nueva != rep:
                flash("Las claves nuevas no coinciden.", "error")
            elif not _clave_valida(nueva):
                flash("La clave debe tener al menos 8 caracteres, con letras y números.", "error")
            elif check_password_hash(u["password_hash"], nueva):
                flash("La nueva clave debe ser distinta de la actual.", "error")
            else:
                db().update("usuarios", u["id"], {"password_hash": generate_password_hash(nueva),
                                                  "debe_cambiar_clave": 0})
                log("Cambio de clave")
                flash("Clave actualizada.", "ok")
                return redirect(url_for("inicio"))
        return render_template("cambiar_clave.html")

    @app.get("/salud")
    def salud():
        return jsonify(ok=True, version=VERSION_APP, hora=datetime.now().isoformat(timespec="seconds"))

    @app.get("/marca/logo")
    def marca_logo():
        p = _logo_path()
        if not p:
            abort(404)
        return send_file(p, max_age=300)

    # ----------------------------------------------------------------- páginas
    @app.get("/")
    def inicio():
        return render_template("inicio.html", pagina="inicio")

    @app.get("/<slug>")
    def maestro_detalle(slug):
        kind = SLUG_KIND.get(slug)
        if not kind:
            abort(404)
        cfg = schemas.MASTER[kind]
        return render_template("maestro_detalle.html", pagina=slug, kind=kind, cfg=cfg,
                               detalle_fields=schemas.DETALLE_FIELDS, detalle_columns=schemas.DETALLE_COLUMNS)

    @app.get("/funcionamiento")
    def funcionamiento():
        return render_template("funcionamiento.html", pagina="funcionamiento", flat=schemas.FLAT)

    @app.get("/resumen")
    def resumen():
        return render_template("resumen.html", pagina="resumen")

    @app.get("/admin/<entidad>")
    def admin(entidad):
        if entidad == "tablero":
            if usuario()["rol"] not in ("ADMIN", "CONSULTA"):
                abort(403)
            return render_template("admin_tablero.html", pagina="tablero")
        if not es_admin():
            abort(403)
        if entidad == "parametros":
            from controllers import PARAM_GENERALES
            return render_template("admin_parametros.html", pagina="parametros", generales=PARAM_GENERALES,
                                   extra=schemas.PARAM_FIELDS_EXTRA)
        if entidad == "bitacora":
            return render_template("admin_bitacora.html", pagina="bitacora")
        ents = schemas.admin_entities()
        if entidad not in ents:
            abort(404)
        return render_template("admin_crud.html", pagina=entidad, entidad=entidad, cfg=ents[entidad])

    @app.get("/ayuda")
    def ayuda():
        return render_template("ayuda.html", pagina="ayuda", manuales=_manuales_visibles())

    @app.get("/ayuda/<nombre>")
    def ayuda_manual(nombre):
        disponibles = {m["archivo"]: m for m in _manuales_visibles()}
        if nombre not in disponibles:
            abort(404)
        ruta = BASE / "manuales" / "pdf" / nombre
        if not ruta.exists():
            abort(404)
        return send_file(ruta, mimetype="application/pdf", download_name=nombre)

    # --------------------------------------------------------- API: contexto
    @app.post("/api/contexto")
    def api_contexto():
        uid = (request.get_json(force=True) or {}).get("uid")
        uid = int(uid) if uid else None
        if uid not in unidades_permitidas():
            raise PermissionError("No tiene acceso a esa unidad.")
        session["uid"] = uid
        return jsonify(ok=True)

    @app.get("/api/listas")
    def api_listas():
        c = ctrl()
        dbx = c.db
        ejes = dbx.catalogo("EJE")
        objs: dict[str, list] = {}
        obj_oe = {}
        for o in dbx.catalogo("OBJ"):
            objs.setdefault(o["eje"], []).append({"value": o["codigo"], "label": f"{o['codigo']} · {o['descripcion']}"})
            obj_oe[o["codigo"]] = c.oe_display(o["padre"]) if o["padre"] else ""
        gg, desc = {}, {}
        for tipo in ("Egresos", "Ingresos", "Inversion"):
            gg[tipo] = [{"value": p, "label": p} for p in c.partidas_gg(tipo)]
            desc[tipo] = {p: c.descripcion_partida(p, tipo) for p in c.partidas_gg(tipo)}
        cat = lambda t: [{"value": x["codigo"], "label": f"{x['codigo']} · {x['descripcion']}"}  # noqa: E731
                         for x in dbx.catalogo(t)]
        plano = lambda t: [{"value": x["codigo"], "label": x["codigo"]} for x in dbx.catalogo(t)]  # noqa: E731
        sust_egr, sust_ing = plano("SUSTENTO_EGR"), plano("SUSTENTO_ING")
        cuentas = lambda t: [{"value": p, "label": p} for p in c.partidas(t)]  # noqa: E731
        return jsonify(
            ejes=[{"value": e["codigo"], "label": f"{e['codigo']} · {e['descripcion']}"} for e in ejes],
            oes=cat("OE"), objs=objs, obj_oe=obj_oe, origenes=plano("ORIGEN"), prioridades=cat("PRIORIDAD"),
            cat_oo=[{"value": x["descripcion"], "label": f"{x['codigo']} · {x['descripcion']}"}
                    for x in dbx.catalogo("CAT_OO")],
            tipos_ap=plano("TIPO_AP"), tipos_inv=plano("TIPO_INV"), gg=gg, gg_desc=desc,
            sustento={"Ingresos": sust_ing, "Egresos": sust_egr, "Inversion": sust_egr},
            sustento_ing=sust_ing, sustento_egr=sust_egr,
            sustento_todos=sust_egr + [x for x in sust_ing if x not in sust_egr],
            cuentas_ing=cuentas("Ingresos"), cuentas_egr=cuentas("Egresos"),
            cuentas_todas=cuentas("Egresos") + cuentas("Ingresos"),
            macrogrupos=[{"value": m, "label": m} for m in c.macrogrupos()],
            macrogrupos_plantilla=[{"value": "*", "label": "* (todos)"}] +
            [{"value": m, "label": m} for m in c.macrogrupos()],
            ambitos=[{"value": "GLOBAL", "label": "GLOBAL"}] +
            [{"value": u["codigo_area"], "label": c.unidad_display(u)} for u in c.unidades()],
            unidades=[{"value": u["id"], "label": c.unidad_display(u)}
                      for u in sorted(c.unidades(), key=lambda x: s(x["codigo_area"]))],
            tipos_detalle={"INI": ["Egresos", "Ingresos"], "OO": ["Egresos", "Ingresos"],
                           "AP": ["Ingresos", "Egresos"], "INV": ["Inversion"]},
            meses=MESES,
        )

    # ------------------------------------------------------------ API: inicio
    @app.get("/api/inicio")
    def api_inicio():
        uid = _uid_param()
        c = ctrl()
        u = c.unidad(uid)
        tpl = []
        for r in c.plantilla_aplicable(uid):
            try:
                monto = c.monto_plantilla(r, u)
            except ValueError:
                monto = None
            tpl.append({"tipo_flujo": r["tipo_flujo"], "partida": r["partida"], "base_calculo": r["base_calculo"],
                        "_monto": monto})
        codigos = {k: {"registrados": len(c._existing_codes(k, uid)), "siguiente": c.next_code(k, uid)}
                   for k in ("INI", "FI", "FE", "OO", "AP", "INV")}
        techos = {k: _techo_resumen(uid, k) for k in ("INI", "FE", "OO", "INV", "AP")}
        return jsonify(unidad={**u, "display": c.unidad_display(u)}, plantilla=tpl, codigos=codigos, techos=techos,
                       editable=puede_editar(uid), anio=c.anio)

    @app.put("/api/unidad/datos")
    def api_unidad_datos():
        uid = _uid_param()
        exigir_edicion(uid)
        data = request.get_json(force=True) or {}
        campos = ("responsable", "cargo", "correo", "fecha_presentacion", "version_ficha", "elaborado_nombre",
                  "elaborado_cargo", "elaborado_fecha", "revisado_nombre", "revisado_cargo", "revisado_fecha")
        db().update("unidades_organicas", uid, {k: s(data.get(k)) for k in campos if k in data})
        log("Datos de la unidad", ctrl().unidad(uid)["codigo_area"])
        return jsonify(ok=True)

    @app.post("/api/codigos")
    def api_codigos():
        uid = _uid_param()
        exigir_edicion(uid)
        q = request.get_json(force=True) or {}
        cant = {k: int(to_float(q.get(k), 0) or 0) for k in ("INI", "FI", "FE", "OO", "AP", "INV")}
        if any(v < 0 or v > 500 for v in cant.values()) or not sum(cant.values()):
            raise ValueError("Indique cantidades entre 0 y 500 (al menos una mayor que cero).")
        creados = ctrl().generar_codigos(uid, cant)
        log("Generar códigos", f"{len(creados)} códigos")
        return jsonify(creados=[{"tipo": k, "codigo": c_, "codigo_bd": bd} for k, c_, bd in creados])

    @app.post("/api/funcionamiento/correr")
    def api_correr():
        uid = _uid_param()
        exigir_edicion(uid)
        reemplazar = bool((request.get_json(force=True) or {}).get("reemplazar", True))
        rep = ctrl().correr_funcionamiento(uid, reemplazar=reemplazar)
        log("Correr funcionamiento", str(rep))
        return jsonify(rep)

    @app.post("/api/estado")
    def api_estado():
        uid = _uid_param()
        data = request.get_json(force=True) or {}
        accion = data.get("accion")
        u = usuario()
        actual = db().scalar("SELECT estado_ficha FROM unidades_organicas WHERE id=?", (uid,)) or "Borrador"
        transiciones = {
            "enviar": ({"Borrador", "Observada"}, "Enviada", ("RESPONSABLE", "ADMIN")),
            "observar": ({"Enviada", "Aprobada"}, "Observada", ("ADMIN",)),
            "aprobar": ({"Enviada"}, "Aprobada", ("ADMIN",)),
            "reabrir": ({"Enviada", "Aprobada", "Observada"}, "Borrador", ("ADMIN",)),
        }
        if accion not in transiciones:
            raise ValueError("Acción no válida.")
        desde, hacia, roles = transiciones[accion]
        if u["rol"] not in roles:
            raise PermissionError("Su rol no puede realizar esta acción.")
        if actual not in desde:
            raise ValueError(f"No se puede «{accion}» una ficha en estado «{actual}».")
        obs = s(data.get("observacion"))
        if accion == "observar" and not obs:
            raise ValueError("Escriba la observación para la unidad.")
        cambios = {"estado_ficha": hacia, "estado_fecha": datetime.now().isoformat(sep=" ", timespec="minutes"),
                   "estado_usuario": u["usuario"]}
        if accion == "observar":
            cambios["estado_obs"] = obs
        elif accion in ("aprobar", "reabrir"):
            cambios["estado_obs"] = ""
        db().update("unidades_organicas", uid, cambios)
        log(f"Ficha {hacia.lower()}", f"{ctrl().unidad(uid)['codigo_area']} {obs}")
        return jsonify(ok=True, estado=hacia)

    # ----------------------------------------------- API: maestro-detalle
    @app.get("/api/md/<kind>")
    def api_md_list(kind):
        _kind(kind)
        uid = _uid_param()
        rows = [_fila_master(kind, r) for r in ctrl().list_masters(kind, uid)]
        return jsonify(rows=rows, techo=_techo_resumen(uid, schemas.MASTER[kind]["techo"]),
                       editable=puede_editar(uid), siguiente=ctrl().next_code(kind, uid))

    @app.post("/api/md/<kind>")
    @app.put("/api/md/<kind>/<int:mid>")
    def api_md_save(kind, mid=None):
        _kind(kind)
        uid = _master_de_unidad(kind, mid)["unidad_id"] if mid else _uid_param()
        exigir_edicion(uid)
        data = {k: v for k, v in (request.get_json(force=True) or {}).items() if not k.startswith("_")}
        if kind == "INI" and not s(data.get("nombre")):
            raise ValueError("Registre el nombre de la iniciativa.")
        new_id = ctrl().save_master(kind, uid, data, mid)
        log(f"Guardar {kind}", s(data.get("codigo")))
        return jsonify(id=new_id, row=_fila_master(kind, ctrl().get_master(kind, new_id)))

    @app.delete("/api/md/<kind>/<int:mid>")
    def api_md_delete(kind, mid):
        _kind(kind)
        m = _master_de_unidad(kind, mid)
        exigir_edicion(m["unidad_id"])
        ctrl().delete_master(kind, mid)
        log(f"Eliminar {kind}", m["codigo"])
        return jsonify(ok=True)

    @app.post("/api/md/<kind>/<int:mid>/duplicar")
    def api_md_dup(kind, mid):
        _kind(kind)
        m = _master_de_unidad(kind, mid)
        exigir_edicion(m["unidad_id"])
        c = ctrl()
        src = c.get_master(kind, mid)
        data = {k: v for k, v in src.items() if not k.startswith("_") and k not in ("id", "codigo")}
        new_id = c.save_master(kind, m["unidad_id"], data)
        for dd in c.list_detalle(kind, mid):
            c.save_detalle(kind, new_id, {k: v for k, v in dd.items() if not k.startswith("_") and k != "id"})
        log(f"Duplicar {kind}", m["codigo"])
        return jsonify(id=new_id)

    @app.post("/api/md/AP/<int:mid>/ingreso-tarifa")
    def api_ap_ingreso(mid):
        m = _master_de_unidad("AP", mid)
        exigir_edicion(m["unidad_id"])
        ctrl().crear_ingreso_desde_tarifa(mid)
        return jsonify(ok=True)

    @app.get("/api/md/<kind>/<int:mid>/detalle")
    def api_det_list(kind, mid):
        _kind(kind)
        _master_de_unidad(kind, mid)
        m = ctrl().get_master(kind, mid)
        return jsonify(rows=[_fila_detalle(r) for r in ctrl().list_detalle(kind, mid)],
                       master=_fila_master(kind, m))

    @app.post("/api/md/<kind>/<int:mid>/detalle")
    @app.put("/api/det/<kind>/<int:did>")
    def api_det_save(kind, mid=None, did=None):
        _kind(kind)
        if did:
            d, m = _detalle_de_unidad(kind, did)
            mid = m["id"]
        else:
            m = _master_de_unidad(kind, mid)
        exigir_edicion(m["unidad_id"])
        data = _con_meses(request.get_json(force=True) or {})
        if not s(data.get("partida")):
            raise ValueError("Seleccione la Partida.")
        new = ctrl().save_detalle(kind, mid, data, did)
        log(f"Guardar detalle {kind}", f"{m['codigo']} {data.get('partida')}")
        return jsonify(id=new)

    @app.delete("/api/det/<kind>/<int:did>")
    def api_det_delete(kind, did):
        _kind(kind)
        d, m = _detalle_de_unidad(kind, did)
        exigir_edicion(m["unidad_id"])
        ctrl().delete_detalle(kind, did)
        log(f"Eliminar detalle {kind}", m["codigo"])
        return jsonify(ok=True)

    # --------------------------------------------------- API: funcionamiento
    @app.get("/api/fun/<kind>")
    def api_fun_list(kind):
        _flat(kind)
        uid = _uid_param()
        c = ctrl()
        rows = [_fila_flat(kind, r) for r in c.list_flat(kind, uid)]
        p = c.params
        info = {"techo": _techo_resumen(uid, "FE"), "reajuste": p.reajuste, "mes_reajuste": MESES[p.mes_reajuste - 1],
                "morosidad": p.morosidad, "meta_ingresos": c.unidad(uid)["meta_ingresos_func"]}
        return jsonify(rows=rows, info=info, editable=puede_editar(uid), siguiente=c.next_code(kind, uid))

    @app.post("/api/fun/<kind>")
    @app.put("/api/fun/<kind>/<int:rid>")
    def api_fun_save(kind, rid=None):
        _flat(kind)
        uid = _flat_de_unidad(kind, rid)["unidad_id"] if rid else _uid_param()
        exigir_edicion(uid)
        data = _con_meses(request.get_json(force=True) or {})
        if not s(data.get("partida")):
            raise ValueError("Seleccione la Partida / rubro.")
        if kind == "FI":
            data["becas_pct"] = parse_pct(data.get("becas_pct"))
        new = ctrl().save_flat(kind, uid, data, rid)
        log(f"Guardar {kind}", s(data.get("partida")))
        return jsonify(id=new)

    @app.delete("/api/fun/<kind>/<int:rid>")
    def api_fun_delete(kind, rid):
        _flat(kind)
        r = _flat_de_unidad(kind, rid)
        exigir_edicion(r["unidad_id"])
        ctrl().delete_flat(kind, rid)
        log(f"Eliminar {kind}", r["codigo"])
        return jsonify(ok=True)

    @app.post("/api/fun/<kind>/renumerar")
    def api_fun_renum(kind):
        _flat(kind)
        uid = _uid_param()
        exigir_edicion(uid)
        ctrl().renumerar_flat(kind, uid)
        return jsonify(ok=True)

    # ------------------------------------------------- API: cálculos en vivo
    @app.post("/api/calc/<tipo>")
    def api_calc(tipo):
        c = ctrl()
        p = c.params
        d = _con_meses(request.get_json(force=True) or {})
        out: dict[str, Any] = {}
        if tipo == "detalle":
            t = d.get("tipo") or "Egresos"
            cc = calc_detalle(d, p)
            cat = c.categoria_de(d.get("partida", ""), t)
            out = {"_categoria": cat or ("(partida fuera del catálogo)" if d.get("partida") else ""),
                   "_desc_partida": c.descripcion_partida(d.get("partida", ""), t) if d.get("partida") else "",
                   "_cu": fmt_money(cc["costo_unit_soles"]), "_total": fmt_money(cc["total"]),
                   "_prog": fmt_money(cc["programado"]), "_control": cc["control"], "total": cc["total"]}
        elif tipo == "FI":
            d["becas_pct"] = parse_pct(d.get("becas_pct"))
            cc = calc_fing(d, p)
            out = {"_categoria": c.categoria_de(d.get("partida", ""), "Ingresos"),
                   "_soles": fmt_money(cc["tarifa_soles"]), "_bruto": fmt_money(cc["bruto"]),
                   "_desc": fmt_money(cc["descuentos"]), "_neto": fmt_money(cc["neto"]),
                   "_cobranza": fmt_money(cc["cobranza_total"]), "_siguiente": fmt_money(cc["por_cobrar_siguiente"]),
                   "_incobrable": fmt_money(cc["incobrable"]),
                   "meses_calc": [fmt_money(v) if v else "" for v in cc["meses"]]}
        elif tipo == "FE":
            cc = calc_fegr(d, p)
            out = {"_categoria": c.categoria_de(d.get("partida", ""), "Egresos"),
                   "_soles": fmt_money(cc["costo_unit_soles"]), "_coste": fmt_money(cc["coste_total"]),
                   "_total": fmt_money(cc["total"]), "_control": cc["control"],
                   "meses_calc": [fmt_money(v) if v else "" for v in cc["meses"]]}
        elif tipo == "INI":
            obj = code_display(s(d.get("obj_poa")))
            oe = c.oe_de_obj(obj) if obj else ""
            out = {"_oe": c.oe_display(oe) if oe else ("(objetivo sin OE vinculado)" if obj else ""),
                   "_control_meta": control_meta(dict(d, nombre=d.get("nombre") or "x"))}
        elif tipo == "AP":
            t = to_float(d.get("tarifa"))
            b = parse_pct(d.get("becas_pct")) or 0.0
            out = {"_tarifa_neta": fmt_money(t * (1 - b)) if t is not None else ""}
        else:
            abort(404)
        return jsonify(out)

    # ----------------------------------------------------------- API: resumen
    @app.get("/api/resumen")
    def api_resumen():
        uid = _uid_param()
        summ = ctrl().build_summary(uid)
        u = summ["unidad"]
        secciones = []
        for sec in summ["sections"]:
            secciones.append({k: (sorted(v) if isinstance(v, set) else v) for k, v in sec.items()})
        return jsonify(unidad={**u, "display": ctrl().unidad_display(u)}, anio=summ["anio"], secciones=secciones,
                       directrices=summ["directrices"], resultado=summ["s3"]["resultado"])

    # ---------------------------------------------------- API: administración
    @app.get("/api/admin/parametros")
    def api_param_get():
        _solo_admin()
        p = dict(db().parametros())
        p.pop("clave_admin", None)
        p.update(tema())
        return jsonify(p)

    @app.put("/api/admin/parametros")
    def api_param_put():
        _solo_admin()
        from controllers import PARAM_GENERALES, PARAM_NUM, PARAM_PCT
        d = request.get_json(force=True) or {}
        gen = {}
        for key, _l, _n in PARAM_GENERALES:
            if key not in d:
                continue
            v = d[key]
            gen[key] = parse_pct(v) if key in PARAM_PCT else to_float(v) if key in PARAM_NUM else s(v)
        if "mes_inicio_reajuste" in d:
            gen["mes_inicio_reajuste"] = int(to_float(d["mes_inicio_reajuste"], 1) or 1)
        if not gen.get("tipo_cambio", 1):
            raise ValueError("El tipo de cambio es obligatorio.")
        db().save_parametros(gen)
        for k in TEMA_DEFECTO:
            if k in d:
                v = s(d[k])
                if k != "tema_fuente" and v and not re.fullmatch(r"#[0-9A-Fa-f]{6}", v):
                    raise ValueError(f"Color no válido: {v} (use el formato #RRGGBB).")
                if k == "tema_fuente" and v and not re.fullmatch(r"[A-Za-z0-9 ]{2,40}", v):
                    raise ValueError("Nombre de tipografía no válido.")
                db().set_meta(k, v)
        log("Parámetros", "actualizados")
        return jsonify(ok=True)

    @app.get("/api/admin/t/<entidad>")
    def api_admin_list(entidad):
        if entidad == "unidades" and usuario()["rol"] == "CONSULTA":
            pass
        else:
            _solo_admin()
        return jsonify(rows=_admin_rows(entidad))

    @app.post("/api/admin/t/<entidad>")
    @app.put("/api/admin/t/<entidad>/<int:rid>")
    def api_admin_save(entidad, rid=None):
        _solo_admin()
        d = {k: v for k, v in (request.get_json(force=True) or {}).items() if not k.startswith("_")}
        new = _admin_save(entidad, d, rid)
        log(f"Guardar {entidad}", str(new))
        extra = {}
        if entidad == "usuarios":
            if g.get("clave_temporal"):
                extra["clave_temporal"] = g.clave_temporal
            fila = db().query_one("SELECT usuario, rol FROM usuarios WHERE id=?", (new,))
            extra["usuario"] = fila["usuario"]
            if fila["rol"] == "RESPONSABLE" and not db().scalar(
                    "SELECT COUNT(*) FROM usuario_unidades WHERE usuario_id=?", (new,)):
                extra["aviso"] = "El responsable no tiene unidades asignadas: no podrá formular hasta asignarle alguna."
        return jsonify(id=new, **extra)

    @app.delete("/api/admin/t/<entidad>/<int:rid>")
    def api_admin_delete(entidad, rid):
        _solo_admin()
        _admin_delete(entidad, rid)
        log(f"Eliminar {entidad}", str(rid))
        return jsonify(ok=True)

    @app.get("/api/admin/bitacora")
    def api_bitacora():
        _solo_admin()
        q = f"%{request.args.get('q', '').strip()}%"
        rows = db().query("SELECT fecha, usuario, accion, detalle, ip FROM bitacora WHERE usuario LIKE ? OR accion "
                          "LIKE ? OR detalle LIKE ? ORDER BY id DESC LIMIT 1000", (q, q, q))
        return jsonify(rows=rows)

    @app.get("/api/admin/tablero")
    def api_tablero():
        if usuario()["rol"] not in ("ADMIN", "CONSULTA"):
            raise PermissionError("Solo para la DPDI.")
        c = ctrl()
        responsables = {}
        for r in db().query("SELECT uu.unidad_id, us.nombre, us.usuario FROM usuario_unidades uu JOIN usuarios us "
                            "ON us.id = uu.usuario_id WHERE us.activo = 1"):
            responsables.setdefault(r["unidad_id"], []).append(r["nombre"] or r["usuario"])
        con_datos = {r["unidad_id"] for t in ("iniciativas", "otras_operativas", "act_propias", "inversiones",
                                               "funcionamiento_ing", "funcionamiento_egr")
                     for r in db().query(f"SELECT DISTINCT unidad_id FROM {t} WHERE anio=?", (c.anio,))}
        filas = []
        for u in sorted(c.unidades(), key=lambda x: s(x["codigo_area"])):
            fila = {"id": u["id"], "unidad": c.unidad_display(u), "macrogrupo": u["macrogrupo"],
                    "responsables": ", ".join(responsables.get(u["id"], [])), "estado": u["estado_ficha"] or "Borrador",
                    "estado_fecha": u["estado_fecha"], "techo": 0.0, "egresos": 0.0, "uso": None, "ingresos": 0.0,
                    "capex": 0.0, "techo_capex": u["techo_inversion"] or 0.0, "revisar": 0, "con_datos": False}
            if u["id"] in con_datos:
                summ = c.build_summary(u["id"])
                opex_i, opex_e, opex_t = summ["s1"]["opex"]
                ap_e = summ["s1"]["ap"][1]
                fila.update(techo=opex_t, egresos=opex_e - ap_e, ingresos=opex_i,
                            uso=(opex_e - ap_e) / opex_t if opex_t else None, capex=summ["s1"]["inv"][1],
                            revisar=sum(1 for d_ in summ["directrices"] if d_["resultado"] == "Revisar"),
                            con_datos=True)
            else:
                fila["techo"] = sum((u[k] or 0) for k in ("techo_iniciativas", "techo_funcionamiento",
                                                          "techo_otras_operativas"))
            filas.append(fila)
        return jsonify(rows=filas)

    # -------------------------------------------------- descargas y cargas
    @app.get("/descargas/libro")
    def desc_libro():
        uid = _uid_param()
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "libro.xlsx"
        export_libro_oficial(ctrl(), uid, tmp)
        log("Descargar libro oficial", ctrl().unidad(uid)["codigo_area"])
        return _enviar(tmp, f"Ficha_Formulacion_POA_{ctrl().anio}_{ctrl().unidad(uid)['codigo_area']}.xlsx")

    @app.get("/descargas/pdf")
    def desc_pdf():
        uid = _uid_param()
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "ficha.pdf"
        export_pdf(ctrl(), uid, tmp)
        log("Descargar PDF", ctrl().unidad(uid)["codigo_area"])
        return _enviar(tmp, f"Ficha_Resumen_POA_{ctrl().anio}_{ctrl().unidad(uid)['codigo_area']}.pdf")

    @app.get("/descargas/bd")
    def desc_bd():
        todas = request.args.get("alcance") == "todas"
        if todas and usuario()["rol"] not in ("ADMIN", "CONSULTA"):
            raise PermissionError("Solo la DPDI descarga la base consolidada de todas las unidades.")
        uids = unidades_permitidas() if todas else [_uid_param()]
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "bd.xlsx"
        export_bd_consolidada(ctrl(), uids, tmp, with_csv=False)
        log("Descargar BD NIIF", "todas" if todas else str(uids[0]))
        return _enviar(tmp, f"BD_CONSOLIDADA_NIIF_{ctrl().anio}{'' if todas else '_' + ctrl().unidad(uids[0])['codigo_area']}.xlsx")

    @app.get("/descargas/plantilla-parametros")
    def desc_plantilla():
        _solo_admin()
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "plantilla.xlsx"
        export_plantilla_parametros(ctrl(), tmp)
        log("Descargar plantilla de parámetros")
        return _enviar(tmp, f"Plantilla_Parametros_Catalogos_POA_{ctrl().anio}.xlsx")

    @app.get("/descargas/respaldo")
    def desc_respaldo():
        _solo_admin()
        import sqlite3
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "respaldo.sqlite3"
        destino = sqlite3.connect(tmp)
        db().conn.backup(destino)  # copia consistente aunque haya usuarios trabajando
        destino.close()
        log("Descargar respaldo de la base de datos")
        return _enviar(tmp, f"respaldo_poa_dpdi_{datetime.now():%Y%m%d_%H%M}.sqlite3")

    @app.post("/cargas/ficha")
    def carga_ficha():
        u = usuario()
        if u["rol"] == "CONSULTA":
            raise PermissionError("El rol de consulta no puede importar fichas.")
        destino = request.form.get("destino", "actual")
        uid = unidad_actual()
        if u["rol"] == "RESPONSABLE" or destino == "actual":
            exigir_edicion(uid)
            target = uid
        else:
            target = None
        params = es_admin() and request.form.get("parametros") == "1"
        nombre = secure_filename(request.files["archivo"].filename) if "archivo" in request.files else ""
        rep = _procesar_upload(lambda p: ctrl().import_ficha(p, unidad_id=target, importar_parametros=params),
                               (".xlsx", ".xlsm"))
        if rep["unidad_id"] not in unidades_permitidas():
            raise PermissionError("La ficha corresponde a una unidad no asignada.")
        session["uid"] = rep["unidad_id"]
        log("Importar ficha", f"{nombre} → {ctrl().unidad(rep['unidad_id'])['codigo_area']}")
        return jsonify(rep)

    @app.post("/cargas/plantilla-parametros")
    @app.post("/cargas/parametros-ficha")
    def carga_plantilla():
        """Acepta la plantilla de Parámetros y Catálogos o una Ficha_Formulacion_POA (hoja «Parámetros»):
        el tipo se reconoce por las hojas del archivo, sin importar el botón usado."""
        _solo_admin()
        rep = _procesar_upload(_importar_parametros_auto, (".xlsx", ".xlsm"))
        log("Importar parámetros", rep["Archivo reconocido"])
        return jsonify(rep)

    @app.get("/descargas/plantilla-usuarios")
    def desc_plantilla_usuarios():
        _solo_admin()
        tmp = Path(tempfile.mkdtemp(prefix="poa_")) / "usuarios.xlsx"
        _plantilla_usuarios(tmp)
        log("Descargar plantilla de usuarios")
        return _enviar(tmp, "Plantilla_Usuarios_POA_DPDI.xlsx")

    @app.post("/cargas/usuarios")
    def carga_usuarios():
        _solo_admin()
        rep = _procesar_upload(_importar_usuarios, (".xlsx",))
        log("Importar usuarios", f"{rep['Usuarios creados']} creados, {rep['Usuarios actualizados']} actualizados")
        return jsonify(rep)

    @app.post("/cargas/niif")
    def carga_niif():
        _solo_admin()
        rep = _procesar_upload(ctrl().carga_masiva_niif)
        log("Carga masiva NIIF", str(rep))
        return jsonify(rep)

    @app.post("/cargas/logo")
    def carga_logo():
        _solo_admin()
        from flask import current_app
        f = request.files.get("archivo")
        if not f or not f.filename:
            raise ValueError("Seleccione la imagen del logo.")
        ext = f.filename.rsplit(".", 1)[-1].lower()
        if ext not in ("png", "svg", "jpg", "jpeg", "webp"):
            raise ValueError("Use una imagen PNG, SVG, JPG o WEBP.")
        d = Path(current_app.config["DATA_DIR"]) / "marca"
        d.mkdir(exist_ok=True)
        for old in d.glob("logo.*"):
            old.unlink()
        f.save(d / f"logo.{ext}")
        log("Actualizar logo")
        return jsonify(ok=True)


MANUALES = [
    {"archivo": "Manual_Usuario_POA_DPDI.pdf", "titulo": "Manual del usuario",
     "descripcion": "Para responsables de unidad orgánica: acceso, formulación de cada tipo de actividad, ficha resumen, "
                    "directrices y envío a revisión.", "roles": ("RESPONSABLE", "CONSULTA", "ADMIN")},
    {"archivo": "Manual_Administrador_POA_DPDI.pdf", "titulo": "Manual del administrador DPDI",
     "descripcion": "Configuración del año, parámetros, catálogos, usuarios, revisión de fichas, tablero, reportes y "
                    "respaldos.", "roles": ("ADMIN",)},
    {"archivo": "Manual_Tecnologia_POA_DPDI.pdf", "titulo": "Manual técnico (Tecnologías de la Información)",
     "descripcion": "Arquitectura, instalación en laptop y servidor, servicio de Windows, HTTPS, respaldos, seguridad "
                    "y solución de problemas.", "roles": ("ADMIN",)},
]


def _manuales_visibles() -> list[dict]:
    rol = usuario()["rol"]
    return [m for m in MANUALES if rol in m["roles"]]


def _kind(kind: str) -> None:
    if kind not in KINDS:
        abort(404)


def _flat(kind: str) -> None:
    if kind not in ("FI", "FE"):
        abort(404)


def _solo_admin() -> None:
    if not es_admin():
        raise PermissionError("Solo el Administrador DPDI puede realizar esta acción.")


def _clave_valida(clave: str) -> bool:
    return len(clave) >= 8 and re.search(r"[A-Za-z]", clave) is not None and re.search(r"\d", clave) is not None


# =============================================================================
# Administración genérica
# =============================================================================
def _admin_rows(entidad: str) -> list[dict]:
    c = ctrl()
    if entidad == "unidades":
        return db().all("unidades_organicas", order="codigo_area")
    if entidad == "cuentas":
        return db().all("catalogo_niif_partidas", order="tipo_flujo, partida_gg='', partida_gg COLLATE NOCASE, "
                                                        "partida_rubro")
    if entidad == "catalogos":
        return db().catalogo(request.args.get("tipo", "EJE"))
    if entidad == "costos":
        return db().all("costos_referencia", order="id")
    if entidad == "morosidad":
        from controllers import rango_dias
        rows = db().rangos_morosidad()
        for r in rows:
            r["_dias"] = rango_dias(r)
            r["_meses"] = "—" if r["tratamiento"] == "Provisión" else \
                f"+{r['meses_desfase']} {'mes' if r['meses_desfase'] == 1 else 'meses'}"
        return rows
    if entidad == "plantilla":
        areas = {u["id"]: u["codigo_area"] for u in c.unidades()}
        rows = db().all("plantilla_funcionamiento", order="unidad_id IS NOT NULL, unidad_id, orden, id")
        for r in rows:
            r["ambito"] = areas.get(r["unidad_id"], "GLOBAL")
        return rows
    if entidad == "usuarios":
        rows = db().query("SELECT id, usuario, nombre, correo, rol, activo, debe_cambiar_clave, ultimo_acceso "
                          "FROM usuarios ORDER BY usuario")
        asign = {}
        for r in db().query("SELECT uu.usuario_id, u.id, u.codigo_area FROM usuario_unidades uu JOIN "
                            "unidades_organicas u ON u.id = uu.unidad_id ORDER BY u.codigo_area"):
            asign.setdefault(r["usuario_id"], []).append(r)
        for r in rows:
            r["unidades"] = [x["id"] for x in asign.get(r["id"], [])]
            r["_unidades"] = ", ".join(x["codigo_area"] for x in asign.get(r["id"], [])) or \
                ("todas" if r["rol"] != "RESPONSABLE" else "—")
        return rows
    abort(404)


def _admin_save(entidad: str, d: dict, rid: Optional[int]) -> int:
    c = ctrl()
    if entidad == "unidades":
        return c.save_unidad(d, rid)
    if entidad == "cuentas":
        d["activo"] = 1 if d.get("activo") in (1, True, "1", "on") else 0
        return c.save_partida_niif(d, rid)
    if entidad == "catalogos":
        tipo = request.args.get("tipo", "EJE")
        data = dict(d, orden=int(to_float(d.get("orden"), 0) or 0))
        if tipo == "OBJ":
            if not s(data.get("eje")) or not s(data.get("padre")):
                raise ValueError("Cada objetivo anual del POA debe tener su Eje estratégico y su OE del PEI.")
            data["eje"], data["padre"] = code_display(s(data["eje"])), code_display(s(data["padre"]))
        else:
            data["eje"], data["padre"] = "", ""
        return c.save_catalogo(tipo, data, rid)
    if entidad == "morosidad":
        d["activo"] = 1 if d.get("activo") in (1, True, "1", "on") else 0
        return c.save_rango_morosidad(d, rid)
    if entidad == "costos":
        if not s(d.get("concepto")):
            raise ValueError("El concepto es obligatorio.")
        d["costo_unitario"] = to_float(d.get("costo_unitario"))
        return db().upsert("costos_referencia", d, rid)
    if entidad == "plantilla":
        if not s(d.get("partida")):
            raise ValueError("Seleccione la cuenta NIIF.")
        amb = s(d.pop("ambito", "GLOBAL")) or "GLOBAL"
        uid = None
        if amb != "GLOBAL":
            r = db().query_one("SELECT id FROM unidades_organicas WHERE codigo_area=?", (amb,))
            if not r:
                raise ValueError(f"El área {amb} no existe.")
            uid = r["id"]
        d.update(unidad_id=uid, porcentaje=to_float(d.get("porcentaje"), 1.0), monto_fijo=to_float(d.get("monto_fijo"), 0.0),
                 orden=int(to_float(d.get("orden"), 0) or 0),
                 activo=1 if d.get("activo") in (1, True, "1", "on") else 0,
                 solo_genera_ingresos=1 if d.get("solo_genera_ingresos") in (1, True, "1", "on") else 0)
        return db().upsert("plantilla_funcionamiento", d, rid)
    if entidad == "usuarios":
        return _guardar_usuario(d, rid)
    abort(404)


def normalizar_usuario(texto: Any) -> str:
    """«Pedro Ruíz » → «pedro.ruiz»: minúsculas, sin tildes y con puntos en lugar de espacios."""
    import unicodedata
    v = unicodedata.normalize("NFKD", s(texto)).encode("ascii", "ignore").decode().lower().strip()
    return re.sub(r"\s+", ".", v)


def _clave_temporal() -> str:
    return secrets.token_urlsafe(7).replace("-", "x").replace("_", "y") + str(secrets.randbelow(10))


def _guardar_usuario(d: dict, rid: Optional[int]) -> int:
    """Crea o actualiza un usuario. Si es nuevo y no se indicó clave, genera una temporal (g.clave_temporal)."""
    nombre_u = normalizar_usuario(d.get("usuario"))
    if not re.fullmatch(r"[a-z0-9._@\-]{3,80}", nombre_u):
        raise ValueError("Usuario no válido: use de 3 a 80 caracteres (letras, números, punto, guion, @).")
    rol = d.get("rol") if d.get("rol") in ROLES else "RESPONSABLE"
    otro = db().query_one("SELECT id FROM usuarios WHERE usuario=?", (nombre_u,))
    if otro and otro["id"] != rid:
        raise ValueError(f"El usuario {nombre_u} ya existe.")
    data = {"usuario": nombre_u, "nombre": s(d.get("nombre")), "correo": s(d.get("correo")), "rol": rol,
            "activo": 1 if d.get("activo") in (1, True, "1", "on") else 0,
            "debe_cambiar_clave": 1 if d.get("debe_cambiar_clave") in (1, True, "1", "on") else 0}
    yo = usuario()
    if rid == yo["id"] and (rol != "ADMIN" or not data["activo"]):
        raise ValueError("No puede quitarse a sí mismo el rol de administrador ni desactivarse.")
    clave = s(d.get("nueva_clave"))
    if not clave and not rid:
        clave = _clave_temporal()
        g.clave_temporal = clave
        data["debe_cambiar_clave"] = 1
    if clave:
        if not _clave_valida(clave):
            raise ValueError("La clave debe tener al menos 8 caracteres, con letras y números.")
        data["password_hash"] = generate_password_hash(clave)
        data.update(intentos_fallidos=0, bloqueado_hasta="")  # restablecer la clave también desbloquea
    if rid:
        db().update("usuarios", rid, data)
    else:
        rid = db().insert("usuarios", data)
    unidades = [int(x) for x in (d.get("unidades") or []) if str(x).isdigit()]
    db().execute("DELETE FROM usuario_unidades WHERE usuario_id=?", (rid,))
    for uid in unidades:
        db().execute("INSERT OR IGNORE INTO usuario_unidades (usuario_id, unidad_id) VALUES (?,?)", (rid, uid))
    db().conn.commit()
    return rid


# ----------------------------------------------------------- carga masiva de usuarios
COLS_USUARIOS = [("usuario", "Usuario", 22), ("nombre", "Nombre completo", 34), ("correo", "Correo", 32),
                 ("rol", "Rol (RESPONSABLE / CONSULTA / ADMIN)", 22),
                 ("unidades", "Códigos de área asignados (separados por coma)", 40),
                 ("clave", "Clave temporal (opcional)", 22), ("activo", "Activo (Sí/No)", 12)]


def _codigo_texto(v: Any) -> str:
    """Los códigos de área numéricos llegan de Excel como 10110101.0."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return s(v)


def _plantilla_usuarios(path: Path) -> None:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    c = ctrl()
    wb = openpyxl.Workbook()
    ins = wb.active
    ins.title = "Instrucciones"
    for fila in (["Carga masiva de usuarios - Sistema POA DPDI"], [],
                 ["1. Complete la hoja «Usuarios»: una fila por persona. Las filas existentes se actualizan por su usuario."],
                 ["2. Usuario: letras, números, punto, guion o @ (por ejemplo ana.alva). Los espacios se cambian por puntos."],
                 ["3. Rol: RESPONSABLE (formula sus unidades), CONSULTA (solo lectura) o ADMIN (DPDI)."],
                 ["4. Códigos de área: los de la hoja «Unidades», separados por coma. Obligatorio para RESPONSABLE."],
                 ["5. Clave temporal: si la deja vacía en un usuario nuevo, el sistema genera una. En usuarios existentes,"
                  " vacía = no cambiar la clave. Todos deben cambiarla en su primer ingreso."],
                 ["6. La hoja «Sugeridos» propone usuarios a partir del responsable y correo de cada unidad: copie a"
                  " «Usuarios» solo las filas que correspondan."],
                 ["7. Importe el archivo en Administración → Usuarios y accesos → Importar usuarios."
                  " Al terminar podrá descargar la lista de claves temporales para entregarlas."]):
        ins.append(fila)
    ins["A1"].font = Font(bold=True, size=14, color="FF6600")
    ins.column_dimensions["A"].width = 120

    cab_fill, cab_font = PatternFill("solid", fgColor="FF6600"), Font(bold=True, color="FFFFFF")

    def hoja(titulo, filas):
        ws = wb.create_sheet(titulo)
        ws.append([h for _k, h, _w in COLS_USUARIOS])
        for i, (_k, _h, w) in enumerate(COLS_USUARIOS, start=1):
            ws.cell(1, i).fill, ws.cell(1, i).font = cab_fill, cab_font
            ws.cell(1, i).alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "A2"
        for f in filas:
            ws.append(f)
        dv_rol = DataValidation(type="list", formula1='"RESPONSABLE,CONSULTA,ADMIN"', allow_blank=True)
        dv_act = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
        ws.add_data_validation(dv_rol)
        ws.add_data_validation(dv_act)
        dv_rol.add("D2:D2000")
        dv_act.add("G2:G2000")
        return ws

    asign: dict[int, list[str]] = {}
    for r in db().query("SELECT uu.usuario_id, u.codigo_area FROM usuario_unidades uu "
                        "JOIN unidades_organicas u ON u.id = uu.unidad_id ORDER BY u.codigo_area"):
        asign.setdefault(r["usuario_id"], []).append(r["codigo_area"])
    existentes = db().query("SELECT * FROM usuarios ORDER BY usuario")
    hoja("Usuarios", [[u["usuario"], u["nombre"], u["correo"], u["rol"], ", ".join(asign.get(u["id"], [])), "",
                       "Sí" if u["activo"] else "No"] for u in existentes])
    ya = {u["usuario"] for u in existentes} | {s(u["correo"]).lower() for u in existentes}
    sugeridos: dict[str, list] = {}
    for un in c.unidades():
        correo = s(un.get("correo")).lower()
        if "@" not in correo or correo in ya:
            continue
        usuario_s = normalizar_usuario(correo.split("@")[0])
        if usuario_s in ya:
            continue
        fila = sugeridos.setdefault(usuario_s, [usuario_s, s(un.get("responsable")), correo, "RESPONSABLE", [], "", "Sí"])
        fila[4].append(un["codigo_area"])
    ws_s = hoja("Sugeridos", [f[:4] + [", ".join(f[4])] + f[5:] for f in sugeridos.values()])
    if not sugeridos:
        ws_s["A2"] = ("Sin sugerencias: registre «Responsable» y «Correo electrónico» de cada unidad (Unidades y techos o "
                      "plantilla de parámetros) para que el sistema proponga los usuarios.")
        ws_s["A2"].font = Font(italic=True, color="6B6B6B")
    wu = wb.create_sheet("Unidades")
    wu.append(["Código de área", "Unidad (código - centro de costo - unidad orgánica)", "Macrogrupo"])
    for i in (1, 2, 3):
        wu.cell(1, i).fill, wu.cell(1, i).font = cab_fill, cab_font
    for un in c.unidades():
        wu.append([un["codigo_area"], c.unidad_display(un), un.get("macrogrupo")])
    wu.column_dimensions["A"].width, wu.column_dimensions["B"].width, wu.column_dimensions["C"].width = 16, 90, 34
    wb.active = 1
    wb.save(path)


def _importar_usuarios(path: Path) -> dict:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "Usuarios" not in wb.sheetnames:
        raise ValueError("El archivo no contiene la hoja «Usuarios». Descargue la plantilla de usuarios del sistema.")
    filas = list(wb["Usuarios"].iter_rows(values_only=True))
    wb.close()
    codigos = {s(u["codigo_area"]).upper(): u["id"] for u in ctrl().unidades()}
    rep: dict[str, Any] = {"Usuarios creados": 0, "Usuarios actualizados": 0, "avisos": [], "credenciales": []}
    for n, fila in enumerate(filas[1:], start=2):
        fila = list(fila) + [None] * (len(COLS_USUARIOS) - len(fila))
        d = {k: fila[i] for i, (k, _h, _w) in enumerate(COLS_USUARIOS)}
        if not any(s(v) for v in d.values()):
            continue
        nombre_u = normalizar_usuario(d["usuario"])
        uids, faltan = [], []
        for cod in re.split(r"[,;\n]+", _codigo_texto(d["unidades"])):
            cod = cod.strip().upper()
            if cod:
                (uids.append(codigos[cod]) if cod in codigos else faltan.append(cod))
        if faltan:
            rep["avisos"].append(f"Fila {n} ({nombre_u}): códigos de área inexistentes: {', '.join(faltan)}.")
        rol = s(d["rol"]).upper() or "RESPONSABLE"
        if rol not in ROLES:
            rep["avisos"].append(f"Fila {n} ({nombre_u}): rol «{d['rol']}» no válido; se omitió la fila.")
            continue
        existente = db().query_one("SELECT id FROM usuarios WHERE usuario=?", (nombre_u,))
        clave = s(d["clave"])
        if isinstance(d["clave"], float) and d["clave"].is_integer():
            clave = str(int(d["clave"]))
        datos = {"usuario": nombre_u, "nombre": d["nombre"], "correo": d["correo"], "rol": rol,
                 "activo": 0 if s(d["activo"]).lower() in ("no", "0", "n") else 1,
                 "debe_cambiar_clave": 1, "nueva_clave": clave, "unidades": [str(x) for x in uids]}
        if existente and not clave:
            datos["debe_cambiar_clave"] = db().scalar("SELECT debe_cambiar_clave FROM usuarios WHERE id=?",
                                                      (existente["id"],))
        g.pop("clave_temporal", None)
        try:
            _guardar_usuario(datos, existente["id"] if existente else None)
        except ValueError as e:
            rep["avisos"].append(f"Fila {n} ({nombre_u or 'sin usuario'}): {e}")
            continue
        rep["Usuarios actualizados" if existente else "Usuarios creados"] += 1
        if rol == "RESPONSABLE" and not uids:
            rep["avisos"].append(f"Fila {n} ({nombre_u}): responsable sin unidades asignadas.")
        if clave or g.get("clave_temporal"):
            rep["credenciales"].append({"usuario": nombre_u, "nombre": s(d["nombre"]), "correo": s(d["correo"]),
                                        "clave": clave or g.clave_temporal})
    g.pop("clave_temporal", None)
    return rep


def _admin_delete(entidad: str, rid: int) -> None:
    c = ctrl()
    if entidad == "unidades":
        c.delete_unidad(rid)
    elif entidad == "cuentas":
        c.delete_partida_niif(rid)
    elif entidad == "catalogos":
        db().delete("catalogos_estrategicos", rid)
    elif entidad == "costos":
        db().delete("costos_referencia", rid)
    elif entidad == "morosidad":
        c.delete_rango_morosidad(rid)
    elif entidad == "plantilla":
        db().delete("plantilla_funcionamiento", rid)
    elif entidad == "usuarios":
        if rid == usuario()["id"]:
            raise ValueError("No puede eliminar su propio usuario.")
        admins = db().scalar("SELECT COUNT(*) FROM usuarios WHERE rol='ADMIN' AND activo=1 AND id<>?", (rid,))
        if db().get("usuarios", rid) and db().get("usuarios", rid)["rol"] == "ADMIN" and not admins:
            raise ValueError("Debe quedar al menos un administrador activo.")
        db().delete("usuarios", rid)
    else:
        abort(404)
