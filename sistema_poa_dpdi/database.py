"""
database.py
===========
Capa de persistencia SQLite del Sistema de Formulación Presupuestal y Operativa
POA - DPDI (Proceso Gestión Estratégica GD-DO-01 / Insumo PP-PR-02).

Define el esquema relacional completo (catálogos NIIF, catálogos estratégicos,
unidades orgánicas con techos, parámetros globales y las tablas maestro-detalle
de las cinco clases de actividad del formato oficial Ficha_Formulacion_POA.xlsx),
crea la base de datos al primer arranque y la pre-puebla con los catálogos
oficiales definidos en catalogs_seed.py.
"""
from __future__ import annotations

import os
import sqlite3
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Optional

APP_NAME = "Sistema_POA_DPDI"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]
MCOLS = [f"m{i:02d}" for i in range(1, 13)]


def app_data_dir() -> Path:
    """Carpeta persistente de datos (junto al .exe o en el perfil del usuario)."""
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA") or Path.home()) / APP_NAME
    else:
        base = Path(__file__).resolve().parent / "data"
    base.mkdir(parents=True, exist_ok=True)
    return base


def default_db_path() -> Path:
    return app_data_dir() / "poa_dpdi.sqlite3"


_MONTH_REAL = ",\n    ".join(f"{c} REAL" for c in MCOLS)
_MONTH_TEXT = ",\n    ".join(f"{c} TEXT" for c in MCOLS)


def _detalle_ddl(table: str, master: str) -> str:
    return f"""
CREATE TABLE IF NOT EXISTS {table} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    master_id INTEGER NOT NULL REFERENCES {master}(id) ON DELETE CASCADE,
    tipo TEXT NOT NULL DEFAULT 'Egresos',
    partida TEXT DEFAULT '',
    partida_gg TEXT DEFAULT '',
    categoria TEXT DEFAULT '',
    descripcion TEXT DEFAULT '',
    cantidad REAL,
    moneda TEXT DEFAULT 'S/',
    costo_unit_moneda REAL,
    {_MONTH_REAL},
    sustento TEXT DEFAULT '',
    observaciones TEXT DEFAULT '',
    orden INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS ix_{table}_master ON {table}(master_id);
"""


SCHEMA = f"""
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS parametros_globales (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    institucion TEXT DEFAULT 'UNIVERSIDAD AUTÓNOMA DEL PERÚ',
    codigo_formato TEXT DEFAULT 'GD-DO-01 / PP-PR-02',
    version_formato TEXT DEFAULT '1.0',
    fecha_aprobacion TEXT DEFAULT '',
    anio_formulacion TEXT DEFAULT '2027',
    fecha_emision TEXT DEFAULT '2026-10-30',
    fecha_limite TEXT DEFAULT '2026-11-20',
    inflacion_proyectada REAL DEFAULT 0.025,
    tipo_cambio REAL DEFAULT 3.50,
    valor_uit REAL,
    rmv REAL,
    reajuste_remunerativo_pct REAL DEFAULT 0,
    mes_inicio_reajuste INTEGER DEFAULT 1,
    tasa_morosidad_pct REAL DEFAULT 0,
    estudiantes_pregrado REAL,
    estudiantes_posgrado REAL,
    pension_ref_pregrado REAL,
    pension_ref_posgrado REAL,
    cuotas_periodo REAL,
    director_dpdi TEXT DEFAULT '',
    cargo_dpdi TEXT DEFAULT 'Director de Planificación y Desarrollo Institucional',
    clave_admin TEXT DEFAULT 'DPDI'
);

CREATE TABLE IF NOT EXISTS unidades_organicas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo_area TEXT NOT NULL UNIQUE,
    macrogrupo TEXT NOT NULL DEFAULT 'Soporte Académico/Administrativo',
    nombre_unidad TEXT NOT NULL,
    centro_costo TEXT DEFAULT '',
    responsable TEXT DEFAULT '',
    cargo TEXT DEFAULT '',
    correo TEXT DEFAULT '',
    genera_ingresos INTEGER NOT NULL DEFAULT 0,
    techo_iniciativas REAL DEFAULT 0,
    techo_funcionamiento REAL DEFAULT 0,
    techo_otras_operativas REAL DEFAULT 0,
    techo_inversion REAL DEFAULT 0,
    techo_planilla_docente REAL DEFAULT 0,
    techo_planilla_adm REAL DEFAULT 0,
    techo_marketing REAL DEFAULT 0,
    meta_ingresos_func REAL DEFAULT 0,
    meta_resultado_otras REAL DEFAULT 0,
    meta_resultado_propias REAL DEFAULT 0,
    fecha_presentacion TEXT DEFAULT '',
    version_ficha TEXT DEFAULT '',
    elaborado_nombre TEXT DEFAULT '',
    elaborado_cargo TEXT DEFAULT '',
    elaborado_fecha TEXT DEFAULT '',
    revisado_nombre TEXT DEFAULT '',
    revisado_cargo TEXT DEFAULT '',
    revisado_fecha TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS catalogo_niif_partidas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo_niif TEXT NOT NULL,
    partida_rubro TEXT NOT NULL,
    partida_gg TEXT DEFAULT '',
    descripcion TEXT DEFAULT '',
    categoria TEXT DEFAULT '',
    tipo_flujo TEXT NOT NULL CHECK (tipo_flujo IN ('Ingresos','Egresos','Inversion')),
    naturaleza TEXT DEFAULT '',
    resultado_funcion TEXT DEFAULT '',
    activo INTEGER DEFAULT 1,
    UNIQUE (codigo_niif, partida_rubro, tipo_flujo)
);
CREATE INDEX IF NOT EXISTS ix_niif_tipo ON catalogo_niif_partidas(tipo_flujo);
CREATE INDEX IF NOT EXISTS ix_niif_gg ON catalogo_niif_partidas(partida_gg);

CREATE TABLE IF NOT EXISTS catalogos_estrategicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_catalogo TEXT NOT NULL,
    codigo TEXT NOT NULL,
    descripcion TEXT DEFAULT '',
    padre TEXT DEFAULT '',
    eje TEXT DEFAULT '',
    extra TEXT DEFAULT '',
    orden INTEGER DEFAULT 0,
    UNIQUE (tipo_catalogo, codigo)
);

CREATE TABLE IF NOT EXISTS costos_referencia (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    concepto TEXT NOT NULL UNIQUE,
    unidad_medida TEXT DEFAULT '',
    costo_unitario REAL,
    fuente TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS plantilla_funcionamiento (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    macrogrupo TEXT DEFAULT '*',
    solo_genera_ingresos INTEGER DEFAULT 0,
    tipo_flujo TEXT NOT NULL CHECK (tipo_flujo IN ('Ingresos','Egresos')),
    partida TEXT NOT NULL,
    base_calculo TEXT NOT NULL DEFAULT 'MONTO_FIJO',
    porcentaje REAL DEFAULT 1,
    monto_fijo REAL DEFAULT 0,
    aplica_ajuste TEXT DEFAULT 'No',
    sustento TEXT DEFAULT 'Histórico de ejecución',
    activo INTEGER DEFAULT 1,
    orden INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS iniciativas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    eje TEXT DEFAULT '',
    nombre TEXT DEFAULT '',
    descripcion TEXT DEFAULT '',
    origen TEXT DEFAULT '',
    directriz TEXT DEFAULT '',
    obj_poa TEXT DEFAULT '',
    oe TEXT DEFAULT '',
    ae TEXT DEFAULT '',
    indicador TEXT DEFAULT '',
    unidad_medida TEXT DEFAULT '',
    linea_base TEXT DEFAULT '',
    meta_anual REAL,
    t1 REAL,
    t2 REAL,
    t3 REAL,
    responsable TEXT DEFAULT '',
    prioridad INTEGER,
    observaciones TEXT DEFAULT '',
    UNIQUE (unidad_id, anio, codigo)
);

CREATE TABLE IF NOT EXISTS otras_operativas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    categoria TEXT DEFAULT '',
    nombre TEXT DEFAULT '',
    descripcion TEXT DEFAULT '',
    responsable TEXT DEFAULT '',
    observaciones TEXT DEFAULT '',
    UNIQUE (unidad_id, anio, codigo)
);

CREATE TABLE IF NOT EXISTS act_propias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    tipo TEXT DEFAULT '',
    nombre TEXT DEFAULT '',
    publico_objetivo TEXT DEFAULT '',
    mes_ejecucion TEXT DEFAULT '',
    participantes REAL,
    tarifa REAL,
    becas_pct REAL,
    responsable TEXT DEFAULT '',
    sustento TEXT DEFAULT '',
    observaciones TEXT DEFAULT '',
    UNIQUE (unidad_id, anio, codigo)
);

CREATE TABLE IF NOT EXISTS inversiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    tipo_inversion TEXT DEFAULT '',
    proyecto TEXT DEFAULT '',
    prioridad INTEGER,
    ficha_sustento TEXT DEFAULT 'No',
    observaciones TEXT DEFAULT '',
    UNIQUE (unidad_id, anio, codigo)
);

CREATE TABLE IF NOT EXISTS funcionamiento_ing (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    partida TEXT DEFAULT '',
    categoria TEXT DEFAULT '',
    cantidad REAL,
    moneda TEXT DEFAULT 'S/',
    tarifa_moneda REAL,
    becas_pct REAL,
    aplica_morosidad TEXT DEFAULT 'No',
    {_MONTH_TEXT},
    sustento TEXT DEFAULT '',
    observaciones TEXT DEFAULT '',
    origen_plantilla INTEGER DEFAULT 0,
    orden INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS funcionamiento_egr (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    anio TEXT NOT NULL,
    codigo TEXT NOT NULL,
    partida TEXT DEFAULT '',
    categoria TEXT DEFAULT '',
    cantidad REAL,
    moneda TEXT DEFAULT 'S/',
    costo_unit_moneda REAL,
    aplica_reajuste TEXT DEFAULT 'No',
    {_MONTH_TEXT},
    sustento TEXT DEFAULT '',
    observaciones TEXT DEFAULT '',
    origen_plantilla INTEGER DEFAULT 0,
    orden INTEGER DEFAULT 0
);

{_detalle_ddl('detalle_iniciativas', 'iniciativas')}
{_detalle_ddl('detalle_otras_operativas', 'otras_operativas')}
{_detalle_ddl('detalle_act_prop', 'act_propias')}
{_detalle_ddl('detalle_inversiones', 'inversiones')}

CREATE TABLE IF NOT EXISTS rangos_morosidad (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT NOT NULL UNIQUE,
    nombre TEXT NOT NULL,
    dias_desde INTEGER NOT NULL DEFAULT 0,
    dias_hasta INTEGER,                         -- NULL = sin límite superior
    tasa_pct REAL DEFAULT 0,                    -- % del ingreso devengado que cae en este rango
    tratamiento TEXT DEFAULT 'Desfase de cobro' CHECK (tratamiento IN ('Desfase de cobro', 'Provisión')),
    meses_desfase INTEGER DEFAULT 0,            -- meses en que se recupera (solo «Desfase de cobro»)
    descripcion TEXT DEFAULT '',
    orden INTEGER DEFAULT 0,
    activo INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS app_meta (
    clave TEXT PRIMARY KEY,
    valor TEXT
);

CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario TEXT NOT NULL UNIQUE COLLATE NOCASE,
    nombre TEXT DEFAULT '',
    correo TEXT DEFAULT '',
    rol TEXT NOT NULL DEFAULT 'RESPONSABLE' CHECK (rol IN ('ADMIN', 'CONSULTA', 'RESPONSABLE')),
    password_hash TEXT NOT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
    intentos_fallidos INTEGER NOT NULL DEFAULT 0,
    bloqueado_hasta TEXT DEFAULT '',
    ultimo_acceso TEXT DEFAULT '',
    creado TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS usuario_unidades (
    usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
    unidad_id INTEGER NOT NULL REFERENCES unidades_organicas(id) ON DELETE CASCADE,
    PRIMARY KEY (usuario_id, unidad_id)
);

CREATE TABLE IF NOT EXISTS bitacora (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT DEFAULT (datetime('now','localtime')),
    usuario TEXT DEFAULT '',
    accion TEXT DEFAULT '',
    detalle TEXT DEFAULT ''
);
"""

# Columnas admitidas por tabla (para validar los diccionarios de entrada).
DETALLE_TABLES = {
    "INI": ("iniciativas", "detalle_iniciativas"),
    "OO": ("otras_operativas", "detalle_otras_operativas"),
    "AP": ("act_propias", "detalle_act_prop"),
    "INV": ("inversiones", "detalle_inversiones"),
}


class Database:
    """Envoltura ligera de sqlite3 con utilidades CRUD genéricas."""

    def __init__(self, path: Optional[os.PathLike | str] = None, seed: bool = True, init: bool = True):
        """init=False abre una conexión a una base ya inicializada (uso por petición en el servidor web)."""
        self.path = str(path or default_db_path())
        self.conn = sqlite3.connect(self.path, timeout=30)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.execute("PRAGMA busy_timeout = 30000")
        self._columns_cache: dict[str, list[str]] = {}
        if not init:
            return
        try:
            self.conn.execute("PRAGMA journal_mode = WAL")  # lectores concurrentes + un escritor (multiusuario)
        except sqlite3.DatabaseError:
            pass
        self.create_schema()
        if seed:
            from catalogs_seed import seed_database
            seed_database(self)

    # ------------------------------------------------------------------ schema
    # Columnas añadidas en versiones posteriores (migración de bases existentes).
    MIGRATIONS = [
        ("catalogo_niif_partidas", "descripcion", "TEXT DEFAULT ''"),
        ("catalogos_estrategicos", "eje", "TEXT DEFAULT ''"),
        # Versión web: flujo de revisión de la ficha y trazabilidad
        ("unidades_organicas", "estado_ficha", "TEXT DEFAULT 'Borrador'"),
        ("unidades_organicas", "estado_obs", "TEXT DEFAULT ''"),
        ("unidades_organicas", "estado_fecha", "TEXT DEFAULT ''"),
        ("unidades_organicas", "estado_usuario", "TEXT DEFAULT ''"),
        ("bitacora", "ip", "TEXT DEFAULT ''"),
    ]

    def create_schema(self) -> None:
        self.conn.executescript(SCHEMA)
        for table, column, ddl in self.MIGRATIONS:
            cols = [r[1] for r in self.conn.execute(f"PRAGMA table_info({table})").fetchall()]
            if column not in cols:
                self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")
        self.conn.commit()
        self._columns_cache.clear()

    def meta(self, clave: str) -> Optional[str]:
        r = self.conn.execute("SELECT valor FROM app_meta WHERE clave = ?", (clave,)).fetchone()
        return r[0] if r else None

    def set_meta(self, clave: str, valor: str) -> None:
        self.conn.execute("INSERT OR REPLACE INTO app_meta (clave, valor) VALUES (?, ?)", (clave, valor))
        self.conn.commit()

    def table_exists(self, table: str) -> bool:
        return bool(self.conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone())

    def columns(self, table: str) -> list[str]:
        if table not in self._columns_cache:
            cur = self.conn.execute(f"PRAGMA table_info({table})")
            self._columns_cache[table] = [r["name"] for r in cur.fetchall()]
        return self._columns_cache[table]

    # ---------------------------------------------------------------- helpers
    @contextmanager
    def transaction(self):
        try:
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def execute(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Cursor:
        return self.conn.execute(sql, tuple(params))

    def query(self, sql: str, params: Iterable[Any] = ()) -> list[dict]:
        return [dict(r) for r in self.conn.execute(sql, tuple(params)).fetchall()]

    def query_one(self, sql: str, params: Iterable[Any] = ()) -> Optional[dict]:
        r = self.conn.execute(sql, tuple(params)).fetchone()
        return dict(r) if r else None

    def scalar(self, sql: str, params: Iterable[Any] = ()) -> Any:
        r = self.conn.execute(sql, tuple(params)).fetchone()
        return r[0] if r else None

    def _clean(self, table: str, data: dict) -> dict:
        cols = set(self.columns(table))
        return {k: v for k, v in data.items() if k in cols and k != "id"}

    def insert(self, table: str, data: dict, commit: bool = True) -> int:
        data = self._clean(table, data)
        keys = ", ".join(data)
        marks = ", ".join("?" for _ in data)
        cur = self.conn.execute(f"INSERT INTO {table} ({keys}) VALUES ({marks})", tuple(data.values()))
        if commit:
            self.conn.commit()
        return int(cur.lastrowid)

    def update(self, table: str, row_id: int, data: dict, commit: bool = True) -> None:
        data = self._clean(table, data)
        if not data:
            return
        sets = ", ".join(f"{k} = ?" for k in data)
        self.conn.execute(f"UPDATE {table} SET {sets} WHERE id = ?", (*data.values(), row_id))
        if commit:
            self.conn.commit()

    def upsert(self, table: str, data: dict, row_id: Optional[int]) -> int:
        if row_id:
            self.update(table, row_id, data)
            return int(row_id)
        return self.insert(table, data)

    def delete(self, table: str, row_id: int, commit: bool = True) -> None:
        self.conn.execute(f"DELETE FROM {table} WHERE id = ?", (row_id,))
        if commit:
            self.conn.commit()

    def get(self, table: str, row_id: int) -> Optional[dict]:
        return self.query_one(f"SELECT * FROM {table} WHERE id = ?", (row_id,))

    def all(self, table: str, where: str = "", params: Iterable[Any] = (), order: str = "id") -> list[dict]:
        sql = f"SELECT * FROM {table}"
        if where:
            sql += f" WHERE {where}"
        if order:
            sql += f" ORDER BY {order}"
        return self.query(sql, params)

    def log(self, usuario: str, accion: str, detalle: str = "", ip: str = "") -> None:
        self.conn.execute("INSERT INTO bitacora (usuario, accion, detalle, ip) VALUES (?,?,?,?)",
                          (usuario, accion, detalle, ip))
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()

    # ------------------------------------------------------------ parameters
    def parametros(self) -> dict:
        row = self.query_one("SELECT * FROM parametros_globales WHERE id = 1")
        if row is None:
            self.conn.execute("INSERT INTO parametros_globales (id) VALUES (1)")
            self.conn.commit()
            row = self.query_one("SELECT * FROM parametros_globales WHERE id = 1")
        return row

    def save_parametros(self, data: dict) -> None:
        self.parametros()
        self.update("parametros_globales", 1, data)
        if "tasa_morosidad_pct" in data:
            self.morosidad_desde_parametro(data["tasa_morosidad_pct"])

    # ------------------------------------------------- morosidad por rangos
    def rangos_morosidad(self, solo_activos: bool = False) -> list[dict]:
        if not self.table_exists("rangos_morosidad"):
            return []
        filtro = " WHERE activo = 1" if solo_activos else ""
        return self.query(f"SELECT * FROM rangos_morosidad{filtro} ORDER BY dias_desde, orden")

    def tasa_provision(self) -> float:
        return sum((r["tasa_pct"] or 0.0) for r in self.rangos_morosidad(True) if r["tratamiento"] == "Provisión")

    def morosidad_desde_rangos(self) -> None:
        """El parámetro «Tasa de morosidad» refleja la suma de los rangos tratados como provisión."""
        self.conn.execute("UPDATE parametros_globales SET tasa_morosidad_pct=? WHERE id=1", (self.tasa_provision(),))
        self.conn.commit()

    def morosidad_desde_parametro(self, valor) -> None:
        """Si se edita la tasa única (parámetros, plantilla o Ficha Excel) y hay un solo rango de provisión,
        ese rango toma el nuevo valor; luego el parámetro se recalcula desde los rangos."""
        try:
            valor = float(valor) if valor not in (None, "") else None
        except (TypeError, ValueError):
            valor = None
        prov = [r for r in self.rangos_morosidad(True) if r["tratamiento"] == "Provisión"]
        if valor is not None and len(prov) == 1 and abs((prov[0]["tasa_pct"] or 0.0) - valor) > 1e-12:
            self.conn.execute("UPDATE rangos_morosidad SET tasa_pct=? WHERE id=?", (valor, prov[0]["id"]))
        if prov:
            self.morosidad_desde_rangos()

    def catalogo(self, tipo_catalogo: str) -> list[dict]:
        return self.all("catalogos_estrategicos", "tipo_catalogo = ?", (tipo_catalogo,), "orden, codigo")
