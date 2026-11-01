"""
controllers.py
==============
Reglas de negocio del Sistema de Formulación Presupuestal y Operativa POA - DPDI.

Replica las fórmulas del formato oficial Ficha_Formulacion_POA.xlsx:
  * cálculos por fila (conversión US$→S/, totales, distribución mensual, controles),
  * reajuste remunerativo por mes, becas/descuentos y morosidad,
  * agregados maestro-detalle (ingresos, gasto, resultado, margen, punto de equilibrio),
  * generador de códigos correlativos por área, plantilla «Correr Actividades de Funcionamiento»,
  * las 12 secciones de la Ficha_Resumen y las 11 directrices automáticas de la DPDI,
  * la base de datos consolidada NIIF mensualizada (Power BI / ERP),
  * el importador de fichas Excel existentes (datos y parámetros).

No depende de la interfaz gráfica: puede usarse desde scripts o pruebas.
"""
from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from catalogs_seed import (BASES_CALCULO, CATEGORIAS_EGR_RESUMEN, CATEGORIAS_ING_RESUMEN, MACROGRUPOS,
                           UNIDADES_MEDIDA)
from database import DETALLE_TABLES, MCOLS, MESES, Database

TOL = 0.01
KINDS_MASTER = ("INI", "OO", "AP", "INV")
CODE_FORMAT = {"INI": ("INI", 2), "FI": ("FI", 3), "FE": ("FE", 3), "OO": ("OO", 3), "AP": ("AP", 3), "INV": ("INV", 3)}
KIND_LABEL = {
    "INI": "Iniciativas con impacto estratégico",
    "FI": "Actividades de funcionamiento",
    "FE": "Actividades de funcionamiento",
    "OO": "Otras actividades operativas",
    "AP": "Actividades propias",
    "INV": "Actividades de inversión",
}
MASTER_TABLE = {k: v[0] for k, v in DETALLE_TABLES.items()}

# Catálogos estratégicos editables (las Acciones estratégicas del PEI ya no se usan: el eje cumple ese rol).
CATALOGOS_ESTRATEGICOS = {
    "EJE": "Eje estratégico", "OE": "Objetivo estratégico del PEI (OE)", "OBJ": "Objetivo anual del POA (OBJ)",
    "ORIGEN": "Origen de iniciativa", "PRIORIDAD": "Nivel de prioridad", "SUSTENTO_EGR": "Sustento de egresos",
    "SUSTENTO_ING": "Sustento de ingresos", "CAT_OO": "Categoría de otras operativas",
    "TIPO_AP": "Tipo de actividad propia", "TIPO_INV": "Tipo de inversión",
    "TIPO_ACTIVIDAD": "Clasificación de actividades",
}

# Datos generales y supuestos (clave, etiqueta, descripción)
PARAM_GENERALES = [
    ("institucion", "Institución", ""), ("codigo_formato", "Código del formato", ""),
    ("version_formato", "Versión del formato", ""), ("fecha_aprobacion", "Fecha de aprobación", "AAAA-MM-DD"),
    ("anio_formulacion", "Año de formulación", "Ejemplo: 2027"),
    ("fecha_emision", "Fecha de emisión de los parámetros", "AAAA-MM-DD"),
    ("fecha_limite", "Fecha límite de presentación de la ficha", "AAAA-MM-DD"),
    ("inflacion_proyectada", "Inflación proyectada (%)", "Referencial para estimar costos sin cotización."),
    ("tipo_cambio", "Tipo de cambio de referencia (S/ por US$)", "Convierte a soles los montos en US$."),
    ("valor_uit", "Valor de la UIT (S/)", ""), ("rmv", "Remuneración mínima vital (S/)", ""),
    ("reajuste_remunerativo_pct", "Reajuste remunerativo (%)", "Filas de Funcionamiento_Egr con «Aplica reajuste = Sí»."),
    ("mes_inicio_reajuste", "Mes desde el que se aplica el reajuste (1–12)", "1 = enero … 12 = diciembre."),
    ("tasa_morosidad_pct", "Tasa de morosidad proyectada (%)", "Filas de Funcionamiento_Ing con «Aplica morosidad = Sí»."),
    ("estudiantes_pregrado", "Estudiantes proyectados — pregrado", ""),
    ("estudiantes_posgrado", "Estudiantes proyectados — posgrado", ""),
    ("pension_ref_pregrado", "Pensión promedio de referencia — pregrado (S/)", ""),
    ("pension_ref_posgrado", "Pensión promedio de referencia — posgrado (S/)", ""),
    ("cuotas_periodo", "N° de cuotas de pensión por periodo académico", ""),
    ("director_dpdi", "Director(a) de Planificación y Desarrollo Institucional", "Firma «Revisado por»."),
    ("cargo_dpdi", "Cargo de quien revisa (DPDI)", ""),
]
PARAM_PCT = {"inflacion_proyectada", "reajuste_remunerativo_pct", "tasa_morosidad_pct"}
PARAM_NUM = {"tipo_cambio", "valor_uit", "rmv", "estudiantes_pregrado", "estudiantes_posgrado", "pension_ref_pregrado",
             "pension_ref_posgrado", "cuotas_periodo"}

# Estructura de la plantilla «Parámetros y Catálogos» (hoja → [(clave, encabezado)])
PLANTILLA_PARAM = {
    "Datos_Generales": [("clave", "Clave"), ("parametro", "Parámetro"), ("valor", "Valor"),
                        ("nota", "Descripción")],
    "Unidades_Techos": [("id", "ID"), ("codigo_area", "Código de área"), ("macrogrupo", "Macrogrupo"),
                        ("nombre_unidad", "Nombre de la unidad orgánica"), ("centro_costo", "Centro de costo"),
                        ("responsable", "Responsable"), ("cargo", "Cargo"), ("correo", "Correo electrónico"),
                        ("genera_ingresos", "¿Genera ingresos? (Sí/No)"),
                        ("techo_iniciativas", "Techo iniciativas (S/)"),
                        ("techo_funcionamiento", "Techo funcionamiento (S/)"),
                        ("techo_otras_operativas", "Techo otras operativas (S/)"),
                        ("techo_inversion", "Techo inversión (S/)"),
                        ("techo_planilla_docente", "Techo planilla docente (S/)"),
                        ("techo_planilla_adm", "Techo planilla administrativa (S/)"),
                        ("techo_marketing", "Techo marketing (S/)"),
                        ("meta_ingresos_func", "Meta ingresos de funcionamiento (S/)"),
                        ("meta_resultado_otras", "Meta ingresos otras operativas (S/)"),
                        ("meta_resultado_propias", "Meta resultado actividades propias (S/)")],
    "Cuentas_NIIF_Partidas_GG": [("id", "ID"), ("tipo_flujo", "Tipo (Egresos/Ingresos/Inversion)"),
                                 ("partida_gg", "Partida (Partida GG que verá el usuario)"),
                                 ("descripcion", "Descripción orientativa"), ("codigo_niif", "Código NIIF"),
                                 ("partida_rubro", "Cuenta contable NIIF"), ("categoria", "Categoría presupuestal"),
                                 ("naturaleza", "Naturaleza"), ("resultado_funcion", "Resultado por función"),
                                 ("activo", "Activo (Sí/No)")],
    "Catalogos_Estrategicos": [("id", "ID"), ("tipo_catalogo", "Tipo de catálogo"), ("codigo", "Código"),
                               ("descripcion", "Descripción"), ("eje", "Eje estratégico (solo objetivos POA)"),
                               ("padre", "OE del PEI (solo objetivos POA)"),
                               ("extra", "Alcance / indicador / criterio"), ("orden", "Orden")],
    "Costos_Referencia": [("id", "ID"), ("concepto", "Concepto"), ("unidad_medida", "Unidad de medida"),
                          ("costo_unitario", "Costo unitario (S/)"), ("fuente", "Fuente / vigencia")],
    "Plantilla_Funcionamiento": [("id", "ID"), ("ambito", "Ámbito (GLOBAL o código de área)"),
                                 ("macrogrupo", "Macrogrupo (* = todos)"),
                                 ("solo_genera_ingresos", "Solo unidades que generan ingresos (Sí/No)"),
                                 ("tipo_flujo", "Tipo (Egresos/Ingresos)"), ("partida", "Cuenta NIIF (partida)"),
                                 ("base_calculo", "Base de cálculo"), ("porcentaje", "Porcentaje de la base"),
                                 ("monto_fijo", "Monto fijo anual (S/)"),
                                 ("aplica_ajuste", "Aplica reajuste / morosidad (Sí/No)"), ("sustento", "Sustento"),
                                 ("activo", "Activo (Sí/No)"), ("orden", "Orden")],
}

DETAIL_TABLE = {k: v[1] for k, v in DETALLE_TABLES.items()}
FLAT_TABLE = {"FI": "funcionamiento_ing", "FE": "funcionamiento_egr"}


# =============================================================================
# Utilidades numéricas y de texto
# =============================================================================
def to_float(value: Any, default: Optional[float] = None) -> Optional[float]:
    """Convierte números de Excel o texto ('1,234.50', 'S/ 200', ' 15 ') en float."""
    if value is None:
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        return float(value)
    text = str(value).strip().replace("S/", "").replace("US$", "").replace(" ", "")
    if not text or text in ("-", "x", "X"):
        return default
    pct = text.endswith("%")
    text = text.rstrip("%")
    if text.count(",") and text.count("."):
        text = text.replace(",", "")
    elif text.count(",") == 1 and len(text.split(",")[1]) != 3:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        num = float(text)
    except ValueError:
        return default
    return num / 100.0 if pct else num


def parse_pct(value: Any) -> Optional[float]:
    """'10%' → 0.10 ; '10' → 0.10 ; '0.1' → 0.10 (las fracciones ≤ 1 se respetan)."""
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, str) and value.strip().endswith("%"):
        return to_float(value)
    num = to_float(value)
    if num is None:
        return None
    return num / 100.0 if abs(num) > 1 else num


def mark_factor(value: Any) -> Optional[float]:
    """Marca mensual de Funcionamiento: «x» = 1 pago/cobro completo; número = factor de distribución."""
    if value is None:
        return None
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        if s.lower() == "x":
            return 1.0
    return to_float(value)


def norm_key(text: Any) -> str:
    """Normaliza para búsquedas insensibles a mayúsculas/tildes/espacios (como VLOOKUP de Excel)."""
    s = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]+", " ", s.upper()).strip()


def s(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def fmt_money(value: Any) -> str:
    num = to_float(value)
    if num is None:
        return ""
    if abs(num) < 0.005:
        num = 0.0
    return f"{num:,.2f}"


def fmt_pct(value: Any) -> str:
    num = to_float(value)
    if num is None:
        return ""
    return f"{num * 100:,.1f}%"


def fmt_mark(value: Any) -> str:
    if value is None or str(value).strip() == "":
        return ""
    if str(value).strip().lower() == "x":
        return "x"
    num = to_float(value)
    if num is None:
        return str(value)
    return f"{num:.8f}".rstrip("0").rstrip(".") if not num.is_integer() else str(int(num))


def split_display(value: str) -> str:
    """'EJ-06 · Transformación Digital' → 'Transformación Digital'."""
    if value and " · " in value:
        return value.split(" · ", 1)[1].strip()
    return (value or "").strip()


def code_display(value: str) -> str:
    """'OE3 · Descripción' → 'OE3'."""
    if value and " · " in value:
        return value.split(" · ", 1)[0].strip()
    return (value or "").strip()


# =============================================================================
# Parámetros
# =============================================================================
@dataclass
class Params:
    anio: str = "2027"
    tc: float = 3.5
    inflacion: float = 0.025
    reajuste: float = 0.0
    mes_reajuste: int = 1
    morosidad: float = 0.0
    raw: Optional[dict] = None

    @classmethod
    def from_db(cls, db: Database) -> "Params":
        p = db.parametros()
        mes = int(to_float(p.get("mes_inicio_reajuste"), 1) or 1)
        return cls(
            anio=s(p.get("anio_formulacion")) or "2027",
            tc=to_float(p.get("tipo_cambio"), 3.5) or 3.5,
            inflacion=to_float(p.get("inflacion_proyectada"), 0.0) or 0.0,
            reajuste=to_float(p.get("reajuste_remunerativo_pct"), 0.0) or 0.0,
            mes_reajuste=min(12, max(1, mes)),
            morosidad=to_float(p.get("tasa_morosidad_pct"), 0.0) or 0.0,
            raw=p,
        )


# =============================================================================
# Cálculos por fila (equivalentes a las fórmulas de cada hoja)
# =============================================================================
def to_soles(amount: Optional[float], moneda: str, tc: float) -> Optional[float]:
    """=IF(F="","",IF(E="US$",F*TC,F))"""
    if amount is None:
        return None
    return amount * tc if (moneda or "").strip() == "US$" else amount


def control_distribucion(total: float, programado: float) -> str:
    """=IF(AND(L=0,Y=0),"",IF(ABS(Y-L)<0.01,"OK","Revisar"))"""
    if abs(total) < 1e-12 and abs(programado) < 1e-12:
        return ""
    return "OK" if abs(programado - total) < TOL else "Revisar"


def calc_detalle(row: dict, p: Params) -> dict:
    """Detalle_Iniciativas / Detalle_Otras_Operativas / Detalle_Act_Prop / Detalle_Inversiones."""
    cant = to_float(row.get("cantidad"))
    cu = to_float(row.get("costo_unit_moneda"))
    cu_s = to_soles(cu, row.get("moneda"), p.tc)
    total = 0.0 if cant is None or cu_s is None else cant * cu_s
    meses = [to_float(row.get(m), 0.0) or 0.0 for m in MCOLS]
    programado = sum(meses)
    return {
        "costo_unit_soles": cu_s,
        "total": total,
        "meses": meses,
        "programado": programado,
        "control": control_distribucion(total, programado),
    }


def calc_fing(row: dict, p: Params) -> dict:
    """Funcionamiento_Ing: ingreso mensual neto = Cant × Tarifa S/ × factor × (1−Becas) × (1−Morosidad)."""
    cant = to_float(row.get("cantidad"))
    tarifa = to_float(row.get("tarifa_moneda"))
    tarifa_s = to_soles(tarifa, row.get("moneda"), p.tc)
    becas = to_float(row.get("becas_pct"), 0.0) or 0.0
    moros = p.morosidad if (row.get("aplica_morosidad") or "").strip() == "Sí" else 0.0
    factores = [mark_factor(row.get(m)) for m in MCOLS]
    meses = []
    for f in factores:
        if cant is None or tarifa_s is None or f is None:
            meses.append(0.0)
        else:
            meses.append(cant * tarifa_s * f * (1 - becas) * (1 - moros))
    suma_f = sum(f for f in factores if f is not None)
    bruto = 0.0 if cant is None or tarifa_s is None else cant * tarifa_s * suma_f
    neto = sum(meses)
    return {
        "tarifa_soles": tarifa_s,
        "meses": meses,
        "bruto": bruto,
        "neto": neto,
        "descuentos": bruto - neto,
        "n_meses": sum(1 for f in factores if f not in (None, 0.0)),
    }


def calc_fegr(row: dict, p: Params) -> dict:
    """Funcionamiento_Egr: monto mensual = factor × Coste total × (1+Reajuste si aplica y mes ≥ MesReajuste)."""
    cant = to_float(row.get("cantidad"))
    cu = to_float(row.get("costo_unit_moneda"))
    cu_s = to_soles(cu, row.get("moneda"), p.tc)
    coste = 0.0 if cant is None or cu_s is None else cant * cu_s
    aplica = (row.get("aplica_reajuste") or "").strip() == "Sí"
    factores = [mark_factor(row.get(m)) for m in MCOLS]
    meses = []
    for i, f in enumerate(factores, start=1):
        if cant is None or cu_s is None or f is None:
            meses.append(0.0)
        else:
            ajuste = (1 + p.reajuste) if (aplica and i >= p.mes_reajuste) else 1.0
            meses.append(f * coste * ajuste)
    total = sum(meses)
    marcas_x = any(isinstance(row.get(m), str) and row.get(m).strip().lower() == "x" for m in MCOLS)
    suma_f = sum(f for f in factores if f is not None)
    if abs(coste) < 1e-12 and abs(total) < 1e-12:
        control = ""
    elif marcas_x:
        control = "OK"
    else:
        control = "OK" if abs(suma_f - 1.0) < 1e-4 else "Revisar"
    return {
        "costo_unit_soles": cu_s,
        "coste_total": coste,
        "meses": meses,
        "total": total,
        "n_pagos": sum(1 for f in factores if f not in (None, 0.0)),
        "control": control,
    }


def control_meta(row: dict) -> str:
    """=IF(C="","",IF(M="","Sin meta",IF(SUM(N:P)=M,"OK","Revisar")))"""
    if not s(row.get("nombre")):
        return ""
    meta = to_float(row.get("meta_anual"))
    if meta is None:
        return "Sin meta"
    suma = sum(to_float(row.get(t), 0.0) or 0.0 for t in ("t1", "t2", "t3"))
    return "OK" if abs(suma - meta) < 1e-9 else "Revisar"


def estado_techo(egresos: float, techo: float) -> str:
    """=IF(F=0,IF(D=0,"","Sin techo"),IF(D<=F,"Dentro del techo","Excede techo"))"""
    if not techo:
        return "" if abs(egresos) < 1e-9 else "Sin techo"
    return "Dentro del techo" if egresos <= techo + 1e-9 else "Excede techo"


def estado_autofin(ingresos: float, egresos: float) -> str:
    if abs(ingresos) < 1e-9 and abs(egresos) < 1e-9:
        return ""
    return "Autofinanciada" if ingresos - egresos >= -1e-9 else "Déficit"


def safe_div(a: float, b: float) -> Optional[float]:
    return None if not b else a / b


def distribuir(total: float, meses_sel: Iterable[int]) -> list[Optional[float]]:
    """Distribuye equitativamente un total en los meses seleccionados (1..12), cuadrando centavos."""
    sel = sorted(set(m for m in meses_sel if 1 <= m <= 12))
    out: list[Optional[float]] = [None] * 12
    if not sel:
        return out
    base = round(total / len(sel), 2)
    for m in sel:
        out[m - 1] = base
    out[sel[-1] - 1] = round(total - base * (len(sel) - 1), 2)
    return out


PRORRATEO_12 = 1.0 / 12.0


# =============================================================================
# Controlador principal
# =============================================================================
class POAController:
    def __init__(self, db: Database):
        self.db = db
        self._niif_cache: Optional[dict] = None

    # ------------------------------------------------------------- contexto
    @property
    def params(self) -> Params:
        return Params.from_db(self.db)

    @property
    def anio(self) -> str:
        return self.params.anio

    def unidades(self) -> list[dict]:
        return self.db.all("unidades_organicas", order="macrogrupo, nombre_unidad")

    def unidad(self, uid: int) -> dict:
        u = self.db.get("unidades_organicas", uid)
        if not u:
            raise ValueError(f"Unidad orgánica {uid} no existe")
        return u

    def unidad_display(self, u: dict) -> str:
        return f"{u['codigo_area']} · {u['nombre_unidad']}"

    def save_unidad(self, data: dict, uid: Optional[int] = None) -> int:
        data = dict(data)
        if not s(data.get("codigo_area")) or not s(data.get("nombre_unidad")):
            raise ValueError("El código de área y el nombre de la unidad son obligatorios.")
        data["codigo_area"] = s(data["codigo_area"]).upper()
        for k in ("techo_iniciativas", "techo_funcionamiento", "techo_otras_operativas", "techo_inversion",
                  "techo_planilla_docente", "techo_planilla_adm", "techo_marketing", "meta_ingresos_func",
                  "meta_resultado_otras", "meta_resultado_propias"):
            if k in data:
                data[k] = to_float(data[k], 0.0) or 0.0
        if "genera_ingresos" in data:
            data["genera_ingresos"] = 1 if str(data["genera_ingresos"]).strip() in ("1", "True", "Sí", "true") else 0
        return self.db.upsert("unidades_organicas", data, uid)

    def delete_unidad(self, uid: int) -> None:
        self.db.delete("unidades_organicas", uid)

    # ------------------------------------------------------------- catálogos
    def _niif(self) -> dict:
        if self._niif_cache is None:
            rows = self.db.all("catalogo_niif_partidas", "activo = 1", order="id")
            by_tipo: dict[str, dict[str, dict]] = {"Ingresos": {}, "Egresos": {}, "Inversion": {}}
            by_gg: dict[str, dict[str, dict]] = {"Ingresos": {}, "Egresos": {}, "Inversion": {}}
            for r in rows:
                by_tipo.setdefault(r["tipo_flujo"], {}).setdefault(norm_key(r["partida_rubro"]), r)
                if s(r["partida_gg"]):
                    by_gg.setdefault(r["tipo_flujo"], {}).setdefault(norm_key(r["partida_gg"]), r)
            self._niif_cache = {"cuenta": by_tipo, "gg": by_gg}
        return self._niif_cache

    def invalidate_cache(self) -> None:
        self._niif_cache = None

    def partidas(self, tipo: str) -> list[str]:
        """Partidas NIIF por tipo de flujo, en orden alfabético (lista desplegable)."""
        rows = self.db.query(
            "SELECT DISTINCT partida_rubro FROM catalogo_niif_partidas WHERE tipo_flujo = ? AND activo = 1 "
            "ORDER BY partida_rubro", (tipo,))
        return [r["partida_rubro"] for r in rows]

    def partidas_gg(self, tipo: str) -> list[str]:
        """Partidas GG (combo «Partida» del detalle) de la base consolidada Cuentas NIIF + Partidas GG."""
        rows = self.db.query("SELECT DISTINCT partida_gg FROM catalogo_niif_partidas WHERE tipo_flujo = ? AND activo = 1 "
                             "AND partida_gg <> '' ORDER BY partida_gg COLLATE NOCASE", (tipo,))
        return [r["partida_gg"] for r in rows]

    def gg_info(self, partida: str, tipo: str = "") -> Optional[dict]:
        """Cuenta NIIF principal (la primera registrada) de una Partida GG: código, cuenta, categoría y descripción."""
        tipos = [tipo] if tipo in ("Ingresos", "Egresos", "Inversion") else []
        tipos += [t for t in ("Egresos", "Ingresos", "Inversion") if t not in tipos]
        key = norm_key(partida)
        gg = self._niif()["gg"]
        for t in tipos:
            if key in gg.get(t, {}):
                return gg[t][key]
        return None

    def descripcion_partida(self, partida: str, tipo: str = "") -> str:
        """Texto orientativo que se muestra al elegir una Partida."""
        r = self.gg_info(partida, tipo) or self.buscar_partida(partida, [tipo] if tipo else
                                                               ("Egresos", "Ingresos", "Inversion"))
        if not r:
            return ""
        desc = s(r.get("descripcion")) or "(sin descripción registrada)"
        return f"{desc}\nCuenta NIIF: {r['codigo_niif']} · {r['partida_rubro']} · Categoría: {r['categoria'] or '—'}"

    def a_partida_gg(self, partida: str, tipo: str) -> str:
        """Normaliza a Partida GG: si se recibe el nombre de una cuenta NIIF, devuelve su Partida GG."""
        if not s(partida):
            return ""
        hit = self.gg_info(partida, tipo)
        if hit:
            return hit["partida_gg"]
        cuenta = self.buscar_partida(partida, [tipo] if tipo else ("Egresos", "Ingresos", "Inversion"))
        if cuenta and s(cuenta["partida_gg"]):
            return cuenta["partida_gg"]
        return s(partida)

    def buscar_partida(self, partida: str, tipos: Iterable[str] = ("Egresos", "Ingresos", "Inversion")) -> Optional[dict]:
        """Busca por nombre de cuenta NIIF y, si no existe, por Partida GG."""
        key = norm_key(partida)
        if not key:
            return None
        cat = self._niif()
        tipos = list(tipos)
        for t in tipos:
            if key in cat["cuenta"].get(t, {}):
                return cat["cuenta"][t][key]
        for t in tipos:
            if key in cat["gg"].get(t, {}):
                return cat["gg"][t][key]
        return None

    def categoria_de(self, partida: str, tipo: str = "") -> str:
        """VLOOKUP(Partida, Parámetros!C117:D408, 2, FALSE) → Categoría."""
        tipos = [tipo] if tipo in ("Ingresos", "Egresos", "Inversion") else []
        tipos += [t for t in ("Egresos", "Ingresos", "Inversion") if t not in tipos]
        r = self.buscar_partida(partida, tipos)
        return (r["categoria"] or "").strip() if r else ""

    def catalogo_values(self, tipo_catalogo: str, display: str = "codigo") -> list[str]:
        rows = self.db.catalogo(tipo_catalogo)
        if display == "codigo":
            return [r["codigo"] for r in rows]
        if display == "descripcion":
            return [r["descripcion"] for r in rows]
        return [f"{r['codigo']} · {r['descripcion']}" for r in rows]

    def ejes_display(self) -> list[str]:
        return self.catalogo_values("EJE", "both")

    def eje_display(self, nombre: str) -> str:
        for r in self.db.catalogo("EJE"):
            if norm_key(r["descripcion"]) == norm_key(nombre) or norm_key(r["codigo"]) == norm_key(nombre):
                return f"{r['codigo']} · {r['descripcion']}"
        return nombre or ""

    def eje_codigo(self, nombre: str) -> str:
        for r in self.db.catalogo("EJE"):
            if norm_key(r["descripcion"]) == norm_key(nombre):
                return r["codigo"]
        return ""

    def eje_ref(self, eje: str) -> Optional[dict]:
        """Eje por código (EJ-06), nombre (Transformación Digital) o texto combinado del combo."""
        key = norm_key(split_display(eje)) if eje and " · " in eje else norm_key(eje)
        cod = norm_key(code_display(eje))
        for r in self.db.catalogo("EJE"):
            if key in (norm_key(r["descripcion"]), norm_key(r["codigo"])) or cod == norm_key(r["codigo"]):
                return r
        return None

    def objs_por_eje(self, eje: str, display: str = "both") -> list[str]:
        """Objetivos anuales del POA vinculados al eje seleccionado (combo dependiente)."""
        ref = self.eje_ref(eje)
        if not ref:
            return []
        rows = [r for r in self.db.catalogo("OBJ") if norm_key(r["eje"]) == norm_key(ref["codigo"])]
        return [f"{r['codigo']} · {r['descripcion']}" if display == "both" else r["codigo"] for r in rows]

    def oe_de_obj(self, obj: str) -> str:
        """OE del PEI vinculado al objetivo anual (se carga automáticamente)."""
        r = self.db.query_one("SELECT padre FROM catalogos_estrategicos WHERE tipo_catalogo='OBJ' AND codigo=?",
                              (code_display(obj),))
        return r["padre"] if r else ""

    def eje_de_obj(self, obj: str) -> str:
        r = self.db.query_one("SELECT eje FROM catalogos_estrategicos WHERE tipo_catalogo='OBJ' AND codigo=?",
                              (code_display(obj),))
        return r["eje"] if r else ""

    def oe_display(self, oe: str) -> str:
        for r in self.db.catalogo("OE"):
            if r["codigo"] == code_display(oe):
                return f"{r['codigo']} · {r['descripcion']}"
        return oe or ""

    @staticmethod
    def unidades_medida() -> list[str]:
        return list(UNIDADES_MEDIDA)

    def save_catalogo(self, tipo_catalogo: str, data: dict, row_id: Optional[int] = None) -> int:
        data = dict(data, tipo_catalogo=tipo_catalogo)
        if not s(data.get("codigo")):
            raise ValueError("El código es obligatorio.")
        return self.db.upsert("catalogos_estrategicos", data, row_id)

    def save_partida_niif(self, data: dict, row_id: Optional[int] = None) -> int:
        if not s(data.get("partida_rubro")) or data.get("tipo_flujo") not in ("Ingresos", "Egresos", "Inversion"):
            raise ValueError("Partida / rubro y Tipo (Ingresos/Egresos/Inversion) son obligatorios.")
        data = dict(data)
        data["codigo_niif"] = s(data.get("codigo_niif")) or "SIN-CODIGO"
        rid = self.db.upsert("catalogo_niif_partidas", data, row_id)
        self.invalidate_cache()
        return rid

    def delete_partida_niif(self, row_id: int) -> None:
        self.db.delete("catalogo_niif_partidas", row_id)
        self.invalidate_cache()

    # ---------------------------------------------------------------- códigos
    def _existing_codes(self, kind: str, uid: int) -> list[str]:
        table = MASTER_TABLE.get(kind) or FLAT_TABLE[kind]
        rows = self.db.query(f"SELECT codigo FROM {table} WHERE unidad_id = ? AND anio = ?", (uid, self.anio))
        return [r["codigo"] for r in rows]

    def next_code(self, kind: str, uid: int, offset: int = 0) -> str:
        prefix, width = CODE_FORMAT[kind]
        nums = []
        for c in self._existing_codes(kind, uid):
            m = re.match(rf"^{prefix}-(\d+)$", (c or "").strip().upper())
            if m:
                nums.append(int(m.group(1)))
        n = (max(nums) if nums else 0) + 1 + offset
        return f"{prefix}-{n:0{width}d}"

    def codigo_bd(self, uid: int, codigo: str) -> str:
        """Código prefijado con el área para la base consolidada (p. ej. OL-INI-01)."""
        return f"{self.unidad(uid)['codigo_area']}-{codigo}"

    def generar_codigos(self, uid: int, cantidades: dict[str, int]) -> list[tuple[str, str, str]]:
        """Crea registros correlativos vacíos (placeholders) por tipo de actividad aprobada."""
        creados = []
        area = self.unidad(uid)["codigo_area"]
        with self.db.transaction():
            for kind in ("INI", "FI", "FE", "OO", "AP", "INV"):
                n = int(cantidades.get(kind) or 0)
                for _ in range(max(0, n)):
                    code = self.next_code(kind, uid)
                    table = MASTER_TABLE.get(kind) or FLAT_TABLE[kind]
                    data = {"unidad_id": uid, "anio": self.anio, "codigo": code}
                    if kind == "INV":
                        data["ficha_sustento"] = "No"
                    if kind in FLAT_TABLE:
                        data["orden"] = self._next_orden(table, uid)
                    self.db.insert(table, data, commit=False)
                    creados.append((kind, code, f"{area}-{code}"))
        return creados

    def _next_orden(self, table: str, uid: int) -> int:
        return int(self.db.scalar(f"SELECT COALESCE(MAX(orden),0)+1 FROM {table} WHERE unidad_id=? AND anio=?",
                                  (uid, self.anio)) or 1)

    # ---------------------------------------------------------- maestros
    def list_masters(self, kind: str, uid: int) -> list[dict]:
        table = MASTER_TABLE[kind]
        rows = self.db.all(table, "unidad_id = ? AND anio = ?", (uid, self.anio), "codigo")
        p = self.params
        for r in rows:
            self._enrich_master(kind, r, self.list_detalle(kind, r["id"], p))
        return rows

    def get_master(self, kind: str, mid: int) -> dict:
        r = self.db.get(MASTER_TABLE[kind], mid)
        if r:
            self._enrich_master(kind, r, self.list_detalle(kind, mid))
        return r

    def _enrich_master(self, kind: str, r: dict, detalles: list[dict]) -> None:
        if kind == "INI":
            ing = sum(d["_c"]["total"] for d in detalles if d["tipo"] == "Ingresos")
            egr = sum(d["_c"]["total"] for d in detalles if d["tipo"] == "Egresos")
            r["_ingresos"], r["_egresos"] = ing, egr
            r["_control_meta"] = control_meta(r)
        elif kind in ("OO", "AP"):
            ing = sum(d["_c"]["programado"] for d in detalles if d["tipo"] == "Ingresos")
            egr = sum(d["_c"]["programado"] for d in detalles if d["tipo"] == "Egresos")
            r["_ingresos"], r["_egresos"] = ing, egr
        else:
            r["_ingresos"] = 0.0
            r["_egresos"] = sum(d["_c"]["programado"] for d in detalles)
        r["_resultado"] = r["_ingresos"] - r["_egresos"]
        r["_margen"] = safe_div(r["_resultado"], r["_ingresos"])
        r["_n_detalle"] = len(detalles)
        if kind == "AP":
            part_det = sum(to_float(d.get("cantidad"), 0.0) or 0.0 for d in detalles if d["tipo"] == "Ingresos")
            participantes = to_float(r.get("participantes"))
            r["_participantes"] = participantes if participantes is not None else part_det
            tarifa = to_float(r.get("tarifa"))
            becas = to_float(r.get("becas_pct"), 0.0) or 0.0
            if tarifa is not None:
                tarifa_neta = tarifa * (1 - becas)
            else:
                tarifa_neta = safe_div(r["_ingresos"], part_det) or 0.0
            r["_tarifa_neta"] = tarifa_neta
            r["_ingreso_proyectado"] = (participantes or 0.0) * (tarifa or 0.0) * (1 - becas)
            r["_punto_equilibrio"] = (math.ceil(round(r["_egresos"] / tarifa_neta, 9))
                                      if tarifa_neta and tarifa_neta > 0 and r["_egresos"] > 0 else None)
            r["_estado"] = estado_autofin(r["_ingresos"], r["_egresos"])

    def save_master(self, kind: str, uid: int, data: dict, mid: Optional[int] = None) -> int:
        table = MASTER_TABLE[kind]
        data = dict(data)
        data["unidad_id"], data["anio"] = uid, self.anio
        data["codigo"] = s(data.get("codigo")).upper() or self.next_code(kind, uid)
        dup = self.db.query_one(f"SELECT id FROM {table} WHERE unidad_id=? AND anio=? AND codigo=?",
                                (uid, self.anio, data["codigo"]))
        if dup and dup["id"] != mid:
            raise ValueError(f"El código {data['codigo']} ya existe en esta unidad.")
        if kind == "INI":
            ref = self.eje_ref(data.get("eje", ""))
            data["eje"] = ref["descripcion"] if ref else split_display(data.get("eje", ""))
            data["obj_poa"] = code_display(data.get("obj_poa", ""))
            if data["obj_poa"]:
                eje_obj = self.eje_de_obj(data["obj_poa"])
                if ref and eje_obj and eje_obj != ref["codigo"]:
                    raise ValueError(f"El objetivo anual {data['obj_poa']} no pertenece al eje {ref['codigo']} "
                                     f"(está vinculado a {eje_obj}).")
                data["oe"] = self.oe_de_obj(data["obj_poa"]) or code_display(data.get("oe", ""))
            else:
                data["oe"] = code_display(data.get("oe", ""))
            data["ae"] = ""
            um = s(data.get("unidad_medida"))
            if um:
                data["unidad_medida"] = UNIDADES_MEDIDA[1] if ("%" in um or "PORCENT" in norm_key(um)) else UNIDADES_MEDIDA[0]
            for k in ("meta_anual", "t1", "t2", "t3"):
                data[k] = to_float(data.get(k))
            pr = to_float(code_display(s(data.get("prioridad"))))
            data["prioridad"] = int(pr) if pr else None
        elif kind == "OO":
            data["categoria"] = split_display(data.get("categoria", ""))
        elif kind == "AP":
            data["participantes"] = to_float(data.get("participantes"))
            data["tarifa"] = to_float(data.get("tarifa"))
            data["becas_pct"] = parse_pct(data.get("becas_pct"))
        elif kind == "INV":
            pr = to_float(code_display(s(data.get("prioridad"))))
            data["prioridad"] = int(pr) if pr else None
            data["ficha_sustento"] = data.get("ficha_sustento") or "No"
        return self.db.upsert(table, data, mid)

    def delete_master(self, kind: str, mid: int) -> None:
        self.db.delete(MASTER_TABLE[kind], mid)

    # ----------------------------------------------------------- detalles
    def list_detalle(self, kind: str, mid: int, p: Optional[Params] = None) -> list[dict]:
        p = p or self.params
        rows = self.db.all(DETAIL_TABLE[kind], "master_id = ?", (mid,), "orden, id")
        for r in rows:
            r["_c"] = calc_detalle(r, p)
        return rows

    def save_detalle(self, kind: str, mid: int, data: dict, did: Optional[int] = None) -> int:
        data = dict(data)
        data["master_id"] = mid
        if kind == "INV":
            data["tipo"] = "Inversion"
        if data.get("tipo") not in ("Ingresos", "Egresos", "Inversion"):
            raise ValueError("Seleccione el Tipo (Ingresos / Egresos).")
        data["partida"] = self.a_partida_gg(data.get("partida", ""), data["tipo"])
        data["partida_gg"] = data["partida"]
        data["categoria"] = self.categoria_de(data["partida"], data["tipo"]) or s(data.get("categoria"))
        data["cantidad"] = to_float(data.get("cantidad"))
        data["costo_unit_moneda"] = to_float(data.get("costo_unit_moneda"))
        data["moneda"] = data.get("moneda") or "S/"
        for m in MCOLS:
            data[m] = to_float(data.get(m))
        if not did:
            data["orden"] = int(self.db.scalar(
                f"SELECT COALESCE(MAX(orden),0)+1 FROM {DETAIL_TABLE[kind]} WHERE master_id=?", (mid,)) or 1)
        return self.db.upsert(DETAIL_TABLE[kind], data, did)

    def delete_detalle(self, kind: str, did: int) -> None:
        self.db.delete(DETAIL_TABLE[kind], did)

    def crear_ingreso_desde_tarifa(self, mid: int) -> Optional[int]:
        """Actividad propia: genera el ítem de ingreso (participantes × tarifa) en el mes de ejecución."""
        ap = self.db.get("act_propias", mid)
        if not ap:
            return None
        part = to_float(ap.get("participantes"))
        tarifa = to_float(ap.get("tarifa"))
        if not part or tarifa is None:
            raise ValueError("Registre participantes proyectados y tarifa por participante.")
        becas = to_float(ap.get("becas_pct"), 0.0) or 0.0
        unit = self.unidad(ap["unidad_id"])
        partida = {
            "Posgrado": "CONGRESOS. SEMINARIOS, TALLERES, OTROS POSGRADO",
            "Educación Continua": "CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PEAS",
        }.get(unit["macrogrupo"], "CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO")
        mes = MESES.index(ap["mes_ejecucion"]) + 1 if ap.get("mes_ejecucion") in MESES else 1
        tarifa_neta = tarifa * (1 - becas)
        data = {"tipo": "Ingresos", "partida": partida, "descripcion": f"Inscripción: {s(part)} participantes",
                "cantidad": part, "moneda": "S/", "costo_unit_moneda": tarifa_neta,
                "sustento": "Tarifario institucional vigente"}
        data[MCOLS[mes - 1]] = part * tarifa_neta
        return self.save_detalle("AP", mid, data)

    # -------------------------------------------------------- funcionamiento
    def list_flat(self, kind: str, uid: int) -> list[dict]:
        table = FLAT_TABLE[kind]
        rows = self.db.all(table, "unidad_id = ? AND anio = ?", (uid, self.anio), "orden, id")
        p = self.params
        for r in rows:
            r["_c"] = calc_fing(r, p) if kind == "FI" else calc_fegr(r, p)
        return rows

    def save_flat(self, kind: str, uid: int, data: dict, rid: Optional[int] = None) -> int:
        table = FLAT_TABLE[kind]
        data = dict(data)
        data["unidad_id"], data["anio"] = uid, self.anio
        data["codigo"] = s(data.get("codigo")).upper() or self.next_code(kind, uid)
        tipo = "Ingresos" if kind == "FI" else "Egresos"
        data["categoria"] = self.categoria_de(data.get("partida", ""), tipo) or s(data.get("categoria"))
        data["cantidad"] = to_float(data.get("cantidad"))
        data["moneda"] = data.get("moneda") or "S/"
        if kind == "FI":
            data["tarifa_moneda"] = to_float(data.get("tarifa_moneda"))
            data["becas_pct"] = parse_pct(data.get("becas_pct"))
            data["aplica_morosidad"] = data.get("aplica_morosidad") or "No"
        else:
            data["costo_unit_moneda"] = to_float(data.get("costo_unit_moneda"))
            data["aplica_reajuste"] = data.get("aplica_reajuste") or "No"
        for m in MCOLS:
            v = data.get(m)
            if v is None or str(v).strip() == "":
                data[m] = None
            elif str(v).strip().lower() == "x":
                data[m] = "x"
            else:
                f = to_float(v)
                data[m] = None if f is None else repr(f)
        if not rid:
            data["orden"] = self._next_orden(table, uid)
        return self.db.upsert(table, data, rid)

    def delete_flat(self, kind: str, rid: int) -> None:
        self.db.delete(FLAT_TABLE[kind], rid)

    def renumerar_flat(self, kind: str, uid: int) -> None:
        """Renumera FI-001.. / FE-001.. en el orden de registro."""
        prefix, width = CODE_FORMAT[kind]
        rows = self.db.all(FLAT_TABLE[kind], "unidad_id = ? AND anio = ?", (uid, self.anio), "orden, id")
        with self.db.transaction():
            for i, r in enumerate(rows, start=1):
                self.db.update(FLAT_TABLE[kind], r["id"], {"codigo": f"{prefix}-{i:0{width}d}", "orden": i},
                               commit=False)

    # -------------------------------------------------- plantilla funcionamiento
    def plantilla(self, uid: Optional[int] = None) -> list[dict]:
        if uid is None:
            return self.db.all("plantilla_funcionamiento", "unidad_id IS NULL", order="orden, id")
        return self.db.all("plantilla_funcionamiento", "unidad_id = ?", (uid,), "orden, id")

    def plantilla_aplicable(self, uid: int) -> list[dict]:
        """Filas globales aplicables al macrogrupo + filas propias de la unidad (que reemplazan por partida)."""
        u = self.unidad(uid)
        propias = {(r["tipo_flujo"], norm_key(r["partida"])): r for r in self.plantilla(uid) if r["activo"]}
        out = []
        for r in self.plantilla(None):
            if not r["activo"]:
                continue
            if r["macrogrupo"] not in ("*", "", None) and r["macrogrupo"] != u["macrogrupo"]:
                continue
            if r["solo_genera_ingresos"] and not u["genera_ingresos"]:
                continue
            key = (r["tipo_flujo"], norm_key(r["partida"]))
            out.append(propias.pop(key, r))
        out.extend(propias.values())
        return out

    def monto_plantilla(self, row: dict, u: dict) -> float:
        p = self.params.raw or {}
        base = row["base_calculo"]
        pct = to_float(row.get("porcentaje"), 1.0)
        pct = 1.0 if pct is None else pct
        cuotas = to_float(p.get("cuotas_periodo"), 0.0) or 0.0
        if base == "MONTO_FIJO":
            return to_float(row.get("monto_fijo"), 0.0) or 0.0
        if base == "TECHO_PLANILLA_ADM":
            return (u["techo_planilla_adm"] or 0.0) * pct
        if base == "TECHO_PLANILLA_DOC":
            return (u["techo_planilla_docente"] or 0.0) * pct
        if base == "TECHO_MARKETING":
            return (u["techo_marketing"] or 0.0) * pct
        if base == "META_INGRESOS_FUNC":
            return (u["meta_ingresos_func"] or 0.0) * pct
        if base == "PENSION_PREGRADO":
            return (to_float(p.get("estudiantes_pregrado"), 0.0) or 0.0) * \
                (to_float(p.get("pension_ref_pregrado"), 0.0) or 0.0) * cuotas * 2 * pct
        if base == "PENSION_POSGRADO":
            return (to_float(p.get("estudiantes_posgrado"), 0.0) or 0.0) * \
                (to_float(p.get("pension_ref_posgrado"), 0.0) or 0.0) * cuotas * 2 * pct
        if base == "PROVISION_MOROSIDAD":
            ing = sum(r["_c"]["bruto"] for r in self.list_flat("FI", u["id"]) if not r["origen_plantilla"]
                      or r["categoria"] in ("PENSIÓN", "MATRICULA"))
            return ing * (to_float(p.get("tasa_morosidad_pct"), 0.0) or 0.0) * pct
        raise ValueError(f"Base de cálculo desconocida: {base} (válidas: {', '.join(BASES_CALCULO)})")

    def correr_funcionamiento(self, uid: int, reemplazar: bool = True) -> dict:
        """«Correr Actividades de Funcionamiento»: inserta la plantilla base prorrateada 1/12 (0.08333333)."""
        u = self.unidad(uid)
        filas = self.plantilla_aplicable(uid)
        resumen = {"FI": 0, "FE": 0, "eliminadas": 0, "monto_FI": 0.0, "monto_FE": 0.0}
        if reemplazar:
            for t in ("funcionamiento_ing", "funcionamiento_egr"):
                cur = self.db.execute(f"DELETE FROM {t} WHERE unidad_id=? AND anio=? AND origen_plantilla=1",
                                      (uid, self.anio))
                resumen["eliminadas"] += cur.rowcount
            self.db.conn.commit()
        ingresos_primero = sorted(filas, key=lambda r: (r["tipo_flujo"] != "Ingresos", r["orden"]))
        for r in ingresos_primero:
            kind = "FI" if r["tipo_flujo"] == "Ingresos" else "FE"
            monto = round(self.monto_plantilla(r, u), 2)
            data = {"partida": r["partida"], "cantidad": 1, "moneda": "S/", "sustento": r["sustento"] or "",
                    "observaciones": "Generado por plantilla base de funcionamiento", "origen_plantilla": 1}
            for m in MCOLS:
                data[m] = PRORRATEO_12
            if kind == "FI":
                data.update(tarifa_moneda=monto, aplica_morosidad=r["aplica_ajuste"] or "No")
            else:
                data.update(costo_unit_moneda=monto, aplica_reajuste=r["aplica_ajuste"] or "No")
            data["codigo"] = self.next_code(kind, uid)
            self.save_flat(kind, uid, data)
            resumen[kind] += 1
            resumen[f"monto_{kind}"] += monto
        return resumen

    # ------------------------------------------------------- techos y consumo
    def consumo_techo(self, uid: int, kind: str) -> dict:
        u = self.unidad(uid)
        techo = {"INI": u["techo_iniciativas"], "FE": u["techo_funcionamiento"], "FI": u["techo_funcionamiento"],
                 "OO": u["techo_otras_operativas"], "INV": u["techo_inversion"], "AP": 0.0}[kind] or 0.0
        if kind in KINDS_MASTER:
            masters = self.list_masters(kind, uid)
            ing = sum(m["_ingresos"] for m in masters)
            egr = sum(m["_egresos"] for m in masters)
        else:
            egr = sum(r["_c"]["total"] for r in self.list_flat("FE", uid))
            ing = sum(r["_c"]["neto"] for r in self.list_flat("FI", uid))
        estado = estado_autofin(ing, egr) if kind == "AP" else estado_techo(egr, techo)
        return {"ingresos": ing, "egresos": egr, "techo": techo, "uso": safe_div(egr, techo),
                "saldo": techo - egr, "estado": estado}

    # ====================================================================
    # Datos completos de una unidad (base para resumen, directrices y BD NIIF)
    # ====================================================================
    def load_unit_data(self, uid: int) -> dict:
        p = self.params
        data = {"unidad": self.unidad(uid), "params": p, "fi": self.list_flat("FI", uid),
                "fe": self.list_flat("FE", uid)}
        for kind in KINDS_MASTER:
            masters = self.db.all(MASTER_TABLE[kind], "unidad_id = ? AND anio = ?", (uid, self.anio), "codigo")
            for m in masters:
                m["_detalle"] = self.list_detalle(kind, m["id"], p)
                self._enrich_master(kind, m, m["_detalle"])
            data[kind] = masters
        return data

    # ====================================================================
    # FICHA RESUMEN — 12 secciones
    # ====================================================================
    def build_summary(self, uid: int, data: Optional[dict] = None) -> dict:
        d = data or self.load_unit_data(uid)
        u, p = d["unidad"], d["params"]
        ini, oo, ap, inv = d["INI"], d["OO"], d["AP"], d["INV"]
        fi, fe = d["fi"], d["fe"]
        out: dict[str, Any] = {"unidad": u, "params": p.raw, "anio": p.anio, "sections": []}

        # ---------------- 1. Resumen por tipo de actividad
        r_ini = (sum(m["_ingresos"] for m in ini), sum(m["_egresos"] for m in ini), u["techo_iniciativas"] or 0.0)
        r_fun = (sum(r["_c"]["neto"] for r in fi), sum(r["_c"]["total"] for r in fe), u["techo_funcionamiento"] or 0.0)
        r_oo = (sum(m["_ingresos"] for m in oo), sum(m["_egresos"] for m in oo), u["techo_otras_operativas"] or 0.0)
        r_ap = (sum(m["_ingresos"] for m in ap), sum(m["_egresos"] for m in ap))
        r_inv = (0.0, sum(m["_egresos"] for m in inv), u["techo_inversion"] or 0.0)
        s1 = []
        for nombre, (i_, e_, t_) in (("1. Iniciativas con impacto estratégico", r_ini),
                                      ("2. Actividades de funcionamiento", r_fun),
                                      ("3. Otras actividades operativas", r_oo)):
            s1.append([nombre, i_, e_, i_ - e_, t_, safe_div(e_, t_), estado_techo(e_, t_)])
        s1.append(["4. Actividades propias", r_ap[0], r_ap[1], r_ap[0] - r_ap[1], "Autofinanciada", None,
                   estado_autofin(*r_ap)])
        opex_i = r_ini[0] + r_fun[0] + r_oo[0] + r_ap[0]
        opex_e = r_ini[1] + r_fun[1] + r_oo[1] + r_ap[1]
        opex_t = r_ini[2] + r_fun[2] + r_oo[2]
        opex_e_techo = opex_e - r_ap[1]
        s1.append(["Total OPEX", opex_i, opex_e, opex_i - opex_e, opex_t, safe_div(opex_e_techo, opex_t),
                   estado_techo(opex_e_techo, opex_t)])
        s1.append(["5. Actividades de inversión", 0.0, r_inv[1], -r_inv[1], r_inv[2], safe_div(r_inv[1], r_inv[2]),
                   estado_techo(r_inv[1], r_inv[2])])
        s1.append(["Total CAPEX", 0.0, r_inv[1], -r_inv[1], r_inv[2], safe_div(r_inv[1], r_inv[2]),
                   estado_techo(r_inv[1], r_inv[2])])
        out["s1"] = {"ini": r_ini, "fun": r_fun, "oo": r_oo, "ap": r_ap, "inv": r_inv,
                     "opex": (opex_i, opex_e, opex_t), "rows": s1}
        out["sections"].append(dict(
            key="s1", title="1. Resumen por tipo de actividad",
            headers=["Tipo de actividad", "Ingresos (S/)", "Egresos (S/)", "Resultado (S/)", "Techo de egresos (S/)",
                     "% de uso del techo", "Estado"],
            formats=["text", "money", "money", "money", "money", "pct", "status"], rows=s1, totals={4, 6},
            note="El uso del techo total excluye las actividades propias, que deben autofinanciarse."))

        # ---------------- 2. Ingresos proyectados por categoría
        s2 = []
        for cat in CATEGORIAS_ING_RESUMEN:
            rows = [r for r in fi if norm_key(r["categoria"]) == norm_key(cat)]
            n = sum(1 for r in rows if to_float(r.get("cantidad")) is not None)
            s2.append([cat, n, sum(r["_c"]["bruto"] for r in rows), sum(r["_c"]["descuentos"] for r in rows),
                       sum(r["_c"]["neto"] for r in rows), None])
        oo_ing = r_oo[0]
        ap_ing = r_ap[0]
        s2.append(["Otras actividades operativas (PEAS, extracurriculares)",
                   sum(1 for m in oo if s(m["nombre"]) and m["_ingresos"] > 0), oo_ing, 0.0, oo_ing, None])
        s2.append(["Actividades propias (congresos, seminarios, talleres)",
                   sum(1 for m in ap if s(m["nombre"]) and m["_ingresos"] > 0), ap_ing, 0.0, ap_ing, None])
        tot_neto = sum(r[4] for r in s2)
        for r in s2:
            r[5] = safe_div(r[4], tot_neto)
        s2.append(["Total", sum(r[1] for r in s2), sum(r[2] for r in s2), sum(r[3] for r in s2), tot_neto,
                   1.0 if tot_neto else None])
        cats9 = {norm_key(c) for c in CATEGORIAS_ING_RESUMEN}
        sin_concepto = sum(r["_c"]["neto"] for r in fi if norm_key(r["categoria"]) not in cats9)
        s2.append(["Control: ingresos de funcionamiento sin concepto asignado (S/)", None, None, None, sin_concepto,
                   None])
        out["sections"].append(dict(
            key="s2", title="2. Ingresos proyectados",
            headers=["Categoría", "N° de actividades", "Ingreso bruto (S/)", "Becas, descuentos y morosidad (S/)",
                     "Ingreso neto (S/)", "% del total"],
            formats=["text", "int", "money", "money", "money", "pct"], rows=s2, totals={len(s2) - 2}))

        # ---------------- 3. Metas de ingresos y resultado
        meta_ing = u["meta_ingresos_func"] or 0.0
        meta_oo = u["meta_resultado_otras"] or 0.0
        meta_ap = u["meta_resultado_propias"] or 0.0
        ing_fun = r_fun[0]

        def estado_meta(valor, meta):
            if not meta:
                return "" if not valor else "Sin meta asignada"
            return "Alcanza la meta" if valor >= meta - 1e-9 else "Bajo la meta"

        res_ap = r_ap[0] - r_ap[1]
        resultado = opex_i - opex_e
        s3 = [
            ["Ingreso neto de funcionamiento", ing_fun, ""],
            ["Meta de ingresos de funcionamiento", meta_ing, ""],
            ["Cumplimiento de la meta de ingresos (%)", safe_div(ing_fun, meta_ing), estado_meta(ing_fun, meta_ing)],
            ["Ingresos de otras actividades operativas", oo_ing, ""],
            ["Meta de ingresos de otras actividades operativas", meta_oo, estado_meta(oo_ing, meta_oo)],
            ["Resultado de actividades propias", res_ap, ""],
            ["Meta de resultado de actividades propias", meta_ap, estado_meta(res_ap, meta_ap) if meta_ap else ""],
            ["Ingresos totales", opex_i, ""],
            ["Egresos totales", opex_e, ""],
            ["Resultado de la unidad (ingresos − egresos)", resultado,
             "" if not opex_i and not opex_e else ("Superávit" if resultado >= 0 else "Déficit")],
            ["Margen sobre ingresos (%)", safe_div(resultado, opex_i), ""],
        ]
        out["s3"] = {"resultado": resultado, "margen": safe_div(resultado, opex_i)}
        out["sections"].append(dict(
            key="s3", title="3. Metas de ingresos y resultado de la unidad",
            headers=["Indicador", "Monto (S/)", "Estado / referencia"],
            formats=["text", "money_or_pct", "status"], rows=s3, totals={9},
            note="Para unidades que no generan ingresos, el resultado equivale a su costo total (negativo)."))

        # ---------------- 4. Iniciativas por eje
        s4 = []
        tot_i4 = sum(m["_ingresos"] for m in ini)
        tot_e4 = sum(m["_egresos"] for m in ini)
        for e in self.db.catalogo("EJE"):
            ms = [m for m in ini if norm_key(m["eje"]) == norm_key(e["descripcion"])]
            i_ = sum(m["_ingresos"] for m in ms)
            e_ = sum(m["_egresos"] for m in ms)
            s4.append([f"{e['codigo']} {e['descripcion']}", sum(1 for m in ms if s(m["nombre"])), i_,
                       safe_div(i_, tot_i4), e_, safe_div(e_, tot_e4), i_ - e_])
        s4.append(["Total", sum(r[1] for r in s4), tot_i4, 1.0 if tot_i4 else None, tot_e4, 1.0 if tot_e4 else None,
                   tot_i4 - tot_e4])
        out["sections"].append(dict(
            key="s4", title="4. Iniciativas con impacto estratégico por eje",
            headers=["Eje estratégico", "N° de iniciativas", "Ingresos Total (S/)", "% del total ingresos",
                     "Costo total (S/)", "% del total costos", "Resultado (S/)"],
            formats=["text", "int", "money", "pct", "money", "pct", "money"], rows=s4, totals={len(s4) - 1}))

        # ---------------- 5. Egresos de funcionamiento por categoría
        tot5 = sum(r["_c"]["total"] for r in fe)
        s5 = []
        cats28 = {norm_key(c) for c in CATEGORIAS_EGR_RESUMEN}
        for cat in CATEGORIAS_EGR_RESUMEN:
            m_ = sum(r["_c"]["total"] for r in fe if norm_key(r["categoria"]) == norm_key(cat))
            s5.append([cat, m_, safe_div(m_, tot5)])
        otros = sum(r["_c"]["total"] for r in fe if norm_key(r["categoria"]) not in cats28)
        if abs(otros) > 1e-9:
            s5.append(["OTRAS CATEGORÍAS (fuera de las 28 / sin categoría)", otros, safe_div(otros, tot5)])
        s5.append(["Total", tot5, 1.0 if tot5 else None])
        out["sections"].append(dict(
            key="s5", title="5. Egresos de funcionamiento por categoría",
            headers=["Categoría", "Monto anual (S/)", "% del total"], formats=["text", "money", "pct"],
            rows=s5, totals={len(s5) - 1}))

        # ---------------- 6. Otras operativas por categoría
        s6 = []
        for c in self.db.catalogo("CAT_OO"):
            ms = [m for m in oo if norm_key(m["categoria"]) == norm_key(c["descripcion"])]
            s6.append([f"{c['codigo']} {c['descripcion']}", sum(m["_ingresos"] for m in ms), None,
                       sum(m["_egresos"] for m in ms), None, None])
        ti6, te6 = sum(r[1] for r in s6), sum(r[3] for r in s6)
        for r in s6:
            r[2], r[4], r[5] = safe_div(r[1], ti6), safe_div(r[3], te6), r[1] - r[3]
        s6.append(["Total", ti6, 1.0 if ti6 else None, te6, 1.0 if te6 else None, ti6 - te6])
        out["sections"].append(dict(
            key="s6", title="6. Otras actividades operativas por categoría",
            headers=["Categoría", "Ingresos Total (S/)", "% del total ingresos", "Costo total (S/)",
                     "% del total costos", "Resultado (S/)"],
            formats=["text", "money", "pct", "money", "pct", "money"], rows=s6, totals={len(s6) - 1}))

        # ---------------- 7. Actividades propias por tipo
        s7 = []
        for t in self.db.catalogo("TIPO_AP"):
            ms = [m for m in ap if norm_key(m["tipo"]) == norm_key(t["codigo"])]
            i_, e_ = sum(m["_ingresos"] for m in ms), sum(m["_egresos"] for m in ms)
            s7.append([t["codigo"], sum(1 for m in ms if s(m["nombre"])), sum(m["_participantes"] for m in ms),
                       i_, e_, i_ - e_, estado_autofin(i_, e_)])
        ti7, te7 = sum(r[3] for r in s7), sum(r[4] for r in s7)
        s7.append(["Total", sum(r[1] for r in s7), sum(r[2] for r in s7), ti7, te7, ti7 - te7,
                   estado_autofin(ti7, te7)])
        out["sections"].append(dict(
            key="s7", title="7. Actividades propias por tipo",
            headers=["Tipo", "N° de actividades", "Participantes", "Ingreso (S/)", "Costo (S/)", "Resultado (S/)",
                     "Estado"],
            formats=["text", "int", "int", "money", "money", "money", "status"], rows=s7, totals={len(s7) - 1}))

        # ---------------- 8. Inversión por tipo
        s8 = []
        tot8 = sum(m["_egresos"] for m in inv)
        for t in self.db.catalogo("TIPO_INV"):
            ms = [m for m in inv if norm_key(m["tipo_inversion"]) == norm_key(t["codigo"])]
            monto = sum(m["_egresos"] for m in ms)
            s8.append([t["codigo"], sum(1 for m in ms if m["_egresos"] > 0), monto, safe_div(monto, tot8),
                       sum(m["_egresos"] for m in ms if m["ficha_sustento"] == "Sí")])
        s8.append(["Total", sum(r[1] for r in s8), sum(r[2] for r in s8), 1.0 if tot8 else None,
                   sum(r[4] for r in s8)])
        out["sections"].append(dict(
            key="s8", title="8. Actividades de inversión por tipo",
            headers=["Tipo de inversión", "N° de líneas", "Monto (S/)", "% del total", "Con ficha de sustento (S/)"],
            formats=["text", "int", "money", "pct", "money"], rows=s8, totals={len(s8) - 1}))

        # ---------------- 9. Iniciativas e inversiones por prioridad
        s9 = []
        for pr in self.db.catalogo("PRIORIDAD"):
            nivel = int(pr["codigo"])
            ms = [m for m in ini if m["prioridad"] == nivel]
            vs = [m for m in inv if m["prioridad"] == nivel]
            i_, e_ = sum(m["_ingresos"] for m in ms), sum(m["_egresos"] for m in ms)
            s9.append([f"{nivel}", pr["descripcion"], sum(1 for m in ms if s(m["nombre"])), i_, e_, i_ - e_,
                       sum(1 for v in vs if v["_egresos"] > 0), sum(v["_egresos"] for v in vs)])
        s9.append(["Total", "", sum(r[2] for r in s9), sum(r[3] for r in s9), sum(r[4] for r in s9),
                   sum(r[5] for r in s9), sum(r[6] for r in s9), sum(r[7] for r in s9)])
        sin_pr_ini = sum(1 for m in ini if s(m["nombre"]) and not m["prioridad"])
        sin_pr_inv = sum(1 for m in inv if m["_egresos"] > 0 and not m["prioridad"])
        s9.append(["Sin prioridad", "Iniciativas / inversiones sin prioridad", sin_pr_ini, None, None, None,
                   sin_pr_inv, None])
        out["sections"].append(dict(
            key="s9", title="9. Iniciativas e inversiones por prioridad",
            headers=["Prioridad", "Categoría", "N° de iniciativas", "Ingresos Total (S/)", "Costo total (S/)",
                     "Total (S/)", "N° de inversiones", "Monto inversión (S/)"],
            formats=["text", "text", "int", "money", "money", "money", "int", "money"], rows=s9,
            totals={len(s9) - 2}))

        # ---------------- 10. Programación mensual
        def mens(items, tipo=None, key="meses"):
            tot = [0.0] * 12
            for it in items:
                if tipo and it.get("tipo") != tipo:
                    continue
                for i, v in enumerate(it["_c"][key]):
                    tot[i] += v
            return tot

        def det(masters):
            return [dd for m in masters for dd in m["_detalle"]]

        ing_f = mens(fi)
        ing_oo = mens(det(oo), "Ingresos")
        ing_ini = mens(det(ini), "Ingresos")
        ing_ap = mens(det(ap), "Ingresos")
        tot_ing = [sum(x) for x in zip(ing_f, ing_oo, ing_ini, ing_ap)]
        egr_ini = mens(det(ini), "Egresos")
        egr_f = mens(fe)
        egr_oo = mens(det(oo), "Egresos")
        egr_ap = mens(det(ap), "Egresos")
        egr_inv = mens(det(inv))
        tot_egr = [sum(x) for x in zip(egr_ini, egr_f, egr_oo, egr_ap, egr_inv)]
        flujo = [a - b for a, b in zip(tot_ing, tot_egr)]
        acum, eacum, ac, ea = [], [], 0.0, 0.0
        for f_, e_ in zip(flujo, tot_egr):
            ac += f_
            ea += e_
            acum.append(ac)
            eacum.append(ea)
        te = sum(tot_egr)
        pct_acum = [safe_div(x, te) for x in eacum]
        s10 = [
            ["Ingresos — funcionamiento", *ing_f, sum(ing_f)],
            ["Ingresos — otras actividades operativas", *ing_oo, sum(ing_oo)],
            ["Ingresos — iniciativas con impacto estratégico", *ing_ini, sum(ing_ini)],
            ["Ingresos — actividades propias", *ing_ap, sum(ing_ap)],
            ["Total ingresos", *tot_ing, sum(tot_ing)],
            ["Egresos — iniciativas con impacto estratégico", *egr_ini, sum(egr_ini)],
            ["Egresos — funcionamiento", *egr_f, sum(egr_f)],
            ["Egresos — otras actividades operativas", *egr_oo, sum(egr_oo)],
            ["Egresos — actividades propias", *egr_ap, sum(egr_ap)],
            ["Egresos — inversión", *egr_inv, sum(egr_inv)],
            ["Total egresos", *tot_egr, te],
            ["Flujo neto mensual (ingresos − egresos)", *flujo, sum(flujo)],
            ["Flujo neto acumulado", *acum, acum[-1]],
            ["Egresos acumulados", *eacum, eacum[-1]],
            ["% acumulado de egresos", *pct_acum, safe_div(eacum[-1], te)],
        ]
        ok10 = abs(round(te - (opex_e + r_inv[1]), 2)) < TOL and abs(round(sum(tot_ing) - opex_i, 2)) < TOL
        control10 = "OK" if ok10 else "Revisar"
        out["s10"] = {"rows": s10, "control": control10}
        out["sections"].append(dict(
            key="s10", title="10. Programación mensual de ingresos y egresos (S/)",
            headers=["Concepto", *MESES, "Total"], formats=["text"] + ["money"] * 13, rows=s10,
            totals={4, 10, 11}, pct_rows={14},
            note=f"Control: ingresos y egresos programados por mes vs. totales de la sección 1 → {control10}"))

        # ---------------- 11. Directrices
        dirs = self.directrices(uid, d, s1_rows=s1)
        out["directrices"] = dirs
        out["sections"].append(dict(
            key="s11", title="11. Verificación automática de directrices",
            headers=["Directriz", "Verificación", "N° de observaciones", "Resultado"],
            formats=["text", "text", "int", "status"],
            rows=[[x["codigo"], x["descripcion"], x["n"], x["resultado"]] for x in dirs], totals=set()))

        # ---------------- 12. Conformidad
        prm = p.raw or {}
        s12 = [
            ["Nombre", u["elaborado_nombre"] or u["responsable"], u["revisado_nombre"] or prm.get("director_dpdi", "")],
            ["Cargo", u["elaborado_cargo"] or u["cargo"], u["revisado_cargo"] or prm.get("cargo_dpdi", "")],
            ["Firma", "", ""],
            ["Fecha", u["elaborado_fecha"] or "", u["revisado_fecha"] or ""],
        ]
        out["sections"].append(dict(
            key="s12", title="12. Conformidad",
            headers=["", "Elaborado por (Responsable de la unidad)",
                     "Revisado por (Director de Planificación y Desarrollo Institucional - DPDI)"],
            formats=["text", "text", "text"], rows=s12, totals=set()))
        return out

    # ====================================================================
    # 11 DIRECTRICES DE LA DPDI
    # ====================================================================
    def directrices(self, uid: int, data: Optional[dict] = None, s1_rows: Optional[list] = None) -> list[dict]:
        d = data or self.load_unit_data(uid)
        ini, oo, ap, inv, fi, fe = d["INI"], d["OO"], d["AP"], d["INV"], d["fi"], d["fe"]
        flags: dict[tuple[str, int], set] = {}

        def flag(key, code):
            flags.setdefault(key, set()).add(code)

        res: list[dict] = []

        def add(code, desc, items):
            res.append({"codigo": code, "descripcion": desc, "n": len(items),
                        "resultado": "Cumple" if not items else "Revisar", "items": items})

        # DF-01
        items = []
        for m in ini:
            if s(m["nombre"]) and (not s(m["eje"]) or not s(m["obj_poa"]) or not s(m["oe"])):
                items.append(f"{m['codigo']} {m['nombre']}: falta eje / objetivo anual POA / OE")
                flag(("INI", m["id"]), "DF-01")
        add("DF-01", "Iniciativas sin eje estratégico, objetivo anual del POA u OE del PEI", items)

        # DF-02
        items = []
        for m in ini:
            if s(m["nombre"]) and (not s(m["indicador"]) or not s(m["responsable"]) or control_meta(m) != "OK"):
                items.append(f"{m['codigo']} {m['nombre']}: indicador/responsable/meta ({control_meta(m)})")
                flag(("INI", m["id"]), "DF-02")
        add("DF-02", "Iniciativas sin indicador, meta o responsable, o con programación trimestral (T1+T2+T3) "
                     "que no suma la meta", items)

        # DP-01
        if s1_rows is None:
            s1_rows = self.build_summary(uid, d)["s1"]["rows"]
        items = [f"{r[0]}: {r[6]}" for r in s1_rows[:6] if r[6] in ("Excede techo", "Sin techo")]
        add("DP-01", "Tipos de actividad que exceden su techo de egresos", items)

        # DP-02
        items = []
        for m in ini:
            if not s(m["nombre"]):
                for dd in m["_detalle"]:
                    if dd["_c"]["total"] > 0:
                        items.append(f"Detalle sin iniciativa válida ({m['codigo']}): {dd['partida']}")
                        flag(("INI-D", dd["id"]), "DP-02")
            elif m["_egresos"] == 0:
                items.append(f"{m['codigo']} {m['nombre']}: iniciativa sin gasto registrado")
                flag(("INI", m["id"]), "DP-02")
        add("DP-02", "Gasto de iniciativas sin código o con código inexistente, e iniciativas sin gasto registrado",
            items)

        # DP-03
        items = []
        for kind, masters in (("INI", ini), ("OO", oo), ("AP", ap), ("INV", inv)):
            for m in masters:
                for dd in m["_detalle"]:
                    if dd["_c"]["total"] > 0 and not s(dd["sustento"]):
                        items.append(f"{m['codigo']} · {dd['partida'] or dd['descripcion']}: sin sustento")
                        flag((f"{kind}-D", dd["id"]), "DP-03")
        for r in fe:
            if r["_c"]["total"] > 0 and not s(r["sustento"]):
                items.append(f"{r['codigo']} · {r['partida']}: sin sustento")
                flag(("FE", r["id"]), "DP-03")
        add("DP-03", "Egresos con monto y sin sustento (todas las hojas de egresos)", items)

        # DP-04
        items = []
        for kind, masters in (("INI", ini), ("OO", oo), ("AP", ap), ("INV", inv)):
            for m in masters:
                for dd in m["_detalle"]:
                    if dd["_c"]["control"] == "Revisar":
                        items.append(f"{m['codigo']} · {dd['partida'] or dd['descripcion']}: total "
                                     f"{fmt_money(dd['_c']['total'])} vs programado {fmt_money(dd['_c']['programado'])}")
                        flag((f"{kind}-D", dd["id"]), "DP-04")
        for r in fe:
            if to_float(r.get("costo_unit_moneda")) is not None and (r["_c"]["n_pagos"] == 0
                                                                      or r["_c"]["control"] == "Revisar"):
                motivo = "sin meses marcados" if r["_c"]["n_pagos"] == 0 else "factores no suman 100 %"
                items.append(f"{r['codigo']} · {r['partida']}: {motivo}")
                flag(("FE", r["id"]), "DP-04")
        add("DP-04", "Montos no distribuidos por mes, o conceptos con costo y sin meses marcados", items)

        # DP-05
        items = []
        for m in inv:
            if m["_egresos"] > 0 and (m["ficha_sustento"] != "Sí" or not s(m["tipo_inversion"]) or not s(m["proyecto"])):
                items.append(f"{m['codigo']} {m['proyecto']}: ficha={m['ficha_sustento'] or '—'}, "
                             f"tipo={m['tipo_inversion'] or '—'}")
                flag(("INV", m["id"]), "DP-05")
        add("DP-05", "Inversiones sin ficha de sustento adjunta, sin tipo o sin proyecto", items)

        # DI-01
        items = []
        for r in fi:
            if to_float(r.get("cantidad")) is not None and r["_c"]["tarifa_soles"] is not None and r["_c"]["n_meses"] == 0:
                items.append(f"{r['codigo']} · {r['partida']}: sin meses de cobro")
                flag(("FI", r["id"]), "DI-01")
        add("DI-01", "Ingresos de funcionamiento con cantidad y tarifa, sin meses de cobro", items)

        # DI-02
        items = []
        for r in fi:
            if r["_c"]["neto"] > 0 and (not s(r["partida"]) or not s(r["sustento"])):
                items.append(f"{r['codigo']} · {r['partida'] or '(sin concepto)'}: sin concepto o sustento")
                flag(("FI", r["id"]), "DI-02")
        add("DI-02", "Ingresos de funcionamiento sin concepto o sin sustento", items)

        # AP-01
        items = []
        for m in ap:
            if m["_estado"] == "Déficit":
                items.append(f"{m['codigo']} {m['nombre']}: resultado {fmt_money(m['_resultado'])}")
                flag(("AP", m["id"]), "AP-01")
        add("AP-01", "Actividades propias en déficit (no autofinanciadas)", items)

        # AP-02
        items = []
        for m in ap:
            if s(m["nombre"]) and (not s(m["tipo"]) or not s(m["mes_ejecucion"])):
                items.append(f"{m['codigo']} {m['nombre']}: sin tipo o mes de ejecución")
                flag(("AP", m["id"]), "AP-02")
        add("AP-02", "Actividades propias sin tipo o sin mes de ejecución", items)

        d["_flags"] = flags
        return res

    # ====================================================================
    # BASE DE DATOS CONSOLIDADA NIIF (1 fila por registro mensualizado)
    # ====================================================================
    BD_COLUMNS = ["Anio", "Codigo_Area", "Macrogrupo", "Unidad_Organica", "Centro_Costo", "Tipo_Actividad_POA",
                  "Codigo_Actividad", "Nombre_Actividad", "Eje_Estrategico", "OE_PEI", "AE_PEI", "OBJ_POA",
                  "Prioridad", "Tipo_Flujo", "Codigo_NIIF", "Partida_Rubro", "Partida_GG", "Categoria_Presupuestal",
                  "Descripcion_Item", "Moneda_Origen", "Cantidad", "Costo_Unitario_Soles", "Mes", "Mes_Numero",
                  "Monto_Mensual_Soles", "Sustento", "Responsable", "Estado_Directrices"]

    def consolidated_rows(self, uids: Iterable[int]) -> list[dict]:
        out = []
        for uid in uids:
            d = self.load_unit_data(uid)
            self.directrices(uid, d)
            flags = d["_flags"]
            u, anio = d["unidad"], d["params"].anio
            area = u["codigo_area"]

            def estado(*keys):
                codes = set()
                for k in keys:
                    codes |= flags.get(k, set())
                return "Cumple" if not codes else "Revisar: " + ", ".join(sorted(codes))

            def emit(tipo_act, codigo, nombre, tipo_flujo, partida, partida_gg, categoria, desc, moneda, cant, cu_s,
                     meses, sustento, responsable, est, eje="", oe="", ae="", obj="", prioridad=None):
                cat = (self.gg_info(partida, tipo_flujo) if partida_gg else None) or \
                    self.buscar_partida(partida, [tipo_flujo, "Egresos", "Ingresos", "Inversion"])
                for i, monto in enumerate(meses):
                    if abs(monto) < 1e-9:
                        continue
                    out.append({
                        "Anio": anio, "Codigo_Area": area, "Macrogrupo": u["macrogrupo"],
                        "Unidad_Organica": u["nombre_unidad"], "Centro_Costo": u["centro_costo"],
                        "Tipo_Actividad_POA": tipo_act, "Codigo_Actividad": f"{area}-{codigo}",
                        "Nombre_Actividad": nombre, "Eje_Estrategico": eje, "OE_PEI": oe, "AE_PEI": ae, "OBJ_POA": obj,
                        "Prioridad": prioridad, "Tipo_Flujo": tipo_flujo,
                        "Codigo_NIIF": cat["codigo_niif"] if cat else "",
                        "Partida_Rubro": cat["partida_rubro"] if cat else partida,
                        "Partida_GG": partida_gg or (cat["partida_gg"] if cat else "") or partida,
                        "Categoria_Presupuestal": categoria or (cat["categoria"] if cat else ""),
                        "Descripcion_Item": desc, "Moneda_Origen": moneda or "S/", "Cantidad": cant,
                        "Costo_Unitario_Soles": cu_s, "Mes": MESES[i], "Mes_Numero": i + 1,
                        "Monto_Mensual_Soles": round(monto, 2), "Sustento": sustento,
                        "Responsable": responsable or u["responsable"], "Estado_Directrices": est,
                    })

            for m in d["INI"]:
                eje = f"{self.eje_codigo(m['eje'])} {m['eje']}".strip()
                for dd in m["_detalle"]:
                    emit(KIND_LABEL["INI"], m["codigo"], m["nombre"], dd["tipo"], dd["partida"], dd["partida"],
                         dd["categoria"], dd["descripcion"], dd["moneda"], to_float(dd["cantidad"]),
                         dd["_c"]["costo_unit_soles"], dd["_c"]["meses"], dd["sustento"], m["responsable"],
                         estado(("INI", m["id"]), ("INI-D", dd["id"])), eje, m["oe"], "", m["obj_poa"],
                         m["prioridad"])
            for r in d["fi"]:
                emit(KIND_LABEL["FI"], r["codigo"], r["partida"], "Ingresos", r["partida"], "", r["categoria"],
                     r["observaciones"] or r["partida"], r["moneda"], to_float(r["cantidad"]), r["_c"]["tarifa_soles"],
                     r["_c"]["meses"], r["sustento"], "", estado(("FI", r["id"])))
            for r in d["fe"]:
                emit(KIND_LABEL["FE"], r["codigo"], r["partida"], "Egresos", r["partida"], "", r["categoria"],
                     r["observaciones"] or r["partida"], r["moneda"], to_float(r["cantidad"]),
                     r["_c"]["costo_unit_soles"], r["_c"]["meses"], r["sustento"], "", estado(("FE", r["id"])))
            for kind in ("OO", "AP", "INV"):
                for m in d[kind]:
                    nombre = m.get("nombre") or m.get("proyecto") or ""
                    for dd in m["_detalle"]:
                        emit(KIND_LABEL[kind], m["codigo"], nombre, dd["tipo"], dd["partida"], dd["partida"],
                             dd["categoria"], dd["descripcion"], dd["moneda"], to_float(dd["cantidad"]),
                             dd["_c"]["costo_unit_soles"], dd["_c"]["meses"], dd["sustento"], m.get("responsable", ""),
                             estado((kind, m["id"]), (f"{kind}-D", dd["id"])),
                             prioridad=m.get("prioridad") if kind == "INV" else None)
        return out

    # ====================================================================
    # IMPORTADOR DE FICHAS EXCEL
    # ====================================================================
    def import_ficha(self, path: str | Path, unidad_id: Optional[int] = None, reemplazar: bool = True,
                     importar_parametros: bool = False) -> dict:
        import openpyxl
        path = Path(path)
        wb = openpyxl.load_workbook(path, data_only=True)
        report = {"archivo": path.name, "unidad_id": None, "creada": False, "conteos": {}, "avisos": []}
        sheets = set(wb.sheetnames)
        requeridas = {"Ficha_Resumen", "Iniciativas", "Funcionamiento_Egr"}
        if not requeridas <= sheets:
            raise ValueError("El archivo no tiene la estructura de Ficha_Formulacion_POA (faltan hojas: "
                             + ", ".join(sorted(requeridas - sheets)) + ").")
        if importar_parametros and "Parámetros" in sheets:
            report["parametros"] = self.import_parametros_wb(wb)

        fr = wb["Ficha_Resumen"]
        nombre_unidad = s(fr["C6"].value)
        uid = unidad_id or self._match_unidad(nombre_unidad, s(fr["C7"].value), path)
        unidad_data = {
            "nombre_unidad": nombre_unidad or None, "centro_costo": s(fr["C7"].value) or None,
            "responsable": s(fr["C8"].value) or None, "cargo": s(fr["C9"].value) or None,
            "correo": s(fr["C10"].value) or None, "fecha_presentacion": s(fr["C12"].value) or None,
            "version_ficha": s(fr["C13"].value) or None,
            "elaborado_nombre": s(fr["C179"].value) or None, "elaborado_cargo": s(fr["C180"].value) or None,
            "elaborado_fecha": s(fr["C182"].value) or None, "revisado_nombre": s(fr["F179"].value) or None,
            "revisado_cargo": s(fr["F180"].value) or None, "revisado_fecha": s(fr["F182"].value) or None,
        }
        if "Parámetros" in sheets:
            pm = wb["Parámetros"]
            for key, cell in (("techo_iniciativas", "C11"), ("techo_funcionamiento", "C12"),
                              ("techo_otras_operativas", "C13"), ("techo_inversion", "C14"),
                              ("meta_ingresos_func", "C16"), ("meta_resultado_otras", "C17"),
                              ("meta_resultado_propias", "C18")):
                v = to_float(pm[cell].value)
                if v is not None:
                    unidad_data[key] = v
        if uid is None:
            area, macro = self._area_desde_archivo(path)
            base_area, n = area, 1
            while self.db.scalar("SELECT 1 FROM unidades_organicas WHERE codigo_area=?", (area,)):
                n += 1
                area = f"{base_area}{n}"
            new = {k: v for k, v in unidad_data.items() if v is not None}
            new.update(codigo_area=area, macrogrupo=macro, nombre_unidad=nombre_unidad or path.stem,
                       genera_ingresos=1 if macro != MACROGRUPOS[0] else 0)
            uid = self.save_unidad(new)
            report["creada"] = True
        else:
            self.db.update("unidades_organicas", uid, {k: v for k, v in unidad_data.items() if v is not None})
        report["unidad_id"] = uid

        if reemplazar:
            with self.db.transaction():
                for t in ("iniciativas", "otras_operativas", "act_propias", "inversiones", "funcionamiento_ing",
                          "funcionamiento_egr"):
                    self.db.execute(f"DELETE FROM {t} WHERE unidad_id=? AND anio=?", (uid, self.anio))

        c = report["conteos"]
        c["INI"], c["INI_det"] = self._imp_master_detail(
            wb, uid, "INI", "Iniciativas", "Detalle_Iniciativas", report)
        c["OO"], c["OO_det"] = self._imp_master_detail(
            wb, uid, "OO", "Otras_Operat", "Detalle_Otras_Operativas", report)
        c["AP"], c["AP_det"] = self._imp_master_detail(wb, uid, "AP", "Act_Prop", "Detalle_Act_Prop", report)
        c["AP_legacy"] = self._imp_act_propias_legacy(wb, uid, report)
        c["INV"], c["INV_det"] = self._imp_master_detail(
            wb, uid, "INV", "Inversion", "Detalle_Inversiones", report)
        c["FI"] = self._imp_flat(wb, uid, "FI", "Funcionamiento_Ing")
        c["FE"] = self._imp_flat(wb, uid, "FE", "Funcionamiento_Egr")
        return report

    def _match_unidad(self, nombre: str, cc: str, path: Path) -> Optional[int]:
        for u in self.unidades():
            if nombre and norm_key(u["nombre_unidad"]) == norm_key(nombre):
                return u["id"]
        for u in self.unidades():
            if cc and norm_key(u["centro_costo"]) == norm_key(cc):
                return u["id"]
        area, _ = self._area_desde_archivo(path)
        r = self.db.query_one("SELECT id FROM unidades_organicas WHERE codigo_area = ?", (area,))
        return r["id"] if r else None

    @staticmethod
    def _area_desde_archivo(path: Path) -> tuple[str, str]:
        stem = re.sub(r"^[0-9a-f]{6,}-", "", path.stem)
        m = re.search(r"macrogrupo_(.+)$", stem, re.I)
        tail = m.group(1) if m else stem
        macro = MACROGRUPOS[0]
        low = tail.lower()
        if low.startswith("pregrado"):
            macro, tail = "Pregrado", tail[len("pregrado"):]
        elif low.startswith("posgrado"):
            macro, tail = "Posgrado", tail[len("posgrado"):]
        elif low.startswith("educ"):
            macro = "Educación Continua"
            tail = re.sub(r"^educ[a-z]*_?cont[a-z]*", "", tail, flags=re.I)
        elif low.startswith("sop"):
            tail = re.sub(r"^sop_acad_adm", "", tail, flags=re.I)
        parts = [p for p in re.split(r"[_\-\s]+", tail) if p]
        if not parts:
            area = "UO"
        elif len(parts) == 1 or (len(parts[-1]) <= 4 and parts[-1].isupper()):
            area = parts[-1].upper()
        else:
            area = "".join(p[0] for p in parts).upper() + ("-PRE" if macro == "Pregrado" else
                                                            "-POS" if macro == "Posgrado" else "")
        return area[:12], macro

    @staticmethod
    def _header_map(ws) -> tuple[int, dict[str, int], list[int]]:
        """Ubica la fila de encabezados («Código…» en la columna A) y mapea nombre → columna."""
        hdr_row = None
        for r in range(1, 15):
            v = s(ws.cell(r, 1).value)
            if v.startswith("Código"):
                hdr_row = r
                break
        if hdr_row is None:
            raise ValueError(f"No se encontró la fila de encabezados en la hoja {ws.title}")
        cols: dict[str, int] = {}
        months: list[int] = []
        for c in range(1, ws.max_column + 1):
            h = s(ws.cell(hdr_row, c).value)
            if not h:
                continue
            key = norm_key(h)
            if h in MESES and len(months) < 12:
                months.append(c)
            cols.setdefault(key, c)
        return hdr_row, cols, months

    @staticmethod
    def _data_rows(ws, hdr_row: int) -> Iterable[int]:
        for r in range(hdr_row + 1, ws.max_row + 1):
            a = s(ws.cell(r, 1).value)
            if a.upper().startswith("TOTAL"):
                break
            if a.upper() == "EJEMPLO":
                continue
            yield r

    @staticmethod
    def _col(cols: dict[str, int], *names: str) -> Optional[int]:
        for n in names:
            k = norm_key(n)
            if k in cols:
                return cols[k]
        for n in names:
            k = norm_key(n)
            for key, c in cols.items():
                if key.startswith(k):
                    return c
        return None

    def _imp_master_detail(self, wb, uid: int, kind: str, sh_m: str, sh_d: str, report: dict) -> tuple[int, int]:
        if sh_m not in wb.sheetnames:
            return 0, 0
        ws = wb[sh_m]
        hdr, cols, _ = self._header_map(ws)
        g = lambda r, *n: (lambda c: ws.cell(r, c).value if c else None)(self._col(cols, *n))  # noqa: E731
        ids: dict[str, int] = {}
        n_m = 0
        for r in self._data_rows(ws, hdr):
            code = s(ws.cell(r, 1).value).upper()
            if kind == "INI":
                data = dict(codigo=code, eje=s(g(r, "Eje estratégico")), nombre=s(g(r, "Nombre de la iniciativa")),
                            descripcion=s(g(r, "Descripción y justificación")), origen=s(g(r, "Origen")),
                            directriz=s(g(r, "Directriz que atiende")), obj_poa=s(g(r, "Objetivo anual del POA")),
                            oe=s(g(r, "OE del PEI")), ae=s(g(r, "AE del PEI")), indicador=s(g(r, "Indicador")),
                            unidad_medida=s(g(r, "Unidad de medida")), linea_base=s(g(r, "Línea base")),
                            meta_anual=g(r, "Meta anual"), t1=g(r, "T1"), t2=g(r, "T2"), t3=g(r, "T3"),
                            responsable=s(g(r, "Responsable")), prioridad=s(g(r, "Prioridad")),
                            observaciones=s(g(r, "Observaciones de la DPDI", "Observaciones")))
                keep = bool(data["nombre"])
            elif kind == "OO":
                data = dict(codigo=code, categoria=s(g(r, "Categoría")), nombre=s(g(r, "Nombre de la actividad")),
                            descripcion=s(g(r, "Descripción y justificación")), responsable=s(g(r, "Responsable")),
                            observaciones=s(g(r, "Observaciones")))
                keep = bool(data["nombre"])
            elif kind == "AP":
                data = dict(codigo=code, tipo=s(g(r, "Tipo")), nombre=s(g(r, "Nombre de la actividad")),
                            participantes=g(r, "Participante proyectados", "Participantes proyectados"),
                            responsable=s(g(r, "Responsable")), observaciones=s(g(r, "Observaciones")))
                keep = bool(data["nombre"])
            else:
                data = dict(codigo=code, tipo_inversion=s(g(r, "Tipo de inversión")),
                            proyecto=s(g(r, "Proyecto de inversión")), prioridad=s(g(r, "Prioridad")),
                            ficha_sustento=s(g(r, "Ficha de sustento adjunta")) or "No",
                            observaciones=s(g(r, "Observaciones")))
                keep = bool(data["proyecto"] or data["tipo_inversion"])
            if not keep or not code:
                continue
            if code in ids:
                report["avisos"].append(f"{sh_m}: código duplicado {code} (se conserva la primera fila)")
                continue
            ids[code] = self.save_master(kind, uid, data)
            n_m += 1
        n_d = 0
        if sh_d in wb.sheetnames:
            wd = wb[sh_d]
            hdr, cols, months = self._header_map(wd)
            gd = lambda r, *n: (lambda c: wd.cell(r, c).value if c else None)(self._col(cols, *n))  # noqa: E731
            for r in self._data_rows(wd, hdr):
                code = s(wd.cell(r, 1).value).upper()
                partida = s(gd(r, "Partida / rubro GG")) or s(gd(r, "Partida", "Partida / rubro"))
                cant = gd(r, "Cantidad")
                cu = gd(r, "Costo unitario (moneda)", "Ingresoss o Costo unitario (moneda)",
                        "Ingresos o Costo unitario (moneda)", "Ingreso o costo unitario (moneda)")
                meses = [wd.cell(r, c).value for c in months]
                if not (partida or to_float(cu) or any(to_float(m) for m in meses)):
                    continue
                if not code:
                    report["avisos"].append(f"{sh_d} fila {r}: detalle sin código (omitido)")
                    continue
                if code not in ids:
                    placeholder = {"codigo": code}
                    if kind == "INV":
                        placeholder["proyecto"] = ""
                    ids[code] = self.save_master(kind, uid, placeholder)
                    report["avisos"].append(f"{sh_d} fila {r}: código {code} no existe en {sh_m} "
                                            "(se creó un registro vacío para conservar el gasto)")
                tipo_raw = s(gd(r, "Tipo")).capitalize()
                tipo = "Inversion" if kind == "INV" else ("Ingresos" if tipo_raw.startswith("Ingreso") else "Egresos")
                data = dict(tipo=tipo, partida=partida,
                            categoria=s(gd(r, "Categoría")), descripcion=s(gd(r, "Descripción del bien o servicio")),
                            cantidad=cant, moneda=s(gd(r, "Moneda")) or "S/", costo_unit_moneda=cu,
                            sustento=s(gd(r, "Sustento")), observaciones=s(gd(r, "Observaciones")))
                for i, v in enumerate(meses[:12]):
                    data[MCOLS[i]] = v
                self.save_detalle(kind, ids[code], data)
                n_d += 1
        if kind == "AP":
            # Act_Prop no tiene «Mes de ejecución»: se infiere del primer mes con ingresos (o egresos) del detalle.
            for code, mid in ids.items():
                m = self.db.get("act_propias", mid)
                if m and not s(m["mes_ejecucion"]):
                    dets = self.list_detalle("AP", mid)
                    for tipo in ("Ingresos", "Egresos"):
                        meses = [sum(dd["_c"]["meses"][i] for dd in dets if dd["tipo"] == tipo) for i in range(12)]
                        if any(meses):
                            mes = MESES[next(i for i, v in enumerate(meses) if v)]
                            self.db.update("act_propias", mid, {"mes_ejecucion": mes})
                            break
        return n_m, n_d

    def _imp_act_propias_legacy(self, wb, uid: int, report: dict) -> int:
        """Hoja oculta Act_Propias (formato anterior): completa datos del maestro o crea actividades."""
        if "Act_Propias" not in wb.sheetnames:
            return 0
        ws = wb["Act_Propias"]
        hdr, cols, _ = self._header_map(ws)
        g = lambda r, *n: (lambda c: ws.cell(r, c).value if c else None)(self._col(cols, *n))  # noqa: E731
        n = 0
        existentes = {m["codigo"]: m for m in self.db.all("act_propias", "unidad_id=? AND anio=?", (uid, self.anio))}
        for r in self._data_rows(ws, hdr):
            code = s(ws.cell(r, 1).value).upper()
            nombre = s(g(r, "Nombre de la actividad"))
            if not nombre or not code:
                continue
            master = dict(codigo=code, tipo=s(g(r, "Tipo")), nombre=nombre,
                          publico_objetivo=s(g(r, "Público objetivo")), mes_ejecucion=s(g(r, "Mes de ejecución")),
                          participantes=g(r, "Participantes proyectados"), tarifa=g(r, "Tarifa por participante"),
                          becas_pct=g(r, "Becas y descuentos"), responsable=s(g(r, "Responsable")),
                          sustento=s(g(r, "Sustento")), observaciones=s(g(r, "Observaciones")))
            if code in existentes:
                old = existentes[code]
                merged = {k: (v if (old.get(k) in (None, "") and v not in (None, "")) else old.get(k))
                          for k, v in master.items()}
                self.save_master("AP", uid, merged, old["id"])
                continue
            mid = self.save_master("AP", uid, master)
            n += 1
            mes = MESES.index(master["mes_ejecucion"]) + 1 if master["mes_ejecucion"] in MESES else 1
            try:
                self.crear_ingreso_desde_tarifa(mid)
            except ValueError:
                pass
            for col_name, desc in (("Ponentes y honorarios", "Ponentes y honorarios"),
                                   ("Local y logística", "Local y logística"), ("Materiales", "Materiales"),
                                   ("Difusión", "Difusión"), ("Otros costos", "Otros costos")):
                monto = to_float(g(r, col_name))
                if monto:
                    dd = {"tipo": "Egresos", "partida": "", "descripcion": desc, "cantidad": 1, "moneda": "S/",
                          "costo_unit_moneda": monto, "sustento": master["sustento"], MCOLS[mes - 1]: monto}
                    self.save_detalle("AP", mid, dd)
            report["avisos"].append(f"Act_Propias: {code} importada desde el formato anterior; asigne partidas NIIF "
                                    "a los costos directos.")
        return n

    def _imp_flat(self, wb, uid: int, kind: str, sheet: str) -> int:
        if sheet not in wb.sheetnames:
            return 0
        ws = wb[sheet]
        hdr, cols, months = self._header_map(ws)
        g = lambda r, *n: (lambda c: ws.cell(r, c).value if c else None)(self._col(cols, *n))  # noqa: E731
        n = 0
        for r in self._data_rows(ws, hdr):
            partida = s(g(r, "Partida / rubro"))
            cant = g(r, "Cantidad proyectada", "Cantidad")
            if kind == "FI":
                monto = g(r, "Tarifa por cobro (moneda)")
            else:
                monto = g(r, "Costo unitario por pago (moneda)")
            marks = [ws.cell(r, c).value for c in months]
            if not (partida or to_float(monto) or to_float(cant)):
                continue
            data = dict(codigo=self.next_code(kind, uid), partida=partida, categoria=s(g(r, "Categoría")),
                        cantidad=cant, moneda=s(g(r, "Moneda")) or "S/", sustento=s(g(r, "Sustento")),
                        observaciones=s(g(r, "Observaciones")))
            if kind == "FI":
                data.update(tarifa_moneda=monto, becas_pct=g(r, "Becas y descuentos"),
                            aplica_morosidad=s(g(r, "¿Aplica morosidad?")) or "No")
            else:
                data.update(costo_unit_moneda=monto, aplica_reajuste=s(g(r, "¿Aplica reajuste?")) or "No")
            for i, v in enumerate(marks[:12]):
                data[MCOLS[i]] = v
            self.save_flat(kind, uid, data)
            n += 1
        return n

    def _upsert_gg(self, nombre: str, descripcion: str, tipo: str, codigo: str = "") -> int:
        """Integra una Partida GG en la base consolidada (Cuentas NIIF + Partidas GG)."""
        hit = self.db.query_one("SELECT id FROM catalogo_niif_partidas WHERE partida_gg=? AND tipo_flujo=? ORDER BY id",
                                (nombre, tipo))
        if hit:
            if descripcion:
                self.db.execute("UPDATE catalogo_niif_partidas SET descripcion=? WHERE partida_gg=? AND tipo_flujo=?",
                                (descripcion, nombre, tipo))
                self.db.conn.commit()
            return 1
        cuenta = self.db.query_one("SELECT id FROM catalogo_niif_partidas WHERE tipo_flujo=? AND "
                                   "(codigo_niif=? OR UPPER(partida_rubro)=UPPER(?)) ORDER BY id",
                                   (tipo, codigo or "#", nombre))
        if cuenta:
            self.db.update("catalogo_niif_partidas", cuenta["id"], {"partida_gg": nombre, "descripcion": descripcion})
        else:
            self.db.insert("catalogo_niif_partidas", {"codigo_niif": codigo or "SIN-CODIGO", "partida_rubro": nombre,
                                                      "partida_gg": nombre, "descripcion": descripcion,
                                                      "categoria": "INVERSIÓN" if tipo == "Inversion" else "",
                                                      "tipo_flujo": tipo})
        self.invalidate_cache()
        return 1

    # ------------------------------------- plantilla de parámetros y catálogos
    def import_plantilla_parametros(self, path: str | Path) -> dict:
        """Lee la plantilla «Parámetros y Catálogos» (descargada desde el sistema) y actualiza la base."""
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        faltan = [sh for sh in PLANTILLA_PARAM if sh not in wb.sheetnames]
        if faltan:
            raise ValueError("El archivo no es la plantilla de Parámetros y Catálogos (faltan hojas: "
                             + ", ".join(faltan) + ").")
        rep: dict[str, Any] = {"avisos": []}

        def tabla(sheet: str) -> list[dict]:
            ws = wb[sheet]
            spec = PLANTILLA_PARAM[sheet]
            hdr = next((r for r in range(1, 12) if s(ws.cell(r, 1).value) in ("ID", "Clave")), None)
            if hdr is None:
                raise ValueError(f"No se encontró el encabezado en la hoja {sheet}.")
            pos = {norm_key(ws.cell(hdr, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(hdr, c).value}
            out = []
            for r in range(hdr + 1, ws.max_row + 1):
                row = {}
                for key, header in spec:
                    c = pos.get(norm_key(header))
                    row[key] = ws.cell(r, c).value if c else None
                if any(v not in (None, "") for k, v in row.items() if k not in ("id", "clave")):
                    row["_fila"] = r
                    out.append(row)
            return out

        def si(v) -> int:
            return 1 if s(v).lower() in ("sí", "si", "1", "true", "x", "s") else 0

        def rid(table: str, value) -> Optional[int]:
            v = to_float(value)
            if v is None:
                return None
            return int(v) if self.db.get(table, int(v)) else None

        with self.db.transaction():
            # --- Datos generales y supuestos
            gen = {}
            for row in tabla("Datos_Generales"):
                key = s(row["clave"])
                if key not in {k for k, _l, _n in PARAM_GENERALES}:
                    continue
                v = row["valor"]
                if key in PARAM_PCT:
                    gen[key] = parse_pct(v)
                elif key in PARAM_NUM:
                    gen[key] = to_float(v)
                elif key == "mes_inicio_reajuste":
                    gen[key] = int(to_float(v, 1) or 1)
                else:
                    gen[key] = s(v)
            if gen:
                self.db.update("parametros_globales", 1, gen, commit=False)
            rep["Datos generales"] = len(gen)

            # --- Unidades y techos
            n = 0
            for row in tabla("Unidades_Techos"):
                area = s(row["codigo_area"]).upper()
                if not area:
                    continue
                data = {k: row[k] for k, _ in PLANTILLA_PARAM["Unidades_Techos"] if k not in ("id",)}
                data["codigo_area"] = area
                data["genera_ingresos"] = si(row["genera_ingresos"])
                data["macrogrupo"] = s(row["macrogrupo"]) or MACROGRUPOS[0]
                for k in list(data):
                    if k.startswith(("techo_", "meta_")):
                        data[k] = to_float(data[k], 0.0) or 0.0
                    elif isinstance(data[k], str) or data[k] is None:
                        data[k] = s(data[k])
                if not data["nombre_unidad"]:
                    rep["avisos"].append(f"Unidades_Techos fila {row['_fila']}: falta el nombre de la unidad.")
                    continue
                uid = rid("unidades_organicas", row["id"])
                if uid is None:
                    old = self.db.query_one("SELECT id FROM unidades_organicas WHERE codigo_area=?", (area,))
                    uid = old["id"] if old else None
                if uid:
                    self.db.update("unidades_organicas", uid, data, commit=False)
                else:
                    self.db.insert("unidades_organicas", data, commit=False)
                n += 1
            rep["Unidades y techos"] = n

            # --- Cuentas NIIF + Partidas GG (base consolidada)
            n = 0
            for row in tabla("Cuentas_NIIF_Partidas_GG"):
                tipo = s(row["tipo_flujo"]).capitalize()
                tipo = "Inversion" if tipo.startswith("Inver") else tipo
                cuenta, gg = s(row["partida_rubro"]), s(row["partida_gg"])
                if tipo not in ("Ingresos", "Egresos", "Inversion") or not (cuenta or gg):
                    rep["avisos"].append(f"Cuentas_NIIF_Partidas_GG fila {row['_fila']}: falta Tipo o Partida/Cuenta.")
                    continue
                data = {"tipo_flujo": tipo, "partida_rubro": cuenta or gg.upper(), "partida_gg": gg,
                        "descripcion": s(row["descripcion"]), "codigo_niif": s(row["codigo_niif"]) or "SIN-CODIGO",
                        "categoria": s(row["categoria"]), "naturaleza": s(row["naturaleza"]),
                        "resultado_funcion": s(row["resultado_funcion"]),
                        "activo": 1 if row["activo"] in (None, "") else si(row["activo"])}
                cid = rid("catalogo_niif_partidas", row["id"])
                if cid is None:
                    old = self.db.query_one("SELECT id FROM catalogo_niif_partidas WHERE codigo_niif=? AND "
                                            "partida_rubro=? AND tipo_flujo=?",
                                            (data["codigo_niif"], data["partida_rubro"], tipo))
                    cid = old["id"] if old else None
                if cid:
                    self.db.update("catalogo_niif_partidas", cid, data, commit=False)
                else:
                    self.db.insert("catalogo_niif_partidas", data, commit=False)
                n += 1
            rep["Cuentas NIIF + Partidas GG"] = n

            # --- Catálogos estratégicos (relación Eje → Objetivo anual POA → OE del PEI)
            n = 0
            tipos_inv = {norm_key(v): k for k, v in CATALOGOS_ESTRATEGICOS.items()}
            filas = tabla("Catalogos_Estrategicos")
            orden_tipo = {"EJE": 0, "OE": 1}
            filas.sort(key=lambda r: orden_tipo.get(tipos_inv.get(norm_key(r["tipo_catalogo"]), ""), 2))
            for row in filas:
                tipo = tipos_inv.get(norm_key(row["tipo_catalogo"])) or s(row["tipo_catalogo"]).upper()
                code = s(row["codigo"])
                if tipo not in CATALOGOS_ESTRATEGICOS or not code:
                    rep["avisos"].append(f"Catalogos_Estrategicos fila {row['_fila']}: tipo de catálogo o código no válido.")
                    continue
                data = {"codigo": code, "descripcion": s(row["descripcion"]) or code, "extra": s(row["extra"]),
                        "orden": int(to_float(row["orden"], 0) or 0), "padre": "", "eje": ""}
                if tipo == "OBJ":
                    eje = self.eje_ref(s(row["eje"]))
                    oe = code_display(s(row["padre"]))
                    data["eje"] = eje["codigo"] if eje else s(row["eje"])
                    data["padre"] = oe
                    if not eje:
                        rep["avisos"].append(f"Objetivo {code}: eje «{s(row['eje'])}» no existe en el catálogo.")
                    if not self.db.query_one("SELECT 1 FROM catalogos_estrategicos WHERE tipo_catalogo='OE' AND codigo=?",
                                             (oe,)):
                        rep["avisos"].append(f"Objetivo {code}: OE «{oe}» no existe en el catálogo.")
                old = self.db.query_one("SELECT id FROM catalogos_estrategicos WHERE tipo_catalogo=? AND codigo=?",
                                        (tipo, code))
                cid = rid("catalogos_estrategicos", row["id"]) or (old["id"] if old else None)
                self.save_catalogo(tipo, data, cid)
                n += 1
            rep["Catálogos estratégicos"] = n

            # --- Costos de referencia
            n = 0
            for row in tabla("Costos_Referencia"):
                if not s(row["concepto"]):
                    continue
                data = {"concepto": s(row["concepto"]), "unidad_medida": s(row["unidad_medida"]),
                        "costo_unitario": to_float(row["costo_unitario"]), "fuente": s(row["fuente"])}
                old = self.db.query_one("SELECT id FROM costos_referencia WHERE concepto=?", (data["concepto"],))
                cid = rid("costos_referencia", row["id"]) or (old["id"] if old else None)
                self.db.upsert("costos_referencia", data, cid)
                n += 1
            rep["Costos de referencia"] = n

            # --- Plantilla de funcionamiento
            n = 0
            areas = {u["codigo_area"]: u["id"] for u in self.unidades()}
            for row in tabla("Plantilla_Funcionamiento"):
                if not s(row["partida"]):
                    continue
                ambito = s(row["ambito"]).upper() or "GLOBAL"
                base = s(row["base_calculo"]).upper() or "MONTO_FIJO"
                if base not in BASES_CALCULO:
                    rep["avisos"].append(f"Plantilla_Funcionamiento fila {row['_fila']}: base «{base}» no válida.")
                    continue
                if ambito != "GLOBAL" and ambito not in areas:
                    rep["avisos"].append(f"Plantilla_Funcionamiento fila {row['_fila']}: el área {ambito} no existe.")
                    continue
                data = {"unidad_id": None if ambito == "GLOBAL" else areas[ambito],
                        "macrogrupo": s(row["macrogrupo"]) or "*", "solo_genera_ingresos": si(row["solo_genera_ingresos"]),
                        "tipo_flujo": "Ingresos" if s(row["tipo_flujo"]).startswith("Ingreso") else "Egresos",
                        "partida": s(row["partida"]), "base_calculo": base,
                        "porcentaje": to_float(row["porcentaje"], 1.0), "monto_fijo": to_float(row["monto_fijo"], 0.0),
                        "aplica_ajuste": "Sí" if si(row["aplica_ajuste"]) else "No", "sustento": s(row["sustento"]),
                        "activo": 1 if row["activo"] in (None, "") else si(row["activo"]),
                        "orden": int(to_float(row["orden"], 0) or 0)}
                self.db.upsert("plantilla_funcionamiento", data, rid("plantilla_funcionamiento", row["id"]))
                n += 1
            rep["Plantilla de funcionamiento"] = n
        self.invalidate_cache()
        return rep

    # ------------------------------------------- carga masiva de cuentas NIIF
    def carga_masiva_niif(self, path: str | Path) -> dict:
        """Carga masiva del índice de cuentas contables NIIF (.xlsx/.csv) con Pandas.

        Columnas reconocidas (formato 1_index_codigos_cc.xlsx): CÓDIGO DE CUENTA CONTABLE, CUENTA CONTABLE,
        LÍNEA DE COSTO (categoría), NATURALEZA, RESULTADO POR FUNCIÓN, TIPO_INVERSION; o bien las columnas
        del sistema: codigo_niif, partida_rubro, partida_gg, categoria, tipo_flujo, naturaleza, resultado_funcion.
        """
        import pandas as pd
        path = Path(path)
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        else:
            df = pd.read_excel(path, dtype=str)
        df.columns = [norm_key(c) for c in df.columns]
        alias = {
            "codigo_niif": ["CODIGO DE CUENTA CONTABLE", "CODIGO NIIF", "CODIGO NIIF", "CODIGO"],
            "partida_rubro": ["CUENTA CONTABLE", "PARTIDA RUBRO", "PARTIDA"],
            "categoria": ["LINEA DE COSTO", "CATEGORIA"],
            "naturaleza": ["NATURALEZA"],
            "resultado_funcion": ["RESULTADO POR FUNCION", "RESULTADO FUNCION"],
            "tipo_inversion": ["TIPO INVERSION"],
            "partida_gg": ["PARTIDA GG", "PARTIDA RUBRO GG"],
            "tipo_flujo": ["TIPO FLUJO", "TIPO"],
        }
        cols = {}
        for key, names in alias.items():
            for n in names:
                if norm_key(n) in df.columns:
                    cols[key] = norm_key(n)
                    break
        if "codigo_niif" not in cols or "partida_rubro" not in cols:
            raise ValueError("No se encontraron las columnas de código y nombre de la cuenta contable.")
        df = df.dropna(subset=[cols["codigo_niif"], cols["partida_rubro"]])
        rep = {"insertadas": 0, "actualizadas": 0, "omitidas": 0}
        with self.db.transaction():
            for _, row in df.iterrows():
                code = s(row[cols["codigo_niif"]])
                name = s(row[cols["partida_rubro"]])
                if not code or not name:
                    rep["omitidas"] += 1
                    continue
                get = lambda k: s(row[cols[k]]) if k in cols and not pd.isna(row[cols[k]]) else ""  # noqa: E731
                tipo = get("tipo_flujo")
                if tipo not in ("Ingresos", "Egresos", "Inversion"):
                    tipo = "Ingresos" if code.startswith("7") else ("Inversion" if code.startswith("3") else "Egresos")
                categoria = get("categoria") or get("tipo_inversion") or ("INVERSIÓN" if tipo == "Inversion" else "")
                data = {"codigo_niif": code, "partida_rubro": name, "categoria": categoria.strip(),
                        "tipo_flujo": tipo, "naturaleza": get("naturaleza") or ("ACTIVO FIJO" if tipo == "Inversion" else ""),
                        "resultado_funcion": get("resultado_funcion") or ("INVERSIÓN" if tipo == "Inversion" else "")}
                if get("partida_gg"):
                    data["partida_gg"] = get("partida_gg")
                old = self.db.query_one("SELECT id FROM catalogo_niif_partidas WHERE codigo_niif=? AND partida_rubro=? "
                                        "AND tipo_flujo=?", (code, name, tipo))
                if old:
                    self.db.update("catalogo_niif_partidas", old["id"], data, commit=False)
                    rep["actualizadas"] += 1
                else:
                    self.db.insert("catalogo_niif_partidas", data, commit=False)
                    rep["insertadas"] += 1
        self.invalidate_cache()
        return rep

    # --------------------------------------------------- importar parámetros
    def import_parametros(self, path: str | Path) -> dict:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        if "Parámetros" not in wb.sheetnames:
            raise ValueError("El archivo no contiene la hoja «Parámetros».")
        return self.import_parametros_wb(wb)

    def import_parametros_wb(self, wb) -> dict:
        ws = wb["Parámetros"]
        rep: dict[str, int] = {}
        gen = {}
        for key, cell, conv in (("anio_formulacion", "C5", s), ("fecha_emision", "C6", s), ("fecha_limite", "C7", s),
                                ("inflacion_proyectada", "C21", to_float), ("tipo_cambio", "C22", to_float),
                                ("valor_uit", "C23", to_float), ("rmv", "C24", to_float),
                                ("reajuste_remunerativo_pct", "C25", to_float), ("mes_inicio_reajuste", "C26", to_float),
                                ("tasa_morosidad_pct", "C27", to_float), ("estudiantes_pregrado", "C28", to_float),
                                ("estudiantes_posgrado", "C29", to_float), ("pension_ref_pregrado", "C30", to_float),
                                ("pension_ref_posgrado", "C31", to_float), ("cuotas_periodo", "C32", to_float)):
            v = conv(ws[cell].value)
            if v not in (None, ""):
                gen[key] = int(v) if key == "mes_inicio_reajuste" else v
        if gen:
            self.db.save_parametros(gen)
        rep["parametros_generales"] = len(gen)

        sections: dict[str, int] = {}
        for r in range(1, ws.max_row + 1):
            v = s(ws.cell(r, 2).value)
            m = re.match(r"^([A-S])\.\s", v)
            if m:
                sections[m.group(1)] = r

        def rows_of(letter):
            start = sections.get(letter)
            if not start:
                return []
            out = []
            for r in range(start + 2, ws.max_row + 1):
                b = ws.cell(r, 2).value
                if b is None and ws.cell(r, 3).value is None and ws.cell(r, 4).value is None:
                    if out or r > start + 40:
                        break
                    continue
                if re.match(r"^[A-S]\.\s", s(b)):
                    break
                out.append((s(b), s(ws.cell(r, 3).value), s(ws.cell(r, 4).value), s(ws.cell(r, 5).value)))
            return out

        def upsert_cat(tipo, code, desc, padre="", extra="", orden=0, eje=""):
            if not code:
                return 0
            old = self.db.query_one("SELECT id, descripcion FROM catalogos_estrategicos WHERE tipo_catalogo=? AND codigo=?",
                                    (tipo, code))
            data = {"codigo": code, "orden": orden}
            if desc:
                data["descripcion"] = desc
            if padre:
                data["padre"] = padre
            if eje:
                data["eje"] = eje
            if extra:
                data["extra"] = extra
            self.save_catalogo(tipo, data, old["id"] if old else None)
            return 1

        def eje_code(value):
            ref = self.eje_ref(value) if value else None
            return ref["codigo"] if ref else s(value)

        # La sección F (acciones estratégicas del PEI) ya no se utiliza: el eje estratégico cumple ese rol.
        rep["EJE"] = sum(upsert_cat("EJE", b, c, "", d, i) for i, (b, c, d, e) in enumerate(rows_of("H")))
        rep["OE"] = sum(upsert_cat("OE", b, c, "", d, i) for i, (b, c, d, e) in enumerate(rows_of("E")))
        rep["OBJ"] = sum(upsert_cat("OBJ", b, c, code_display(d), "", i, eje_code(e))
                         for i, (b, c, d, e) in enumerate(rows_of("G")))
        rep["CAT_OO"] = sum(upsert_cat("CAT_OO", b, c, "", d, i) for i, (b, c, d, e) in enumerate(rows_of("K")))
        rep["TIPO_AP"] = sum(upsert_cat("TIPO_AP", b, b, "", c, i) for i, (b, c, d, e) in enumerate(rows_of("L")))
        rep["TIPO_INV"] = sum(upsert_cat("TIPO_INV", b, b, "", c, i) for i, (b, c, d, e) in enumerate(rows_of("M")))
        rep["PRIORIDAD"] = sum(upsert_cat("PRIORIDAD", b, c, "", d, i) for i, (b, c, d, e) in enumerate(rows_of("O")))
        rep["ORIGEN"] = sum(upsert_cat("ORIGEN", b, b, "", "", i) for i, (b, c, d, e) in enumerate(rows_of("P")))
        rep["SUSTENTO_EGR"] = sum(upsert_cat("SUSTENTO_EGR", b, b, "", "", i)
                                  for i, (b, c, d, e) in enumerate(rows_of("Q")))
        rep["SUSTENTO_ING"] = sum(upsert_cat("SUSTENTO_ING", b, b, "", "", i)
                                  for i, (b, c, d, e) in enumerate(rows_of("R")))

        def upsert_niif(code, partida, categoria, tipo):
            if not partida:
                return 0
            old = self.db.query_one("SELECT id FROM catalogo_niif_partidas WHERE partida_rubro=? AND tipo_flujo=? "
                                    "ORDER BY id LIMIT 1", (partida, tipo))
            data = {"codigo_niif": code or "SIN-CODIGO", "partida_rubro": partida, "categoria": categoria,
                    "tipo_flujo": tipo}
            if old:
                self.db.update("catalogo_niif_partidas", old["id"], data)
            else:
                self.db.insert("catalogo_niif_partidas", data)
            return 1

        rep["NIIF_egresos"] = sum(upsert_niif(b, c, d, "Egresos") for b, c, d, e in rows_of("I"))
        rep["NIIF_ingresos"] = sum(upsert_niif(b, c, d, "Ingresos") for b, c, d, e in rows_of("J"))
        self.invalidate_cache()
        n_gg = 0
        for b, c, d, e in rows_of("N"):
            if b:
                tipo = d if d in ("Egresos", "Ingresos") else "Inversion"
                n_gg += self._upsert_gg(b, c, tipo, d if tipo == "Inversion" else "")
        rep["Partidas_GG"] = n_gg
        n_cr = 0
        for b, c, d, e in rows_of("S"):
            if not b:
                continue
            m = re.match(r"^(.*?)\s*[—-]\s*S/\s*([\d.,]*)\s*$", c)
            unidad, costo = (m.group(1), to_float(m.group(2))) if m else (c, None)
            old = self.db.query_one("SELECT id FROM costos_referencia WHERE concepto=?", (b,))
            self.db.upsert("costos_referencia", {"concepto": b, "unidad_medida": unidad, "costo_unitario": costo,
                                                  "fuente": d}, old["id"] if old else None)
            n_cr += 1
        rep["Costos_referencia"] = n_cr
        self.invalidate_cache()
        return rep
