"""
Definición declarativa de formularios y tablas de la versión web.

Cada campo: {key, label, type, ...}
  type: text | textarea | number | money | pct | select | calc | check | months | marks
  src (select): nombre de la lista en /api/listas; dep: campo del que depende la lista (eje, tipo…)
Cada columna: {key, label, fmt, w}; fmt: text | money | pct | int | num | status
Las mismas reglas de negocio se ejecutan en el servidor (controllers.py).
"""
from __future__ import annotations

from typing import Optional

from catalogs_seed import UNIDADES_MEDIDA
from database import MESES

MONEDAS = [{"value": "S/", "label": "S/"}, {"value": "US$", "label": "US$"}]
SINO = [{"value": "Sí", "label": "Sí"}, {"value": "No", "label": "No"}]
MESES_OPT = [{"value": m, "label": m} for m in MESES]


def opts(values):
    return [{"value": v, "label": v} for v in values]


MONEY = [{"key": "_ingresos", "label": "Ingresos (S/)", "fmt": "money"},
         {"key": "_egresos", "label": "Costo / gasto (S/)", "fmt": "money"},
         {"key": "_resultado", "label": "Resultado (S/)", "fmt": "money"},
         {"key": "_margen", "label": "Margen", "fmt": "pct"}]

MASTER = {
    "INI": {
        "titulo": "Iniciativas con impacto estratégico",
        "hoja": "Iniciativas + Detalle_Iniciativas",
        "ayuda": "Seleccione primero el eje estratégico; el objetivo anual del POA se elige entre los objetivos de "
                 "ese eje y el OE del PEI se completa automáticamente.",
        "tipos": ["Egresos", "Ingresos"],
        "techo": "INI",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "eje", "label": "1) Eje estratégico", "type": "select", "src": "ejes", "required": True},
            {"key": "obj_poa", "label": "2) Objetivo anual del POA (del eje)", "type": "select", "src": "objs",
             "dep": "eje", "wide": True},
            {"key": "_oe", "label": "3) OE del PEI (automático)", "type": "calc", "wide": True},
            {"key": "nombre", "label": "Nombre de la iniciativa", "type": "text", "wide": True, "required": True},
            {"key": "descripcion", "label": "Descripción y justificación", "type": "textarea", "wide": True},
            {"key": "origen", "label": "Origen", "type": "select", "src": "origenes"},
            {"key": "directriz", "label": "Directriz que atiende", "type": "text"},
            {"key": "indicador", "label": "Indicador", "type": "text"},
            {"key": "unidad_medida", "label": "Unidad de medida", "type": "select",
             "options": opts(UNIDADES_MEDIDA)},
            {"key": "linea_base", "label": "Línea base", "type": "text"},
            {"key": "meta_anual", "label": "Meta anual", "type": "number"},
            {"key": "t1", "label": "T1", "type": "number"},
            {"key": "t2", "label": "T2", "type": "number"},
            {"key": "t3", "label": "T3", "type": "number"},
            {"key": "_control_meta", "label": "Control de meta", "type": "calc"},
            {"key": "responsable", "label": "Responsable", "type": "text"},
            {"key": "prioridad", "label": "Prioridad (1–4)", "type": "select", "src": "prioridades"},
            {"key": "observaciones", "label": "Observaciones DPDI", "type": "textarea", "wide": True},
        ],
        "columns": [
            {"key": "codigo", "label": "Código"}, {"key": "eje_nombre", "label": "Eje estratégico"},
            {"key": "nombre", "label": "Nombre"}, {"key": "obj_poa", "label": "OBJ POA"},
            {"key": "oe", "label": "OE"}, {"key": "meta_anual", "label": "Meta", "fmt": "num"},
            {"key": "_control_meta", "label": "Control", "fmt": "status"},
            {"key": "prioridad", "label": "Prior.", "fmt": "int"}, *MONEY,
        ],
    },
    "OO": {
        "titulo": "Otras actividades operativas",
        "hoja": "Otras_Operat + Detalle_Otras_Operativas",
        "ayuda": "Licencias, suscripciones, marketing, mantenimientos y similares, sin objetivos asociados.",
        "tipos": ["Egresos", "Ingresos"],
        "techo": "OO",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "categoria", "label": "Categoría", "type": "select", "src": "cat_oo"},
            {"key": "nombre", "label": "Nombre de la actividad", "type": "text", "wide": True, "required": True},
            {"key": "descripcion", "label": "Descripción y justificación", "type": "textarea", "wide": True},
            {"key": "responsable", "label": "Responsable", "type": "text"},
            {"key": "observaciones", "label": "Observaciones", "type": "textarea", "wide": True},
        ],
        "columns": [{"key": "codigo", "label": "Código"}, {"key": "categoria", "label": "Categoría"},
                    {"key": "nombre", "label": "Nombre"}, {"key": "responsable", "label": "Responsable"}, *MONEY],
    },
    "AP": {
        "titulo": "Actividades propias",
        "hoja": "Act_Prop + Detalle_Act_Prop",
        "ayuda": "Congresos, seminarios, talleres y cursos cortos. Deben autofinanciarse.",
        "tipos": ["Ingresos", "Egresos"],
        "techo": "AP",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "tipo", "label": "Tipo", "type": "select", "src": "tipos_ap"},
            {"key": "nombre", "label": "Nombre de la actividad", "type": "text", "wide": True, "required": True},
            {"key": "publico_objetivo", "label": "Público objetivo", "type": "text"},
            {"key": "mes_ejecucion", "label": "Mes de ejecución", "type": "select", "options": MESES_OPT},
            {"key": "participantes", "label": "Participantes proyectados", "type": "number"},
            {"key": "tarifa", "label": "Tarifa por participante (S/)", "type": "money"},
            {"key": "becas_pct", "label": "Becas y descuentos (%)", "type": "pct"},
            {"key": "_tarifa_neta", "label": "Tarifa neta (S/)", "type": "calc"},
            {"key": "responsable", "label": "Responsable", "type": "text"},
            {"key": "sustento", "label": "Sustento", "type": "select", "src": "sustento_ing"},
            {"key": "observaciones", "label": "Observaciones", "type": "textarea", "wide": True},
        ],
        "columns": [{"key": "codigo", "label": "Código"}, {"key": "tipo", "label": "Tipo"},
                    {"key": "nombre", "label": "Nombre"}, {"key": "mes_ejecucion", "label": "Mes"},
                    {"key": "_participantes", "label": "Particip.", "fmt": "num"},
                    {"key": "_tarifa_neta", "label": "Tarifa neta", "fmt": "money"}, *MONEY,
                    {"key": "_punto_equilibrio", "label": "P. equilibrio", "fmt": "int"},
                    {"key": "_estado", "label": "Estado", "fmt": "status"}],
    },
    "INV": {
        "titulo": "Actividades de inversión (CAPEX)",
        "hoja": "Inversion + Detalle_Inversiones",
        "ayuda": "Infraestructura física y tecnológica. Adjunte la ficha de sustento.",
        "tipos": ["Inversion"],
        "techo": "INV",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "tipo_inversion", "label": "Tipo de inversión", "type": "select", "src": "tipos_inv"},
            {"key": "proyecto", "label": "Proyecto de inversión", "type": "text", "wide": True, "required": True},
            {"key": "prioridad", "label": "Prioridad (1–4)", "type": "select", "src": "prioridades"},
            {"key": "ficha_sustento", "label": "Ficha de sustento adjunta", "type": "select", "options": SINO},
            {"key": "observaciones", "label": "Observaciones", "type": "textarea", "wide": True},
        ],
        "columns": [{"key": "codigo", "label": "Código"}, {"key": "tipo_inversion", "label": "Tipo de inversión"},
                    {"key": "proyecto", "label": "Proyecto"},
                    {"key": "_egresos", "label": "Programado (S/)", "fmt": "money"},
                    {"key": "prioridad", "label": "Prior.", "fmt": "int"},
                    {"key": "ficha_sustento", "label": "Ficha"}, {"key": "_n_detalle", "label": "Ítems", "fmt": "int"}],
    },
}

DETALLE_FIELDS = [
    {"key": "tipo", "label": "Tipo", "type": "select", "src": "tipos_detalle"},
    {"key": "partida", "label": "Partida", "type": "select", "src": "gg", "dep": "tipo", "wide": True,
     "search": True, "required": True},
    {"key": "_desc_partida", "label": "Descripción orientativa", "type": "calc", "wide": True},
    {"key": "_categoria", "label": "Categoría", "type": "calc"},
    {"key": "descripcion", "label": "Descripción del bien o servicio", "type": "textarea", "wide": True},
    {"key": "cantidad", "label": "Cantidad", "type": "number"},
    {"key": "moneda", "label": "Moneda", "type": "select", "options": MONEDAS},
    {"key": "costo_unit_moneda", "label": "Costo unitario (moneda)", "type": "money"},
    {"key": "_cu", "label": "Costo unitario (S/)", "type": "calc"},
    {"key": "_total", "label": "Total (S/)", "type": "calc"},
    {"key": "meses", "label": "Distribución mensual del monto (S/)", "type": "months", "wide": True},
    {"key": "_prog", "label": "Programado (S/)", "type": "calc"},
    {"key": "_control", "label": "Control de distribución", "type": "calc"},
    {"key": "sustento", "label": "Sustento", "type": "select", "src": "sustento", "dep": "tipo"},
    {"key": "observaciones", "label": "Observaciones", "type": "text", "wide": True},
]

DETALLE_COLUMNS = [
    {"key": "tipo", "label": "Tipo"}, {"key": "partida", "label": "Partida"}, {"key": "categoria", "label": "Categoría"},
    {"key": "descripcion", "label": "Descripción"}, {"key": "cantidad", "label": "Cant.", "fmt": "num"},
    {"key": "moneda", "label": "Mon."}, {"key": "_cu", "label": "C.U. (S/)", "fmt": "money"},
    {"key": "_total", "label": "Total (S/)", "fmt": "money"},
    *[{"key": f"_m{i}", "label": m, "fmt": "money"} for i, m in enumerate(MESES)],
    {"key": "_prog", "label": "Programado", "fmt": "money"}, {"key": "_control", "label": "Control", "fmt": "status"},
    {"key": "sustento", "label": "Sustento"},
]

FLAT = {
    "FI": {
        "titulo": "Ingresos de funcionamiento (Funcionamiento_Ing)",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "partida", "label": "Partida / rubro (cuenta NIIF)", "type": "select", "src": "cuentas_ing",
             "wide": True, "search": True, "required": True},
            {"key": "_categoria", "label": "Categoría", "type": "calc"},
            {"key": "cantidad", "label": "Cantidad proyectada", "type": "number"},
            {"key": "moneda", "label": "Moneda", "type": "select", "options": MONEDAS},
            {"key": "tarifa_moneda", "label": "Tarifa por cobro (moneda)", "type": "money"},
            {"key": "_soles", "label": "Tarifa por cobro (S/)", "type": "calc"},
            {"key": "becas_pct", "label": "Becas y descuentos (%)", "type": "pct"},
            {"key": "aplica_morosidad", "label": "¿Aplica morosidad?", "type": "select", "options": SINO},
            {"key": "meses", "label": "Meses de cobro: «x» o factor (ej. 0.08333333)", "type": "marks", "wide": True},
            {"key": "_bruto", "label": "Ingreso bruto anual (S/)", "type": "calc"},
            {"key": "_desc", "label": "Becas, descuentos y morosidad (S/)", "type": "calc"},
            {"key": "_neto", "label": "Ingreso neto anual (S/)", "type": "calc"},
            {"key": "sustento", "label": "Sustento", "type": "select", "src": "sustento_ing"},
            {"key": "observaciones", "label": "Observaciones", "type": "text", "wide": True},
        ],
        "columns": [{"key": "codigo", "label": "Código"}, {"key": "partida", "label": "Partida / rubro"},
                    {"key": "categoria", "label": "Categoría"}, {"key": "cantidad", "label": "Cant.", "fmt": "num"},
                    {"key": "_soles", "label": "Tarifa (S/)", "fmt": "money"},
                    {"key": "becas_pct", "label": "Becas", "fmt": "pct"},
                    {"key": "aplica_morosidad", "label": "Moros."},
                    *[{"key": f"_m{i}", "label": m, "fmt": "money"} for i, m in enumerate(MESES)],
                    {"key": "_bruto", "label": "Bruto anual", "fmt": "money"},
                    {"key": "_neto", "label": "Neto anual", "fmt": "money"}, {"key": "sustento", "label": "Sustento"}],
    },
    "FE": {
        "titulo": "Egresos de funcionamiento (Funcionamiento_Egr)",
        "fields": [
            {"key": "codigo", "label": "Código", "type": "text"},
            {"key": "partida", "label": "Partida / rubro (cuenta NIIF)", "type": "select", "src": "cuentas_egr",
             "wide": True, "search": True, "required": True},
            {"key": "_categoria", "label": "Categoría", "type": "calc"},
            {"key": "cantidad", "label": "Cantidad", "type": "number"},
            {"key": "moneda", "label": "Moneda", "type": "select", "options": MONEDAS},
            {"key": "costo_unit_moneda", "label": "Costo unitario por pago (moneda)", "type": "money"},
            {"key": "_soles", "label": "Costo unitario por pago (S/)", "type": "calc"},
            {"key": "_coste", "label": "Coste total (S/)", "type": "calc"},
            {"key": "aplica_reajuste", "label": "¿Aplica reajuste?", "type": "select", "options": SINO},
            {"key": "meses", "label": "Meses de pago: «x» o factor (ej. 0.08333333)", "type": "marks", "wide": True},
            {"key": "_total", "label": "Total anual (S/)", "type": "calc"},
            {"key": "_control", "label": "Control de distribución", "type": "calc"},
            {"key": "sustento", "label": "Sustento", "type": "select", "src": "sustento_egr"},
            {"key": "observaciones", "label": "Observaciones", "type": "text", "wide": True},
        ],
        "columns": [{"key": "codigo", "label": "Código"}, {"key": "partida", "label": "Partida / rubro"},
                    {"key": "categoria", "label": "Categoría"}, {"key": "cantidad", "label": "Cant.", "fmt": "num"},
                    {"key": "_soles", "label": "C.U. por pago", "fmt": "money"},
                    {"key": "_coste", "label": "Coste total", "fmt": "money"},
                    {"key": "aplica_reajuste", "label": "Reaj."},
                    *[{"key": f"_m{i}", "label": m, "fmt": "money"} for i, m in enumerate(MESES)],
                    {"key": "_total", "label": "Total anual", "fmt": "money"},
                    {"key": "_control", "label": "Control", "fmt": "status"}, {"key": "sustento", "label": "Sustento"}],
    },
}

# ------------------------------------------------------------------ administración
TECHOS = ["techo_iniciativas", "techo_funcionamiento", "techo_otras_operativas", "techo_inversion",
          "techo_planilla_docente", "techo_planilla_adm", "techo_marketing", "meta_ingresos_func",
          "meta_resultado_otras", "meta_resultado_propias"]


# Combo especial del Registro Maestro de «Otras operativas»: solo para unidades del macrogrupo Marketing.
CAMPO_ACT_COMERCIAL = {"key": "actividad_comercial", "label": "Tipo de Actividad Comercial", "type": "select",
                       "src": "act_comercial", "required": True}
COLUMNA_ACT_COMERCIAL = {"key": "actividad_comercial", "label": "Tipo de actividad comercial"}


def master_para_unidad(kind: str, unidad: Optional[dict]) -> dict:
    """Configuración del maestro según la unidad: agrega el combo comercial en Otras operativas de Marketing."""
    from controllers import es_marketing
    cfg = MASTER[kind]
    if kind != "OO" or not es_marketing(unidad):
        return cfg
    cfg = dict(cfg)
    campos = list(cfg["fields"])
    campos.insert(next(i for i, f in enumerate(campos) if f["key"] == "categoria") + 1, CAMPO_ACT_COMERCIAL)
    columnas = list(cfg["columns"])
    columnas.insert(next(i for i, c in enumerate(columnas) if c["key"] == "categoria") + 1, COLUMNA_ACT_COMERCIAL)
    cfg.update(fields=campos, columns=columnas)
    return cfg


def admin_entities() -> dict:
    from controllers import CATALOGOS_ESTRATEGICOS, PLANTILLA_PARAM
    from catalogs_seed import BASES_CALCULO
    unidades_fields, unidades_cols = [], []
    for key, header in PLANTILLA_PARAM["Unidades_Techos"]:
        if key == "id":
            continue
        f = {"key": key, "label": header, "type": "text"}
        if key == "macrogrupo":
            f.update(type="select", src="macrogrupos", free=True)
        elif key == "genera_ingresos":
            f.update(type="check", label="¿Genera ingresos?")
        elif key in TECHOS:
            f["type"] = "money"
        unidades_fields.append(f)
        unidades_cols.append({"key": key, "label": header, "fmt": "money" if key in TECHOS else
                              ("sino" if key == "genera_ingresos" else "text")})
    return {
        "unidades": {"titulo": "Unidades orgánicas y techos", "fields": unidades_fields, "columns": unidades_cols,
                     "ayuda": "Misma estructura que la hoja Unidades_Techos de la plantilla de parámetros."},
        "cuentas": {
            "titulo": "Cuentas NIIF + Partidas GG",
            "ayuda": "Base única: cada cuenta NIIF lleva la Partida (GG) que se elige en el detalle y su descripción "
                     "orientativa. Las cuentas sin Partida GG solo se usan en Funcionamiento.",
            "fields": [{"key": "tipo_flujo", "label": "Tipo", "type": "select",
                        "options": opts(["Egresos", "Ingresos", "Inversion"])},
                       {"key": "partida_gg", "label": "Partida (GG)", "type": "text", "wide": True},
                       {"key": "descripcion", "label": "Descripción orientativa", "type": "textarea", "wide": True},
                       {"key": "codigo_niif", "label": "Código NIIF", "type": "text"},
                       {"key": "partida_rubro", "label": "Cuenta contable NIIF", "type": "text", "wide": True},
                       {"key": "categoria", "label": "Categoría", "type": "text"},
                       {"key": "naturaleza", "label": "Naturaleza", "type": "text"},
                       {"key": "resultado_funcion", "label": "Resultado por función", "type": "text"},
                       {"key": "activo", "label": "Activo", "type": "check"}],
            "columns": [{"key": "tipo_flujo", "label": "Tipo"}, {"key": "partida_gg", "label": "Partida (GG)"},
                        {"key": "descripcion", "label": "Descripción"}, {"key": "codigo_niif", "label": "Código NIIF"},
                        {"key": "partida_rubro", "label": "Cuenta NIIF"}, {"key": "categoria", "label": "Categoría"},
                        {"key": "activo", "label": "Activo", "fmt": "sino"}],
        },
        "catalogos": {
            "titulo": "Catálogos estratégicos",
            "ayuda": "Para cada objetivo anual del POA registre su Eje estratégico y su OE del PEI (relación "
                     "Eje → Objetivo anual → OE).",
            "filtro": {"key": "tipo_catalogo", "options": [{"value": k, "label": v}
                                                           for k, v in CATALOGOS_ESTRATEGICOS.items()]},
            "fields": [{"key": "codigo", "label": "Código", "type": "text"},
                       {"key": "descripcion", "label": "Descripción", "type": "textarea", "wide": True},
                       {"key": "eje", "label": "Eje estratégico (solo objetivos POA)", "type": "select",
                        "src": "ejes", "free": True},
                       {"key": "padre", "label": "OE del PEI (solo objetivos POA)", "type": "select", "src": "oes",
                        "free": True},
                       {"key": "extra", "label": "Alcance / indicador / criterio", "type": "textarea", "wide": True},
                       {"key": "orden", "label": "Orden", "type": "number"}],
            "columns": [{"key": "codigo", "label": "Código"}, {"key": "descripcion", "label": "Descripción"},
                        {"key": "eje", "label": "Eje"}, {"key": "padre", "label": "OE"},
                        {"key": "extra", "label": "Extra"}, {"key": "orden", "label": "Orden", "fmt": "int"}],
        },
        "act_comercial": {
            "titulo": "Tipos de actividad comercial (Marketing)",
            "ayuda": "Opciones del combo «Tipo de Actividad Comercial» del Registro Maestro de Otras operativas. Solo lo "
                     "ven y completan las unidades del macrogrupo Marketing. Si renombra una opción, se actualizan las "
                     "actividades que ya la usan; una opción en uso no se puede eliminar.",
            "fields": [{"key": "codigo", "label": "Tipo de actividad comercial", "type": "text", "wide": True},
                       {"key": "orden", "label": "Orden en el combo", "type": "number"}],
            "columns": [{"key": "codigo", "label": "Tipo de actividad comercial"},
                        {"key": "_uso", "label": "Actividades que lo usan", "fmt": "int"},
                        {"key": "orden", "label": "Orden", "fmt": "int"}],
        },
        "costos": {
            "titulo": "Costos unitarios de referencia",
            "fields": [{"key": "concepto", "label": "Concepto", "type": "text", "wide": True},
                       {"key": "unidad_medida", "label": "Unidad de medida", "type": "text"},
                       {"key": "costo_unitario", "label": "Costo unitario (S/)", "type": "money"},
                       {"key": "fuente", "label": "Fuente / vigencia", "type": "text", "wide": True}],
            "columns": [{"key": "concepto", "label": "Concepto"}, {"key": "unidad_medida", "label": "Unidad"},
                        {"key": "costo_unitario", "label": "Costo (S/)", "fmt": "money"},
                        {"key": "fuente", "label": "Fuente"}],
        },
        "plantilla": {
            "titulo": "Plantilla de «Correr Actividades de Funcionamiento»",
            "ayuda": "Monto = Base de cálculo × Porcentaje (o Monto fijo). Se prorratea 1/12 en Ene–Dic.",
            "fields": [{"key": "ambito", "label": "Ámbito (GLOBAL o código de área)", "type": "select",
                        "src": "ambitos"},
                       {"key": "macrogrupo", "label": "Macrogrupo (* = todos)", "type": "select",
                        "src": "macrogrupos_plantilla", "free": True},
                       {"key": "solo_genera_ingresos", "label": "Solo unidades que generan ingresos", "type": "check"},
                       {"key": "tipo_flujo", "label": "Tipo", "type": "select", "options": opts(["Egresos", "Ingresos"])},
                       {"key": "partida", "label": "Cuenta NIIF", "type": "select", "src": "cuentas_todas",
                        "search": True, "wide": True},
                       {"key": "base_calculo", "label": "Base de cálculo", "type": "select",
                        "options": opts(BASES_CALCULO)},
                       {"key": "porcentaje", "label": "Porcentaje de la base (1 = 100 %)", "type": "number"},
                       {"key": "monto_fijo", "label": "Monto fijo anual (S/)", "type": "money"},
                       {"key": "aplica_ajuste", "label": "Aplica reajuste / morosidad", "type": "select",
                        "options": SINO},
                       {"key": "sustento", "label": "Sustento", "type": "select", "src": "sustento_todos"},
                       {"key": "activo", "label": "Activo", "type": "check"},
                       {"key": "orden", "label": "Orden", "type": "number"}],
            "columns": [{"key": "ambito", "label": "Ámbito"}, {"key": "macrogrupo", "label": "Macrogrupo"},
                        {"key": "tipo_flujo", "label": "Tipo"}, {"key": "partida", "label": "Cuenta NIIF"},
                        {"key": "base_calculo", "label": "Base"}, {"key": "porcentaje", "label": "%", "fmt": "pct"},
                        {"key": "monto_fijo", "label": "Monto fijo", "fmt": "money"},
                        {"key": "activo", "label": "Activo", "fmt": "sino"}],
        },
        "usuarios": {
            "titulo": "Usuarios y accesos",
            "ayuda": "Administrador DPDI: acceso total. Consulta DPDI: ve todas las unidades sin modificar. "
                     "Responsable: formula solo las unidades asignadas. La clave inicial debe cambiarse al ingresar.",
            "fields": [{"key": "usuario", "label": "Usuario (correo o código)", "type": "text"},
                       {"key": "nombre", "label": "Nombre completo", "type": "text", "wide": True},
                       {"key": "correo", "label": "Correo electrónico", "type": "text"},
                       {"key": "rol", "label": "Rol", "type": "select",
                        "options": [{"value": "RESPONSABLE", "label": "Responsable de unidad orgánica"},
                                    {"value": "CONSULTA", "label": "Consulta DPDI (solo lectura)"},
                                    {"value": "ADMIN", "label": "Administrador DPDI"}]},
                       {"key": "unidades", "label": "Unidades asignadas (responsables)", "type": "multiselect",
                        "src": "unidades", "wide": True},
                       {"key": "activo", "label": "Activo", "type": "check"},
                       {"key": "nueva_clave", "label": "Nueva clave (dejar vacío para no cambiar)", "type": "password"},
                       {"key": "debe_cambiar_clave", "label": "Debe cambiar la clave al ingresar", "type": "check"}],
            "columns": [{"key": "usuario", "label": "Usuario"}, {"key": "nombre", "label": "Nombre"},
                        {"key": "rol", "label": "Rol"}, {"key": "_unidades", "label": "Unidades"},
                        {"key": "activo", "label": "Activo", "fmt": "sino"},
                        {"key": "ultimo_acceso", "label": "Último acceso"}],
        },
    }


PARAM_FIELDS_EXTRA = [
    {"key": "tema_primario", "label": "Color de marca (logo, menú activo, acentos) — portal: #FF6600", "type": "color"},
    {"key": "tema_secundario", "label": "Color de títulos y encabezados — portal: #333333", "type": "color"},
    {"key": "tema_acento", "label": "Color de botones principales — portal: #FF6600", "type": "color"},
    {"key": "tema_fuente", "label": "Tipografía (nombre de Google Fonts, opcional; requiere internet)", "type": "text"},
]
