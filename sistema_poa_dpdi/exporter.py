"""
exporter.py
===========
Exportaciones del Sistema POA - DPDI:

1. Libro oficial Ficha_Formulacion_POA (.xlsx) con las 16 pestañas del formato oficial,
   estilos (amarillo = dato, gris = fórmula, verde = dato vinculado), listas desplegables,
   nombres definidos (TC, Morosidad, Reajuste, MesReajuste, Techo*, L_*) y fórmulas vivas,
   MÁS la pestaña BD_CONSOLIDADA_NIIF (tabla plana mensualizada para ERP/NIIF y Power BI).
2. Base de datos NIIF consolidada de varias unidades (.xlsx y .csv).
3. Ficha_Resumen firmada en PDF (12 secciones + bloque de conformidad) con ReportLab.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from catalogs_seed import CATEGORIAS_EGR_RESUMEN, CATEGORIAS_ING_RESUMEN
from controllers import POAController, fmt_money, fmt_pct, norm_key, s, to_float
from database import MCOLS, MESES

# =============================================================================
# Estilos del formato oficial
# =============================================================================
NAVY = "1F3864"
F_TITLE = Font(name="Calibri", size=14, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=9, italic=True, color="595959")
F_HDR = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
F_GRP = Font(name="Calibri", size=9, bold=True, color=NAVY)
F_INPUT = Font(name="Calibri", size=9, color="0000FF")
F_CALC = Font(name="Calibri", size=9, color="000000")
F_LINK = Font(name="Calibri", size=9, color="00804A")
F_LABEL = Font(name="Calibri", size=9, bold=True, color="000000")
F_EX = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
F_TOTAL = Font(name="Calibri", size=9, bold=True, color="000000")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_GRP = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_CALC = PatternFill("solid", fgColor="EDEDED")
FILL_EX = PatternFill("solid", fgColor="F2F2F2")
FILL_TOTAL = PatternFill("solid", fgColor="D9D9D9")
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
MONEY = "#,##0.00"
PCT = "0.0%"
FACTOR = "0.00000000"
INT = "0"

OK_WORDS = ['"OK"', '"Cumple"', '"Dentro del techo"', '"Autofinanciada"', '"Superávit"', '"Alcanza la meta"']
BAD_WORDS = ['"Revisar"', '"Excede techo"', '"Déficit"', '"Sin techo"', '"Bajo la meta"', '"Código no existe"',
             '"Sin meta"']

CAPACITY = {"INI": 30, "DET_INI": 150, "FI": 60, "FE": 79, "OO": 30, "DET_OO": 150, "AP": 30, "DET_AP": 150,
            "APL": 25, "INV": 60, "DET_INV": 150, "OOL": 60}


def col(n: int) -> str:
    return get_column_letter(n)


def cidx(letter: str) -> int:
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def shift(letter: str, k: int) -> str:
    return col(cidx(letter) + k)


class XL:
    """Ayudas de escritura con estilo."""

    def __init__(self, ws):
        self.ws = ws

    def put(self, ref: str, value: Any, kind: str = "calc", fmt: Optional[str] = None, bold: bool = False):
        c = self.ws[ref]
        c.value = value
        c.border = BORDER
        if kind == "input":
            c.font, c.fill = F_INPUT, FILL_INPUT
        elif kind == "calc":
            c.font, c.fill = F_CALC, FILL_CALC
        elif kind == "link":
            c.font, c.fill = F_LINK, FILL_CALC
        elif kind == "label":
            c.font = F_LABEL
        elif kind == "example":
            c.font, c.fill = F_EX, FILL_EX
        elif kind == "total":
            c.font, c.fill = F_TOTAL, FILL_TOTAL
        elif kind == "plain":
            c.font = F_CALC
        if bold:
            c.font = Font(name="Calibri", size=9, bold=True, color=c.font.color.rgb if c.font.color else "000000")
        if fmt:
            c.number_format = fmt
        c.alignment = WRAP
        return c

    def title(self, text: str, sub: str = "", unidad: bool = True, col_unidad: str = "B"):
        self.ws["A1"] = text
        self.ws["A1"].font = F_TITLE
        if sub:
            self.ws["A2"] = sub
            self.ws["A2"].font = F_SUB
        if unidad:
            self.ws["A3"] = "Unidad:"
            self.ws["A3"].font = F_LABEL
            c = self.ws[f"{col_unidad}3"]
            c.value = '=IF(Ficha_Resumen!$C$6="","(complete Ficha_Resumen)",Ficha_Resumen!$C$6)'
            c.font = F_LINK

    def group(self, row: int, c1: str, c2: str, text: str):
        self.ws[f"{c1}{row}"] = text
        if c1 != c2:
            self.ws.merge_cells(f"{c1}{row}:{c2}{row}")
        cell = self.ws[f"{c1}{row}"]
        cell.font, cell.fill, cell.alignment = F_GRP, FILL_GRP, CENTER

    def headers(self, row: int, headers: list[str], start: int = 1):
        for i, h in enumerate(headers):
            c = self.ws.cell(row, start + i, h)
            c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER
        self.ws.row_dimensions[row].height = 42

    def widths(self, widths: dict[str, float]):
        for k, v in widths.items():
            self.ws.column_dimensions[k].width = v

    def status_cf(self, rng: str):
        for w in OK_WORDS:
            self.ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[w], fill=FILL_OK))
        for w in BAD_WORDS:
            self.ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[w], fill=FILL_BAD))

    def dv_list(self, rng: str, formula: str):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=False)
        self.ws.add_data_validation(dv)
        dv.add(rng)


def num_or_text(v: Any) -> Any:
    """Marca mensual para Excel: 'x' como texto, factor como número."""
    if v is None or str(v).strip() == "":
        return None
    if str(v).strip().lower() == "x":
        return "x"
    f = to_float(v)
    return f if f is not None else str(v)


# =============================================================================
# 1. LIBRO OFICIAL + BD_CONSOLIDADA_NIIF
# =============================================================================
class OfficialWorkbookExporter:
    def __init__(self, ctrl: POAController, uid: int):
        self.c = ctrl
        self.uid = uid
        self.d = ctrl.load_unit_data(uid)
        self.summary = ctrl.build_summary(uid, self.d)
        self.u = self.d["unidad"]
        self.p = self.d["params"]
        self.prm = self.p.raw or {}
        self.L: dict[str, tuple[int, int, int]] = {}
        self.P: dict[str, tuple[int, int]] = {}

    # ------------------------------------------------------------- layouts
    def _layout(self, key: str, first: int, n: int) -> tuple[int, int, int]:
        rows = max(CAPACITY[key], n)
        lay = (first, first + rows - 1, first + rows)
        self.L[key] = lay
        return lay

    def _rng(self, sheet: str, key: str, letter: str, first: Optional[int] = None) -> str:
        fr, lr, _ = self.L[key]
        return f"{sheet}!${letter}${first or fr}:${letter}${lr}"

    def export(self, path: str | Path, incluir_bd: bool = True) -> Path:
        path = Path(path)
        d = self.d
        n_det = lambda k: sum(len(m["_detalle"]) for m in d[k])  # noqa: E731
        self._layout("INI", 8, len(d["INI"]))
        self._layout("DET_INI", 8, n_det("INI"))
        self._layout("FI", 8, len(d["fi"]))
        self._layout("FE", 8, len(d["fe"]))
        self._layout("OO", 8, len(d["OO"]))
        self._layout("DET_OO", 9, n_det("OO"))
        self._layout("AP", 8, len(d["AP"]))
        self._layout("DET_AP", 9, n_det("AP"))
        self._layout("APL", 8, len(d["AP"]))
        self._layout("INV", 8, len(d["INV"]))
        self._layout("DET_INV", 8, n_det("INV"))
        self._layout("OOL", 8, sum(1 for m in d["OO"] for x in m["_detalle"] if x["tipo"] == "Egresos"))

        wb = Workbook()
        order = ["Instrucciones", "Ficha_Resumen", "Parámetros", "Hoja6", "Funcionamiento_Ing", "Funcionamiento_Egr",
                 "Iniciativas", "Detalle_Iniciativas", "Otras_Operat", "Detalle_Otras_Operativas", "Otras_Operativas",
                 "Act_Prop", "Detalle_Act_Prop", "Act_Propias", "Inversion", "Detalle_Inversiones"]
        ws0 = wb.active
        ws0.title = order[0]
        sheets = {order[0]: ws0}
        for name in order[1:]:
            sheets[name] = wb.create_sheet(name)
        self._parametros(sheets["Parámetros"])
        self._hoja6(sheets["Hoja6"])
        self._names(wb)
        self._iniciativas(sheets["Iniciativas"])
        self._det_iniciativas(sheets["Detalle_Iniciativas"])
        self._fun_ing(sheets["Funcionamiento_Ing"])
        self._fun_egr(sheets["Funcionamiento_Egr"])
        self._otras_operat(sheets["Otras_Operat"])
        self._det_generic(sheets["Detalle_Otras_Operativas"], "OO", "DET_OO", "Otras_Operat")
        self._otras_operativas_legacy(sheets["Otras_Operativas"])
        self._act_prop(sheets["Act_Prop"])
        self._det_generic(sheets["Detalle_Act_Prop"], "AP", "DET_AP", "Act_Prop")
        self._act_propias_legacy(sheets["Act_Propias"])
        self._inversion(sheets["Inversion"])
        self._det_inversiones(sheets["Detalle_Inversiones"])
        self._ficha_resumen(sheets["Ficha_Resumen"])
        self._instrucciones(ws0)
        sheets["Otras_Operativas"].sheet_state = "hidden"
        sheets["Act_Propias"].sheet_state = "hidden"
        if incluir_bd:
            rows = self.c.consolidated_rows([self.uid])
            write_bd_sheet(wb.create_sheet("BD_CONSOLIDADA_NIIF"), rows)
        for ws in wb.worksheets:
            ws.sheet_view.zoomScale = 90
            if ws.title not in ("Instrucciones", "Ficha_Resumen", "Parámetros", "Hoja6", "BD_CONSOLIDADA_NIIF"):
                ws.freeze_panes = "B7"
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
        return path

    # ------------------------------------------------------------ Parámetros
    def _parametros(self, ws):
        x = XL(ws)
        u, prm = self.u, self.prm
        x.title("PARÁMETROS DE PLANIFICACIÓN OPERATIVA Y PRESUPUESTAL",
                "Completa la DPDI según el Instructivo de formulación aprobado. Estos valores alimentan los "
                "cálculos y las listas desplegables.", unidad=False)
        x.widths({"A": 2, "B": 52, "C": 44, "D": 70})
        x.group(4, "B", "D", "A. Datos generales")
        rows = [
            (5, "Año de formulación", s(prm.get("anio_formulacion")), "Ejemplo: 2027 (se registra como texto).", None),
            (6, "Fecha de emisión de los parámetros", s(prm.get("fecha_emision")), "", None),
            (7, "Fecha límite de presentación de la ficha", s(prm.get("fecha_limite")), "", None),
            (8, "Unidad orgánica destinataria", u["nombre_unidad"],
             "Unidad a la que corresponden los techos y metas de la sección B.", None),
        ]
        for r, lab, val, note, fmt in rows:
            x.put(f"B{r}", lab, "label")
            x.put(f"C{r}", val, "input", fmt)
            x.put(f"D{r}", note, "plain")
        x.group(10, "B", "D", "B. Techos de egresos y metas de ingresos asignados a la unidad")
        techos = [
            (11, "Techo — iniciativas con impacto estratégico (S/)", u["techo_iniciativas"], "Gasto corriente de las iniciativas."),
            (12, "Techo — actividades de funcionamiento (S/)", u["techo_funcionamiento"],
             "Remuneraciones, servicios básicos y regulares, depreciaciones y tributos."),
            (13, "Techo — otras actividades operativas (S/)", u["techo_otras_operativas"],
             "Licencias, suscripciones, marketing, mantenimientos y similares."),
            (14, "Techo — actividades de inversión (S/)", u["techo_inversion"], "Infraestructura física y tecnológica."),
        ]
        for r, lab, val, note in techos:
            x.put(f"B{r}", lab, "label")
            x.put(f"C{r}", val or 0, "input", MONEY)
            x.put(f"D{r}", note, "plain")
        x.put("B15", "Techo total de egresos (S/)", "label")
        x.put("C15", "=SUM(C11:C14)", "calc", MONEY)
        x.put("D15", "Las actividades propias no tienen techo: deben autofinanciarse.", "plain")
        for r, lab, val, note in ((16, "Meta de ingresos de funcionamiento (S/)", u["meta_ingresos_func"],
                                   "Solo unidades que generan ingresos regulares."),
                                  (17, "Meta de ingresos de otras actividades operativas (S/)",
                                   u["meta_resultado_otras"], ""),
                                  (18, "Meta de resultado de actividades propias (S/)", u["meta_resultado_propias"],
                                   "Opcional: excedente esperado de congresos, seminarios y talleres.")):
            x.put(f"B{r}", lab, "label")
            x.put(f"C{r}", val or None, "input", MONEY)
            x.put(f"D{r}", note, "plain")
        # Sub-techos (columnas F:G, referencia de la DPDI)
        x.group(10, "F", "G", "Sub-techos de la unidad (referencia)")
        for i, (lab, val) in enumerate((("Techo planilla docente (S/)", u["techo_planilla_docente"]),
                                        ("Techo planilla administrativa (S/)", u["techo_planilla_adm"]),
                                        ("Techo marketing (S/)", u["techo_marketing"]))):
            x.put(f"F{11 + i}", lab, "label")
            x.put(f"G{11 + i}", val or 0, "input", MONEY)
        x.widths({"F": 34, "G": 16})

        x.group(20, "B", "D", "C. Supuestos de formulación")
        sup = [
            (21, "Inflación proyectada (%)", prm.get("inflacion_proyectada"), PCT, "Referencial para estimar costos sin cotización."),
            (22, "Tipo de cambio de referencia (S/ por US$)", prm.get("tipo_cambio"), "0.000",
             "Convierte a soles los montos registrados en US$."),
            (23, "Valor de la UIT (S/)", prm.get("valor_uit"), MONEY, ""),
            (24, "Remuneración mínima vital (S/)", prm.get("rmv"), MONEY, ""),
            (25, "Reajuste remunerativo (%)", prm.get("reajuste_remunerativo_pct"), PCT,
             "Se aplica en Funcionamiento_Egr a las filas con «Aplica reajuste = Sí»."),
            (26, "Mes desde el que se aplica el reajuste (1–12)", prm.get("mes_inicio_reajuste"), INT,
             "1 = enero … 12 = diciembre."),
            (27, "Tasa de morosidad proyectada (%)", prm.get("tasa_morosidad_pct"), PCT,
             "Se descuenta en Funcionamiento_Ing a las filas con «Aplica morosidad = Sí»."),
            (28, "Estudiantes proyectados — pregrado", prm.get("estudiantes_pregrado"), INT,
             "Referencia para las cantidades de Funcionamiento_Ing."),
            (29, "Estudiantes proyectados — posgrado", prm.get("estudiantes_posgrado"), INT, ""),
            (30, "Pensión promedio de referencia — pregrado (S/)", prm.get("pension_ref_pregrado"), MONEY, ""),
            (31, "Pensión promedio de referencia — posgrado (S/)", prm.get("pension_ref_posgrado"), MONEY, ""),
            (32, "N° de cuotas de pensión por periodo académico", prm.get("cuotas_periodo"), INT, ""),
        ]
        for r, lab, val, fmt, note in sup:
            x.put(f"B{r}", lab, "label")
            x.put(f"C{r}", val, "input", fmt)
            x.put(f"D{r}", note, "plain")

        x.group(34, "B", "D", "D. Clasificación de actividades (referencia)")
        x.headers(35, ["Tipo de actividad", "Definición", "Hoja"], 2)
        r = 36
        for it in self.c.db.catalogo("TIPO_ACTIVIDAD"):
            x.put(f"B{r}", it["codigo"], "label")
            x.put(f"C{r}", it["descripcion"], "plain")
            x.put(f"D{r}", it["extra"], "plain")
            r += 1

        def block(letter_title: str, headers: list[str], rows: list[tuple], key: str, start: int) -> int:
            x.group(start, "B", "D", letter_title)
            x.headers(start + 1, headers, 2)
            rr = start + 2
            for t in rows:
                for j, v in enumerate(t):
                    x.put(f"{'BCD'[j]}{rr}", v, "input")
                rr += 1
            self.P[key] = (start + 2, max(start + 2, rr - 1))
            return rr + 1

        db = self.c.db
        r = block("E. Objetivos estratégicos del PEI", ["Código", "Descripción del objetivo estratégico",
                                                        "Indicador del PEI asociado"],
                  [(o["codigo"], o["descripcion"], o["extra"]) for o in db.catalogo("OE")], "oe", r + 1)
        r = block("F. Acciones estratégicas del PEI", ["Código de AE", "OE al que pertenece",
                                                       "Descripción de la acción estratégica"],
                  [(o["codigo"], o["padre"], o["descripcion"]) for o in db.catalogo("AE")], "ae", r)
        r = block("G. Objetivos anuales del POA", ["Código", "Objetivo anual del POA", "OE del PEI al que contribuye"],
                  [(o["codigo"], o["descripcion"], o["padre"]) for o in db.catalogo("OBJ")], "obj", r)
        r = block("H. Ejes de las iniciativas con impacto estratégico", ["Código", "Eje estratégico",
                                                                          "Alcance referencial"],
                  [(o["codigo"], o["descripcion"], o["extra"]) for o in db.catalogo("EJE")], "eje", r)
        egr = db.query("SELECT codigo_niif, partida_rubro, categoria FROM catalogo_niif_partidas "
                       "WHERE tipo_flujo='Egresos' AND activo=1 ORDER BY partida_rubro, id")
        r = block("I. Categorías de egresos de funcionamiento", ["Código", "Partida / rubro", "Categoría"],
                  [(e["codigo_niif"], e["partida_rubro"], e["categoria"]) for e in egr], "cat_fe", r)
        ing = db.query("SELECT codigo_niif, partida_rubro, categoria FROM catalogo_niif_partidas "
                       "WHERE tipo_flujo='Ingresos' AND activo=1 ORDER BY id")
        r = block("J. Conceptos de ingresos de funcionamiento", ["Código", "Partida / rubro", "Categoría"],
                  [(e["codigo_niif"], e["partida_rubro"], e["categoria"]) for e in ing], "conc_fi", r)
        r = block("K. Categorías de otras actividades operativas", ["Código", "Categoría", "Incluye"],
                  [(o["codigo"], o["descripcion"], o["extra"]) for o in db.catalogo("CAT_OO")], "cat_oo", r)
        r = block("L. Tipos de actividades propias", ["Tipo", "Descripción", ""],
                  [(o["codigo"], o["extra"], "") for o in db.catalogo("TIPO_AP")], "tipo_ap", r)
        r = block("M. Tipos de inversión", ["Tipo", "Incluye", ""],
                  [(o["codigo"], o["extra"], "") for o in db.catalogo("TIPO_INV")], "tipo_inv", r)
        gg = []
        for tipo in ("Egresos", "Inversion", "Ingresos"):
            for g in db.query("SELECT * FROM catalogo_partidas_gg WHERE tipo=? ORDER BY id", (tipo,)):
                gg.append((g["partida_gg"], g["descripcion"], g["codigo_niif"] if tipo == "Inversion" else tipo))
        r = block("N. Partidas / rubros presupuestales", ["Partida", "Descripción", "Observación"], gg, "gg", r)
        g0 = self.P["gg"][0]
        n_e = sum(1 for t in gg if t[2] == "Egresos")
        n_v = sum(1 for t in gg if t[2] not in ("Egresos", "Ingresos"))
        self.P["gg_egr"] = (g0, g0 + n_e - 1)
        self.P["gg_inv"] = (g0 + n_e, g0 + n_e + n_v - 1)
        self.P["gg_ing"] = (g0 + n_e + n_v, self.P["gg"][1])
        r = block("O. Niveles de prioridad", ["Nivel", "Categoría", "Criterio"],
                  [(int(o["codigo"]), o["descripcion"], o["extra"]) for o in db.catalogo("PRIORIDAD")], "prio", r)
        r = block("P. Origen de las iniciativas", ["Origen", "", ""],
                  [(o["codigo"], "", "") for o in db.catalogo("ORIGEN")], "origen", r)
        r = block("Q. Tipos de sustento de egresos", ["Sustento", "", ""],
                  [(o["codigo"], "", "") for o in db.catalogo("SUSTENTO_EGR")], "sust", r)
        r = block("R. Tipos de sustento de ingresos", ["Sustento", "", ""],
                  [(o["codigo"], "", "") for o in db.catalogo("SUSTENTO_ING")], "sust_ing", r)
        costos = db.all("costos_referencia", order="id")
        block("S. Costos unitarios de referencia", ["Concepto", "Unidad de medida / Costo unitario (S/)",
                                                    "Fuente / vigencia"],
              [(c["concepto"], f"{c['unidad_medida']} — S/ {fmt_money(c['costo_unitario'])}".strip(), c["fuente"])
               for c in costos], "costos", r)

    # ---------------------------------------------------------------- Hoja6
    def _hoja6(self, ws):
        x = XL(ws)
        x.widths({"C": 48, "D": 48, "E": 44, "F": 52, "G": 36, "H": 34})
        x.headers(5, ["ingresos", "Egresos", "Inversion", "Partida NIIF (todas)", "Categoría", "Sustentos"], 3)
        gi, ge, gv = self.P["gg_ing"], self.P["gg_egr"], self.P["gg_inv"]
        for i, rr in enumerate(range(gi[0], gi[1] + 1)):
            x.put(f"C{6 + i}", f"=+Parámetros!B{rr}", "link")
        for i, rr in enumerate(range(ge[0], ge[1] + 1)):
            x.put(f"D{6 + i}", f"=+Parámetros!B{rr}", "link")
        for i, rr in enumerate(range(gv[0], gv[1] + 1)):
            x.put(f"E{6 + i}", f"=+Parámetros!B{rr}", "link")
        self.P["h6_ing"] = (6, 6 + gi[1] - gi[0])
        self.P["h6_egr"] = (6, 6 + ge[1] - ge[0])
        self.P["h6_inv"] = (6, 6 + gv[1] - gv[0])
        allp = self.c.db.query("SELECT partida_rubro, categoria FROM catalogo_niif_partidas WHERE activo=1 "
                               "ORDER BY CASE tipo_flujo WHEN 'Egresos' THEN 0 WHEN 'Ingresos' THEN 1 ELSE 2 END, "
                               "partida_rubro, id")
        seen, r = set(), 6
        for p in allp:
            k = norm_key(p["partida_rubro"])
            if k in seen:
                continue
            seen.add(k)
            x.put(f"F{r}", p["partida_rubro"], "plain")
            x.put(f"G{r}", p["categoria"], "plain")
            r += 1
        self.P["h6_all"] = (6, r - 1)
        sust = [o["codigo"] for o in self.c.db.catalogo("SUSTENTO_EGR")]
        sust += [o["codigo"] for o in self.c.db.catalogo("SUSTENTO_ING") if o["codigo"] not in sust]
        for i, v in enumerate(sust):
            x.put(f"H{6 + i}", v, "plain")
        self.P["h6_sust"] = (6, 5 + len(sust))

    def _names(self, wb):
        P = self.P

        def add(name, ref):
            wb.defined_names[name] = DefinedName(name, attr_text=ref)

        add("AnioPOA", "Parámetros!$C$5")
        add("TechoIniciativas", "Parámetros!$C$11")
        add("TechoFuncionamiento", "Parámetros!$C$12")
        add("TechoOtras", "Parámetros!$C$13")
        add("TechoInversion", "Parámetros!$C$14")
        add("MetaIngresos", "Parámetros!$C$16")
        add("MetaOtras", "Parámetros!$C$17")
        add("MetaPropias", "Parámetros!$C$18")
        add("TC", "Parámetros!$C$22")
        add("Reajuste", "Parámetros!$C$25")
        add("MesReajuste", "Parámetros!$C$26")
        add("Morosidad", "Parámetros!$C$27")
        for name, key, c in (("L_OE", "oe", "B"), ("L_AE", "ae", "B"), ("L_ObjPOA", "obj", "B"),
                             ("L_Ejes", "eje", "C"), ("L_CatFE", "cat_fe", "C"), ("L_ConcFI", "conc_fi", "C"),
                             ("L_CatOO", "cat_oo", "C"), ("L_TipoAP", "tipo_ap", "B"), ("L_TipoINV", "tipo_inv", "B"),
                             ("L_Partidas", "gg", "B"), ("L_Origen", "origen", "B"), ("L_Sustento", "sust", "B"),
                             ("L_SustIng", "sust_ing", "B")):
            a, b = P[key]
            add(name, f"Parámetros!${c}${a}:${c}${b}")
        add("Ingresos", f"Hoja6!$C${P['h6_ing'][0]}:$C${P['h6_ing'][1]}")
        add("Egresos", f"Hoja6!$D${P['h6_egr'][0]}:$D${P['h6_egr'][1]}")
        add("Inversion", f"Hoja6!$E${P['h6_inv'][0]}:$E${P['h6_inv'][1]}")
        add("L_PartidasNIIF", f"Hoja6!$F${P['h6_all'][0]}:$F${P['h6_all'][1]}")
        add("L_SustAll", f"Hoja6!$H${P['h6_sust'][0]}:$H${P['h6_sust'][1]}")
        fr, lr, _ = self.L["INI"]
        add("L_INI", f"Iniciativas!$A${fr}:$A${lr}")

    def _cat_formula(self, ref: str) -> str:
        a, b = self.P["h6_all"]
        return f'=IFERROR(VLOOKUP({ref},Hoja6!$F${a}:$G${b},2,FALSE),"")'

    # --------------------------------------------------------- Iniciativas
    def _iniciativas(self, ws):
        x = XL(ws)
        x.title("1. INICIATIVAS CON IMPACTO ESTRATÉGICO",
                "Iniciativas cuya realización impacta en los objetivos anuales del POA. Una fila por iniciativa; el "
                "gasto corriente se calcula desde Detalle_Iniciativas y la inversión vinculada desde Inversion.")
        x.group(5, "A", "I", "Identificación y alineamiento (POA y PEI)")
        x.group(5, "J", "Q", "Indicador y meta física")
        x.group(5, "R", "S", "Gestión")
        x.group(5, "T", "W", "Presupuesto (calculado)")
        x.group(5, "X", "X", "Revisión")
        x.headers(6, ["Código", "Eje estratégico", "Nombre de la iniciativa", "Descripción y justificación", "Origen",
                      "Directriz que atiende", "Objetivo anual del POA", "OE del PEI", "AE del PEI", "Indicador",
                      "Unidad de medida", "Línea base", "Meta anual", "T1", "T2", "T3", "Control de meta",
                      "Responsable", "Prioridad (1–4)", "Ingresos Total (S/)", "Gasto corriente (S/)",
                      "Resultado (S/)", "Margen (%)", "Observaciones de la DPDI"])
        x.widths({"A": 10, "B": 22, "C": 34, "D": 44, "E": 24, "F": 14, "G": 12, "H": 10, "I": 10, "J": 30, "K": 14,
                  "L": 10, "M": 10, "N": 8, "O": 8, "P": 8, "Q": 12, "R": 26, "S": 10, "T": 14, "U": 14, "V": 14,
                  "W": 10, "X": 30})
        dfr, dlr, _ = self.L["DET_INI"]
        fr, lr, tr = self.L["INI"]

        def formulas(r, det_a, det_b, kind):
            x.put(f"Q{r}", f'=IF(C{r}="","",IF(M{r}="","Sin meta",IF(SUM(N{r}:P{r})=M{r},"OK","Revisar")))', kind)
            for colx, tipo in (("T", "Ingresos"), ("U", "Egresos")):
                x.put(f"{colx}{r}", f"=SUMIFS(Detalle_Iniciativas!$L${det_a}:$L${det_b},Detalle_Iniciativas!$A${det_a}"
                                    f":$A${det_b},$A{r},Detalle_Iniciativas!$C${det_a}:$C${det_b},\"{tipo}\")",
                      kind, MONEY)
            x.put(f"V{r}", f"=T{r}-U{r}", kind, MONEY)
            x.put(f"W{r}", f'=IF(T{r}=0,"",V{r}/T{r})', kind, PCT)

        ex = ["EJEMPLO", "Transformación Digital", "Implementación de aulas virtuales con analítica de aprendizaje",
              "Ampliar el uso de aulas virtuales para mejorar el seguimiento del aprendizaje; responde a la debilidad "
              "D3 del FODA.", "Diagnóstico estratégico (FODA)", "DE-02", "OBJ-03", "OE3", "AE3.1",
              "Cursos con aula virtual implementada", "Curso", 20, 60, 10, 20, 20, None,
              "Jefe de Tecnologías Educativas", 2]
        for i, v in enumerate(ex):
            if v is not None:
                x.put(f"{col(i + 1)}7", v, "example")
        formulas(7, 7, 7, "example")
        recs = self.d["INI"]
        nxt = len(recs)
        for i, r in enumerate(range(fr, lr + 1)):
            m = recs[i] if i < len(recs) else None
            if m:
                vals = [m["codigo"], m["eje"], m["nombre"], m["descripcion"], m["origen"], m["directriz"],
                        m["obj_poa"], m["oe"], m["ae"], m["indicador"], m["unidad_medida"],
                        to_float(m["linea_base"]) if to_float(m["linea_base"]) is not None else m["linea_base"],
                        m["meta_anual"], m["t1"], m["t2"], m["t3"]]
                for j, v in enumerate(vals):
                    x.put(f"{col(j + 1)}{r}", v if v != "" else None, "input")
                x.put(f"R{r}", m["responsable"] or None, "input")
                x.put(f"S{r}", m["prioridad"], "input")
                x.put(f"X{r}", m["observaciones"] or None, "input")
            else:
                nxt += 1
                x.put(f"A{r}", f"INI-{nxt:02d}", "input")
                for j in list(range(2, 17)) + [18, 19, 24]:
                    x.put(f"{col(j)}{r}", None, "input")
            formulas(r, dfr, dlr, "calc")
        x.put(f"A{tr}", "TOTAL INICIATIVAS CON IMPACTO ESTRATÉGICO (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:S{tr}")
        for c_ in "TUV":
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        x.dv_list(f"B{fr}:B{lr}", "=L_Ejes")
        x.dv_list(f"E{fr}:E{lr}", "=L_Origen")
        x.dv_list(f"G{fr}:G{lr}", "=L_ObjPOA")
        x.dv_list(f"H{fr}:H{lr}", "=L_OE")
        x.dv_list(f"I{fr}:I{lr}", "=L_AE")
        x.dv_list(f"S{fr}:S{lr}", '"1,2,3,4"')
        x.status_cf(f"Q7:Q{lr}")

    # ------------------------------------------------- Detalle_Iniciativas
    def _det_iniciativas(self, ws):
        x = XL(ws)
        x.title("1. DETALLE DEL GASTO CORRIENTE DE LAS INICIATIVAS",
                "Varias filas por iniciativa: un bien o servicio por fila. Distribuya el total en los meses de "
                "ejecución. Los bienes de inversión se registran en la hoja Inversion.")
        x.group(5, "A", "D", "Iniciativa")
        x.group(5, "E", "L", "Bien o servicio y costo")
        x.group(5, "M", "X", "Distribución mensual del monto (S/)")
        x.group(5, "Y", "AB", "Control y sustento")
        x.headers(6, ["Código de iniciativa", "Nombre de la iniciativa", "Tipo", "Partida / rubro",
                      "Partida / rubro GG", "Categoría", "Descripción del bien o servicio", "Cantidad", "Moneda",
                      "Costo unitario (moneda)", "Costo unitario (S/)", "Total (S/)", *MESES, "Programado (S/)",
                      "Control de distribución", "Sustento", "Observaciones"])
        x.widths({"A": 11, "B": 30, "C": 10, "D": 34, "E": 30, "F": 22, "G": 40, "H": 9, "I": 7, "J": 13, "K": 13,
                  "L": 14, **{col(13 + i): 11 for i in range(12)}, "Y": 14, "Z": 12, "AA": 20, "AB": 30})
        fr, lr, tr = self.L["DET_INI"]
        ifr, ilr, _ = self.L["INI"]

        def formulas(r, kind, ex=False):
            if not ex:
                x.put(f"B{r}", f'=IF(A{r}="","",IFERROR(INDEX(Iniciativas!$C$7:$C${ilr},MATCH(A{r},'
                               f'Iniciativas!$A$7:$A${ilr},0))&"","Código no existe"))', "link")
            x.put(f"F{r}", self._cat_formula(f"D{r}"), "link" if kind == "calc" else kind)
            x.put(f"K{r}", f'=IF(J{r}="","",IF(I{r}="US$",J{r}*TC,J{r}))', kind, MONEY)
            x.put(f"L{r}", f'=IF(OR(H{r}="",K{r}=""),0,H{r}*K{r})', kind, MONEY)
            x.put(f"Y{r}", f"=SUM(M{r}:X{r})", kind, MONEY)
            x.put(f"Z{r}", f'=IF(AND(L{r}=0,Y{r}=0),"",IF(ABS(Y{r}-L{r})<0.01,"OK","Revisar"))', kind)

        for c_, v in (("A", "EJEMPLO"), ("B", "Renovación Licencia Educativa"), ("C", "Egresos"),
                      ("D", "LICENCIAS Y DERECHOS DE VIGENCIA"), ("E", "Licencias y derechos de vigencia"),
                      ("G", "Licencia anual de plataforma de analítica de aprendizaje"), ("H", 1), ("I", "S/"),
                      ("J", 12000), ("O", 6000), ("T", 6000), ("AA", "Cotización"),
                      ("AB", "Pago en dos armadas (marzo y agosto).")):
            x.put(f"{c_}7", v, "example")
        formulas(7, "example", ex=True)
        rows = [(m, dd) for m in self.d["INI"] for dd in m["_detalle"]]
        self._detail_rows(x, rows, fr, lr, formulas, letters=dict(code="A", tipo="C", partida="D", gg="E",
                                                                     desc="G", cant="H", mon="I", cu="J", m0="M",
                                                                     sust="AA", obs="AB"))
        x.put(f"A{tr}", "TOTAL (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:K{tr}")
        for i in range(14):
            c_ = col(12 + i)
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        x.dv_list(f"A{fr}:A{lr}", "=L_INI")
        x.dv_list(f"C{fr}:C{lr}", '"Ingresos,Egresos"')
        x.dv_list(f"D{fr}:D{lr}", "=L_PartidasNIIF")
        x.dv_list(f"E{fr}:E{lr}", "=L_Partidas")
        x.dv_list(f"I{fr}:I{lr}", '"S/,US$"')
        x.dv_list(f"AA{fr}:AA{lr}", "=L_SustAll")
        x.status_cf(f"Z7:Z{lr}")

    def _detail_rows(self, x: XL, rows: list, fr: int, lr: int, formulas, letters: dict):
        for i, r in enumerate(range(fr, lr + 1)):
            item = rows[i] if i < len(rows) else None
            if item:
                m, dd = item
                vals = {letters["code"]: m["codigo"], letters["tipo"]: dd["tipo"], letters["partida"]: dd["partida"],
                        letters["desc"]: dd["descripcion"], letters["cant"]: to_float(dd["cantidad"]),
                        letters["mon"]: dd["moneda"] or "S/", letters["cu"]: to_float(dd["costo_unit_moneda"]),
                        letters["sust"]: dd["sustento"], letters["obs"]: dd["observaciones"]}
                if "gg" in letters:
                    vals[letters["gg"]] = dd["partida_gg"]
                for c_, v in vals.items():
                    x.put(f"{c_}{r}", v if v != "" else None, "input",
                          MONEY if c_ == letters["cu"] else None)
                for k in range(12):
                    v = to_float(dd[MCOLS[k]])
                    x.put(f"{shift(letters['m0'], k)}{r}", v, "input", MONEY)
            else:
                for key in ("code", "tipo", "partida", "desc", "cant", "mon", "cu", "sust", "obs") + \
                        (("gg",) if "gg" in letters else ()):
                    x.put(f"{letters[key]}{r}", None, "input")
                for k in range(12):
                    x.put(f"{shift(letters['m0'], k)}{r}", None, "input", MONEY)
            formulas(r, "calc")

    # ---------------------------------------------------- Funcionamiento_Ing
    def _fun_ing(self, ws):
        x = XL(ws)
        x.title("2. ACTIVIDADES DE FUNCIONAMIENTO — INGRESOS",
                "Ingresos de naturaleza regular: matrículas, pensiones, tasas educativas y otros. Una fila por "
                "concepto y programa o servicio; marque con «x» los meses de cobro.")
        x.group(5, "A", "I", "Concepto y tarifa")
        x.group(5, "J", "V", "Meses de cobro (marque x)")
        x.group(5, "W", "AH", "Ingreso mensual neto calculado (S/)")
        x.group(5, "AI", "AK", "Ingreso anual (S/)")
        x.group(5, "AL", "AM", "Sustento")
        x.headers(6, ["Código", "Partida / rubro", "Categoría", "Cantidad proyectada", "Moneda",
                      "Tarifa por cobro (moneda)", "Tarifa por cobro (S/)", "Becas y descuentos (%)",
                      "¿Aplica morosidad?", *MESES, "Ingreso bruto anual (S/)", *[f"{m} (S/)" for m in MESES],
                      "Ingreso neto anual (S/)", "Becas, descuentos y morosidad (S/)", "Ingreso neto anual (S/)",
                      "Sustento", "Observaciones"])
        x.widths({"A": 9, "B": 34, "C": 22, "D": 11, "E": 7, "F": 14, "G": 14, "H": 10, "I": 10,
                  **{col(10 + i): 7 for i in range(12)}, "V": 14, **{col(23 + i): 12 for i in range(12)},
                  "AI": 14, "AJ": 14, "AK": 14, "AL": 26, "AM": 30})
        fr, lr, tr = self.L["FI"]

        def formulas(r, kind):
            x.put(f"C{r}", self._cat_formula(f"B{r}"), "link" if kind == "calc" else kind)
            x.put(f"G{r}", f'=IF(F{r}="","",IF(E{r}="US$",F{r}*TC,F{r}))', kind, MONEY)
            x.put(f"V{r}", f'=IF(OR($D{r}="",$G{r}=""),0,$D{r}*$G{r}*(COUNTIF(J{r}:U{r},"x")+SUMIF(J{r}:U{r},">0")))',
                  kind, MONEY)
            for k in range(12):
                mc = col(10 + k)
                x.put(f"{col(23 + k)}{r}",
                      f'=IF(OR($D{r}="",$G{r}="",{mc}{r}=""),0,$D{r}*$G{r}*IF({mc}{r}="x",1,N({mc}{r}))'
                      f'*(1-IF($H{r}="",0,$H{r}))*(1-IF($I{r}="Sí",Morosidad,0)))', kind, MONEY)
            x.put(f"AI{r}", f"=SUM(W{r}:AH{r})", kind, MONEY)
            x.put(f"AJ{r}", f"=V{r}-AI{r}", kind, MONEY)
            x.put(f"AK{r}", f"=SUM(W{r}:AH{r})", kind, MONEY)

        ex = {"A": "EJEMPLO", "B": "PENSION PREGRADO", "D": 350, "E": "S/", "F": 650, "H": 0.1, "I": "Sí",
              "AL": "Proyección de matrícula aprobada"}
        for c_, v in ex.items():
            x.put(f"{c_}7", v, "example", PCT if c_ == "H" else None)
        for k in range(2, 12):
            x.put(f"{col(10 + k)}7", 0.1, "example", FACTOR)
        formulas(7, "example")
        recs = self.d["fi"]
        for i, r in enumerate(range(fr, lr + 1)):
            rec = recs[i] if i < len(recs) else None
            vals = {}
            if rec:
                vals = {"A": rec["codigo"], "B": rec["partida"], "D": to_float(rec["cantidad"]),
                        "E": rec["moneda"] or "S/", "F": to_float(rec["tarifa_moneda"]), "H": rec["becas_pct"],
                        "I": rec["aplica_morosidad"], "AL": rec["sustento"], "AM": rec["observaciones"]}
                for k in range(12):
                    vals[col(10 + k)] = num_or_text(rec[MCOLS[k]])
            for c_ in ["A", "B", "D", "E", "F", "H", "I", "AL", "AM"] + [col(10 + k) for k in range(12)]:
                v = vals.get(c_)
                fmt = MONEY if c_ == "F" else PCT if c_ == "H" else FACTOR if cidx(c_) in range(10, 22) else None
                x.put(f"{c_}{r}", v if v != "" else None, "input", fmt)
            formulas(r, "calc")
        x.put(f"A{tr}", "TOTAL ACTIVIDADES DE FUNCIONAMIENTO — INGRESOS (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:U{tr}")
        for c_ in ["V"] + [col(23 + k) for k in range(12)] + ["AI", "AJ", "AK"]:
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        x.dv_list(f"B{fr}:B{lr}", "=L_ConcFI")
        x.dv_list(f"E{fr}:E{lr}", '"S/,US$"')
        x.dv_list(f"I{fr}:I{lr}", '"Sí,No"')
        x.dv_list(f"AL{fr}:AL{lr}", "=L_SustIng")

    # ---------------------------------------------------- Funcionamiento_Egr
    def _fun_egr(self, ws):
        x = XL(ws)
        x.title("2. ACTIVIDADES DE FUNCIONAMIENTO — EGRESOS",
                "Remuneraciones, servicios básicos y regulares, depreciaciones y tributos. Marque con «x» los meses de "
                "pago (o el factor de prorrateo, p. ej. 0.08333333); el reajuste remunerativo se aplica a las filas "
                "marcadas «Sí».")
        x.put("V4", "N° de mes →", "label")
        for k in range(12):
            x.put(f"{col(23 + k)}4", k + 1, "calc", INT)
        x.group(5, "A", "I", "Concepto y costo")
        x.group(5, "J", "V", "Meses de pago (marque x)")
        x.group(5, "W", "AH", "Monto mensual calculado (S/)")
        x.group(5, "AI", "AJ", "Total y control")
        x.group(5, "AK", "AL", "Sustento")
        x.headers(6, ["Código", "Partida / rubro", "Categoría", "Cantidad", "Moneda", "Costo unitario por pago (moneda)",
                      "Costo unitario por pago (S/)", "Coste total", "¿Aplica reajuste?", *MESES, "N° de pagos",
                      *[f"{m} (S/)" for m in MESES], "Total anual (S/)", "Control de distribución", "Sustento",
                      "Observaciones"])
        x.widths({"A": 9, "B": 36, "C": 24, "D": 9, "E": 7, "F": 14, "G": 14, "H": 14, "I": 10,
                  **{col(10 + i): 10 for i in range(12)}, "V": 9, **{col(23 + i): 12 for i in range(12)},
                  "AI": 14, "AJ": 12, "AK": 24, "AL": 30})
        fr, lr, tr = self.L["FE"]

        def formulas(r, kind):
            x.put(f"C{r}", self._cat_formula(f"B{r}"), "link" if kind == "calc" else kind)
            x.put(f"G{r}", f'=IF(F{r}="","",IF(E{r}="US$",F{r}*TC,F{r}))', kind, MONEY)
            x.put(f"H{r}", f'=IF(OR(D{r}="",G{r}=""),0,G{r}*D{r})', kind, MONEY)
            x.put(f"V{r}", f'=COUNTIF(J{r}:U{r},"x")+COUNTIF(J{r}:U{r},">0")', kind, INT)
            for k in range(12):
                mc, oc = col(10 + k), col(23 + k)
                x.put(f"{oc}{r}", f'=IF(OR($D{r}="",$H{r}="",{mc}{r}=""),0,IF({mc}{r}="x",1,N({mc}{r}))*$H{r}'
                                  f'*IF(AND($I{r}="Sí",{oc}$4>=MesReajuste),1+Reajuste,1))', kind, MONEY)
            x.put(f"AI{r}", f"=SUM(W{r}:AH{r})", kind, MONEY)
            x.put(f"AJ{r}", f'=IF(AND(H{r}=0,AI{r}=0),"",IF(COUNTIF(J{r}:U{r},"x")>0,"OK",'
                            f'IF(ABS(SUM(J{r}:U{r})-1)<0.0001,"OK","Revisar")))', kind)

        for c_, v in {"A": "EJEMPLO", "B": "CTS ADMINISTRATIVOS", "D": 1, "E": "S/", "F": 2500, "I": "No",
                      "AK": "Contrato vigente"}.items():
            x.put(f"{c_}7", v, "example")
        for k in range(12):
            x.put(f"{col(10 + k)}7", 1 / 12, "example", FACTOR)
        formulas(7, "example")
        recs = self.d["fe"]
        for i, r in enumerate(range(fr, lr + 1)):
            rec = recs[i] if i < len(recs) else None
            vals = {}
            if rec:
                vals = {"A": rec["codigo"], "B": rec["partida"], "D": to_float(rec["cantidad"]),
                        "E": rec["moneda"] or "S/", "F": to_float(rec["costo_unit_moneda"]),
                        "I": rec["aplica_reajuste"], "AK": rec["sustento"], "AL": rec["observaciones"]}
                for k in range(12):
                    vals[col(10 + k)] = num_or_text(rec[MCOLS[k]])
            for c_ in ["A", "B", "D", "E", "F", "I", "AK", "AL"] + [col(10 + k) for k in range(12)]:
                v = vals.get(c_)
                fmt = MONEY if c_ == "F" else FACTOR if cidx(c_) in range(10, 22) else None
                x.put(f"{c_}{r}", v if v != "" else None, "input", fmt)
            formulas(r, "calc")
        x.put(f"A{tr}", "TOTAL ACTIVIDADES DE FUNCIONAMIENTO — EGRESOS (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:V{tr}")
        for c_ in [col(23 + k) for k in range(12)] + ["AI"]:
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        x.dv_list(f"B{fr}:B{lr}", "=L_CatFE")
        x.dv_list(f"E{fr}:E{lr}", '"S/,US$"')
        x.dv_list(f"I{fr}:I{lr}", '"Sí,No"')
        x.dv_list(f"AK{fr}:AK{lr}", "=L_Sustento")
        x.status_cf(f"AJ7:AJ{lr}")

    # -------------------------------------------------------- Otras_Operat
    def _otras_operat(self, ws):
        x = XL(ws)
        x.title("3. OTRAS ACTIVIDADES OPERATIVAS",
                "Licencias, suscripciones, marketing, mantenimientos y otras actividades regulares sin objetivos "
                "asociados. El ingreso y el costo se calculan desde Detalle_Otras_Operativas.", col_unidad="C")
        x.group(5, "A", "D", "Identificación de la actividad")
        x.group(5, "E", "I", "Gestión y resultado (calculado)")
        x.group(5, "J", "J", "Revisión")
        x.headers(6, ["Código", "Categoría", "Nombre de la actividad", "Descripción y justificación", "Responsable",
                      "Ingresos Total (S/)", "Costo total (S/)", "Resultado (S/)", "Margen (%)", "Observaciones"])
        x.widths({"A": 10, "B": 30, "C": 36, "D": 48, "E": 26, "F": 14, "G": 14, "H": 14, "I": 10, "J": 34})
        fr, lr, tr = self.L["OO"]
        dfr, dlr, _ = self.L["DET_OO"]

        def formulas(r, a, b, kind):
            for c_, tipo in (("F", "Ingresos"), ("G", "Egresos")):
                x.put(f"{c_}{r}", f"=SUMIFS(Detalle_Otras_Operativas!$X${a}:$X${b},Detalle_Otras_Operativas!$A${a}:"
                                  f"$A${b},$A{r},Detalle_Otras_Operativas!$C${a}:$C${b},\"{tipo}\")", kind, MONEY)
            x.put(f"H{r}", f"=F{r}-G{r}", kind, MONEY)
            x.put(f"I{r}", f'=IF(F{r}=0,"",H{r}/F{r})', kind, PCT)

        for c_, v in (("A", "EJEMPLO"), ("B", "Licencias de software"), ("C", "Licencias Software Estadístico"),
                      ("D", "Licencia anual de software estadístico y mensajería de correos."),
                      ("E", "Jefe de Tecnologías Educativas")):
            x.put(f"{c_}7", v, "example")
        formulas(7, 7, 8, "example")
        recs = self.d["OO"]
        nxt = len(recs)
        for i, r in enumerate(range(fr, lr + 1)):
            m = recs[i] if i < len(recs) else None
            if m:
                vals = [m["codigo"], m["categoria"], m["nombre"], m["descripcion"], m["responsable"]]
            else:
                nxt += 1
                vals = [f"OO-{nxt:03d}", None, None, None, None]
            for j, v in enumerate(vals):
                x.put(f"{col(j + 1)}{r}", v or None, "input")
            x.put(f"J{r}", (m["observaciones"] if m else None) or None, "input")
            formulas(r, dfr, dlr, "calc")
        x.put(f"A{tr}", "TOTAL OTRAS ACTIVIDADES OPERATIVAS (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:E{tr}")
        for c_ in "FGH":
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        x.dv_list(f"B{fr}:B{lr}", "=L_CatOO")

    # ------------------------------------- Detalle_Otras_Operativas / Act_Prop
    def _det_generic(self, ws, kind: str, key: str, master_sheet: str):
        x = XL(ws)
        if kind == "OO":
            x.title("3. DETALLE DE OTRAS ACTIVIDADES OPERATIVAS",
                    'Varias filas por cada actividad calificada como "otras actividades operativas": un bien o servicio '
                    "por fila. Distribuya el total en los meses de ejecución. Los bienes de inversión se registran en "
                    "la hoja Inversion.")
        else:
            x.title("4. DETALLE DE ACTIVIDADES PROPIAS",
                    "Ingresos y costos de cada actividad propia (congresos, seminarios, talleres, cursos cortos): una "
                    "fila por concepto, distribuida en los meses de ejecución. Deben autofinanciarse.")
        x.group(5, "A", "C", "Actividad")
        x.group(5, "D", "K", "Ingresos / Bien o servicio y costo")
        x.group(5, "L", "W", "Distribución mensual del monto (S/)")
        x.group(5, "X", "AA", "Control y sustento")
        if kind == "OO":
            h_code, h_name, h_cu, h_cus = ("Código de iniciativa", "Nombre de la iniciativa", "Costo unitario (moneda)",
                                           "Costo unitario (S/)")
        else:
            h_code, h_name, h_cu, h_cus = ("Código", "Nombre de la iniciativa", "Ingresoss o Costo unitario (moneda)",
                                           "Ingresos o   Costo unitario  (S/)")
        x.headers(6, [h_code, h_name, "Tipo", "Partida / rubro", "Categoría", "Descripción del bien o servicio",
                      "Cantidad", "Moneda", h_cu, h_cus, "Total (S/)", *MESES, "Programado (S/)",
                      "Control de distribución", "Sustento", "Observaciones"])
        x.widths({"A": 10, "B": 32, "C": 10, "D": 36, "E": 22, "F": 40, "G": 9, "H": 7, "I": 14, "J": 14, "K": 14,
                  **{col(12 + i): 11 for i in range(12)}, "X": 14, "Y": 12, "Z": 22, "AA": 30})
        fr, lr, tr = self.L[key]
        mlr = self.L[kind][1]
        mname = "C"

        def formulas(r, k, ex=False):
            if not ex:
                x.put(f"B{r}", f'=IF(A{r}="","",IFERROR(INDEX({master_sheet}!${mname}$7:${mname}${mlr},MATCH(A{r},'
                               f'{master_sheet}!$A$7:$A${mlr},0))&"","Código no existe"))', "link")
            x.put(f"E{r}", self._cat_formula(f"D{r}"), "link" if k == "calc" else k)
            x.put(f"J{r}", f'=IF(I{r}="","",IF(H{r}="US$",I{r}*TC,I{r}))', k, MONEY)
            x.put(f"K{r}", f'=IF(OR(G{r}="",J{r}=""),0,G{r}*J{r})', k, MONEY)
            x.put(f"X{r}", f"=SUM(L{r}:W{r})", k, MONEY)
            x.put(f"Y{r}", f'=IF(AND(K{r}=0,X{r}=0),"",IF(ABS(X{r}-K{r})<0.01,"OK","Revisar"))', k)

        if kind == "OO":
            ex_rows = [
                {"A": "EJEMPLO", "B": "Licencias Software Estadístico", "C": "Ingresos", "D": "PENSION PEA",
                 "F": "Recaudación pensión", "G": 2, "H": "S/", "I": 1200, "N": 1200, "S": 1200, "Z": "Histórico de recaudación",
                 "AA": "Cobro en dos armadas (marzo y agosto)."},
                {"A": "EJEMPLO", "B": "Licencias Software Estadístico", "C": "Egresos", "D": "CORREOS",
                 "F": "Licencia mensual mensajería correos", "G": 12, "H": "S/", "I": 60,
                 **{col(12 + k): 60 for k in range(12)}, "Z": "Cotización", "AA": "Pago en doce meses (ene-dic)"},
            ]
        else:
            ex_rows = [
                {"A": "EJEMPLO", "B": "Seminario internacional de innovación educativa", "C": "Ingresos",
                 "D": "CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO", "F": "Inscripción", "G": 150, "H": "S/",
                 "I": 50, "N": 5000, "S": 2500, "Z": "Tarifario institucional vigente",
                 "AA": "Cobro en dos armadas (marzo y agosto)."},
                {"A": "EJEMPLO", "B": "Seminario internacional de innovación educativa", "C": "Egresos",
                 "D": "SERVICIO DESARROLLO DE TALLERES", "F": "Pago ponencia", "G": 3, "H": "S/", "I": 250, "N": 500,
                 "S": 250, "Z": "Cotización", "AA": "Pago en dos armadas (marzo y agosto)."},
            ]
        for i, exr in enumerate(ex_rows):
            for c_, v in exr.items():
                x.put(f"{c_}{7 + i}", v, "example")
            formulas(7 + i, "example", ex=True)
        rows = [(m, dd) for m in self.d[kind] for dd in m["_detalle"]]
        self._detail_rows(x, rows, fr, lr, formulas, letters=dict(code="A", tipo="C", partida="D", desc="F", cant="G",
                                                                     mon="H", cu="I", m0="L", sust="Z", obs="AA"))
        x.put(f"A{tr}", "TOTAL (no incluye las filas de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:J{tr}")
        for i in range(14):
            c_ = col(11 + i)
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        a, b = self.L[kind][0], self.L[kind][1]
        x.dv_list(f"A{fr}:A{lr}", f"={master_sheet}!$A${a}:$A${b}")
        x.dv_list(f"C{fr}:C{lr}", '"Ingresos,Egresos"')
        x.dv_list(f"D{fr}:D{lr}", "=L_PartidasNIIF")
        x.dv_list(f"H{fr}:H{lr}", '"S/,US$"')
        x.dv_list(f"Z{fr}:Z{lr}", "=L_SustAll")
        x.status_cf(f"Y7:Y{lr}")

    # --------------------------------------- Otras_Operativas (oculta, legado)
    def _otras_operativas_legacy(self, ws):
        x = XL(ws)
        x.title("3. OTRAS ACTIVIDADES OPERATIVAS",
                "Vista plana (formato anterior) generada desde Detalle_Otras_Operativas. Marque con «x» los meses de pago.")
        x.group(5, "A", "K", "Concepto y costo")
        x.group(5, "L", "X", "Meses de pago (marque x)")
        x.group(5, "Y", "AK", "Monto mensual calculado (S/)")
        x.group(5, "AL", "AM", "Sustento")
        x.headers(6, ["Código", "Categoría", "Concepto / descripción", "Partida / rubro", "Proveedor / contrato",
                      "Unidad de medida", "Cantidad", "Moneda", "Costo unitario por pago (moneda)",
                      "Costo unitario por pago (S/)", "Responsable", *MESES, "N° de pagos",
                      *[f"{m} (S/)" for m in MESES], "Total anual (S/)", "Sustento", "Observaciones"])
        fr, lr, tr = self.L["OOL"]
        rows = [(m, dd) for m in self.d["OO"] for dd in m["_detalle"] if dd["tipo"] == "Egresos"]
        for i, r in enumerate(range(fr, lr + 1)):
            item = rows[i] if i < len(rows) else None
            if item:
                m, dd = item
                vals = {"A": m["codigo"], "B": m["categoria"], "C": dd["descripcion"], "D": dd["partida"],
                        "G": to_float(dd["cantidad"]), "H": dd["moneda"], "I": to_float(dd["costo_unit_moneda"]),
                        "K": m["responsable"], "AL": dd["sustento"], "AM": dd["observaciones"]}
                for c_, v in vals.items():
                    x.put(f"{c_}{r}", v or None, "input")
                for k in range(12):
                    v = dd["_c"]["meses"][k]
                    x.put(f"{col(12 + k)}{r}", "x" if v else None, "input")
                    x.put(f"{col(25 + k)}{r}", v or 0, "calc", MONEY)
            x.put(f"J{r}", f'=IF(I{r}="","",IF(H{r}="US$",I{r}*TC,I{r}))', "calc", MONEY)
            x.put(f"X{r}", f'=COUNTIF(L{r}:W{r},"x")', "calc", INT)
            x.put(f"AK{r}", f"=SUM(Y{r}:AJ{r})", "calc", MONEY)
        x.put(f"A{tr}", "TOTAL OTRAS ACTIVIDADES OPERATIVAS (no incluye la fila de ejemplo)", "total")
        for k in range(13):
            c_ = col(25 + k)
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)

    # ------------------------------------------------------------- Act_Prop
    def _act_prop(self, ws):
        x = XL(ws)
        x.title("4. ACTIVIDADES PROPIAS",
                "Actividades de corta duración que generan ingresos (congresos, seminarios, talleres). Deben "
                "autofinanciarse: el ingreso proyectado debe cubrir sus costos.")
        x.group(5, "A", "C", "Identificación")
        x.group(5, "D", "J", "Resultado (calculado desde Detalle_Act_Prop)")
        x.group(5, "K", "L", "Gestión y revisión")
        x.group(5, "M", "Q", "Datos de la actividad y tarifa")
        x.headers(6, ["Código", "Tipo", "Nombre de la actividad", "Ingresos Total (S/)", "Costo total (S/)",
                      "Resultado (S/)", "Margen (%)", "Punto de equilibrio (participantes)",
                      "Participantes proyectados", "Estado", "Responsable", "Observaciones", "Público objetivo",
                      "Mes de ejecución", "Tarifa por participante (S/)", "Becas y descuentos (%)",
                      "Tarifa neta (S/)"])
        x.widths({"A": 10, "B": 14, "C": 40, "D": 14, "E": 14, "F": 14, "G": 10, "H": 14, "I": 12, "J": 14, "K": 26,
                  "L": 30, "M": 24, "N": 10, "O": 14, "P": 10, "Q": 12})
        fr, lr, tr = self.L["AP"]
        dfr, dlr, _ = self.L["DET_AP"]

        def formulas(r, a, b, kind, part_value=None):
            det = "Detalle_Act_Prop"
            for c_, tipo in (("D", "Ingresos"), ("E", "Egresos")):
                x.put(f"{c_}{r}", f"=SUMIFS({det}!$X${a}:$X${b},{det}!$A${a}:$A${b},$A{r},{det}!$C${a}:$C${b},\"{tipo}\")",
                      kind, MONEY)
            x.put(f"F{r}", f"=D{r}-E{r}", kind, MONEY)
            x.put(f"G{r}", f'=IF(D{r}=0,"",F{r}/D{r})', kind, PCT)
            qty = f'SUMIFS({det}!$G${a}:$G${b},{det}!$A${a}:$A${b},$A{r},{det}!$C${a}:$C${b},"Ingresos")'
            if part_value is None:
                x.put(f"I{r}", f"={qty}", kind, INT)
            else:
                x.put(f"I{r}", part_value, "input", INT)
            x.put(f"Q{r}", f'=IF(O{r}="",IF({qty}=0,"",D{r}/{qty}),O{r}*(1-IF(P{r}="",0,P{r})))', kind, MONEY)
            x.put(f"H{r}", f'=IF(OR(E{r}=0,Q{r}=""),"",IF(N(Q{r})<=0,"",ROUNDUP(E{r}/Q{r},0)))', kind, INT)
            x.put(f"J{r}", f'=IF(AND(D{r}=0,E{r}=0),"",IF(F{r}>=0,"Autofinanciada","Déficit"))', kind)

        for c_, v in (("A", "EJEMPLO"), ("B", "Seminario"), ("C", "Seminario internacional de innovación educativa"),
                      ("K", "Coordinación de Educación Continua"), ("M", "Docentes y egresados"), ("N", "Mar"),
                      ("O", 50)):
            x.put(f"{c_}7", v, "example")
        formulas(7, 7, 8, "example")
        recs = self.d["AP"]
        nxt = len(recs)
        for i, r in enumerate(range(fr, lr + 1)):
            m = recs[i] if i < len(recs) else None
            if m:
                vals = {"A": m["codigo"], "B": m["tipo"], "C": m["nombre"], "K": m["responsable"],
                        "L": m["observaciones"], "M": m["publico_objetivo"], "N": m["mes_ejecucion"],
                        "O": to_float(m["tarifa"]), "P": m["becas_pct"]}
                part = to_float(m["participantes"])
            else:
                nxt += 1
                vals = {"A": f"AP-{nxt:03d}"}
                part = None
            for c_ in "ABCKLMNOP":
                x.put(f"{c_}{r}", vals.get(c_) if vals.get(c_) != "" else None, "input",
                      MONEY if c_ == "O" else PCT if c_ == "P" else None)
            formulas(r, dfr, dlr, "calc", part)
        x.put(f"A{tr}", "TOTAL ACTIVIDADES PROPIAS (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:C{tr}")
        for c_ in "DEFI":
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", INT if c_ == "I" else MONEY)
        x.put(f"G{tr}", f'=IF(D{tr}=0,"",F{tr}/D{tr})', "total", PCT)
        x.dv_list(f"B{fr}:B{lr}", "=L_TipoAP")
        x.dv_list(f"N{fr}:N{lr}", '"' + ",".join(MESES) + '"')
        x.status_cf(f"J7:J{lr}")

    # ----------------------------------------- Act_Propias (oculta, legado)
    def _act_propias_legacy(self, ws):
        x = XL(ws)
        x.title("4. ACTIVIDADES PROPIAS", "Vista resumida (formato anterior) generada desde Act_Prop y Detalle_Act_Prop.")
        x.group(5, "A", "E", "Actividad")
        x.group(5, "F", "I", "Ingreso")
        x.group(5, "J", "O", "Costos directos")
        x.group(5, "P", "S", "Resultado (calculado)")
        x.group(5, "T", "V", "Gestión y sustento")
        x.headers(6, ["Código", "Tipo", "Nombre de la actividad", "Público objetivo", "Mes de ejecución",
                      "Participantes proyectados", "Tarifa por participante (S/)", "Becas y descuentos (%)",
                      "Ingreso proyectado (S/)", "Ponentes y honorarios (S/)", "Local y logística (S/)",
                      "Materiales (S/)", "Difusión (S/)", "Otros costos (S/)", "Costo total (S/)", "Resultado (S/)",
                      "Margen (%)", "Punto de equilibrio (participantes)", "Estado", "Responsable", "Sustento",
                      "Observaciones"])
        fr, lr, tr = self.L["APL"]
        recs = self.d["AP"]
        for i, r in enumerate(range(fr, lr + 1)):
            m = recs[i] if i < len(recs) else None
            if m:
                buckets = {"J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0}
                for dd in m["_detalle"]:
                    if dd["tipo"] == "Egresos":
                        buckets[_bucket_ap(dd)] += dd["_c"]["programado"]
                part = m["_participantes"]
                tarifa = to_float(m["tarifa"])
                if tarifa is None and part:
                    tarifa = m["_ingresos"] / part
                vals = {"A": m["codigo"], "B": m["tipo"], "C": m["nombre"], "D": m["publico_objetivo"],
                        "E": m["mes_ejecucion"], "F": part or None, "G": tarifa, "H": m["becas_pct"] if
                        to_float(m["tarifa"]) is not None else None, "T": m["responsable"], "U": m["sustento"],
                        "V": m["observaciones"], **{k: (v or None) for k, v in buckets.items()}}
                for c_, v in vals.items():
                    x.put(f"{c_}{r}", v if v != "" else None, "input")
            x.put(f"I{r}", f'=IF(OR(F{r}="",G{r}=""),0,F{r}*G{r}*(1-IF(H{r}="",0,H{r})))', "calc", MONEY)
            x.put(f"O{r}", f"=SUM(J{r}:N{r})", "calc", MONEY)
            x.put(f"P{r}", f"=I{r}-O{r}", "calc", MONEY)
            x.put(f"Q{r}", f'=IF(I{r}=0,"",P{r}/I{r})', "calc", PCT)
            x.put(f"R{r}", f'=IF(OR(G{r}="",O{r}=0),"",IF(G{r}*(1-IF(H{r}="",0,H{r}))<=0,"",'
                           f'ROUNDUP(O{r}/(G{r}*(1-IF(H{r}="",0,H{r}))),0)))', "calc", INT)
            x.put(f"S{r}", f'=IF(AND(I{r}=0,O{r}=0),"",IF(P{r}>=0,"Autofinanciada","Déficit"))', "calc")
        x.put(f"A{tr}", "TOTAL ACTIVIDADES PROPIAS (no incluye la fila de ejemplo)", "total")
        for c_ in "FIJKLMNOP":
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)

    # ------------------------------------------------------------ Inversion
    def _inversion(self, ws):
        x = XL(ws)
        x.title("5. ACTIVIDADES DE INVERSIÓN",
                "Adquisición e implementación de infraestructura física y tecnológica. Una fila por proyecto; el detalle "
                "de bienes u obras se registra en Detalle_Inversiones.")
        x.group(5, "A", "C", "Inversión y vinculación")
        x.group(5, "D", "G", "Control y sustento")
        x.headers(6, ["Código", "Tipo de inversión", "Proyecto de inversión", "Programado (S/)", "Prioridad (1–4)",
                      "Ficha de sustento adjunta", "Observaciones"])
        x.widths({"A": 10, "B": 28, "C": 52, "D": 16, "E": 12, "F": 14, "G": 36})
        fr, lr, tr = self.L["INV"]
        dfr, dlr, _ = self.L["DET_INV"]
        for c_, v in (("A", "EJEMPLO"), ("B", "Infraestructura tecnológica"),
                      ("C", "Renovación del laboratorio de cómputo 3"), ("E", 2), ("F", "Sí")):
            x.put(f"{c_}7", v, "example")
        x.put("D7", "=SUMIFS(Detalle_Inversiones!$W$7:$W$7,Detalle_Inversiones!$A$7:$A$7,$A7)", "example", MONEY)
        recs = self.d["INV"]
        nxt = len(recs)
        for i, r in enumerate(range(fr, lr + 1)):
            m = recs[i] if i < len(recs) else None
            if m:
                vals = {"A": m["codigo"], "B": m["tipo_inversion"], "C": m["proyecto"], "E": m["prioridad"],
                        "F": m["ficha_sustento"], "G": m["observaciones"]}
            else:
                nxt += 1
                vals = {"A": f"INV-{nxt:03d}"}
            for c_ in "ABCEFG":
                x.put(f"{c_}{r}", vals.get(c_) if vals.get(c_) != "" else None, "input")
            x.put(f"D{r}", f"=SUMIFS(Detalle_Inversiones!$W${dfr}:$W${dlr},Detalle_Inversiones!$A${dfr}:$A${dlr},$A{r})",
                  "calc", MONEY)
        x.put(f"A{tr}", "TOTAL INVERSIÓN (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:C{tr}")
        x.put(f"D{tr}", f"=SUM(D{fr}:D{lr})", "total", MONEY)
        x.dv_list(f"B{fr}:B{lr}", "=L_TipoINV")
        x.dv_list(f"E{fr}:E{lr}", '"1,2,3,4"')
        x.dv_list(f"F{fr}:F{lr}", '"Sí,No"')

    def _det_inversiones(self, ws):
        x = XL(ws)
        x.title("5. DETALLE DE LAS ACTIVIDADES DE INVERSIÓN",
                "Varias filas por proyecto: un bien u obra por fila (admite listas multilínea de equipos). Distribuya "
                "el total en los meses de ejecución.")
        x.group(5, "A", "C", "Proyecto")
        x.group(5, "D", "J", "Bien o servicio y costo")
        x.group(5, "K", "V", "Distribución mensual del monto (S/)")
        x.group(5, "W", "Z", "Control y sustento")
        x.headers(6, ["Código", "Nombre de la iniciativa", "Tipo", "Partida / rubro", "Descripción del bien o servicio",
                      "Cantidad", "Moneda", "Costo unitario (moneda)", "Costo unitario (S/)", "Total (S/)", *MESES,
                      "Programado (S/)", "Control de distribución", "Sustento", "Observaciones"])
        x.widths({"A": 10, "B": 34, "C": 10, "D": 36, "E": 56, "F": 9, "G": 7, "H": 14, "I": 14, "J": 14,
                  **{col(11 + i): 11 for i in range(12)}, "W": 14, "X": 12, "Y": 22, "Z": 30})
        fr, lr, tr = self.L["DET_INV"]
        ilr = self.L["INV"][1]

        def formulas(r, k, ex=False):
            if not ex:
                x.put(f"B{r}", f'=IF(A{r}="","",IFERROR(INDEX(Inversion!$C$7:$C${ilr},MATCH(A{r},Inversion!$A$7:'
                               f'$A${ilr},0))&"","Código no existe"))', "link")
            x.put(f"I{r}", f'=IF(H{r}="","",IF(G{r}="US$",H{r}*TC,H{r}))', k, MONEY)
            x.put(f"J{r}", f'=IF(OR(F{r}="",I{r}=""),0,F{r}*I{r})', k, MONEY)
            x.put(f"W{r}", f"=SUM(K{r}:V{r})", k, MONEY)
            x.put(f"X{r}", f'=IF(AND(J{r}=0,W{r}=0),"",IF(ABS(W{r}-J{r})<0.01,"OK","Revisar"))', k)

        for c_, v in (("A", "EJEMPLO"), ("B", "Renovación del laboratorio de cómputo 3"), ("C", "Inversion"),
                      ("D", "Equipos para procesamiento de información"), ("E", "Computadora"), ("F", 40), ("G", "S/"),
                      ("H", 1200), ("M", 48000), ("Y", "Cotización"), ("Z", "Entrega en marzo.")):
            x.put(f"{c_}7", v, "example")
        formulas(7, "example", ex=True)
        rows = [(m, dd) for m in self.d["INV"] for dd in m["_detalle"]]
        self._detail_rows(x, rows, fr, lr, formulas, letters=dict(code="A", tipo="C", partida="D", desc="E", cant="F",
                                                                     mon="G", cu="H", m0="K", sust="Y", obs="Z"))
        x.put(f"A{tr}", "TOTAL (no incluye la fila de ejemplo)", "total")
        ws.merge_cells(f"A{tr}:I{tr}")
        for i in range(14):
            c_ = col(10 + i)
            x.put(f"{c_}{tr}", f"=SUM({c_}{fr}:{c_}{lr})", "total", MONEY)
        a, b = self.L["INV"][0], self.L["INV"][1]
        x.dv_list(f"A{fr}:A{lr}", f"=Inversion!$A${a}:$A${b}")
        x.dv_list(f"D{fr}:D{lr}", "=Inversion")
        x.dv_list(f"G{fr}:G{lr}", '"S/,US$"')
        x.dv_list(f"Y{fr}:Y{lr}", "=L_Sustento")
        x.status_cf(f"X7:X{lr}")

    # ---------------------------------------------------------- Ficha_Resumen
    def _ficha_resumen(self, ws):
        x = XL(ws)
        u = self.u
        x.title("FICHA DE FORMULACIÓN DEL POA PRESUPUESTAL — RESUMEN",
                "Complete solo los datos de la unidad. Todo lo demás se calcula a partir de las hojas de registro.",
                unidad=False)
        x.widths({"A": 2, "B": 50, **{col(3 + i): 15 for i in range(13)}})
        x.group(5, "B", "C", "Datos de la unidad")
        datos = [("Unidad orgánica", u["nombre_unidad"]), ("Centro de costo", u["centro_costo"]),
                 ("Responsable de la unidad", u["responsable"]), ("Cargo", u["cargo"]),
                 ("Correo electrónico", u["correo"]), ("Año de formulación", None),
                 ("Fecha de presentación", u["fecha_presentacion"]), ("Versión de la ficha", u["version_ficha"])]
        for i, (lab, val) in enumerate(datos):
            r = 6 + i
            x.put(f"B{r}", lab, "label")
            if lab == "Año de formulación":
                x.put(f"C{r}", '=IF(AnioPOA="","(definir en Parámetros)",AnioPOA)', "link")
            else:
                x.put(f"C{r}", val or None, "input")
            ws.merge_cells(f"C{r}:H{r}")
        x.put("J5", "Código de área", "label")
        x.put("K5", u["codigo_area"], "input")
        x.put("J6", "Macrogrupo", "label")
        x.put("K6", u["macrogrupo"], "input")
        L = self.L
        T = lambda k: L[k][2]  # noqa: E731
        R = lambda sh, k, c: self._rng(sh, k, c)  # noqa: E731

        # ---- 1
        r = 15
        x.group(r, "B", "H", "1. Resumen por tipo de actividad")
        x.headers(r + 1, ["Tipo de actividad", "Ingresos (S/)", "Egresos (S/)", "Resultado (S/)",
                          "Techo de egresos (S/)", "% de uso del techo", "Estado"], 2)
        r1 = r + 2

        def est(rr, d_expr, f_ref):
            return (f'=IF({f_ref}=0,IF({d_expr}=0,"","Sin techo"),IF({d_expr}<={f_ref},"Dentro del techo",'
                    f'"Excede techo"))')

        lines = [
            ("1. Iniciativas con impacto estratégico", f"=Iniciativas!$T${T('INI')}", f"=Iniciativas!$U${T('INI')}",
             "=TechoIniciativas"),
            ("2. Actividades de funcionamiento", f"=Funcionamiento_Ing!$AK${T('FI')}",
             f"=Funcionamiento_Egr!$AI${T('FE')}", "=TechoFuncionamiento"),
            ("3. Otras actividades operativas", f"=Otras_Operat!$F${T('OO')}", f"=Otras_Operat!$G${T('OO')}",
             "=TechoOtras"),
        ]
        for i, (lab, ci, di, fi) in enumerate(lines):
            rr = r1 + i
            x.put(f"B{rr}", lab, "label")
            x.put(f"C{rr}", ci, "link", MONEY)
            x.put(f"D{rr}", di, "link", MONEY)
            x.put(f"E{rr}", f"=C{rr}-D{rr}", "calc", MONEY)
            x.put(f"F{rr}", fi, "link", MONEY)
            x.put(f"G{rr}", f'=IF(F{rr}=0,"",D{rr}/F{rr})', "calc", PCT)
            x.put(f"H{rr}", est(rr, f"D{rr}", f"F{rr}"), "calc")
        ra = r1 + 3
        x.put(f"B{ra}", "4. Actividades propias", "label")
        x.put(f"C{ra}", f"=Act_Prop!$D${T('AP')}", "link", MONEY)
        x.put(f"D{ra}", f"=Act_Prop!$E${T('AP')}", "link", MONEY)
        x.put(f"E{ra}", f"=C{ra}-D{ra}", "calc", MONEY)
        x.put(f"F{ra}", "Autofinanciada", "calc")
        x.put(f"G{ra}", None, "calc")
        x.put(f"H{ra}", f'=IF(AND(C{ra}=0,D{ra}=0),"",IF(E{ra}>=0,"Autofinanciada","Déficit"))', "calc")
        ro = ra + 1
        x.put(f"B{ro}", "Total OPEX", "total")
        for c_ in "CDE":
            x.put(f"{c_}{ro}", f"=SUM({c_}{r1}:{c_}{ra})", "total", MONEY)
        x.put(f"F{ro}", f"=F{r1}+F{r1 + 1}+F{r1 + 2}", "total", MONEY)
        x.put(f"G{ro}", f'=IF(F{ro}=0,"",(D{ro}-D{ra})/F{ro})', "total", PCT)
        x.put(f"H{ro}", est(ro, f"(D{ro}-D{ra})", f"F{ro}"), "total")
        ri = ro + 1
        x.put(f"B{ri}", "5. Actividades de inversión", "label")
        x.put(f"C{ri}", 0, "calc", MONEY)
        x.put(f"D{ri}", f"=Inversion!$D${T('INV')}", "link", MONEY)
        x.put(f"E{ri}", f"=C{ri}-D{ri}", "calc", MONEY)
        x.put(f"F{ri}", "=TechoInversion", "link", MONEY)
        x.put(f"G{ri}", f'=IF(F{ri}=0,"",D{ri}/F{ri})', "calc", PCT)
        x.put(f"H{ri}", est(ri, f"D{ri}", f"F{ri}"), "calc")
        rc = ri + 1
        x.put(f"B{rc}", "Total CAPEX", "total")
        x.put(f"C{rc}", f"=C{ri}", "total", MONEY)
        x.put(f"D{rc}", f"=D{ri}", "total", MONEY)
        x.put(f"E{rc}", f"=E{ri}", "total", MONEY)
        x.put(f"F{rc}", f"=F{ri}", "total", MONEY)
        x.put(f"G{rc}", f'=IF(F{rc}=0,"",D{rc}/F{rc})', "total", PCT)
        x.put(f"H{rc}", est(rc, f"D{rc}", f"F{rc}"), "total")
        x.put(f"B{rc + 1}", "El uso del techo total excluye las actividades propias, que deben autofinanciarse.", "plain")
        x.status_cf(f"H{r1}:H{rc}")
        s1_first, s1_last = r1, ri
        tot_opex_row, ap_row, inv_row = ro, ra, ri

        # ---- 2
        r = rc + 3
        x.group(r, "B", "G", "2. Ingresos proyectados")
        x.headers(r + 1, ["Categoría", "N° de actividades", "Ingreso bruto (S/)", "Becas, descuentos y morosidad (S/)",
                          "Ingreso neto (S/)", "% del total"], 2)
        r2 = r + 2
        n2 = len(CATEGORIAS_ING_RESUMEN) + 2
        t2 = r2 + n2
        fi_c, fi_d = R("Funcionamiento_Ing", "FI", "C"), R("Funcionamiento_Ing", "FI", "D")
        for i, cat in enumerate(CATEGORIAS_ING_RESUMEN):
            rr = r2 + i
            x.put(f"B{rr}", cat, "label")
            x.put(f"C{rr}", f'=COUNTIFS({fi_c},B{rr},{fi_d},"<>")', "calc", INT)
            x.put(f"D{rr}", f"=SUMIFS({R('Funcionamiento_Ing', 'FI', 'V')},{fi_c},B{rr})", "calc", MONEY)
            x.put(f"E{rr}", f"=SUMIFS({R('Funcionamiento_Ing', 'FI', 'AJ')},{fi_c},B{rr})", "calc", MONEY)
            x.put(f"F{rr}", f"=SUMIFS({R('Funcionamiento_Ing', 'FI', 'AI')},{fi_c},B{rr})", "calc", MONEY)
            x.put(f"G{rr}", f'=IF($F${t2}=0,"",F{rr}/$F${t2})', "calc", PCT)
        roo = r2 + len(CATEGORIAS_ING_RESUMEN)
        x.put(f"B{roo}", "Otras actividades operativas (PEAS, extracurriculares)", "label")
        x.put(f"C{roo}", f'=COUNTIFS({R("Otras_Operat", "OO", "C")},"<>",{R("Otras_Operat", "OO", "F")},">0")', "calc", INT)
        x.put(f"D{roo}", f"=Otras_Operat!$F${T('OO')}", "link", MONEY)
        x.put(f"E{roo}", f"=D{roo}-F{roo}", "calc", MONEY)
        x.put(f"F{roo}", f"=D{roo}", "calc", MONEY)
        x.put(f"G{roo}", f'=IF($F${t2}=0,"",F{roo}/$F${t2})', "calc", PCT)
        rap = roo + 1
        x.put(f"B{rap}", "Actividades propias (congresos, seminarios, talleres)", "label")
        x.put(f"C{rap}", f'=COUNTIFS({R("Act_Prop", "AP", "C")},"<>",{R("Act_Prop", "AP", "D")},">0")', "calc", INT)
        x.put(f"D{rap}", f"=Act_Prop!$D${T('AP')}", "link", MONEY)
        x.put(f"E{rap}", f"=D{rap}-F{rap}", "calc", MONEY)
        x.put(f"F{rap}", f"=D{rap}", "calc", MONEY)
        x.put(f"G{rap}", f'=IF($F${t2}=0,"",F{rap}/$F${t2})', "calc", PCT)
        x.put(f"B{t2}", "Total", "total")
        for c_ in "CDEF":
            x.put(f"{c_}{t2}", f"=SUM({c_}{r2}:{c_}{t2 - 1})", "total", INT if c_ == "C" else MONEY)
        x.put(f"G{t2}", f'=IF(F{t2}=0,"",SUM(G{r2}:G{t2 - 1}))', "total", PCT)
        x.put(f"B{t2 + 1}", "Control: ingresos de funcionamiento sin concepto asignado (S/)", "label")
        x.put(f"F{t2 + 1}", f"=Funcionamiento_Ing!$AI${T('FI')}-SUM(F{r2}:F{roo - 1})", "calc", MONEY)

        # ---- 3
        r = t2 + 3
        x.group(r, "B", "D", "3. Metas de ingresos y resultado de la unidad")
        x.headers(r + 1, ["Indicador", "Monto (S/)", "Estado / referencia"], 2)
        r3 = r + 2

        def meta_est(v, m):
            return f'=IF({m}=0,IF({v}=0,"","Sin meta asignada"),IF({v}>={m},"Alcanza la meta","Bajo la meta"))'

        rows3 = [
            ("Ingreso neto de funcionamiento", f"=Funcionamiento_Ing!$AK${T('FI')}", MONEY, None),
            ("Meta de ingresos de funcionamiento", "=MetaIngresos", MONEY, None),
            ("Cumplimiento de la meta de ingresos (%)", f'=IF(C{r3 + 1}=0,"",C{r3}/C{r3 + 1})', PCT,
             meta_est(f"C{r3}", f"C{r3 + 1}")),
            ("Ingresos de otras actividades operativas", f"=F{roo}", MONEY, None),
            ("Meta de ingresos de otras actividades operativas", "=MetaOtras", MONEY,
             meta_est(f"C{r3 + 3}", f"C{r3 + 4}")),
            ("Resultado de actividades propias", f"=E{ap_row}", MONEY, None),
            ("Meta de resultado de actividades propias", "=MetaPropias", MONEY,
             f'=IF(C{r3 + 6}=0,"",IF(C{r3 + 5}>=C{r3 + 6},"Alcanza la meta","Bajo la meta"))'),
            ("Ingresos totales", f"=C{tot_opex_row}", MONEY, None),
            ("Egresos totales", f"=D{tot_opex_row}", MONEY, None),
            ("Resultado de la unidad (ingresos − egresos)", f"=C{r3 + 7}-C{r3 + 8}", MONEY,
             f'=IF(AND(C{r3 + 7}=0,C{r3 + 8}=0),"",IF(C{r3 + 9}>=0,"Superávit","Déficit"))'),
            ("Margen sobre ingresos (%)", f'=IF(C{r3 + 7}=0,"",C{r3 + 9}/C{r3 + 7})', PCT, None),
        ]
        for i, (lab, f_, fmt, st) in enumerate(rows3):
            rr = r3 + i
            x.put(f"B{rr}", lab, "label")
            x.put(f"C{rr}", f_, "calc", fmt)
            x.put(f"D{rr}", st, "calc")
        x.status_cf(f"D{r3}:D{r3 + 10}")
        x.put(f"B{r3 + 11}", "Para unidades que no generan ingresos, el resultado equivale a su costo total.", "plain")

        # ---- 4
        r = r3 + 13
        x.group(r, "B", "H", "4. Iniciativas con impacto estratégico por eje")
        x.headers(r + 1, ["Eje estratégico", "N° de iniciativas", "Ingresos Total (S/)", "% del total ingresos",
                          "Costo total (S/)", "% del total costos", "Resultado (S/)"], 2)
        r4 = r + 2
        ejes = self.c.db.catalogo("EJE")
        t4 = r4 + len(ejes)
        e0 = self.P["eje"][0]
        ib, ic = R("Iniciativas", "INI", "B"), R("Iniciativas", "INI", "C")
        for i, _e in enumerate(ejes):
            rr = r4 + i
            x.put(f"B{rr}", f"=Parámetros!C{e0 + i}", "link")
            x.put(f"C{rr}", f'=COUNTIFS({ib},B{rr},{ic},"<>")', "calc", INT)
            x.put(f"D{rr}", f"=SUMIFS({R('Iniciativas', 'INI', 'T')},{ib},B{rr})", "calc", MONEY)
            x.put(f"E{rr}", f'=IF($D${t4}=0,"",D{rr}/$D${t4})', "calc", PCT)
            x.put(f"F{rr}", f"=SUMIFS({R('Iniciativas', 'INI', 'U')},{ib},B{rr})", "calc", MONEY)
            x.put(f"G{rr}", f'=IF($F${t4}=0,"",F{rr}/$F${t4})', "calc", PCT)
            x.put(f"H{rr}", f"=D{rr}-F{rr}", "calc", MONEY)
        x.put(f"B{t4}", "Total", "total")
        for c_ in "CDFH":
            x.put(f"{c_}{t4}", f"=SUM({c_}{r4}:{c_}{t4 - 1})", "total", INT if c_ == "C" else MONEY)
        for c_ in "EG":
            x.put(f"{c_}{t4}", f'=IF(SUM({c_}{r4}:{c_}{t4 - 1})=0,"",SUM({c_}{r4}:{c_}{t4 - 1}))', "total", PCT)
        x.put(f"B{t4 + 1}", "La inversión vinculada ya está incluida en «Actividades de inversión»; aquí solo se "
                            "muestra el gasto corriente.", "plain")

        # ---- 5
        r = t4 + 3
        x.group(r, "B", "D", "5. Egresos de funcionamiento por categoría")
        x.headers(r + 1, ["Categoría", "Monto anual (S/)", "% del total"], 2)
        r5 = r + 2
        t5 = r5 + len(CATEGORIAS_EGR_RESUMEN) + 1
        fe_ai, fe_c = R("Funcionamiento_Egr", "FE", "AI"), R("Funcionamiento_Egr", "FE", "C")
        for i, cat in enumerate(CATEGORIAS_EGR_RESUMEN):
            rr = r5 + i
            x.put(f"B{rr}", cat, "label")
            x.put(f"C{rr}", f"=SUMIFS({fe_ai},{fe_c},B{rr})", "calc", MONEY)
            x.put(f"D{rr}", f'=IF($C${t5}=0,"",C{rr}/$C${t5})', "calc", PCT)
        ro5 = t5 - 1
        x.put(f"B{ro5}", "OTRAS CATEGORÍAS (fuera de las 28 / sin categoría)", "label")
        x.put(f"C{ro5}", f"=Funcionamiento_Egr!$AI${T('FE')}-SUM(C{r5}:C{ro5 - 1})", "calc", MONEY)
        x.put(f"D{ro5}", f'=IF($C${t5}=0,"",C{ro5}/$C${t5})', "calc", PCT)
        x.put(f"B{t5}", "Total", "total")
        x.put(f"C{t5}", f"=SUM(C{r5}:C{ro5})", "total", MONEY)
        x.put(f"D{t5}", f'=IF(C{t5}=0,"",SUM(D{r5}:D{ro5}))', "total", PCT)

        # ---- 6
        r = t5 + 2
        x.group(r, "B", "G", "6. Otras actividades operativas por categoría")
        x.headers(r + 1, ["Categoría", "Ingresos Total (S/)", "% del total ingresos", "Costo total (S/)",
                          "% del total costos", "Resultado (S/)"], 2)
        r6 = r + 2
        cats = self.c.db.catalogo("CAT_OO")
        t6 = r6 + len(cats)
        c0 = self.P["cat_oo"][0]
        ob = R("Otras_Operat", "OO", "B")
        for i, _c in enumerate(cats):
            rr = r6 + i
            x.put(f"B{rr}", f"=Parámetros!C{c0 + i}", "link")
            x.put(f"C{rr}", f"=SUMIFS({R('Otras_Operat', 'OO', 'F')},{ob},B{rr})", "calc", MONEY)
            x.put(f"D{rr}", f'=IF($C${t6}=0,"",C{rr}/$C${t6})', "calc", PCT)
            x.put(f"E{rr}", f"=SUMIFS({R('Otras_Operat', 'OO', 'G')},{ob},B{rr})", "calc", MONEY)
            x.put(f"F{rr}", f'=IF($E${t6}=0,"",E{rr}/$E${t6})', "calc", PCT)
            x.put(f"G{rr}", f"=C{rr}-E{rr}", "calc", MONEY)
        x.put(f"B{t6}", "Total", "total")
        for c_ in "CEG":
            x.put(f"{c_}{t6}", f"=SUM({c_}{r6}:{c_}{t6 - 1})", "total", MONEY)
        x.put(f"D{t6}", f'=IF(C{t6}=0,"",SUM(D{r6}:D{t6 - 1}))', "total", PCT)
        x.put(f"F{t6}", f'=IF(E{t6}=0,"",SUM(F{r6}:F{t6 - 1}))', "total", PCT)

        # ---- 7
        r = t6 + 2
        x.group(r, "B", "H", "7. Actividades propias por tipo")
        x.headers(r + 1, ["Tipo", "N° de actividades", "Participantes", "Ingreso (S/)", "Costo (S/)", "Resultado (S/)",
                          "Estado"], 2)
        r7 = r + 2
        tipos = self.c.db.catalogo("TIPO_AP")
        t7 = r7 + len(tipos)
        a0 = self.P["tipo_ap"][0]
        apb = R("Act_Prop", "AP", "B")
        for i, _t in enumerate(tipos):
            rr = r7 + i
            x.put(f"B{rr}", f"=Parámetros!B{a0 + i}", "link")
            x.put(f"C{rr}", f'=COUNTIFS({apb},B{rr},{R("Act_Prop", "AP", "C")},"<>")', "calc", INT)
            x.put(f"D{rr}", f"=SUMIFS({R('Act_Prop', 'AP', 'I')},{apb},B{rr})", "calc", INT)
            x.put(f"E{rr}", f"=SUMIFS({R('Act_Prop', 'AP', 'D')},{apb},B{rr})", "calc", MONEY)
            x.put(f"F{rr}", f"=SUMIFS({R('Act_Prop', 'AP', 'E')},{apb},B{rr})", "calc", MONEY)
            x.put(f"G{rr}", f"=E{rr}-F{rr}", "calc", MONEY)
            x.put(f"H{rr}", f'=IF(AND(E{rr}=0,F{rr}=0),"",IF(G{rr}>=0,"Autofinanciada","Déficit"))', "calc")
        x.put(f"B{t7}", "Total", "total")
        for c_ in "CDEFG":
            x.put(f"{c_}{t7}", f"=SUM({c_}{r7}:{c_}{t7 - 1})", "total", INT if c_ in "CD" else MONEY)
        x.put(f"H{t7}", f'=IF(AND(E{t7}=0,F{t7}=0),"",IF(G{t7}>=0,"Autofinanciada","Déficit"))', "total")
        x.status_cf(f"H{r7}:H{t7}")

        # ---- 8
        r = t7 + 2
        x.group(r, "B", "F", "8. Actividades de inversión por tipo")
        x.headers(r + 1, ["Tipo de inversión", "N° de líneas", "Monto (S/)", "% del total",
                          "Con ficha de sustento (S/)"], 2)
        r8 = r + 2
        tinv = self.c.db.catalogo("TIPO_INV")
        t8 = r8 + len(tinv)
        v0 = self.P["tipo_inv"][0]
        vb, vd, vf = R("Inversion", "INV", "B"), R("Inversion", "INV", "D"), R("Inversion", "INV", "F")
        for i, _t in enumerate(tinv):
            rr = r8 + i
            x.put(f"B{rr}", f"=Parámetros!B{v0 + i}", "link")
            x.put(f"C{rr}", f'=COUNTIFS({vb},B{rr},{vd},">0")', "calc", INT)
            x.put(f"D{rr}", f"=SUMIFS({vd},{vb},B{rr})", "calc", MONEY)
            x.put(f"E{rr}", f'=IF($D${t8}=0,"",D{rr}/$D${t8})', "calc", PCT)
            x.put(f"F{rr}", f'=SUMIFS({vd},{vb},B{rr},{vf},"Sí")', "calc", MONEY)
        x.put(f"B{t8}", "Total", "total")
        for c_ in "CDF":
            x.put(f"{c_}{t8}", f"=SUM({c_}{r8}:{c_}{t8 - 1})", "total", INT if c_ == "C" else MONEY)
        x.put(f"E{t8}", f'=IF(D{t8}=0,"",SUM(E{r8}:E{t8 - 1}))', "total", PCT)

        # ---- 9
        r = t8 + 3
        x.group(r, "B", "I", "9. Iniciativas e inversiones por prioridad")
        x.headers(r + 1, ["Prioridad", "Categoría", "N° de iniciativas", "Ingresos Total (S/)", "Costo total (S/)",
                          "Total (S/)", "N° de inversiones", "Monto inversión (S/)"], 2)
        r9 = r + 2
        prios = self.c.db.catalogo("PRIORIDAD")
        t9 = r9 + len(prios)
        p0 = self.P["prio"][0]
        isx = R("Iniciativas", "INI", "S")
        ve = R("Inversion", "INV", "E")
        for i, pr in enumerate(prios):
            rr = r9 + i
            x.put(f"B{rr}", int(pr["codigo"]), "label", INT)
            x.put(f"C{rr}", f"=Parámetros!C{p0 + i}", "link")
            x.put(f"D{rr}", f'=COUNTIFS({isx},B{rr},{ic},"<>")', "calc", INT)
            x.put(f"E{rr}", f"=SUMIFS({R('Iniciativas', 'INI', 'T')},{isx},B{rr})", "calc", MONEY)
            x.put(f"F{rr}", f"=SUMIFS({R('Iniciativas', 'INI', 'U')},{isx},B{rr})", "calc", MONEY)
            x.put(f"G{rr}", f"=E{rr}-F{rr}", "calc", MONEY)
            x.put(f"H{rr}", f'=COUNTIFS({ve},B{rr},{vd},">0")', "calc", INT)
            x.put(f"I{rr}", f"=SUMIFS({vd},{ve},B{rr})", "calc", MONEY)
        x.put(f"B{t9}", "Total", "total")
        for c_ in "DEFGHI":
            x.put(f"{c_}{t9}", f"=SUM({c_}{r9}:{c_}{t9 - 1})", "total", INT if c_ in "DH" else MONEY)
        x.put(f"B{t9 + 1}", "Iniciativas sin prioridad / inversiones sin prioridad:", "label")
        x.put(f"D{t9 + 1}", f'=COUNTIFS({ic},"<>",{isx},"")', "calc", INT)
        x.put(f"H{t9 + 1}", f'=COUNTIFS({vd},">0",{ve},"")', "calc", INT)

        # ---- 10
        r = t9 + 3
        x.group(r, "B", "O", "10. Programación mensual de ingresos y egresos (S/)")
        x.headers(r + 1, ["Concepto", *MESES, "Total"], 2)
        r10 = r + 2

        def det_sum(sheet, key, first_letter, k, tipo):
            fr_, lr_, _ = L[key]
            mcol = shift(first_letter, k)
            return (f'=SUMIFS({sheet}!${mcol}${fr_}:${mcol}${lr_},{sheet}!$C${fr_}:$C${lr_},"{tipo}")')

        conceptos = [
            ("Ingresos — funcionamiento", lambda k: f"=Funcionamiento_Ing!{col(23 + k)}{T('FI')}"),
            ("Ingresos — otras actividades operativas",
             lambda k: det_sum("Detalle_Otras_Operativas", "DET_OO", "L", k, "Ingresos")),
            ("Ingresos — iniciativas con impacto estratégico",
             lambda k: det_sum("Detalle_Iniciativas", "DET_INI", "M", k, "Ingresos")),
            ("Ingresos — actividades propias", lambda k: det_sum("Detalle_Act_Prop", "DET_AP", "L", k, "Ingresos")),
            ("Total ingresos", lambda k, rr=r10: f"=SUM({col(3 + k)}{rr}:{col(3 + k)}{rr + 3})"),
            ("Egresos — iniciativas con impacto estratégico",
             lambda k: det_sum("Detalle_Iniciativas", "DET_INI", "M", k, "Egresos")),
            ("Egresos — funcionamiento", lambda k: f"=Funcionamiento_Egr!{col(23 + k)}{T('FE')}"),
            ("Egresos — otras actividades operativas",
             lambda k: det_sum("Detalle_Otras_Operativas", "DET_OO", "L", k, "Egresos")),
            ("Egresos — actividades propias", lambda k: det_sum("Detalle_Act_Prop", "DET_AP", "L", k, "Egresos")),
            ("Egresos — inversión", lambda k: f"=Detalle_Inversiones!{col(11 + k)}{T('DET_INV')}"),
            ("Total egresos", lambda k, rr=r10: f"=SUM({col(3 + k)}{rr + 5}:{col(3 + k)}{rr + 9})"),
            ("Flujo neto mensual (ingresos − egresos)",
             lambda k, rr=r10: f"={col(3 + k)}{rr + 4}-{col(3 + k)}{rr + 10}"),
            ("Flujo neto acumulado", lambda k, rr=r10: f"=SUM($C${rr + 11}:{col(3 + k)}{rr + 11})"),
            ("Egresos acumulados", lambda k, rr=r10: f"=SUM($C${rr + 10}:{col(3 + k)}{rr + 10})"),
            ("% acumulado de egresos", lambda k, rr=r10: f'=IF($O${rr + 10}=0,"",{col(3 + k)}{rr + 13}/$O${rr + 10})'),
        ]
        for i, (lab, fn) in enumerate(conceptos):
            rr = r10 + i
            kind = "total" if lab.startswith(("Total", "Flujo neto mensual")) else "calc"
            x.put(f"B{rr}", lab, "total" if kind == "total" else "label")
            fmt = PCT if lab.startswith("%") else MONEY
            for k in range(12):
                x.put(f"{col(3 + k)}{rr}", fn(k), kind, fmt)
            if lab in ("Flujo neto acumulado", "Egresos acumulados"):
                x.put(f"O{rr}", f"=N{rr}", kind, MONEY)
            elif lab.startswith("%"):
                x.put(f"O{rr}", f'=IF(O{r10 + 10}=0,"",O{r10 + 13}/O{r10 + 10})', kind, PCT)
            else:
                x.put(f"O{rr}", f"=SUM(C{rr}:N{rr})", kind, MONEY)
        rk = r10 + len(conceptos)
        x.put(f"B{rk}", "Control: ingresos y egresos programados por mes vs. totales de la sección 1", "label")
        x.put(f"O{rk}", f'=IF(AND(ROUND(O{r10 + 10}-(D{tot_opex_row}+D{inv_row}),2)=0,'
                        f'ROUND(O{r10 + 4}-C{tot_opex_row},2)=0),"OK","Revisar")', "calc")
        x.status_cf(f"O{rk}:O{rk}")

        # ---- 11
        r = rk + 2
        x.group(r, "B", "H", "11. Verificación automática de directrices")
        x.headers(r + 1, ["Directriz", "Verificación", "", "", "", "N° de observaciones", "Resultado"], 2)
        r11 = r + 2

        def sp(*terms):
            return "+".join(terms)

        I_ = lambda c_: R("Iniciativas", "INI", c_)  # noqa: E731
        DI = lambda c_: R("Detalle_Iniciativas", "DET_INI", c_)  # noqa: E731
        DO = lambda c_: R("Detalle_Otras_Operativas", "DET_OO", c_)  # noqa: E731
        DA = lambda c_: R("Detalle_Act_Prop", "DET_AP", c_)  # noqa: E731
        DV = lambda c_: R("Detalle_Inversiones", "DET_INV", c_)  # noqa: E731
        FE = lambda c_: R("Funcionamiento_Egr", "FE", c_)  # noqa: E731
        FI = lambda c_: R("Funcionamiento_Ing", "FI", c_)  # noqa: E731
        IV = lambda c_: R("Inversion", "INV", c_)  # noqa: E731
        AP = lambda c_: R("Act_Prop", "AP", c_)  # noqa: E731
        meses_fi = "+".join(f"({FI(col(10 + k))}<>\"\")" for k in range(12))
        reglas = [
            ("DF-01", "Iniciativas sin objetivo anual del POA, OE o AE del PEI",
             f'=SUMPRODUCT(({I_("C")}<>"")*((({I_("G")}="")+({I_("H")}="")+({I_("I")}=""))>0))'),
            ("DF-02", "Iniciativas sin indicador, meta o responsable, o con programación trimestral que no suma la meta",
             f'=SUMPRODUCT(({I_("C")}<>"")*((({I_("J")}="")+({I_("R")}="")+({I_("Q")}<>"OK"))>0))'),
            ("DP-01", "Tipos de actividad que exceden su techo de egresos",
             f'=COUNTIF(H{s1_first}:H{s1_last},"Excede techo")+COUNTIF(H{s1_first}:H{s1_last},"Sin techo")'),
            ("DP-02", "Gasto de iniciativas sin código o con código inexistente, e iniciativas sin gasto registrado",
             "=" + sp(f'SUMPRODUCT(({DI("L")}>0)*({DI("A")}=""))', f'COUNTIF({DI("B")},"Código no existe")',
                      f'SUMPRODUCT(({DI("L")}>0)*({DI("A")}<>"")*({DI("B")}=""))',
                      f'SUMPRODUCT(({I_("C")}<>"")*({I_("U")}=0))')),
            ("DP-03", "Egresos con monto y sin sustento (todas las hojas de egresos)",
             "=" + sp(f'SUMPRODUCT(({DI("L")}>0)*({DI("AA")}=""))', f'SUMPRODUCT(({FE("AI")}>0)*({FE("AK")}=""))',
                      f'SUMPRODUCT(({DO("K")}>0)*({DO("Z")}=""))', f'SUMPRODUCT(({DA("K")}>0)*({DA("Z")}=""))',
                      f'SUMPRODUCT(({DV("J")}>0)*({DV("Y")}=""))')),
            ("DP-04", "Montos no distribuidos por mes, o conceptos con costo y sin meses marcados",
             "=" + sp(f'COUNTIF({DI("Z")},"Revisar")', f'COUNTIF({DV("X")},"Revisar")', f'COUNTIF({DO("Y")},"Revisar")',
                      f'COUNTIF({DA("Y")},"Revisar")',
                      f'SUMPRODUCT(({FE("F")}<>"")*((({FE("V")}=0)+({FE("AJ")}="Revisar"))>0))')),
            ("DP-05", "Inversiones sin ficha de sustento adjunta, sin tipo o sin proyecto",
             f'=SUMPRODUCT(({IV("D")}>0)*((({IV("F")}<>"Sí")+({IV("B")}="")+({IV("C")}=""))>0))'),
            ("DI-01", "Ingresos de funcionamiento con cantidad y tarifa, sin meses de cobro",
             f'=SUMPRODUCT(({FI("D")}<>"")*({FI("G")}<>"")*(({meses_fi})=0))'),
            ("DI-02", "Ingresos de funcionamiento sin concepto o sin sustento",
             f'=SUMPRODUCT(({FI("AK")}>0)*((({FI("B")}="")+({FI("AL")}=""))>0))'),
            ("AP-01", "Actividades propias en déficit (no autofinanciadas)", f'=COUNTIF({AP("J")},"Déficit")'),
            ("AP-02", "Actividades propias sin tipo o sin mes de ejecución",
             f'=SUMPRODUCT(({AP("C")}<>"")*((({AP("B")}="")+({AP("N")}=""))>0))'),
        ]
        for i, (code, desc, f_) in enumerate(reglas):
            rr = r11 + i
            x.put(f"B{rr}", code, "label")
            x.put(f"C{rr}", desc, "plain")
            ws.merge_cells(f"C{rr}:F{rr}")
            x.put(f"G{rr}", f_, "calc", INT)
            x.put(f"H{rr}", f'=IF(G{rr}=0,"Cumple","Revisar")', "calc")
        x.status_cf(f"H{r11}:H{r11 + len(reglas) - 1}")

        # ---- 12
        r = r11 + len(reglas) + 1
        x.group(r, "B", "H", "12. Conformidad")
        x.put(f"C{r + 1}", "Elaborado por (responsable de la unidad)", "label")
        x.put(f"F{r + 1}", "Revisado por (Director de Planificación y Desarrollo Institucional - DPDI)", "label")
        ws.merge_cells(f"C{r + 1}:E{r + 1}")
        ws.merge_cells(f"F{r + 1}:H{r + 1}")
        conf = self.summary["sections"][-1]["rows"]
        for i, (lab, a, b) in enumerate(conf):
            rr = r + 2 + i
            x.put(f"B{rr}", lab, "label")
            x.put(f"C{rr}", a or None, "input")
            x.put(f"F{rr}", b or None, "input")
            ws.merge_cells(f"C{rr}:E{rr}")
            ws.merge_cells(f"F{rr}:H{rr}")
            if lab == "Firma":
                ws.row_dimensions[rr].height = 40

    # ---------------------------------------------------------- Instrucciones
    def _instrucciones(self, ws):
        x = XL(ws)
        prm = self.prm
        x.title("FICHA DE FORMULACIÓN DEL POA PRESUPUESTAL",
                "Proceso Gestión Estratégica (GD-DO-01) · Actividad: Definición de Parámetros y Directrices para la "
                "Planificación Operativa (POA) y Presupuesto · Insumo para PP-PR-02", unidad=False)
        x.widths({"A": 2, "B": 30, "C": 90, "D": 36})
        for r, lab, val in ((4, "Institución", prm.get("institucion")), (5, "Código del formato", self.u["codigo_area"]),
                            (6, "Versión", prm.get("version_formato")), (7, "Fecha de aprobación", prm.get("fecha_aprobacion"))):
            x.put(f"B{r}", lab, "label")
            x.put(f"C{r}", val or None, "plain")
        x.group(9, "B", "D", "¿Para qué sirve esta ficha?")
        x.put("B10", "Cada unidad orgánica (centro de costo) la usa para formular su POA y su presupuesto del año, "
                     "aplicando los parámetros, techos y directrices comunicados por la DPDI en el Instructivo de "
                     "formulación. Las actividades se registran según la siguiente clasificación; cada tipo tiene su "
                     "propia hoja.", "plain")
        ws.merge_cells("B10:D10")
        ws.row_dimensions[10].height = 42
        x.group(12, "B", "D", "Clasificación de actividades")
        x.headers(13, ["Tipo de actividad", "Definición", "Hoja(s) donde se registra"], 2)
        r = 14
        for it in self.c.db.catalogo("TIPO_ACTIVIDAD"):
            x.put(f"B{r}", it["codigo"], "label")
            x.put(f"C{r}", it["descripcion"], "plain")
            x.put(f"D{r}", it["extra"], "plain")
            r += 1
        x.group(20, "B", "D", "Hojas y quién las completa")
        x.headers(21, ["Hoja", "Contenido", "Completa"], 2)
        hojas = [
            ("Parámetros", "Año, techos de egresos por tipo de actividad, metas de ingresos, supuestos, objetivos del PEI "
                           "y del POA, y listas desplegables.", "DPDI"),
            ("Ficha_Resumen", "Datos de la unidad. Todo lo demás se calcula: resumen por tipo de actividad, ingresos, "
                              "resultado, distribución, programación mensual y verificación de directrices.",
             "Unidad (solo datos)"),
            ("Iniciativas", "Una fila por iniciativa con impacto estratégico: objetivo anual del POA, OE/AE del PEI, "
                            "indicador, meta trimestral, responsable y prioridad.", "Unidad"),
            ("Detalle_Iniciativas", "Gasto corriente de cada iniciativa (bienes y servicios) con distribución mensual. "
                                    "Los bienes de inversión de una iniciativa se registran en Inversion.", "Unidad"),
            ("Funcionamiento_Ing", "Ingresos regulares: matrículas, pensiones, tasas educativas y otros. Solo unidades "
                                   "que generan ingresos.", "Unidad"),
            ("Funcionamiento_Egr", "Remuneraciones, servicios básicos y regulares, depreciaciones y tributos. Se marcan "
                                   "con «x» los meses de pago.", "Unidad"),
            ("Otras_Operat", "Licencias, suscripciones, marketing, mantenimientos y similares, sin objetivos asociados "
                             "(detalle en Detalle_Otras_Operativas).", "Unidad"),
            ("Act_Prop", "Congresos, seminarios, talleres y otras actividades cortas que generan ingresos: ingreso, "
                         "costos, resultado y punto de equilibrio (detalle en Detalle_Act_Prop).", "Unidad"),
            ("Inversion", "Adquisición e implementación de infraestructura física y tecnológica, con distribución "
                          "mensual y ficha de sustento (detalle en Detalle_Inversiones).", "Unidad"),
            ("BD_CONSOLIDADA_NIIF", "Tabla plana mensualizada (1 fila por registro y mes) para carga masiva al ERP/NIIF "
                                    "y modelos de Power BI.", "Sistema POA-DPDI"),
        ]
        for i, (a, b, c_) in enumerate(hojas):
            x.put(f"B{22 + i}", a, "label")
            x.put(f"C{22 + i}", b, "plain")
            x.put(f"D{22 + i}", c_, "plain")
        r = 22 + len(hojas) + 1
        x.group(r, "B", "D", "Pasos para completar")
        pasos = [
            "Verifique que la DPDI haya completado Parámetros (techos, metas y supuestos). No modifique esa hoja salvo "
            "indicación de la DPDI.",
            "Complete los datos de la unidad en Ficha_Resumen.",
            "Clasifique cada actividad según la tabla «Clasificación de actividades» y regístrela en su hoja. Una misma "
            "actividad no debe registrarse en dos hojas.",
            "En las hojas de funcionamiento, marque con «x» los meses de pago o cobro (o el factor de prorrateo). En las "
            "hojas de detalle, distribuya el monto en los meses de ejecución.",
            "Revise en Ficha_Resumen la sección 1 (techos por tipo de actividad) y la sección 11 (verificación de "
            "directrices): todo debe mostrar «Cumple» o «Dentro del techo».",
            "Firme la conformidad y remita la ficha a la DPDI, o regístrela en SGEI, dentro del plazo indicado en "
            "Parámetros.",
        ]
        for i, t in enumerate(pasos):
            x.put(f"B{r + 1 + i}", f"Paso {i + 1}", "label")
            x.put(f"C{r + 1 + i}", t, "plain")
        r = r + len(pasos) + 2
        x.group(r, "B", "D", "Leyenda de celdas")
        ley = [("Texto azul", "input", "Celda para completar (fondo amarillo). Muchas tienen lista desplegable."),
               (1234, "calc", "Celda calculada (fondo gris). No la modifique: contiene fórmulas."),
               ("Texto verde", "link", "Dato traído de otra hoja (por ejemplo, de Parámetros)."),
               ("Fila de ejemplo", "example", "Muestra cómo llenar cada hoja. No se suma en los totales."),
               ("OK / Cumple", "ok", "Control conforme."), ("Revisar", "bad", "Control con observación: corrija antes de remitir.")]
        for i, (a, kind, b) in enumerate(ley):
            rr = r + 1 + i
            if kind in ("ok", "bad"):
                c_ = x.put(f"B{rr}", a, "plain")
                c_.fill = FILL_OK if kind == "ok" else FILL_BAD
            else:
                x.put(f"B{rr}", a, kind, MONEY if kind == "calc" else None)
            x.put(f"C{rr}", b, "plain")


def _bucket_ap(dd: dict) -> str:
    """Clasifica un egreso de actividad propia en las columnas del formato anterior (Act_Propias)."""
    txt = norm_key(f"{dd['partida']} {dd['categoria']} {dd['descripcion']}")
    if any(k in txt for k in ("PONENT", "HONORAR", "TALLER", "CONFEREN", "EXPOSITOR", "DOCEN")):
        return "J"
    if any(k in txt for k in ("ALQUIL", "LOCAL", "ALOJAM", "ALIMENT", "TRANSPORTE", "PASAJE", "LOGIST")):
        return "K"
    if any(k in txt for k in ("MATERIAL", "SUMINIST", "IMPRES", "IMPRENTA", "FOTOCOP", "ADQUISIC")):
        return "L"
    if any(k in txt for k in ("PUBLICID", "DIFUS", "PROMOCI", "PUBLICAC", "MARKETING")):
        return "M"
    return "N"


# =============================================================================
# 2. BD CONSOLIDADA NIIF
# =============================================================================
def write_bd_sheet(ws, rows: list[dict]) -> None:
    cols = POAController.BD_COLUMNS
    for j, h in enumerate(cols, start=1):
        c = ws.cell(1, j, h)
        c.font, c.fill, c.alignment = F_HDR, FILL_HDR, CENTER
    for i, row in enumerate(rows, start=2):
        for j, h in enumerate(cols, start=1):
            c = ws.cell(i, j, row.get(h))
            if h in ("Monto_Mensual_Soles", "Costo_Unitario_Soles"):
                c.number_format = MONEY
    widths = {"Unidad_Organica": 40, "Nombre_Actividad": 40, "Partida_Rubro": 36, "Descripcion_Item": 40,
              "Tipo_Actividad_POA": 30, "Categoria_Presupuestal": 26, "Estado_Directrices": 22, "Centro_Costo": 34}
    for j, h in enumerate(cols, start=1):
        ws.column_dimensions[col(j)].width = widths.get(h, 15)
    ws.freeze_panes = "A2"
    if rows:
        ref = f"A1:{col(len(cols))}{len(rows) + 1}"
        t = Table(displayName="BD_CONSOLIDADA_NIIF", ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)


def export_bd_consolidada(ctrl: POAController, uids: Iterable[int], path: str | Path, with_csv: bool = True) -> list[Path]:
    path = Path(path)
    import pandas as pd
    rows = ctrl.consolidated_rows(list(uids))
    wb = Workbook()
    ws = wb.active
    ws.title = "BD_CONSOLIDADA_NIIF"
    write_bd_sheet(ws, rows)
    df = pd.DataFrame(rows, columns=POAController.BD_COLUMNS)
    wp = wb.create_sheet("PIVOT_MENSUAL")
    if not df.empty:
        df["Flujo_Firmado"] = df.apply(lambda r: r["Monto_Mensual_Soles"] * (1 if r["Tipo_Flujo"] == "Ingresos"
                                                                             else -1), axis=1)
        pv = pd.pivot_table(df, index=["Codigo_Area", "Tipo_Actividad_POA", "Tipo_Flujo"], columns="Mes_Numero",
                            values="Monto_Mensual_Soles", aggfunc="sum", fill_value=0.0)
        pv = pv.reindex(columns=range(1, 13), fill_value=0.0)
        pv["Total"] = pv.sum(axis=1)
        headers = ["Codigo_Area", "Tipo_Actividad_POA", "Tipo_Flujo", *MESES, "Total"]
        for j, h in enumerate(headers, start=1):
            c = wp.cell(1, j, h)
            c.font, c.fill, c.alignment = F_HDR, FILL_HDR, CENTER
        for i, (idx, vals) in enumerate(pv.iterrows(), start=2):
            for j, v in enumerate(list(idx) + [round(float(x), 2) for x in vals.tolist()], start=1):
                c = wp.cell(i, j, v)
                if j > 3:
                    c.number_format = MONEY
        for j, w in enumerate([12, 34, 12] + [13] * 13, start=1):
            wp.column_dimensions[col(j)].width = w
        wp.freeze_panes = "D2"
    wb.save(path)
    out = [path]
    if with_csv:
        csv_path = path.with_suffix(".csv")
        df.to_csv(csv_path, sep=";", index=False, encoding="utf-8-sig", columns=POAController.BD_COLUMNS)
        out.append(csv_path)
    return out


def export_libro_oficial(ctrl: POAController, uid: int, path: str | Path) -> Path:
    return OfficialWorkbookExporter(ctrl, uid).export(path)


# =============================================================================
# 3. PDF — FICHA RESUMEN FIRMADA
# =============================================================================
def format_cell(value: Any, fmt: str) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value
    if fmt == "money":
        return fmt_money(value)
    if fmt == "pct":
        return fmt_pct(value)
    if fmt == "int":
        return f"{value:,.0f}"
    if fmt == "money_or_pct":
        return fmt_money(value)
    return str(value)


def export_pdf(ctrl: POAController, uid: int, path: str | Path) -> Path:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table as RLTable,
                                    TableStyle)

    path = Path(path)
    summary = ctrl.build_summary(uid)
    u, prm = summary["unidad"], summary["params"] or {}
    styles = getSampleStyleSheet()
    st_title = ParagraphStyle("t", parent=styles["Title"], fontSize=14, textColor=colors.HexColor("#" + NAVY),
                              spaceAfter=2)
    st_sub = ParagraphStyle("s", parent=styles["Normal"], fontSize=7.5, textColor=colors.HexColor("#595959"),
                            alignment=TA_CENTER)
    st_h = ParagraphStyle("h", parent=styles["Heading3"], fontSize=10, textColor=colors.HexColor("#" + NAVY),
                          spaceBefore=8, spaceAfter=3)
    st_cell = ParagraphStyle("c", parent=styles["Normal"], fontSize=6.8, leading=8)
    st_cell_b = ParagraphStyle("cb", parent=st_cell, fontName="Helvetica-Bold", textColor=colors.white)
    st_note = ParagraphStyle("n", parent=styles["Normal"], fontSize=6.8, textColor=colors.HexColor("#595959"))
    green, red = colors.HexColor("#C6EFCE"), colors.HexColor("#FFC7CE")
    ok_words = {"OK", "Cumple", "Dentro del techo", "Autofinanciada", "Superávit", "Alcanza la meta"}
    bad_words = {"Revisar", "Excede techo", "Déficit", "Sin techo", "Bajo la meta"}

    page_w = landscape(A4)[0] - 2.4 * cm
    story = [Paragraph(prm.get("institucion") or "", st_sub),
             Paragraph("FICHA DE FORMULACIÓN DEL POA PRESUPUESTAL — RESUMEN", st_title),
             Paragraph("Proceso Gestión Estratégica (GD-DO-01) · Insumo PP-PR-02 · Dirección de Planificación y "
                       "Desarrollo Institucional (DPDI)", st_sub), Spacer(1, 6)]
    datos = [["Unidad orgánica", u["nombre_unidad"], "Código de área", u["codigo_area"]],
             ["Centro de costo", u["centro_costo"], "Macrogrupo", u["macrogrupo"]],
             ["Responsable de la unidad", u["responsable"], "Año de formulación", summary["anio"]],
             ["Cargo", u["cargo"], "Fecha de presentación", u["fecha_presentacion"]],
             ["Correo electrónico", u["correo"], "Versión de la ficha", u["version_ficha"]]]
    t = RLTable([[Paragraph(f"<b>{a}</b>", st_cell), Paragraph(s(b), st_cell), Paragraph(f"<b>{c}</b>", st_cell),
                  Paragraph(s(d), st_cell)] for a, b, c, d in datos],
                colWidths=[page_w * 0.17, page_w * 0.43, page_w * 0.15, page_w * 0.25])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
                           ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#D9E1F2")),
                           ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#D9E1F2")),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [t, Spacer(1, 4)]

    for sec in summary["sections"]:
        if sec["key"] == "s12":
            continue
        headers, fmts, rows = sec["headers"], sec["formats"], sec["rows"]
        ncols = len(headers)
        if sec["key"] == "s10":
            widths = [page_w * 0.19] + [page_w * 0.81 / 13] * 13
        else:
            first = 0.34 if ncols > 4 else 0.45
            widths = [page_w * first] + [page_w * (1 - first) / (ncols - 1)] * (ncols - 1)
            if sec["key"] == "s11":
                widths = [page_w * 0.08, page_w * 0.62, page_w * 0.14, page_w * 0.16]
        data = [[Paragraph(h, st_cell_b) for h in headers]]
        style = [("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                 ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
                 ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                 ("FONTSIZE", (0, 0), (-1, -1), 6.8), ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                 ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]
        pct_rows = sec.get("pct_rows", set())
        for i, row in enumerate(rows, start=1):
            cells = []
            for j, v in enumerate(row):
                fmt = fmts[j] if j < len(fmts) else "text"
                if sec["key"] == "s3" and j == 1 and "(%)" in s(row[0]):
                    fmt = "pct"
                if (i - 1) in pct_rows and j > 0:
                    fmt = "pct"
                txt = format_cell(v, fmt)
                cells.append(Paragraph(txt, st_cell) if j == 0 or (sec["key"] == "s11" and j == 1) else txt)
                if isinstance(v, str) and v in ok_words:
                    style.append(("BACKGROUND", (j, i), (j, i), green))
                elif isinstance(v, str) and v in bad_words:
                    style.append(("BACKGROUND", (j, i), (j, i), red))
            data.append(cells)
            if (i - 1) in sec.get("totals", set()):
                style += [("BACKGROUND", (0, i), (-1, i), colors.HexColor("#D9D9D9")),
                          ("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")]
        tbl = RLTable(data, colWidths=widths, repeatRows=1)
        tbl.setStyle(TableStyle(style))
        block = [Paragraph(sec["title"], st_h), tbl]
        if sec.get("note"):
            block.append(Paragraph(sec["note"], st_note))
        if sec["key"] == "s11":
            obs = []
            for dct in summary["directrices"]:
                for it in dct["items"][:6]:
                    obs.append(f"<b>{dct['codigo']}</b>: {it}")
                if len(dct["items"]) > 6:
                    obs.append(f"<b>{dct['codigo']}</b>: … y {len(dct['items']) - 6} observaciones más.")
            if obs:
                block.append(Spacer(1, 3))
                block.append(Paragraph("Detalle de observaciones:", st_note))
                block += [Paragraph(o, st_note) for o in obs]
        story.append(KeepTogether(block[:3]) if len(block) >= 3 else KeepTogether(block))
        story += block[3:]

    conf = summary["sections"][-1]["rows"]
    firm = [[Paragraph("<b>Elaborado por (Responsable de la unidad)</b>", st_cell), "",
             Paragraph("<b>Revisado por (Director de Planificación y Desarrollo Institucional - DPDI)</b>", st_cell)],
            ["", "", ""],
            ["_" * 48, "", "_" * 48],
            [Paragraph(s(conf[0][1]), st_cell), "", Paragraph(s(conf[0][2]), st_cell)],
            [Paragraph(s(conf[1][1]), st_cell), "", Paragraph(s(conf[1][2]), st_cell)],
            [Paragraph(f"Fecha: {s(conf[3][1])}", st_cell), "", Paragraph(f"Fecha: {s(conf[3][2])}", st_cell)]]
    ft = RLTable(firm, colWidths=[page_w * 0.42, page_w * 0.16, page_w * 0.42],
                 rowHeights=[None, 1.4 * cm, None, None, None, None])
    ft.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                            ("FONTSIZE", (0, 0), (-1, -1), 7.5)]))
    story += [KeepTogether([Paragraph("12. Conformidad", st_h), Spacer(1, 4), ft])]

    generado = datetime.now().strftime("%Y-%m-%d %H:%M")

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(colors.HexColor("#7F7F7F"))
        canvas.drawString(1.2 * cm, 0.7 * cm, f"Sistema de Formulación Presupuestal y Operativa POA - DPDI · "
                                               f"{u['codigo_area']} · Generado: {generado}")
        canvas.drawRightString(landscape(A4)[0] - 1.2 * cm, 0.7 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.0 * cm, bottomMargin=1.2 * cm,
                            title=f"Ficha Resumen POA {summary['anio']} - {u['codigo_area']}", author="DPDI")
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return path
