"""
Pruebas automáticas del núcleo (sin interfaz gráfica).

    python -m unittest discover -s tests -v
"""
import math
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from controllers import (PRORRATEO_12, Params, POAController, calc_detalle, calc_fegr, calc_fing,  # noqa: E402
                         control_meta, distribuir, parse_pct, to_float)
from database import MCOLS, Database  # noqa: E402
from exporter import (OfficialWorkbookExporter, export_bd_consolidada, export_pdf,  # noqa: E402
                      export_plantilla_parametros)


def months(values):
    return {m: v for m, v in zip(MCOLS, values)}


class CalculosTest(unittest.TestCase):
    def test_conversion_y_control_detalle(self):
        p = Params(tc=3.5)
        r = dict(cantidad=10, moneda="US$", costo_unit_moneda=12.5, **months([None, None, 437.5] + [None] * 9))
        c = calc_detalle(r, p)
        self.assertAlmostEqual(c["costo_unit_soles"], 43.75)
        self.assertAlmostEqual(c["total"], 437.5)
        self.assertEqual(c["control"], "OK")
        r["m03"] = 400
        self.assertEqual(calc_detalle(r, p)["control"], "Revisar")
        self.assertEqual(calc_detalle({"cantidad": None}, p)["control"], "")

    def test_reajuste_por_mes(self):
        p = Params(reajuste=0.10, mes_reajuste=7)
        r = dict(cantidad=1, moneda="S/", costo_unit_moneda=1200, aplica_reajuste="Sí", **months([PRORRATEO_12] * 12))
        c = calc_fegr(r, p)
        self.assertAlmostEqual(c["meses"][0], 100.0)
        self.assertAlmostEqual(c["meses"][6], 110.0)
        self.assertAlmostEqual(c["total"], 6 * 100 + 6 * 110)
        self.assertEqual(c["control"], "OK")
        r["aplica_reajuste"] = "No"
        self.assertAlmostEqual(calc_fegr(r, p)["total"], 1200.0)

    def test_marcas_x_en_egresos(self):
        r = dict(cantidad=2, moneda="S/", costo_unit_moneda=50, aplica_reajuste="No", **months(["x", "x"] + [None] * 10))
        c = calc_fegr(r, Params())
        self.assertAlmostEqual(c["total"], 200.0)
        self.assertEqual(c["n_pagos"], 2)
        self.assertEqual(c["control"], "OK")

    def test_ingresos_becas_y_morosidad(self):
        p = Params(morosidad=0.05)
        r = dict(cantidad=100, moneda="S/", tarifa_moneda=500, becas_pct=0.10, aplica_morosidad="Sí",
                 **months([None, None, "x", "x", "x", "x", "x"] + [None] * 5))
        c = calc_fing(r, p)
        self.assertAlmostEqual(c["bruto"], 250000.0)
        self.assertAlmostEqual(c["neto"], 250000 * 0.9 * 0.95)
        self.assertAlmostEqual(c["descuentos"], 250000 - 250000 * 0.9 * 0.95)
        self.assertEqual(c["n_meses"], 5)

    def test_utilidades(self):
        self.assertEqual(parse_pct("10%"), 0.10)
        self.assertEqual(parse_pct("10"), 0.10)
        self.assertEqual(parse_pct("0.1"), 0.1)
        self.assertEqual(to_float("1,234.50"), 1234.5)
        d = distribuir(100, [1, 2, 3])
        self.assertAlmostEqual(sum(x for x in d if x), 100.0)
        self.assertEqual(control_meta({"nombre": "x", "meta_anual": 60, "t1": 10, "t2": 20, "t3": 30}), "OK")
        self.assertEqual(control_meta({"nombre": "x", "meta_anual": 60, "t1": 10}), "Revisar")
        self.assertEqual(control_meta({"nombre": "x"}), "Sin meta")


class FlujoCompletoTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Database(Path(self.tmp.name) / "t.sqlite3")
        self.c = POAController(self.db)
        self.db.save_parametros({"tipo_cambio": 3.5, "tasa_morosidad_pct": 0.05, "reajuste_remunerativo_pct": 0.03,
                                 "mes_inicio_reajuste": 4})
        self.uid = self.c.save_unidad({"codigo_area": "TST", "macrogrupo": "Pregrado", "nombre_unidad": "UNIDAD DE PRUEBA",
                                       "centro_costo": "CC PRUEBA", "genera_ingresos": 1, "techo_iniciativas": 10000,
                                       "techo_funcionamiento": 50000, "techo_otras_operativas": 5000,
                                       "techo_inversion": 3000, "techo_planilla_adm": 30000,
                                       "meta_ingresos_func": 100000, "meta_resultado_propias": 500})
        c, uid = self.c, self.uid
        ini = c.save_master("INI", uid, {"eje": "EJ-01 · Adaptación académica", "nombre": "Aulas virtuales",
                                         "origen": "Otro", "obj_poa": "OBJ-01 · Objetivo anual 01 del POA",
                                         "indicador": "Cursos", "unidad_medida": "Cantidad",
                                         "meta_anual": 10, "t1": 5, "t2": 5,
                                         "responsable": "Jefe TI", "prioridad": "2"})
        c.save_detalle("INI", ini, {"tipo": "Egresos", "partida": "MATERIAL DE OFICINA", "cantidad": 10, "moneda": "US$",
                                    "costo_unit_moneda": 100, "sustento": "Cotización", "m03": 1750, "m08": 1750})
        c.save_master("INI", uid, {"nombre": "Sin alineamiento", "eje": "Investigación"})
        c.save_flat("FI", uid, {"partida": "PENSION PREGRADO", "cantidad": 100, "tarifa_moneda": 300, "becas_pct": "10%",
                                "aplica_morosidad": "Sí", "sustento": "Histórico de recaudación",
                                **months(["x"] * 10 + [None] * 2)})
        c.save_flat("FE", uid, {"partida": "AGUA - DIRECTO", "cantidad": 1, "costo_unit_moneda": 2400,
                                "aplica_reajuste": "No", "sustento": "Histórico de ejecución",
                                **months([PRORRATEO_12] * 12)})
        c.save_flat("FE", uid, {"partida": "INTERNET", "cantidad": 1, "costo_unit_moneda": 1200, "aplica_reajuste": "No"})
        oo = c.save_master("OO", uid, {"categoria": "OO-01 · Licencias de software", "nombre": "Licencia estadística"})
        c.save_detalle("OO", oo, {"tipo": "Egresos", "partida": "LICENCIAS Y DERECHOS DE VIGENCIA", "cantidad": 12,
                                  "costo_unit_moneda": 500, "sustento": "Cotización", **months([500] * 12)})
        ap = c.save_master("AP", uid, {"tipo": "Seminario", "nombre": "Seminario de innovación", "mes_ejecucion": "Set",
                                       "participantes": 40, "tarifa": 100, "becas_pct": "10%"})
        c.crear_ingreso_desde_tarifa(ap)
        c.save_detalle("AP", ap, {"tipo": "Egresos", "partida": "SERVICIO DESARROLLO DE TALLERES", "cantidad": 1,
                                  "costo_unit_moneda": 4000, "sustento": "Cotización", "m09": 4000})
        inv = c.save_master("INV", uid, {"tipo_inversion": "Infraestructura tecnológica", "proyecto": "Laptops",
                                         "prioridad": "1", "ficha_sustento": "No"})
        c.save_detalle("INV", inv, {"partida": "Equipos para procesamiento de información", "cantidad": 1,
                                    "costo_unit_moneda": 4500, "sustento": "Costo de referencia", "m02": 4500})

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_resumen_y_directrices(self):
        summ = self.c.build_summary(self.uid)
        s1 = summ["s1"]
        self.assertAlmostEqual(s1["ini"][1], 3500.0)
        self.assertAlmostEqual(s1["fun"][0], 100 * 300 * 10 * 0.9 * 0.95)
        self.assertAlmostEqual(s1["fun"][1], 2400.0)
        self.assertAlmostEqual(s1["oo"][1], 6000.0)
        self.assertEqual(s1["rows"][2][6], "Excede techo")
        ap = self.c.list_masters("AP", self.uid)[0]
        self.assertAlmostEqual(ap["_ingresos"], 3600.0)
        self.assertEqual(ap["_punto_equilibrio"], math.ceil(4000 / 90))
        self.assertEqual(ap["_estado"], "Déficit")
        dirs = {d["codigo"]: d["n"] for d in summ["directrices"]}
        self.assertEqual(dirs["DF-01"], 1)
        self.assertEqual(dirs["DF-02"], 1)
        self.assertEqual(dirs["DP-01"], 2)
        self.assertEqual(dirs["DP-02"], 1)
        self.assertEqual(dirs["DP-03"], 0)
        self.assertEqual(dirs["DP-04"], 1)
        self.assertEqual(dirs["DP-05"], 1)
        self.assertEqual(dirs["DI-01"], 0)
        self.assertEqual(dirs["AP-01"], 1)
        self.assertEqual(dirs["AP-02"], 0)
        self.assertEqual(summ["s10"]["control"], "OK")

    def test_codigos_y_plantilla(self):
        creados = self.c.generar_codigos(self.uid, {"INI": 2, "FE": 1})
        self.assertEqual([x[1] for x in creados], ["INI-03", "INI-04", "FE-003"])
        self.assertEqual(creados[0][2], "TST-INI-03")
        rep = self.c.correr_funcionamiento(self.uid)
        self.assertGreater(rep["FE"], 10)
        fe = [r for r in self.c.list_flat("FE", self.uid) if r["origen_plantilla"]]
        planilla_adm = sum(r["_c"]["coste_total"] for r in fe if r["categoria"] == "PLANILLA ADMINISTRATIVA")
        self.assertAlmostEqual(planilla_adm, 30000.0, delta=0.05)
        self.assertTrue(all(abs(float(r["m01"]) - PRORRATEO_12) < 1e-12 for r in fe))
        rep2 = self.c.correr_funcionamiento(self.uid)
        self.assertEqual(rep2["eliminadas"], rep["FE"] + rep["FI"])

    def test_exportar_e_importar_ida_y_vuelta(self):
        out = Path(self.tmp.name)
        libro = OfficialWorkbookExporter(self.c, self.uid).export(out / "libro.xlsx")
        import openpyxl
        wb = openpyxl.load_workbook(libro)
        self.assertEqual(wb.sheetnames[:16], ["Instrucciones", "Ficha_Resumen", "Parámetros", "Hoja6",
                                              "Funcionamiento_Ing", "Funcionamiento_Egr", "Iniciativas",
                                              "Detalle_Iniciativas", "Otras_Operat", "Detalle_Otras_Operativas",
                                              "Otras_Operativas", "Act_Prop", "Detalle_Act_Prop", "Act_Propias",
                                              "Inversion", "Detalle_Inversiones"])
        self.assertEqual(wb.sheetnames[16], "BD_CONSOLIDADA_NIIF")
        bd = wb["BD_CONSOLIDADA_NIIF"]
        self.assertEqual([c.value for c in bd[1]], POAController.BD_COLUMNS)
        before = self.c.build_summary(self.uid)
        uid2 = self.c.save_unidad({"codigo_area": "TST2", "nombre_unidad": "COPIA", "macrogrupo": "Pregrado",
                                   "genera_ingresos": 1})
        self.c.import_ficha(libro, unidad_id=uid2)
        after = self.c.build_summary(uid2)
        for a, b in zip(before["s1"]["rows"], after["s1"]["rows"]):
            for x, y in zip(a[1:4], b[1:4]):
                self.assertAlmostEqual(x, y, places=2)
        self.assertEqual([d["n"] for d in before["directrices"]], [d["n"] for d in after["directrices"]])
        self.assertTrue(export_pdf(self.c, self.uid, out / "ficha.pdf").stat().st_size > 5000)
        files = export_bd_consolidada(self.c, [self.uid, uid2], out / "bd.xlsx")
        self.assertTrue(all(f.exists() for f in files))

    def test_eje_objetivo_oe_y_partida_gg(self):
        ini = self.c.list_masters("INI", self.uid)[0]
        self.assertEqual((ini["eje"], ini["obj_poa"], ini["oe"]), ("Adaptación académica", "OBJ-01", "OE1"))
        self.assertEqual(ini["unidad_medida"], "Unidad")
        self.assertEqual(self.c.objs_por_eje("EJ-01", "codigo"), ["OBJ-01", "OBJ-10"])
        with self.assertRaises(ValueError):
            self.c.save_master("INI", self.uid, {"eje": "EJ-03", "nombre": "x", "obj_poa": "OBJ-01"})
        det = self.c.list_detalle("INI", ini["id"])[0]
        self.assertEqual(det["partida"], "Material de oficina")
        self.assertEqual(det["categoria"], "SUMINISTROS")
        self.assertIn("Cuenta NIIF: 6941010100", self.c.descripcion_partida("Material de oficina", "Egresos"))
        self.assertIn("Material de oficina", self.c.partidas_gg("Egresos"))
        self.assertNotIn("SUELDOS ADMINISTRATIVOS", self.c.partidas_gg("Egresos"))
        bd = [r for r in self.c.consolidated_rows([self.uid]) if r["Codigo_Actividad"] == "TST-INI-01"]
        self.assertEqual({r["Codigo_NIIF"] for r in bd}, {"6941010100"})
        self.assertEqual({r["Partida_GG"] for r in bd}, {"Material de oficina"})

    def test_plantilla_unidades_y_techos(self):
        import openpyxl
        path = export_plantilla_parametros(self.c, Path(self.tmp.name) / "plantilla.xlsx")
        wb = openpyxl.load_workbook(path)
        ws = wb["Unidades_Techos"]
        cab = {ws.cell(4, j).value: j for j in range(1, ws.max_column + 1)}
        fila = {ws.cell(r, cab["Código de área"]).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 2).value}
        tst = fila["TST"]
        ws.cell(tst, cab["Techo funcionamiento (S/)"], 77777)
        # fila con ID de OTRA instalación (apunta a la unidad OL de esta base) y código nuevo: debe crear otra unidad
        ol = next(u for u in self.c.unidades() if u["codigo_area"] == "OL")
        r = ws.max_row + 1
        for k, v in (("ID", ol["id"]), ("Código de área", "601102"), ("Macrogrupo", "MARKETING"),
                     ("Nombre de la unidad orgánica", "Lima Sur"), ("Centro de costo", "DIRECCION DE MARKETING"),
                     ("Techo funcionamiento (S/)", 5000)):
            ws.cell(r, cab[k], v)
        # cambio real de código: misma unidad (mismo centro de costo) con código nuevo
        ecc = next(u for u in self.c.unidades() if u["codigo_area"] == "ECC-PRE")
        ws.cell(fila["ECC-PRE"], cab["Código de área"], "4011")
        wb.save(path)
        rep = self.c.import_plantilla_parametros(path)
        self.assertIn("1 nuevas", rep["Unidades y techos"])
        self.assertTrue(any("TST - UNIDAD DE PRUEBA: Techo funcionamiento (S/)" in x and "77,777.00" in x
                            for x in rep["Cambios en unidades y techos"]))
        por_area = {u["codigo_area"]: u for u in self.c.unidades()}
        self.assertEqual(por_area["TST"]["techo_funcionamiento"], 77777)
        self.assertEqual(por_area["OL"]["id"], ol["id"])  # OL no fue sobrescrita
        self.assertEqual(por_area["601102"]["techo_funcionamiento"], 5000)
        self.assertEqual(por_area["4011"]["id"], ecc["id"])  # ECC-PRE cambió de código
        # sincronizar: eliminar las unidades que no están en el archivo y no tienen actividades
        ws.delete_rows(fila["OL"])
        wb.save(path)
        rep = self.c.import_plantilla_parametros(path, eliminar_unidades_ausentes=True)
        self.assertNotIn("OL", {u["codigo_area"] for u in self.c.unidades()})
        self.assertEqual(rep["Unidades eliminadas (no estaban en el archivo)"], 1)
        # sin la opción, una unidad ausente se conserva y se avisa
        ws.delete_rows(fila["TST"] - (1 if fila["OL"] < fila["TST"] else 0))
        wb.save(path)
        rep = self.c.import_plantilla_parametros(path, eliminar_unidades_ausentes=True)
        self.assertIn("TST", {u["codigo_area"] for u in self.c.unidades()})  # tiene actividades: se conserva
        self.assertIn("TST", rep["avisos"][0])

    def test_plantilla_parametros_ida_y_vuelta(self):
        import openpyxl
        path = export_plantilla_parametros(self.c, Path(self.tmp.name) / "plantilla.xlsx")
        wb = openpyxl.load_workbook(path)
        self.assertEqual(wb.sheetnames, ["Instrucciones", "Datos_Generales", "Unidades_Techos",
                                         "Cuentas_NIIF_Partidas_GG", "Catalogos_Estrategicos", "Costos_Referencia",
                                         "Plantilla_Funcionamiento", "Listas"])
        ws = wb["Catalogos_Estrategicos"]
        r = 5
        while ws.cell(r, 2).value:
            r += 1
        ws.cell(r, 2, "Eje estratégico"); ws.cell(r, 3, "EJ-10"); ws.cell(r, 4, "Sostenibilidad")  # noqa: E702
        for j, v in enumerate(["", "Objetivo anual del POA (OBJ)", "OBJ-16", "Reducir la huella de carbono",
                               "EJ-10", "OE3"], start=1):
            ws.cell(r + 1, j, v or None)
        ws = wb["Cuentas_NIIF_Partidas_GG"]
        r = 5
        while ws.cell(r, 2).value:
            r += 1
        for j, v in enumerate(["", "Egresos", "Servicio de prueba", "Descripción de prueba", "999999",
                               "SERVICIO DE PRUEBA", "OTROS SERVICIOS", "", "", "Sí"], start=1):
            ws.cell(r, j, v or None)
        ws = wb["Datos_Generales"]
        for rr in range(5, ws.max_row + 1):
            if ws.cell(rr, 1).value == "tipo_cambio":
                ws.cell(rr, 3, 3.8)
        edit = Path(self.tmp.name) / "plantilla_editada.xlsx"
        wb.save(edit)
        n_cuentas = self.db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas")
        rep = self.c.import_plantilla_parametros(edit)
        self.assertEqual(rep["avisos"], [])
        self.assertAlmostEqual(self.c.params.tc, 3.8)
        self.assertEqual(self.c.objs_por_eje("Sostenibilidad", "codigo"), ["OBJ-16"])
        self.assertEqual(self.c.oe_de_obj("OBJ-16"), "OE3")
        self.assertIn("Servicio de prueba", self.c.partidas_gg("Egresos"))
        self.assertEqual(self.db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas"), n_cuentas + 1)
        self.c.import_plantilla_parametros(edit)
        self.assertEqual(self.db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas"), n_cuentas + 1)

    def test_unidades_y_techos_al_importar(self):
        # «Nuevo» con un código de área existente actualiza la unidad (no IntegrityError)
        uid = self.c.save_unidad({"codigo_area": "tst", "nombre_unidad": "UNIDAD DE PRUEBA", "techo_iniciativas": 777})
        self.assertEqual(uid, self.uid)
        self.assertEqual(self.c.unidad(self.uid)["techo_iniciativas"], 777)
        otra = self.c.save_unidad({"codigo_area": "OTRA", "nombre_unidad": "OTRA UNIDAD"})
        with self.assertRaises(ValueError):  # mismo código y misma unidad orgánica (sede) que otra unidad
            self.c.save_unidad({"codigo_area": "TST", "nombre_unidad": "UNIDAD DE PRUEBA"}, otra)
        # el mismo código en otra sede es otra unidad
        norte = self.c.save_unidad({"codigo_area": "TST", "nombre_unidad": "Lima Norte", "techo_iniciativas": 5})
        self.assertNotEqual(norte, self.uid)
        self.assertEqual(self.c.unidad(self.uid)["techo_iniciativas"], 777)
        self.assertEqual(self.c.buscar_unidad("TST")[0], None)  # ambiguo: pide la sede
        self.assertEqual(self.c.buscar_unidad("TST - Lima Norte")[0], norte)
        self.assertEqual(self.c.clave_unidad(self.c.unidad(norte)), "TST - Lima Norte")
        self.c.delete_unidad(norte)
        # Los techos y sub-techos viajan en el libro exportado y se recuperan al importarlo
        self.db.update("unidades_organicas", self.uid, {"techo_planilla_docente": 1234, "techo_marketing": 700})
        libro = OfficialWorkbookExporter(self.c, self.uid).export(Path(self.tmp.name) / "libro.xlsx")
        self.db.update("unidades_organicas", self.uid, {"techo_iniciativas": 1, "techo_planilla_docente": 0,
                                                        "techo_marketing": 0})
        n = self.db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas")
        rep = self.c.import_parametros(libro)  # partidas repetidas con distinto código no deben chocar
        self.assertIn("TST", rep["techos_unidad"])
        u = self.c.unidad(self.uid)
        self.assertEqual((u["techo_iniciativas"], u["techo_planilla_docente"], u["techo_marketing"]),
                         (777, 1234, 700))
        self.assertEqual(self.db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas"), n)
        self.db.update("unidades_organicas", self.uid, {"techo_planilla_docente": 0})
        self.c.import_ficha(libro)
        self.assertEqual(self.c.unidad(self.uid)["techo_planilla_docente"], 1234)


if __name__ == "__main__":
    unittest.main()
