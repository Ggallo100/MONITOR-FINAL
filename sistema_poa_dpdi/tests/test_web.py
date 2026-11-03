"""
Pruebas de la versión web multiusuario (Flask test client).

    python -m unittest discover -s tests -v
"""
import io
import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import importlib.util

HAY_FLASK = importlib.util.find_spec("flask") is not None  # la versión de escritorio no requiere Flask

CLAVE_ADMIN = "Admin2027x"


@unittest.skipUnless(HAY_FLASK, "Flask no instalado")
class WebTest(unittest.TestCase):
    def setUp(self):
        import os
        from web.app import create_app
        self.tmp = tempfile.TemporaryDirectory()
        os.environ["POA_ADMIN_PASSWORD"] = CLAVE_ADMIN
        self.app = create_app(data_dir=self.tmp.name, testing=True)
        os.environ.pop("POA_ADMIN_PASSWORD")
        self.admin = self.cliente("admin", CLAVE_ADMIN, nueva="Dpdi2027ok")

    def tearDown(self):
        self.tmp.cleanup()

    # ----------------------------------------------------------- utilidades
    def cliente(self, usuario, clave, nueva=None):
        c = self.app.test_client()
        tok = self.token(c, "/login")
        r = c.post("/login", data={"usuario": usuario, "clave": clave, "csrf_token": tok})
        self.assertEqual(r.status_code, 302, r.get_data(as_text=True)[:300])
        if nueva:
            tok = self.token(c, "/cambiar-clave")
            r = c.post("/cambiar-clave", data={"actual": clave, "nueva": nueva, "repetir": nueva, "csrf_token": tok})
            self.assertEqual(r.headers.get("Location"), "/")
        c.csrf = self.token(c, "/")
        return c

    @staticmethod
    def token(c, url):
        html = c.get(url).get_data(as_text=True)
        m = re.search(r'name="csrf-token" content="([^"]+)"', html) or re.search(r'name="csrf_token" value="([^"]+)"', html)
        return m.group(1)

    def api(self, c, metodo, url, json=None, codigo=200, **kw):
        r = c.open(url, method=metodo, json=json, headers={"X-CSRF-Token": c.csrf}, **kw)
        self.assertEqual(r.status_code, codigo, f"{metodo} {url}: {r.get_data(as_text=True)[:400]}")
        return r.get_json() if r.is_json else r

    def crear_responsable(self, usuario, unidad_id):
        self.api(self.admin, "POST", "/api/admin/t/usuarios",
                 {"usuario": usuario, "nombre": "Responsable " + usuario, "rol": "RESPONSABLE", "activo": 1,
                  "debe_cambiar_clave": 0, "nueva_clave": "Clave2027a", "unidades": [str(unidad_id)]})
        return self.cliente(usuario, "Clave2027a")

    # --------------------------------------------------------------- pruebas
    def test_seguridad_basica(self):
        anon = self.app.test_client()
        self.assertEqual(anon.get("/").status_code, 302)
        self.assertEqual(anon.get("/api/listas").status_code, 401)
        # CSRF obligatorio en escrituras
        r = self.admin.post("/api/contexto", json={"uid": 1})
        self.assertEqual(r.status_code, 400)
        # bloqueo tras 5 intentos fallidos
        c = self.app.test_client()
        for _ in range(5):
            tok = self.token(c, "/login")
            c.post("/login", data={"usuario": "admin", "clave": "mala", "csrf_token": tok})
        tok = self.token(c, "/login")
        r = c.post("/login", data={"usuario": "admin", "clave": "Dpdi2027ok", "csrf_token": tok})
        self.assertEqual(r.status_code, 429)

    def test_permisos_por_unidad_y_flujo_de_revision(self):
        unidades = self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
        ol = next(u for u in unidades if u["codigo_area"] == "OL")["id"]
        ecc = next(u for u in unidades if u["codigo_area"] == "ECC-PRE")["id"]
        resp = self.crear_responsable("ana.alva", ol)
        # solo ve su unidad
        self.api(resp, "POST", "/api/contexto", {"uid": ecc}, codigo=403)
        self.api(resp, "GET", f"/api/md/INI?u={ecc}", codigo=403)
        self.api(resp, "GET", "/api/admin/t/usuarios", codigo=403)
        self.assertEqual(resp.get("/admin/parametros").status_code, 403)
        # formula su unidad: maestro + detalle con cálculo
        ini = self.api(resp, "POST", "/api/md/INI", {"eje": "EJ-01", "obj_poa": "OBJ-01", "nombre": "Aulas virtuales",
                                                     "indicador": "Cursos", "unidad_medida": "Unidad",
                                                     "meta_anual": "10", "t1": "5", "t2": "5", "responsable": "Ana",
                                                     "prioridad": "2"})
        self.assertEqual(ini["row"]["oe"], "OE1")
        det = {"tipo": "Egresos", "partida": "Material de oficina", "cantidad": "10", "moneda": "US$",
               "costo_unit_moneda": "12.5", "sustento": "Cotización",
               "meses": ["", "", "145.83", "145.83", "145.84"] + [""] * 7}
        calc = self.api(resp, "POST", "/api/calc/detalle", det)
        self.assertEqual((calc["_total"], calc["_control"], calc["_categoria"]), ("437.50", "OK", "SUMINISTROS"))
        self.api(resp, "POST", f"/api/md/INI/{ini['id']}/detalle", det)
        lista = self.api(resp, "GET", "/api/md/INI")
        self.assertAlmostEqual(lista["rows"][0]["_egresos"], 437.5)
        # el administrador no puede ser editado por el responsable ni otras unidades
        otra = self.api(self.admin, "POST", f"/api/md/OO?u={ecc}", {"nombre": "Licencia", "categoria": "Licencias de software"})
        self.api(resp, "DELETE", f"/api/md/OO/{otra['id']}", codigo=403)
        # flujo: enviar → solo lectura → observar → editable → enviar → aprobar
        self.api(resp, "POST", "/api/estado", {"accion": "enviar"})
        self.api(resp, "PUT", f"/api/md/INI/{ini['id']}", {"eje": "EJ-01", "nombre": "Cambio"}, codigo=403)
        self.api(resp, "POST", "/api/estado", {"accion": "aprobar"}, codigo=403)
        self.api(self.admin, "POST", "/api/contexto", {"uid": ol})
        self.api(self.admin, "POST", "/api/estado", {"accion": "observar", "observacion": "Falta sustento"})
        self.api(resp, "PUT", f"/api/md/INI/{ini['id']}", {"eje": "EJ-01", "obj_poa": "OBJ-01", "nombre": "Cambio"})
        self.api(resp, "POST", "/api/estado", {"accion": "enviar"})
        self.api(self.admin, "POST", "/api/estado", {"accion": "aprobar"})
        tablero = self.api(self.admin, "GET", "/api/admin/tablero")["rows"]
        self.assertEqual(next(t for t in tablero if t["id"] == ol)["estado"], "Aprobada")
        bit = self.api(self.admin, "GET", "/api/admin/bitacora?q=ana.alva")["rows"]
        self.assertTrue(any(b["accion"] == "Ficha enviada" for b in bit))

    def test_consulta_solo_lectura_y_descargas(self):
        self.api(self.admin, "POST", "/api/admin/t/usuarios",
                 {"usuario": "consulta", "rol": "CONSULTA", "activo": 1, "debe_cambiar_clave": 0,
                  "nueva_clave": "Consulta2027"})
        con = self.cliente("consulta", "Consulta2027")
        self.assertGreaterEqual(len(self.api(con, "GET", "/api/admin/tablero")["rows"]), 2)
        self.api(con, "POST", "/api/md/INI", {"eje": "EJ-01", "nombre": "x"}, codigo=403)
        for url, firma in (("/descargas/libro", b"PK"), ("/descargas/pdf", b"%PDF"), ("/descargas/bd?alcance=todas", b"PK")):
            r = con.get(url)
            self.assertEqual(r.status_code, 200, url)
            self.assertTrue(r.data.startswith(firma), url)
        self.assertEqual(con.get("/descargas/plantilla-parametros").status_code, 403)
        self.assertEqual(con.get("/descargas/respaldo").status_code, 403)
        r = self.admin.get("/descargas/respaldo")
        self.assertTrue(r.data.startswith(b"SQLite format 3"))

    def test_paginas_e_importacion(self):
        for url in ("/", "/iniciativas", "/otras-operativas", "/actividades-propias", "/inversion", "/funcionamiento",
                    "/resumen", "/admin/tablero", "/admin/parametros", "/admin/unidades", "/admin/cuentas",
                    "/admin/catalogos", "/admin/costos", "/admin/plantilla", "/admin/usuarios", "/admin/bitacora"):
            self.assertEqual(self.admin.get(url).status_code, 200, url)
        for k in ("INI", "OO", "AP", "INV"):
            self.api(self.admin, "GET", f"/api/md/{k}")
        for k in ("FI", "FE"):
            self.api(self.admin, "GET", f"/api/fun/{k}")
        self.api(self.admin, "GET", "/api/resumen")
        self.api(self.admin, "GET", "/api/inicio")
        # exportar el libro y volver a importarlo en la unidad actual
        fe = {"partida": "AGUA - DIRECTO", "cantidad": "1", "costo_unit_moneda": "2400", "aplica_reajuste": "No",
              "sustento": "Histórico de ejecución", "meses": ["0.08333333"] * 12}
        self.api(self.admin, "POST", "/api/fun/FE", fe)
        libro = self.admin.get("/descargas/libro").data
        rep = self.api(self.admin, "POST", "/cargas/ficha",
                       data={"archivo": (io.BytesIO(libro), "ficha.xlsx"), "destino": "actual"},
                       content_type="multipart/form-data")
        self.assertEqual(rep["conteos"]["FE"], 1)
        # plantilla de parámetros: descarga e importación
        plantilla = self.admin.get("/descargas/plantilla-parametros").data
        rep = self.api(self.admin, "POST", "/cargas/plantilla-parametros",
                       data={"archivo": (io.BytesIO(plantilla), "plantilla.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(rep["avisos"], [])

    def test_ayuda_y_manuales_por_rol(self):
        from web.app import BASE, MANUALES
        html = self.admin.get("/ayuda").get_data(as_text=True)
        for m in MANUALES:
            self.assertIn(m["titulo"], html)
            if (BASE / "manuales" / "pdf" / m["archivo"]).exists():
                r = self.admin.get(f"/ayuda/{m['archivo']}")
                self.assertTrue(r.data.startswith(b"%PDF"), m["archivo"])
                r.close()
        self.assertEqual(self.admin.get("/ayuda/../web/app.py").status_code, 404)
        unidades = self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
        resp = self.crear_responsable("luis.rios", unidades[0]["id"])
        html = resp.get("/ayuda").get_data(as_text=True)
        self.assertIn("Manual del usuario", html)
        self.assertNotIn("Manual técnico", html)
        self.assertEqual(resp.get("/ayuda/Manual_Tecnologia_POA_DPDI.pdf").status_code, 404)

    def test_restablecer_clave_desbloquea(self):
        unidades = self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
        self.crear_responsable("rosa.paz", unidades[0]["id"])
        c = self.app.test_client()
        for _ in range(5):
            c.post("/login", data={"usuario": "rosa.paz", "clave": "mala", "csrf_token": self.token(c, "/login")})
        usuarios = self.api(self.admin, "GET", "/api/admin/t/usuarios")["rows"]
        rosa = next(u for u in usuarios if u["usuario"] == "rosa.paz")
        self.api(self.admin, "PUT", f"/api/admin/t/usuarios/{rosa['id']}",
                 {"usuario": "rosa.paz", "rol": "RESPONSABLE", "activo": 1, "debe_cambiar_clave": 0,
                  "nueva_clave": "Nueva2027b", "unidades": [str(unidades[0]["id"])]})
        self.cliente("rosa.paz", "Nueva2027b")

    def test_usuario_con_clave_temporal_y_nombre_normalizado(self):
        unidades = self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
        r = self.api(self.admin, "POST", "/api/admin/t/usuarios",
                     {"usuario": "Pedro Ruíz", "nombre": "Pedro Ruiz", "rol": "RESPONSABLE", "activo": 1,
                      "debe_cambiar_clave": 0, "nueva_clave": "", "unidades": [str(unidades[0]["id"])]})
        self.assertEqual(r["usuario"], "pedro.ruiz")
        self.assertNotIn("aviso", r)
        c = self.app.test_client()
        c.post("/login", data={"usuario": "pedro.ruiz", "clave": r["clave_temporal"], "csrf_token": self.token(c, "/login")})
        self.assertIn("cambiar-clave", c.get("/").headers.get("Location", ""))  # debe cambiarla
        r = self.api(self.admin, "POST", "/api/admin/t/usuarios", {"usuario": "sin.unidad", "rol": "RESPONSABLE",
                                                                    "activo": 1, "unidades": []})
        self.assertIn("aviso", r)

    def test_carga_masiva_de_usuarios(self):
        import openpyxl
        plantilla = self.admin.get("/descargas/plantilla-usuarios").data
        wb = openpyxl.load_workbook(io.BytesIO(plantilla))
        self.assertEqual(wb.sheetnames, ["Instrucciones", "Usuarios", "Sugeridos", "Unidades"])
        ws = wb["Usuarios"]
        self.assertEqual(ws.cell(2, 1).value, "admin")
        codigos = [r[0] for r in wb["Unidades"].iter_rows(min_row=2, values_only=True)]
        ws.append(["carla.diaz", "Carla Díaz", "carla@uni.pe", "RESPONSABLE", f"{codigos[0]}, {codigos[1]}", "", "Sí"])
        ws.append(["Luis Soto", "Luis Soto", "", "CONSULTA", "", "Consulta2027", "Sí"])
        ws.append(["mal.rol", "X", "", "JEFE", "", "", "Sí"])
        ws.append(["eva.paz", "Eva", "", "RESPONSABLE", "NOEXISTE", "", "Sí"])
        buf = io.BytesIO()
        wb.save(buf)
        rep = self.api(self.admin, "POST", "/cargas/usuarios", data={"archivo": (io.BytesIO(buf.getvalue()), "u.xlsx")},
                       content_type="multipart/form-data")
        self.assertEqual((rep["Usuarios creados"], rep["Usuarios actualizados"]), (3, 1))
        self.assertEqual({c["usuario"] for c in rep["credenciales"]}, {"carla.diaz", "luis.soto", "eva.paz"})
        self.assertTrue(any("JEFE" in a for a in rep["avisos"]))
        self.assertTrue(any("NOEXISTE" in a for a in rep["avisos"]))
        usuarios = {u["usuario"]: u for u in self.api(self.admin, "GET", "/api/admin/t/usuarios")["rows"]}
        self.assertEqual(len(usuarios["carla.diaz"]["unidades"]), 2)
        clave = next(c["clave"] for c in rep["credenciales"] if c["usuario"] == "carla.diaz")
        c = self.app.test_client()
        r = c.post("/login", data={"usuario": "carla.diaz", "clave": clave, "csrf_token": self.token(c, "/login")})
        self.assertEqual(r.status_code, 302)
        self.cliente("admin", "Dpdi2027ok")  # el administrador conserva su clave (fila sin clave)

    def test_importar_parametros_reconoce_el_archivo(self):
        plantilla = self.admin.get("/descargas/plantilla-parametros").data
        libro = self.admin.get("/descargas/libro").data
        for url in ("/cargas/plantilla-parametros", "/cargas/parametros-ficha"):  # cualquiera de los dos botones
            rep = self.api(self.admin, "POST", url, data={"archivo": (io.BytesIO(plantilla), "plantilla.xlsx")},
                           content_type="multipart/form-data")
            self.assertEqual(rep["Archivo reconocido"], "Plantilla de Parámetros y Catálogos")
            self.assertIn("filas leídas", rep["Unidades y techos"])
            rep = self.api(self.admin, "POST", url, data={"archivo": (io.BytesIO(libro), "ficha.xlsx")},
                           content_type="multipart/form-data")
            self.assertEqual(rep["Archivo reconocido"], "Parámetros de una Ficha Excel")
        import openpyxl
        otro = io.BytesIO()
        openpyxl.Workbook().save(otro)
        rep = self.api(self.admin, "POST", "/cargas/plantilla-parametros", codigo=400,
                       data={"archivo": (io.BytesIO(otro.getvalue()), "otro.xlsx")}, content_type="multipart/form-data")
        self.assertIn("no es la plantilla", str(rep))

    def test_migra_colores_de_la_version_anterior(self):
        from database import Database
        from web.app import create_app
        bd = Database(self.app.config["DB_PATH"], seed=False, init=False)
        for k, v in {"tema_primario": "#1F3864", "tema_secundario": "#2E5597", "tema_acento": "#123456"}.items():
            bd.set_meta(k, v)
        bd.set_meta("tema_version", "")
        bd.close()
        app = create_app(data_dir=self.tmp.name, testing=True)  # reinicio del servidor tras actualizar
        html = app.test_client().get("/login").get_data(as_text=True)
        self.assertIn("--primario: #FF6600", html)
        self.assertIn("--secundario: #333333", html)
        self.assertIn("#123456", html)  # un color personalizado por la DPDI se respeta
        self.assertRegex(html, r"css/app\.css\?v=\d+")

    def test_tipo_actividad_comercial_marketing(self):
        from catalogs_seed import ACTIVIDADES_COMERCIALES
        # catálogo inicial y mantenedor (Catálogos filtrado por ACT_COM)
        cat = self.api(self.admin, "GET", "/api/admin/t/catalogos?tipo=ACT_COM")["rows"]
        self.assertEqual([x["codigo"] for x in cat], ACTIVIDADES_COMERCIALES)
        self.assertEqual(len(self.api(self.admin, "GET", "/api/listas")["act_comercial"]), 18)
        self.assertEqual(self.admin.get("/admin/act_comercial").status_code, 200)
        nuevo = self.api(self.admin, "POST", "/api/admin/t/act_comercial", {"codigo": "ventas  lima norte"})["id"]
        self.assertIn("VENTAS LIMA NORTE", [x["value"] for x in self.api(self.admin, "GET", "/api/listas")["act_comercial"]])
        self.api(self.admin, "POST", "/api/admin/t/act_comercial", {"codigo": "ATL"}, codigo=400)  # duplicado
        self.api(self.admin, "DELETE", f"/api/admin/t/act_comercial/{nuevo}")
        mkt = self.api(self.admin, "POST", "/api/admin/t/unidades",
                       {"codigo_area": "601102", "nombre_unidad": "Lima Sur", "centro_costo": "DIRECCION DE MARKETING",
                        "macrogrupo": "MARKETING"})["id"]
        otra = next(u for u in self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
                    if u["codigo_area"] == "OL")["id"]
        # el combo solo aparece en Otras operativas de unidades de Marketing
        self.api(self.admin, "POST", "/api/contexto", {"uid": mkt})
        self.assertIn("Tipo de Actividad Comercial", self.admin.get("/otras-operativas").get_data(as_text=True))
        self.assertNotIn("Tipo de Actividad Comercial", self.admin.get("/iniciativas").get_data(as_text=True))
        base = {"categoria": "Licencias de software", "nombre": "Campaña admisión 2027"}
        self.api(self.admin, "POST", "/api/md/OO", base, codigo=400)  # obligatorio en Marketing
        self.api(self.admin, "POST", "/api/md/OO", dict(base, actividad_comercial="RADIO"), codigo=400)  # no existe
        r = self.api(self.admin, "POST", "/api/md/OO", dict(base, actividad_comercial="campañas"))
        fila = next(x for x in self.api(self.admin, "GET", "/api/md/OO")["rows"] if x["id"] == r["id"])
        self.assertEqual(fila["actividad_comercial"], "CAMPAÑAS")
        self.api(self.admin, "POST", f"/api/md/OO/{r['id']}/detalle",
                 {"tipo": "Egresos", "partida": "Material de oficina", "cantidad": "1", "costo_unit_moneda": "1000",
                  "sustento": "Cotización", "meses": ["1000"] + [""] * 11})
        sec = next(x for x in self.api(self.admin, "GET", "/api/resumen")["secciones"] if x["key"] == "s6m")
        self.assertEqual(next(f for f in sec["rows"] if f[0] == "CAMPAÑAS")[1:4], [1, 0.0, 1000.0])
        # renombrar la opción en el mantenedor actualiza los registros
        camp = next(x for x in cat if x["codigo"] == "CAMPAÑAS")
        self.api(self.admin, "DELETE", f"/api/admin/t/act_comercial/{camp['id']}", codigo=400)  # en uso
        self.api(self.admin, "PUT", f"/api/admin/t/act_comercial/{camp['id']}", dict(camp, codigo="Campañas digitales"))
        fila = next(x for x in self.api(self.admin, "GET", "/api/md/OO")["rows"] if x["id"] == r["id"])
        self.assertEqual(fila["actividad_comercial"], "CAMPAÑAS DIGITALES")
        # Excel: columna en Otras_Operat y en la BD consolidada; ida y vuelta conserva el valor
        import openpyxl
        libro = self.admin.get("/descargas/libro").data
        wb = openpyxl.load_workbook(io.BytesIO(libro))
        self.assertEqual(wb["Otras_Operat"]["K6"].value, "Tipo de actividad comercial (solo Marketing)")
        bd = wb["BD_CONSOLIDADA_NIIF"]
        col_ac = [c.value for c in bd[1]].index("Tipo_Actividad_Comercial") + 1
        self.assertEqual(bd.cell(2, col_ac).value, "CAMPAÑAS DIGITALES")
        rep = self.api(self.admin, "POST", "/cargas/ficha", data={"archivo": (io.BytesIO(libro), "f.xlsx"),
                                                                  "destino": "actual"},
                       content_type="multipart/form-data")
        self.assertEqual(rep["conteos"]["OO"], 1)
        self.assertEqual(self.api(self.admin, "GET", "/api/md/OO")["rows"][0]["actividad_comercial"],
                         "CAMPAÑAS DIGITALES")
        # en una unidad que no es de Marketing no se muestra ni se guarda
        self.api(self.admin, "POST", "/api/contexto", {"uid": otra})
        self.assertNotIn("Tipo de Actividad Comercial", self.admin.get("/otras-operativas").get_data(as_text=True))
        r2 = self.api(self.admin, "POST", "/api/md/OO", dict(base, actividad_comercial="ATL"))
        fila = next(x for x in self.api(self.admin, "GET", "/api/md/OO")["rows"] if x["id"] == r2["id"])
        self.assertEqual(fila["actividad_comercial"], "")
        self.assertFalse(any(x["key"] == "s6m" for x in self.api(self.admin, "GET", "/api/resumen")["secciones"]))

    def test_libro_consolidado_todas_las_unidades(self):
        import openpyxl
        unidades = self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]
        ol = next(u for u in unidades if u["codigo_area"] == "OL")["id"]
        ecc = next(u for u in unidades if u["codigo_area"] == "ECC-PRE")["id"]
        for uid, nombre in ((ol, "Feria laboral"), (ecc, "Licencia de edición")):
            m = self.api(self.admin, "POST", f"/api/md/OO?u={uid}", {"categoria": "Licencias de software", "nombre": nombre})
            self.api(self.admin, "POST", f"/api/md/OO/{m['id']}/detalle",
                     {"tipo": "Egresos", "partida": "Material de oficina", "cantidad": "1", "costo_unit_moneda": "600",
                      "sustento": "Cotización", "meses": ["300", "300"] + [""] * 10})
        # solo el administrador
        self.api(self.admin, "POST", "/api/admin/t/usuarios",
                 {"usuario": "consulta", "rol": "CONSULTA", "activo": 1, "debe_cambiar_clave": 0, "nueva_clave": "Consulta2027"})
        self.assertEqual(self.cliente("consulta", "Consulta2027").get("/descargas/consolidado").status_code, 403)
        # todas las unidades: maestros, detalles y BD con el código de área de cada fila
        wb = openpyxl.load_workbook(io.BytesIO(self.admin.get("/descargas/consolidado?estado=todas").data))
        for hoja in ("Instrucciones", "Resumen_Unidades", "Iniciativas", "Detalle_Iniciativas", "Otras_Operativas",
                     "Detalle_Otras_Operativas", "Actividades_Propias", "Inversiones", "Funcionamiento_Ing",
                     "Funcionamiento_Egr", "BD_CONSOLIDADA_NIIF"):
            self.assertIn(hoja, wb.sheetnames)
        oo = wb["Otras_Operativas"]
        cab = [c.value for c in oo[1]]
        nombres = {(r[cab.index("Codigo_Area")], r[cab.index("nombre")]) for r in oo.iter_rows(min_row=2, values_only=True)}
        self.assertTrue({("OL", "Feria laboral"), ("ECC-PRE", "Licencia de edición")} <= nombres)
        bd = wb["BD_CONSOLIDADA_NIIF"]
        cab = [c.value for c in bd[1]]
        areas = {r[cab.index("Codigo_Area")] for r in bd.iter_rows(min_row=2, values_only=True)}
        self.assertTrue({"OL", "ECC-PRE"} <= areas)
        self.assertEqual(cab[-1], "Estado_Ficha")
        self.assertEqual(len(wb["Resumen_Unidades"]["A"]) - 1, len(unidades))
        # solo aprobadas
        self.api(self.admin, "GET", "/descargas/consolidado?estado=aprobadas", codigo=400)
        self.api(self.admin, "POST", "/api/estado?u=" + str(ol), {"accion": "enviar"})
        self.api(self.admin, "POST", "/api/estado?u=" + str(ol), {"accion": "aprobar"})
        wb = openpyxl.load_workbook(io.BytesIO(self.admin.get("/descargas/consolidado?estado=aprobadas").data))
        bd = wb["BD_CONSOLIDADA_NIIF"]
        cab = [c.value for c in bd[1]]
        self.assertEqual({r[cab.index("Codigo_Area")] for r in bd.iter_rows(min_row=2, values_only=True)}, {"OL"})
        # CSV para carga masiva
        csv = self.admin.get("/descargas/consolidado?estado=todas&formato=csv").data.decode("utf-8-sig")
        self.assertTrue(csv.startswith("Anio;Codigo_Area;"))
        self.assertIn(";ECC-PRE;", csv)
        # el tablero informa la observación vigente para la ventana de revisión
        self.api(self.admin, "POST", "/api/estado?u=" + str(ol), {"accion": "observar", "observacion": "Falta sustento"})
        fila = next(x for x in self.api(self.admin, "GET", "/api/admin/tablero")["rows"] if x["id"] == ol)
        self.assertEqual((fila["estado"], fila["estado_obs"]), ("Observada", "Falta sustento"))

    def test_mismo_codigo_en_dos_sedes(self):
        import openpyxl
        sur = self.api(self.admin, "POST", "/api/admin/t/unidades",
                       {"codigo_area": "10110101", "nombre_unidad": "Lima Sur", "centro_costo": "EP INGENIERÍA DE SISTEMAS",
                        "macrogrupo": "PREGRADO", "techo_funcionamiento": "1000"})["id"]
        norte = self.api(self.admin, "POST", "/api/admin/t/unidades",
                         {"codigo_area": "10110101", "nombre_unidad": "Lima Norte", "centro_costo": "EP INGENIERÍA DE SISTEMAS",
                          "macrogrupo": "PREGRADO", "techo_funcionamiento": "2000"})["id"]
        self.assertNotEqual(sur, norte)
        filas = {u["id"]: u for u in self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]}
        self.assertEqual((filas[sur]["techo_funcionamiento"], filas[norte]["techo_funcionamiento"]), (1000, 2000))
        # plantilla de parámetros: cada sede conserva sus techos
        plantilla = self.admin.get("/descargas/plantilla-parametros").data
        wb = openpyxl.load_workbook(io.BytesIO(plantilla))
        ws = wb["Unidades_Techos"]
        cab = {ws.cell(4, j).value: j for j in range(1, ws.max_column + 1)}
        for r in range(5, ws.max_row + 1):
            if ws.cell(r, cab["Código de área"]).value == "10110101":
                nuevo = 3000 if ws.cell(r, cab["Nombre de la unidad orgánica"]).value == "Lima Sur" else 4000
                ws.cell(r, cab["Techo funcionamiento (S/)"], nuevo)
        buf = io.BytesIO()
        wb.save(buf)
        rep = self.api(self.admin, "POST", "/cargas/plantilla-parametros",
                       data={"archivo": (io.BytesIO(buf.getvalue()), "p.xlsx")}, content_type="multipart/form-data")
        self.assertIn("2 actualizadas", rep["Unidades y techos"])
        filas = {u["id"]: u for u in self.api(self.admin, "GET", "/api/admin/t/unidades")["rows"]}
        self.assertEqual((filas[sur]["techo_funcionamiento"], filas[norte]["techo_funcionamiento"]), (3000, 4000))
        # carga masiva de usuarios: la clave «código - sede» resuelve la unidad; el código solo es ambiguo
        wb = openpyxl.load_workbook(io.BytesIO(self.admin.get("/descargas/plantilla-usuarios").data))
        claves = [r[0] for r in wb["Unidades"].iter_rows(min_row=2, values_only=True)]
        self.assertIn("10110101 - Lima Norte", claves)
        wb["Usuarios"].append(["juan.norte", "Juan", "", "RESPONSABLE", "10110101 - Lima Norte", "", "Sí"])
        wb["Usuarios"].append(["juan.duda", "Juan", "", "RESPONSABLE", "10110101", "", "Sí"])
        buf = io.BytesIO()
        wb.save(buf)
        rep = self.api(self.admin, "POST", "/cargas/usuarios", data={"archivo": (io.BytesIO(buf.getvalue()), "u.xlsx")},
                       content_type="multipart/form-data")
        usuarios = {u["usuario"]: u for u in self.api(self.admin, "GET", "/api/admin/t/usuarios")["rows"]}
        self.assertEqual(usuarios["juan.norte"]["unidades"], [norte])
        self.assertEqual(usuarios["juan.norte"]["_unidades"], "10110101 - Lima Norte")
        self.assertTrue(any("indique la sede" in a for a in rep["avisos"]))

    def test_identidad_grafica(self):
        self.api(self.admin, "PUT", "/api/admin/parametros", {"tema_primario": "#123456", "tema_acento": "#ABCDEF"})
        self.assertIn("--primario: #123456", self.admin.get("/").get_data(as_text=True))
        self.api(self.admin, "PUT", "/api/admin/parametros", {"tema_primario": "rojo"}, codigo=400)
        png = (b"\x89PNG\r\n\x1a\n" + b"\x00" * 32)
        self.api(self.admin, "POST", "/cargas/logo", data={"archivo": (io.BytesIO(png), "logo.png")},
                 content_type="multipart/form-data")
        self.assertEqual(self.app.test_client().get("/marca/logo").status_code, 200)


if __name__ == "__main__":
    unittest.main()
