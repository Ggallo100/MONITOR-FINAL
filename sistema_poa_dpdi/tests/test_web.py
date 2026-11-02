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
