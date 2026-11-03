"""
catalogs_seed.py
================
Catálogos oficiales del formato Ficha_Formulacion_POA.xlsx (hoja «Parámetros»,
secciones E a S, y hoja «Hoja6») y del índice de códigos de cuentas contables
NIIF (1_index_codigos_cc.xlsx), embebidos para que el ejecutable funcione sin
archivos externos. La función seed_database() es idempotente: solo inserta lo
que falta, por lo que los cambios hechos desde el Mantenedor se conservan.
"""
from __future__ import annotations

MACROGRUPOS = [
    "Soporte Académico/Administrativo",
    "Pregrado",
    "Posgrado",
    "Educación Continua",
]

# Sección D — Clasificación de actividades (referencia)
TIPOS_ACTIVIDAD = [
    ("Iniciativas con impacto estratégico",
     "Iniciativas cuya realización impacta sobre el cumplimiento de los objetivos anuales establecidos en el POA.",
     "Iniciativas + Detalle_Iniciativas"),
    ("Actividades de funcionamiento",
     "Actividades regularmente ejecutadas para la prestación del servicio académico o administrativo. Egresos: "
     "remuneraciones, pago de servicios básicos y regulares, depreciaciones y tributos. Ingresos: matrículas, "
     "pensiones y tasas educativas, entre otros de naturaleza regular.",
     "Funcionamiento_Ing + Funcionamiento_Egr"),
    ("Otras actividades operativas",
     "Actividades regularmente ejecutadas para la prestación del servicio académico o administrativo que no "
     "comprenden los ingresos y egresos de funcionamiento y que no tienen objetivos estratégicos ni operativos "
     "asociados (licencias de software, renovación de suscripciones, marketing, mantenimientos, entre otras).",
     "Otras_Operat + Detalle_Otras_Operativas"),
    ("Actividades propias",
     "Actividades de corta duración que generan ingresos a la universidad (congresos, seminarios, talleres).",
     "Act_Prop + Detalle_Act_Prop"),
    ("Actividades de inversión",
     "Actividades vinculadas a la adquisición e implementación de infraestructura física y tecnológica.",
     "Inversion + Detalle_Inversiones"),
]

# Sección H — Ejes de las iniciativas con impacto estratégico
EJES = [
    ("EJ-01", "Adaptación académica", "Actualización curricular, nuevas modalidades, innovación pedagógica y recursos de aprendizaje."),
    ("EJ-02", "Gestión docente", "Selección, capacitación, evaluación, carrera y reconocimiento docente."),
    ("EJ-03", "Investigación", "Proyectos, publicaciones, semilleros, laboratorios y fondos concursables."),
    ("EJ-04", "Responsabilidad Social", "Proyectos de responsabilidad social universitaria, extensión y voluntariado."),
    ("EJ-05", "Internacionalización", "Movilidad estudiantil y docente, convenios, doble titulación y redes internacionales."),
    ("EJ-06", "Transformación Digital", "Plataformas, sistemas de información, automatización de procesos y analítica de datos."),
    ("EJ-07", "Empleabilidad", "Prácticas preprofesionales, bolsa de trabajo, seguimiento de egresados y vínculo con empleadores."),
    ("EJ-08", "Aseguramiento de la Calidad", "Licenciamiento, acreditación, sistema de gestión de la calidad y auditorías."),
    ("EJ-09", "Crecimiento Institucional", "Nuevos programas y sedes, captación de estudiantes, posicionamiento."),
]

# Sección E — Objetivos estratégicos del PEI (la DPDI completa la descripción e indicador)
OBJETIVOS_PEI = [(f"OE{i}", f"Objetivo estratégico {i} del PEI", "") for i in range(1, 9)]

# Sección G — Objetivos anuales del POA: cada objetivo pertenece a un Eje estratégico y contribuye a un OE del PEI.
# Relación de dependencia en la formulación: Eje → Objetivo anual del POA → OE del PEI (automático).
# (codigo, descripcion, oe_pei, eje). Vínculos provisionales: la DPDI los actualiza con la plantilla de parámetros.
OBJETIVOS_POA = [(f"OBJ-{i:02d}", f"Objetivo anual {i:02d} del POA", f"OE{min(8, (i - 1) // 2 + 1)}",
                  f"EJ-{(i - 1) % 9 + 1:02d}") for i in range(1, 16)]

# Unidades de medida del indicador de las iniciativas
UNIDADES_MEDIDA = ["Unidad", "Porcentaje (%)"]

# Sección O — Niveles de prioridad
PRIORIDADES = [
    ("1", "Obligatorio", "Exigencias legales y normativas (licenciamiento, acreditación, seguridad) y continuidad operativa."),
    ("2", "Estratégico", "Contribuye directamente a un objetivo anual del POA o a un OE del PEI con meta en riesgo."),
    ("3", "Mejora", "Atiende planes de mejora, riesgos u oportunidades registrados o necesidades priorizadas."),
    ("4", "Deseable", "Aporta valor, pero puede postergarse sin afectar objetivos ni obligaciones."),
]

# Sección P — Origen de las iniciativas
ORIGENES = [
    "Diagnóstico estratégico (FODA)",
    "Riesgo u oportunidad (GC-MT-01)",
    "Plan de mejora / auditoría",
    "Licenciamiento o acreditación",
    "Necesidad de partes interesadas",
    "Lineamiento de la Alta Dirección",
    "Otro",
]

# Secciones Q y R — Tipos de sustento
SUSTENTOS_EGRESOS = ["Cotización", "Costo de referencia", "Contrato vigente", "Histórico de ejecución", "Otro"]
SUSTENTOS_INGRESOS = ["Proyección de matrícula aprobada", "Tarifario institucional vigente", "Histórico de recaudación",
                      "Convenio o contrato", "Otro"]

# Sección K — Categorías de otras actividades operativas
CATEGORIAS_OO = [
    ("OO-01", "Licencias de software", "Licencias anuales o por usuario."),
    ("OO-02", "Renovación de suscripciones", "Bases de datos, plataformas, membresías y servicios en la nube."),
    ("OO-03", "Marketing y publicidad", "Campañas, medios, material promocional y ferias."),
    ("OO-04", "Mantenimiento de infraestructura", "Mantenimiento preventivo y correctivo de locales."),
    ("OO-05", "Mantenimiento de equipos", "Equipos de cómputo, laboratorio y mobiliario."),
    ("OO-06", "Materiales y suministros", "Útiles de oficina, materiales de enseñanza y limpieza."),
    ("OO-07", "Eventos y actividades institucionales", "Ceremonias, actividades de integración y eventos sin cobro."),
    ("OO-08", "Otras actividades operativas", "Actividades operativas no incluidas en las categorías anteriores."),
    ("OO-09", "Programas especializados", "Ingresos por PEAS"),
    ("OO-10", "Programas extracurriculares", "Ingresos por Cursos de Idiomas, Cursos Extracurriculares"),
    ("OO-11", "Otroas tasas educación continua", "Ingresos por libros, constancias, examenes suficiencia"),
]

# Sección L — Tipos de actividades propias
TIPOS_AP = [("Congreso", ""), ("Seminario", ""), ("Taller", ""), ("Curso corto", ""), ("Otro", "")]

# Sección M — Tipos de inversión
TIPOS_INVERSION = [
    ("Infraestructura física", "Obras, ampliaciones, remodelaciones, mobiliario y equipamiento de oficina"),
    ("Equipamiento", "Equipamiento de laboratorios y talleres"),
    ("Infraestructura tecnológica", "Equipos de cómputo, redes, servidores, software de uso prolongado"),
]

# Ficha_Resumen sección 2 — las 9 categorías de ingresos de funcionamiento
CATEGORIAS_ING_RESUMEN = [
    "MATRICULA", "PENSIÓN", "DERECHOS", "DESCUENTOS", "OTROS INGRESOS", "FALTAS Y/O TARDANZAS",
    "RECUPERACION CUENTAS DE COBRANZA DUDOSA", "DEPÓSITOS EN INSTITUCIONES FINANCIERAS",
    "GANANCIA DIFERENCIA EN CAMBIO",
]

# Ficha_Resumen sección 5 — las 28 categorías presupuestales de egresos de funcionamiento
CATEGORIAS_EGR_RESUMEN = [
    "ADQUISICIONES", "ALQUILERES", "ASESORIAS Y CONSULTORIAS", "CAPACITACIONES", "CARGAS FINANCIERAS",
    "COMBUSTIBLES Y PEAJES", "DEPRECIACIÓN", "GASTOS ADMINISTRACIÓN", "LICENCIAS", "MANT, Y REPARACIONES",
    "OTRAS CARGAS ADMINISTRATIVAS", "OTROS SERVICIOS", "PASAJES", "PASAJES Y VIATICOS", "PLANILLA ADMINISTRATIVA",
    "PLANILLA OBREROS", "PLANILLA PERSONAL DOCENTE", "PROVISIONES", "PUBLICIDAD", "SEGURIDAD", "SEGUROS",
    "SERVICIOS BASICOS", "SERVICIOS DE TERCEROS", "SUMINISTROS", "SUSCRIPCIONES", "TRAMITES",
    "TRANSPORTE, CORREOS Y GASTOS DE VIAJE", "TRIBUTOS",
]

# Sección S — Costos unitarios de referencia (la DPDI completa costo y vigencia)
COSTOS_REFERENCIA = [
    ("Capacitación externa", "Participante", None, ""),
    ("Pasaje nacional", "Tramo", None, ""),
    ("Viáticos nacionales", "Día", None, ""),
    ("Servicio de consultoría", "Hora", None, ""),
    ("Licencia de software", "Usuario / año", None, ""),
    ("Equipo de cómputo estándar", "Unidad", None, ""),
]

# Unidades orgánicas iniciales (techos tomados de las fichas oficiales de referencia).
UNIDADES_INICIALES = [
    dict(codigo_area="OL", macrogrupo="Soporte Académico/Administrativo",
         nombre_unidad="DIRECCIÓN DE OPORTUNIDADES LABORALES Y SEGUIMIENTO DEL GRADUADO",
         centro_costo="DIRECCIÓN DE OPORTUNIDADES LABORALES Y SEGUIMIENTO DEL GRADUADO", genera_ingresos=0,
         techo_iniciativas=195543, techo_funcionamiento=338121, techo_otras_operativas=37200, techo_inversion=10000),
    dict(codigo_area="ECC-PRE", macrogrupo="Pregrado",
         nombre_unidad="ESCUELA DE CIENCIAS DE LA COMUNICACIÓN",
         centro_costo="ESCUELA DE CIENCIAS DE LA COMUNICACIÓN PRESENCIAL", genera_ingresos=1,
         techo_iniciativas=5300, techo_funcionamiento=1266050, techo_otras_operativas=101154, techo_inversion=10000,
         meta_resultado_propias=2000),
]

# Estructura anual del costo de planilla (régimen laboral peruano, en sueldos mensuales):
# 11 sueldos + 1 vacaciones + 2 gratificaciones con bonificación extraordinaria 9 % (2.18)
# + CTS (14/12 = 1.1667) + EsSalud 9 % sobre 12 remuneraciones (1.08).
PLANILLA_FACTORES = {"SUELDOS": 11.0, "VACACIONES": 1.0, "GRATIFICACIONES": 2.18, "CTS": 14.0 / 12.0, "ESSALUD": 1.08}
_PL_TOTAL = sum(PLANILLA_FACTORES.values())


def _pl(concepto: str) -> float:
    return round(PLANILLA_FACTORES[concepto] / _PL_TOTAL, 10)


# Plantilla base de «Correr Actividades de Funcionamiento».
# (macrogrupo, solo_genera_ingresos, tipo_flujo, partida, base_calculo, porcentaje, monto_fijo, aplica_ajuste, sustento)
# base_calculo: MONTO_FIJO | TECHO_PLANILLA_ADM | TECHO_PLANILLA_DOC | TECHO_MARKETING | META_INGRESOS_FUNC
#               | PENSION_PREGRADO | PENSION_POSGRADO | PROVISION_MOROSIDAD
PLANTILLA_FUNCIONAMIENTO = [
    ("*", 0, "Egresos", "SUELDOS ADMINISTRATIVOS", "TECHO_PLANILLA_ADM", _pl("SUELDOS"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "VACACIONES ADMINISTRACION", "TECHO_PLANILLA_ADM", _pl("VACACIONES"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "GRATIFICACIONES ADMINISTRACION", "TECHO_PLANILLA_ADM", _pl("GRATIFICACIONES"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "CTS ADMINISTRATIVOS", "TECHO_PLANILLA_ADM", _pl("CTS"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "ESSALUD ADMINISTRACION", "TECHO_PLANILLA_ADM", _pl("ESSALUD"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "SUELDOS TIEMPO COMP", "TECHO_PLANILLA_DOC", _pl("SUELDOS"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "VACACIONES TIEMPO COMPLETO", "TECHO_PLANILLA_DOC", _pl("VACACIONES"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "GRATIFICACIONES TIEMPO COMPLETO", "TECHO_PLANILLA_DOC", _pl("GRATIFICACIONES"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "CTS PERSONAL TIEMPO COMPLETO", "TECHO_PLANILLA_DOC", _pl("CTS"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "ESSALUD TIEMPO COMPLETO", "TECHO_PLANILLA_DOC", _pl("ESSALUD"), 0, "Sí", "Contrato vigente"),
    ("*", 0, "Egresos", "PUBLICIDAD", "TECHO_MARKETING", 1.0, 0, "No", "Costo de referencia"),
    ("*", 0, "Egresos", "EDIFICACIONES", "MONTO_FIJO", 1.0, 0, "No", "Histórico de ejecución"),
    ("*", 0, "Egresos", "MUEBLES Y ENSERES", "MONTO_FIJO", 1.0, 0, "No", "Histórico de ejecución"),
    ("*", 0, "Egresos", "EQUIPOS DIVERSOS", "MONTO_FIJO", 1.0, 0, "No", "Histórico de ejecución"),
    ("*", 0, "Egresos", "AGUA - DIRECTO", "MONTO_FIJO", 1.0, 0, "No", "Histórico de ejecución"),
    ("*", 0, "Egresos", "ENERGIA ELECTRICA ADMINISTRACION", "MONTO_FIJO", 1.0, 0, "No", "Histórico de ejecución"),
    ("*", 0, "Egresos", "INTERNET", "MONTO_FIJO", 1.0, 0, "No", "Contrato vigente"),
    ("*", 0, "Egresos", "TELEFONO VOIP", "MONTO_FIJO", 1.0, 0, "No", "Contrato vigente"),
    ("*", 1, "Egresos", "CUENTAS POR COBRAR COMERCIALES - TERCEROS", "PROVISION_MOROSIDAD", 1.0, 0, "No", "Histórico de ejecución"),
    ("Pregrado", 1, "Ingresos", "PENSION PREGRADO", "PENSION_PREGRADO", 1.0, 0, "Sí", "Proyección de matrícula aprobada"),
    ("Posgrado", 1, "Ingresos", "PENSION POST GRADO", "PENSION_POSGRADO", 1.0, 0, "Sí", "Proyección de matrícula aprobada"),
    ("Educación Continua", 1, "Ingresos", "PENSION PEA", "META_INGRESOS_FUNC", 1.0, 0, "Sí", "Histórico de recaudación"),
]

BASES_CALCULO = ["MONTO_FIJO", "TECHO_PLANILLA_ADM", "TECHO_PLANILLA_DOC", "TECHO_MARKETING", "META_INGRESOS_FUNC",
                 "PENSION_PREGRADO", "PENSION_POSGRADO", "PROVISION_MOROSIDAD"]

# Secciones I, J y N de «Parámetros» + índice de códigos NIIF (naturaleza / resultado por función).
NIIF_PARTIDAS = [
    # (codigo_niif, partida_rubro, partida_gg, categoria, tipo_flujo, naturaleza, resultado_funcion)
    ('6327', 'ACREDITACION', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6321', 'ADMINISTRATIVA', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63631', 'AGUA - DIRECTO', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63633', 'AGUA - VENTAS', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('63142', 'ALIMENTACION ADMINISTRACION', 'Alimentacion administracion', 'ALIMENTACIÓN', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63141', 'ALIMENTACION DIRECTO', 'Alimentacion directo', 'ALIMENTACIÓN', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63143', 'ALIMENTACION VENTAS', 'Alimentacion ventas', 'ALIMENTACIÓN', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('63132', 'ALOJAMIENTO ADMINISTRACION', 'Alojamiento administracion', 'ALOJAMIENTO', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63131', 'ALOJAMIENTO DIRECTO', 'Alojamiento directo', 'ALOJAMIENTO', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63133', 'ALOJAMIENTO VENTAS', '', 'ALOJAMIENTO', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('63532', 'ALQUILER DE MAQUINARIA Y EQUIPOS - DIRECTO', '', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6357', 'ALQUILERES DE ACTIVOS PUBLICITARIOS', 'Alquileres de activos publicitarios', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6358', 'ALQUILERES DE OTROS ACTIVOS', '', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6352', 'ALQUILERES EDIFICACIONES', 'Alquileres edificaciones', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6354', 'ALQUILERES EQUIPO DE TRANSPORTE', '', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6355', 'ALQUILERES EQUIPOS DIVERSOS', 'Alquileres equipos diversos', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6356', 'ALQUILERES MUEBLES Y ENSERES', 'Alquileres muebles y enseres', 'ALQUILERES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('643202', 'AREAS VERDES', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('62502', 'ATENCION AL PERSONAL ADMINISTRACION', '', 'GASTOS ADMINISTRACIÓN', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62505', 'ATENCION AL PERSONAL MODALIDAD FORMATIVA', '', 'GASTOS ADMINISTRACIÓN', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6323', 'AUDITORIA Y CONTABLE', 'Auditoria Contable', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639915', 'AULA VIRTUAL ZOOM', 'Aula Virtual Zoom', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6941010800', 'BEBIDAS Y COMESTIBLES', 'Bebidas y comestibles', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('637107', 'BRANDEO INSTITUCIONAL', '', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('6367', 'CABLE', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('62402', 'CAPACITACION ADMINISTRACION', 'Capacitacion administracion', 'CAPACITACIONES', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62401', 'CAPACITACION FUNCIONARIOS', '', 'CAPACITACIONES', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62403', 'CAPACITACION TIEMPO COMPLETO', 'Capacitacion tiempo completo', 'CAPACITACIONES', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('62404', 'CAPACITACION TIEMPO PARCIAL', 'Capacitacion tiempo parcial', 'CAPACITACIONES', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('659905', 'CARGAS DIV. DE EJERC. ANTERIORES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('659999', 'CARGAS DIVERSAS NO DEDUCIBLES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('62209', 'CARGAS PERSONAL NO DEDUCIBLES', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('637402', 'CESION EN USO MARCA EDUCACION CONTINUA', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('637401', 'CESION EN USO MARCA POSGRADO', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('69111031', 'CHAQUETAS Y SIMILARES', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('656013', 'COMBUSTIBLES Y LUBRICANTES - LICUADOS', '', 'COMBUSTIBLES Y PEAJES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('656012', 'COMBUSTIBLES Y LUBRICANTES - LIQUIDOS', '', 'COMBUSTIBLES Y PEAJES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'COSTO DE VENTAS'),
    ('656014', 'COMBUSTIBLES Y LUBRICANTES - OTROS', '', 'COMBUSTIBLES Y PEAJES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE VENTAS'),
    ('62122', 'COMISIONES ADMINISTRATIVOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62121', 'COMISIONES FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6033', 'COMPRAS DIRECTAS', 'Compras directas', 'ADQUISICIONES', 'Egresos', 'MATERIALES Y SUMINISTROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('68211', 'CONCESIONES, LICENCIAS Y OTROS DERECHOS', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('6211108', 'CONDICION DE TRABAJO - VIVIENDA', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('67321', 'CONTRATOS DE ARRENDAMIENTO FINANCIERO', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6312', 'CORREOS', 'Correos', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('629102', 'CTS ADMINISTRATIVOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('629101', 'CTS FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('629106', 'CTS OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('629103', 'CTS PERSONAL TIEMPO COMPLETO', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('68411', 'CUENTAS POR COBRAR COMERCIALES - TERCEROS', '', 'PROVISIONES', 'Egresos', 'CUENTAS POR COBRAR', 'GASTOS DE ADMINISTRACIÓN'),
    ('68415', 'CUENTAS POR COBRAR DIVERSAS - RELACIONADAS', '', 'PROVISIONES', 'Egresos', 'CUENTAS POR COBRAR', 'GASTOS DE ADMINISTRACIÓN'),
    ('681332', 'DEPRECIACION ARRENDAMIENTO BCP 54673AFB - AUTOMOVIL', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('681331', 'DEPRECIACION ARRENDAMIENTO BCP-51119AFB CAMIONETA', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('681353', 'DEPRECIACION ARRENDAMIENTO FINANCIERO BCP- 63133AFB EQUIPOS DE COMPUTO', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('681354', 'DEPRECIACION ARRENDAMIENTO FINANCIERO BCP- 64546AFB EQUIPOS DE COMPUTO', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('681351', 'DEPRECIACION ARRENDAMIENTO FINANCIERO BCP-51093AFB PC', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('639910', 'DESARROLLOS DE TALLERES', '', 'SERVICIOS DE TERCEROS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6941010900', 'DESCARTABLES', 'Descartables', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6751', 'DESCUENTOS CONCEDIDOS POR PRONTO PAGO', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6281', 'DIETAS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('639914', 'DISPOSICION DE RESIDUOS', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('65934', 'DISTINCION CONDECORAC. Y PRESENTES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE VENTAS'),
    ('65911', 'DONACIONES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('68161', 'EDIFICACIONES', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68141', 'EDIFICACIONES DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68167', 'EDIFICACIONES LEY 30264', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('631114', 'ENCOMIENDAS', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63612', 'ENERGIA ELECTRICA ADMINISTRACION', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63611', 'ENERGIA ELECTRICA DIRECTO', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63613', 'ENERGIA ELECTRICA VENTAS', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('6364040000', 'ENTEL PERU SA', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('68165', 'EQUIPOS DIVERSOS', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68145', 'EQUIPOS DIVERSOS DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('627801', 'ESSALUD + VIDA CONSTRUCCION CIVIL', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('627102', 'ESSALUD ADMINISTRACION', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('627101', 'ESSALUD FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('627106', 'ESSALUD OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('627103', 'ESSALUD TIEMPO COMPLETO', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('627104', 'ESSALUD TIEMPO PARCIAL', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('631113', 'ESTACIONAMIENTOS', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('639903', 'FOTOCOPIAS,IMPRESIONES, PLOTTERS', 'Fotocopias,impresiones, plotters', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639803', 'GASTOS DE FOTOCOPIADO', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639801', 'GASTOS DE IMPRENTA', 'Gastos de imprenta', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6392', 'GASTOS DE LABORATORIO', '', 'ADQUISICIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('659912', 'GASTOS DE PUBLICACIONES CIENTIFICAS', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('65933', 'GASTOS DE REPRESENTACIÓN', 'Gastos de representación', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE VENTAS'),
    ('6221', 'GASTOS DE VIAJE', 'Gastos de viaje', 'PASAJES Y VIATICOS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6394', 'GASTOS EQUIPOS DIVERSOS', 'Gastos equipos diversos', 'ADQUISICIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('6395', 'GASTOS MUEBLES', '', 'ADQUISICIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('659906', 'GASTOS POR PREMIOS DE INVESTIGACION E INNOVACION', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6490', 'GASTOS POR TRIBUTOS NO DEDUCIBLES', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('65931', 'GASTOS DE PROMOCIÓN CAPTACIÓN', 'Gastos de promoción captación', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE VENTAS'),
    ('659911', 'GASTOS PROYECTOS DE RESPONSABILIDAD SOCIAL', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('659910', 'GASTOS RECREATIVOS', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6391', 'GASTOS,COMISIONES, PORTES, MANTENIMIENTO BANCARIOS', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS FINANCIEROS'),
    ('621408', 'GRATIFICACION ADMINISTRATIVO TIEMPO PARCIAL', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('621402', 'GRATIFICACIONES ADMINISTRACION', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621401', 'GRATIFICACIONES FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621405', 'GRATIFICACIONES MODALIDAD FORMATIVA', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621406', 'GRATIFICACIONES OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('621407', 'GRATIFICACIONES RESP. DE CARRERA', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621403', 'GRATIFICACIONES TIEMPO COMPLETO', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('621404', 'GRATIFICACIONES TIEMPO PARCIAL', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('6418', 'IGV NO DOMICILIADOS', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6417', 'IMPUESTO A LA RENTA NO DOMICILIADOS', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6412', 'IMPUESTO A LAS TRANSACCIONES FINANCIERAS', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS FINANCIEROS'),
    ('6416', 'IMPUESTO AL CONSUMO DE LAS BOLSAS DE PLASTICO', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6433', 'IMPUESTO AL PATRIMONIO VEHICULAR', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6411', 'IMPUESTO GENERAL A LAS VENTAS Y SELECTIVO AL CONSUMO', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6431', 'IMPUESTO PREDIAL', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('62302', 'INDEMNIZACIONES ADMINISTRACION', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6941011400', 'INSUMOS QUIMICOS FISCALIZADOS', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('659907', 'INTERES LEGAL DEVUELTO', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('659211', 'INTERESES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('673111', 'INTERESES Y GASTOS DE PRESTAMO', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('673112', 'INTERESES Y GASTOS DE SOBREGIROS', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6365', 'INTERNET', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6326', 'INVESTIGACION Y DESARROLLO', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6322', 'LEGAL Y TRIBUTARIA', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('68166', 'LIBROS', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('69111014', 'LIBROS DE FRANCES', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6911101010', 'LIBROS DE INGLES', 'Libros de ingles', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('69111012', 'LIBROS DE PORTUGUES', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('69111013', 'LIBROS DIGITALES', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('68148', 'LIBROS DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('65401', 'LICENCIAS Y DERECHOS DE VIGENCIA', 'Licencias y derechos de vigencia', 'LICENCIAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('643201', 'LIMPIEZA PUBLICA', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('634601', 'MANT. Y REPAR. EDIFICACIONES', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634607', 'MANT. Y REPAR. EQUIPOS DE COMPUTO', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634605', 'MANT. Y REPAR. EQUIPOS DIVERSOS', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634602', 'MANT. Y REPAR. MAQUINARIAS Y EQUIPOS', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634604', 'MANT. Y REPAR. MUEBLES Y ENSERES', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634606', 'MANT. Y REPAR. OTROS ACTIVOS', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('634603', 'MANT. Y REPAR. UNIDAD DE TRANSPORTE', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('63433', 'MANTENIMIENTO DE MAQUINARIA Y EQUIPO ADMINISTRACION', 'Mantenimiento de maquinaria y equipo administracion', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63432', 'MANTENIMIENTO MAQUINARIA Y  EQUIPO - DIRECTO', '', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6346', 'MANTENIMIENTO Y REPARAC. INMUEBLES MAQ. Y EQUIPOS', 'Mantenimiento y reparac. Inmuebles maq. Y equipos', 'MANT, Y REPARACIONES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('681423', 'MAQ. Y EQUIPO - ADMINISTRACIÓN DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68162', 'MAQUINARIAS Y EQUIPOS', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('6941010500', 'MAT. LAB. DIVERSO', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941010200', 'MATERIAL DE ASEO', 'Material de aseo', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941011100', 'MATERIAL DE COMPUTO', 'Material de computo', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941010300', 'MATERIAL DE FERRETERIA', 'Material de ferreteria', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941020300', 'MATERIAL DE FERRETERIA', 'Material de ferreteria', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941010100', 'MATERIAL DE OFICINA', 'Material de oficina', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941020100', 'MATERIAL DE OFICINA', 'Material de oficina', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941010400', 'MATERIAL PUBLICITARIO', 'Material publicitario', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6031', 'MATERIALES AUXILIARES', '', 'ADQUISICIONES', 'Egresos', 'MATERIALES Y SUMINISTROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6941010700', 'MATERIALES DE SEGURIDAD', 'Materiales de seguridad', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6941010600', 'MEDICINA GENERAL', 'Medicina general', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6325', 'MEDIOAMBIENTAL', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('691110401', 'MERCADERIAS', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('6324', 'MERCADOTECNIA', 'Mercadotecnia', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('6941011000', 'MODULOS E IMPRESOS', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('631124', 'MOVILIDAD', 'Movilidad', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('68164', 'MUEBLES Y ENSERES', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('685242', 'MUEBLES Y ENSERES - DIRECTO', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68144', 'MUEBLES Y ENSERES DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('67371', 'OBLIGACIONES TRIBUTARIAS', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6211201', 'OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('6223', 'OTRAS CARGAS DE PERSONAL', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('659909', 'OTRAS CARGAS DIVERSAS DE GESTION', 'Otras cargas diversas de gestion', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('673121', 'OTRAS ENTIDADES', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6911199', 'OTRAS MERCADERIAS', '', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('62202', 'OTRAS REMUNERACIONES ADMINISTRACION', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62201', 'OTRAS REMUNERACIONES FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62206', 'OTRAS REMUNERACIONES OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('62203', 'OTRAS REMUNERACIONES TIEMPO COMPLETO', 'Otras remuneraciones tiempo completo', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('62204', 'OTRAS REMUNERACIONES TIEMPO PARCIAL', 'Otras remuneraciones tiempo parcial', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('65929', 'OTRAS SANCIONES', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6329', 'OTROS', '', 'ASESORIAS Y CONSULTORIAS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6419', 'OTROS', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6439', 'OTROS', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('629301', 'OTROS BENEFICIOS POST-EMPLEO', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('637109', 'OTROS GASTOS DE PUBLICIDAD', 'Otros gastos de publicidad', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('63152', 'OTROS GASTOS DE VIAJE ADMINISTRACION', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63151', 'OTROS GASTOS DE VIAJE DIRECTO', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('6449', 'OTROS GASTOS POR TRIBUTOS', 'Otros gastos por tributos', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639909', 'OTROS SERVICIOS', 'Otros servicios', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63809', 'OTROS SERVICIOS DE PERSONAL', '', 'SEGURIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6941011200', 'OTROS SUM. DIVERSOS', 'Otros sum. Diversos', 'SUMINISTROS', 'Egresos', 'MATERIALES Y SUMINISTROS', 'COSTO DE VENTAS'),
    ('62208', 'PARTICIPACION DE LOS TRABAJADORES', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62210102', 'PASAJES EXTERIOR', 'Pasajes exterior', 'PASAJES Y VIATICOS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('631611', 'PASAJES EXTERIOR', 'Pasajes exterior', 'PASAJES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('62210101', 'PASAJES NACIONAL', 'Pasajes nacionales', 'PASAJES Y VIATICOS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('631610', 'PASAJES NACIONALES', 'Pasajes nacionales', 'PASAJES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('657', 'PEAJES', '', 'COMBUSTIBLES Y PEAJES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6761', 'PERDIDA DIFERENCIA DE CAMBIO', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('6721', 'PÉRDIDA POR INSTRUMENTOS FINANCIEROS DERIVADOS', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('659903', 'PERDIDA POR REDONDEO', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6711', 'PRÉSTAMOS DE INSTITUCIONES FINANCIERAS Y OTRAS ENTIDADES', '', 'CARGAS FINANCIERAS', 'Egresos', 'CARGAS FINANCIERAS', 'GASTOS FINANCIEROS'),
    ('68213', 'PROGRAMAS DE COMPUTADORA (SOFTWARE)', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'GASTOS DE ADMINISTRACIÓN'),
    ('6372', 'PUBLICACIONES', 'Publicaciones', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637202', 'PUBLICACIONES REVISTAS', '', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637203', 'PUBLICACIONES VARIAS', '', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('6371', 'PUBLICIDAD', 'Publicidad', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637104', 'PUBLICIDAD PANELES', 'Publicidad paneles', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637106', 'PUBLICIDAD PLATAFORMAS DIGITALES', 'Publicidad plataformas digitales', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637102', 'PUBLICIDAD RADIO', 'Publicidad radio', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('637101', 'PUBLICIDAD TELEVISION', 'Publicidad televisión', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('659901', 'REFRIGERIOS', 'Refrigerios', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('63731', 'RELACIONES PUBLICAS', '', 'PUBLICIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('65921', 'SANCIONES SUNAT', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('627301', 'SEGURO COMPLEMENTARIO DE TRABAJO DE RIESGO, ACCIDENTES DE TRABAJO Y ENFERMEDADES PROFESIONALES', '', 'SEGUROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('627401', 'SEGURO DE VIDA', '', 'SEGUROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('627601', 'SEGURO PARTICULAR DE PRESTACION DE SALUD - MODALIDAD FORMATIVA', '', 'SEGUROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('651022', 'SEGUROS  MAQ. Y EQUIPO - DIRECTO', '', 'SEGUROS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'COSTO DE VENTAS'),
    ('627501', 'SEGUROS PARTICULARES DE PRESTACIONES DE SALUD - EPS Y OTROS PARTÍCULARES', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('65109', 'SEGUROS VARIOS', '', 'SEGUROS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('65103', 'SEGUROS VEHICULOS', '', 'SEGUROS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('643203', 'SERENAZGO', '', 'TRIBUTOS', 'Egresos', 'TRIBUTOS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639916', 'SERVICIO DE AMBULANCIA', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639910', 'SERVICIO DESARROLLO DE TALLERES', 'Servicio desarrollo de talleres', 'SERVICIOS DE TERCEROS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('639913', 'SERVICIO DE SEGURIDAD INFORMATICA', 'Servicio de seguridad informatica', 'SEGURIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639911', 'SERVICIO DE SEGURIDAD Y VIGILANCIA', '', 'SEGURIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639912', 'SERVICIO MONITOREO GPS', '', 'SEGURIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('659902', 'SERVICIOS - BREAK', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('639906', 'SERVICIOS DE LAVANDERIA', 'Servicios de lavanderia', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639999', 'SERVICIOS DE TERCEROS NO DEDUCIBLES', '', 'OTROS SERVICIOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('621122', 'SOBRE TIEMPO PERSONAL OBRERO', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('62205', 'SUBVENCIONES MODALIDAD FORMATIVA', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6211107', 'SUELDOS ADMINISTRATIVO TIEMPO PARCIAL', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('6211102', 'SUELDOS ADMINISTRATIVOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6211101', 'SUELDOS FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6211105', 'SUELDOS MODALIDAD FORMATIVA', 'Sueldos modalidad formativa', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('6211103', 'SUELDOS TIEMPO COMP', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('6211104', 'SUELDOS TIEMPO PARCIAL', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('656031', 'SUMINISTROS DIVERSOS', '', 'OTRAS CARGAS ADMINISTRATIVAS', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6532', 'SUSCRIPCIONES EXTRANJERAS', 'Suscripciones extranjeras', 'SUSCRIPCIONES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6531', 'SUSCRIPCIONES NACIONALES', 'Suscripciones nacionales', 'SUSCRIPCIONES', 'Egresos', 'CARGAS DIVERSAS DE GESTIÓN', 'GASTOS DE ADMINISTRACIÓN'),
    ('6364', 'TELÉFONO', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('6364020000', 'TELÉFONO BITEL', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('636406', 'TELEFONO VOIP', '', 'SERVICIOS BASICOS', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639902', 'TRAMITES ADMINISTRATIVOS', 'Tramites administrativos', 'TRAMITES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('639901', 'TRAMITES JUDICIALES Y NOTARIAL', '', 'TRAMITES', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('63111', 'TRANSPORTE DE CARGA', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('631122', 'TRANSPORTE DE PASAJEROS ADMINISTRACION', 'Transporte de pasajeros administracion', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('631121', 'TRANSPORTE DE PASAJEROS DIRECTO', 'Transporte de pasajeros directo', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('631123', 'TRANSPORTE DE PASAJEROS VENTAS', 'Transporte de pasajeros ventas', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE VENTAS'),
    ('631131', 'TRANSPORTE DE RESIDUOS', '', 'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'COSTO DE VENTAS'),
    ('68163', 'UNIDADES DE TRANSPORTE', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('68143', 'UNIDADES DE TRANSPORTE DL 882', '', 'DEPRECIACIÓN', 'Egresos', 'DEPRECIACION', 'COSTO DE VENTAS'),
    ('621502', 'VACACIONES ADMINISTRACION', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621501', 'VACACIONES FUNCIONARIOS', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621505', 'VACACIONES MODALIDAD FORMATIVA', '', 'PLANILLA ADMINISTRATIVA', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('621506', 'VACACIONES OBREROS', '', 'PLANILLA OBREROS', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('621503', 'VACACIONES TIEMPO COMPLETO', '', 'PLANILLA PERSONAL DOCENTE', 'Egresos', 'CARGAS DE PERSONAL', 'COSTO DE VENTAS'),
    ('62210202', 'VIATICOS PERSONAL ACADEMICO', '', 'PASAJES Y VIATICOS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('62210201', 'VIATICOS PERSONAL ADMINISTRATIVO', '', 'PASAJES Y VIATICOS', 'Egresos', 'CARGAS DE PERSONAL', 'GASTOS DE ADMINISTRACIÓN'),
    ('63801', 'VIGILANCIA Y SEGURIDAD', '', 'SEGURIDAD', 'Egresos', 'SERVICIOS PRESTADOS POR TERCEROS', 'GASTOS DE ADMINISTRACIÓN'),
    ('7011101', 'LIBROS', '', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041040300', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - CENTRO DE IDIOMAS', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - CENTRO DE IDIOMAS', 'PRODUCCIÓN', 'Ingresos', '', ''),
    ('7041040200', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - FUTURE', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - FUTURE', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041046000', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PEAS', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PEAS', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041040100', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041040700', 'CONGRESOS. SEMINARIOS, TALLERES, OTROS POSGRADO', 'CONGRESOS, SEMINARIOS, TALLERES Y OTROS - POSGRADO', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('XXXXXXXXXX', 'Materiales', 'Materiales', 'PRODUCCIÓN', 'Ingresos', '', ''),
    ('7041010200', 'MATRICULA FUTURE', 'MATRICULA FUTURE', 'MATRICULA', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041010400', 'MATRICULA POST GRADO', 'MATRICULA POSGRADO', 'MATRICULA', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041010100', 'MATRICULA PREGRADO', 'MATRICULA PREGRADO', 'MATRICULA', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041010500', 'MATRICULA SEGUNDA ESPECIALIDAD', 'MATRICULA SEGUNDA ESPECIALIDAD', 'MATRICULA', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('704113010', 'MERCHANDISING', 'MERCHANDISING', 'PRODUCCIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020200', 'PENSION FUTURE', 'PENSION FUTURE', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020300', 'PENSION CENTRO DE IDIOMAS', 'PENSION CENTRO DE IDIOMAS', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020600', 'PENSION PEA', 'PENSION PEA', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020700', 'PENSION POST GRADO', 'PENSION POSGRADO', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020100', 'PENSION PREGRADO', 'PENSION PREGRADO', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041020800', 'PENSION SEGUNDA ESPECIALIDAD', 'PENSION SEGUNDA ESPECIALIDAD', 'PENSIÓN', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041031000', 'TASAS CARPETA SEGUNDA ESPECIALIDAD', 'TASAS CARPETA SEGUNDA ESPECIALIDAD', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041030200', 'TASAS CARPETAS  FUTURE', 'TASAS CARPETAS FUTURE', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041030300', 'TASAS CARPETAS CENTRO DE IDIOMAS', 'TASAS CARPETAS CENTRO DE IDIOMAS', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041030900', 'TASAS CARPETAS PEA', 'TASAS CARPETAS PEA', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041030800', 'TASAS CARPETAS POST GRADO', 'TASAS CARPETAS POSGRADO', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7041030100', 'TASAS CARPETAS PREGRADO', 'TASAS CARPETAS PREGRADO', 'DERECHOS', 'Ingresos', 'INGRESO POR VENTAS', 'INGRESO BRUTO POR VENTAS'),
    ('7411', 'DSTOS,  REBAJ, BONIF CONCEDIDOS TERCEROS', '', 'DESCUENTOS', 'Ingresos', 'SUBVENCIONES', 'DESCUENTOS'),
    ('7542', 'EDIFICACIONES', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7598', 'FALTAS Y/O TARDANZAS', '', 'FALTAS Y/O TARDANZAS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599070000', 'INGRESOS MODALIDADES FORMATIVAS', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599010000', 'INGRESOS PERSONAL ADMINISTRATIVOS', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599040000', 'INGRESOS PERSONAL FUNCIONARIOS', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599020000', 'INGRESOS PERSONAL TIEMPO COMPLETO', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599030000', 'INGRESOS PERSONAL TIEMPO PARCIAL', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599', 'OTROS INGRESOS', '', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7599990000', 'OTROS INGRESOS GESTION', 'OTROS INGRESOS DE GESTIÓN', 'OTROS INGRESOS', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7551', 'RECUPERACIÓN - CUENTAS DE COBRANZA DUDOSA', '', 'RECUPERACION CUENTAS DE COBRANZA DUDOSA', 'Ingresos', 'INGRESOS DIVERSOS', 'OTROS INGRESOS'),
    ('7721', 'DEPÓSITOS EN INSTITUCIONES FINANCIERAS', '', 'DEPÓSITOS EN INSTITUCIONES FINANCIERAS', 'Ingresos', 'INGRESOS FINANCIEROS', 'INGRESOS FINANCIEROS'),
    ('7761', 'GANANCIA DIFERENCIA EN CAMBIO', '', 'GANANCIA DIFERENCIA EN CAMBIO', 'Ingresos', 'INGRESOS FINANCIEROS', 'INGRESOS FINANCIEROS'),
    ('381', 'Bienes de arte y cultutra', 'Bienes de arte y cultutra', 'BIBLIOGRAFIA', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('33921504', 'COSTOS INDIRECTOS DE OBRA', 'COSTOS INDIRECTOS DE OBRA', 'INFRAESTRUCTURA', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('3362', 'Equipos de comunicación', 'Equipos de comunicación', 'EQUIPOS', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('3361', 'Equipos para procesamiento de información', 'Equipos para procesamiento de información', 'EQUIPOS DE CÓMPUTO', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('333', 'Maquinarias y equipos de explotación', 'Maquinarias y equipos de explotación', 'EQUIPOS', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('33921501', 'MATERIALES DE OBRA', 'MATERIALES DE OBRA', 'INFRAESTRUCTURA', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('335', 'MUEBLES Y ENSERES', 'MUEBLES Y ENSERES', 'INFRAESTRUCTURA', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('336911', 'OTROS EQUIPOS COSTO', 'OTROS EQUIPOS COSTO', 'EQUIPOS', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
    ('33921503', 'SERVICIOS DE TERCEROS PARA OBRA', 'SERVICIOS DE TERCEROS PARA OBRA', 'INFRAESTRUCTURA', 'Inversion', 'ACTIVO FIJO', 'INVERSIÓN'),
]

PARTIDAS_GG = [
    # (partida_gg, descripcion, tipo / codigo NIIF de inversión)
    ('Alimentacion administracion', 'Alimentación del personal administrativo.', 'Egresos'),
    ('Alimentacion directo', 'Alimentacion de ponentes, conferencistas, docentes.', 'Egresos'),
    ('Alimentacion ventas', 'Gastos por alimentación personal de apoyo en ventas.', 'Egresos'),
    ('Alojamiento administracion', 'Hospedaje personal administrativo.', 'Egresos'),
    ('Alojamiento directo', 'Hospedaje ponentes, expositores, conferencistas nacionales e internacionales.', 'Egresos'),
    ('Alquiler maquinaria y equipos', 'Alquiler de maquinarias y equipos para la construcción de edificios, etc.', 'Egresos'),
    ('Alquileres de activos publicitarios', 'Alquiler y diseño de módulo / stand', 'Egresos'),
    ('Alquileres edificaciones', 'Alquiler sala de conferencias (ejemplo: westin, etc.)', 'Egresos'),
    ('Alquileres equipos diversos', 'Alquiler de equipos para laboratorios de ciencias e ingenierías, laboratorios de ciencias de la salud', 'Egresos'),
    ('Alquileres muebles y enseres', 'Alquileres de auditorio /sonido/ sillas/ toldos/ ecram.', 'Egresos'),
    ('Auditoria Contable', 'Asesoria y auditoria contable', 'Egresos'),
    ('Aula Virtual Zoom', 'Servicio Plataforma Zoom', 'Egresos'),
    ('Bebidas y comestibles', 'Agua, gaseosa,panetón,jugo, bidon de agua,galletas,chocolate, azucar,café...', 'Egresos'),
    ('Capacitacion administracion', 'Gastos de capacitación personal administrativo', 'Egresos'),
    ('Capacitacion tiempo completo', 'Gastos de capacitación dtc', 'Egresos'),
    ('Capacitacion tiempo parcial', 'Gastos de capacitación dtp', 'Egresos'),
    ('Compras directas', 'Compras entregadas directamente al área, no ingresan al almacén', 'Egresos'),
    ('Correos', 'Mailing, envio correos', 'Egresos'),
    ('Descartables', 'Plato,cucharita,sorbetes,bolsa,vaso plático,etc.', 'Egresos'),
    ('Fotocopias,impresiones, plotters', 'Fotocopias certificados, constancias,copias, diplomas, examenes, impresión (fotos,revistas, pioner), ploteos,test psicológicos wais, tarjetas aprobados y desaprobados..', 'Egresos'),
    ('Gastos de imprenta', 'Gastos de imprenta', 'Egresos'),
    ('Gastos de promoción captación', 'Auspicios (banderola, torta, torta), trofeos expositores, gastos de representación, eventos de lanzamiento.', 'Egresos'),
    ('Gastos de representación', 'Gastos de representación', 'Egresos'),
    ('Gastos de viaje', 'Gastos de viaje personal administrativo en planilla.', 'Egresos'),
    ('Gastos equipos diversos', 'Gastos equipos diversos', 'Egresos'),
    ('Libros de ingles', 'Libros de inglés', 'Egresos'),
    ('Licencias y derechos de vigencia', 'Pago de licencias', 'Egresos'),
    ('Mantenimiento de maquinaria y equipo administracion', 'Mantenimiento de maquinaria y equipo administracion', 'Egresos'),
    ('Mantenimiento y reparac. Inmuebles maq. Y equipos', 'Servicio mantenimiento y reparación de equipos y edificaciones block abcdegi- servicio de mantenimiento preventivo', 'Egresos'),
    ('Material de aseo', 'Escoba, baldes, guantes,papelera,trapeador, escobilla,recogedor, ambientador,etc.', 'Egresos'),
    ('Material de computo', 'Disco duro, memoria usb, mouse, estabilizador,tablet, teléfono ip, placa multimedia, switch, splitter, teclado, cable de red...', 'Egresos'),
    ('Material de ferreteria', 'Adaptador,tubo,valvula, accesorios pvc,llave temporizada,bisagra,candado,acople,interruptor,zoquete,foco,tomacorriente,fluorescente,pegamento,tinner,lija,cable....', 'Egresos'),
    ('Material de oficina', 'Papel bond,toner, cinta, plumones,resaltador,lapicero,sello,folder,archivador,cartulina,clips,post it,tijera,vinifan...', 'Egresos'),
    ('Material publicitario', 'Porta fotocheck, medalla de premiación, lapicero con logo,tomatodo,bolsa con logo, calendario, pioner con logo,folder con logo, polos, gorra, invitaciones, diptico,volantes,afiche,gigantografías,triptico, pases....', 'Egresos'),
    ('Materiales de seguridad', 'Micas para gabinetes de extintores, pedestal metálico para extintor, embudo de vidrio, laminilla cubre objeto...', 'Egresos'),
    ('Materiales e insumos quimicos', 'Embudo de vidrio, pipeta, espátula serologica,matraz,pinza,tapón, fiola,probeta, vaso de precipitación, formol, tubo de vidrio, alambre de cobre, gradilla, pipeta, electrodo, capsula de evaporación, sodio, potasio,manganeso, amonio, tricloruro, cloroformo', 'Egresos'),
    ('Medicina general', 'Ibuprofeno, hirudoid, naproxeno, amoxicilina,dexametasona,mentholatun,jarabe,venda,cateter,agua oxigenada,esparadrapo,gasa...', 'Egresos'),
    ('Mercadotecnia', 'Estudio de mercado', 'Egresos'),
    ('Modulos academicos', 'Módulos académicos e impresos', 'Egresos'),
    ('Movilidad', 'Movilidad local', 'Egresos'),
    ('Otras cargas diversas de gestion', 'Otras cargas', 'Egresos'),
    ('Otras cargas diversas de gestion', 'Otras cargas diversas', 'Egresos'),
    ('Otras remuneraciones tiempo completo', 'Otras remuneraciones tiempo completo', 'Egresos'),
    ('Otras remuneraciones tiempo parcial', 'Pago por asesoría de tesis', 'Egresos'),
    ('Otros gastos de publicidad', 'Otros gastos de publicidad', 'Egresos'),
    ('Otros gastos por tributos', 'Otros gastos por tributos.', 'Egresos'),
    ('Otros servicios', 'Mensaje texto,graduación alquiler de vestimenta, shows,danza, arbitraje,pago docentes de talleres artísticos, pago consultoría nuevas carreras, constanias, diplomas, filmación,animacion de eventos, servicio de lavandería,courier, video', 'Egresos'),
    ('Otros sum. Diversos', 'Encendedor, llaveros, jarra de plástico,removedor bambu,azafate,mantel,vinchas navideñas, azicarero, acetona, pelota, toalla, inflador...', 'Egresos'),
    ('Pasajes exterior', 'Pasajes internacionales', 'Egresos'),
    ('Pasajes nacionales', 'Pasajes nacionales', 'Egresos'),
    ('Publicaciones', 'Publicación en revistas y suplementos.', 'Egresos'),
    ('Publicidad', 'Sponsor deportivo,publicidad plataformas digitales,publicidad en prensa escrita,paneles, auspicios de tv, fanpage, campaña inicio de ciclo, carnet universitario.', 'Egresos'),
    ('Publicidad paneles', 'Publicación en paneles', 'Egresos'),
    ('Publicidad plataformas digitales', 'Publicación en medios digitales', 'Egresos'),
    ('Publicidad radio', 'Publicidad radial', 'Egresos'),
    ('Publicidad televisión', 'Publicidad en tv', 'Egresos'),
    ('Refrigerios', 'Refrigerios de personal', 'Egresos'),
    ('Servicio de seguridad informatica', 'Servicio de seguridad informatica', 'Egresos'),
    ('Servicio desarrollo de talleres', 'Pago a ponentes de talleres, charlas, clases magistrales, conferencias, seminarios', 'Egresos'),
    ('Servicios de lavanderia', 'Servicios de lavanderia', 'Egresos'),
    ('Sueldos modalidad formativa', 'Pago a practicantes', 'Egresos'),
    ('Suscripciones extranjeras', 'Pago suscripciones, membresia y revistas internacionales', 'Egresos'),
    ('Suscripciones nacionales', 'Pago suscripciones,inscripciones, membresia y revistas nacionales', 'Egresos'),
    ('Tramites administrativos', 'Pago de isbn(libros biblioteca nacional)pagio issn (revistas científicas)', 'Egresos'),
    ('Transporte de pasajeros administracion', 'Movilidad personal administrativo', 'Egresos'),
    ('Transporte de pasajeros directo', 'Alquiler de buses para visitas académicas estudiantes,movilidad para traslado de ponentes, expositores, conferencistas y docentes.', 'Egresos'),
    ('Transporte de pasajeros ventas', 'Gastos de movilidad del personal de apoyo en ventas.', 'Egresos'),
    ('Bienes de arte y cultutra', 'activos con valor histórico, artístico o cultural propiedad de la institución, adquiridos para su conservación y exhibición.', '381'),
    ('COSTOS INDIRECTOS DE OBRA', '', '33921504'),
    ('Equipos de comunicación', '', '3362'),
    ('Equipos para procesamiento de información', 'Equipos de cómputo, Equipos informáticos', '3361'),
    ('Maquinarias y equipos de explotación', '', '333'),
    ('MATERIALES DE OBRA', '', '33921501'),
    ('MUEBLES Y ENSERES', '', '335'),
    ('OTROS EQUIPOS COSTO', '', '336911'),
    ('SERVICIOS DE TERCEROS PARA OBRA', '', '33921503'),
    ('CONGRESOS, SEMINARIOS, TALLERES Y OTROS - CENTRO DE IDIOMAS', '', 'Ingresos'),
    ('CONGRESOS, SEMINARIOS, TALLERES Y OTROS - FUTURE', '', 'Ingresos'),
    ('CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PEAS', '', 'Ingresos'),
    ('CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO', '', 'Ingresos'),
    ('MATRICULA FUTURE', '', 'Ingresos'),
    ('MATRICULA POSGRADO', '', 'Ingresos'),
    ('MATRICULA PREGRADO', '', 'Ingresos'),
    ('MATRICULA SEGUNDA ESPECIALIDAD', '', 'Ingresos'),
    ('MERCHANDISING', '', 'Ingresos'),
    ('OTROS INGRESOS DE GESTIÓN', '', 'Ingresos'),
    ('PENSION CENTRO DE IDIOMAS', '', 'Ingresos'),
    ('PENSION FUTURE', '', 'Ingresos'),
    ('PENSION PEA', '', 'Ingresos'),
    ('PENSION POSGRADO', '', 'Ingresos'),
    ('PENSION PREGRADO', '', 'Ingresos'),
    ('PENSION SEGUNDA ESPECIALIDAD', '', 'Ingresos'),
    ('TASAS CARPETAS CENTRO DE IDIOMAS', '', 'Ingresos'),
    ('TASAS CARPETAS FUTURE', '', 'Ingresos'),
    ('TASAS CARPETAS PEA', '', 'Ingresos'),
    ('TASAS CARPETAS POSGRADO', '', 'Ingresos'),
    ('TASAS CARPETAS PREGRADO', '', 'Ingresos'),
    ('TASAS CARPETA SEGUNDA ESPECIALIDAD', '', 'Ingresos'),
    ('Materiales', '', 'Ingresos'),
]



# Partidas GG oficiales que no tenían una cuenta NIIF con el mismo nombre: cuenta asignada en la consolidación.
GG_CUENTA_MANUAL = {
    ("Alquiler maquinaria y equipos", "Egresos"): "63532",
    ("Materiales e insumos quimicos", "Egresos"): "6941011400",
    ("Modulos academicos", "Egresos"): "6941011000",
}


def _tipo_gg(tipo: str) -> str:
    return tipo if tipo in ("Egresos", "Ingresos") else "Inversion"


def consolidar_niif_gg(conn) -> None:
    """Base única «Cuentas NIIF + Partidas GG»: cada cuenta NIIF lleva su Partida GG y la descripción orientativa."""
    from controllers import norm_key
    rows = conn.execute("SELECT id, codigo_niif, partida_rubro, partida_gg, descripcion, tipo_flujo "
                        "FROM catalogo_niif_partidas ORDER BY id").fetchall()
    gg_oficial = {}
    for nombre, descripcion, tipo in PARTIDAS_GG:
        gg_oficial.setdefault((norm_key(nombre), _tipo_gg(tipo)), (nombre, descripcion))
    # 1) cuentas sin GG cuyo nombre coincide con una Partida GG oficial
    for rid, code, cuenta, gg, desc, tipo in rows:
        if not gg and (norm_key(cuenta), tipo) in gg_oficial:
            conn.execute("UPDATE catalogo_niif_partidas SET partida_gg=? WHERE id=?",
                         (gg_oficial[(norm_key(cuenta), tipo)][0], rid))
    # 2) asignaciones manuales
    for (nombre, tipo), code in GG_CUENTA_MANUAL.items():
        conn.execute("UPDATE catalogo_niif_partidas SET partida_gg=? WHERE codigo_niif=? AND tipo_flujo=? AND "
                     "(partida_gg IS NULL OR partida_gg='')", (nombre, code, tipo))
    # 3) descripciones orientativas de las Partidas GG
    for (key, tipo), (nombre, descripcion) in gg_oficial.items():
        if descripcion:
            conn.execute("UPDATE catalogo_niif_partidas SET descripcion=? WHERE partida_gg=? AND tipo_flujo=? AND "
                         "(descripcion IS NULL OR descripcion='')", (descripcion, nombre, tipo))


def _migrar_bd_existente(db) -> None:
    """Migraciones de versiones anteriores (se ejecutan una sola vez)."""
    conn = db.conn
    if db.meta("version_catalogos") == "2":
        return
    from controllers import norm_key
    # a) la tabla separada de Partidas GG se integra en la base consolidada
    if db.table_exists("catalogo_partidas_gg"):
        for nombre, descripcion, tipo, codigo in conn.execute(
                "SELECT partida_gg, descripcion, tipo, codigo_niif FROM catalogo_partidas_gg").fetchall():
            tipo = _tipo_gg(tipo)
            hit = conn.execute("SELECT id FROM catalogo_niif_partidas WHERE partida_gg=? AND tipo_flujo=?",
                               (nombre, tipo)).fetchone()
            if hit:
                if descripcion:
                    conn.execute("UPDATE catalogo_niif_partidas SET descripcion=? WHERE partida_gg=? AND tipo_flujo=? "
                                 "AND (descripcion IS NULL OR descripcion='')", (descripcion, nombre, tipo))
                continue
            conn.execute("INSERT OR IGNORE INTO catalogo_niif_partidas (codigo_niif, partida_rubro, partida_gg, "
                         "descripcion, categoria, tipo_flujo) VALUES (?,?,?,?,?,?)",
                         (codigo or "SIN-CODIGO", nombre, nombre, descripcion or "", "", tipo))
        conn.execute("DROP TABLE catalogo_partidas_gg")
    conn.execute("UPDATE catalogo_niif_partidas SET partida_gg='CONGRESOS, SEMINARIOS, TALLERES Y OTROS - POSGRADO' "
                 "WHERE codigo_niif='7041040700' AND partida_gg='CONGRESOS, SEMINARIOS, TALLERES Y OTROS - PREGRADO'")
    # b) las Acciones estratégicas (AE) ya no se utilizan: el eje estratégico cumple ese rol
    conn.execute("DELETE FROM catalogos_estrategicos WHERE tipo_catalogo='AE'")
    conn.execute("UPDATE iniciativas SET ae=''")
    # c) objetivos anuales sin eje: vínculo provisional
    for codigo, d, oe, eje in OBJETIVOS_POA:
        conn.execute("UPDATE catalogos_estrategicos SET eje=? WHERE tipo_catalogo='OBJ' AND codigo=? AND "
                     "(eje IS NULL OR eje='')", (eje, codigo))
    consolidar_niif_gg(conn)
    # d) el detalle usa una sola «Partida» (Partida GG)
    gg_por_cuenta, cat_por_gg = {}, {}
    for cuenta, gg, tipo, cat in conn.execute("SELECT partida_rubro, partida_gg, tipo_flujo, categoria FROM "
                                              "catalogo_niif_partidas WHERE partida_gg<>'' ORDER BY id").fetchall():
        gg_por_cuenta.setdefault((norm_key(cuenta), tipo), gg)
        cat_por_gg.setdefault((norm_key(gg), tipo), cat)
    for t in ("detalle_iniciativas", "detalle_otras_operativas", "detalle_act_prop", "detalle_inversiones"):
        for rid, partida, gg, tipo, cat in conn.execute(
                f"SELECT id, partida, partida_gg, tipo, categoria FROM {t}").fetchall():
            nuevo = gg or gg_por_cuenta.get((norm_key(partida), tipo)) or partida
            nueva_cat = cat_por_gg.get((norm_key(nuevo), tipo)) or cat
            if nuevo != partida or nuevo != gg or nueva_cat != cat:
                conn.execute(f"UPDATE {t} SET partida=?, partida_gg=?, categoria=? WHERE id=?",
                             (nuevo, nuevo, nueva_cat, rid))
    conn.commit()
    db.set_meta("version_catalogos", "2")


# Tipo de actividad comercial: combo del Registro Maestro de «Otras operativas» de las unidades del macrogrupo
# Marketing. Se siembra una sola vez; luego la DPDI lo mantiene en Administración → Tipos de actividad comercial.
ACTIVIDADES_COMERCIALES = [
    "ATL", "BRANDING", "BTL", "CAMPAÑAS", "EVENTOS", "INVESTIGACIÓN", "LICENCIAS", "PRODUCCIÓN", "SERVICIOS",
    "VENTAS AREQUIPA", "VENTAS JULIACA", "VENTAS AYACUCHO", "VENTAS CUZCO", "VENTAS CAJAMARCA",
    "VENTAS EDUCACIÓN CONTINUA", "VENTAS MAESTRÍAS", "VENTAS PREGRADO", "VENTAS SEGUNDAS ESPECIALIDADES",
]


def _sembrar_actividades_comerciales(db) -> None:
    if db.meta("actividades_comerciales") == "1":
        return
    if not db.scalar("SELECT COUNT(*) FROM catalogos_estrategicos WHERE tipo_catalogo='ACT_COM'"):
        db.conn.executemany(
            "INSERT OR IGNORE INTO catalogos_estrategicos (tipo_catalogo, codigo, descripcion, padre, eje, extra, orden) "
            "VALUES ('ACT_COM', ?, ?, '', '', '', ?)", [(x, x, i) for i, x in enumerate(ACTIVIDADES_COMERCIALES)])
        db.conn.commit()
    db.set_meta("actividades_comerciales", "1")


def seed_database(db) -> None:
    """Pre-puebla (idempotente) catálogos, parámetros y unidades iniciales."""
    conn = db.conn
    db.parametros()
    nueva = not db.scalar("SELECT COUNT(*) FROM catalogo_niif_partidas")

    conn.executemany(
        "INSERT OR IGNORE INTO catalogo_niif_partidas "
        "(codigo_niif, partida_rubro, partida_gg, categoria, tipo_flujo, naturaleza, resultado_funcion) "
        "VALUES (?,?,?,?,?,?,?)",
        NIIF_PARTIDAS,
    )
    if nueva:
        consolidar_niif_gg(conn)

    def cat(tipo_catalogo, rows):
        conn.executemany(
            "INSERT OR IGNORE INTO catalogos_estrategicos (tipo_catalogo, codigo, descripcion, padre, eje, extra, orden) "
            "VALUES (?,?,?,?,?,?,?)",
            [(tipo_catalogo, *r, i) for i, r in enumerate(rows)],
        )

    cat("EJE", [(c, d, "", "", a) for c, d, a in EJES])
    cat("OE", [(c, d, "", "", ind) for c, d, ind in OBJETIVOS_PEI])
    cat("OBJ", [(c, d, oe, eje, "") for c, d, oe, eje in OBJETIVOS_POA])
    cat("PRIORIDAD", [(c, d, "", "", crit) for c, d, crit in PRIORIDADES])
    cat("ORIGEN", [(o, o, "", "", "") for o in ORIGENES])
    cat("SUSTENTO_EGR", [(x, x, "", "", "") for x in SUSTENTOS_EGRESOS])
    cat("SUSTENTO_ING", [(x, x, "", "", "") for x in SUSTENTOS_INGRESOS])
    cat("CAT_OO", [(c, d, "", "", inc) for c, d, inc in CATEGORIAS_OO])
    cat("TIPO_AP", [(t, t, "", "", d) for t, d in TIPOS_AP])
    cat("TIPO_INV", [(t, t, "", "", d) for t, d in TIPOS_INVERSION])
    cat("TIPO_ACTIVIDAD", [(t, d, "", "", h) for t, d, h in TIPOS_ACTIVIDAD])

    conn.executemany(
        "INSERT OR IGNORE INTO costos_referencia (concepto, unidad_medida, costo_unitario, fuente) VALUES (?,?,?,?)",
        COSTOS_REFERENCIA,
    )

    if not db.scalar("SELECT COUNT(*) FROM unidades_organicas"):
        for u in UNIDADES_INICIALES:
            db.insert("unidades_organicas", u, commit=False)

    if not db.scalar("SELECT COUNT(*) FROM plantilla_funcionamiento WHERE unidad_id IS NULL"):
        conn.executemany(
            "INSERT INTO plantilla_funcionamiento (macrogrupo, solo_genera_ingresos, tipo_flujo, partida, base_calculo, "
            "porcentaje, monto_fijo, aplica_ajuste, sustento, orden) VALUES (?,?,?,?,?,?,?,?,?,?)",
            [(*r, i) for i, r in enumerate(PLANTILLA_FUNCIONAMIENTO)],
        )
    conn.commit()
    _migrar_bd_existente(db)
    _sembrar_actividades_comerciales(db)
