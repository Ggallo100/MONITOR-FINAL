# Sistema de Formulación Presupuestal y Operativa POA - DPDI

Proceso Gestión Estratégica **GD-DO-01** · Insumo **PP-PR-02**

Aplicación de escritorio (Python 3.11+, CustomTkinter, SQLite3, Pandas, OpenPyXL y ReportLab) que reproduce el formato
oficial `Ficha_Formulacion_POA.xlsx`: estructura maestro-detalle, catálogos NIIF, reglas de negocio, las 12 secciones
de la Ficha_Resumen y las 11 directrices automáticas de la DPDI. Se puede empaquetar como `Sistema_POA_DPDI.exe`.

## Archivos

| Archivo | Contenido |
|---|---|
| `database.py` | Esquema SQLite (unidades y techos, parámetros, catálogo NIIF, partidas GG, catálogos estratégicos, costos de referencia, plantilla de funcionamiento, maestros/detalles de las 5 clases de actividad, bitácora) y CRUD genérico. |
| `catalogs_seed.py` | Catálogos oficiales embebidos: base consolidada de 296 cuentas NIIF con sus 99 Partidas GG y descripciones, ejes EJ-01…EJ-09, OE y objetivos anuales POA (vinculados a eje y OE), orígenes, prioridades, sustentos, categorías OO-01…OO-11, tipos de actividad propia e inversión, 9 + 28 categorías de la Ficha_Resumen y plantilla base de funcionamiento. Incluye la migración automática de bases de versiones anteriores. |
| `controllers.py` | Reglas de negocio: cálculos por fila, generador de códigos, «Correr Actividades de Funcionamiento», Ficha_Resumen (12 secciones), 11 directrices, BD NIIF mensualizada, importador de fichas y parámetros, carga masiva NIIF (Pandas). |
| `exporter.py` | Libro oficial `.xlsx` (16 pestañas con fórmulas vivas, listas desplegables —incluida la dependiente Eje → Objetivo anual—, nombres definidos y formato condicional) + `BD_CONSOLIDADA_NIIF`; BD consolidada multiunidad (`.xlsx` + `.csv` + pivot mensual); PDF de la Ficha Resumen con bloque de firmas; **plantilla descargable de Parámetros y Catálogos**. |
| `main_app.py` | Interfaz CustomTkinter (login por rol y pestañas 0–7). |
| `build_exe.bat` | Crea el entorno, ejecuta las pruebas y genera `dist\Sistema_POA_DPDI.exe`. |
| `tests/test_core.py` | Pruebas del núcleo (cálculos, directrices, códigos, plantilla, exportar→importar ida y vuelta, PDF). |

## Ejecución

```bash
pip install -r requirements.txt
python main_app.py            # opcional: python main_app.py ruta\a\otra_base.sqlite3
```

La base se crea al primer arranque en `sistema_poa_dpdi/data/` (en el `.exe`: `%APPDATA%\Sistema_POA_DPDI\`).

**Roles**

* **Responsable de Unidad Orgánica**: elige su unidad y formula su POA (pestañas 0 y 2–7).
* **Administrador DPDI**: clave inicial `DPDI` (cámbiela en *Parámetros → Datos generales*). Acceso total al mantenedor
  de parámetros, techos, cuentas NIIF, catálogos, carga masiva y cambio de unidad.

## Plantilla de Parámetros y Catálogos (actualización masiva)

En *1. Parámetros y catálogos NIIF* (rol Administrador DPDI):

1. **⬇ Descargar plantilla de Parámetros y Catálogos (.xlsx)**: genera un Excel con los valores vigentes en las hojas
   `Datos_Generales`, `Unidades_Techos`, `Cuentas_NIIF_Partidas_GG`, `Catalogos_Estrategicos`, `Costos_Referencia`,
   `Plantilla_Funcionamiento` (más `Instrucciones` y `Listas` para las listas desplegables).
2. Edite o agregue filas (amarillo = editable; gris = calculado). La columna **ID** identifica cada registro: déjela
   vacía para crear uno nuevo.
3. **⬆ Importar plantilla de Parámetros y Catálogos (.xlsx)**: actualiza o agrega registros (no elimina; para retirar
   una cuenta use `Activo = No`). Volver a importar el mismo archivo no duplica registros.

## Alineamiento de las iniciativas: Eje → Objetivo anual del POA → OE del PEI

* Cada objetivo anual del POA (OBJ) se registra con el **Eje estratégico** al que pertenece y el **OE del PEI** al que
  contribuye (mantenedor *Catálogos estratégicos* o plantilla, hoja `Catalogos_Estrategicos`).
* En la iniciativa el usuario elige primero el **eje**; el combo **Objetivo anual del POA** solo muestra los objetivos
  de ese eje, y el **OE del PEI se completa automáticamente**. En el libro Excel, la columna G usa una lista dependiente
  del eje (columna B) y la columna H calcula el OE.
* Las **Acciones estratégicas del PEI (AE)** ya no se usan: el eje estratégico cumple ese rol.
* **Unidad de medida** del indicador: lista de dos opciones, `Unidad` o `Porcentaje (%)`.

## Partida única en el detalle (base consolidada Cuentas NIIF + Partidas GG)

* El mantenedor maneja una sola base: cada cuenta contable NIIF lleva su **Partida GG** y una **descripción
  orientativa**.
* En el detalle de iniciativas, otras operativas, actividades propias e inversión hay un solo combo **«Partida»** con las
  Partidas GG; al elegirla se muestra su descripción orientativa (también en el buscador 🔍), y la categoría y el código
  NIIF se derivan de la base consolidada.
* En el libro Excel: columna **Partida** con lista según el tipo (Ingresos / Egresos / Inversión) y una columna de
  **descripción orientativa** calculada.
* Las hojas de Funcionamiento siguen usando las cuentas NIIF (secciones I y J del formato oficial).

## Empaquetado (.exe sin consola)

```bat
build_exe.bat
```

Comando PyInstaller exacto que ejecuta el script:

```bat
pyinstaller --noconfirm --clean --onefile --windowed --name Sistema_POA_DPDI --collect-all customtkinter --collect-data reportlab --hidden-import openpyxl --hidden-import pandas main_app.py
```

### Si `build_exe.bat` muestra «Error durante la construcción»

* `PermissionError: Acceso denegado … dist\Sistema_POA_DPDI.exe`: el sistema estaba abierto o el antivirus bloqueaba el
  ejecutable anterior. La versión actual del `.bat` cierra el programa y reemplaza el `.exe` automáticamente; si aun así
  falla, cierre el sistema y vuelva a ejecutarlo.
* La advertencia de Windows «No se pudo comprobar el editor» aparece porque el `.bat` no tiene firma digital: pulse
  **Ejecutar**.
* Los datos no se pierden al reconstruir: la base está en `%APPDATA%\Sistema_POA_DPDI\poa_dpdi.sqlite3`.

## Reglas de cálculo replicadas

* **Costo/Tarifa (S/)** `=IF(F="","",IF(E="US$",F*TC,F))`; **Total** `= Cantidad × Costo S/`.
* **Detalles** (iniciativas, otras operativas, act. propias, inversión): `Programado = Σ Ene..Dic`;
  `Control = OK si |Programado − Total| < 0.01`, si no `Revisar`.
* **Funcionamiento_Ing**: ingreso mensual neto = `Cant × Tarifa S/ × factor_mes × (1 − Becas) × (1 − Morosidad si aplica)`;
  «x» = factor 1; bruto = `Cant × Tarifa × Σ factores`; descuentos = bruto − neto.
* **Funcionamiento_Egr**: monto del mes = `factor_mes × Coste total × (1 + Reajuste si aplica y mes ≥ MesReajuste)`.
* **Iniciativas**: `Control de meta = OK si T1+T2+T3 = Meta anual`; ingresos/gasto desde el detalle; resultado y margen.
* **Actividades propias**: resultado, margen, `Punto de equilibrio = ROUNDUP(Costo total / Tarifa neta)` y estado
  `Autofinanciada` / `Déficit`.
* **Techos**: `Dentro del techo` / `Excede techo` / `Sin techo`; las actividades propias se excluyen del techo OPEX.
* **Códigos**: `INI-NN`, `FI/FE/OO/AP/INV-NNN` correlativos por unidad y año; en la BD consolidada se prefijan con el área
  (`OL-INI-01`).
* **Correr Actividades de Funcionamiento**: inserta la plantilla base (planilla administrativa y docente distribuidas
  como 11 sueldos + vacaciones + 2.18 gratificaciones + CTS 14/12 + EsSalud 9 %, marketing, depreciación, servicios
  básicos, provisión de cobranza dudosa y pensiones según el macrogrupo) con factor `1/12 = 0.08333333` en Ene–Dic.
  La plantilla es editable (global o por unidad) en *Parámetros → Plantilla de funcionamiento*.

## Verificación realizada

* Las dos fichas oficiales de referencia se importaron y la sección 1 de la Ficha_Resumen coincide al céntimo con los
  valores calculados por Excel (unidad OL: 11 iniciativas, 44 ítems, 23 egresos de funcionamiento, etc.).
* El libro exportado se recalculó con LibreOffice: las 576 cifras comparadas por unidad coinciden con el motor Python y
  no hay celdas con error en ninguna de las 17 hojas.
* Exportar → importar devuelve los mismos totales y directrices (prueba automatizada).

## Diferencias deliberadas respecto de la plantilla Excel

Durante la réplica se encontraron fórmulas inconsistentes en la plantilla; el sistema aplica la regla que describe el
propio formato:

1. **DP-04 y DI-01** cuentan en Excel solo las marcas «x» (`COUNTIF(J:U,"x")`), por lo que marcan como observación toda
   fila prorrateada con 0.08333333. El sistema considera mes marcado tanto «x» como un factor > 0.
2. **Total CAPEX – % de uso**: Excel calcula `(D23−D22)/F23`, que siempre da 0 %; el sistema usa `D23/F23`.
3. **Sección 10**: el total de ingresos de Excel omite los ingresos de otras operativas y el control compara egresos con
   inversión contra el total OPEX; el sistema incluye ambos conceptos.
4. **Sección 4**: la columna «% del total costos» repetía el monto; ahora es porcentaje y se agrega el resultado.
5. **Sección 3**: el estado de la meta de otras operativas comparaba celdas equivocadas; el resultado de actividades
   propias se mide como ingresos − costos (excedente), según la definición de la meta.
6. **Control de distribución de Funcionamiento_Egr**: verifica que los factores sumen 100 % (o que haya meses con «x»),
   para no marcar «Revisar» cuando hay reajuste.
7. En la ficha de Pregrado, la fila «SUELDOS ADMINISTRATIVOS» no tenía la fórmula de Costo unitario (S/), por lo que
   Excel omitía S/ 15,400; el sistema siempre la calcula. La celda del techo de inversión de esa Ficha_Resumen también
   había sido sobrescrita (61,512.50 frente a 10,000 en Parámetros); el sistema usa Parámetros.

## Pruebas

```bash
python -m unittest discover -s tests -v
```
