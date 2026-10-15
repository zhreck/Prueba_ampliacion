# Prueba_ampliacion

Web app (Flask) para **ampliar materiales a los centros SAP** y generar la planilla
de carga **MM01** (151 columnas) de **Modelos** (Vehículos, Camiones y Maquinarias)
y de **Repuestos**, sin depender de una conexión directa a la base de datos. Replica
la lógica de negocio de la vista SQL `VW_ZRP1_TAU_AMPLIADO_DELTA` (MASTERDATA).

Sitio publicado: https://zhreck.pythonanywhere.com/

## Qué hace

1. El usuario elige **Modelos** o **Repuestos**, descarga la plantilla de input
   (hoja `INPUT`, con listas desplegables y fórmulas de texto breve/NPF) y la sube llena.
2. El motor traduce los nombres del input a códigos (filial, marca, fabricante,
   grupo de artículo, jerarquía), busca en las tablas de referencia a qué **centros**
   corresponde cada material (con su **centro de beneficio** y categoría de valoración)
   y genera **una fila por centro** (y por canal de distribución).
3. Asigna un **número de material (correlativo)** por rango y tipo de material; el
   mismo número se repite en todas las filas ampliadas de ese material.
4. Con la opción *"Generar formato de carga SAP (MM01)"* convierte el resultado a las
   151 columnas del formato real de SAP. Si un dato no se puede completar, la celda
   queda en blanco y se avisa al procesar.
5. Se descarga el Excel. Además, la planilla SAP queda guardada **48 horas** y se
   puede volver a ver y descargar desde el **Historial**.

### Tipos de material y rangos de correlativo

| Tipo | Qué es | Filiales | Rango |
|---|---|---|---|
| ZVEH | Vehículos y motos | VF00 | 10.000.000 – 19.999.999 |
| ZMAQ | Maquinarias | VC00, VD00, VE00 | 20.000.000 – 29.999.999 |
| ZCAM | Camiones | VA00 | 30.000.000 – 39.999.999 |
| ZUSA | Unidades usadas (aún no habilitado) | — | 40.000.000 – 49.999.999 |
| ZRP1 / ZRP2 / ZRP3 | Repuestos (normal / lote / serie), comparten rango | todas | 50.000.000 – 79.999.999 |

Los rangos están en `correlativo.py` (`DEFAULT_RANGOS`).

## Reglas de negocio principales

- **Varias filiales en un mismo Excel**: todo sale junto en una sola hoja (`ZMAQ`, o
  `MODELOS` si el archivo mezcla tipos de material). Cada fila usa los centros,
  canales y grupos de su filial.
- **'Tattersall Automotriz y Maco Tattersall'** (Vehículos y Camiones) amplía a VF00 y
  VA00, cada una con sus centros, su tipo de material (ZVEH/ZCAM) y su número.
- **Maquinarias**: MAXIMAL (180) desde VC00 se amplía también a VE00 y VE01; LIUGONG
  (150, F0016) se amplía a VC00 (sin VC08) y VD00 desde cualquiera de las dos
  sociedades. Reglas en `config/tipos/modelos_maquinaria.json`.
- **Centro de beneficio de electromovilidad** (…02) si el NPF termina en `BEV`,
  `PHEV` o `MHEV`; en otro caso el de unidades nuevas (…01).
- **NPF único por tipo de material**: se rechaza el archivo si un NPF se repite dentro
  del Excel o si ya tiene número asignado en el historial para ese tipo.
- **Canal de distribución**: ZVEH y ZCAM 20 y 30; ZMAQ solo 30; Repuestos según filial.
- **Grupo de compras**: Maquinarias por marca (`Q19` si la marca no está en la tabla);
  Camiones por marca (`M28` por defecto); Vehículos fijo (`T03`); Repuestos por filial.
- **Jerarquías** sin relleno de ceros en los niveles vacíos (códigos de 5, 10 o 18
  dígitos, igual que el diccionario SAP).
- **Categoría Clase** (columna BK) = `300` en Modelos.
- **Disponibilidad centro × fabricante** (`Centros_UN_RP.xlsx`, hoja `Unidades`): un 0
  bloquea ese fabricante en ese centro (salvo las reglas de Maximal/Liugong de arriba).

## Historial de correlativos (`/historial`)

- Tarjetas con el último número y los disponibles de cada rango.
- Tabla con todos los materiales: número, tipo, **filiales**, NPF, fabricante, texto
  breve, fecha (DD-MM-AAAA) y hora. Búsqueda por texto/filial, **filtro por fecha de
  creación** (desde/hasta), paginación y **descarga en Excel** respetando el filtro.
- Al hacer clic en un número se abre la página del material, con **Vista previa** y
  **Descargar planilla** de la planilla en que se generó (se conserva 48 horas).
- Los correlativos que ya existen en PRD se cargan desde
  `data/reference/Repuestos/CorrelativosRepuestos.xlsx` y
  `data/reference/Unidades/RangoModelos.xlsx` (si el segundo cambia, se vuelve a
  importar solo en el siguiente arranque). Los materiales ocultos (cargas mal hechas)
  conservan su número pero no se muestran.
- Las fechas se guardan en **hora de Chile** (`America/Santiago`), no en la del servidor.

## Base de datos

**SQLite**, un archivo en `data/db/correlativos.db` (se crea solo; no va a git, así que
local y producción tienen datos independientes). Guarda contadores, historial,
planillas generadas y migraciones ya aplicadas. Las planillas se guardan como archivos
en `data/generaciones/` (tampoco va a git) y se borran solas a las 48 horas.

## Ajustes de contador y migraciones al arrancar

Al arrancar, `app.py` ejecuta (una sola vez cada una, quedan registradas en la tabla
`migraciones`): corrección de fechas a hora de Chile, importación de los correlativos
de PRD, limpieza de planillas vencidas y los ajustes de
`config/ajustes_contadores.json`. En ese archivo:

- `siguiente`: fija el próximo número (solo mueve el contador hacia adelante).
- `avanzar`: salta N números desde donde esté el contador.
- `reiniciar`: fija el contador aunque esté más adelante y oculta los registros desde
  ese número (para volver a empezar, ej. una capacitación).

Para volver a mover un contador hay que agregar una entrada con un `id` **nuevo**.

## Correr localmente

```bash
python -m venv .venv
.venv\Scripts\activate          # Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Abrir `http://localhost:5000`.

## Publicación (PythonAnywhere)

El servidor se actualiza solo con cada `push` a GitHub: un webhook llama a
`/deploy-hook`, que hace `git pull` y recarga la app (requiere la variable de entorno
`GITHUB_WEBHOOK_SECRET` en el servidor). Al subir el código, las migraciones y ajustes
de arriba se aplican en el primer arranque. La base y las planillas guardadas viven
solo en el servidor.

## Estructura

```
app.py                  # rutas Flask: selección, carga, resultado, historial, deploy-hook
engine.py               # motor de ampliación (lee config, traduce nombres, hace el match)
salida_sap.py           # conversión a las 151 columnas SAP (MM01) por tipo de material
correlativo.py          # SQLite: contadores, historial, planillas, importaciones y ajustes
config/tipos/           # un JSON por tipo de input (Modelos: Vehículos/Camiones, Maquinarias; Repuestos)
config/salida_sap/      # un JSON por tipo de material SAP (zveh, zcam, zmaq, zusa, zrp1/2/3)
config/*.json           # marcas, categoría de valoración, grupo de compras, ajustes de contador
data/reference/         # tablas de referencia y planillas de input (Excel)
data/db/                # SQLite (se genera solo, no va a git)
data/generaciones/      # planillas generadas guardadas 48 h (no va a git)
docs/                   # matriz de confirmación de campos, header de 151 columnas, ejemplos
templates/, static/     # interfaz
```

Archivos de Excel clave en `data/reference/`:

- `DiccionarioSAP.xlsx`: centros y centro de beneficio por filial/marca/fabricante
  (hoja `cebe3.0`) y jerarquías SAP (hoja `jerarquia`).
- `Unidades/plantilla_modelos_ZVEH_ZCAM.xlsx` y `Unidades/plantilla_modelos_maquinarias.xlsx`:
  plantillas de input de Modelos (sus hojas de diccionarios alimentan las listas
  desplegables y la traducción de nombres a códigos).
- `Unidades/Centros_UN_RP.xlsx`: matriz de disponibilidad centro × fabricante.
- `Unidades/RangoModelos.xlsx` y `Repuestos/CorrelativosRepuestos.xlsx`: correlativos
  ya creados en PRD.

> Si cambias una plantilla de input, mantén los nombres de las hojas y de las columnas
> del `INPUT`, y revisa que los rangos de las listas desplegables en
> `config/tipos/*.json` (`validaciones_lista_extra`) sigan apuntando a las filas correctas.

## Cómo agregar un tipo nuevo

La mayor parte es configuración: copiar la estructura de un JSON de
`config/tipos/` (columnas del input, archivo de referencia, columnas de match,
traducciones de nombres a códigos, validaciones) y, si se genera planilla SAP, agregar
su plantilla en `config/salida_sap/` y el rango en `correlativo.py`. El tipo aparece solo
en el selector de la página principal.

## Pendiente / ideas a futuro

- ZUSA (unidades usadas): sin input ni plantilla todavía.
- Rangos de correlativo reales de Modelos: hoy los contadores quedan en el último número de `RangoModelos.xlsx` (puede cambiar hasta el 02/11).
- Autenticación básica si se despliega fuera de la red interna.
- Si crece el uso simultáneo, evaluar pasar de SQLite a una base con servidor.

## ¿Se puede leer este código sin saber programar?

**No en general** para los archivos `.py` (`engine.py`, `salida_sap.py`, `app.py`,
`correlativo.py`): es código Python real y entenderlo requiere saber programar.

**La excepción son los archivos `config/*.json`.** Ahí vive casi toda la lógica de
negocio (qué campo SAP sale de qué columna, qué categoría de valoración o grupo de
compras le corresponde a cada marca, qué reglas especiales hay por marca o filial) y
cada decisión viene con una nota en español explicando de dónde salió el dato y quién
lo confirmó (`_nota`, `_fuente`, `notas_pendientes`). Quien no programa puede **leerlos
como una minuta** para revisar las reglas, aunque no conviene que los modifique sin
ayuda: un JSON mal cerrado rompe la app.
