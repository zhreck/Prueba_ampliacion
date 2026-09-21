# Prueba_ampliacion

Web app para ampliar materiales (Modelos, Repuestos, ...) a los centros SAP
correspondientes, replicando la lógica de negocio de la vista SQL
`VW_ZRP1_TAU_AMPLIADO_DELTA` (MASTERDATA), pero sin depender de conexión
directa a la base de datos.

## Cómo funciona

1. El usuario sube un Excel con **una o varias filas de input** (hoja `INPUT`).
   No llena `CENTRO`, `COD_CEBE` ni `NUMERO MATERIAL` — el sistema los calcula.
2. El motor busca el `FABRICANTE CODIGO` del input en la tabla de referencia
   (`data/reference/modelos_centros_ampliacion.xlsx`). Si no lo encuentra,
   usa el fabricante comodín `F9999` (TODAS/OTRAS MARCAS).
3. Genera **una fila de salida por cada centro** que le corresponda a ese
   fabricante, agregando `CENTRO` y `COD_CEBE`.
4. Asigna un **número de material** nuevo (correlativo global 1–5.000.000,
   guardado en SQLite) — el mismo número se repite en todas las filas
   ampliadas de ese material, porque es el mismo material en distintos centros.
5. Devuelve el Excel resultante para descargar.

## Correr localmente

```bash
cd app
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python3 app.py
```

Abrir `http://localhost:5000`.

La base `data/db/correlativos.db` se crea sola la primera vez que se procesa
un archivo. En `/historial` se ve el último número usado por rango y el log
de asignaciones (útil para auditar quién generó qué número y cuándo).

## Cómo agregar un tipo nuevo (ej. Repuestos)

No hay que tocar `engine.py` ni `app.py`. Solo:

1. Crear `app/config/tipos/repuestos.json` copiando la estructura de
   `modelos.json`, con:
   - `input_columns`: columnas del Excel de input de Repuestos.
   - `reference_file`: ruta al Excel de referencia de Repuestos.
   - `key_input` / `key_reference`: columnas por las que se hace el match
     (ej. fabricante, o lo que corresponda).
   - `fallback_value`: código comodín si aplica.
   - `output_columns_from_reference`: columnas a copiar desde la referencia.
   - `correlativo.nombre`: si Repuestos debe compartir el mismo rango
     global que Modelos, usar `"material_global"`. Si necesita su propio
     rango independiente, poner otro nombre (ej. `"material_repuestos"`)
     y su propio `range_min` / `range_max`.
2. Dejar el Excel de referencia de Repuestos en `app/data/reference/`.

El tipo nuevo aparece automáticamente en el selector de la página principal.

## Estructura

```
app/
  app.py            # rutas Flask (/, /procesar, /historial)
  engine.py         # motor genérico de ampliación (lee config + hace el match)
  correlativo.py     # manejo del correlativo en SQLite
  config/tipos/      # un JSON por tipo de ampliación
  data/reference/    # tablas de referencia (Excel) por tipo
  data/db/            # SQLite del correlativo (se genera solo)
  templates/, static/ # UI
```

## Pendiente / ideas a futuro

- Módulo de Repuestos (input distinto, cebe distinto — usa el mismo motor).
- Descarga de plantilla vacía de input desde la propia web.
- Autenticación básica si se despliega fuera de la red interna.

## ¿Se puede leer este código sin saber programar?

**No, no en general.** Los archivos `.py` (`engine.py`, `salida_sap.py`,
`app.py`, `correlativo.py`) son código Python real — funciones, bucles,
manejo de tablas (pandas) — y entenderlos requiere saber programar, sin
importar cuán bien nombradas estén las variables o cuántos comentarios
tengan. Eso es así con cualquier código, no algo puntual de este proyecto.

**Con una excepción real: los archivos de configuración en `config/*.json`.**
Ahí vive casi toda la lógica de negocio (qué campo SAP sale de qué columna,
qué categoría de valoración le corresponde a cada marca, qué rango de
material usa cada tipo, etc.), y cada decisión viene acompañada de una nota
en español explicando de dónde salió el dato, quién lo confirmó y por qué
("_nota", "_fuente", "notas_pendientes"). Alguien sin programar no puede
*modificar* esos archivos con confianza (un JSON mal cerrado rompe todo),
pero sí puede **leerlos y entender las reglas de negocio** casi como si
fueran una minuta — de hecho estas notas se escribieron pensando en que
Seba (que no programa) las pueda revisar directamente.
