"""
Manejo del correlativo de "Número de material".

Soporta múltiples rangos con nombres distintos:
- "material_global" (1 - 5.000.000): Original para ampliación sin SAP
- "zmaq_material" (20000000 - 29999999): Maquinarias SAP
- "zcam_material" (30000000 - 39999999): Camiones SAP
- "zveh_material" (10000000 - 19999999): Vehículos y Motos SAP
- "repuestos_material" (50000000 - 79999999): Repuestos SAP (ZRP1/ZRP2/ZRP3,
  comparten rango — son el mismo universo de materiales, solo cambia si
  llevan serie, lote o ninguno). Rango CONFIRMADO por Seba.

Cada tipo de material puede tener su propio correlativo nombrado.
"""

import json
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "db" / "correlativos.db"
RANGOS_CONFIG_PATH = Path(__file__).parent / "config" / "rangos_materiales.json"

_lock = threading.Lock()

# Rangos default - estos están disponibles siempre
DEFAULT_RANGOS = {
    "material_global": {"min": 1, "max": 5000000, "descripcion": "Ampliación original"},
    "zmaq_material": {"min": 20000000, "max": 29999999, "descripcion": "ZMAQ - Maquinarias (VC00, VD00, VE00)"},
    "zcam_material": {"min": 30000000, "max": 39999999, "descripcion": "ZCAM - Camiones (VA00)"},
    "zveh_material": {"min": 10000000, "max": 19999999, "descripcion": "ZVEH - Vehículos y Motos (VF00)"},
    "zusa_material": {"min": 40000000, "max": 49999999, "descripcion": "ZUSA - Unidades usados"},
    "repuestos_material": {"min": 50000000, "max": 79999999, "descripcion": "ZRP1/ZRP2/ZRP3 - Repuestos (confirmado por Seba)"},
}


def _get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS contadores (
            nombre TEXT PRIMARY KEY,
            ultimo_valor INTEGER NOT NULL,
            range_min INTEGER NOT NULL,
            range_max INTEGER NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS historial_asignaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_contador TEXT NOT NULL,
            numero_asignado INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            texto_breve TEXT,
            fabricante_codigo TEXT,
            fecha TEXT NOT NULL
        )
    """)
    # Columnas agregadas después (el historial ahora guarda qué se cargó, para
    # poder sacar un mini reporte): se agregan a bases ya existentes.
    existentes = {r[1] for r in conn.execute("PRAGMA table_info(historial_asignaciones)")}
    for columna, definicion in (
        ("tipo_material", "TEXT"),
        ("npf", "TEXT"),
        ("fabricante_desc", "TEXT"),
        ("oculto", "INTEGER NOT NULL DEFAULT 0"),
    ):
        if columna not in existentes:
            conn.execute(f"ALTER TABLE historial_asignaciones ADD COLUMN {columna} {definicion}")
    conn.execute("CREATE TABLE IF NOT EXISTS migraciones (nombre TEXT PRIMARY KEY, fecha TEXT NOT NULL)")
    conn.commit()
    return conn


class RangoAgotadoError(Exception):
    pass


class RangoNoConfiguradoError(Exception):
    pass


def obtener_rango(nombre_rango: str) -> dict:
    """Obtiene la configuración de un rango por nombre."""
    if nombre_rango in DEFAULT_RANGOS:
        return DEFAULT_RANGOS[nombre_rango]
    
    # Si está en la config file, usarla
    if RANGOS_CONFIG_PATH.exists():
        with open(RANGOS_CONFIG_PATH, encoding="utf-8") as fh:
            custom_rangos = json.load(fh)
            if nombre_rango in custom_rangos:
                return custom_rangos[nombre_rango]
    
    # No encontrado
    raise RangoNoConfiguradoError(
        f"Rango de material '{nombre_rango}' no está configurado. "
        f"Disponibles: {list(DEFAULT_RANGOS.keys())}"
    )


def siguiente_numero(nombre_contador: str, range_min: int = None, range_max: int = None,
                      tipo: str = "", texto_breve: str = "", fabricante_codigo: str = "",
                      tipo_material: str = "", npf: str = "", fabricante_desc: str = "") -> int:
    """
    Entrega atómicamente el siguiente número disponible del rango indicado,
    dejándolo guardado para que nadie más lo reutilice.
    
    Args:
        nombre_contador: Nombre único del rango (ej. "zmaq_material", "zcam_material")
        range_min: Valor mínimo (opcional, se obtiene de config si no se proporciona)
        range_max: Valor máximo (opcional, se obtiene de config si no se proporciona)
        tipo: Tipo de material para historial
        texto_breve: Descripción breve para historial
        fabricante_codigo: Código fabricante para historial
        tipo_material, npf, fabricante_desc: datos extra del material para el historial
    
    Returns:
        int: Siguiente número disponible
    
    Raises:
        RangoAgotadoError: Si se agotó el rango
        RangoNoConfiguradoError: Si no existe configuración para el rango
    """
    # Si no se proporcionan los rangos, obtenerlos de la config
    if range_min is None or range_max is None:
        rango_config = obtener_rango(nombre_contador)
        range_min = rango_config["min"]
        range_max = rango_config["max"]
    
    with _lock:
        conn = _get_conn()
        try:
            cur = conn.execute(
                "SELECT ultimo_valor, range_min, range_max FROM contadores WHERE nombre = ?",
                (nombre_contador,),
            )
            row = cur.fetchone()

            if row is None:
                nuevo_valor = range_min
                conn.execute(
                    "INSERT INTO contadores (nombre, ultimo_valor, range_min, range_max) VALUES (?, ?, ?, ?)",
                    (nombre_contador, nuevo_valor, range_min, range_max),
                )
            else:
                ultimo_valor, r_min, r_max = row
                nuevo_valor = ultimo_valor + 1
                if nuevo_valor > r_max:
                    raise RangoAgotadoError(
                        f"El rango '{nombre_contador}' ({r_min}-{r_max}) está agotado."
                    )
                conn.execute(
                    "UPDATE contadores SET ultimo_valor = ? WHERE nombre = ?",
                    (nuevo_valor, nombre_contador),
                )

            conn.execute(
                """INSERT INTO historial_asignaciones
                   (nombre_contador, numero_asignado, tipo, texto_breve, fabricante_codigo, fecha,
                    tipo_material, npf, fabricante_desc)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (nombre_contador, nuevo_valor, tipo, texto_breve, fabricante_codigo,
                 datetime.now().isoformat(timespec="seconds"), tipo_material, npf, fabricante_desc),
            )
            conn.commit()
            return nuevo_valor
        finally:
            conn.close()


def estado_contador(nombre_contador: str, range_min: int, range_max: int) -> dict:
    conn = _get_conn()
    try:
        cur = conn.execute(
            "SELECT ultimo_valor FROM contadores WHERE nombre = ?", (nombre_contador,)
        )
        row = cur.fetchone()
        ultimo = row[0] if row else range_min - 1
        return {
            "nombre": nombre_contador,
            "ultimo_valor": ultimo,
            "range_min": range_min,
            "range_max": range_max,
            "disponibles": range_max - ultimo,
        }
    finally:
        conn.close()


_COLUMNAS_BUSQUEDA = ("CAST(numero_asignado AS TEXT)", "tipo", "tipo_material", "npf", "texto_breve",
                      "fabricante_codigo", "fabricante_desc")


def _filtro_historial(buscar: str):
    """WHERE del historial: sin los materiales ocultos (cargas mal hechas que se
    conservan solo para que su número no se reutilice) ni el contador interno."""
    where = "nombre_contador != 'material_global' AND oculto = 0"
    params: list = []
    for palabra in buscar.split():
        where += " AND (" + " OR ".join(f"{c} LIKE ?" for c in _COLUMNAS_BUSQUEDA) + ")"
        params += [f"%{palabra}%"] * len(_COLUMNAS_BUSQUEDA)
    return where, params


def contar_historial(buscar: str = "") -> int:
    where, params = _filtro_historial(buscar)
    conn = _get_conn()
    try:
        return conn.execute(f"SELECT COUNT(*) FROM historial_asignaciones WHERE {where}", params).fetchone()[0]
    finally:
        conn.close()


def historial(limit: int = 100, offset: int = 0, buscar: str = "") -> list:
    where, params = _filtro_historial(buscar)
    conn = _get_conn()
    try:
        cur = conn.execute(
            f"""SELECT numero_asignado, tipo, texto_breve, fabricante_codigo, fecha,
                       npf, fabricante_desc
                FROM historial_asignaciones
                WHERE {where}
                ORDER BY fecha DESC, numero_asignado DESC LIMIT ? OFFSET ?""",
            params + [limit, offset],
        )
        cols = ["numero_asignado", "tipo", "texto_breve", "fabricante_codigo", "fecha", "npf", "fabricante_desc"]
        registros = [dict(zip(cols, r)) for r in cur.fetchall()]
        for reg in registros:
            # Se guarda en formato ISO ("2026-10-01T14:32:07") para que el TEXT
            # ordene bien en SQLite; para mostrarlo se separa fecha y hora en
            # columnas propias (ver historial.html) así la hora no se corta a
            # la mitad si el ancho de columna queda justo.
            fecha, _, hora = reg["fecha"].partition("T")
            reg["fecha_fmt"] = fecha
            reg["hora_fmt"] = hora
            for campo in ("texto_breve", "fabricante_codigo", "npf", "fabricante_desc"):
                reg[campo] = reg[campo] or ""
        return registros
    finally:
        conn.close()


CORRELATIVOS_REPUESTOS_PATH = Path(__file__).parent / "data" / "reference" / "Repuestos" / "CorrelativosRepuestos.xlsx"
IMPORTACION_REPUESTOS_NOMBRE = "importar_correlativos_repuestos_prd_2026-10-06"
IMPORTACION_REPUESTOS_FECHA = "2026-10-06T12:30:00"


def importar_correlativos_repuestos() -> int:
    """
    Carga UNA sola vez los correlativos de Repuestos que ya existen en PRD
    (data/reference/Repuestos/CorrelativosRepuestos.xlsx): quedan en el
    historial con fecha/hora fija (el Excel no trae cuándo se crearon) y el
    contador de 'repuestos_material' avanza al último número del Excel, así
    nunca se vuelve a asignar uno ya usado.

    Las combinaciones NPF:Fabricante repetidas (cargas mal hechas, todas
    'NO UTILIZAR:...') se guardan igual para conservar su número, pero con
    oculto=1: no se muestran en el historial.

    Devuelve cuántas filas se agregaron (0 si ya estaba importado o falta el archivo).
    """
    if not CORRELATIVOS_REPUESTOS_PATH.exists():
        return 0
    import pandas as pd

    df = pd.read_excel(CORRELATIVOS_REPUESTOS_PATH, dtype=str, keep_default_na=False)
    col_material, col_tipo, col_npf, col_fab, col_desc, col_clave, col_texto = df.columns[:7]
    df["_num"] = df[col_material].astype(int)
    df["_oculto"] = df[col_clave].str.strip().str.upper().duplicated(keep=False).astype(int)
    nombre_contador = "repuestos_material"
    rango = obtener_rango(nombre_contador)

    with _lock:
        conn = _get_conn()
        try:
            if conn.execute("SELECT 1 FROM migraciones WHERE nombre = ?", (IMPORTACION_REPUESTOS_NOMBRE,)).fetchone():
                return 0
            ya_registrados = {
                r[0] for r in conn.execute(
                    "SELECT numero_asignado FROM historial_asignaciones WHERE nombre_contador = ?", (nombre_contador,)
                )
            }
            filas = [
                (nombre_contador, int(f["_num"]), f"repuestos · {f[col_tipo].strip()}", f[col_texto].strip(), f[col_fab].strip(),
                 IMPORTACION_REPUESTOS_FECHA, f[col_tipo].strip(), f[col_npf].strip(), f[col_desc].strip(), int(f["_oculto"]))
                for _, f in df.iterrows() if int(f["_num"]) not in ya_registrados
            ]
            conn.executemany(
                """INSERT INTO historial_asignaciones
                   (nombre_contador, numero_asignado, tipo, texto_breve, fabricante_codigo, fecha,
                    tipo_material, npf, fabricante_desc, oculto)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                filas,
            )
            maximo = int(df["_num"].max())
            actual = conn.execute("SELECT ultimo_valor FROM contadores WHERE nombre = ?", (nombre_contador,)).fetchone()
            if actual is None:
                conn.execute(
                    "INSERT INTO contadores (nombre, ultimo_valor, range_min, range_max) VALUES (?, ?, ?, ?)",
                    (nombre_contador, maximo, rango["min"], rango["max"]),
                )
            elif actual[0] < maximo:
                conn.execute("UPDATE contadores SET ultimo_valor = ? WHERE nombre = ?", (maximo, nombre_contador))
            conn.execute(
                "INSERT INTO migraciones (nombre, fecha) VALUES (?, ?)",
                (IMPORTACION_REPUESTOS_NOMBRE, datetime.now().isoformat(timespec="seconds")),
            )
            conn.commit()
            return len(filas)
        finally:
            conn.close()
