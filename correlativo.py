"""
Manejo del correlativo de "Número de material".

Soporta múltiples rangos con nombres distintos:
- "material_global" (1 - 5.000.000): Original para ampliación sin SAP
- "zmaq_material" (20000000 - 29999999): Maquinarias SAP
- "zcam_material" (30000000 - 39999999): Camiones SAP
- "zvhe_material_PENDIENTE": ERROR - Automotriz no tiene rango confirmado todavía

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
                      tipo: str = "", texto_breve: str = "", fabricante_codigo: str = "") -> int:
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
                   (nombre_contador, numero_asignado, tipo, texto_breve, fabricante_codigo, fecha)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (nombre_contador, nuevo_valor, tipo, texto_breve, fabricante_codigo,
                 datetime.now().isoformat(timespec="seconds")),
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


def historial(limit: int = 100) -> list:
    conn = _get_conn()
    try:
        cur = conn.execute(
            """SELECT numero_asignado, tipo, texto_breve, fabricante_codigo, fecha
               FROM historial_asignaciones ORDER BY id DESC LIMIT ?""",
            (limit,),
        )
        cols = ["numero_asignado", "tipo", "texto_breve", "fabricante_codigo", "fecha"]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()
