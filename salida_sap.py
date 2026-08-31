"""
Conversión de datos ampliados a formato SAP MM01 (transacción de creación de materiales).

Toma el resultado de engine.procesar() (fila ampliada a N centros) y lo convierte
a las ~150 columnas del formato SAP real, aplicando plantillas por filial/tipo de material.

Las plantillas están en config/salida_sap/*.json y contienen:
- Valores defaults (fijos por filial)
- Reglas dinámicas (mapeo de columnas input -> columnas SAP)
- Campos pendientes de lookup
"""

import json
from pathlib import Path
from typing import Optional, Dict, List, Tuple
import pandas as pd

import correlativo

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "config" / "salida_sap"
HEADER_PATH = BASE_DIR / "docs" / "header_151_columnas.json"
CONFIRMACION_PATH = BASE_DIR / "docs" / "confirmacion_campos_parsed.json"

class FormatoSAPError(Exception):
    pass


def cargar_header_151() -> List[str]:
    """Carga la lista de 151 nombres de columnas SAP en orden exacto."""
    with open(HEADER_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def cargar_confirmacion_sap() -> dict:
    """Carga la matriz de confirmación de campos por bloque."""
    with open(CONFIRMACION_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def cargar_plantilla_sap(tipo_material: str) -> dict:
    """
    Carga la plantilla SAP para un tipo de material.
    
    Args:
        tipo_material: Tipo de material SAP (ZVHE, ZMAQ, ZCAM)
    
    Returns:
        Diccionario con la plantilla de configuración
        
    Raises:
        FormatoSAPError: Si la plantilla no existe o hay problemas
    """
    # Mapear tipo material a archivo de config
    tipo_a_config = {
        "ZVHE": "zvhe.json",
        "ZMAQ": "zmaq.json",
        "ZCAM": "zcam.json",
    }
    
    config_file = tipo_a_config.get(tipo_material)
    if not config_file:
        raise FormatoSAPError(f"Tipo material '{tipo_material}' no tiene plantilla SAP configurada.")
    
    path = CONFIG_DIR / config_file
    if not path.exists():
        raise FormatoSAPError(f"Archivo de plantilla no encontrado: {path}")
    
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def filial_a_tipo_material(filial: str) -> str:
    """Mapea código de filial a tipo de material SAP."""
    mapeo = {
        "VF00": "ZVHE",
        "VA00": "ZCAM",
        "VC00": "ZMAQ",
        "VD00": "ZMAQ",
        "VE00": "ZMAQ",
    }
    
    tipo = mapeo.get(filial)
    if not tipo:
        raise FormatoSAPError(
            f"Filial '{filial}' no tiene mapeo a tipo de material SAP. "
            f"Filiales soportadas: {list(mapeo.keys())}"
        )
    return tipo


def aplicar_plantilla_sap(
    df_ampliado: pd.DataFrame,
    tipo_material: str,
    centro_cebe_lookup: Optional[Dict] = None,
) -> Tuple[pd.DataFrame, Dict]:
    """
    Convierte un DataFrame ampliado al formato SAP MM01.
    
    Args:
        df_ampliado: DataFrame con las filas ya ampliadas a N centros (output de engine.procesar())
        tipo_material: Tipo de material SAP (ZVHE, ZMAQ, ZCAM)
        centro_cebe_lookup: Dict opcional {centro: cod_cebe} para mapeo de centros a CEBE
    
    Returns:
        Tupla (DataFrame de 151 columnas en formato SAP, diccionario de metadatos)
    
    Raises:
        FormatoSAPError: Si hay problemas de configuración
    """
    
    # Cargar configuraciones
    plantilla = cargar_plantilla_sap(tipo_material)
    header = cargar_header_151()
    confirmacion = cargar_confirmacion_sap()
    
    # Determinar qué bloque de la confirmación usar
    bloque_key = plantilla.get("bloque_confirmacion_sap")
    if not bloque_key:
        raise FormatoSAPError(f"Plantilla de {tipo_material} no especifica 'bloque_confirmacion_sap'")
    
    bloque = confirmacion.get(bloque_key)
    if not bloque:
        raise FormatoSAPError(f"Bloque '{bloque_key}' no encontrado en confirmacion_campos_parsed.json")
    
    # Validar que el rango de material esté configurado
    rango_corr_nombre = plantilla.get("rango_correlativo_nombre")
    if "_PENDIENTE" in rango_corr_nombre:
        raise FormatoSAPError(
            f"PENDIENTE: Rango de material para {tipo_material} no está confirmado. "
            f"Seba debe entregar: rango_min y rango_max para {rango_corr_nombre}"
        )
    
    # Obtener transacción
    transaccion = plantilla.get("transaccion", "MM01")
    
    # Resultado: lista de filas SAP
    resultado = []
    numeros_asignados = {}  # Para tracking
    
    # Procesar cada material (agrupar por NUMERO MATERIAL interno)
    for material_idx, grupo in df_ampliado.groupby("NUMERO MATERIAL", sort=False):
        # Obtener el número de material SAP del correlativo
        try:
            numero_sap = correlativo.siguiente_numero(rango_corr_nombre)
        except correlativo.RangoNoConfiguradoError as e:
            raise FormatoSAPError(f"Error de correlativo para {tipo_material}: {e}")
        except correlativo.RangoAgotadoError as e:
            raise FormatoSAPError(f"Rango agotado para {tipo_material}: {e}")
        
        # Procesar cada fila del grupo (todas comparten número SAP)
        for _, row in grupo.iterrows():
            fila_sap = generar_fila_sap(
                row=row,
                plantilla=plantilla,
                bloque=bloque,
                header=header,
                numero_material=numero_sap,
                transaccion=transaccion,
                centro_cebe_lookup=centro_cebe_lookup
            )
            resultado.append(fila_sap)
        
        numeros_asignados[material_idx] = numero_sap
    
    # Crear DataFrame con columnas en orden exacto
    df_sap = pd.DataFrame(resultado)
    
    # Asegurar que todas las 151 columnas estén presentes en orden exacto
    columnas_ordenadas = [col for col in header if col in df_sap.columns]
    columnas_extras = [col for col in df_sap.columns if col not in header]
    df_sap = df_sap[columnas_ordenadas + columnas_extras]
    
    # Preparar metadatos de pendientes y gaps
    metadatos = {
        "tipo_material": tipo_material,
        "bloque": bloque_key,
        "transaccion": transaccion,
        "total_filas": len(df_sap),
        "numeros_asignados": numeros_asignados,
        "campos_pendientes": list(bloque.get("needs_review_or_lookup", {}).keys()),
        "columnas_vacias": identificar_columnas_vacias(df_sap, header),
        "columnas_totales": len(header),
    }
    
    return df_sap, metadatos


def generar_fila_sap(
    row: pd.Series,
    plantilla: dict,
    bloque: dict,
    header: List[str],
    numero_material: int,
    transaccion: str,
    centro_cebe_lookup: Optional[Dict] = None
) -> Dict:
    """
    Convierte una fila ampliada a una fila SAP completa (151 columnas).
    
    Args:
        row: Serie de pandas con los datos ampliados
        plantilla: Configuración de la plantilla SAP
        bloque: Bloque de confirmación (defaults, dynamic_rules, needs_review_or_lookup)
        header: Lista de 151 nombres de columnas SAP
        numero_material: Número de material asignado del correlativo
        transaccion: Código de transacción SAP
        centro_cebe_lookup: Dict opcional para mapear Centro -> COD_CEBE
    
    Returns:
        Diccionario con todas las 151 columnas completadas
    """
    fila_sap = {}
    
    # 1. Inicializar todas las columnas con vacío
    for col in header:
        fila_sap[col] = ""
    
    # 2. Aplicar valores por defecto del bloque
    defaults = bloque.get("defaults", {})
    for col, valor in defaults.items():
        if col in fila_sap:
            fila_sap[col] = valor
    
    # 3. Asignar transacción y número de material
    fila_sap["Transaccion"] = transaccion
    if "Material" in fila_sap:
        fila_sap["Material"] = str(numero_material)
    
    # 4. Aplicar mapeos definidos en la plantilla
    mapeos = plantilla.get("mapeos", {})
    for col_sap, config_mapeo in mapeos.items():
        if col_sap not in fila_sap:
            continue
        
        # Ignorar campos meta
        if config_mapeo == {"from": "numero_sap", "description": "Asignado por correlativo.siguiente_numero()"}:
            continue
        
        if isinstance(config_mapeo, dict):
            # Mapeo complejo
            col_input = config_mapeo.get("from")
            default_val = config_mapeo.get("default")
            
            if col_input and col_input in row.index:
                fila_sap[col_sap] = str(row[col_input])
            elif default_val is not None:
                fila_sap[col_sap] = default_val
        else:
            # Mapeo simple: nombre de columna como string
            if config_mapeo in row.index:
                fila_sap[col_sap] = str(row[config_mapeo])
    
    # 5. Aplicar lookups si están disponibles
    if centro_cebe_lookup and "Centro" in row.index:
        centro = str(row["Centro"])
        if centro in centro_cebe_lookup:
            if "Centro de beneficio" in fila_sap:
                fila_sap["Centro de beneficio"] = centro_cebe_lookup[centro]
    
    return fila_sap


def identificar_columnas_vacias(df: pd.DataFrame, header: List[str]) -> List[str]:
    """Identifica qué columnas de SAP quedaron todas vacías en el output."""
    vacias = []
    for col in header:
        if col in df.columns:
            valores_no_vacios = df[col].astype(str).str.strip()
            if (valores_no_vacios == "").all():
                vacias.append(col)
    return vacias

