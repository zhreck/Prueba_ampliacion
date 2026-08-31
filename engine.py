"""
Motor genérico de ampliación de materiales a centros.

La idea: un "tipo" (Modelos, Repuestos, ...) se define solo con un JSON en
config/tipos/*.json. Este motor no conoce reglas de negocio de ningún tipo
en particular; todo sale de la config + la tabla de referencia (Excel).

Para agregar un tipo nuevo (ej. Repuestos) más adelante:
  1. Crear config/tipos/repuestos.json con sus propias columnas de input,
     su propio archivo de referencia y su propia key de matching.
  2. Dejar el archivo de referencia en data/reference/.
  3. Listo — aparece automático en el selector de la web.
"""

import json
from pathlib import Path

import pandas as pd

import correlativo

BASE_DIR = Path(__file__).parent
TIPOS_DIR = BASE_DIR / "config" / "tipos"


class TipoNoEncontradoError(Exception):
    pass


class InputInvalidoError(Exception):
    pass


def listar_tipos() -> list[dict]:
    tipos = []
    for f in sorted(TIPOS_DIR.glob("*.json")):
        with open(f, encoding="utf-8") as fh:
            cfg = json.load(fh)
        tipos.append({"id": cfg["id"], "nombre": cfg["nombre"], "descripcion": cfg.get("descripcion", "")})
    return tipos


def cargar_config(tipo_id: str) -> dict:
    path = TIPOS_DIR / f"{tipo_id}.json"
    if not path.exists():
        raise TipoNoEncontradoError(f"No existe configuración para el tipo '{tipo_id}'.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _leer_input(file_storage, cfg: dict) -> pd.DataFrame:
    # Preparar dtype dict y converters para especificar columnas de texto
    text_columns = cfg.get("text_columns", [])
    dtype_dict = {col: str for col in text_columns}
    converters_dict = {col: str for col in text_columns}
    
    try:
        df = pd.read_excel(
            file_storage, 
            sheet_name=cfg["input_sheet"], 
            dtype=dtype_dict,
            converters=converters_dict
        )
    except ValueError as e:
        raise InputInvalidoError(
            f"No se encontró la hoja '{cfg['input_sheet']}' en el Excel subido."
        ) from e

    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    columnas_esperadas = cfg["input_columns"]
    faltantes = [c for c in columnas_esperadas if c not in df.columns]
    if faltantes:
        raise InputInvalidoError(
            "El Excel de input no tiene las columnas esperadas: " + ", ".join(faltantes)
        )

    # Solo nos quedamos con las columnas definidas por la plantilla, en su orden,
    # y descartamos cualquier fila que venga completamente vacía en la key de match.
    df = df[columnas_esperadas].copy()
    
    # Convertir columnas especificadas a texto para evitar notación científica y pérdida de ceros
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
    
    df = df.dropna(subset=[cfg["key_input"]])
    if df.empty:
        raise InputInvalidoError("El Excel no tiene ninguna fila de datos para procesar.")
    
    return df


def _cargar_referencia(ref_file: str, ref_sheet) -> pd.DataFrame:
    ref_path = BASE_DIR / ref_file
    if not ref_path.exists():
        raise InputInvalidoError(f"Falta el archivo de referencia: {ref_path}")
    return pd.read_excel(ref_path, sheet_name=ref_sheet or 0)


def _resolver_reglas_referencia(cfg: dict, filial: str | None) -> dict:
    """
    Resuelve qué reference_file / key_reference / fallback_value / output_columns_from_reference
    / expandir_todos usar para una fila, según su FILIAL CODIGO.

    Por defecto usa los valores "planos" del tipo (reference_file, key_reference, ...).
    Si el tipo define "reference_by_filial", la filial de la fila puede pisar cualquiera
    de esos campos (por ejemplo, ZMAQ no hace match por fabricante: siempre expande a los
    mismos centros, así que usa "expandir_todos": true con su propia tabla centro->cebe).
    """
    base = {
        "reference_file": cfg.get("reference_file"),
        "reference_sheet": cfg.get("reference_sheet"),
        "key_reference": cfg.get("key_reference"),
        "fallback_value": cfg.get("fallback_value"),
        "output_columns_from_reference": cfg.get("output_columns_from_reference", []),
        "expandir_todos": False,
    }
    overrides_por_filial = cfg.get("reference_by_filial", {})
    if filial and filial in overrides_por_filial:
        base.update(overrides_por_filial[filial])
    return base


def procesar(tipo_id: str, file_storage) -> pd.DataFrame:
    """
    Punto de entrada principal: recibe el Excel del usuario (file-like) y
    devuelve el DataFrame ya ampliado, listo para exportar.
    """
    cfg = cargar_config(tipo_id)
    input_df = _leer_input(file_storage, cfg)

    key_input = cfg["key_input"]
    text_columns = cfg.get("text_columns", [])

    corr_cfg = cfg.get("correlativo", {"enabled": False})

    # Cache de tablas de referencia ya cargadas, por (reference_file, reference_sheet),
    # para no releer el mismo Excel por cada fila del input.
    _referencias_cache: dict[tuple, pd.DataFrame] = {}

    def _referencia_para(reglas: dict) -> pd.DataFrame:
        clave = (reglas["reference_file"], reglas["reference_sheet"])
        if clave not in _referencias_cache:
            _referencias_cache[clave] = _cargar_referencia(reglas["reference_file"], reglas["reference_sheet"])
        return _referencias_cache[clave]

    filas_salida = []
    avisos = []

    for _, fila_input in input_df.iterrows():
        filial_fila = str(fila_input.get("FILIAL CODIGO", "")).strip() or None
        reglas = _resolver_reglas_referencia(cfg, filial_fila)
        ref_df = _referencia_para(reglas)
        key_ref = reglas["key_reference"]
        fallback = reglas["fallback_value"]
        cols_from_ref = reglas["output_columns_from_reference"]

        valor_key = fila_input[key_input]

        if reglas["expandir_todos"]:
            # Esta filial no hace match por fabricante: todas las filas de
            # referencia (ej. la lista fija de centros de la filial) aplican
            # a todos los materiales por igual.
            matches = ref_df
            uso_fallback = False
        else:
            matches = ref_df[ref_df[key_ref] == valor_key]

            uso_fallback = False
            if matches.empty and fallback:
                matches = ref_df[ref_df[key_ref] == fallback]
                uso_fallback = True

        if matches.empty:
            avisos.append(
                f"Fabricante '{valor_key}' no encontrado en la tabla de referencia "
                f"(y no hay fallback disponible) — fila omitida."
            )
            continue

        # Un mismo material conserva UN solo número, repetido en todas sus
        # filas ampliadas (no uno distinto por centro).
        numero_material = None
        if corr_cfg.get("enabled"):
            numero_material = correlativo.siguiente_numero(
                nombre_contador=corr_cfg["nombre"],
                range_min=corr_cfg["range_min"],
                range_max=corr_cfg["range_max"],
                tipo=tipo_id,
                texto_breve=str(fila_input.get("TEXTO BREVE", "")),
                fabricante_codigo=str(valor_key),
            )

        for _, fila_ref in matches.iterrows():
            fila_out = {}
            # Copiar valores preservando tipos de texto
            for col in fila_input.index:
                if col in text_columns:
                    fila_out[col] = str(fila_input[col])
                else:
                    fila_out[col] = fila_input[col]
            
            for col in cols_from_ref:
                if col in text_columns:
                    fila_out[col] = str(fila_ref.get(col, ""))
                else:
                    fila_out[col] = fila_ref.get(col)
            
            if numero_material is not None:
                fila_out[corr_cfg["column_name"]] = numero_material
            if uso_fallback:
                fila_out["_FALLBACK_USADO"] = "SI"
            filas_salida.append(fila_out)

    if not filas_salida:
        raise InputInvalidoError(
            "No se generó ninguna fila de salida. " + " ".join(avisos)
        )

    resultado = pd.DataFrame(filas_salida)
    
    # Aplicar tipos de datos específicos
    for col in text_columns:
        if col in resultado.columns:
            resultado[col] = resultado[col].astype(str)
    
    resultado.attrs["avisos"] = avisos
    return resultado
