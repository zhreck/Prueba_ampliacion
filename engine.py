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

import copy
import unicodedata
import json
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import correlativo

BASE_DIR = Path(__file__).parent
TIPOS_DIR = BASE_DIR / "config" / "tipos"


class TipoNoEncontradoError(Exception):
    pass


class InputInvalidoError(Exception):
    pass


def _sin_tildes(texto: str) -> str:
    """Quita tildes/diéresis (á->a, ü->u) conservando la Ñ."""
    texto = texto.replace("Ñ", "\0N").replace("ñ", "\0n")
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    return texto.replace("\0N", "Ñ").replace("\0n", "ñ")


def normalizar_valor_texto(v):
    """Sin espacios al inicio/final, MAYÚSCULAS y sin tildes. Solo toca strings."""
    if isinstance(v, str):
        return _sin_tildes(v.strip()).upper()
    return v


def normalizar_df_texto(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(normalizar_valor_texto)
    return df


def _formatear_costo(v):
    """Costo como texto numérico limpio: 5000.0 -> '5000', '10,99' -> '10.99'; vacío se queda vacío."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    texto = str(v).strip().replace(",", ".")
    if texto == "":
        return ""
    try:
        numero = float(texto)
    except ValueError:
        return texto
    return str(int(numero)) if numero == int(numero) else str(numero)


def listar_tipos() -> list[dict]:
    tipos = []
    for f in sorted(TIPOS_DIR.glob("*.json")):
        with open(f, encoding="utf-8") as fh:
            cfg = json.load(fh)
        if cfg.get("oculto"):
            continue
        tipos.append({
            "id": cfg["id"],
            "nombre": cfg["nombre"],
            "descripcion": cfg.get("descripcion", ""),
            # Solo mostrar un botón de diccionario APARTE cuando no viene ya
            # bundleado en la plantilla (Modelos: diccionario_referencia,
            # singular, no se bundlea). Repuestos (diccionarios_referencia,
            # plural) ya trae todo en el mismo archivo de la plantilla — un
            # segundo botón sería redundante y confuso.
            "tiene_diccionario": bool(cfg.get("diccionario_referencia")),
        })
    return tipos


def cargar_config(tipo_id: str) -> dict:
    path = TIPOS_DIR / f"{tipo_id}.json"
    if not path.exists():
        raise TipoNoEncontradoError(f"No existe configuración para el tipo '{tipo_id}'.")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def generar_plantilla_vacia(tipo_id: str, filas_vacias: int = 200) -> "openpyxl.Workbook":
    """
    Arma un Excel vacío con la hoja y columnas de input del tipo, para que
    alguien sin el Excel original tenga de dónde partir.

    Si la config tiene "plantilla_desde_archivo" (Repuestos), se CLONA la
    hoja de input real tal cual (mismo formato, colores, anchos de columna y
    validaciones/listas desplegables de Excel — ej. Serie o Lote?, Unidad
    Medida, Moneda) y solo se le borran los valores de las filas de ejemplo,
    en vez de armar una hoja nueva desde cero que perdía todo eso. Si no
    (Modelos), se sigue armando desde cero como antes, con formato de celda
    texto ('@') en las columnas de "text_columns" para evitar el problema de
    siempre (ceros iniciales / notación científica).
    """
    cfg = cargar_config(tipo_id)
    columnas = cfg["input_columns"]
    text_columns = set(cfg.get("text_columns", []))

    plantilla_real = cfg.get("plantilla_desde_archivo")
    if plantilla_real:
        origen_path = BASE_DIR / plantilla_real["reference_file"]
        wb = openpyxl.load_workbook(origen_path)
        ws = wb[plantilla_real["sheet"]]
        if ws.title != cfg["input_sheet"]:
            ws.title = cfg["input_sheet"]
        # Borra los valores de las filas de ejemplo (fila 2 en adelante) sin
        # tocar formato, anchos de columna ni validaciones — esas quedan
        # intactas porque son propiedades de la hoja/rango, no de la celda.
        for fila in ws.iter_rows(min_row=2, max_row=max(ws.max_row, filas_vacias + 1)):
            for celda in fila:
                celda.value = None
        # Las hojas de diccionario del archivo real ya vienen con formato
        # correcto — no hace falta reconstruirlas con _agregar_hoja_diccionario.
        for nombre_col in cfg.get("columnas_excluir_plantilla", []):
            _quitar_columna_plantilla(ws, nombre_col)
        for nombre_hoja in plantilla_real.get("hojas_excluir", []):
            if nombre_hoja in wb.sheetnames:
                del wb[nombre_hoja]
        _agregar_validaciones_desde_traduccion(wb, ws, cfg)
        return wb

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cfg["input_sheet"]
    ws.append(columnas)

    for row in range(2, 2 + filas_vacias):
        for col_idx, nombre_col in enumerate(columnas, start=1):
            if nombre_col in text_columns:
                ws.cell(row=row, column=col_idx).number_format = "@"

    for col_idx, nombre_col in enumerate(columnas, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = max(14, len(nombre_col) + 2)

    # Modelos: el diccionario NO se bundlea acá (ya tiene un botón de
    # descarga aparte) — ver diccionario_referencia (singular) vs.
    # diccionarios_referencia (plural, Repuestos) en app.py/_agregar_hoja_diccionario.
    for diccionario_cfg in cfg.get("diccionarios_referencia", []):
        _agregar_hoja_diccionario(wb, diccionario_cfg)

    return wb


def _quitar_columna_plantilla(ws, nombre_col: str) -> None:
    """Borra de la hoja de input la columna con ese encabezado (si existe) y las
    listas desplegables que la cubrían."""
    idx = next((c for c in range(1, ws.max_column + 1)
                if str(ws.cell(row=1, column=c).value or "").strip() == nombre_col), None)
    if idx is None:
        return
    for dv in list(ws.data_validations.dataValidation):
        if any(r.min_col <= idx <= r.max_col for r in dv.sqref.ranges):
            ws.data_validations.dataValidation.remove(dv)
    ws.delete_cols(idx)


def _agregar_validaciones_desde_traduccion(wb: "openpyxl.Workbook", ws, cfg: dict) -> None:
    """
    Agrega listas desplegables de Excel a las columnas del input que se
    traducen por nombre (ver "traduccion_nombres" en config/tipos/*.json),
    para que el usuario solo pueda elegir un valor que el backend después
    sabe traducir a código — evita que escriba algo que no está en ningún
    diccionario (pedido explícito de Seba para Filial, Grupo de Artículo,
    Texto jerarquía, ORIGINAL_ALTERNATIVO, Marca y Fabricante). La lista de
    cada columna sale de la MISMA regla que usa el backend para traducir
    (el "mapa" a mano, o la hoja+columna del diccionario real), así que
    nunca puede quedar desincronizada con lo que realmente se acepta.

    Columnas que ya traían una validación funcionando en el archivo
    original (Serie o Lote?, Unidad Medida, Moneda) no se tocan. Una que
    esté rota (ej. la de Marca/Fabricante venía con "#REF!", una referencia
    perdida del archivo original) se reemplaza por una que sí funciona.
    """
    columnas = cfg["input_columns"]

    columnas_con_validacion_propia = set()
    for dv in list(ws.data_validations.dataValidation):
        if str(dv.formula1) == "#REF!":
            ws.data_validations.dataValidation.remove(dv)
            continue
        for rango in dv.sqref.ranges:
            columnas_con_validacion_propia.add(rango.min_col)

    for regla in cfg.get("traduccion_nombres", []):
        col_origen = regla["columna_origen"]
        if col_origen not in columnas:
            continue
        col_idx = columnas.index(col_origen) + 1
        if col_idx in columnas_con_validacion_propia:
            continue
        letra = get_column_letter(col_idx)
        rango_destino = f"{letra}2:{letra}1048576"

        if regla["tipo"] == "mapa":
            opciones = ",".join(str(k) for k in regla["mapa"].keys())
            dv = DataValidation(type="list", formula1=f'"{opciones}"', allow_blank=True)
        else:  # "diccionario" o "jerarquia": la lista sale de la hoja real bundleada
            nombre_hoja = regla["diccionario_sheet"]
            if nombre_hoja not in wb.sheetnames:
                continue
            ws_dicc = wb[nombre_hoja]
            col_nombre = regla.get("col_nombre") or regla.get("col_texto_completo")
            col_letra_dicc = next(
                (get_column_letter(c) for c in range(1, ws_dicc.max_column + 1)
                 if ws_dicc.cell(row=1, column=c).value == col_nombre),
                None,
            )
            if not col_letra_dicc:
                continue
            col_num_dicc = openpyxl.utils.column_index_from_string(col_letra_dicc)
            ultima_fila_con_dato = max(
                (r for r in range(1, ws_dicc.max_row + 1) if ws_dicc.cell(row=r, column=col_num_dicc).value not in (None, "")),
                default=1,
            )
            dv = DataValidation(
                type="list",
                formula1=f"'{nombre_hoja}'!${col_letra_dicc}$2:${col_letra_dicc}${ultima_fila_con_dato}",
                allow_blank=True,
            )

        dv.add(rango_destino)
        ws.add_data_validation(dv)


def _agregar_hoja_diccionario(wb: "openpyxl.Workbook", diccionario_cfg: dict) -> None:
    """Copia una hoja de referencia (ej. un diccionario de códigos) como hoja
    extra de la plantilla descargable, preservando formato (colores, anchos
    de columna, formato de celda texto '@' para que no se pierdan ceros a la
    izquierda) — no solo los valores."""
    origen_path = BASE_DIR / diccionario_cfg["reference_file"]
    if not origen_path.exists():
        return
    wb_origen = openpyxl.load_workbook(origen_path, data_only=True)
    ws_origen = wb_origen[diccionario_cfg["sheet"]]

    ws_destino = wb.create_sheet(title=diccionario_cfg.get("titulo_hoja", "DICCIONARIO")[:31])
    _clonar_hoja(ws_origen, ws_destino)


def _clonar_hoja(ws_origen, ws_destino, max_filas: int | None = None) -> None:
    """Copia valores + formato (fuente, relleno, alineación, formato de
    número) celda por celda, más anchos de columna. `max_filas` limita
    cuántas filas de datos copiar (útil para no clonar decenas de miles de
    filas vacías de una plantilla SAP real)."""
    tope = min(ws_origen.max_row, max_filas) if max_filas else ws_origen.max_row
    for fila in ws_origen.iter_rows(min_row=1, max_row=tope):
        for celda_origen in fila:
            celda_destino = ws_destino.cell(row=celda_origen.row, column=celda_origen.column)
            celda_destino.value = celda_origen.value
            if celda_origen.has_style:
                celda_destino.font = copy.copy(celda_origen.font)
                celda_destino.fill = copy.copy(celda_origen.fill)
                celda_destino.border = copy.copy(celda_origen.border)
                celda_destino.alignment = copy.copy(celda_origen.alignment)
                celda_destino.number_format = celda_origen.number_format

    for letra, dim in ws_origen.column_dimensions.items():
        ws_destino.column_dimensions[letra].width = dim.width

    for rango in ws_origen.merged_cells.ranges:
        ws_destino.merge_cells(str(rango))


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
            converters=converters_dict,
            keep_default_na=False,
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
    df = df.reset_index(drop=True)

    # Algunos tipos tienen columnas de input con nombres largos/propios (ej.
    # Repuestos: "Texto breve (máximo 40 caracteres)") que el resto del motor
    # y salida_sap.py ya conocen por un nombre corto interno (ej. "TEXTO
    # BREVE", igual que en Modelos) — este rename es solo cosmético, pasa
    # antes de traducir nombres a códigos.
    renombrar = cfg.get("renombrar_columnas", {})
    df = df.rename(columns=renombrar)

    # Repuestos: sin espacios sobrantes (NPF incluido), mayúsculas y sin tildes
    # en todo texto, antes de traducir/validar.
    if cfg.get("normalizar_texto"):
        df = normalizar_df_texto(df)
    columna_costo = cfg.get("costo_columna")
    if columna_costo and columna_costo in df.columns:
        df[columna_costo] = df[columna_costo].map(_formatear_costo)

    # Algunos tipos (Repuestos) reciben el input por NOMBRE (marca, fabricante,
    # filial...) y necesitan traducirlo a código antes de que el resto del
    # motor pueda hacer matching contra la tabla de referencia (que sí es por
    # código). Filas cuyo nombre no existe en ningún diccionario se excluyen
    # acá mismo (nunca se inventa un código a partir de un nombre parecido).
    avisos_traduccion: list[str] = []
    for regla in cfg.get("traduccion_nombres", []):
        df, avisos_regla = _traducir_columna(df, regla)
        avisos_traduccion.extend(avisos_regla)

    respaldo = cfg.get("columna_con_respaldo")
    if respaldo:
        preferida = df[respaldo["columna_preferida"]].astype(str).str.strip()
        df[respaldo["columna_destino"]] = preferida.where(preferida != "", df[respaldo["columna_respaldo"]].astype(str))

    avisos_validacion: list[str] = []
    validaciones = cfg.get("validaciones")
    if validaciones:
        df, avisos_validacion = _validar_filas(df, validaciones)

    # Cada tipo puede llamar distinto a su columna de filial (Modelos: "FILIAL
    # CODIGO", Repuestos: ya llega como "FILIAL CODIGO" después de traducir el
    # nombre de la sociedad...). Hacia adentro del sistema siempre se trabaja
    # con "FILIAL CODIGO", en mayúsculas (Vc00, vC00, VC00... se tratan
    # igual), para no tener que enseñarle el nombre real a app.py/salida_sap.py.
    columna_filial = cfg.get("columna_filial", "FILIAL CODIGO")
    if columna_filial in df.columns:
        df[columna_filial] = df[columna_filial].astype(str).str.strip().str.upper()
        if columna_filial != "FILIAL CODIGO":
            df = df.rename(columns={columna_filial: "FILIAL CODIGO"})

    # Convertir columnas especificadas a texto para evitar notación científica y pérdida de ceros
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    if df.empty:
        raise InputInvalidoError(
            "Ninguna fila del Excel pasó las validaciones. " + " ".join(avisos_traduccion + avisos_validacion)
        )

    df = df.dropna(subset=[cfg["key_input"]])
    if df.empty:
        raise InputInvalidoError("El Excel no tiene ninguna fila de datos para procesar.")

    df.attrs["avisos_lectura"] = avisos_traduccion + avisos_validacion
    return df


def _traducir_columna(df: pd.DataFrame, regla: dict) -> tuple[pd.DataFrame, list[str]]:
    """
    Traduce una columna del input "por nombre" a la columna de código que el
    resto del motor espera, según una regla de config/tipos/*.json
    ("traduccion_nombres"). Dos tipos de regla:
    - "mapa": diccionario chico definido a mano en el JSON (ej. Filial -> cod.
      sociedad, con una excepción de negocio como VE00 -> VC00; o Serie o
      Lote? -> tipo de material ZRP1/ZRP2/ZRP3).
    - "diccionario": lee nombre->código desde una hoja de Excel (Marca,
      Fabricante, Grupo de Artículo, Original/Alternativo).
    En ambos casos, si el valor de la fila no matchea ninguna entrada
    (comparando sin mayúsculas/minúsculas ni espacios extra), la fila se
    excluye con un aviso — nunca se inventa ni se aproxima un código.
    """
    col_origen = regla["columna_origen"]
    col_destino = regla["columna_destino"]

    if regla["tipo"] == "mapa":
        mapa = regla["mapa"]
    elif regla["tipo"] == "diccionario":
        origen_path = BASE_DIR / regla["diccionario_file"]
        dicc_df = pd.read_excel(origen_path, sheet_name=regla["diccionario_sheet"], dtype=str, keep_default_na=False)
        dicc_df = dicc_df.dropna(subset=[regla["col_nombre"], regla["col_codigo"]])
        mapa = dict(zip(dicc_df[regla["col_nombre"]], dicc_df[regla["col_codigo"]]))
    elif regla["tipo"] == "jerarquia":
        # Caso especial: el código no es una columna del diccionario, sino la
        # concatenación de 3 (Nivel 1/2/3, cada uno con cero a la izquierda a
        # un ancho fijo) — evidenciado en el único ejemplo real disponible
        # (docs/ejemplo_output_ZMAQ_VC00.xlsx: "002030000100000002" = Nivel 1
        # "00203" + Nivel 2 "00001" + Nivel 3 "00000002"). El match es exacto
        # porque el input trae el camino COMPLETO (ej. "FAW TRUCK REPUESTOS -
        # REFRIGERACION - BOMBAS"), igual que la columna "Texto Jerarquía"
        # del diccionario -- no el nombre de la hoja suelto (que sí se repite
        # muchas veces entre marcas/negocios).
        origen_path = BASE_DIR / regla["diccionario_file"]
        dicc_df = pd.read_excel(origen_path, sheet_name=regla["diccionario_sheet"], dtype=str, keep_default_na=False)
        dicc_df = dicc_df.dropna(subset=[regla["col_texto_completo"]])
        if regla.get("sin_relleno_niveles_vacios"):
            # Un nivel vacío no se rellena con ceros: la jerarquía de 1 o 2
            # niveles simplemente es más corta (5 / 10 / 18 dígitos).
            def _nivel(col, ancho):
                return dicc_df[col].fillna("").astype(str).str.strip().map(lambda v: v.zfill(ancho) if v else "")
        else:
            def _nivel(col, ancho):
                return dicc_df[col].fillna("").str.zfill(ancho)
        codigo = (
            dicc_df[regla["col_nivel1"]].str.zfill(regla["ancho_nivel1"])
            + _nivel(regla["col_nivel2"], regla["ancho_nivel2"])
            + _nivel(regla["col_nivel3"], regla["ancho_nivel3"])
        )
        mapa = dict(zip(dicc_df[regla["col_texto_completo"]], codigo))
    else:
        raise InputInvalidoError(f"Tipo de traducción desconocido: {regla['tipo']}")

    mapa_normalizado = {_sin_tildes(str(k).strip()).upper(): v for k, v in mapa.items()}

    if col_origen not in df.columns:
        return df, []

    valores_normalizados = df[col_origen].astype(str).map(lambda v: _sin_tildes(v.strip()).upper())
    codigos = valores_normalizados.map(mapa_normalizado)

    avisos = []
    sin_match = codigos.isna() & df[col_origen].notna() & (df[col_origen].astype(str).str.strip() != "")
    if sin_match.any():
        valores_no_reconocidos = sorted(df.loc[sin_match, col_origen].astype(str).unique())
        avisos.append(
            f"'{col_origen}' con valor(es) no reconocido(s) en el diccionario, fila(s) omitida(s): "
            + ", ".join(valores_no_reconocidos)
        )

    df[col_destino] = codigos
    df = df[~sin_match].copy()
    return df, avisos


def _validar_filas(df: pd.DataFrame, cfg_validaciones: dict) -> tuple[pd.DataFrame, list[str]]:
    """
    Validaciones genéricas de negocio sobre el input ya traducido, config-driven
    (ver "validaciones" en config/tipos/*.json). Nunca corrige el dato: si una
    fila no pasa, se excluye y se reporta en el aviso — el usuario tiene que
    arreglar el Excel de origen.
    """
    avisos = []
    filas_validas = pd.Series(True, index=df.index)

    for col in cfg_validaciones.get("campos_obligatorios", []):
        if col not in df.columns:
            continue
        vacio = df[col].isna() | (df[col].astype(str).str.strip() == "")
        if vacio.any():
            avisos.append(f"'{col}' vacío en {vacio.sum()} fila(s) — fila(s) omitida(s) (campo obligatorio).")
            filas_validas &= ~vacio

    eq_req = cfg_validaciones.get("equivalencia_requerida_si")
    if eq_req and eq_req["columna"] in df.columns and eq_req["columna_requerida"] in df.columns:
        aplica = df[eq_req["columna"]].astype(str).str.strip().str.upper() == eq_req["valor"].upper()
        falta = aplica & (df[eq_req["columna_requerida"]].isna() | (df[eq_req["columna_requerida"]].astype(str).str.strip() == ""))
        if falta.any():
            avisos.append(
                f"'{eq_req['columna_requerida']}' vacío en {falta.sum()} fila(s) con "
                f"{eq_req['columna']}='{eq_req['valor']}' — fila(s) omitida(s)."
            )
            filas_validas &= ~falta

    col_texto = cfg_validaciones.get("texto_breve_columna")
    max_len = cfg_validaciones.get("texto_breve_max_len")
    if col_texto and max_len and col_texto in df.columns:
        largo = df[col_texto].astype(str).str.len()
        muy_largo = largo > max_len
        if muy_largo.any():
            avisos.append(
                f"'{col_texto}' con más de {max_len} caracteres en {muy_largo.sum()} fila(s) — fila(s) omitida(s)."
            )
            filas_validas &= ~muy_largo

    costo_moneda = cfg_validaciones.get("costo_moneda")
    if costo_moneda:
        col_costo = costo_moneda["columna_costo"]
        col_moneda = costo_moneda["columna_moneda"]
        reglas = costo_moneda["reglas"]
        if col_costo in df.columns and col_moneda in df.columns:
            tiene_ambos = (
                df[col_costo].notna() & (df[col_costo].astype(str).str.strip() != "")
                & df[col_moneda].notna() & (df[col_moneda].astype(str).str.strip() != "")
            )
            malos = pd.Series(False, index=df.index)
            for i in df.index[tiene_ambos]:
                moneda = str(df.at[i, col_moneda]).strip().upper()
                regla = reglas.get(moneda)
                if not regla:
                    continue
                crudo = str(df.at[i, col_costo]).strip()
                try:
                    valor = float(crudo.replace(",", "."))
                except ValueError:
                    malos.at[i] = True
                    continue
                # Una vez que Excel tiene la celda como número real, no hay
                # forma de saber si el usuario "escribió" decimales o no
                # (1000 y 1000.00 son el mismo float) — la única regla
                # verificable de forma confiable es matemática: si es CLP,
                # no puede tener centavos reales (el peso chileno no tiene
                # decimales). USD sí puede tener centavos, pero no los
                # exige — un monto entero en USD (1000.00) es válido igual
                # (confirmado por Seba), así que del lado "float" no se
                # rechaza nada, solo se valida que sea un número.
                es_entero = valor == int(valor)
                if regla == "int" and not es_entero:
                    malos.at[i] = True
            if malos.any():
                avisos.append(
                    f"'{col_costo}'/'{col_moneda}' inconsistentes (CLP no puede tener centavos) "
                    f"en {malos.sum()} fila(s) — fila(s) omitida(s)."
                )
                filas_validas &= ~malos

    if col_texto and col_texto in df.columns:
        df[col_texto] = df[col_texto].astype(str).str.upper()

    return df[filas_validas].copy(), avisos


def _cargar_referencia(ref_file: str, ref_sheet, text_columns: list[str] | None = None) -> pd.DataFrame:
    ref_path = BASE_DIR / ref_file
    if not ref_path.exists():
        raise InputInvalidoError(f"Falta el archivo de referencia: {ref_path}")
    dtype = {col: str for col in text_columns} if text_columns else None
    return pd.read_excel(ref_path, sheet_name=ref_sheet or 0, dtype=dtype)


def _cargar_disponibilidad(cfg: dict) -> dict[tuple[str, str], float] | None:
    """
    Matriz Centro x Fabricante confirmada por negocio (1 = habilitado,
    0.5 = por confirmar pero se trata como habilitado, 0 = bloqueado: ese
    fabricante no tiene CEBE para ese centro y NO puede generarse ahí,
    aunque la tabla de referencia fabricante->centro/cebe tenga una fila
    para esa combinación (puede ser un error de carga en esa tabla).

    Layout fijo de la hoja (ver data/reference/Centros_UN_RP.xlsx, hoja
    "Unidades"): fila 1 = códigos de fabricante desde la columna 3 en
    adelante, filas 3+ = un centro por fila (columna 1 = CENTRO).

    Devuelve None si el tipo no configuró "disponibilidad" (nadie valida
    nada, comportamiento anterior sin cambios).
    """
    disp_cfg = cfg.get("disponibilidad")
    if not disp_cfg:
        return None

    path = BASE_DIR / disp_cfg["reference_file"]
    if not path.exists():
        raise InputInvalidoError(f"Falta el archivo de disponibilidad: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[disp_cfg["sheet"]]

    fabricantes = [ws.cell(row=1, column=c).value for c in range(3, ws.max_column + 1)]
    lookup: dict[tuple[str, str], float] = {}
    for r in range(3, ws.max_row + 1):
        centro = ws.cell(row=r, column=1).value
        if not centro:
            continue
        for i, fab in enumerate(fabricantes):
            if not fab:
                continue
            valor = ws.cell(row=r, column=3 + i).value
            if valor is not None:
                lookup[(str(centro).strip(), str(fab).strip())] = float(valor)
    return lookup


def _etiqueta_tipo_historial(tipo_id: str, fila_input) -> str:
    """Tipo que se muestra en el historial de correlativos: el tipo de
    ampliación + el tipo de material SAP de la fila (ej. 'repuestos · ZRP1')."""
    col = {"repuestos": "TIPO MATERIAL REPUESTO", "modelos": "TIPO MATERIAL"}.get(tipo_id)
    tipo_sap = str(fila_input.get(col, "")).strip() if col else ""
    return f"{tipo_id} · {tipo_sap}" if tipo_sap else tipo_id


def procesar(tipo_id: str, file_storage) -> pd.DataFrame:
    """
    Punto de entrada principal: recibe el Excel del usuario (file-like) y
    devuelve el DataFrame ya ampliado, listo para exportar.
    """
    cfg = cargar_config(tipo_id)
    input_df = _leer_input(file_storage, cfg)
    ref_df_completa = _cargar_referencia(
        cfg["reference_file"], cfg.get("reference_sheet"), cfg.get("reference_text_columns")
    )

    key_input = cfg["key_input"]
    key_ref = cfg["key_reference"]
    fallback = cfg.get("fallback_value")
    cols_from_ref = cfg["output_columns_from_reference"]
    text_columns = cfg.get("text_columns", [])
    filial_column = cfg.get("filial_column")
    reference_filter = cfg.get("reference_filter")

    if reference_filter:
        ref_df_completa = ref_df_completa[
            ref_df_completa[reference_filter["column"]] == reference_filter["value"]
        ]

    disponibilidad = _cargar_disponibilidad(cfg)
    bloqueado_valor = cfg.get("disponibilidad", {}).get("valor_bloqueado", 0)

    corr_cfg = cfg.get("correlativo", {"enabled": False})

    filas_salida = []
    avisos = list(input_df.attrs.get("avisos_lectura", []))

    def _filtrar_por_disponibilidad(matches_df, fabricante):
        """Saca del match los centros que la matriz de disponibilidad marca
        como bloqueados para ese fabricante, aunque la tabla de referencia
        traiga una fila (dato posiblemente mal cargado ahí)."""
        if disponibilidad is None or matches_df.empty:
            return matches_df, []
        centros_excluidos = []
        indices_ok = []
        for idx, fila_ref in matches_df.iterrows():
            centro = str(fila_ref["CENTRO"]).strip()
            estado = disponibilidad.get((centro, str(fabricante).strip()))
            if estado == bloqueado_valor:
                centros_excluidos.append(centro)
            else:
                indices_ok.append(idx)
        return matches_df.loc[indices_ok], centros_excluidos

    def _fabricante_bloqueado_en_toda_la_filial(ref_df_filial, fabricante) -> bool:
        """
        True si la matriz de disponibilidad conoce a este fabricante para AL
        MENOS uno de los centros de esta filial, y lo marca en 0 para TODOS
        ellos (nunca en 1 o 0.5). Cubre el caso en que la tabla de referencia
        fabricante->centro/cebe ni siquiera tiene una fila para esta
        combinación (por eso no basta con filtrar matches_crudos): un
        fabricante conocido-pero-no-habilitado en esta filial no debe caer en
        el comodín F9999, tiene que rechazarse.
        Si la matriz no conoce el fabricante para ningún centro de la filial,
        devuelve False (es un fabricante nuevo/no catalogado: sigue el
        comportamiento normal de fallback).
        """
        if disponibilidad is None or ref_df_filial.empty:
            return False
        centros_filial = ref_df_filial["CENTRO"].astype(str).str.strip().unique()
        estados = [
            disponibilidad[(centro, str(fabricante).strip())]
            for centro in centros_filial
            if (centro, str(fabricante).strip()) in disponibilidad
        ]
        if not estados:
            return False
        return all(e == bloqueado_valor for e in estados)

    for _, fila_input in input_df.iterrows():
        valor_key = fila_input[key_input]

        ref_df = ref_df_completa
        if filial_column:
            filial_fila = str(fila_input.get("FILIAL CODIGO", "")).strip()
            ref_df = ref_df[ref_df[filial_column] == filial_fila]

        matches_crudos = ref_df[ref_df[key_ref] == valor_key]
        matches, centros_excluidos = _filtrar_por_disponibilidad(matches_crudos, valor_key)

        if not matches_crudos.empty and matches.empty:
            # El fabricante SÍ tiene fila(s) en la tabla de referencia, pero la
            # matriz de disponibilidad los bloquea a todos: no caer al comodín,
            # sería disfrazar un fabricante conocido-pero-no-habilitado como
            # "todas/otras marcas".
            avisos.append(
                f"Fabricante '{valor_key}' no habilitado para {', '.join(sorted(set(centros_excluidos)))} "
                f"(matriz de disponibilidad) — fila omitida."
            )
            continue

        if matches.empty and _fabricante_bloqueado_en_toda_la_filial(ref_df, valor_key):
            avisos.append(
                f"Fabricante '{valor_key}' no habilitado en ningún centro de "
                f"'{filial_fila if filial_column else ''}' (matriz de disponibilidad) — fila omitida."
            )
            continue

        uso_fallback = False
        if matches_crudos.empty and fallback:
            matches_crudos = ref_df[ref_df[key_ref] == fallback]
            matches, _ = _filtrar_por_disponibilidad(matches_crudos, fallback)
            uso_fallback = True

        if matches.empty:
            avisos.append(f"Fabricante '{valor_key}' no encontrado — fila omitida.")
            continue

        if centros_excluidos:
            avisos.append(
                f"Fabricante '{valor_key}': se excluyeron {', '.join(sorted(set(centros_excluidos)))} "
                f"(no habilitados), se amplió al resto."
            )

        # Un mismo material conserva UN solo número, repetido en todas sus
        # filas ampliadas (no uno distinto por centro).
        numero_material = None
        if corr_cfg.get("enabled"):
            numero_material = correlativo.siguiente_numero(
                nombre_contador=corr_cfg["nombre"],
                range_min=corr_cfg["range_min"],
                range_max=corr_cfg["range_max"],
                tipo=_etiqueta_tipo_historial(tipo_id, fila_input),
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
    
    if cfg.get("normalizar_texto"):
        resultado = normalizar_df_texto(resultado)

    resultado.attrs["avisos"] = avisos
    return resultado
