"""
Conversión de datos ampliados a formato SAP MM01 (transacción de creación de materiales).

Toma el resultado de engine.procesar() (fila ya ampliada a N centros) y lo convierte
a las 151 columnas del formato SAP real, aplicando plantillas por filial/tipo de material.

Fuentes de verdad (no se inventan valores fuera de estas):
- docs/header_151_columnas.json: orden y nombres exactos de columnas (incluye nombres
  repetidos a propósito — así es el layout real de SAP, con grupos de columnas que
  comparten etiqueta, ej. 3 slots de "Unidad de Medida Alternativa (UMA)"). Por eso
  las filas se arman POSICIONALMENTE (por índice), no por diccionario: un dict no
  puede tener dos claves iguales y perdería la distinción entre esas columnas.
- docs/confirmacion_campos_parsed.json: matriz de negocio confirmada, por bloque de
  filial. "defaults" son valores fijos confirmados. "needs_review_or_lookup" son
  campos que el negocio todavía no resolvió — esos se dejan en blanco y se listan en
  la hoja PENDIENTES, nunca se inventa un valor SAP para ellos.
- config/salida_sap/<filial>.json: qué columnas de la fila ya ampliada van a qué
  columna SAP (config-driven, sin lógica por filial hardcodeada en Python), más un
  puñado de valores "defaults_confirmados_extra" evidenciados en un ejemplo de
  salida real (docs/ejemplo_output_ZMAQ_VC00.xlsx) que no estaban en la matriz de
  confirmación pero sí en el output real ya usado en producción.

Reglas universales (no dependen de la plantilla, aplican siempre):
- "Transaccion" = "MM01" para toda fila. Los nombres "tau"/"ttm"/"tti"/"ttmq"/"mtt"
  que aparecen en confirmacion_campos_parsed.json NO son el valor del campo: son la
  etiqueta interna de cada bloque/sub-tipo dentro de confirmacionCampos.xlsx. Esto se
  confirma con el propio Excel (fila de ZCAM dice literalmente
  'Obligatorio\\nValor por defecto "MM01"'), con la vista SQL histórica
  (docs/vista_sql_formato_sap.sql, 'MM01' AS TRANSACCION) y con el ejemplo real
  de ZMAQ/VC00 (columna Transaccion = 'MM01' en las 70 filas).
- "Material" = número asignado por correlativo.siguiente_numero() del rango de la
  filial/tipo de material.
- "Org. Ventas" = el código de FILIAL CODIGO de la fila ampliada (evidenciado en el
  ejemplo real: Org. Ventas es 'VC00' constante para todos los centros VC00..VC07,
  es decir la organización de ventas "propia" de la filial, no el centro expandido).
- "Jerarquía de productos SD\\nMVKE-PRODH" = mismo valor que
  "Jerarquía productos\\n MARA-PRDHA\\nReplicar en\\nMVKE-PRODH". Esto SÍ está resuelto
  en confirmacion_campos_parsed.json (nota: "Tomar el dato que se registró en
  MARA-PRDHA"), aunque el parser lo haya agrupado bajo needs_review_or_lookup.
- "Categoría Clase", "Clase", "Unidad medida pedido" y "Unidad med.salida" quedan
  SIEMPRE vacíos, para las tres plantillas. Seba confirmó esto directamente
  (arreglos_notas.txt) y pisa
  lo que dice confirmacion_campos_parsed.json (que trae "300"/"ZMAQUINAS"/"UN" en
  "defaults" para estos campos) — el ejemplo real de ZMAQ ya los traía vacíos.

ZRP1/ZRP3 (Repuestos, normal/seriado) no tienen bloque en
confirmacion_campos_parsed.json — no existe un confirmacionCampos.xlsx
equivalente para Repuestos, solo dos ejemplos de salida real
(data/reference/Repuestos/campos_repuestos_zrp{1,3}.xlsx). "bloque_confirmacion_sap"
queda sin definir en esas plantillas y todo sale de config/salida_sap/zrp{1,3}.json
(defaults_confirmados_extra + campos_desde_ampliado + obligatorios propios). Además,
ahí se evidenció que Repuestos repite cada centro una vez por Canal distribución
(20 y 30) — plantilla config "canales_distribucion": ["20","30"] activa esa
expansión adicional, a diferencia de Modelos/Maquinaria donde el canal es un
único valor constante.
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
CATEGORIA_VALORACION_PATH = BASE_DIR / "config" / "categoria_valoracion.json"
MARCAS_PATH = BASE_DIR / "config" / "marcas.json"
GRUPO_COMPRAS_PATH = BASE_DIR / "config" / "grupo_compras.json"

COL_CATEGORIA_VALORACION = "Categoría valoración"
COL_MARCA = "Grupo materiales 1"
COL_GRUPO_COMPRAS = "Grupo de compras"

COL_TRANSACCION = "Transaccion"
COL_MATERIAL = "Material"
COL_ORG_VENTAS = "Org. Ventas"
COL_CANAL = "Canal distribución"
COL_JERARQUIA_MARA = "Jerarquía productos\n MARA-PRDHA\nReplicar en\nMVKE-PRODH"
COL_JERARQUIA_SD = "Jerarquía de productos SD\nMVKE-PRODH"

TRANSACCION_VALOR = "MM01"

# Confirmado directamente por Seba (arreglos_notas.txt): estos campos van SIEMPRE
# vacíos en las tres plantillas, pase lo que diga confirmacion_campos_parsed.json.
CAMPOS_FORZAR_VACIO = {"Categoría Clase", "Clase", "Unidad medida pedido", "Unidad med.salida"}

# Filiales que ya tienen tipo de material y plantilla configurados. ZUSA queda
# fuera a propósito (fuera de alcance, ver docs/mapeo_filiales.json).
FILIAL_A_TIPO_MATERIAL = {
    "VF00": "ZVHE",
    "VA00": "ZCAM",
    "VC00": "ZMAQ",
    "VD00": "ZMAQ",
    "VE00": "ZMAQ",
}

TIPO_MATERIAL_A_CONFIG = {
    "ZVHE": "zvhe.json",
    "ZMAQ": "zmaq.json",
    "ZCAM": "zcam.json",
    "ZRP1": "zrp1.json",
    "ZRP3": "zrp3.json",
}


class FormatoSAPError(Exception):
    pass


def cargar_header_151() -> List[str]:
    """Carga la lista de 151 nombres de columnas SAP, en orden exacto y CON duplicados."""
    with open(HEADER_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def cargar_categoria_valoracion() -> dict:
    """Carga config/categoria_valoracion.json: {grupo: {cod_marca: cod_carVal}}."""
    with open(CATEGORIA_VALORACION_PATH, encoding="utf-8") as fh:
        return json.load(fh).get("grupos", {})


def cargar_marcas() -> dict:
    """Carga config/marcas.json: {cod_marca: nombre_marca} (solo para avisos legibles)."""
    with open(MARCAS_PATH, encoding="utf-8") as fh:
        return json.load(fh).get("marcas", {})


def cargar_grupo_compras() -> dict:
    """Carga config/grupo_compras.json: {filial: {cod_marca: grupo} | {'_default': grupo}}."""
    with open(GRUPO_COMPRAS_PATH, encoding="utf-8") as fh:
        return json.load(fh).get("filiales", {})


def cargar_confirmacion_sap() -> dict:
    """Carga la matriz de confirmación de campos por bloque."""
    with open(CONFIRMACION_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def cargar_plantilla_sap(tipo_material: str) -> dict:
    config_file = TIPO_MATERIAL_A_CONFIG.get(tipo_material)
    if not config_file:
        raise FormatoSAPError(f"Tipo material '{tipo_material}' no tiene plantilla SAP configurada.")

    path = CONFIG_DIR / config_file
    if not path.exists():
        raise FormatoSAPError(f"Archivo de plantilla no encontrado: {path}")

    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def filial_a_tipo_material(filial: str) -> str:
    """Mapea código de filial a tipo de material SAP."""
    tipo = FILIAL_A_TIPO_MATERIAL.get(filial)
    if not tipo:
        raise FormatoSAPError(
            f"Filial '{filial}' no tiene mapeo a tipo de material SAP. "
            f"Filiales soportadas: {list(FILIAL_A_TIPO_MATERIAL.keys())}"
        )
    return tipo


def _primeras_ocurrencias(header: List[str]) -> Dict[str, int]:
    """Índice de la PRIMERA posición de cada nombre de columna en el header.

    El header tiene nombres repetidos a propósito (así es el layout real de SAP).
    Toda vez que este módulo escribe un valor "por nombre", lo hace en la primera
    ocurrencia; las ocurrencias siguientes del mismo nombre quedan en blanco, que es
    exactamente el patrón observado en docs/ejemplo_output_ZMAQ_VC00.xlsx (ej.
    'Centro de beneficio' solo trae valor en su primera aparición, la segunda
    siempre viene vacía en las 70 filas del ejemplo).
    """
    primeras = {}
    for i, nombre in enumerate(header):
        if nombre not in primeras:
            primeras[nombre] = i
    return primeras


def _valor_texto(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _normalizar_codigo_marca(valor) -> str:
    """
    Normaliza un código de marca a 3 dígitos con cero a la izquierda (ej. "80"
    o "80.0" -> "080"), que es el formato real que trae
    data/reference/Diccionario_marca_jerarquia.xlsx (columna CODIGO, celda de
    texto tipo '080') y con el que están armados config/marcas.json,
    categoria_valoracion.json y grupo_compras.json. Si el valor no es
    numérico se devuelve tal cual (no debería pasar con un código de marca
    real, pero evita reventar si llega algo raro).
    """
    texto = str(valor).strip()
    if not texto:
        return ""
    try:
        return f"{int(float(texto)):03d}"
    except ValueError:
        return texto


def aplicar_plantilla_sap(
    df_ampliado: pd.DataFrame,
    tipo_material: str,
) -> Tuple[pd.DataFrame, Dict]:
    """
    Convierte un DataFrame ampliado al formato SAP MM01 (151 columnas, en el orden
    y con los nombres exactos de docs/header_151_columnas.json).

    Args:
        df_ampliado: filas ya ampliadas a N centros (salida de engine.procesar()).
        tipo_material: ZVHE, ZMAQ o ZCAM.

    Returns:
        (DataFrame de 151 columnas, metadatos con pendientes/avisos)
    """
    plantilla = cargar_plantilla_sap(tipo_material)
    header = cargar_header_151()
    primeras = _primeras_ocurrencias(header)

    bloque_key = plantilla.get("bloque_confirmacion_sap")
    if bloque_key:
        confirmacion = cargar_confirmacion_sap()
        bloque = confirmacion.get(bloque_key)
        if not bloque:
            raise FormatoSAPError(f"Bloque '{bloque_key}' no encontrado en confirmacion_campos_parsed.json")
    else:
        # Plantillas sin equivalente en confirmacionCampos.xlsx (ej. ZRP1/ZRP3):
        # todo sale directo de esta plantilla (defaults_confirmados_extra +
        # campos_desde_ampliado + obligatorios), evidenciado en un ejemplo de
        # salida real en vez de en la matriz de confirmación de negocio.
        bloque = {}

    rango_corr_nombre = plantilla.get("rango_correlativo_nombre")
    if not rango_corr_nombre or "PENDIENTE" in rango_corr_nombre:
        raise FormatoSAPError(
            f"PENDIENTE: el rango de número de material para {tipo_material} no está "
            f"confirmado todavía (rango_correlativo_nombre='{rango_corr_nombre}'). "
            f"No se puede generar el formato SAP para esta filial hasta tener el rango real."
        )
    # Falla fuerte y explícita si el rango no quedó registrado en correlativo.py.
    correlativo.obtener_rango(rango_corr_nombre)

    campos_pendientes = _calcular_campos_pendientes(bloque, plantilla)

    if "FILIAL CODIGO" not in df_ampliado.columns:
        raise FormatoSAPError("La fila ampliada no tiene columna 'FILIAL CODIGO'; no se puede resolver Org. Ventas.")

    resultado_filas = []
    numeros_asignados = {}

    for material_idx, grupo in df_ampliado.groupby("NUMERO MATERIAL", sort=False):
        try:
            numero_sap = correlativo.siguiente_numero(rango_corr_nombre)
        except correlativo.RangoNoConfiguradoError as e:
            raise FormatoSAPError(f"Error de correlativo para {tipo_material}: {e}") from e
        except correlativo.RangoAgotadoError as e:
            raise FormatoSAPError(f"Rango agotado para {tipo_material}: {e}") from e

        numeros_asignados[str(material_idx)] = numero_sap

        canales = plantilla.get("canales_distribucion")
        for _, row in grupo.iterrows():
            if canales:
                # Ej. Repuestos: cada centro se repite una vez por canal (20 y
                # 30), evidenciado en docs.../campos_repuestos_zrp{1,3}.xlsx.
                for canal in canales:
                    resultado_filas.append(
                        _generar_fila_sap(row, plantilla, bloque, header, primeras, numero_sap, campos_pendientes, canal_forzado=canal)
                    )
            else:
                resultado_filas.append(
                    _generar_fila_sap(row, plantilla, bloque, header, primeras, numero_sap, campos_pendientes)
                )

    df_sap = pd.DataFrame(resultado_filas, columns=header)

    marcas_sin_categoria = _resolver_categoria_valoracion(df_sap, plantilla, primeras)
    marcas_sin_grupo_compras = _resolver_grupo_compras(df_sap, plantilla, primeras)

    metadatos = {
        "tipo_material": tipo_material,
        "bloque": bloque_key,
        "transaccion": TRANSACCION_VALOR,
        "total_filas": len(df_sap),
        "numeros_asignados": numeros_asignados,
        "campos_pendientes": campos_pendientes,
        "columnas_obligatorias_vacias": _obligatorios_vacios(df_sap, header, plantilla, bloque, primeras),
        "marcas_sin_categoria_valoracion": marcas_sin_categoria,
        "marcas_sin_grupo_compras": marcas_sin_grupo_compras,
    }

    return df_sap, metadatos


def _resolver_categoria_valoracion(df_sap: pd.DataFrame, plantilla: dict, primeras: Dict[str, int]) -> List[str]:
    """
    Llena "Categoría valoración" según la MARCA de cada fila (columna "Grupo
    materiales 1", ya cargada por campos_desde_ampliado con MARCA CODIGO /
    CODIGO MARCA), usando config/categoria_valoracion.json. Se usa la marca y
    no el Fabricante porque calVal.xlsx describe categorías por marca de
    vehículo/máquina (ej. "MAQ. HYSTER", "Camiones Faw"), y en Repuestos el
    Fabricante es quien fabrica/provee la pieza — puede ser un tercero sin
    relación con la marca (evidenciado en campos_repuestos_zrp3.xlsx:
    Fabricante=F0058/NACIONAL pero marca=110/INTERNATIONAL).

    Solo pone valores confirmados sin ambigüedad — una marca sin entrada en
    el grupo que corresponda queda con el campo vacío y se reporta en la
    lista que devuelve esta función (ver docs/categoria_valoracion_pendientes.md).

    Dos formas de configurar "categoria_valoracion" en la plantilla:
    - {"grupo": "ZMAQ"}: un solo grupo fijo (Modelos — cada plantilla ya es
      de una sola filial/negocio).
    - {"grupos_por_filial": {"VA00": "ZREP_CAMIONES", ...}}: el grupo depende
      del negocio de la filial de cada fila (Repuestos: un mismo archivo
      puede traer materiales de varias filiales). Se lee de "Org. Ventas",
      que la regla universal ya deja igual al código de filial de la fila.
    """
    cfg = plantilla.get("categoria_valoracion")
    if not cfg or df_sap.empty:
        return []

    todos_los_grupos = cargar_categoria_valoracion()
    marcas = cargar_marcas()
    grupo_fijo = cfg.get("grupo")
    grupos_por_filial = cfg.get("grupos_por_filial")

    idx_categoria = primeras[COL_CATEGORIA_VALORACION]
    idx_marca = primeras[COL_MARCA]
    idx_filial = primeras.get(COL_ORG_VENTAS) if grupos_por_filial else None

    sin_categoria = set()
    for i in range(len(df_sap)):
        marca = _normalizar_codigo_marca(df_sap.iat[i, idx_marca])
        if not marca:
            continue

        if grupo_fijo:
            nombre_grupo = grupo_fijo
        else:
            filial = str(df_sap.iat[i, idx_filial]).strip()
            nombre_grupo = grupos_por_filial.get(filial)
            if not nombre_grupo:
                sin_categoria.add(marcas.get(marca, marca))
                continue

        valor = todos_los_grupos.get(nombre_grupo, {}).get(marca)
        if valor:
            df_sap.iat[i, idx_categoria] = valor
        else:
            sin_categoria.add(marcas.get(marca, marca))

    return sorted(sin_categoria)


def _resolver_grupo_compras(df_sap: pd.DataFrame, plantilla: dict, primeras: Dict[str, int]) -> List[str]:
    """
    Llena "Grupo de compras" según filial ('Org. Ventas') + marca ('Grupo
    materiales 1'), usando config/grupo_compras.json — tabla que Seba pasó en
    arreglos_notas.txt (grupo de compras de Unidades, no de Repuestos: no se
    usa para ZRP1/ZRP3, que ya tienen su propio valor confirmado por ejemplo
    real). Solo aplica si la plantilla tiene "grupo_de_compras_dinamico": true.

    Si la filial no está en la tabla, no se toca el campo (sigue el
    comportamiento anterior: vacío/pendiente). Si la filial tiene una sola
    opción fija (clave "_default", ej. VF00) se usa esa sin mirar la marca.
    Si la filial tiene varias opciones por marca (ej. VA00/VC00) y la marca de
    la fila no está en la tabla, el campo queda vacío y se reporta en la
    lista que devuelve esta función.
    """
    if not plantilla.get("grupo_de_compras_dinamico") or df_sap.empty:
        return []

    tabla = cargar_grupo_compras()
    marcas = cargar_marcas()

    idx_grupo = primeras[COL_GRUPO_COMPRAS]
    idx_marca = primeras[COL_MARCA]
    idx_filial = primeras[COL_ORG_VENTAS]

    sin_grupo = set()
    for i in range(len(df_sap)):
        filial = str(df_sap.iat[i, idx_filial]).strip()
        opciones = tabla.get(filial)
        if not opciones:
            continue

        if "_default" in opciones:
            df_sap.iat[i, idx_grupo] = opciones["_default"]
            continue

        marca = _normalizar_codigo_marca(df_sap.iat[i, idx_marca])
        valor = opciones.get(marca)
        if valor:
            df_sap.iat[i, idx_grupo] = valor
        else:
            sin_grupo.add(marcas.get(marca, marca))

    return sorted(sin_grupo)


def _calcular_campos_pendientes(bloque: dict, plantilla: dict) -> Dict[str, str]:
    """
    Campos que quedan en blanco a propósito porque el negocio todavía no dio una
    regla, con el motivo. Nunca se rellenan con un valor inventado.

    Un campo deja de estar "pendiente" apenas la plantilla de la filial le da una
    regla propia (defaults_confirmados_extra o campos_desde_ampliado) — así, cuando
    Seba confirma un valor para una filial puntual (ej. "Grupo de compras" solo para
    ZMAQ), alcanza con agregarlo al JSON de esa plantilla; no hace falta tocar esta
    función ni la lista needs_review_or_lookup de confirmacion_campos_parsed.json.
    """
    resueltos_aparte = {COL_TRANSACCION, COL_MATERIAL, COL_JERARQUIA_SD} | CAMPOS_FORZAR_VACIO
    if plantilla.get("categoria_valoracion"):
        # No es un "pendiente" de plantilla completa: se resuelve fila por fila
        # según la marca (ver _resolver_categoria_valoracion). Las marcas que
        # no tengan categoría confirmada se reportan aparte, en
        # metadatos["marcas_sin_categoria_valoracion"].
        resueltos_aparte = resueltos_aparte | {COL_CATEGORIA_VALORACION}
    if plantilla.get("grupo_de_compras_dinamico"):
        # Mismo patrón que categoria_valoracion: se resuelve fila por fila
        # según filial + marca (ver _resolver_grupo_compras), no es un
        # "pendiente" de plantilla completa. Lo que no matchee se reporta en
        # metadatos["marcas_sin_grupo_compras"].
        resueltos_aparte = resueltos_aparte | {COL_GRUPO_COMPRAS}
    resueltos_por_plantilla = (
        set(plantilla.get("defaults_confirmados_extra", {}).keys())
        | set(plantilla.get("campos_desde_ampliado", {}).keys())
    )

    pendientes = {}
    for campo, nota in bloque.get("needs_review_or_lookup", {}).items():
        if campo in resueltos_aparte or campo in resueltos_por_plantilla:
            continue
        pendientes[campo] = nota

    for campo in plantilla.get("pendientes_extra", []):
        if campo in resueltos_por_plantilla:
            continue
        pendientes.setdefault(campo, "Obligatorio sin regla confirmada (no evidenciado en un ejemplo real de esta filial).")

    return pendientes


def _generar_fila_sap(
    row: pd.Series,
    plantilla: dict,
    bloque: dict,
    header: List[str],
    primeras: Dict[str, int],
    numero_material: int,
    campos_pendientes: Dict[str, str],
    canal_forzado: Optional[str] = None,
) -> List[str]:
    fila = [""] * len(header)

    def set_valor(nombre_columna: str, valor) -> None:
        if valor is None:
            return
        idx = primeras.get(nombre_columna)
        if idx is None:
            return
        texto = _valor_texto(valor)
        if texto == "":
            return
        fila[idx] = texto

    # 1) Defaults confirmados por negocio (confirmacion_campos_parsed.json), salvo
    #    los que Seba pidió dejar siempre vacíos (pisan lo que diga este bloque).
    for campo, valor in bloque.get("defaults", {}).items():
        if campo in CAMPOS_FORZAR_VACIO:
            continue
        set_valor(campo, valor)

    # 2) Defaults extra evidenciados en un ejemplo real de esta filial (config-driven).
    for campo, valor in plantilla.get("defaults_confirmados_extra", {}).items():
        set_valor(campo, valor)

    # 3) Columnas que vienen directo de la fila ya ampliada (mapeo config-driven).
    for col_sap, col_ampliado in plantilla.get("campos_desde_ampliado", {}).items():
        if col_ampliado in row.index:
            set_valor(col_sap, row[col_ampliado])

    # 4) Campos pendientes de negocio: se dejan explícitamente en blanco, aunque
    #    algún paso anterior haya intentado poner algo (nunca se inventa un valor).
    for campo in campos_pendientes:
        idx = primeras.get(campo)
        if idx is not None:
            fila[idx] = ""

    # 5) Reglas universales, siempre pisan lo anterior.
    set_valor(COL_TRANSACCION, TRANSACCION_VALOR)
    set_valor(COL_MATERIAL, numero_material)
    if "FILIAL CODIGO" in row.index:
        set_valor(COL_ORG_VENTAS, row["FILIAL CODIGO"])
    if canal_forzado is not None:
        set_valor(COL_CANAL, canal_forzado)
    idx_jerarquia_mara = primeras.get(COL_JERARQUIA_MARA)
    if idx_jerarquia_mara is not None:
        set_valor(COL_JERARQUIA_SD, fila[idx_jerarquia_mara])

    return fila


def _obligatorios_vacios(df: pd.DataFrame, header: List[str], plantilla: dict, bloque: dict, primeras: Dict[str, int]) -> List[str]:
    """De la lista de 'obligatorios' (de la plantilla, o si no del bloque de
    confirmación), cuáles quedaron vacíos en TODAS las filas."""
    vacios = []
    obligatorios = plantilla.get("obligatorios") or bloque.get("obligatorios", [])
    for campo in obligatorios:
        idx = primeras.get(campo)
        if idx is None:
            continue
        columna = df.iloc[:, idx].astype(str).str.strip()
        if (columna == "").all():
            vacios.append(campo)
    return vacios
