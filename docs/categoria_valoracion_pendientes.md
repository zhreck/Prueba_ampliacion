Combinaciones de `data/reference/calVal.xlsx` que NO se cargaron en
`config/categoria_valoracion.json` porque no son un match único —
completar/confirmar con Seba antes de agregarlas.

## Cambio de clave: Fabricante -> Marca (2026-09-02)

Hasta esta fecha, la categoría se resolvía por el código de **Fabricante**
de cada fila. Seba hizo notar que era más simple resolverla por **Marca**
(la columna "Grupo materiales 1" de la salida SAP, ya cargada desde MARCA
CODIGO / CODIGO MARCA del input) — y revisando la evidencia real
(`data/reference/Repuestos/campos_repuestos_zrp3.xlsx`) se confirmó que
además es lo correcto: en esa fila el Fabricante es `F0058` (NACIONAL, un
proveedor genérico de repuestos) pero la Marca es `110` (INTERNATIONAL, la
marca del camión). calVal.xlsx clasifica por marca de vehículo/máquina
("MAQ. HYSTER", "Camiones Faw"), no por quién fabricó la pieza — usar
Fabricante como clave dejaba "sin categoría" filas de Repuestos que sí
tenían una categoría válida por marca.

De paso, este cambio resolvió sin necesidad de una regla especial el caso
LIFAN: el diccionario de marcas (`data/reference/Diccionario_marca_jerarquia.xlsx`)
trae **dos códigos separados** — `LIFAN` (140, motos) y `LIFAN AUTO` (290,
autos). Antes había que asumir "Lifan = motos" siempre porque el
Fabricante no distinguía el caso; Seba confirmó que el input real de
Repuestos sí distingue estas dos marcas, así que ahora cada código resuelve
solo a su categoría (1700 motos / 1575 autos), sin regla especial.

## Resueltas (ya no están pendientes)

- **ZREP / Dongfeng y Faw**: se usan tanto en Camiones como en
  Vehículos/Repuestos de vehículos. Se resuelven con la regla de filial que
  confirmó Seba: la categoría depende del negocio de la filial de la fila
  (VA00=Camiones, VC00/VD00/VE00=Maquinarias, VF00=Vehículos/Motos) — ver
  `ZREP_CAMIONES`/`ZREP_MAQUINARIAS`/`ZREP_VEHICULOS` en
  `categoria_valoracion.json`, elegido según "Org. Ventas".
- **ZREP_MAQUINARIAS / HYSTER (marca 100)**: calVal.xlsx traía dos filas
  duplicadas con la misma descripción "Rptos MAQ. HYSTER" (códigos 1600 y
  1620). Seba confirmó que el correcto es **1600**.
- **ZREP_VEHICULOS / LIFAN (140) y LIFAN AUTO (290)**: ver sección de
  arriba — ya no es una regla asumida, cada marca resuelve a su propia
  categoría (1700 motos / 1575 autos).
- **ZVEH (Modelos, Automotriz) / FAW (marca 80) y JETOUR (marca 120)**:
  calVal.xlsx no tiene fila `ZVEH` propia para ninguna de las dos marcas.
  Seba confirmó que van con **1150 "VEH. Otras Marcas"** (la misma
  categoría que el comodín de autos), no con la de Camiones ni con una fila
  nueva.

## Siguen pendientes

- **ZVEH / marca "OTRAS MARCAS" (280)** y **ZREP_VEHICULOS / marca "OTRAS
  MARCAS" (280)**: en ambos casos hay DOS filas candidatas en calVal.xlsx —
  auto vs moto (`1150`/`1190` en ZVEH, `1585`/`1690` en Repuestos). Ni la
  filial (VF00) ni la marca (ambas vienen genéricas como "Otras Marcas")
  alcanzan para distinguir auto de moto — haría falta otro dato del
  material (ej. el grupo de artículos) para saber cuál aplica, y no se
  quiso asumir uno sin confirmación.
- **ZREP / cod_carVal 1730 "Kit Tattersall"**: no corresponde a ninguna
  marca del diccionario (es un tipo de kit, no una marca) — queda fuera de
  alcance de este cruce por marca.

Todas estas quedan con "Categoría valoración" vacía en el Excel de salida
(reportadas en el aviso "sin categoría de valoración" cuando corresponda),
nunca con un valor inventado.
