Combinaciones de `data/reference/calVal.xlsx` que NO se cargaron en
`config/categoria_valoracion.json` porque no son un match único —
completar/confirmar con Seba antes de agregarlas.

## Resueltas (ya no están pendientes)

Estas quedaron sin resolver en una primera pasada, pero se cerraron con
información que Seba dio después:

- **ZREP / F0022 (DFAC) y F0028 (FAW)**: tenían dos filas candidatas cada
  uno (Camiones vs Vehículos) porque ambos fabricantes se usan en los dos
  negocios. Se resolvió con la regla de filial que confirmó Seba: la
  categoría depende del negocio de la filial de la fila (VA00=Camiones,
  VC00/VD00/VE00=Maquinarias, VF00=Vehículos/Motos) — ver
  `ZREP_CAMIONES`/`ZREP_MAQUINARIAS`/`ZREP_VEHICULOS` en
  `categoria_valoracion.json`, elegido según "Org. Ventas".
- **ZREP / F0012 (HYSTER)**: calVal.xlsx traía dos filas duplicadas con la
  misma descripción "Rptos MAQ. HYSTER" (códigos 1600 y 1620). Seba
  confirmó que el correcto es **1600**.
- **ZREP / F0021 (LIFAN)**: dos filas candidatas — "1575 Rptos Vehiculos
  LIFAN" y "1700 Repuestos Motos Lifan". Seba confirmó usar **1700**, por
  la misma clasificación ya confirmada en ZVEH (LIFAN = moto, "1170 Motos
  Lifan").

## Siguen pendientes

- **ZVEH (Modelos, Automotriz) / F0028 (FAW)**: no hay ninguna fila `ZVEH`
  para FAW en calVal.xlsx (solo existe `1010 ZCAM Camiones Faw`, del lado
  Camiones). Es un caso distinto al de Repuestos: para Repuestos SÍ existe
  `1670 ZREP Repuestos Vehículos FAW`, pero para Modelos/ZVEH no hay
  equivalente. ¿Falta cargar esa fila en calVal.xlsx, o un vehículo FAW
  ampliado por VF00 debería usar la categoría de Camiones (1010) igual que
  Repuestos usa la de Camiones para VA00?
- **ZVEH / F0029 (JETOUR)**: mismo caso — no hay fila `ZVEH` para Jetour
  (solo `1710 ZREP Repuestos Vehíc. Jetour`, que es de Repuestos).
- **ZVEH / F9999 (comodín)** y **ZREP_VEHICULOS / F9999 (comodín)**: en
  ambos casos hay DOS filas candidatas — auto vs moto (`1150`/`1190` en
  ZVEH, `1585`/`1690` en Repuestos). La filial (VF00) no alcanza para
  distinguir auto de moto dentro del mismo VF00 — haría falta otro dato
  del material (ej. el grupo de artículos) para saber cuál aplica, y no se
  quiso asumir uno sin confirmación.

Todas estas quedan con "Categoría valoración" vacía en el Excel de salida
(reportadas en el aviso "sin categoría de valoración" cuando corresponda),
nunca con un valor inventado.
