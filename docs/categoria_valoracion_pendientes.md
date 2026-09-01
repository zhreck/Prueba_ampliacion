Combinaciones de `data/reference/calVal.xlsx` que NO se cargaron en
`config/categoria_valoracion.json` porque no son un match único —
completar/confirmar con Seba antes de agregarlas.

- **ZVEH / F0028 (FAW)**: no hay ninguna fila `ZVEH` para FAW en calVal.xlsx
  (solo existe `1010 ZCAM Camiones Faw`). ¿Un vehículo FAW ampliado por VF00
  usa esa misma categoría de Camiones, o falta una fila `ZVEH` para FAW?
- **ZVEH / F0029 (JETOUR)**: mismo caso — no hay fila `ZVEH` para Jetour
  (solo `1710 ZREP Repuestos Vehíc. Jetour`, que es de Repuestos).
- **ZVEH / F9999 (comodín)**: hay DOS filas candidatas — `1150 VEH. Otras
  Marcas` y `1190 Motos Otras Marcas`. No se puede saber si un material con
  fabricante comodín es auto o moto solo con el código.
- **ZREP / F0012 (HYSTER)**: calVal.xlsx trae DOS filas con la misma
  descripción `Rptos MAQ. HYSTER` pero códigos distintos — `1600` y `1620`.
  Parece un duplicado de carga en la planilla; falta saber cuál es el
  correcto (o si ambos son válidos y hace falta otro criterio para elegir).
- **ZREP / F0021 (LIFAN)**: dos filas candidatas — `1575 Rptos Vehiculos
  LIFAN` y `1700 Repuestos Motos Lifan`. En el resto del sistema, LIFAN está
  clasificado como moto (ver `ZVEH`: `1170 Motos Lifan`), lo que sugeriría
  `1700`, pero no se asumió sin confirmación porque calVal.xlsx sí tiene una
  fila separada "Vehiculos LIFAN" que podría ser intencional.
- **ZREP / F0022 (DFAC - Dong Feng)**: dos filas candidatas — `1520 Rptos
  Camiones Dongfeng` y `1720 Repuestos Vehíc. Dongfeng`. A diferencia de
  F0001 (DFCV, que solo se usa en camiones), F0022 aparece tanto en la tabla
  de Camiones (VA00) como en la de Automotriz (VF00) — la categoría correcta
  depende de la filial del material, no solo del fabricante. Se podría
  resolver mirando la filial de la fila ampliada, pero se dejó pendiente
  hasta confirmar que esa es la regla correcta.
- **ZREP / F0028 (FAW)**: mismo caso que F0022 — `1510 Rptos Camiones Faw`
  vs `1670 Repuestos Vehículos FAW`, FAW aparece en ambos contextos.

Todas estas quedan con "Categoría valoración" vacía en el Excel de salida
(reportadas en el aviso "sin categoría de valoración" cuando corresponda),
nunca con un valor inventado.
