# Cambios Implementados en Ampliación de Materiales

Fecha: 31 de Agosto de 2026

## 📋 Resumen de Cambios

Se han implementado 3 cambios principales en la configuración y procesamiento de datos:

---

## 1️⃣ JERARQUIA como Texto (sin notación científica)

**Problema:** Valores numéricos largos en JERARQUIA se mostraban en notación científica.

**Solución:** 
- Ahora JERARQUIA se procesa como texto desde la lectura del Excel
- Se preservan todos los dígitos y caracteres especiales (guiones, etc.)
- Ejemplo: `VEH-001` se mantiene tal cual (no se convierte a notación científica)

**Implementación:**
- Archivo: `config/tipos/modelos.json`
- Campo agregado: `"text_columns": ["JERARQUIA", "MARCA CODIGO", "NPF:Fabricante"]`
- Engine: `engine.py` - función `_leer_input()` y `procesar()`

---

## 2️⃣ MARCA CODIGO como Texto (preserva ceros iniciales)

**Problema:** MARCA CODIGO con ceros iniciales (ej: 090) se guardaba como 90, perdiendo el cero.

**Solución:**
- MARCA CODIGO ahora se trata como texto en todo el proceso
- Se preservan los ceros iniciales (090 → "090")
- Se preservan otros caracteres especiales si los hubiera

**Implementación:**
- Agregado a `text_columns` en `config/tipos/modelos.json`
- Pandas lee con `dtype={'MARCA CODIGO': str}`
- El valor se mantiene como string en todo el procesamiento

**⚠️ Instrucción Importante para el Usuario:**

Cuando crees el archivo Excel en Microsoft Excel o Calc, **DEBES formatear las celdas de MARCA CODIGO como TEXTO** antes de escribir los valores:

1. Selecciona la columna MARCA CODIGO
2. Click derecho → Formato de celdas
3. Pestaña "Número" → Categoría "Texto"
4. Escribe tus valores (090, 091, etc.)

O simplemente **usa los archivos de ejemplo** (`EJEMPLO_INPUT.xlsx`) que ya tienen el formato correcto.

---

## 3️⃣ Nuevos Campos: NPF y NPF:Fabricante

**Campos Agregados:**
- `NPF`: Número de referencia del proveedor (text, ingresado por el usuario)
- `NPF:Fabricante`: Referencia adicional (text, ingresado por el usuario)

---

## 📊 Estructura de Columnas del Excel de Input

El Excel debe tener la hoja `INPUT` con estas columnas (en cualquier orden, pero todos deben estar presentes):

| # | Columna | Tipo | Requerido | Notas |
|----|---------|------|-----------|-------|
| 1 | FILIAL CODIGO | Texto | ✓ | Código de la filial |
| 2 | UNIDAD DE MEDIDA | Texto | ✓ | ej: UN, KG |
| 3 | TEXTO BREVE | Texto | ✓ | Descripción breve del material |
| 4 | GRUPO ARTICULO | Texto | ✓ | Clasificación |
| 5 | NOMBRE JERARQUIA | Texto | ✓ | Nombre de la jerarquía |
| 6 | JERARQUIA | **Texto** | ✓ | Código jerarquía (⚠️ **FORMATO TEXTO**) |
| 7 | MARCA CODIGO | **Texto** | ✓ | Código marca, ej: 090 (⚠️ **FORMATO TEXTO**) |
| 8 | NOMBRE MARCA | Texto | ✓ | Nombre de la marca |
| 9 | FABRICANTE CODIGO | Texto | ✓ | Código fabricante |
| 10 | NOMBRE FABRICANTE | Texto | ✓ | Nombre fabricante |
| 11 | NPF | **Texto** | ✓ | Campo nuevo |
| 12 | NPF:Fabricante | **Texto** | ✓ | Campo nuevo |

---

## 🎯 Resultado del Procesamiento

El Excel generado incluirá todas las columnas del input + las siguientes columnas generadas automáticamente:

- `CENTRO`: Agregado desde tabla de referencia
- `COD_CEBE`: Agregado desde tabla de referencia
- `NOM_CODIGO_CEBE`: Agregado desde tabla de referencia
- `NUMERO MATERIAL`: Generado automáticamente (número correlativo único)

---

## 📥 Cómo Usar

### Opción 1: Usar archivos de ejemplo (Recomendado)

1. Descarga `EJEMPLO_INPUT.xlsx` de la carpeta `download`
2. Abre con Excel/Calc
3. Reemplaza los datos manteniendo la estructura
4. Sube y procesa

### Opción 2: Crear desde cero

1. Crea un Excel con la hoja `INPUT`
2. Agrega los encabezados de columnas (ver tabla arriba)
3. **Importante:** Formatea las columnas 6, 7 y 12 como TEXTO
4. Ingresa tus datos
5. Sube y procesa

---

## ✅ Test de Validación

Para verificar que todo funciona:

1. Descarga `EJEMPLO_INPUT.xlsx` → Procesa → ✓ Debe funcionar
2. Descarga `EJEMPLO_DUPLICADO.xlsx` → Procesa → ✗ Debe rechazarse con error

---

## 🔧 Archivos Modificados

```
config/tipos/modelos.json          ← Agregados: text_columns, unique_columns, validation_errors
engine.py                          ← Modificada función _leer_input() y procesar()
download/EJEMPLO_INPUT.xlsx        ← Actualizado con nuevos campos
download/EJEMPLO_DUPLICADO.xlsx    ← Nuevo archivo de prueba
```

---

## 💡 Notas Técnicas

- Pandas lee Excel con `dtype` especificado para preservar tipos
- Los archivos de ejemplo usan openpyxl con formato de celda '@' (texto)
- Validación de duplicados en lectura de input antes de procesar
- Los mensajes de error son claros y sugieren correcciones

---

## ❓ Preguntas Frecuentes

**P: ¿Por qué mi archivo con 090 se convierte en 90?**  
R: Porque Excel lo está guardando como número. Formatea la columna como TEXTO antes de ingresar datos.

**P: ¿Qué pasa si tengo dos filas con el mismo NPF:Fabricante?**  
R: El sistema rechaza el archivo y muestra un error. Cada NPF:Fabricante debe ser único.

**P: ¿Puedo cambiar el orden de las columnas en el Excel?**  
R: Sí, el orden no importa. El sistema busca por nombre de columna.

**P: ¿Qué es NPF?**  
R: Es un nuevo campo que debes proporcionar. Actualmente es solo para información/trazabilidad.

---

*Documento generado automáticamente. Para soporte, contacta al equipo de desarrollo.*
