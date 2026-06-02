# Documentación: Slicers Dinámicos en Reportes Excel

## 📋 Resumen
Esta funcionalidad permite crear **filtros dinámicos (slicers)** en los reportes Excel generados, específicamente en la hoja **T+G** (Tablas y Gráficos). Los slicers permiten filtrar interactivamente todos los cálculos y métricas sin modificar la estructura visual del reporte.

## 🎯 Objetivo
Permitir al usuario analizar los datos segmentados por diferentes dimensiones (ej: por programa académico, por sede, por tipo de estudiante, etc.) sin necesidad de generar múltiples reportes. Un solo archivo Excel con slicers reemplaza decenas de reportes estáticos.

## ⚙️ Implementación Técnica

### 1. **Columna Auxiliar `_VISIBLE`**
Se agrega automáticamente una columna oculta en la tabla TB que detecta qué filas están visibles:

```excel
Columna: _VISIBLE
Fórmula: =SUBTOTAL(103, A2)
```

- `SUBTOTAL(103, ...)` cuenta celdas visibles (no ocultas por filtros)
- Devuelve `1` si la fila está visible, `0` si está oculta por un slicer
- Esta columna se oculta automáticamente para no afectar la visualización

### 2. **Función Helper `countif_visible()`**
Todas las fórmulas de conteo se generan dinámicamente:

**Antes (sin slicers):**
```python
=COUNTIF(TB[Pregunta1], 5)
```

**Después (con slicers):**
```python
=COUNTIFS(TB[Pregunta1], 5, TB[_VISIBLE], 1)
```

La función `countif_visible()` genera automáticamente la fórmula correcta dependiendo de si hay filtros dinámicos o no.

### 3. **Conexión de Slicers**
Los slicers se conectan directamente a la tabla TB usando win32com:

```python
slicer_cache = wb.SlicerCaches.Add(
    Source=tabla_tb,
    SourceField=columna_filtro
)
slicer = slicer_cache.Slicers.Add(SlicerDestination=ws_tg)
```

## 🚀 Uso desde la Interfaz Streamlit

### Paso 1: Cargar Datos
1. Seleccionar método de procesamiento (Procesar/Pivotear)
2. Cargar archivo Excel (.xlsx)
3. Configurar oficina y proceso

### Paso 2: Seleccionar Columnas para Filtros
En la sección **"🎛️ Seleccionar columnas para filtros dinámicos"**:
- Aparece un multiselect con todas las columnas del DataFrame
- Seleccionar las columnas que se desean usar como filtros
- Ejemplos típicos:
  - `Programa Académico`
  - `Sede`
  - `Tipo de Estudiante`
  - `Semestre`
  - `Jornada`

### Paso 3: Generar Reporte
1. Presionar botón **"🚀 Ejecutar función excel_exportar"**
2. El sistema:
   - Genera el archivo Excel normalmente
   - Agrega columna `_VISIBLE` (oculta)
   - Modifica todas las fórmulas para considerar filtros
   - Crea slicers en la hoja T+G
   - Conecta slicers a la tabla TB

### Paso 4: Usar Slicers en Excel
1. Abrir el archivo generado
2. Ir a la hoja **T+G**
3. Los slicers aparecen en la parte superior izquierda
4. **Hacer clic en cualquier valor del slicer**
5. ✅ **TODAS las métricas se actualizan automáticamente:**
   - Porcentajes de satisfacción
   - Conteos absolutos
   - NPS (Net Promoter Score)
   - Correlaciones
   - Gráficos
   - Índices ponderados

## 📊 Métricas que Responden a Slicers

### ✅ Impacto Completo
Todas estas métricas se recalculan automáticamente al filtrar:

1. **Satisfacción General**
   - % por cada nivel (1-5)
   - Conteos absolutos
   - No Aplica

2. **Satisfacción por Pregunta**
   - Todas las preguntas individuales
   - Porcentajes y conteos

3. **Indicadores Calculados**
   - NIP (Nivel de Importancia Ponderado)
   - NSP (Nivel de Satisfacción Ponderado)
   - Peso de Correlaciones
   - ISC (Índice de Satisfacción del Cliente)

4. **Gráficos**
   - Los gráficos se basan en las fórmulas, por lo que también se actualizan

## 🔧 Archivos Modificados

### `Cargue.py`
**Cambios:**
- Agregada sección UI para selección de columnas de filtro
- Guardado de `columnas_filtros_dinamicos` en `session_state`
- Paso del parámetro a `excel_exportar()`
- Corrección de deprecaciones de Streamlit

**Líneas clave:**
```python
# Línea ~310: UI para selección de filtros
columnas_filtros_dinamicos = st.multiselect(
    "📊 Selecciona columnas para crear filtros desplegables en T+G:",
    options=df.columns.tolist()
)

# Línea ~335: Llamada a función con filtros
modulo.excel_exportar(..., filtros_dinamicos)
```

### `Generararchivoexcel_generico.py`
**Cambios principales:**

1. **Función helper (línea ~12-18):**
```python
def countif_visible(columna, criterio):
    if tiene_filtros:
        return f'COUNTIFS(TB[{columna}],{criterio},TB[_VISIBLE],1)'
    else:
        return f'COUNTIF(TB[{columna}],{criterio})'
```

2. **Columna auxiliar (línea ~45-65):**
```python
if tiene_filtros:
    Dijitacion.write(0, n_cols, "_VISIBLE")
    for row in range(1, n_rows + 1):
        Dijitacion.write_formula(row, n_cols, f'=SUBTOTAL(103,A{row+1})')
    Dijitacion.set_column(n_cols, n_cols, None, None, {'hidden': True})
```

3. **Fórmulas actualizadas (múltiples líneas):**
```python
# Antes
TG.write_formula(6, col, f'=COUNTIF(TB[{general}],5)')

# Después
TG.write_formula(6, col, f'={countif_visible(general, "5")}')
```

4. **Creación de slicers (línea ~1200-1280):**
```python
slicer_cache = wb.SlicerCaches.Add(
    Source=tabla_tb,
    SourceField=col_filtro
)
slicer = slicer_cache.Slicers.Add(SlicerDestination=ws_tg)
```

## 🐛 Problemas Resueltos

### 1. **COM Threading Error**
**Error:** `pywintypes.com_error: (-2147221008, 'No se ha llamado a CoInitialize.')`

**Solución:** Agregado `pythoncom.CoInitialize()` antes de usar win32com en entorno Streamlit multi-threaded.

### 2. **SlicerCaches.Add2() Fallaba**
**Error:** `(-2147352567, 'Ocurrió una excepción.', ..., -2147024809)`

**Solución:** Usar `SlicerCaches.Add()` con parámetros nombrados en lugar de `Add2()`.

### 3. **COUNTIF No Respetaba Filtros**
**Problema:** Las fórmulas `COUNTIF(TB[col], valor)` no cambiaban al usar slicers.

**Solución:** Usar `COUNTIFS` con columna `_VISIBLE` que usa `SUBTOTAL` para detectar filas visibles.

### 4. **Deprecaciones de Streamlit**
**Warnings:**
- `DataFrame.applymap` → `DataFrame.map`
- `use_container_width=True` → `width='stretch'`

**Solución:** Actualizado código para usar nuevas APIs.

## 📈 Ventajas del Sistema

### Para el Usuario Final
✅ **Un solo archivo** en lugar de decenas de reportes segmentados  
✅ **Análisis interactivo** sin necesidad de Excel avanzado  
✅ **Actualizaciones instantáneas** al cambiar filtros  
✅ **Múltiples dimensiones** de análisis simultáneas  
✅ **Estructura familiar** - el reporte se ve igual que siempre  

### Para el Equipo Técnico
✅ **Código mantenible** - función helper centralizada  
✅ **Retrocompatible** - funciona sin filtros también  
✅ **Escalable** - fácil agregar más columnas de filtro  
✅ **Sin duplicación** - una sola función genera todas las fórmulas  

## 🔄 Flujo Completo del Sistema

```
1. Usuario carga Excel en Streamlit
          ↓
2. Selecciona columnas para filtros
          ↓
3. Presiona "Ejecutar"
          ↓
4. Sistema genera Excel con xlsxwriter
          ↓
5. Sistema agrega columna _VISIBLE
          ↓
6. Fórmulas generadas con countif_visible()
          ↓
7. workbook.close() guarda el Excel
          ↓
8. win32com abre el Excel guardado
          ↓
9. Crea slicers conectados a tabla TB
          ↓
10. Guarda y cierra Excel
          ↓
11. Usuario descarga archivo final
          ↓
12. Al usar slicers: columna _VISIBLE cambia
          ↓
13. COUNTIFS evalúa _VISIBLE=1
          ↓
14. Todas las métricas se actualizan ✨
```

## 🎓 Ejemplo de Uso Real

### Caso: Encuesta de Satisfacción por Programa
**Columnas de filtro seleccionadas:**
- `Programa Académico`
- `Sede`

**Resultado:**
- Slicer 1: Lista todos los programas (Ing. Industrial, Medicina, Derecho, etc.)
- Slicer 2: Lista todas las sedes (Barranquilla, Soledad, etc.)

**Análisis posible:**
1. Seleccionar "Ing. Industrial" → Ver satisfacción del programa
2. Agregar "Barranquilla" → Ver satisfacción de Ing. Industrial en Barranquilla
3. Cambiar a "Medicina" → Satisfacción cambia instantáneamente
4. Borrar filtros → Volver a vista completa

## 📝 Notas Importantes

### Limitaciones
- Solo funciona en "Oficina Genérica / Personalizada"
- Requiere pywin32 instalado
- Solo funciona en Windows (win32com)
- Excel debe estar instalado en el sistema

### Rendimiento
- La columna `_VISIBLE` usa `SUBTOTAL` que es eficiente
- Los slicers son nativos de Excel, rendimiento óptimo
- Sin impacto en velocidad de generación del reporte

### Compatibilidad
- Excel 2010 o superior
- Funciona con cualquier cantidad de columnas de filtro
- Compatible con todos los formatos actuales del sistema

## 🚦 Estado del Proyecto

### ✅ Completado
- [x] Implementación de columna `_VISIBLE`
- [x] Función `countif_visible()` 
- [x] Actualización de fórmulas en T+G
- [x] Creación automática de slicers
- [x] Conexión de slicers a tabla TB
- [x] UI en Streamlit para selección
- [x] Corrección de bugs COM
- [x] Documentación completa

### 🎯 Próximas Mejoras Potenciales
- [ ] Extender a otras oficinas específicas
- [ ] Posicionamiento personalizado de slicers
- [ ] Estilos personalizados para slicers
- [ ] Guardar preferencias de columnas de filtro

## 👥 Créditos
**Desarrollado por:** Equipo de Automatización - Universidad del Norte  
**Fecha:** Diciembre 2025  
**Versión:** 2.0 con Slicers Dinámicos

---

Para preguntas o soporte, contactar al equipo de desarrollo.
