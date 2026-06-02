# 📘 Instructivo: Sistema de Slicers Dinámicos en Reportes Excel

## 📋 Instalación en otro equipo

### **Requisitos previos:**
- ✅ Python 3.8 o superior instalado
- ✅ Windows (para funcionalidad completa de slicers)
- ✅ Microsoft Excel instalado (2010 o superior)

### **Pasos de instalación:**

1. **Copiar la carpeta completa** del proyecto a la nueva ubicación

2. **Abrir terminal/PowerShell** en la carpeta del proyecto:
   ```bash
   cd ruta\a\Automatizacion
   ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```
   
   Esto instalará todas las librerías necesarias:
   - streamlit, pandas, xlsxwriter, openpyxl
   - pywin32 (para slicers), numpy, matplotlib
   - seaborn, st-aggrid, textwrap3

4. **Ejecutar la aplicación:**
   
   Prueba estos comandos en orden hasta que funcione:
   
   **Opción 1 (más común en Windows):**
   ```bash
   py -m streamlit run Cargue.py
   ```
   
   **Opción 2:**
   ```bash
   python -m streamlit run Cargue.py
   ```
   
   **Opción 3:**
   ```bash
   streamlit run Cargue.py
   ```
   
   **Nota:** Depende de cómo esté configurado Python en tu sistema. Usa el que funcione.

5. **Acceder a la app:**
   - Se abrirá automáticamente en tu navegador
   - URL: `http://localhost:8501`

### **Verificación:**
- ✅ Si ves la interfaz con "Exportador de Excel", la instalación fue exitosa
- ✅ Puedes cargar archivos y generar reportes
- ✅ Si estás en Windows con Excel, los slicers funcionarán

---

## 🎯 ¿Qué son los Slicers Dinámicos?

Los **slicers** (o segmentadores de datos) son botones de filtro interactivos en Excel que permiten:
- Filtrar datos con un solo clic
- Ver resultados inmediatos sin modificar fórmulas
- Analizar diferentes segmentos sin generar múltiples reportes
- Compartir un solo archivo con capacidad de exploración interactiva

**Ejemplo:** Un reporte con slicers de "Programa" y "Sede" permite ver la satisfacción de:
- Ingeniería Industrial en Barranquilla
- Medicina en todas las sedes
- Todos los programas en Soledad
- Y cualquier combinación con solo hacer clic

---

## 🚀 Cómo usar la funcionalidad

### **Paso 1: Cargar el archivo de encuesta**
1. Abre la aplicación Streamlit
2. En la sección "1️⃣ Cargar archivo Excel", sube tu archivo de respuestas
3. Selecciona el método de procesamiento (Procesar o Pivotear)

### **Paso 2: Configurar parámetros básicos**
1. Selecciona la oficina
2. Selecciona el proceso asociado
3. Escribe el nombre del archivo de salida
4. Ingresa el número de población
5. Escribe el periodo de la encuesta

### **Paso 3: Configurar columnas**
1. En "Vista previa", verás todas las columnas de tu archivo
2. Selecciona las preguntas que quieres analizar
3. Selecciona las columnas de observaciones
4. Elige la columna general de satisfacción
5. Selecciona los tipos de gráfica

### **Paso 4: ⭐ Activar Slicers (NUEVO)**
1. En la sección **"📊 Selecciona columnas para crear filtros desplegables en T+G"**
2. Marca las columnas que quieres usar como filtros interactivos
3. Ejemplos comunes:
   - Programa académico
   - Sede
   - Género
   - Tipo de estudiante
   - Semestre
   - Jornada
   - Cualquier columna categórica de tu encuesta

**⚠️ Nota:** Puedes seleccionar 1, 5, 10 o más columnas. No hay límite técnico.

### **Paso 5: Generar el reporte**
1. Haz clic en **"🚀 Ejecutar generación"**
2. Espera a que se genere el archivo
3. Verás un mensaje de éxito y un botón de descarga
4. Descarga el archivo Excel

### **Paso 6: Usar los slicers en Excel**
1. Abre el archivo descargado en Excel
2. Ve a la hoja **"T+G"** (Tablas y Gráficos)
3. Verás los slicers a la izquierda de la hoja
4. Haz clic en cualquier valor para filtrar:
   - **Un solo valor:** Clic en el valor
   - **Múltiples valores:** Mantén Ctrl + clic en varios valores
   - **Limpiar filtro:** Clic en el ícono de funnel con X
5. Todos los cálculos, porcentajes, gráficos y métricas se actualizan automáticamente

---

## 🔧 Cómo funciona técnicamente

### **Arquitectura del sistema**

```
Usuario selecciona columnas → App guarda en session_state → 
Pasa a script de oficina → Script detecta si hay filtros →
Genera Excel con columna _VISIBLE → Crea slicers con pywin32 →
Usuario hace clic en slicer → Excel filtra automáticamente
```

### **Componentes clave**

#### **1. Columna auxiliar `_VISIBLE`**
- Se agrega automáticamente a la tabla de datos (hoja "Digitación")
- Usa la fórmula: `=SUBTOTAL(103,A2)`
- Devuelve `1` si la fila está visible, `0` si está oculta por un filtro
- Se oculta automáticamente para no afectar la visualización

#### **2. Función helper `countif_visible()`**
Cada script tiene esta función que genera fórmulas dinámicas:

```python
def countif_visible(columna, criterio):
    if tiene_filtros:
        return f'COUNTIFS(TB[{columna}],{criterio},TB[_VISIBLE],1)'
    else:
        return f'COUNTIF(TB[{columna}],{criterio})'
```

**Sin filtros:**
```excel
=COUNTIF(TB[Pregunta1], 5)
```

**Con filtros:**
```excel
=COUNTIFS(TB[Pregunta1], 5, TB[_VISIBLE], 1)
```

#### **3. Variable `tiene_filtros`**
```python
tiene_filtros = len(columnas_filtros_dinamicos) > 0
```
- `True` si el usuario seleccionó columnas → genera fórmulas especiales
- `False` si no hay columnas → funciona como antes (retrocompatible)

#### **4. Creación de slicers con pywin32**
Al final del proceso, el script:

```python
import win32com.client
import pythoncom

# Inicializar COM
pythoncom.CoInitialize()
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(ruta_archivo)

# Crear cada slicer
for idx, columna in enumerate(columnas_filtros_dinamicos):
    slicer_cache = wb.SlicerCaches.Add(
        Source=tabla_tb,
        SourceField=columna
    )
    slicer = slicer_cache.Slicers.Add(SlicerDestination=ws_tg)
    slicer.Top = 50 + (idx * 220)  # Apilar verticalmente
    slicer.Left = 50
    slicer.Height = 200
    slicer.Width = 250

wb.Save()
wb.Close()
excel.Quit()
```

---

## 📁 Archivos modificados

### **1. Cargue.py** (líneas ~280, ~370)
**Cambios:**
- Agregado `st.multiselect()` para selección de columnas de filtro
- Guardado en `session_state["columnas_filtros_dinamicos"]`
- Paso del parámetro a todas las oficinas

### **2. Todos los scripts de oficinas** (19 archivos)
**Cambios en cada uno:**
- Parámetro `columnas_filtros_dinamicos=[]` en `excel_exportar()`
- Variable `tiene_filtros` (línea ~18)
- Función `countif_visible()` (líneas ~20-26)
- Columna `_VISIBLE` (líneas ~45-65)
- Fórmulas actualizadas usando `countif_visible()` (múltiples líneas)
- Código de creación de slicers (líneas ~1265-1320 en generico.py)

**Scripts modificados:**
- Generararchivoexcel_generico.py
- Generararchivoexcel_admisiones_posgrado.py
- Generararchivoexcel_Tesoreria.py
- Generararchivoexcel_Almacen.py
- Generararchivoexcel_Adquisicion_bienes.py
- Generararchivoexcel_certificaciones.py
- Generararchivoexcel_coordinadores.py
- Generararchivoexcel_Financiamiento_Empresarial.py
- Generararchivoexcel_laboratori_cimm.py
- Generararchivoexcel_laboratorio_geotecnia.py
- Generararchivoexcel_mantenimientoDSA.py
- Generararchivoexcel_mantenimiento_tic_CSU.py
- Generararchivoexcel_mantenimiento_tic_trimestre.py
- Generararchivoexcel_movilidad_entrante.py
- Generararchivoexcel_oficinaregistro_grado.py
- Generararchivoexcel_planeacion.py
- Generararchivoexcel_prueba.py
- Generararchivoexcel_registro_provedores.py
- Generararchivoexcel_servicio_transporte_seguridad_Aseo.py

### **3. requirements.txt**
**Agregado:**
```
pywin32==306
```

---

## ⚠️ Requisitos y limitaciones

### **Requisitos del sistema:**
- ✅ **Windows** (win32com solo funciona en Windows)
- ✅ **Microsoft Excel instalado** (2010 o superior)
- ✅ **Python 3.8+**
- ✅ **pywin32 instalado** (`pip install pywin32==306`)

### **Limitaciones:**
- ❌ No funciona en Mac/Linux (limitación de pywin32)
- ⚠️ Si creas muchos slicers (10+), necesitarás hacer scroll en Excel para verlos todos
- ✅ Sin límite en cantidad de columnas de filtro
- ✅ Sin impacto en rendimiento

### **Compatibilidad:**
- ✅ Funciona en todas las 19 oficinas
- ✅ Retrocompatible (si no seleccionas columnas, funciona como antes)
- ✅ Compatible con todos los tipos de gráficos existentes

---

## 🐛 Solución de problemas

### **Problema: Los slicers no aparecen en Excel**
**Posibles causas:**
1. No seleccionaste columnas en el paso 4
2. pywin32 no está instalado
3. Excel no está instalado en el sistema
4. Estás en Mac/Linux

**Solución:**
- Verifica que hayas seleccionado columnas antes de generar
- Ejecuta: `pip install pywin32==306`
- Asegúrate de estar en Windows con Excel instalado

### **Problema: Error #REF! en las fórmulas**
**Causa:** Este error ya fue corregido en el commit `a7c53e7`

**Solución:**
- Asegúrate de tener la última versión del código
- Ejecuta: `git pull origin main`

### **Problema: Las fórmulas no actualizan al filtrar**
**Causa:** La columna `_VISIBLE` no se creó correctamente

**Solución:**
- Verifica que seleccionaste columnas de filtro
- Regenera el reporte
- Revisa que la variable `tiene_filtros = True`

### **Problema: Error "module 'win32com' not found"**
**Solución:**
```bash
pip install pywin32==306
```

---

## 📊 Qué se actualiza con los slicers

Cuando aplicas un filtro con los slicers, se actualizan automáticamente:

✅ **Ficha Técnica:**
- Muestra alcanzada (G11)
- Todos los conteos y porcentajes

✅ **Tabla General de Satisfacción:**
- Porcentajes de cada respuesta (Muy satisfecho, Satisfecho, etc.)
- Conteos absolutos
- Total de respuestas

✅ **Indicadores Calculados:**
- NIP (Nivel de Importancia Ponderado)
- NSP (Nivel de Satisfacción Ponderado)
- Peso de Correlaciones
- ISC (Índice de Satisfacción del Cliente)

✅ **Gráficos:**
- Gráfico general de satisfacción
- Gráficos por pregunta individual
- Gráfico de importancia vs satisfacción

✅ **Preguntas Individuales:**
- Todas las tablas de frecuencia
- Todos los porcentajes
- Todos los gráficos asociados

---

## 🎓 Ejemplos de uso

### **Caso 1: Análisis por Programa**
**Objetivo:** Ver la satisfacción de cada programa académico

**Pasos:**
1. Selecciona columna "Programa" en los slicers
2. Genera el reporte
3. En Excel, haz clic en "Ingeniería Industrial"
4. Resultado: Ves solo la satisfacción de ese programa
5. Cambia a "Medicina" → Todo se actualiza instantáneamente

### **Caso 2: Análisis Multidimensional**
**Objetivo:** Ver satisfacción de estudiantes de pregrado en Barranquilla

**Pasos:**
1. Selecciona columnas "Tipo_estudiante" y "Sede" en los slicers
2. Genera el reporte
3. En Excel:
   - Clic en "Pregrado" en slicer de Tipo_estudiante
   - Clic en "Barranquilla" en slicer de Sede
4. Resultado: Ves solo pregrado de Barranquilla
5. Agrega "Jornada: Diurna" → Se filtra aún más

### **Caso 3: Comparar Sedes**
**Objetivo:** Comparar satisfacción entre sedes

**Pasos:**
1. Selecciona columna "Sede" en los slicers
2. Genera el reporte
3. Anota el ISC para "Barranquilla"
4. Cambia a "Soledad" → Anota el ISC
5. Compara los resultados sin necesidad de generar 2 reportes

---

## 📝 Commits principales

**Historial de desarrollo:**

1. **`3d589eb`** - Implementar slicers dinámicos funcionales con filtrado reactivo
2. **`f99bdc3`** - Corregir ficha técnica y porcentajes para que sean dinámicos con slicers
3. **`dae140d`** - Extender filtros dinámicos a TODAS las oficinas
4. **`369eb9f`** - Agregar pywin32 a requirements.txt
5. **`0a394f1`**, **`3983d05`**, **`483733a`** - Corregir errores en admisiones_posgrado
6. **`fe9b512`** - Corregir referencias #REF! en formulas de admisiones_posgrado
7. **`a7c53e7`** - Corregir referencias #REF! en las 17 oficinas restantes

**Total de trabajo:** ~3 días, 3800+ líneas de código modificadas

---

## 📞 Soporte

**Si tienes dudas o problemas:**
1. Revisa la sección "Solución de problemas"
2. Verifica que tienes la última versión: `git pull origin main`
3. Consulta la documentación técnica en `SLICERS_DINAMICOS.md`
4. Contacta al equipo de desarrollo

---

## ✨ Ventajas del sistema

- 📊 **Un reporte reemplaza docenas** - Ya no necesitas generar un archivo por cada segmento
- ⚡ **Análisis instantáneo** - Cambias el filtro y ves resultados en milisegundos
- 🎯 **Sin errores humanos** - Las fórmulas se actualizan automáticamente
- 💾 **Menor almacenamiento** - Un archivo vs. 20+ archivos estáticos
- 🔄 **Fácil de compartir** - Envías un solo archivo con toda la capacidad de análisis
- 📈 **Exploración libre** - El usuario final puede analizar sin necesidad de conocimientos técnicos

---

**Desarrollado por:** Equipo de Automatización - Universidad del Norte  
**Fecha:** Diciembre 2025  
**Versión:** 2.0 con Slicers Dinámicos

