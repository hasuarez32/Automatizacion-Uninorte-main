import pandas as pd
import sys
import os

# Cargar el archivo de Qualtrics
archivo = r"C:\Users\ecpereira\Desktop\Automatizacion\Calidad_Biblioteca_Estudiantes_Profesores_V2025_24+de+octubre+de+2025_08.00 (2).xlsx"

# Verificar que el archivo existe
if not os.path.exists(archivo):
    print(f"ERROR: No se encuentra el archivo: {archivo}")
    sys.exit(1)

print("=" * 80)
print("DIAGNÓSTICO DE PROCESAMIENTO QUALTRICS")
print("=" * 80)

# Leer archivo completo sin procesar
df_raw = pd.read_excel(archivo, header=None)
print(f"\n1. ARCHIVO RAW:")
print(f"   Total de filas (incluyendo encabezados): {len(df_raw)}")
print(f"   Total de columnas: {len(df_raw.columns)}")
print(f"\n   Primeras 3 filas:")
for i in range(min(3, len(df_raw))):
    print(f"   Fila {i}: {df_raw.iloc[i, 0]} | {df_raw.iloc[i, 1]} | {df_raw.iloc[i, 2]}")

# Simular procesamiento Qualtrics (CORREGIDO)
df = df_raw.copy()
df.columns = df.iloc[0]
# Eliminar ambas filas de encabezado (0 y 1)
df = df[2:].reset_index(drop=True)

print(f"\n2. DESPUÉS DE USAR FILA 0 COMO ENCABEZADO:")
print(f"   Total de filas de datos: {len(df)}")
print(f"   Nombres de primeras 5 columnas: {list(df.columns[:5])}")

# Manejar columnas duplicadas
cols = pd.Series(df.columns)
duplicados = cols[cols.duplicated()].unique()
print(f"\n3. COLUMNAS DUPLICADAS ENCONTRADAS: {len(duplicados)}")
if len(duplicados) > 0:
    print(f"   Ejemplos: {list(duplicados[:5])}")

for dup in duplicados:
    cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
df.columns = cols

# Limpiar espacios
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

# Buscar columnas que podrían ser la "pregunta general"
print(f"\n4. BUSCANDO COLUMNA DE PREGUNTA GENERAL:")
posibles_generales = []
for col in df.columns:
    col_str = str(col).lower()
    if 'general' in col_str or 'satisfacción' in col_str or 'satisfaccion' in col_str:
        posibles_generales.append(col)
        print(f"   Columna encontrada: '{col}'")
        # Contar valores antes de reemplazos
        print(f"   Valores únicos: {df[col].value_counts().to_dict()}")

# Aplicar reemplazos
reemplazos_qualtrics = {
    '5- Totalmente satisfecho': 5,
    '4- Satisfecho': 4,
    '3- Neutral': 3,
    '2- Insatisfecho': 2,
    '1-Totalmente insatisfecho': 1,
    '5- Estás totalmente Satisfecho': 5,
    '5- Estás totalmente satisfecho': 5,
    '4- Estás satisfecho': 4,
    '3- Estás neutral': 3,
    '2- Estás insatisfecho': 2,
    '1- Estás totalmente insatisfecho': 1,
    '1-Estás totalmente insatisfecho': 1,
    '5 - Muy satisfecho': 5,
    '4 - Satisfecho': 4,
    '3 - Neutral': 3,
    '2 - Insatisfecho': 2,
    '1 - Muy insatisfecho': 1,
    'Supera notablemente las expectativas': 5,
    '5. Supera notablemente las expectativas': 5,
    '5 (Supera las expectativas)': 5,
    'Cumple las expectativas': 3,
    'Por debajo de las expectativas': 2,
    'Muy por debajo de las expectativas': 1,
    '5': 5,
    '4': 4,
    '3': 3,
    '2': 2,
    '1': 1,
    'Mucho': 4,
    'Algo': 3,
    'Poco': 2,
    'Nada': 1
}

df = df.replace(reemplazos_qualtrics).infer_objects(copy=False)

print(f"\n5. DESPUÉS DE REEMPLAZOS (PREGUNTA GENERAL):")
for col in posibles_generales:
    if col in df.columns:
        conteo = df[col].value_counts().sort_index()
        print(f"   Columna: '{col}'")
        print(f"   Conteo de valores:")
        for val, count in conteo.items():
            print(f"      {val}: {count}")

# Verificar si hay filas que parecen ser de prueba o metadata
print(f"\n6. VERIFICACIÓN DE FILAS SOSPECHOSAS:")
print(f"   Revisando primeras 5 filas para detectar patrones de prueba/metadata...")
for i in range(min(5, len(df))):
    fila = df.iloc[i]
    # Contar cuántos valores son NaN o vacíos
    vacios = fila.isna().sum()
    total_cols = len(fila)
    porcentaje_vacios = (vacios / total_cols) * 100
    print(f"   Fila {i}: {porcentaje_vacios:.1f}% vacíos ({vacios}/{total_cols})")
    
    # Mostrar algunas columnas clave si existen
    if 'StartDate' in df.columns:
        print(f"      StartDate: {df.iloc[i]['StartDate']}")
    if 'EndDate' in df.columns:
        print(f"      EndDate: {df.iloc[i]['EndDate']}")
    if len(posibles_generales) > 0 and posibles_generales[0] in df.columns:
        print(f"      Pregunta general: {df.iloc[i][posibles_generales[0]]}")

print(f"\n7. RESUMEN FINAL:")
print(f"   Total de filas procesadas: {len(df)}")
print(f"   Total de columnas: {len(df.columns)}")
print(f"   Columnas con duplicados renombradas: {len(duplicados)}")

print("\n" + "=" * 80)
