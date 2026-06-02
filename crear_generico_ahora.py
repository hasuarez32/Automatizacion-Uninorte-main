# Script para crear el archivo genérico
import os

# Leer el archivo de admisiones
ruta_origen = r'Script de los formatos\Generararchivoexcel_admisiones_posgrado.py'
ruta_destino = r'Script de los formatos\Generararchivoexcel_generico.py'

with open(ruta_origen, 'r', encoding='utf-8') as f:
    contenido = f.read()

# Reemplazar la primera línea (firma de la función)
contenido_modificado = contenido.replace(
    'def excel_exportar(data, nombre_archivo,numerodepoblacion, Preguntas,columnas_observaciones,general,oficina,proceso, perido,tipos_grafica):',
    'def excel_exportar(data, nombre_archivo,numerodepoblacion, Preguntas,columnas_observaciones,general,oficina,proceso, perido,tipos_grafica, columnas_filtros_dinamicos=[]):'
)

# Guardar como generico
with open(ruta_destino, 'w', encoding='utf-8') as f:
    f.write(contenido_modificado)

print(f'✅ Archivo genérico creado exitosamente')
print(f'📄 Archivo origen: {os.path.getsize(ruta_origen)} bytes')
print(f'📄 Archivo destino: {os.path.getsize(ruta_destino)} bytes')
