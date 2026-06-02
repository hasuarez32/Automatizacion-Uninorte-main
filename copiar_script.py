import shutil

# Leer el archivo de admisiones
with open('Script de los formatos/Generararchivoexcel_admisiones_posgrado.py', 'r', encoding='utf-8') as f:
    contenido = f.read()

# Reemplazar la primera línea (firma de la función)
contenido_modificado = contenido.replace(
    'def excel_exportar(data, nombre_archivo,numerodepoblacion, Preguntas,columnas_observaciones,general,oficina,proceso, perido,tipos_grafica):',
    'def excel_exportar(data, nombre_archivo,numerodepoblacion, Preguntas,columnas_observaciones,general,oficina,proceso, perido,tipos_grafica, columnas_filtros_dinamicos=[]):'
)

# Guardar como generico
with open('Script de los formatos/Generararchivoexcel_generico.py', 'w', encoding='utf-8') as f:
    f.write(contenido_modificado)

print("✅ Archivo genérico creado exitosamente")
