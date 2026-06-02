import streamlit as st

st.set_page_config(page_title="📘 Instructivo de Uso", layout="wide")

st.markdown("""
<style>
.guide-title {
    font-size: 32px;
    color: #2f3e75;
    font-weight: 700;
    text-align: center;
    background: linear-gradient(to right, #f0f4fc, #e1e8f7);
    padding: 18px;
    margin: 30px auto 20px auto;
    border-radius: 14px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    border: 1px solid #d3ddf0;
    max-width: 800px;
}
.guide-subtitle {
    font-size: 20px;
    color: #3b4d80;
    margin-top: 22px;
    font-weight: 600;
    border-left: 4px solid #3b4d80;
    padding-left: 12px;
}
ul {
    padding-left: 20px;
    margin-bottom: 16px;
}
ul li {
    padding: 6px 0;
    font-size: 16px;
    line-height: 1.5;
}
table.example-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
    margin-bottom: 25px;
}
table.example-table th, table.example-table td {
    border: 1px solid #ccc;
    padding: 10px;
    text-align: center;
    font-size: 14px;
}
table.example-table th {
    background-color: #e3e7f1;
}
</style>
""", unsafe_allow_html=True)

with st.container():

    st.markdown("<div class='guide-title'>🧾 Instructivo de Uso de la Herramienta</div>", unsafe_allow_html=True)

    st.markdown("<div class='guide-subtitle'>🏢 Selección de Oficina y Proceso</div>", unsafe_allow_html=True)
    st.markdown("""
- Usa los menús desplegables para elegir correctamente:
  - La **oficina responsable del proceso**.
  - El **proceso específico** dentro de dicha oficina.
- Esta selección carga automáticamente el script de procesamiento asociado para generar el informe correcto.
    """)

    st.markdown("<div class='guide-subtitle'>📤 Carga del Archivo</div>", unsafe_allow_html=True)
    st.markdown("""
- Carga un archivo Excel (`.xlsx`) que contenga los resultados de la encuesta aplicada.
- El archivo debe tener una estructura clara, con preguntas como títulos de columna.
- Valores válidos para las preguntas deben ser: `1`, `2`, `3`, `4`, `5`, `No Aplica`.
""")

    st.markdown("""
<strong>Ejemplo de tabla esperada:</strong>
<table class='example-table'>
    <tr>
        <th>Fecha</th>
        <th>Nombre</th>
        <th>¿Está satisfecho con el servicio?</th>
        <th>¿Recomendaría el servicio?</th>
        <th>Satisfacción General</th>
    </tr>
    <tr>
        <td>2025-06-15</td>
        <td>Juan Pérez</td>
        <td>5</td>
        <td>4</td>
        <td>5</td>
    </tr>
    <tr>
        <td>2025-06-16</td>
        <td>Ana Torres</td>
        <td>3</td>
        <td>3</td>
        <td>4</td>
    </tr>
</table>
<p>- Si tu archivo contiene errores ortográficos, saltos de línea o valores no válidos, <strong>debes limpiarlo antes de continuar</strong>.</p>
""", unsafe_allow_html=True)

    st.markdown("<div class='guide-subtitle'>🔍 Aplicación de Filtros (Opcional)</div>", unsafe_allow_html=True)
    st.markdown("""
- Algunos procesos permiten aplicar filtros personalizados para mostrar solo ciertos registros.
- Por ejemplo, puedes filtrar por programa, jornada, sede, o ciclo.
    """)

    st.markdown("<div class='guide-subtitle'>📌 Selección de Preguntas</div>", unsafe_allow_html=True)
    st.markdown("""
- El sistema detecta automáticamente las columnas con valores válidos como posibles preguntas.
- Puedes seleccionar otras columnas manualmente si es necesario.
- También debes indicar qué columna corresponde a la **pregunta de satisfacción general**, si aplica.
    """)

    st.markdown("<div class='guide-subtitle'>💬 Comentarios y Observaciones</div>", unsafe_allow_html=True)
    st.markdown("""
- Se detectan automáticamente columnas con términos como `comentario`, `sugerencia`, `observación`.
- Estas columnas pueden ser seleccionadas para su inclusión en el informe exportado.
    """)

    st.markdown("<div class='guide-subtitle'>📝 Datos Complementarios</div>", unsafe_allow_html=True)
    st.markdown("""
- Asigna el nombre del archivo de salida (sin extensión `.xlsx`).
- Ingresa el número de personas que conforman la población total.
- Especifica el período en que fue aplicada la encuesta (ej. `2025-1`).
    """)

    st.markdown("<div class='guide-subtitle'>📊 Visualización Opcional</div>", unsafe_allow_html=True)
    st.markdown("""
- Puedes elegir preguntas específicas para visualizar en gráficos.
- Tipos de visualización disponibles: `pie` o `column`.
- Esto no afecta el archivo exportado, solo es útil para análisis exploratorio.
    """)

    st.markdown("<div class='guide-subtitle'>🚀 Ejecución y Descarga</div>", unsafe_allow_html=True)
    st.markdown("""
- Presiona el botón "Ejecutar función excel_exportar".
- El sistema procesará los datos y generará un archivo Excel con la estructura adecuada.
- Si todo va bien, se mostrará un botón de descarga del archivo generado.
    """)

    st.markdown("<div class='guide-subtitle'>⚠️ Recomendaciones Finales</div>", unsafe_allow_html=True)
    st.markdown("""
- Verifica que las columnas no tengan valores vacíos o mal escritos.
- Asegúrate de que el archivo esté limpio y en el formato correcto antes de subirlo.
- Si hay errores, realiza la limpieza en Excel o una herramienta de tu preferencia.
    """)

    st.markdown("</div>", unsafe_allow_html=True)
