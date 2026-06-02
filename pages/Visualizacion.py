import streamlit as st
import matplotlib.pyplot as plt
import textwrap
import numpy as np
import pandas as pd
import seaborn as sns
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
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
</style>
""", unsafe_allow_html=True)
st.set_page_config(page_title="Visualización de Resultados", layout="wide")
st.markdown("<div class='guide-title'>📊 Vista previa informe</div>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; height: 1px; background: linear-gradient(to right, #ccc, #eee, #ccc); margin: 120px 0;'>",
                     unsafe_allow_html=True)
if "df_encuesta" in st.session_state:
    df = st.session_state["df_encuesta"]
    columna_general = st.session_state["nombre_columna_general"]

    oficina_seleccionada=st.session_state["oficina_seleccionada"]
    proceso_seleccionado=st.session_state["proceso_seleccionado"]
    oficina_filtros=["Almacen","Prueba","Sección de Servicios Generales","Registro"]
    proceso_filtros=["Entrega de Insumos y Compras Directas - Entrega de Activos","Prueba grafica","Servicio de transporte","Seguridad en Uninorte","Servicios de Aseo","Certificaciones académicas por vía Web"]
   
    if oficina_seleccionada in oficina_filtros and proceso_seleccionado in proceso_filtros:
        if df is not None and not df.empty:
            columnas_filtrar = st.sidebar.multiselect("Selecciona las columnas para aplicar filtros:", options=df.columns.tolist())

            for columna in columnas_filtrar:
                valores_unicos = df[columna].dropna().unique().tolist()
                valores_seleccionados = st.sidebar.multiselect(f"Selecciona los valores para '{columna}':", options=valores_unicos)

                if valores_seleccionados:
                    df = df[df[columna].isin(valores_seleccionados)]
    posibles_valores = {"1", "2", "3", "4", "5", "No Aplica"}
    columnas_preguntas = []
    if columna_general in df.columns:
        conteo = df[columna_general].astype(str).value_counts()

        n1 = conteo.get("1", 0)
        n2 = conteo.get("2", 0)
        n3 = conteo.get("3", 0)
        n4 = conteo.get("4", 0)
        n5 = conteo.get("5", 0)
        total = n1 + n2 + n3 + n4 + n5

        if total > 0:
            nsp_100 = (100*n5 + 80*n4 + 60*n3 + 40*n2 + 20*n1) / total
            nsp_5 = nsp_100 / 20

            st.markdown(f"""
            <style>
            .tabla-satisfaccion {{
                border-collapse: collapse;
                width: 50%;
                margin: 30px auto;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
            }}
            .tabla-satisfaccion th, .tabla-satisfaccion td {{
                border: 1px solid black;
                padding: 10px;
            }}
            .tabla-satisfaccion th {{
                background-color: #e0e0e0;
                text-align: center;
                font-weight: bold;
            }}
            .tabla-satisfaccion td {{
                background-color: #d9e9c1;
            }}
            .tabla-satisfaccion td:nth-child(2) {{
                text-align: right;
                font-weight: bold;
            }}
            </style>

            <table class="tabla-satisfaccion">
                <thead>
                    <tr><th colspan="2">CÁLCULO NIVEL DE SATISFACCIÓN GENERAL</th></tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Nivel de Satisfacción General (Escala 0:100)</td>
                        <td>{nsp_100:.2f}</td>
                    </tr>
                    <tr>
                        <td>Nivel de Satisfacción General (Escala 0:5)</td>
                        <td>{nsp_5:.2f}</td>
                    </tr>
                </tbody>
            </table>
            """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ No hay suficientes datos en la columna de satisfacción general.")
    else:
        st.warning("❌ No se encontró la columna de satisfacción general en el DataFrame.")
    columnas_seleccionadas=st.session_state["columnas_seleccionadas"]
    st.markdown("<hr style='border: none; height: 1px; background: linear-gradient(to right, #ccc, #eee, #ccc); margin: 120px 0;'>",
                     unsafe_allow_html=True)
    # Visualización por pregunta
    if columnas_seleccionadas:
        col1, col2 = st.columns(2)

        for i, pregunta in enumerate(columnas_seleccionadas):
            data = df[pregunta].dropna().astype(str)
            conteo = data.value_counts()
            orden = ["5", "4", "3", "2", "1", "No Aplica"]
            conteo = conteo.reindex(orden, fill_value=0)

            n_valores = sum([conteo.get(v, 0) for v in ["1", "2", "3", "4", "5"]])
            n_total = n_valores + conteo.get("No Aplica", 0)

            porcentajes = []
            for opcion in orden:
                valor = conteo.get(opcion, 0)
                if opcion == "No Aplica" and n_total > 0:
                    pct = (valor / n_total) * 100
                elif opcion != "No Aplica" and n_valores > 0:
                    pct = (valor / n_valores) * 100
                else:
                    pct = 0
                porcentajes.append(pct)

            # Título ajustado por largo
            titulo_wrap = "\n".join(textwrap.wrap(pregunta, width=40))
 
            # 🎯 Gráfico más pequeño
            fig, ax = plt.subplots(figsize=(3, 1.7))
            colores = ["#0f6df1", "#0f6df1", "#0f6df1", "#0f6df1", "#0f6df1", "#0f6df1"]
            bars = ax.bar(orden, porcentajes, color=colores)
            
            for bar, pct in zip(bars, porcentajes):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2, height + 1, f"{pct:.1f}%", ha='center', va='bottom', fontsize=5)

            ax.set_title(titulo_wrap, fontsize=6, pad=4)
            ax.set_ylim(0, 110)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_visible(False)
            ax.spines["bottom"].set_color("#ccc")
            ax.tick_params(left=False, labelleft=False, bottom=False, labelbottom=True)
            ax.set_xticks(range(len(orden)))
            ax.set_xticklabels(orden, fontsize=5)
            ax.set_xlabel("")  # Eliminamos el eje X
            fig.tight_layout()

            if i % 2 == 0:
                col1.pyplot(fig, clear_figure=True, bbox_inches="tight")

            else:
                col2.pyplot(fig, clear_figure=True, bbox_inches="tight")
        st.markdown(
                    "<hr style='border: none; height: 1px; background: linear-gradient(to right, #ccc, #eee, #ccc); margin: 120px 0;'>",
                     unsafe_allow_html=True)



st.markdown("### 📋 Tabla de correlaciones y pesos (NIP)")

if "df_encuesta" in st.session_state and "nombre_columna_general" in st.session_state:
    columna_general = st.session_state["nombre_columna_general"]
    columnas_seleccionadas=st.session_state["columnas_seleccionadas"]
    if columna_general not in df.columns:
        st.warning("❌ No se encontró la columna de satisfacción general.")
    else:
        preguntas_sin_general = [col for col in columnas_seleccionadas if col != columna_general]
        correlaciones = []
        pesos = []

        suma_corr = 0
        porcentaje_respuestas = {}

        for pregunta in preguntas_sin_general:
            try:
                serie_pregunta = pd.to_numeric(df[pregunta], errors="coerce")
                serie_general = pd.to_numeric(df[columna_general], errors="coerce")

                # Combinar y eliminar filas con NaN
                validas = pd.concat([serie_pregunta, serie_general], axis=1).dropna()

                if not validas.empty:
                    corr = abs(validas.iloc[:, 0].corr(validas.iloc[:, 1]))
                else:
                    corr = 0.0
            except Exception:
                corr = 0.0
            correlaciones.append(corr)
            suma_corr += corr

            # Calcular porcentaje de respuestas por valor
            conteo = df[pregunta].value_counts(normalize=True)
            porcentajes = {
                str(i): conteo.get(i, 0.0)
                for i in ["5", "4", "3", "2", "1"]
            }
            porcentaje_respuestas[pregunta] = porcentajes

        for c in correlaciones:
            peso = c / suma_corr if suma_corr > 0 else 0.0
            pesos.append(peso)

        pnips = [(peso / (np.median(pesos))*100 if np.median(pesos) > 0 else 0) for peso in pesos]
        
        nsps_100 = []
        nsps_5 = []

        for pregunta in preguntas_sin_general:
            try:
                # Convertir a string por si hay números o mezcla
                valores_validos = df[pregunta].astype(str).isin(["1", "2", "3", "4", "5"])
                serie_filtrada = df.loc[valores_validos, pregunta].astype(str)

                conteo = serie_filtrada.value_counts(normalize=True)

                score_100 = (
                    conteo.get("5", 0.0) * 100 +
                    conteo.get("4", 0.0) * 80 +
                    conteo.get("3", 0.0) * 60 +
                    conteo.get("2", 0.0) * 40 +
                    conteo.get("1", 0.0) * 20
)

                score_5 = score_100 / 20
            except Exception as e:
                score_100 = ""
                score_5 = ""

            nsps_100.append(score_100)
            nsps_5.append(score_5)
        pnsp = [(nsp5 / (np.median(nsps_5))*100 if np.median(nsps_5) > 0 else 0) for nsp5 in nsps_5]
        codigos = [f"V{i+1}" for i in range(len(preguntas_sin_general))]
        # Crear DataFrame
        tabla_resultado = pd.DataFrame({
            "Atributo": preguntas_sin_general,
            "CORR": correlaciones,
            "PESO": pesos,
            "P(NIP)": pnips,
            "NSP_100": nsps_100,
            "NSP_5": nsps_5,
            "P(NSP)":pnsp,
            "Código": codigos
        })

        # Agregar fila total
        fila_total = pd.DataFrame([{
            "Atributo": "",
            "CORR": np.sum(correlaciones),
            "PESO": np.median(pesos),
            "P(NIP)": np.average(pnips),
            "NSP_100": np.median(nsps_100),
            "NSP_5": np.median(nsps_5),
            "P(NSP)": np.average(pnsp),
            "Código": ""
        }])
        tabla_resultado = pd.concat([tabla_resultado, fila_total], ignore_index=True)

        # HTML estilo
        html = """
        <style>
        .tabla-nip {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
        }
        .tabla-nip th, .tabla-nip td {
            border: 1px solid #ddd;
            text-align: center;
            padding: 8px;
        }
        .tabla-nip th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        .tabla-nip thead tr:first-child th {
            background-color: #f9f9f9;
            border-bottom: none;
        }
        .tabla-nip thead tr:nth-child(2) th {
            border-top: none;
        }
        .tabla-nip td:first-child {
            text-align: left;
        }
        .tabla-nip tbody tr:last-child {
            background-color: #fcfcfc;
            font-weight: bold;
        }
        </style>
        <table class="tabla-nip">
            <thead>
                <tr>
                    <th rowspan="2">Atributo</th>
                    <th colspan="2">NIP</th>
                    <th rowspan="2">P(NIP)</th>
                    <th colspan="2">NSP</th>
                    <th rowspan="2">P(NSP)</th>
                    <th rowspan="2">Código</th>
                </tr>
                <tr>
                    <th>CORR</th>
                    <th>PESO</th>
                    <th>0:100</th>
                    <th>0:5</th>
                </tr>
            </thead>
            <tbody>
        """

        for i, row in tabla_resultado.iterrows():
            pnip=row['P(NIP)']
            pnsp_val=row['P(NSP)']
            bg_color = ""
            if row["Atributo"] != "":
                if isinstance(pnip, (int, float)) and isinstance(pnsp_val, (int, float)):
                    if pnip > 100 and pnsp_val < 100:
                        bg_color = "#ff4d4d"
                    elif pnip <= 100 and pnsp_val < 100:
                        bg_color = "#ffff66"
                    elif pnip >= 100 and pnsp_val >= 100:
                        bg_color = "#a2d96c"
            html += f"<tr>"
            html += f"<td>{row['Atributo']}</td>"
            html += f"<td>{row['CORR']:.2f}</td>"
            html += f"<td>{row['PESO']:.2%}</td>"
            html += f"<td>{row['P(NIP)']:.2f}</td>"
            html += f"<td>{'' if row['NSP_100'] == '' else format(row['NSP_100'], '.1f')}</td>"
            html += f"<td>{'' if row['NSP_5'] == '' else format(row['NSP_5'], '.2f')}</td>"
            html += f"<td>{row['P(NSP)']:.2f}</td>"
            html += f"<td style='background-color:{bg_color}'>{row['Código']}</td>"
            html += f"</tr>"

        html += "</tbody></table>"
        st.markdown(html, unsafe_allow_html=True)

        st.markdown(
                    "<hr style='border: none; height: 1px; background: linear-gradient(to right, #ccc, #eee, #ccc); margin: 120px 0;'>",
                     unsafe_allow_html=True)
        # Cálculo de ISC usando P(NIP) y NSP_100
        isc_100 = sum(p * nsp for p, nsp in zip(pesos, nsps_100) if isinstance(nsp, (int, float)))
        isc_5 = isc_100 / 20

        # Mostrar tabla ISC
        st.markdown(f"""
        <style>
        .tabla-isc {{
            border-collapse: collapse;
            width: 50%;
            margin: 30px auto;
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
        }}
        .tabla-isc th, .tabla-isc td {{
            border: 1px solid black;
            padding: 10px;
        }}
        .tabla-isc td {{
            background-color: #d9e9c1;
        }}
        .tabla-isc td:nth-child(2) {{
            text-align: right;
            font-weight: bold;
        }}
        </style>

        <table class="tabla-isc">
            <tbody>
                <tr>
                    <td>ISC Ponderado por NIP</td>
                    <td>{isc_100:.2f}</td>
                </tr>
                <tr>
                    <td>ISC Ponderado por NIP (1:5)</td>
                    <td>{isc_5:.2f}</td>
                </tr>
            </tbody>
        </table>
        """, unsafe_allow_html=True)

        st.markdown(
                    "<hr style='border: none; height: 1px; background: linear-gradient(to right, #ccc, #eee, #ccc); margin: 120px 0;'>",
                     unsafe_allow_html=True)
        with st.container():

            fig, ax = plt.subplots(figsize=(15, 7))

            # Límites de los ejes
            x_min, x_max = 0, 200
            y_min, y_max = 60, 140

            ax.set_xlim(x_min, x_max)
            ax.set_ylim(y_min, y_max)

            # Ticks explícitos cada 20 en X y cada 10 en Y
            ax.set_xticks(list(range(x_min, x_max + 1, 20)))
            ax.set_yticks(list(range(y_min, y_max + 1, 10)))

            # Líneas principales
            ax.axvline(x=100, color='black', linewidth=1)
            ax.axhline(y=100, color='black', linewidth=1)

            # Colores de cuadrantes
            ax.fill_between([100, x_max], y_max, 100, color='#a2d96c', alpha=1.0)
            ax.fill_between([100, x_max], y_min, 100, color='#ff4d4d', alpha=1.0)
            ax.fill_between([x_min, 100], y_min, 100, color='#ffff66', alpha=1.0)
            ax.fill_between([x_min, 100], y_max, 100, color='#ffffff', alpha=1.0)

            # Texto en cuadrantes
            text_props = {'ha': 'center', 'va': 'center', 'fontsize': 7, 'weight': 'bold'}
            padding_x = (x_max - x_min) / 6
            padding_y = (y_max - y_min) / 6

            ax.text(100 + padding_x, 100 + padding_y, "Muy importantes\nAlta satisfacción", **text_props)
            ax.text(100 - padding_x, 100 + padding_y, "Poco importantes\nAlta satisfacción", **text_props)
            ax.text(100 + padding_x, 100 - padding_y, "Muy importantes\nBaja satisfacción", **text_props, color='white')
            ax.text(100 - padding_x, 100 - padding_y, "Poco importantes\nBaja satisfacción", **text_props, color='darkred')

            # Líneas intermedias tenues
            for x in range(x_min, x_max + 1, 20):
                ax.axvline(x, color='gray', linestyle='--', linewidth=0.3)
            for y in range(y_min, y_max + 1, 10):
                ax.axhline(y, color='gray', linestyle='--', linewidth=0.3)

            # Filtrar y graficar puntos dentro del rango
            puntos_filtrados = [(x, y) for x, y in zip(pnips, pnsp) if x_min <= x <= x_max and y_min <= y <= y_max]

            for x, y in puntos_filtrados:
                ax.scatter(x, y, color='navy', s=45, marker='D', zorder=5)
                ax.text(x + 1.2, y + 1.2, f"({x:.1f}, {y:.1f})", fontsize=6, color='black', zorder=6)

            # Estética
            ax.set_xlabel("Importancia", fontsize=9, weight='bold')
            ax.set_ylabel("Satisfacción", fontsize=9, weight='bold')
            ax.set_title("Matriz Importancia - Satisfacción", fontsize=11, weight='bold', pad=10)
            ax.grid(False)
            ax.set_aspect('equal', adjustable='box')

            st.pyplot(fig)
            st.markdown("</div>", unsafe_allow_html=True)        
else:
    st.warning("⚠️ No se pudo construir la tabla NIP porque falta la columna general o el DataFrame.")
