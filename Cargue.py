import streamlit as st
import pandas as pd
import importlib.util
import tempfile
import os
import ast
import json
 # --- Funciones de procesamiento y pivotaje ---
import io
def procesar_excel(df):
    # Procesamiento tipo Procesar.ipynb: limpieza, renombrado, reemplazo de valores
    nuevos_encabezados = [str(columna).split('-')[-1].strip() if '-' in str(columna) else str(columna) for columna in df.columns]
    df.columns = nuevos_encabezados
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    reemplazos = {
        "5 (Supera las expectativas)":5,"5 Supera notablemente las expectativas": 5,"5 Supera notablemente las expectativas": 5 , "5. Supera notablemente las expectativas": 5, "5 - Muy satisfecho":5, "Muy satisfecho":5, "5. Supera notablemente las expectativas":5,
        "Supera notablemente las expectativas":5, "5 - Supera notablemente las expectativas":5, "5-Supera notablemente las expectativas":5,
        "Supera las expectativas / Exceeds expectations":5, "5. Supera notablemente mis expectativas":5, "5.Muy por encima de sus expectativas":5,
        "5- Supera notablemente las expectativas":5, "4":4, "4 - Satisfecho":4, "Satisfecho":4, "Supera las expectativas / Exceeds expectations4":4,
        "Supera las expectativas / Exceeds expectations\n4":4, "3":3, "Cumple las expectativas / Meets expectations3":3,
        "Cumple las expectativas / Meets expectations\n3":3, "Ni satisfecho - ni insatisfecho":3, "3 - Ni satisfecho - ni insatisfecho":3,
        "Cumple las expectativas / Meets expectations":3, "2":2, "Por debajo de las expectativas / Below expectations\n2":2,
        "Muy insatisfecho":2, "2 - Insatisfecho":2, "Insatisfecho":2, "1 (Muy por debajo de las expectativas)":1, "1 - Muy insatisfecho":1,
        "Muy por debajo de las expectativas":1, "1 - Muy por Debajo de las Expectativas":1, "1. Muy por debajo de las expectativas":1,
        "1. Muy por debajo de mis expectativas":1,"1 Muy por debajo de las expectativas":1, "1.Muy por debajo de sus expectativas":1, "Muy por debajo de las expectativas / Far below expectations":1
    }
    df = df.replace(reemplazos)
    return df

def pivotear_excel(df):
    # Procesamiento tipo pivotaje 1.ipynb: pivotar según columnas detectadas
    def detectar_columnas_caso(df):
        casos=[]
        for columns in df.columns:
            if 'caso' in columns.strip().lower():
                casos.append(columns)
        return casos
    def detectar_columnas_respuesta(df):
        for columns in df.columns:
            if 'respuesta' in columns.strip().lower():
                return columns
    def detectar_column_pregunta(df):
        for columns in df.columns:
            if 'pregunta' in columns.strip().lower():
                return columns
    iindexc=detectar_columnas_caso(df)
    icolumnsc=detectar_column_pregunta(df)
    ivaluesc=detectar_columnas_respuesta(df)
    if iindexc and icolumnsc and ivaluesc:
        data = df.pivot_table(index=iindexc, columns=icolumnsc, values=ivaluesc, aggfunc='first').reset_index()
        data = procesar_excel(data)
        return data
    else:
        return df

def _df_arrow_safe(d):
    """Copia solo para visualización, compatible con Arrow.

    Las columnas de preguntas mezclan enteros (1-5) y texto ('No aplica'), lo
    que hace fallar la serialización de st.dataframe (ArrowInvalid en la
    terminal). Convertir los valores no-texto de esas columnas a str evita el
    error sin tocar los datos reales del reporte.
    """
    d2 = d.copy()
    for c in d2.columns[d2.dtypes.eq("object")]:
        d2[c] = d2[c].map(lambda v: v if (isinstance(v, str) or pd.isna(v)) else str(v))
    return d2

def procesar_qualtrics(df):
    # Procesamiento tipo Procesar_qualtrix.ipynb: limpieza de archivos exportados de Qualtrics
    #
    # Existen dos variantes del formato Qualtrics:
    #
    # Variante A — Encabezados técnicos (este caso):
    #   Fila 1 Excel (header pandas): códigos técnicos  StartDate, Q7, Q8_1, DistributionChannel...
    #   Fila 2 Excel (df.iloc[0]):    descripciones ES   Fecha de inicio, ¿Cómo califica...
    #   Fila 3+ Excel:                datos reales
    #   → Hay que reemplazar los headers técnicos por las descripciones en español.
    #
    # Variante B — Encabezados ya en español (NO usar este modo):
    #   Fila 1 Excel (header pandas): Fecha de inicio, ¿Cómo califica... (ya son los correctos)
    #   Fila 2 Excel (df.iloc[0]):    datos reales
    #   → Usar modo PROCESAR, no Qualtrics.
    #
    # Detección: si los nombres de columna son códigos técnicos de Qualtrics
    qualtrics_system_cols = {
        'startdate', 'enddate', 'status', 'ipaddress', 'progress',
        'duration (in seconds)', 'finished', 'recordeddate', 'responseid',
        'recipientlastname', 'recipientfirstname', 'recipientemail',
        'externaldatareference', 'locationlatitude', 'locationlongitude',
        'distributionchannel', 'userlanguage'
    }
    cols_lower = {str(c).lower() for c in df.columns}
    tiene_tecnicos = bool(cols_lower & qualtrics_system_cols)

    if tiene_tecnicos:
        # Usar la fila de descripciones en español como nuevos encabezados
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        # Simplificar nombres largos de Qualtrics: tomar solo la parte después del último " - "
        # Ej: "Por favor, califique cada atributo... - Los recursos empleados..."
        #  →  "Los recursos empleados durante la prestación del servicio evaluado fueron adecuados"
        # Esto evita que los nombres largos con caracteres especiales rompan las fórmulas de Excel
        def simplificar_col(col):
            col = str(col).strip()
            if ' - ' in col:
                return col.split(' - ')[-1].strip()
            return col
        df.columns = [simplificar_col(c) for c in df.columns]

    # Manejar columnas duplicadas agregando sufijos
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    
    # Normalizar floats enteros: 4.0 → 4, 3.0 → 3, etc.
    import math
    def norm(x):
        if isinstance(x, float) and not math.isnan(x) and x == int(x):
            return int(x)
        return x
    df = df.map(norm)

    # Limpiar espacios y saltos de línea de todas las celdas string
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # Diccionario de reemplazos exhaustivo para Qualtrics
    reemplazos_qualtrics = {
        # Formatos con saltos de línea al inicio/final
        '5- Totalmente satisfecho': 5,
        '4- Satisfecho': 4,
        '3- Neutral': 3,
        '2- Insatisfecho': 2,
        '1-Totalmente insatisfecho': 1,
        # Variantes con "Estás"
        '5- Estás totalmente Satisfecho': 5,
        '5- Estás totalmente satisfecho': 5,
        '4- Estás satisfecho': 4,
        '3- Estás neutral': 3,
        '2- Estás insatisfecho': 2,
        '1- Estás totalmente insatisfecho': 1,
        '1-Estás totalmente insatisfecho': 1,
        # Formatos con "Muy" y standalone (sin número prefijo)
        '5 - Muy satisfecho': 5,
        'Muy satisfecho': 5,
        '4 - Satisfecho': 4,
        'Satisfecho': 4,
        '3 - Neutral': 3,
        '3 - Ni satisfecho - ni insatisfecho': 3,
        'Ni satisfecho - ni insatisfecho': 3,
        'Neutral': 3,
        '2 - Insatisfecho': 2,
        'Insatisfecho': 2,
        '1 - Muy insatisfecho': 1,
        'Muy insatisfecho': 1,
        # Formatos de expectativas (todas las variantes)
        '5 - Supera notablemente las expectativas': 5,
        '5- Supera notablemente las expectativas': 5,
        '5 Supera notablemente las expectativas': 5,
        'Supera notablemente las expectativas': 5,
        '5. Supera notablemente las expectativas': 5,
        '5 (Supera las expectativas)': 5,
        '4 - Cumple las expectativas': 4,
        '4- Cumple las expectativas': 4,
        'Cumple las expectativas': 3,
        '3 - Cumple las expectativas': 3,
        '3- Cumple las expectativas': 3,
        '2 - Por debajo de las expectativas': 2,
        '2- Por debajo de las expectativas': 2,
        'Por debajo de las expectativas': 2,
        '1 - Muy por debajo de las expectativas': 1,
        '1- Muy por debajo de las expectativas': 1,
        'Muy por debajo de las expectativas': 1,
        'Cumple las expectativas': 3,
        'Por debajo de las expectativas': 2,
        'Muy por debajo de las expectativas': 1,
        # Ya vienen como números pero como string
        '5': 5,
        '4': 4,
        '3': 3,
        '2': 2,
        '1': 1,
        # Contribución (escala diferente)
        'Mucho': 4,
        'Algo': 3,
        'Poco': 2,
        'Nada': 1
    }
    
    # Aplicar reemplazos y convertir tipos explícitamente
    df = df.replace(reemplazos_qualtrics).infer_objects(copy=False)
    
    # Convertir formato de fecha de Qualtrics si existe columna 'Fecha registrada'
    if 'Fecha registrada' in df.columns:
        try:
            df['Fecha registrada'] = df['Fecha registrada'].astype(str)
            df['Fecha registrada'] = pd.to_datetime(df['Fecha registrada'], format='mixed', errors='coerce').dt.strftime('%d/%m/%Y')
        except:
            pass  # Si falla, mantener valores originales
    
    return df

# Diccionario de oficinas vinculado a un único script común y procesos asociados
diccionario_oficinas = {
    "Dirección de Tecnología Informática y Comunicaciones": {
        "script": "Generararchivoexcel_mantenimiento_tic_trimestre",
        "procesos": ["Sistemas de Información desarrollo, mantenimiento y soporte a usuarios"]
    },
    "Operaciones Tic": {
        "script": "Generararchivoexcel_mantenimiento_tic_CSU",
        "procesos": ["Soporte de Servicios TIC"]
    },
    "Dirección de servicios Administrativos": {
        "script": "Generararchivoexcel_mantenimientoDSA",
        "procesos": ["Mantenimiento DSA"]
    },
    "Admisiones": {
        "script": "Generararchivoexcel_admisiones_posgrado",
        "procesos": ["Satisfacción con respecto al servicio recibido durante el proceso de admisión a posgrado"]
    },
    "Financiamiento Empresarial": {
        "script": "Generararchivoexcel_Financiamiento_Empresarial",
        "procesos": ["Servicio prestado en facturación a través de distintas plataformas"]
    },
    "Dirección Financiera": {
        "script": "Generararchivoexcel_registro_provedores",
        "procesos": ["Registro de Proveedores"]
    },
    "Sección de Compras": {
        "script": "Generararchivoexcel_Adquisicion_bienes",
        "procesos": ["Adquisición de bienes"]
    },
    "Laboratorio de Geotecnia y Materiales de Construcción": {
        "script": "Generararchivoexcel_laboratorio_geotecnia",
        "procesos": ["Servicios del Laboratorio de Geotecnia y Materiales de Construcción"]
    },
    "Departamento de Registro": {
        "script": "Generararchivoexcel_coordinadores",
        "procesos": ["Informe de calidad coordinadores de Pregrado"]
    },
    "Departamenro de Ingeniería Mecánica": {
        "script": "Generararchivoexcel_laboratori_cimm",
        "procesos": ["CIMM"]
    },
    "Tesorería": {
        "script": "Generararchivoexcel_Tesoreria",
        "procesos": ["Recaudo Web - Zona Pago", "Recaudo Web - Place to Pay", "Presencial Caja"]
    },
    "Oficina de Registro": {
        "script": "Generararchivoexcel_oficinaregistro_grado",
        "procesos": ["Trámite de Grado"]
    },
    "Oficina de Planeación": {
        "script": "Generararchivoexcel_planeacion",
        "procesos": ["Satisfación de estadisticas institucionales"]
    },
    "Dirección de Gestión y Relaciones Internacionales": {
        "script": "Generararchivoexcel_movilidad_entrante",
        "procesos": ["Movilidad Estudiantil Internacional Entrante No 1","Movilidad Estudiantil Internacional Entrante No 2","Gestión de la Movilidad Estudiantil Internacional (saliente)"]
    },
    "Prueba": {
        "script": "Generararchivoexcel_prueba",
        "procesos": ["Prueba grafica"]
    },
    "Almacen": {
        "script": "Generararchivoexcel_Almacen",
        "procesos": ["Entrega de Insumos y Compras Directas - Entrega de Activos"]
    },
    "Sección de Servicios Generales": {
        "script": "Generararchivoexcel_servicio_transporte_seguridad_Aseo",
        "procesos": ["Servicio de transporte","Seguridad en Uninorte","Servicios de Aseo"]
    },
    "Registro": {
        "script": "Generararchivoexcel_certificaciones",
        "procesos": ["Certificaciones académicas por vía Web"]
    },
    "Oficina Genérica / Personalizada": {
        "script": "Generararchivoexcel_generico",
        "procesos": ["Proceso Personalizado"]
    }
}

# ── Oficinas / procesos personalizados por el usuario (persistentes) ──────────
# Las oficinas de arriba son fijas. Desde la interfaz el usuario puede crear
# oficinas nuevas o agregar procesos a oficinas existentes; esas adiciones se
# guardan en 'oficinas_personalizadas.json' y se fusionan aquí para que
# sobrevivan reinicios de la aplicación.
OFICINAS_BASE = tuple(diccionario_oficinas.keys())
RUTA_OFICINAS_PERSONALIZADAS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "oficinas_personalizadas.json")


def cargar_oficinas_personalizadas():
    """Devuelve el dict de personalizaciones guardadas ({} si no hay o está dañado)."""
    try:
        if os.path.exists(RUTA_OFICINAS_PERSONALIZADAS):
            with open(RUTA_OFICINAS_PERSONALIZADAS, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def guardar_oficinas_personalizadas(data):
    """Persiste el dict de personalizaciones en disco."""
    with open(RUTA_OFICINAS_PERSONALIZADAS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def fusionar_oficinas_personalizadas(base, personalizadas):
    """Mezcla en 'base' (in-place) las oficinas/procesos añadidos por el usuario.

    - Oficina que ya existe en base -> se agregan solo sus procesos nuevos.
    - Oficina nueva                 -> se crea con su script (por defecto el
                                       genérico) y sus procesos.
    """
    for oficina, cfg in personalizadas.items():
        if not isinstance(cfg, dict):
            continue
        procesos = [str(p) for p in cfg.get("procesos", []) if str(p).strip()]
        if oficina in base:
            for p in procesos:
                if p not in base[oficina]["procesos"]:
                    base[oficina]["procesos"].append(p)
        else:
            base[oficina] = {
                "script": cfg.get("script") or "Generararchivoexcel_generico",
                "procesos": procesos or ["Proceso Personalizado"],
            }
    return base


fusionar_oficinas_personalizadas(diccionario_oficinas, cargar_oficinas_personalizadas())

st.set_page_config(page_title="Exportador de Excel", page_icon="📁", layout="wide")



# --- Estilos personalizados ---
st.markdown("""
    <style>
    .main-title {
        text-align: center;
        font-size: 36px;
        color: #2c3e50;
        margin-top: 20px;
        font-weight: bold;
    }
    .section-title {
        font-size: 20px;
        margin-top: 25px;
        color: #34495e;
        border-bottom: 1px solid #ccc;
        padding-bottom: 6px;
        max-width: 90%;
        margin-left: auto;
        margin-right: auto;
    }
    .footer {
        text-align: center;
        font-size: 13px;
        color: #999;
        margin-top: 30px;
    }
    .block-container {
        max-width: 1200px;
        margin: auto;
    }
    </style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
.main-title {
    font-size: 36px;
    color: #2f3e75;
    text-align: center;
    font-weight: 700;
    margin-top: 30px;
    margin-bottom: 20px;
    background: linear-gradient(to right, #f5f9ff, #e8effc);
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
    border: 1px solid #d0d8ec;
    max-width: 900px;
    margin-left: auto;
    margin-right: auto;
}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="main-title">📤 Exportador Automático de Datos a Excel</div>', unsafe_allow_html=True)

# --- Selector de método de procesamiento ---
st.markdown('<div class="section-title">0️⃣ Selecciona el método de procesamiento</div>', unsafe_allow_html=True)
metodo = st.radio("¿Cómo deseas procesar el archivo?", ["Procesar", "Pivotear", "Qualtrics"], horizontal=True)

# --- Filtro de período (solo Qualtrics) ---
fecha_inicio_q = None
fecha_fin_q = None
if metodo == "Qualtrics":
    from datetime import date as _date
    st.markdown('<div class="section-title">📅 Período de análisis</div>', unsafe_allow_html=True)
    st.caption("Qualtrics almacena todo el histórico. Selecciona el rango de fechas de inicio de las encuestas que quieres analizar.")
    _col1, _col2 = st.columns(2)
    with _col1:
        fecha_inicio_q = st.date_input(
            "Fecha inicio",
            value=st.session_state.get("fecha_inicio_q", _date(_date.today().year, 1, 1)),
            key="fecha_inicio_q"
        )
    with _col2:
        fecha_fin_q = st.date_input(
            "Fecha fin",
            value=st.session_state.get("fecha_fin_q", _date.today()),
            key="fecha_fin_q"
        )

# --- Sección: Carga del archivo ---
st.markdown('<div class="section-title">1️⃣ Subir archivo Excel</div>', unsafe_allow_html=True)
archivo_excel = st.file_uploader("Cargar archivo .xlsx", type=["xlsx"])
if archivo_excel:
    st.session_state["archivo_excel"] = archivo_excel
if "archivo_excel" in st.session_state and archivo_excel is None:
    archivo_excel = st.session_state["archivo_excel"]

if archivo_excel is not None:
    try:
        df = pd.read_excel(archivo_excel)
        # Procesar según método seleccionado
        if metodo == "Procesar":
            df = procesar_excel(df)
            st.success("✅ Archivo procesado correctamente (Procesar)")
        elif metodo == "Pivotear":
            df = pivotear_excel(df)
            st.success("✅ Archivo procesado correctamente (Pivotear)")
        elif metodo == "Qualtrics":
            total_bruto = len(df) - 1  # -1 por la fila de descripciones de Qualtrics
            df_data = df.iloc[1:].copy()  # saltar fila de descripciones

            # 1) Filtros de validez fijos
            if 'Status' in df_data.columns:
                df_data = df_data[df_data['Status'].astype(str).str.strip() == 'IP Address']
            if 'Progress' in df_data.columns:
                df_data = df_data[pd.to_numeric(df_data['Progress'], errors='coerce') == 100]
            if 'Finished' in df_data.columns:
                df_data = df_data[df_data['Finished'].astype(str).str.strip().str.lower().isin(['true', '1'])]

            # 2) Filtro de fechas
            _fi = st.session_state.get("fecha_inicio_q")
            _ff = st.session_state.get("fecha_fin_q")
            if _fi and _ff and 'StartDate' in df_data.columns:
                _fechas = pd.to_datetime(df_data['StartDate'], errors='coerce')
                df_data = df_data[
                    (_fechas.dt.date >= _fi) &
                    (_fechas.dt.date <= _ff)
                ]

            encuestas_validas = len(df_data)
            descartadas = total_bruto - encuestas_validas

            # Reconstruir con fila de descripciones para que procesar_qualtrics renombre columnas
            df = procesar_qualtrics(pd.concat([df.iloc[[0]], df_data]).reset_index(drop=True))

            st.success(f"✅ Qualtrics procesado: **{encuestas_validas}** encuestas válidas | **{descartadas}** descartadas")
            if _fi and _ff:
                st.caption(f"Período: {_fi.strftime('%d/%m/%Y')} – {_ff.strftime('%d/%m/%Y')}")
        st.session_state["df_encuesta"] = df
    except Exception as e:
        st.error(f"❌ Error al leer o procesar el archivo: {e}")
elif "df_encuesta" in st.session_state:
    df = st.session_state["df_encuesta"]

# --- Sección: Configuración ---
st.markdown('<div class="section-title">2️⃣ Seleccionar oficina y parámetros</div>', unsafe_allow_html=True)
_oficinas_disponibles = list(diccionario_oficinas.keys())
_oficina_guardada = st.session_state.get("oficina_seleccionada", _oficinas_disponibles[0])
if _oficina_guardada not in _oficinas_disponibles:
    _oficina_guardada = _oficinas_disponibles[0]
oficina_seleccionada = st.selectbox(
    "🏢 Selecciona la oficina",
    options=_oficinas_disponibles,
    index=_oficinas_disponibles.index(_oficina_guardada)
)
st.session_state["oficina_seleccionada"] = oficina_seleccionada
procesos_disponibles = diccionario_oficinas[oficina_seleccionada]["procesos"]
valor_guardado = st.session_state.get("proceso_seleccionado", procesos_disponibles[0])
if valor_guardado not in procesos_disponibles:
    valor_guardado = procesos_disponibles[0]

proceso_seleccionado = st.selectbox(
    "🧪 Selecciona el proceso asociado",
    options=procesos_disponibles,
    index=procesos_disponibles.index(valor_guardado)
)
st.session_state["proceso_seleccionado"] = proceso_seleccionado

# --- Agregar / gestionar oficinas y procesos ---
with st.expander("➕ Agregar o gestionar oficinas y procesos"):
    st.caption(
        "Crea una oficina nueva, agrega un proceso a una oficina existente o "
        "elimina las personalizaciones que hayas creado. Los cambios quedan "
        "guardados de forma permanente."
    )
    _modo_gestion = st.radio(
        "¿Qué deseas hacer?",
        ["Nueva oficina", "Agregar proceso a oficina existente", "Eliminar personalizados"],
        horizontal=True,
        key="modo_gestion_oficina"
    )

    if _modo_gestion == "Nueva oficina":
        _nueva_of = st.text_input("🏢 Nombre de la nueva oficina", key="in_nueva_oficina")
        _nuevo_pr = st.text_input("🧪 Primer proceso asociado", key="in_nueva_oficina_proc")
        _opts_formato = list(diccionario_oficinas.keys())
        _idx_formato = (_opts_formato.index("Oficina Genérica / Personalizada")
                        if "Oficina Genérica / Personalizada" in _opts_formato else 0)
        _formato = st.selectbox(
            "📄 Formato base del informe",
            options=_opts_formato,
            index=_idx_formato,
            help="La nueva oficina generará el informe con el mismo formato/script de "
                 "la oficina elegida aquí. Si no estás seguro, deja 'Oficina Genérica / "
                 "Personalizada'."
        )
        if st.button("💾 Guardar nueva oficina", key="btn_nueva_oficina"):
            _n = _nueva_of.strip()
            _p = _nuevo_pr.strip()
            if not _n:
                st.warning("⚠️ Escribe el nombre de la oficina.")
            elif not _p:
                st.warning("⚠️ Escribe el primer proceso asociado.")
            elif _n in diccionario_oficinas:
                st.warning(f"⚠️ La oficina '{_n}' ya existe. Usa 'Agregar proceso a oficina existente'.")
            else:
                _custom = cargar_oficinas_personalizadas()
                _custom[_n] = {
                    "script": diccionario_oficinas[_formato]["script"],
                    "procesos": [_p],
                }
                guardar_oficinas_personalizadas(_custom)
                st.session_state["oficina_seleccionada"] = _n
                st.session_state["proceso_seleccionado"] = _p
                st.success(f"✅ Oficina '{_n}' creada con el proceso '{_p}'.")
                st.rerun()

    elif _modo_gestion == "Agregar proceso a oficina existente":
        _of_dest = st.selectbox(
            "🏢 Oficina", options=list(diccionario_oficinas.keys()), key="sel_of_dest")
        _pr_nuevo = st.text_input("🧪 Nuevo proceso a agregar", key="in_proc_existente")
        if st.button("💾 Agregar proceso", key="btn_agregar_proceso"):
            _p = _pr_nuevo.strip()
            if not _p:
                st.warning("⚠️ Escribe el nombre del proceso.")
            elif _p in diccionario_oficinas[_of_dest]["procesos"]:
                st.warning(f"⚠️ El proceso '{_p}' ya existe en '{_of_dest}'.")
            else:
                _custom = cargar_oficinas_personalizadas()
                _entry = _custom.get(_of_dest, {})
                _procs = _entry.get("procesos", [])
                _procs.append(_p)
                _entry["procesos"] = _procs
                _custom[_of_dest] = _entry
                guardar_oficinas_personalizadas(_custom)
                st.session_state["oficina_seleccionada"] = _of_dest
                st.session_state["proceso_seleccionado"] = _p
                st.success(f"✅ Proceso '{_p}' agregado a '{_of_dest}'.")
                st.rerun()

    else:  # Eliminar personalizados
        _custom = cargar_oficinas_personalizadas()
        _items = []  # (label, tipo, oficina, proceso)
        for _off, _cfg in _custom.items():
            if not isinstance(_cfg, dict):
                continue
            if _off not in OFICINAS_BASE:
                _items.append((f"🏢 Oficina completa: {_off}", "oficina", _off, None))
            for _p in _cfg.get("procesos", []):
                _items.append((f"🧪 Proceso: {_p}  —  ({_off})", "proceso", _off, _p))
        if not _items:
            st.info("No hay oficinas ni procesos personalizados para eliminar.")
        else:
            _sel_label = st.selectbox(
                "Selecciona qué eliminar", options=[it[0] for it in _items], key="sel_eliminar")
            st.caption("Solo se pueden eliminar oficinas y procesos que tú hayas agregado; "
                       "las oficinas y procesos originales no se ven afectados.")
            if st.button("🗑️ Eliminar", key="btn_eliminar_personalizado"):
                _tipo, _off, _p = next((it[1], it[2], it[3]) for it in _items if it[0] == _sel_label)
                _custom2 = cargar_oficinas_personalizadas()
                if _tipo == "oficina":
                    _custom2.pop(_off, None)
                else:
                    _entry = _custom2.get(_off, {})
                    _procs = [x for x in _entry.get("procesos", []) if x != _p]
                    if _procs:
                        _entry["procesos"] = _procs
                        _custom2[_off] = _entry
                    else:
                        # sin procesos añadidos: quitar la entrada (revierte a base o
                        # elimina la oficina personalizada por completo)
                        _custom2.pop(_off, None)
                guardar_oficinas_personalizadas(_custom2)
                st.success("✅ Elemento eliminado.")
                st.rerun()

# Campo adicional para nombre de oficina personalizado (solo para Oficina Genérica)
if oficina_seleccionada == "Oficina Genérica / Personalizada":
    nombre_oficina_personalizado = st.text_input(
        "🏷️ Nombre de la oficina para el informe",
        value=st.session_state.get("nombre_oficina_personalizado", ""),
        placeholder="Ej: Oficina de Relaciones Internacionales",
        help="Este nombre aparecerá en el informe Excel generado"
    )
    st.session_state["nombre_oficina_personalizado"] = nombre_oficina_personalizado
else:
    # Para otras oficinas, usar el nombre de la oficina seleccionada
    st.session_state["nombre_oficina_personalizado"] = oficina_seleccionada

nombre_archivo = st.text_input("📝 Nombre del archivo de salida (sin extensión)", value="exportado")
numerodepoblacion = st.number_input("👥 Número de población", min_value=1, step=1)

# --- Selección de periodo ---
periodo_unico  = st.text_input("📝 Escribir periodo en que se relizo la encuesta", value="Periodo")

# --- Vista previa del archivo cargado ---
if archivo_excel is not None:
    # Usar el dataframe ya procesado del session_state
    df = st.session_state.get("df_encuesta", pd.DataFrame())
    
    if df.empty:
        st.error("❌ Error: No se pudo procesar el archivo")
    else:
        st.success("✅ Archivo cargado exitosamente")
    
    # Inicializar variables que se usan en este bloque
    columnas_pregunta_detectadas = []
    columnas_observaciones_detectadas = []
    nombre_columna_general = ""
    #st.dataframe(df.head(), use_container_width=True)

    #Oficina con filtros
    oficina_filtros=["Almacen","Prueba","Sección de Servicios Generales","Registro"]
    proceso_filtros=["Entrega de Insumos y Compras Directas - Entrega de Activos","Prueba grafica","Servicio de transporte","Seguridad en Uninorte","Servicios de Aseo","Certificaciones académicas por vía Web"]
   
    if oficina_seleccionada in oficina_filtros and proceso_seleccionado in proceso_filtros:
        st.markdown('<div class="section-title">🔍 Filtros específicos para prueba</div>', unsafe_allow_html=True)

        if archivo_excel is not None and not df.empty:

            # Filtros eliminados de la barra lateral, solo se muestra la tabla filtrada si aplica
            st.dataframe(_df_arrow_safe(df), width='stretch')

        else:
            st.sidebar.warning("⚠️ Sube un archivo Excel para aplicar los filtros.")
    else:   
        st.dataframe(_df_arrow_safe(df), width='stretch')
    columnas_todo_no_aplica = [col for col in df.columns if (df[col] == "No Aplica").all()]
    # --- Detectar columnas de preguntas automáticamente ---
    posibles_valores = {"1", "2", "3", "4", "5", "No Aplica",
                        "1.0", "2.0", "3.0", "4.0", "5.0"}
    for col in df.columns:
        valores = set(df[col].dropna().astype(str).unique())
        contiene_valores = valores.issubset(posibles_valores) or len(valores.intersection(posibles_valores)) >= 3
        es_general = "general" in col.lower()
        es_numero = "numero" in col.lower()
        # Excluir columnas que terminan en _1, _2, etc. (duplicadas de Qualtrics)
        es_duplicada = col.endswith('_1') or col.endswith('_2') or col.endswith('_3') or col.endswith('_4')
        if contiene_valores and not es_general and not es_numero and not es_duplicada and not df[col].isna().all():
            columnas_pregunta_detectadas.append(col)
        elif "general" in col.lower() and not es_duplicada:
            # Solo aceptar como columna general si tiene valores numéricos (no texto libre)
            if contiene_valores and not df[col].isna().all():
                nombre_columna_general = col
    columnas_pregunta_detectadas=[x for x in columnas_pregunta_detectadas if x not in columnas_todo_no_aplica]

    # Nota sobre truncado de nombres (solo Qualtrics) — más visible
    if metodo == "Qualtrics":
        st.warning(
            "⚠️ **Atención — los nombres de columna cambian al procesar Qualtrics**\n\n"
            "Los encabezados largos se **recortan desde el último \" - \"** hacia adelante.\n\n"
            "**Ejemplo:** *\"Por favor califique cada atributo del servicio — Los recursos empleados durante la prestación\"*"
            " → aparece como **\"Los recursos empleados durante la prestación\"**\n\n"
            "Usa este nombre acortado al seleccionar preguntas, columna general y filtros abajo."
        )

    # --- Sección: Eliminar columnas no relevantes (opcional) ---
    st.markdown('<div class="section-title">🗑️ Eliminar columnas innecesarias (opcional)</div>', unsafe_allow_html=True)
    st.caption("Puedes excluir columnas que no aporten información al informe (metadatos de Qualtrics, campos vacíos, etc.).")
    cols_a_eliminar = st.multiselect(
        "Selecciona columnas a eliminar del análisis:",
        options=df.columns.tolist(),
        default=[],
        key="cols_eliminar"
    )
    if cols_a_eliminar:
        df = df.drop(columns=[c for c in cols_a_eliminar if c in df.columns])
        st.session_state["df_encuesta"] = df
        # Recalcular listas de preguntas y observaciones tras eliminar
        columnas_pregunta_detectadas = [c for c in columnas_pregunta_detectadas if c not in cols_a_eliminar]
        st.success(f"Se eliminaron {len(cols_a_eliminar)} columna(s). Quedan {len(df.columns)} columnas.")

    st.markdown('<div class="section-title">🧮 Columnas detectadas como preguntas</div>', unsafe_allow_html=True)
    st.info(f"Preguntas detectadas automáticamente (Sin incluir la pregunta de satisfacción general ): {columnas_pregunta_detectadas}")
    # ── Atributos del reporte + categorías (comparten estado) ────────────────
    # El pool de arriba y las categorías se comportan como un solo conjunto:
    # asignar un atributo a una categoría lo QUITA del pool (movimiento real),
    # quitarlo de la categoría lo DEVUELVE al pool. Nada se duplica ni se pierde.
    _opts_cols = df.columns.tolist()

    # Cambio de archivo/columnas -> reiniciar pool y categorías
    _fp_cols = hash(tuple(_opts_cols))
    if st.session_state.get("_fp_cols_adicionales") != _fp_cols:
        st.session_state["_fp_cols_adicionales"] = _fp_cols
        st.session_state["cols_adicionales"] = list(columnas_pregunta_detectadas)
        for _k in [k for k in st.session_state if str(k).startswith("cat_attrs_")]:
            st.session_state[_k] = []
        st.session_state["_cat_attrs_snapshot"] = {}

    _usar_cat_estado = bool(st.session_state.get("usar_categorias", False))
    _idx_cats = sorted({int(str(k).rsplit("_", 1)[1]) for k in st.session_state
                        if str(k).startswith("cat_attrs_")})

    if not _usar_cat_estado and st.session_state.get("_usar_cat_last", False):
        # Se acaba de DESACTIVAR el modo: devolver todo al pool y limpiar categorías
        _pool = [c for c in st.session_state.get("cols_adicionales", []) if c in _opts_cols]
        for _ci in _idx_cats:
            for a in st.session_state.get(f"cat_attrs_{_ci}", []):
                if a in _opts_cols and a not in _pool:
                    _pool.append(a)
            st.session_state[f"cat_attrs_{_ci}"] = []
        st.session_state["cols_adicionales"] = _pool
        st.session_state["_cat_attrs_snapshot"] = {}
    elif _usar_cat_estado:
        # ── Reconciliación ANTES de dibujar los widgets (semántica de 'mover') ──
        _snap = st.session_state.get("_cat_attrs_snapshot", {})
        _n_cat_vig = int(st.session_state.get("n_categorias", 1) or 1)
        _cur, _recien, _devueltos = {}, {}, []
        for _ci in _idx_cats:
            _lista = [a for a in st.session_state.get(f"cat_attrs_{_ci}", []) if a in _opts_cols]
            if _ci >= _n_cat_vig:
                # categoría eliminada (bajó el número): sus atributos vuelven al pool
                _devueltos.extend(_lista)
                _cur[_ci] = []
            else:
                _cur[_ci] = _lista
            _recien[_ci] = [a for a in _cur[_ci] if a not in _snap.get(_ci, [])]
        # 1) atributo recién agregado a una categoría -> se quita de las demás
        for _ci in _idx_cats:
            _mover = {a for _cj, _ads in _recien.items() if _cj != _ci for a in _ads}
            _cur[_ci] = [a for a in _cur[_ci] if a not in _mover or a in _recien.get(_ci, [])]
        _asignados_now = {a for _l in _cur.values() for a in _l}
        # 2) atributo quitado de una categoría -> vuelve al pool
        for _ci in _idx_cats:
            for a in _snap.get(_ci, []):
                if a not in _asignados_now and a in _opts_cols:
                    _devueltos.append(a)
        # 3) actualizar pool: los asignados salen, los devueltos entran
        _pool = [c for c in st.session_state.get("cols_adicionales", []) if c in _opts_cols]
        _pool = [c for c in _pool if c not in _asignados_now]
        for a in _devueltos:
            if a not in _pool and a not in _asignados_now:
                _pool.append(a)
        st.session_state["cols_adicionales"] = _pool
        for _ci in _idx_cats:
            st.session_state[f"cat_attrs_{_ci}"] = _cur[_ci]
        st.session_state["_cat_attrs_snapshot"] = {_ci: list(_cur[_ci]) for _ci in _idx_cats}
    st.session_state["_usar_cat_last"] = _usar_cat_estado

    _lbl_extra = " — nombre = texto después del último \" - \"" if metodo == "Qualtrics" else ""
    _lbl_pool = ("🧾 Atributos sin categoría (irán a GENERAL)" if _usar_cat_estado
                 else "🧾 Selecciona columnas adicionales (opcional)") + _lbl_extra
    _pool_sel = st.multiselect(_lbl_pool, options=_opts_cols, key="cols_adicionales")

    # Lista COMPLETA de atributos del reporte = pool + asignados a categorías
    _asignados_all = []
    for _ci in _idx_cats:
        _asignados_all.extend(st.session_state.get(f"cat_attrs_{_ci}", []))
    columnas_seleccionadas = list(dict.fromkeys(list(_pool_sel) + _asignados_all))

    #--- Segmentar atributos por categorías (opcional) ---
    st.markdown('<div class="section-title">🗂️ Segmentar atributos por categorías (opcional)</div>', unsafe_allow_html=True)
    usar_categorias = st.checkbox(
        "Agrupar las tarjetas de atributos bajo encabezados de categoría (como en el informe OFE)",
        key="usar_categorias"
    )
    categorias_def = []
    if usar_categorias:
        if not columnas_seleccionadas:
            st.warning("⚠️ Primero selecciona las columnas de preguntas (atributos) justo arriba.")
        else:
            st.caption(
                "Selecciona atributos en cada categoría: **se mueven automáticamente** — desaparecen "
                "de la lista de arriba (o de la otra categoría donde estaban). Si los quitas de la "
                "categoría, vuelven a la lista de arriba. Los que queden arriba irán a **GENERAL**."
            )
            _n_cat = st.number_input(
                "¿Cuántas categorías quieres crear?", min_value=1, max_value=40,
                step=1, key="n_categorias"
            )
            for _ci in range(int(_n_cat)):
                _c1, _c2 = st.columns([1, 2])
                with _c1:
                    _nom = st.text_input(f"Nombre de la categoría {_ci + 1}", key=f"cat_nombre_{_ci}")
                with _c2:
                    _sel = st.multiselect(
                        f"Atributos de la categoría {_ci + 1}",
                        options=columnas_seleccionadas,
                        key=f"cat_attrs_{_ci}",
                        help="Al seleccionar un atributo aquí, se mueve a esta categoría."
                    )
                if _nom.strip() and _sel:
                    categorias_def.append({"nombre": _nom.strip(), "atributos": list(_sel)})
            if _pool_sel:
                st.info(
                    f"📦 {len(_pool_sel)} atributo(s) sin asignar irán a la categoría **GENERAL**: "
                    + ", ".join(list(_pool_sel)[:5]) + ("…" if len(_pool_sel) > 5 else "")
                )
            else:
                st.success("✅ Todos los atributos están asignados a una categoría.")
    st.session_state["categorias_def"] = categorias_def if usar_categorias else []
    st.session_state["usar_categorias_flag"] = bool(usar_categorias)

    # --- Detectar columnas de observaciones ---
    palabras_clave_obs = ["comentario", "sugerencia", "observacion"]
    for col in df.columns:
        if any(palabra in col.lower() for palabra in palabras_clave_obs):
            columnas_observaciones_detectadas.append(col)

    st.markdown('<div class="section-title">📝 Columnas de comentarios u observaciones</div>', unsafe_allow_html=True)
    columnas_observaciones = st.multiselect("✏️ Selecciona columnas de observación", options=df.columns.tolist(), default=columnas_observaciones_detectadas)

    #st.markdown(f"**📌 Columna general detectada:** `{nombre_columna_general}`")
    # Validar que el default exista en las columnas
    default_general = [nombre_columna_general] if nombre_columna_general and nombre_columna_general in df.columns else []
    columna_general_seleccion = st.multiselect(
        "📌 Selecciona la columna general (opcional)"
        + (" — nombre = texto después del último \" - \"" if metodo == "Qualtrics" else ""),
        options=df.columns.tolist(), default=default_general
    )
    nombre_columna_general = columna_general_seleccion[0] if columna_general_seleccion else ""
        
   # Guardar en session_state para que otras páginas puedan acceder
    st.session_state["oficina_seleccionada"]=oficina_seleccionada
    st.session_state["proceso_seleccionado"]=proceso_seleccionado
    st.session_state["df_encuesta"] = df
    st.session_state["columnas_seleccionadas"] = columnas_seleccionadas
    st.session_state["columnas_observaciones"] = columnas_observaciones
    st.session_state["nombre_columna_general"] = nombre_columna_general

    #---Sección de seleccionar columnas para filtros dinámicos (slicers)---------------------------
    st.markdown('<div class="section-title">🎛️ Seleccionar columnas para filtros dinámicos (opcional)</div>', unsafe_allow_html=True)
    st.info("Los filtros dinámicos permiten segmentar los datos en la hoja T+G (ej: por programa, sede, tipo de estudiante, etc.)")
    columnas_filtros_dinamicos = st.multiselect("📊 Selecciona columnas para crear filtros desplegables en T+G:", options=df.columns.tolist())
    st.session_state["columnas_filtros_dinamicos"] = columnas_filtros_dinamicos

    # --- Sección: Segmentación por población ------------------------------------------------
    st.markdown('<div class="section-title">🏫 Segmentación por población (opcional)</div>', unsafe_allow_html=True)
    st.caption(
        "Úsalo cuando el proceso abarca poblaciones con tamaños distintos "
        "(ej: pregrado, posgrado, docentes). Selecciona la columna que las identifica "
        "y escribe el N de cada grupo. La ficha técnica y la muestra mínima se actualizan "
        "dinámicamente según el filtro activo en el Excel."
    )
    _opciones_seg = ["(No segmentar)"] + df.columns.tolist()
    col_segmentacion_ui = st.selectbox(
        "Columna que identifica la población:",
        options=_opciones_seg,
        index=0,
        key="col_seg_select"
    )

    poblaciones_por_segmento_ui = {}
    if col_segmentacion_ui and col_segmentacion_ui != "(No segmentar)" and col_segmentacion_ui in df.columns:
        # Agregar automáticamente al filtro dinámico si aún no está
        if col_segmentacion_ui not in columnas_filtros_dinamicos:
            columnas_filtros_dinamicos = [col_segmentacion_ui] + [c for c in columnas_filtros_dinamicos if c != col_segmentacion_ui]
            st.session_state["columnas_filtros_dinamicos"] = columnas_filtros_dinamicos
            st.caption(f"✅ Se agregó **{col_segmentacion_ui}** como filtro dinámico automáticamente.")

        valores_seg = sorted(df[col_segmentacion_ui].fillna("Sin especificar").astype(str).unique().tolist())
        st.write(f"Ingresa el **N (población total)** de cada segmento:")
        _ncols = min(len(valores_seg), 3)
        _cols_inputs = st.columns(_ncols)
        for _i, _val in enumerate(valores_seg):
            with _cols_inputs[_i % _ncols]:
                _n_val = st.number_input(
                    f"N — {_val}",
                    min_value=1, step=1,
                    value=int(numerodepoblacion),
                    key=f"n_seg_{_val}"
                )
                poblaciones_por_segmento_ui[str(_val)] = int(_n_val)

        _total_segmentos = sum(poblaciones_por_segmento_ui.values())
        st.info(
            f"📊 **Población total automática:** {_total_segmentos:,} "
            f"(suma de los segmentos: {' + '.join(f'{v:,}' for v in poblaciones_por_segmento_ui.values())} = {_total_segmentos:,}). "
            "El campo **'Número de población'** de arriba se ignora cuando hay segmentación activa."
        )

        st.session_state["col_segmentacion"] = col_segmentacion_ui
        st.session_state["poblaciones_por_segmento"] = poblaciones_por_segmento_ui
        st.session_state["n_poblacion_efectivo"] = _total_segmentos
    else:
        st.session_state["col_segmentacion"] = ""
        st.session_state["poblaciones_por_segmento"] = {}
        st.session_state["n_poblacion_efectivo"] = int(numerodepoblacion)

    #---Sección de selcionar graficas---------------------------
    st.markdown('<div class="section-title">📊 Seleccionar gráficas (opcional)</div>', unsafe_allow_html=True)
    seleccionadas = st.multiselect("Selecciona las métricas que deseas visualizar:", options=df.columns.tolist())
    # Paso 2: Para cada opción seleccionada, elegir entre 'bar' o 'column'
    tipos_grafica = {}

    for item in seleccionadas:
        tipo = st.selectbox(
            f"Selecciona el tipo de gráfica para '{item}':",
            ['pie', 'column'],
            key=item  # Importante: usar key único para evitar conflictos en el renderizado
        )
        tipos_grafica[item] = tipo
    # --- Sección: Ejecutar proceso ---
st.markdown('<div class="section-title">3️⃣ Ejecutar proceso</div>', unsafe_allow_html=True)
if st.button("🚀 Ejecutar función excel_exportar"):
    if not archivo_excel:
        st.warning("⚠️ Por favor sube el archivo Excel.")
    else:
        try:
            # Obtener datos del session_state
            df = st.session_state.get("df_encuesta", pd.DataFrame())
            preguntas = st.session_state.get("columnas_seleccionadas", [])
            comentarios = st.session_state.get("columnas_observaciones", [])
            general = st.session_state.get("nombre_columna_general", "")
            
            # Usar nombre personalizado si está disponible, sino usar oficina_seleccionada
            oficina = st.session_state.get("nombre_oficina_personalizado", oficina_seleccionada)
            
            proceso = proceso_seleccionado
            script_name = diccionario_oficinas[oficina_seleccionada]["script"]
            ruta_script = f"Script de los formatos/{script_name}.py"

            if not os.path.isfile(ruta_script):
                st.error(f"❌ El script '{ruta_script}' no existe.")
            else:
                spec = importlib.util.spec_from_file_location("modulo_dinamico", ruta_script)
                modulo = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(modulo)

                if hasattr(modulo, "excel_exportar"):
                    # Con filtros + VBA el script guarda como .xlsm; sin VBA como .xlsx
                    ruta_salida = f"{nombre_archivo}.xlsx"
                    # Si es Operaciones Tic, poner 'N_Caso' (o la columna que corresponda) de primera
                    if oficina_seleccionada == "Operaciones Tic":
                        col_primera = None
                        for col in df.columns:
                            if col.strip().lower() == "n_caso" or col.strip().lower() == "n caso":
                                col_primera = col
                                break
                        if col_primera:
                            otras = [c for c in df.columns if c != col_primera]
                            df = df[[col_primera] + otras]
                    
                    # Obtener filtros dinámicos (disponibles para TODAS las oficinas)
                    filtros_dinamicos = st.session_state.get("columnas_filtros_dinamicos", [])

                    # Categorías de atributos (opcional): lista de {nombre, atributos}
                    _usar_cat = st.session_state.get("usar_categorias_flag", False)
                    _categorias = st.session_state.get("categorias_def", []) if _usar_cat else None

                    # Sanitizar nombres de columna: eliminar caracteres que rompen
                    # referencias estructurales Excel TB[col] ([, ], #, &)
                    def _sanitize_col(c):
                        for ch in ['[', ']', '#', '&']:
                            c = str(c).replace(ch, '')
                        return c.strip() or 'Columna'
                    _rename_map = {c: _sanitize_col(c) for c in df.columns if c != _sanitize_col(c)}
                    if _rename_map:
                        df = df.rename(columns=_rename_map)
                        preguntas   = [_rename_map.get(p, p) for p in preguntas]
                        comentarios = [_rename_map.get(c, c) for c in comentarios]
                        general     = _rename_map.get(general, general)
                        filtros_dinamicos   = [_rename_map.get(f, f) for f in filtros_dinamicos]
                        if _categorias:
                            _categorias = [
                                {"nombre": _c.get("nombre", ""),
                                 "atributos": [_rename_map.get(a, a) for a in _c.get("atributos", [])]}
                                for _c in _categorias
                            ]

                    # Llamar excel_exportar para TODAS las oficinas con soporte de filtros
                    import io, contextlib
                    log_buffer = io.StringIO()
                    with st.spinner("Generando archivo Excel..." + (" y slicers..." if filtros_dinamicos else "")):
                        _n_efectivo = st.session_state.get("n_poblacion_efectivo", int(numerodepoblacion))
                        with contextlib.redirect_stdout(log_buffer):
                            modulo.excel_exportar(
                                df, nombre_archivo, _n_efectivo, preguntas, comentarios,
                                general, oficina, proceso, periodo_unico, tipos_grafica,
                                filtros_dinamicos,
                                col_segmentacion=st.session_state.get("col_segmentacion", ""),
                                poblaciones_por_segmento=st.session_state.get("poblaciones_por_segmento", {}),
                                categorias=_categorias
                            )
                    # Si se generó .xlsm (VBA + dropdowns), apuntar al nuevo archivo
                    ruta_xlsm = f"{nombre_archivo}.xlsm"
                    if filtros_dinamicos and os.path.exists(ruta_xlsm):
                        ruta_salida = ruta_xlsm
                    st.session_state["ruta_archivo_generado"] = ruta_salida
                    st.success(f"✅ Función ejecutada y archivo generado como '{os.path.basename(ruta_salida)}'")
                    # Mostrar log de slicers si hay filtros (para detectar errores)
                    if filtros_dinamicos:
                        log_texto = log_buffer.getvalue()
                        if "❌" in log_texto or "Error" in log_texto:
                            st.warning("⚠️ Hubo un problema al crear los segmentadores. Detalle:")
                            st.code(log_texto)
                        elif "✅ PROCESO COMPLETADO" in log_texto:
                            st.info("🎛️ Segmentadores creados correctamente en la hoja T+G")
                else:
                    st.error("❌ El script no contiene una función llamada 'excel_exportar'.")
        except Exception as e:
            st.error(f"🚨 Error al ejecutar: {e}")

# --- Botón de descarga si se ha generado el archivo ---
if "ruta_archivo_generado" in st.session_state:
    ruta_archivo_generado = st.session_state["ruta_archivo_generado"]
    if os.path.exists(ruta_archivo_generado):
        with open(ruta_archivo_generado, "rb") as f:
            st.download_button(
                label="📥 Descargar archivo generado",
                data=f,
                file_name=os.path.basename(ruta_archivo_generado),
                mime=("application/vnd.ms-excel.sheet.macroEnabled.12"
                  if ruta_archivo_generado.endswith(".xlsm")
                  else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            )

# --- Pie de página ---
st.markdown('<div class="footer">Desarrollado con ❤️ usando Streamlit · Universidad del Norte · 2025</div>', unsafe_allow_html=True)
