import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata
import os
import zipfile

# --- 1. ESTILOS VISUALES (Elegancia Restaurada y Sobria) ---
st.set_page_config(page_title="Sistema IRIDEM", page_icon="📄", layout="centered")

st.markdown("""
    <style>
    /* Fondo general oscuro */
    .stApp {
        background-color: #424549;
    }
    
    /* Tarjeta central - VERDE LIMÓN */
    .block-container {
        background-color: #a8cf45;
        padding: 3rem;
        border-radius: 20px;
        margin-top: 2rem;
        box-shadow: 0 10px 25px rgba(0,0,0,0.4);
    }
    
    /* TÍTULOS EN BLANCO CON SOMBRA */
    h1, h2 {
        color: #ffffff !important;
        text-align: center;
        font-family: 'Helvetica', sans-serif;
        font-weight: 800;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        margin-bottom: 0px !important;
        border: none !important;
    }
    h1 { font-size: 2.1rem !important; padding-bottom: 5px !important; }
    h2 { font-size: 1.8rem !important; margin-top: -10px !important; padding-bottom: 20px !important;}

    /* ESTILO DE LAS PESTAÑAS (TABS) */
    .stTabs [data-baseweb="tab-list"] {
        background-color: transparent;
        justify-content: center;
    }
    .stTabs [data-baseweb="tab"] {
        color: #ffffff !important;
        font-weight: bold !important;
        font-size: 1.1rem;
    }
    /* LÍNEA DE SELECCIÓN NEGRA */
    .stTabs [data-baseweb="tab-highlight"] {
        background-color: #000000 !important;
    }

    /* CAJA DE CARGA CLARA (Gris azulado) */
    [data-testid="stFileUploader"] {
        background-color: #eff4f9 !important;
        padding: 1.5rem;
        border-radius: 12px;
        border: none !important;
        margin-top: 1rem;
    }

    /* BOTÓN DE UPLOAD BLANCO */
    [data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #31333f !important;
        border: 1px solid #d3dae0 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }

    /* Textos internos en color oscuro */
    [data-testid="stFileUploader"] label, 
    [data-testid="stFileUploader"] p, 
    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] span {
        color: #555555 !important;
    }

    /* Botones de acción principales (Negro sólido) */
    div.stButton > button {
        background-color: #1e1e1e;
        color: white;
        width: 100%;
        border-radius: 10px;
        height: 3.5em;
        font-weight: bold;
        border: 1px solid #ffffff;
        margin-top: 1rem;
        text-transform: uppercase;
    }
    
    /* Botón Descargar (Gris oscuro) */
    div.stDownloadButton > button {
        background-color: #555555;
        color: white;
        width: 100%;
        border-radius: 10px;
        height: 3.5em;
        font-weight: bold;
        border: none;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE LIMPIEZA Y GRAMÁTICA (CÓDIGO MAESTRO) ---
def limpiar_y_asegurar_unicos(columnas):
    nombres_limpios = []
    for i, col in enumerate(columnas):
        texto = str(col).lower()
        texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
        texto = texto.replace(':', '').replace('/', '_').replace('.', '').replace('°', '')
        texto = re.sub(r'[^a-z0-9\s_]', '', texto)
        nuevo_nombre = "_".join(texto.split())
        if nuevo_nombre in nombres_limpios:
            nuevo_nombre = f"{nuevo_nombre}_{i}"
        nombres_limpios.append(nuevo_nombre)
    return nombres_limpios

def corregir_mayusculas(texto, es_nombre=False):
    if not isinstance(texto, str) or texto == "-": return texto
    texto = texto.lower().strip()
    if not texto: return "-"
    if es_nombre: return texto.title()
    return re.sub(r'(^|[.!?]\s+)(\w)', lambda m: m.group(1) + m.group(2).upper(), texto)

def limpiar_dato(dato, nombre_columna):
    if isinstance(dato, float) and dato.is_integer(): dato = int(dato)
    s = str(dato).strip()
    if s in ['0', '0.0', '00:00:00', 'nan', 'NaT', 'None', '1900-01-01', '01/01/1900']: return "-"
    if len(s) >= 10 and re.match(r'\d{4}-\d{2}-\d{2}', s):
        try: return pd.to_datetime(s).strftime('%d/%m/%Y')
        except: return s
    es_col_nombre = any(x in nombre_columna for x in ['nombre', 'apellido', 'paterno', 'materno'])
    return corregir_mayusculas(s, es_nombre=es_col_nombre)

def extraer_objetivo_al_inicio(texto):
    texto = str(texto).strip()
    match = re.match(r'^\((.*?)\)', texto)
    if match: return corregir_mayusculas(match.group(1).strip())
    return "-"

def limpiar_descripcion_original(texto):
    texto = str(texto).strip()
    return re.sub(r'^\((.*?)\)', '', texto).strip()

def formatear_fecha_larga(valor):
    try:
        meses = {1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio", 7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"}
        dias = {0:"lunes", 1:"martes", 2:"miércoles", 3:"jueves", 4:"viernes", 5:"sábado", 6:"domingo"}
        dt = pd.to_datetime(valor, dayfirst=True)
        return f"{dias[dt.weekday()]}, {dt.day} de {meses[dt.month]} de {dt.year}"
    except: return valor

# --- 3. LÓGICA DE LA PÁGINA ---
st.markdown("<h1>Sistema de Automatización Documental IRIDEM y Reportes de Intervención</h1>", unsafe_allow_html=True)
st.markdown("<h2>FAE DEM Cerrillos</h2>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["Informe de evaluación IRIDEM", "Registro de intervención"])

with tab1:
    excel_1 = st.file_uploader("Seleccionar archivo Excel para Evaluación", type=["xlsx"], key="ind")
    if excel_1:
        df1 = pd.read_excel(excel_1)
        df1.columns = limpiar_y_asegurar_unicos(df1.columns)
        if 'descripcionevento' in df1.columns:
            df1['objetivo'] = df1['descripcionevento'].apply(extraer_objetivo_al_inicio)
            df1['descripcionevento'] = df1['descripcionevento'].apply(limpiar_descripcion_original)
        for col in df1.columns: df1[col] = df1.apply(lambda row: limpiar_dato(row[col], col), axis=1)
        if 'fecha' in df1.columns: df1['fecha'] = df1['fecha'].apply(formatear_fecha_larga)
        
        datos1 = df1.iloc[0].to_dict()
        fn = f"{datos1.get('nombres', '')} {datos1.get('apellido_paterno', '')} {datos1.get('apellido_materno', '')}".strip()
        fn = fn if fn else "Reporte"
        if st.button(f"GENERAR INFORME DE: {fn}", key="btn_ind"):
            if os.path.exists("plantilla.docx"):
                doc = DocxTemplate("plantilla.docx")
                doc.render(datos1)
                out = io.BytesIO(); doc.save(out); out.seek(0)
                st.success("✅ Documento listo!")
                st.download_button("GUARDAR WORD", out, f"Informe_{fn}.docx", key="dl_ind")
            else: st.error("No se encontró 'plantilla.docx'")

with tab2:
    excel_2 = st.file_uploader("Seleccionar archivo Excel para Registros", type=["xlsx"], key="mas")
    if excel_2:
        df2 = pd.read_excel(excel_2)
        df2.columns = limpiar_y_asegurar_unicos(df2.columns)
        if 'descripcionevento' in df2.columns:
            df2['objetivo'] = df2['descripcionevento'].apply(extraer_objetivo_al_inicio)
            df2['descripcionevento'] = df2['descripcionevento'].apply(limpiar_descripcion_original)
        for col in df2.columns: df2[col] = df2.apply(lambda row: limpiar_dato(row[col], col), axis=1)
        if 'fecha' in df2.columns: df2['fecha'] = df2['fecha'].apply(formatear_fecha_larga)
        
        if st.button("GENERAR REGISTROS MASIVOS", key="btn_mas"):
            if os.path.exists("plantilla_2.docx"):
                buf_z = io.BytesIO()
                with zipfile.ZipFile(buf_z, "w") as zf:
                    for i, fila in df2.iterrows():
                        doc = DocxTemplate("plantilla_2.docx")
                        df_f = fila.to_dict()
                        doc.render(df_f)
                        buf_w = io.BytesIO(); doc.save(buf_w)
                        name = f"{df_f.get('nombres', '')}_{df_f.get('apellido_paterno', '')}_{df_f.get('apellido_materno', '')}".strip()
                        name = name if name else f"Registro_{i+1}"
                        zf.writestr(f"Registro_{name}.docx", buf_w.getvalue())
                buf_z.seek(0)
                st.success("✅ Registros procesados!")
                st.download_button("DESCARGAR ZIP", buf_z, "registros_intervencion.zip", key="dl_mas")
            else: st.error("No se encontró 'plantilla_2.docx'")