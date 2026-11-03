import os
import re
import io
from datetime import datetime
import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Wolf Travel 2026 - Generador de Vouchers",
    page_icon="🐺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CONFIGURACIÓN DE GOOGLE DRIVE ---
ID_CONTROL_PASAJEROS = "1bg6yaJ1V4pd2g_5rufnByz3sVe-ovGP6ckWYEB62l0g"
ID_DESGLOSE_SERVICIOS = "1baRRKiC_-fZb5t_FmQtM-1kMK7epKHx4zckGlCVkhHo"

URL_PASAJEROS = f"https://docs.google.com/spreadsheets/d/{ID_CONTROL_PASAJEROS}/export?format=xlsx"
URL_DESGLOSE = f"https://docs.google.com/spreadsheets/d/{ID_DESGLOSE_SERVICIOS}/export?format=xlsx"

# --- FUNCIONES DE CARGA AUTOMÁTICA DESDE DRIVE ---
@st.cache_data(ttl=300)  # Recarga automáticamente los datos de Drive cada 5 minutos
def cargar_datos_drive():
    try:
        df_pasajeros = pd.read_excel(URL_PASAJEROS)
        df_desglose = pd.read_excel(URL_DESGLOSE)
        return df_pasajeros, df_desglose, None
    except Exception as e:
        return None, None, str(e)

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/616/616408.png", width=60) # Icono lobo de referencia o título
    st.title("Wolf Travel 2026")
    
    st.markdown("---")
    st.subheader("Archivos de Origen (Google Drive)")
    
    # Enlaces informativos sin opción de carga manual
    st.markdown("🟢 **Control de Pasajeros**")
    st.caption("[Abrir en Google Drive](https://docs.google.com/spreadsheets/d/1bg6yaJ1V4pd2g_5rufnByz3sVe-ovGP6ckWYEB62l0g/edit?usp=sharing)")
    
    st.markdown("🟢 **Desglose de Servicios 2026**")
    st.caption("[Abrir en Google Drive](https://docs.google.com/spreadsheets/d/1baRRKiC_-fZb5t_FmQtM-1kMK7epKHx4zckGlCVkhHo/edit?usp=sharing)")

    st.markdown("---")
    if st.button("🔄 Sincronizar datos de Drive"):
        st.cache_data.clear()
        st.rerun()

# --- HEADER PRINCIPAL ---
st.title("🐺 Wolf Travel 2026 - Generador de Vouchers")

# Cargar datos
with st.spinner("Conectando y descargando datos desde Google Drive..."):
    df_pasajeros, df_desglose, error = cargar_datos_drive()

if error:
    st.error(f"Error al conectar con Google Drive: {error}")
    st.warning("Verifica que las planillas en Google Drive tengan activado el acceso por enlace ('Cualquier persona con el enlace puede ver').")
    st.stop()

st.success("✅ Datos sincronizados correctamente desde Google Drive.")

# --- AQUÍ CONTINÚA LA LÓGICA DE TU INTERFAZ DE GENERACIÓN DE VOUCHERS Y PDF ---
# (Formularios de Staff Empresa, TOTAL calculado, Observaciones y st.download_button para los PDF)