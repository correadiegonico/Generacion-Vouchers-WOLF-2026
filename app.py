import os
import io
import pandas as pd
import streamlit as st

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Wolf Travel 2026 - Generador de Vouchers",
    page_icon="🐺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- IDs DE GOOGLE DRIVE (NUEVOS) ---
ID_CONTROL_PASAJEROS = "1cfMOu1NAvfvbiuDB7QRSvLsMdui4LMqfFGciptQr5Xw"
ID_DESGLOSE_SERVICIOS = "1ObkoFqie4MV7xZ2PL9I1wqKXZm-qRjnrfAllT_fn2s4"

URL_PASAJEROS = f"https://docs.google.com/spreadsheets/d/{ID_CONTROL_PASAJEROS}/export?format=xlsx"
URL_DESGLOSE = f"https://docs.google.com/spreadsheets/d/{ID_DESGLOSE_SERVICIOS}/export?format=xlsx"

# --- FUNCIONES DE CARGA AUTOMÁTICA DESDE DRIVE ---
@st.cache_data(ttl=300)
def cargar_datos_drive():
    try:
        df_pasajeros = pd.read_excel(URL_PASAJEROS)
        df_desglose = pd.read_excel(URL_DESGLOSE)
        return df_pasajeros, df_desglose, None
    except Exception as e:
        return None, None, str(e)

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.title("🐺 Wolf Travel 2026")
    st.markdown("---")
    st.subheader("Archivos de Origen (Google Drive)")
    
    st.markdown("🟢 **Control de Pasajeros**")
    st.caption(f"[Abrir en Google Drive](https://docs.google.com/spreadsheets/d/{ID_CONTROL_PASAJEROS}/edit)")
    
    st.markdown("🟢 **Desglose de Servicios 2026**")
    st.caption(f"[Abrir en Google Drive](https://docs.google.com/spreadsheets/d/{ID_DESGLOSE_SERVICIOS}/edit)")

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
    st.warning("Verifica que las planillas tengan activado el acceso por enlace ('Cualquier persona con el enlace puede ver').")
    st.stop()

st.success("✅ Datos sincronizados correctamente desde Google Drive.")
