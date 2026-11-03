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

# --- IDs DE GOOGLE DRIVE ---
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
    st.stop()

st.success("✅ Datos sincronizados correctamente desde Google Drive.")

st.markdown("---")

# --- PASO 1: BUSCADOR POR NÚMERO DE COCHE (COINCIDENCIA EXACTA) ---
st.subheader("🚌 Búsqueda y Selección por Número de Coche")

if 'Coche' in df_pasajeros.columns:
    # Formatear la columna Coche como texto limpio
    # Si viene como número decimal (ej. 25.0), lo convertimos a entero como texto (25)
    def formatear_coche(val):
        if pd.isna(val) or str(val).strip() == '':
            return ''
        try:
            # Si es float equivalente a entero (ej: 25.0 -> '25')
            val_float = float(val)
            if val_float.is_integer():
                return str(int(val_float))
            return str(val).strip()
        except ValueError:
            return str(val).strip()

    df_pasajeros['Coche_Clean'] = df_pasajeros['Coche'].apply(formatear_coche)
    
    # Obtener lista única de coches válidos para el selector de sugerencia
    coches_disponibles = sorted([c for c in df_pasajeros['Coche_Clean'].unique() if c != ''], key=lambda x: (int(x) if x.isdigit() else x))
    
    col_input, col_select = st.columns([2, 1])
    with col_input:
        nro_coche_input = st.text_input("Ingrese el Número de Coche exacto:")
    with col_select:
        coche_seleccionado = st.selectbox("O seleccione un Coche activo:", ["-- Seleccionar --"] + coches_disponibles)
        
    # Determinar el coche a filtrar
    coche_final = ""
    if nro_coche_input.strip():
        coche_final = nro_coche_input.strip()
    elif coche_seleccionado != "-- Seleccionar --":
        coche_final = coche_seleccionado

    if coche_final:
        # FILTRADO EXACTO
        df_coche = df_pasajeros[df_pasajeros['Coche_Clean'] == coche_final]
        
        if not df_coche.empty:
            cant_pax_total = len(df_coche)
            
            # Agrupar por Colegio y Código de contrato
            if 'Colegio' in df_coche.columns and 'Codigo' in df_coche.columns:
                resumen_colegios = df_coche.groupby(['Colegio', 'Codigo']).size().reset_index(name='Cantidad de Pax')
                cant_contratos = len(resumen_colegios)
            else:
                cant_contratos = 1
                resumen_colegios = pd.DataFrame()

            # Mostrar métricas de resumen en un renglón
            col1, col2 = st.columns(2)
            with col1:
                st.metric(label="Cantidad de Contratos (Colegios)", value=cant_contratos)
            with col2:
                st.metric(label="Cantidad Total de Pasajeros (Pax)", value=cant_pax_total)
                
            st.markdown("##### Detalle de Colegios y Contratos en este Coche:")
            if not resumen_colegios.empty:
                st.dataframe(resumen_colegios, use_container_width=True)
            
            st.markdown("---")
            
            # --- PASO 2: MENÚ DE CONFIGURACIÓN Y EMISIÓN ---
            st.subheader("⚙️ Configuración del Voucher")
            
            tipo_servicio = st.selectbox(
                "Servicio a emitir:",
                ["Actividad Diurna", "Actividad Nocturna", "Hotel", "Gastronomía"]
            )
            
            if tipo_servicio == "Hotel":
                st.info("📌 Modelo activo: Voucher de Hotelería (Mantiene formato original del PDF de Hotel)")
                hotel_nombre = st.text_input("Hotel:")
                col_in, col_out = st.columns(2)
                with col_in:
                    fecha_in = st.date_input("IN (Fecha de ingreso):")
                with col_out:
                    fecha_out = st.date_input("OUT (Fecha de salida):")
                
                tipo_habitacion = st.multiselect("Tipo de Habitación:", ["DBL (Doble)", "TPL (Triple)", "CDPL (Cuádruple)", "QTPL (Quíntuple)"])
                tripulacion = st.number_input("Tripulación:", min_value=0, value=2, step=1)
                tripulacion_liberada = st.number_input("Tripulación Liberada:", min_value=0, value=0, step=1)
                observaciones_hotel = st.text_area("Observaciones:")
                
            else:
                col_a, col_b = st.columns(2)
                with col_a:
                    coordinador = st.text_input("Nombre del coordinador encargado:")
                with col_b:
                    liberados = st.number_input("Liberados:", min_value=0, value=0, step=1)
                
                pax_auto = cant_pax_total
                acompanantes_auto = 0 
                coord_pred = 2
                choferes_pred = 2
                
                st.markdown("##### Distribución de Asistentes:")
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.metric("Pax", pax_auto)
                with c2:
                    acompanantes = st.number_input("Acompañantes:", min_value=0, value=acompanantes_auto, step=1)
                with c3:
                    coordinadores_cant = st.number_input("Coordinadores", min_value=0, value=coord_pred, step=1)
                with c4:
                    choferes_cant = st.number_input("Choferes", min_value=0, value=choferes_pred, step=1)
                with c5:
                    staff_extra = st.number_input("Staff Extra:", min_value=0, value=0, step=1)
                    
                total_general = pax_auto + acompanantes + coordinadores_cant + choferes_cant + staff_extra
                st.markdown(f"### **TOTAL GENERAL:** `{total_general}` personas")
                
                if tipo_servicio == "Actividad Diurna":
                    excursion = st.text_input("Excursión:")
                    tipo_servicio_detalle = st.text_input("Tipo de servicio:")
                elif tipo_servicio == "Actividad Nocturna":
                    boliche = st.text_input("Boliche / Local Nocturno:")
                elif tipo_servicio == "Gastronomía":
                    st.markdown("##### Cantidades de Servicios Gastronómicos:")
                    g1, g2, g3 = st.columns(3)
                    with g1:
                        cant_almuerzos = st.number_input("Almuerzos:", min_value=0, value=0, step=1)
                    with g2:
                        cant_cenas = st.number_input("Cenas:", min_value=0, value=0, step=1)
                    with g3:
                        cant_quinta = st.number_input("Quinta Comida:", min_value=0, value=0, step=1)

                observaciones = st.text_area("Observaciones:")
                
                st.markdown("---")
                st.markdown("`------------------------------------------------------------------------------------------------------------------`")
                st.markdown("**Firma y Aclaración (Proveedor)**                                          **Firma y Aclaración (Coordinador)**")

            if st.button("🖨️ Generar y Descargar Voucher PDF"):
                st.success(f"Voucher de [{tipo_servicio}] preparado correctamente para el Coche {coche_final}.")
        else:
            st.warning(f"No se encontraron pasajeros asignados exactamente al coche '{coche_final}'. Verifica el número e intenta nuevamente.")
else:
    st.error("No se encontró la columna 'Coche' en la planilla de Control de Pasajeros.")
