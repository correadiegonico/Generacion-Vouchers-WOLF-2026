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

# --- PASO 1: BUSCADOR POR NÚMERO DE COCHE ---
st.subheader("🚌 Búsqueda y Selección por Número de Coche")

if 'Coche' in df_pasajeros.columns:
    # Formatear la columna Coche como texto limpio
    def formatear_coche(val):
        if pd.isna(val) or str(val).strip() == '':
            return ''
        try:
            val_float = float(val)
            if val_float.is_integer():
                return str(int(val_float))
            return str(val).strip()
        except ValueError:
            return str(val).strip()

    df_pasajeros['Coche_Clean'] = df_pasajeros['Coche'].apply(formatear_coche)
    
    # Filtrar únicamente pasajeros cuyo Estado (Columna K) sea "Activo"
    col_estado = 'Estado' if 'Estado' in df_pasajeros.columns else df_pasajeros.columns[10] # Columna K es índice 10
    df_activos = df_pasajeros[df_pasajeros[col_estado].astype(str).str.strip().str.upper() == 'ACTIVO'].copy()
    
    # Lista única de coches válidos activos
    coches_disponibles = sorted([c for c in df_activos['Coche_Clean'].unique() if c != ''], key=lambda x: (int(x) if x.isdigit() else x))
    
    col_input, col_select = st.columns([2, 1])
    with col_input:
        nro_coche_input = st.text_input("Ingrese el Número de Coche exacto:")
    with col_select:
        coche_seleccionado = st.selectbox("O seleccione un Coche activo:", ["-- Seleccionar --"] + coches_disponibles)
        
    coche_final = nro_coche_input.strip() if nro_coche_input.strip() else (coche_seleccionado if coche_seleccionado != "-- Seleccionar --" else "")

    if coche_final:
        # FILTRADO EXACTO SOBRE PASAJEROS ACTIVOS
        df_coche = df_activos[df_activos['Coche_Clean'] == coche_final]
        
        if not df_coche.empty:
            # Separar Pax y Acompañantes si existe columna 'Tipo' o 'Categoria'
            col_tipo = 'Tipo' if 'Tipo' in df_coche.columns else ('Categoria' if 'Categoria' in df_coche.columns else None)
            
            if col_tipo:
                mask_acomp = df_coche[col_tipo].astype(str).str.upper().str.contains('ACOMPAÑANTE|ACOMPANANTE|PADRE|MADRE', na=False)
                cant_pax_solo = len(df_coche[~mask_acomp])
                cant_acomp_solo = len(df_coche[mask_acomp])
            else:
                cant_pax_solo = len(df_coche)
                cant_acomp_solo = 0

            cant_pax_total = len(df_coche)
            
            # Agrupar por Colegio y Código de contrato
            if 'Colegio' in df_coche.columns and 'Codigo' in df_coche.columns:
                resumen_colegios = df_coche.groupby(['Colegio', 'Codigo']).size().reset_index(name='Cantidad de Pax Activos')
                cant_contratos = len(resumen_colegios)
                codigos_contratos = resumen_colegios['Codigo'].unique().tolist()
            else:
                cant_contratos = 1
                resumen_colegios = pd.DataFrame()
                codigos_contratos = []

            # Mostrar métricas de resumen
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                st.metric(label="Cantidad de Contratos", value=cant_contratos)
            with c_m2:
                st.metric(label="Pasajeros (Pax Activos)", value=cant_pax_solo)
            with c_m3:
                st.metric(label="Acompañantes Activos", value=cant_acomp_solo)
                
            st.markdown("##### Detalle de Colegios y Contratos en este Coche:")
            if not resumen_colegios.empty:
                st.dataframe(resumen_colegios, use_container_width=True)
            
            st.markdown("---")
            
            # --- PASO 2: MENÚ DE CONFIGURACIÓN Y EMISIÓN ---
            st.subheader("⚙️ Configuración del Voucher")
            
            col_tipo_serv, col_chk = st.columns([2, 2])
            with col_tipo_serv:
                tipo_servicio = st.selectbox(
                    "Tipo de servicio:",
                    ["Actividades", "Hotel", "Gastronomía"]
                )
            with col_chk:
                st.write("") # Espaciador
                emitir_todos = st.checkbox("Emitir todos los vouchers del coche")

            # Obtención dinámica de opciones de actividades desde Desglose de Servicios
            opciones_actividades = []
            if 'Hotel' in [str(c).strip() for c in df_desglose.columns]:
                columnas_lista = [str(c).strip() for c in df_desglose.columns]
                idx_hotel = columnas_lista.index('Hotel')
                brutas = columnas_lista[idx_hotel + 1:]
                
                # Desestimar las indicadas por usuario
                desestimar = {'TOTAL MERIEN', 'TOTAL 5TA', 'REGRESO BUS', 'HTL PARANORMAL'}
                opciones_actividades = [col for col in brutas if col.upper() not in desestimar and not col.startswith('Unnamed')]

            # Selección de Servicio a emitir
            servicio_especifico = ""
            if not emitir_todos:
                if tipo_servicio == "Actividades":
                    servicio_especifico = st.selectbox(
                        "Servicio a emitir:",
                        ["-- Seleccionar Actividad --"] + opciones_actividades
                    )
                else:
                    st.text_input("Servicio a emitir:", value="", disabled=True, placeholder="No requiere selección específica")
            
            st.markdown("---")
            
            # Formulario según Tipo de Servicio elegido
            if tipo_servicio == "Hotel":
                st.info("📌 Modelo activo: Voucher de Hotelería")
                hotel_nombre = st.text_input("Hotel:")
                col_in, col_out = st.columns(2)
                with col_in:
                    fecha_in = st.date_input("IN (Fecha de ingreso):")
                with col_out:
                    fecha_out = st.date_input("OUT (Fecha de salida):")
                
                tipo_habitacion = st.multiselect("Tipo de Habitación:", ["DBL (Doble)", "TPL (Triple)", "CDPL (Cuádruple)", "QTPL (Quíntuple)"])
                tripulacion = st.number_input("Tripulación:", min_value=0, value=2, step=1)
                tripulacion_liberada = st.number_input("Tripulación Liberada:", min_value=0, value=0, step=1)
                observaciones = st.text_area("Observaciones:")

            elif tipo_servicio == "Gastronomía":
                st.info("📌 Modelo activo: Voucher de Servicios Gastronómicos")
                coordinador = st.text_input("Coordinador a cargo:")
                st.markdown("##### Cantidades de Servicios Gastronómicos:")
                g1, g2, g3 = st.columns(3)
                with g1:
                    cant_almuerzos = st.number_input("Cantidad de Almuerzos:", min_value=0, value=cant_pax_total, step=1)
                with g2:
                    cant_cenas = st.number_input("Cantidad de Cenas:", min_value=0, value=cant_pax_total, step=1)
                with g3:
                    cant_quinta = st.number_input("Cantidad de Quinta Comida:", min_value=0, value=0, step=1)

                observaciones = st.text_area("Observaciones:")

            else:  # Actividades
                st.info("📌 Modelo activo: Voucher de Actividades (Diurnas y Nocturnas unificadas)")
                
                coordinador = st.text_input("Coordinador a cargo:")
                
                st.markdown("##### Distribución y Conteo de Asistentes:")
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    pax_val = st.number_input("Pax (Automático):", min_value=0, value=cant_pax_solo, step=1)
                with c2:
                    acomp_val = st.number_input("Acompañantes (Automático):", min_value=0, value=cant_acomp_solo, step=1)
                with c3:
                    coord_val = st.number_input("Coordinadores:", min_value=0, value=2, step=1)
                with c4:
                    choferes_val = st.number_input("Choferes:", min_value=0, value=2, step=1)
                with c5:
                    staff_val = st.number_input("Staff Extra:", min_value=0, value=0, step=1)
                    
                total_general = pax_val + acomp_val + coord_val + choferes_val + staff_val
                st.markdown(f"### **TOTAL:** `{total_general}` personas")
                
                liberados = st.number_input("Liberados:", min_value=0, value=0, step=1)
                observaciones = st.text_area("Observaciones:")

            st.markdown("---")
            st.markdown("`------------------------------------------------------------------------------------------------------------------`")
            st.markdown("**Firma y Aclaración (Coordinador)**                                          **Firma y Aclaración (Proveedor)**")

            if st.button("🖨️ Generar y Descargar Voucher PDF"):
                if emitir_todos:
                    st.success(f"Emisión por LOTE iniciada: Generando todos los vouchers de [{tipo_servicio}] para el Coche {coche_final}...")
                else:
                    st.success(f"Generando voucher de [{tipo_servicio}] - {servicio_especifico} para el Coche {coche_final}...")

        else:
            st.warning(f"No se encontraron pasajeros ACTIVOS asignados exactamente al coche '{coche_final}'.")
else:
    st.error("No se encontró la columna 'Coche' en la planilla de Control de Pasajeros.")
