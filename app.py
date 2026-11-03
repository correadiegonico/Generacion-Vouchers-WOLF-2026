import os
import io
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A5, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

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

# --- FUNCIÓN GENERADORA DE PDF EN MEMORIA ---
def generar_pdf_voucher(tipo_servicio, datos):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A5),
        rightMargin=1*cm, leftMargin=1*cm,
        topMargin=1*cm, bottomMargin=1*cm
    )
    styles = getSampleStyleSheet()
    story = []

    # Estilos
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor('#1E3A8A'))
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Heading2'], fontSize=12, leading=14, textColor=colors.HexColor('#374151'))
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=10, leading=13)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=10, leading=13, fontName='Helvetica-Bold')

    # Encabezado
    story.append(Paragraph("<b>WOLF TRAVEL 2026</b>", title_style))
    story.append(Paragraph(f"VOUCHER DE SERVICIO - {tipo_servicio.upper()}", subtitle_style))
    story.append(Spacer(1, 0.4*cm))

    # Construcción de la tabla de datos
    tabla_datos = []
    for k, v in datos.items():
        tabla_datos.append([Paragraph(f"<b>{k}:</b>", bold_style), Paragraph(str(v), body_style)])

    t = Table(tabla_datos, colWidths=[5*cm, 12*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.8*cm))

    # Bloque de Firmas y Aclaración
    firmas_data = [
        ["........................................................................", "........................................................................"],
        ["Firma y Aclaración (Coordinador)", "Firma y Aclaración (Proveedor)"]
    ]
    t_firmas = Table(firmas_data, colWidths=[8.5*cm, 8.5*cm])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TEXTCOLOR', (0,1), (-1,1), colors.HexColor('#4B5563')),
    ]))
    story.append(t_firmas)

    doc.build(story)
    buffer.seek(0)
    return buffer

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
    
    # Filtrar únicamente Estado == "Activo" (Columna K)
    col_estado = df_pasajeros.columns[10] if len(df_pasajeros.columns) > 10 else 'Estado'
    df_activos = df_pasajeros[df_pasajeros[col_estado].astype(str).str.strip().str.upper() == 'ACTIVO'].copy()
    
    coches_disponibles = sorted([c for c in df_activos['Coche_Clean'].unique() if c != ''], key=lambda x: (int(x) if x.isdigit() else x))
    
    col_input, col_select = st.columns([2, 1])
    with col_input:
        nro_coche_input = st.text_input("Ingrese el Número de Coche exacto:")
    with col_select:
        coche_seleccionado = st.selectbox("O seleccione un Coche activo:", ["-- Seleccionar --"] + coches_disponibles)
        
    coche_final = nro_coche_input.strip() if nro_coche_input.strip() else (coche_seleccionado if coche_seleccionado != "-- Seleccionar --" else "")

    if coche_final:
        df_coche = df_activos[df_activos['Coche_Clean'] == coche_final]
        
        if not df_coche.empty:
            # Columna P (16ª columna, índice 15) para Categoría de Acompañantes
            col_categoria = df_coche.columns[15] if len(df_coche.columns) > 15 else 'Categoria'
            
            mask_acomp = df_coche[col_categoria].astype(str).str.upper().str.contains('ACOMPAÑANTE|ACOMPANANTE|PADRE|MADRE', na=False)
            cant_acomp_solo = len(df_coche[mask_acomp])
            cant_pax_solo = len(df_coche[~mask_acomp])
            cant_pax_total = len(df_coche)
            
            if 'Colegio' in df_coche.columns and 'Codigo' in df_coche.columns:
                resumen_colegios = df_coche.groupby(['Colegio', 'Codigo']).size().reset_index(name='Cantidad de Pax Activos')
                cant_contratos = len(resumen_colegios)
            else:
                cant_contratos = 1
                resumen_colegios = pd.DataFrame()

            # Métricas
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                st.metric(label="Cantidad de Contratos", value=cant_contratos)
            with c_m2:
                st.metric(label="Pasajeros (Pax)", value=cant_pax_solo)
            with c_m3:
                st.metric(label="Acompañantes", value=cant_acomp_solo)
                
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
                st.write("")
                emitir_todos = st.checkbox("Emitir todos los vouchers del coche")

            # Extracción dinámica de actividades de Desglose de Servicios
            opciones_actividades = []
            cols_desglose = [str(c).strip() for c in df_desglose.columns]
            
            # Buscar celda "Hotel" en cabecera
            idx_hotel = -1
            for i, col_name in enumerate(cols_desglose):
                if 'HOTEL' in col_name.upper():
                    idx_hotel = i
                    break

            if idx_hotel != -1:
                brutas = cols_desglose[idx_hotel + 1:]
                desestimar = {'TOTAL MERIEN', 'TOTAL 5TA', 'REGRESO BUS', 'HTL PARANORMAL'}
                opciones_actividades = [col for col in brutas if col.upper() not in desestimar and not col.startswith('Unnamed')]

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
            
            # Diccionario de datos para compilar el PDF
            datos_voucher = {"Coche": coche_final}
            
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

                datos_voucher.update({
                    "Hotel": hotel_nombre,
                    "Fecha IN": str(fecha_in),
                    "Fecha OUT": str(fecha_out),
                    "Tipo Habitación": ", ".join(tipo_habitacion),
                    "Tripulación": tripulacion,
                    "Tripulación Liberada": tripulacion_liberada,
                    "Observaciones": observaciones
                })

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

                datos_voucher.update({
                    "Coordinador": coordinador,
                    "Almuerzos": cant_almuerzos,
                    "Cenas": cant_cenas,
                    "Quinta Comida": cant_quinta,
                    "Observaciones": observaciones
                })

            else:  # Actividades
                st.info("📌 Modelo activo: Voucher de Actividades")
                coordinador = st.text_input("Coordinador a cargo:")
                
                st.markdown("##### Distribución y Conteo de Asistentes:")
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    pax_val = st.number_input("Pax:", min_value=0, value=cant_pax_solo, step=1)
                with c2:
                    acomp_val = st.number_input("Acompañantes:", min_value=0, value=cant_acomp_solo, step=1)
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

                datos_voucher.update({
                    "Actividad / Servicio": servicio_especifico if not emitir_todos else "TODAS LAS ACTIVIDADES DEL COCHE",
                    "Coordinador a cargo": coordinador,
                    "Pax": pax_val,
                    "Acompañantes": acomp_val,
                    "Coordinadores": coord_val,
                    "Choferes": choferes_val,
                    "Staff Extra": staff_val,
                    "TOTAL GENERAL": total_general,
                    "Liberados": liberados,
                    "Observaciones": observaciones
                })

            st.markdown("---")
            
            # --- BOTÓN DE GENERACIÓN Y DESCARGA DIRECTA DE PDF ---
            pdf_bytes = generar_pdf_voucher(tipo_servicio, datos_voucher)
            
            st.download_button(
                label="🖨️ Generar y Descargar Voucher PDF",
                data=pdf_bytes,
                file_name=f"Voucher_{tipo_servicio}_Coche_{coche_final}.pdf",
                mime="application/pdf"
            )

        else:
            st.warning(f"No se encontraron pasajeros ACTIVOS asignados exactamente al coche '{coche_final}'.")
else:
    st.error("No se encontró la columna 'Coche' en la planilla de Control de Pasajeros.")
