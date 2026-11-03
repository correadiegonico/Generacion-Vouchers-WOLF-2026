import os
import io
import zipfile
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

# --- FUNCIÓN GENERADORA DE PDF EN MEMORIA (NUEVO DISEÑO) ---
def generar_pdf_voucher(tipo_servicio, datos):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A5),
        rightMargin=1*cm, leftMargin=1*cm,
        topMargin=0.8*cm, bottomMargin=0.8*cm
    )
    styles = getSampleStyleSheet()
    story = []

    # Estilos
    logo_style = ParagraphStyle('LogoStyle', parent=styles['Normal'], fontSize=20, leading=22, fontName='Helvetica-Bold', textColor=colors.HexColor('#1E3A8A'))
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=9, leading=12, alignment=2) # Alineado a la derecha
    
    banner_title = ParagraphStyle('BannerTitle', parent=styles['Heading1'], fontSize=14, leading=16, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    banner_sub = ParagraphStyle('BannerSub', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=11)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')

    # --- 1. ENCABEZADO SUPERIOR (Logo izq / Fecha y Nº Voucher der) ---
    header_data = [
        [
            Paragraph("<b>Wolf</b> <font size=10>travel</font>", logo_style),
            Paragraph("<b>FECHA:</b> ____ / ____ / ________<br/><br/><b>Nº VOUCHER:</b> ________________", meta_style)
        ]
    ]
    t_header = Table(header_data, colWidths=[9*cm, 8*cm])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 0.1*cm))

    # --- 2. BANNER CENTRAL (Título y Subtítulo) ---
    # Normalizar subtítulo según tipo de servicio
    sub_texto = "HOTELERÍA" if tipo_servicio.lower() == "hotel" else ("SERVICIOS GASTRONÓMICOS" if tipo_servicio.lower() == "gastronomía" else "ACTIVIDADES")
    
    banner_content = [
        [Paragraph("VOUCHER DE SERVICIO", banner_title)],
        [Paragraph(sub_texto, banner_sub)]
    ]
    t_banner = Table(banner_content, colWidths=[17*cm])
    t_banner.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_banner)
    story.append(Spacer(1, 0.4*cm))

    # --- 3. TABLA DE DATOS DEL VOUCHER ---
    tabla_datos = []
    for k, v in datos.items():
        tabla_datos.append([Paragraph(f"<b>{k}:</b>", bold_style), Paragraph(str(v), body_style)])

    t = Table(tabla_datos, colWidths=[5.5*cm, 11.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 3),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.8*cm)) # Espacio antes de firmas

    # --- 4. BLOQUE DE FIRMAS Y ACLARACIÓN (Con renglón extra de espacio) ---
    firmas_data = [
        ["", ""], # Renglón extra de espacio solicitado
        ["........................................................................", "........................................................................"],
        ["Firma y Aclaración (Coordinador)", "Firma y Aclaración (Proveedor)"]
    ]
    t_firmas = Table(firmas_data, colWidths=[8.5*cm, 8.5*cm])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TEXTCOLOR', (0,2), (-1,2), colors.HexColor('#4B5563')),
        ('FONTNAME', (0,2), (-1,2), 'Helvetica'),
        ('FONTSIZE', (0,2), (-1,2), 8),
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
            
            coordinador_input = st.text_input("Coordinador a cargo:", value="Coordinador General")
            observaciones_input = st.text_area("Observaciones generales:")
            
            st.markdown("---")
            
            # Botón de Descarga (Único o Lote en ZIP)
            if emitir_todos:
                st.info(f"📦 Se generará un paquete ZIP con: 1 Voucher de Hotel, 1 de Gastronomía y {len(opciones_actividades)} Vouchers de Actividades para el Coche {coche_final}.")
                
                if st.button("🖨️ Generar y Descargar Todos los Vouchers (ZIP)"):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        
                        # 1. Voucher Hotel
                        datos_hotel = {
                            "Coche": coche_final,
                            "Hotel": "A definir / Según Operativa",
                            "Tripulación": 2,
                            "Tripulación Liberada": 0,
                            "Observaciones": observaciones_input
                        }
                        pdf_hotel = generar_pdf_voucher("Hotel", datos_hotel)
                        zip_file.writestr(f"Voucher_Hotel_Coche_{coche_final}.pdf", pdf_hotel.getvalue())
                        
                        # 2. Voucher Gastronomía
                        datos_gas = {
                            "Coche": coche_final,
                            "Coordinador": coordinador_input,
                            "Almuerzos": cant_pax_total,
                            "Cenas": cant_pax_total,
                            "Quinta Comida": 0,
                            "Observaciones": observaciones_input
                        }
                        pdf_gas = generar_pdf_voucher("Gastronomía", datos_gas)
                        zip_file.writestr(f"Voucher_Gastronomia_Coche_{coche_final}.pdf", pdf_gas.getvalue())
                        
                        # 3. Vouchers de Actividades
                        for act in opciones_actividades:
                            datos_act = {
                                "Coche": coche_final,
                                "Actividad / Servicio": act,
                                "Coordinador a cargo": coordinador_input,
                                "Pax": cant_pax_solo,
                                "Acompañantes": cant_acomp_solo,
                                "Coordinadores": 2,
                                "Choferes": 2,
                                "Staff Extra": 0,
                                "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4,
                                "Liberados": 0,
                                "Observaciones": observaciones_input
                            }
                            pdf_act = generar_pdf_voucher("Actividades", datos_act)
                            nombre_seguro = "".join([c if c.isalnum() else "_" for c in act])
                            zip_file.writestr(f"Voucher_Actividad_{nombre_seguro}_Coche_{coche_final}.pdf", pdf_act.getvalue())
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label="⬇️ Descargar Paquete ZIP con Todos los Vouchers",
                        data=zip_buffer,
                        file_name=f"Vouchers_Completos_Coche_{coche_final}.zip",
                        mime="application/zip"
                    )
            else:
                if tipo_servicio == "Hotel":
                    datos_voucher = {
                        "Coche": coche_final,
                        "Hotel": st.text_input("Hotel específico:"),
                        "Tripulación": 2,
                        "Tripulación Liberada": 0,
                        "Observaciones": observaciones_input
                    }
                elif tipo_servicio == "Gastronomía":
                    datos_voucher = {
                        "Coche": coche_final,
                        "Coordinador": coordinador_input,
                        "Almuerzos": cant_pax_total,
                        "Cenas": cant_pax_total,
                        "Quinta Comida": 0,
                        "Observaciones": observaciones_input
                    }
                else:
                    datos_voucher = {
                        "Coche": coche_final,
                        "Actividad / Servicio": servicio_especifico if servicio_especifico != "-- Seleccionar Actividad --" else "A DEFINIR",
                        "Coordinador a cargo": coordinador_input,
                        "Pax": cant_pax_solo,
                        "Acompañantes": cant_acomp_solo,
                        "Coordinadores": 2,
                        "Choferes": 2,
                        "Staff Extra": 0,
                        "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4,
                        "Liberados": 0,
                        "Observaciones": observaciones_input
                    }

                pdf_bytes = generar_pdf_voucher(tipo_servicio, datos_voucher)
                
                st.download_button(
                    label=f"🖨️ Generar y Descargar Voucher [{tipo_servicio}]",
                    data=pdf_bytes,
                    file_name=f"Voucher_{tipo_servicio}_Coche_{coche_

cat << 'EOF' > app.py
import os
import io
import zipfile
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

# --- FUNCIÓN GENERADORA DE PDF EN MEMORIA (NUEVO DISEÑO) ---
def generar_pdf_voucher(tipo_servicio, datos):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A5),
        rightMargin=1*cm, leftMargin=1*cm,
        topMargin=0.8*cm, bottomMargin=0.8*cm
    )
    styles = getSampleStyleSheet()
    story = []

    # Estilos
    logo_style = ParagraphStyle('LogoStyle', parent=styles['Normal'], fontSize=20, leading=22, fontName='Helvetica-Bold', textColor=colors.HexColor('#1E3A8A'))
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=9, leading=12, alignment=2) # Alineado a la derecha
    
    banner_title = ParagraphStyle('BannerTitle', parent=styles['Heading1'], fontSize=14, leading=16, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    banner_sub = ParagraphStyle('BannerSub', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=11)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')

    # --- 1. ENCABEZADO SUPERIOR (Logo izq / Fecha y Nº Voucher der) ---
    header_data = [
        [
            Paragraph("<b>Wolf</b> <font size=10>travel</font>", logo_style),
            Paragraph("<b>FECHA:</b> ____ / ____ / ________<br/><br/><b>Nº VOUCHER:</b> ________________", meta_style)
        ]
    ]
    t_header = Table(header_data, colWidths=[9*cm, 8*cm])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 0.1*cm))

    # --- 2. BANNER CENTRAL (Título y Subtítulo) ---
    # Normalizar subtítulo según tipo de servicio
    sub_texto = "HOTELERÍA" if tipo_servicio.lower() == "hotel" else ("SERVICIOS GASTRONÓMICOS" if tipo_servicio.lower() == "gastronomía" else "ACTIVIDADES")
    
    banner_content = [
        [Paragraph("VOUCHER DE SERVICIO", banner_title)],
        [Paragraph(sub_texto, banner_sub)]
    ]
    t_banner = Table(banner_content, colWidths=[17*cm])
    t_banner.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_banner)
    story.append(Spacer(1, 0.4*cm))

    # --- 3. TABLA DE DATOS DEL VOUCHER ---
    tabla_datos = []
    for k, v in datos.items():
        tabla_datos.append([Paragraph(f"<b>{k}:</b>", bold_style), Paragraph(str(v), body_style)])

    t = Table(tabla_datos, colWidths=[5.5*cm, 11.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 3),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.8*cm)) # Espacio antes de firmas

    # --- 4. BLOQUE DE FIRMAS Y ACLARACIÓN (Con renglón extra de espacio) ---
    firmas_data = [
        ["", ""], # Renglón extra de espacio solicitado
        ["........................................................................", "........................................................................"],
        ["Firma y Aclaración (Coordinador)", "Firma y Aclaración (Proveedor)"]
    ]
    t_firmas = Table(firmas_data, colWidths=[8.5*cm, 8.5*cm])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TEXTCOLOR', (0,2), (-1,2), colors.HexColor('#4B5563')),
        ('FONTNAME', (0,2), (-1,2), 'Helvetica'),
        ('FONTSIZE', (0,2), (-1,2), 8),
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
            
            coordinador_input = st.text_input("Coordinador a cargo:", value="Coordinador General")
            observaciones_input = st.text_area("Observaciones generales:")
            
            st.markdown("---")
            
            # Botón de Descarga (Único o Lote en ZIP)
            if emitir_todos:
                st.info(f"📦 Se generará un paquete ZIP con: 1 Voucher de Hotel, 1 de Gastronomía y {len(opciones_actividades)} Vouchers de Actividades para el Coche {coche_final}.")
                
                if st.button("🖨️ Generar y Descargar Todos los Vouchers (ZIP)"):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        
                        # 1. Voucher Hotel
                        datos_hotel = {
                            "Coche": coche_final,
                            "Hotel": "A definir / Según Operativa",
                            "Tripulación": 2,
                            "Tripulación Liberada": 0,
                            "Observaciones": observaciones_input
                        }
                        pdf_hotel = generar_pdf_voucher("Hotel", datos_hotel)
                        zip_file.writestr(f"Voucher_Hotel_Coche_{coche_final}.pdf", pdf_hotel.getvalue())
                        
                        # 2. Voucher Gastronomía
                        datos_gas = {
                            "Coche": coche_final,
                            "Coordinador": coordinador_input,
                            "Almuerzos": cant_pax_total,
                            "Cenas": cant_pax_total,
                            "Quinta Comida": 0,
                            "Observaciones": observaciones_input
                        }
                        pdf_gas = generar_pdf_voucher("Gastronomía", datos_gas)
                        zip_file.writestr(f"Voucher_Gastronomia_Coche_{coche_final}.pdf", pdf_gas.getvalue())
                        
                        # 3. Vouchers de Actividades
                        for act in opciones_actividades:
                            datos_act = {
                                "Coche": coche_final,
                                "Actividad / Servicio": act,
                                "Coordinador a cargo": coordinador_input,
                                "Pax": cant_pax_solo,
                                "Acompañantes": cant_acomp_solo,
                                "Coordinadores": 2,
                                "Choferes": 2,
                                "Staff Extra": 0,
                                "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4,
                                "Liberados": 0,
                                "Observaciones": observaciones_input
                            }
                            pdf_act = generar_pdf_voucher("Actividades", datos_act)
                            nombre_seguro = "".join([c if c.isalnum() else "_" for c in act])
                            zip_file.writestr(f"Voucher_Actividad_{nombre_seguro}_Coche_{coche_final}.pdf", pdf_act.getvalue())
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label="⬇️ Descargar Paquete ZIP con Todos los Vouchers",
                        data=zip_buffer,
                        file_name=f"Vouchers_Completos_Coche_{coche_final}.zip",
                        mime="application/zip"
                    )
            else:
                if tipo_servicio == "Hotel":
                    datos_voucher = {
                        "Coche": coche_final,
                        "Hotel": st.text_input("Hotel específico:"),
                        "Tripulación": 2,
                        "Tripulación Liberada": 0,
                        "Observaciones": observaciones_input
                    }
                elif tipo_servicio == "Gastronomía":
                    datos_voucher = {
                        "Coche": coche_final,
                        "Coordinador": coordinador_input,
                        "Almuerzos": cant_pax_total,
                        "Cenas": cant_pax_total,
                        "Quinta Comida": 0,
                        "Observaciones": observaciones_input
                    }
                else:
                    datos_voucher = {
                        "Coche": coche_final,
                        "Actividad / Servicio": servicio_especifico if servicio_especifico != "-- Seleccionar Actividad --" else "A DEFINIR",
                        "Coordinador a cargo": coordinador_input,
                        "Pax": cant_pax_solo,
                        "Acompañantes": cant_acomp_solo,
                        "Coordinadores": 2,
                        "Choferes": 2,
                        "Staff Extra": 0,
                        "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4,
                        "Liberados": 0,
                        "Observaciones": observaciones_input
                    }

                pdf_bytes = generar_pdf_voucher(tipo_servicio, datos_voucher)
                
                st.download_button(
                    label=f"🖨️ Generar y Descargar Voucher [{tipo_servicio}]",
                    data=pdf_bytes,
                    file_name=f"Voucher_{tipo_servicio}_Coche_{coche_final}.pdf",
                    mime="application/pdf"
                )

        else:
            st.warning(f"No se encontraron pasajeros ACTIVOS asignados exactamente al coche '{coche_final}'.")
else:
    st.error("No se encontró la columna 'Coche' en la planilla de Control de Pasajeros.")
