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
    page_title="Wolf Travel 2026 - Vouchers",
    page_icon="🐺",
    layout="centered", # Usamos centered para que en celulares se vea ordenado verticalmente sin desbordes
    initial_sidebar_state="collapsed" # Contraído por defecto para ganar velocidad en móviles
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
        topMargin=0.8*cm, bottomMargin=0.8*cm
    )
    styles = getSampleStyleSheet()
    story = []

    logo_style = ParagraphStyle('LogoStyle', parent=styles['Normal'], fontSize=20, leading=22, fontName='Helvetica-Bold', textColor=colors.HexColor('#1E3A8A'))
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=9, leading=12, alignment=2)
    
    banner_title = ParagraphStyle('BannerTitle', parent=styles['Heading1'], fontSize=14, leading=16, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    banner_sub = ParagraphStyle('BannerSub', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, textColor=colors.white, fontName='Helvetica-Bold')
    
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=11)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')

    # Encabezado Superior
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
    story.append(Spacer(1, 0.8*cm))

    firmas_data = [
        ["", ""],
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
    st.title("🐺 Wolf Travel")
    if st.button("🔄 Sincronizar datos"):
        st.cache_data.clear()
        st.rerun()

# --- HEADER PRINCIPAL ---
st.title("🐺 Wolf Travel 2026")
st.caption("Generador Móvil de Vouchers")

# Cargar datos
with st.spinner("Cargando base de datos..."):
    df_pasajeros, df_desglose, error = cargar_datos_drive()

if error:
    st.error(f"Error de conexión: {error}")
    st.stop()

# --- BUSCADOR POR NÚMERO DE COCHE ---
st.markdown("---")
st.subheader("🚌 Buscar Coche")

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
    
    coche_seleccionado = st.selectbox("Seleccione el Número de Coche:", ["-- Seleccionar --"] + coches_disponibles)
        
    if coche_seleccionado != "-- Seleccionar --":
        coche_final = coche_seleccionado
        df_coche = df_activos[df_activos['Coche_Clean'] == coche_final]
        
        if not df_coche.empty:
            col_categoria = df_coche.columns[15] if len(df_coche.columns) > 15 else 'Categoria'
            
            mask_acomp = df_coche[col_categoria].astype(str).str.upper().str.contains('ACOMPAÑANTE|ACOMPANANTE|PADRE|MADRE', na=False)
            cant_acomp_solo = len(df_coche[mask_acomp])
            cant_pax_solo = len(df_coche[~mask_acomp])
            cant_pax_total = len(df_coche)
            
            if 'Colegio' in df_coche.columns and 'Codigo' in df_coche.columns:
                cant_contratos = df_coche['Codigo'].nunique()
            else:
                cant_contratos = 1

            # Métricas adaptadas a móviles
            col1, col2, col3 = st.columns(3)
            col1.metric("Contratos", cant_contratos)
            col2.metric("Pax", cant_pax_solo)
            col3.metric("Acomp.", cant_acomp_solo)
            
            st.markdown("---")
            st.subheader("⚙️ Configuración del Voucher")
            
            tipo_servicio = st.selectbox("Tipo de servicio:", ["Actividades", "Hotel", "Gastronomía"])
            emitir_todos = st.checkbox("Emitir todos los vouchers del coche")

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
                    servicio_especifico = st.selectbox("Servicio a emitir:", ["-- Seleccionar Actividad --"] + opciones_actividades)
            
            st.markdown("---")
            coordinador_input = st.text_input("Coordinador a cargo:", value="Coordinador General")
            observaciones_input = st.text_area("Observaciones:")
            
            st.markdown("---")
            
            if emitir_todos:
                st.info(f"📦 Paquete ZIP para Coche {coche_final} (Hotel, Gastronomía y Actividades).")
                
                if st.button("⬇️ Generar y Descargar Paquete ZIP", type="primary", use_container_width=True):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        
                        datos_hotel = {"Coche": coche_final, "Hotel": "A definir", "Tripulación": 2, "Tripulación Liberada": 0, "Observaciones": observaciones_input}
                        pdf_hotel = generar_pdf_voucher("Hotel", datos_hotel)
                        zip_file.writestr(f"Voucher_Hotel_Coche_{coche_final}.pdf", pdf_hotel.getvalue())
                        
                        datos_gas = {"Coche": coche_final, "Coordinador": coordinador_input, "Almuerzos": cant_pax_total, "Cenas": cant_pax_total, "Quinta Comida": 0, "Observaciones": observaciones_input}
                        pdf_gas = generar_pdf_voucher("Gastronomía", datos_gas)
                        zip_file.writestr(f"Voucher_Gastronomia_Coche_{coche_final}.pdf", pdf_gas.getvalue())
                        
                        for act in opciones_actividades:
                            datos_act = {
                                "Coche": coche_final, "Actividad / Servicio": act, "Coordinador a cargo": coordinador_input,
                                "Pax": cant_pax_solo, "Acompañantes": cant_acomp_solo, "Coordinadores": 2, "Choferes": 2,
                                "Staff Extra": 0, "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4, "Liberados": 0, "Observaciones": observaciones_input
                            }
                            pdf_act = generar_pdf_voucher("Actividades", datos_act)
                            nombre_seguro = "".join([c if c.isalnum() else "_" for c in act])
                            zip_file.writestr(f"Voucher_Actividad_{nombre_seguro}_Coche_{coche_final}.pdf", pdf_act.getvalue())
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 Toca aquí para descargar el archivo ZIP",
                        data=zip_buffer,
                        file_name=f"Vouchers_Coche_{coche_final}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            else:
                if tipo_servicio == "Hotel":
                    datos_voucher = {"Coche": coche_final, "Hotel": st.text_input("Hotel específico:"), "Tripulación": 2, "Tripulación Liberada": 0, "Observaciones": observaciones_input}
                elif tipo_servicio == "Gastronomía":
                    datos_voucher = {"Coche": coche_final, "Coordinador": coordinador_input, "Almuerzos": cant_pax_total, "Cenas": cant_pax_total, "Quinta Comida": 0, "Observaciones": observaciones_input}
                else:
                    datos_voucher = {
                        "Coche": coche_final, "Actividad / Servicio": servicio_especifico if servicio_especifico != "-- Seleccionar Actividad --" else "A DEFINIR",
                        "Coordinador a cargo": coordinador_input, "Pax": cant_pax_solo, "Acompañantes": cant_acomp_solo, "Coordinadores": 2,
                        "Choferes": 2, "Staff Extra": 0, "TOTAL GENERAL": cant_pax_solo + cant_acomp_solo + 4, "Liberados": 0, "Observaciones": observaciones_input
                    }

                pdf_bytes = generar_pdf_voucher(tipo_servicio, datos_voucher)
                
                st.download_button(
                    label=f"📥 Descargar Voucher PDF [{tipo_servicio}]",
                    data=pdf_bytes,
                    file_name=f"Voucher_{tipo_servicio}_Coche_{coche_final}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.warning("No hay pasajeros activos en este coche.")
