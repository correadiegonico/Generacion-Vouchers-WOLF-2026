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

# --- CARPETAS Y RUTAS DE PERSISTENCIA ---
UPLOAD_DIR = "uploaded_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

SAVED_DESGLOSE_PATH = os.path.join(UPLOAD_DIR, "desglose_servicios.xlsx")
SAVED_PASAJEROS_PATH = os.path.join(UPLOAD_DIR, "control_pasajeros.xlsx")

# --- GENERADOR DE PDF EN MEMORIA (RETORNA BYTES) ---
def generate_pdf_voucher_bytes(data):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=landscape(letter), 
        rightMargin=30, 
        leftMargin=30, 
        topMargin=20, 
        bottomMargin=20
    )
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor("#0B2545"), alignment=1, spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor("#134074"), alignment=1, spaceAfter=8
    )
    bold_style = ParagraphStyle(
        'BoldStyle', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#1d3557"), leading=11
    )
    normal_style = ParagraphStyle(
        'NormalStyle', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#333333"), leading=11
    )
    
    story.append(Paragraph("<b>WOLF TRAVEL 2026</b>", title_style))
    story.append(Paragraph("VOUCHER DE SERVICIOS Y EXCURSIONES", subtitle_style))
    
    table_data = [
        [Paragraph("<b>Fecha:</b>", bold_style), Paragraph(str(data['date']), normal_style)],
        [Paragraph("<b>N° de coche:</b>", bold_style), Paragraph(str(data['coche']), normal_style)],
        [Paragraph("<b>Servicio:</b>", bold_style), Paragraph(str(data['service']), normal_style)],
        [Paragraph("<b>Colegios / Grupos:</b>", bold_style), Paragraph(str(data['colegios_str']), normal_style)],
        [Paragraph("<b>Cantidad de pasajeros (PAX):</b>", bold_style), Paragraph(str(data['pax']), normal_style)],
        [Paragraph("<b>Cantidad de acompañantes:</b>", bold_style), Paragraph(str(data['acompanantes']), normal_style)],
        [Paragraph("<b>Tripulación / choferes:</b>", bold_style), Paragraph(str(data['tripulacion']), normal_style)],
        [Paragraph("<b>Coordinadores:</b>", bold_style), Paragraph(str(data['coordinadores']), normal_style)],
        [Paragraph("<b>Liberados:</b>", bold_style), Paragraph(str(data['liberados']), normal_style)],
        [Paragraph("<b>Staff empresa:</b>", bold_style), Paragraph(str(data['staff_empresa']), normal_style)],
        [Paragraph("<b>TOTAL =:</b>", bold_style), Paragraph(f"<b>{data['total_personas']}</b>", bold_style)],
    ]
    
    t = Table(table_data, colWidths=[200, 560])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8f9fa")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#dee2e6")),
        ('PADDING', (0,0), (-1,-1), 3.0),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0, 10), (-1, 10), colors.HexColor("#e2eafc")),
    ]))
    
    story.append(t)
    story.append(Spacer(1, 10))
    
    # Campo de Observaciones (dos renglones)
    obs_data = [
        [Paragraph(f"<b>Observaciones:</b><br/>{data['observaciones']}", normal_style)]
    ]
    t_obs = Table(obs_data, colWidths=[760])
    t_obs.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fff3cd")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#ffeeba")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_obs)
    story.append(Spacer(1, 15))
    
    sig_data = [
        [Paragraph("___________________________________<br/><b>Firma Coordinador</b>", bold_style),
         Paragraph("___________________________________<br/><b>Firma Proveedor / Prestador</b>", bold_style)]
    ]
    t_sig = Table(sig_data, colWidths=[380, 380])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_sig)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.sidebar.image("https://img.icons8.com/color/96/wolf.png", width=70)
st.sidebar.title("Wolf Travel 2026")
st.sidebar.markdown("---")

st.sidebar.markdown("### Archivos de Origen")
desglose_file = st.sidebar.file_uploader("1. Desglose Servicios 2026 (.xlsx)", type=["xlsx", "xls"])
pasajeros_file = st.sidebar.file_uploader("2. Control de Pasajeros (.xlsx)", type=["xlsx", "xls"])

if desglose_file is not None:
    with open(SAVED_DESGLOSE_PATH, "wb") as f:
        f.write(desglose_file.getbuffer())

if pasajeros_file is not None:
    with open(SAVED_PASAJEROS_PATH, "wb") as f:
        f.write(pasajeros_file.getbuffer())

has_desglose = os.path.exists(SAVED_DESGLOSE_PATH)
has_pasajeros = os.path.exists(SAVED_PASAJEROS_PATH)

if has_desglose:
    try:
        xls = pd.ExcelFile(SAVED_DESGLOSE_PATH)
        dfs = []
        for sheet in xls.sheet_names:
            df_t = pd.read_excel(SAVED_DESGLOSE_PATH, sheet_name=sheet)
            df_t['DESTINO_ORIGEN'] = sheet
            df_t.columns = [f"{col}_{i}" if df_t.columns.duplicated()[i] else col for i, col in enumerate(df_t.columns)]
            dfs.append(df_t)
        
        df_desglose = pd.concat(dfs, ignore_index=True, join='outer')
        
        d_map = {}
        for col in df_desglose.columns:
            cu = str(col).strip().upper()
            if 'CODIGO' in cu and 'CODIGO' not in d_map.values(): d_map[col] = 'CODIGO'
            elif 'COLEGIO' in cu and 'COLEGIO' not in d_map.values(): d_map[col] = 'COLEGIO'
            elif 'PAX' in cu and 'PAX' not in d_map.values(): d_map[col] = 'PAX'
            elif 'HOTEL' in cu and 'HOTEL' not in d_map.values(): d_map[col] = 'HOTEL'
        df_desglose = df_desglose.rename(columns=d_map)

        df_pasajeros = None
        if has_pasajeros:
            try:
                xls_pax = pd.ExcelFile(SAVED_PASAJEROS_PATH)
                sheet_to_use = 'Listado Completo' if 'Listado Completo' in xls_pax.sheet_names else xls_pax.sheet_names[0]
                df_pasajeros = pd.read_excel(SAVED_PASAJEROS_PATH, sheet_name=sheet_to_use)
                
                p_map = {}
                for col in df_pasajeros.columns:
                    cu = str(col).strip().upper()
                    if cu == 'CODIGO': p_map[col] = 'CODIGO'
                    elif cu == 'COLEGIO': p_map[col] = 'COLEGIO_PAX'
                    elif cu == 'COCHE': p_map[col] = 'NRO_COCHE'
                    elif cu == 'CATEGORIA': p_map[col] = 'CATEGORIA'
                df_pasajeros = df_pasajeros.rename(columns=p_map)
            except Exception as e:
                st.sidebar.warning(f"Aviso al leer Control de Pasajeros: {e}")

        # --- VISTA PRINCIPAL: GENERACIÓN DE VOUCHERS ---
        st.title("🎫 Wolf Travel 2026 - Generador de Vouchers")
        st.markdown("Seleccione el **Número de Coche** para generar y descargar los vouchers operativos correspondientes.")
        
        if df_pasajeros is not None and 'NRO_COCHE' in df_pasajeros.columns and 'CODIGO' in df_pasajeros.columns:
            coches_disponibles = sorted([str(int(c)) for c in df_pasajeros['NRO_COCHE'].dropna().unique() if pd.notna(c)])
            
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                selected_coche = st.selectbox("Número de Coche:", coches_disponibles)
            
            df_pasajeros['NRO_COCHE_STR'] = df_pasajeros['NRO_COCHE'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
            df_coche_pax = df_pasajeros[df_pasajeros['NRO_COCHE_STR'] == str(selected_coche)]
            contratos_en_coche = df_coche_pax['CODIGO'].dropna().unique().tolist()
            
            df_contratos_coche = df_desglose[df_desglose['CODIGO'].isin(contratos_en_coche)]
            
            total_pax_coche = int(df_contratos_coche['PAX'].sum()) if 'PAX' in df_contratos_coche.columns and pd.api.types.is_numeric_dtype(df_contratos_coche['PAX']) else 0
            
            cant_acompanantes_auto = 0
            if 'CATEGORIA' in df_coche_pax.columns:
                cant_acompanantes_auto = int((df_coche_pax['CATEGORIA'].astype(str).str.lower() == 'acompañante').sum())

            meta_cols = ['CODIGO', 'COLEGIO', 'PAX', 'HOTEL', 'DESTINO_ORIGEN', 'FECHA', 'DIAS', 'NTS']
            service_cols = [c for c in df_desglose.columns if c not in meta_cols and not str(c).startswith('Unnamed')]

            with col_s2:
                todos_los_servicios = st.checkbox("Generar confirmación de TODOS los servicios de este coche", value=True)
            
            selected_service = None
            if not todos_los_servicios:
                selected_service = st.selectbox("Seleccionar Servicio Específico:", service_cols)

            st.info(f"🚍 **Coche #{selected_coche}** | Contratos asociados: {len(contratos_en_coche)} | **Total PAX: {total_pax_coche}**")
            
            if contratos_en_coche and not df_contratos_coche.empty:
                st.dataframe(df_contratos_coche[['CODIGO', 'COLEGIO', 'PAX']], use_container_width=True)
                
                st.markdown("---")
                st.subheader("Datos Operativos del Voucher")
                
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    nombre_coord = st.text_input("Nombre Coordinador a Cargo:", value="Coordinador Operativo")
                    cant_acompanantes = st.number_input("Cantidad de Acompañantes:", min_value=0, value=cant_acompanantes_auto)
                    liberados = st.number_input("Liberados:", min_value=0, value=0)
                with col_p2:
                    tripulacion = st.text_input("Tripulación / Choferes:", value="2")
                    coordinadores = st.number_input("Coordinadores:", min_value=1, value=2)
                    staff_empresa = st.text_input("Staff empresa:", value="0")
                
                observaciones_input = st.text_area("Observaciones (2 renglones appx):", value="Sin observaciones particulares.", height=70)
                
                # Cálculo en tiempo real del TOTAL
                try: t_val = int(tripulacion)
                except: t_val = 2
                try: c_val = int(coordinadores)
                except: c_val = 2
                try: l_val = int(liberados)
                except: l_val = 0
                try: s_val = int(staff_empresa)
                except: s_val = 0
                
                total_personas_calculado = total_pax_coche + cant_acompanantes + t_val + c_val + l_val + s_val
                st.metric("TOTAL DE PERSONAS (Calculado)", total_personas_calculado)

                if st.button("🚀 Generar Vouchers para Descarga", type="primary"):
                    servicios_a_procesar = service_cols if todos_los_servicios else [selected_service]
                    colegios_str = ", ".join(df_contratos_coche['COLEGIO'].dropna().unique().astype(str))
                    
                    st.success("¡Vouchers listos para descargar individualmente abajo!")
                    
                    for serv in servicios_a_procesar:
                        excursion_first_word = str(serv).split()[0].upper()
                        voucher_code = f"ch#{selected_coche}-{excursion_first_word}"
                        
                        v_data = {
                            'voucher_code': voucher_code,
                            'coche': selected_coche,
                            'colegios_str': colegios_str,
                            'service': serv,
                            'date': datetime.today().strftime("%d/%m/%Y"),
                            'coordinador': nombre_coord,
                            'pax': total_pax_coche,
                            'acompanantes': cant_acompanantes,
                            'tripulacion': str(t_val),
                            'coordinadores': str(c_val),
                            'liberados': str(l_val),
                            'staff_empresa': str(s_val),
                            'total_personas': total_personas_calculado,
                            'observaciones': observaciones_input
                        }
                        
                        pdf_bytes = generate_pdf_voucher_bytes(v_data)
                        safe_service_name = re.sub(r'[\\/*?:"<>|]', '_', str(serv)).replace(' ', '_')
                        pdf_filename = f"Voucher_Coche{selected_coche}_{safe_service_name}.pdf"
                        
                        # Botón de descarga directa infalible para cada PDF
                        st.download_button(
                            label=f"📥 Descargar PDF: {serv}",
                            data=pdf_bytes,
                            file_name=pdf_filename,
                            mime="application/pdf",
                            key=f"dl_{selected_coche}_{serv}"
                        )
            else:
                st.warning("No se encontraron contratos asociados a este número de coche.")
        else:
            st.warning("⚠️ Cargue el archivo **Control de Pasajeros.xlsx** en la barra lateral para habilitar el sistema.")

    except Exception as e:
        st.error(f"Error procesando los archivos Excel: {e}")
else:
    st.title("🐺 Wolf Travel 2026 - Generador de Vouchers")
    st.info("📂 Utiliza los botones de la barra lateral izquierda para cargar los archivos **Desglose Servicios** y **Control de Pasajeros** por primera vez.")
