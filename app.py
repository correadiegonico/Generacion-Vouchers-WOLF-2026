import os
import sqlite3
import re
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

# --- CONFIGURACIÓN DE BASE DE DATOS PERSISTENTE (SQLite) ---
DB_PATH = "wolf_travel_itineraries.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS service_dates (
            contract_code TEXT,
            service_name TEXT,
            service_date TEXT,
            PRIMARY KEY (contract_code, service_name)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_service_date(contract_code, service_name, service_date):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO service_dates (contract_code, service_name, service_date)
        VALUES (?, ?, ?)
    ''', (contract_code, service_name, service_date))
    conn.commit()
    conn.close()

def get_saved_dates():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM service_dates", conn)
    conn.close()
    return df

# --- GENERADOR DE PDF (VOUCHER HORIZONTAL 50% A4) ---
def generate_pdf_voucher(data, output_folder):
    safe_service_name = re.sub(r'[\\/*?:"<>|]', '_', str(data['service'])).replace(' ', '_')
    pdf_filename = f"Voucher_Coche{data['coche']}_{safe_service_name}.pdf"
    full_path = os.path.join(output_folder, pdf_filename)
    
    # Orientación horizontal (Landscape) y dimensiones ajustadas para ocupar el 50% de la hoja A4 (apaisada)
    # A4 landscape es 841.89 x 595.27 puntos. 50% de altura son aprox 280 puntos.
    doc = SimpleDocTemplate(
        full_path, 
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
    
    # Orden estricto solicitado para los ítems
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
        [Paragraph("<b>TOTAL DE PERSONAS:</b>", bold_style), Paragraph(f"<b>{data['total_personas']}</b>", bold_style)],
    ]
    
    t = Table(table_data, colWidths=[200, 560])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8f9fa")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#dee2e6")),
        ('PADDING', (0,0), (-1,-1), 3.5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0, 9), (-1, 9), colors.HexColor("#e2eafc")),
    ]))
    
    story.append(t)
    
    # Espacio de unos 3 renglones antes de las firmas
    story.append(Spacer(1, 20))
    
    # Líneas para firmas tal como están
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
    return full_path

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.sidebar.image("https://img.icons8.com/color/96/wolf.png", width=70)
st.sidebar.title("Wolf Travel 2026")
st.sidebar.markdown("---")

menu = st.sidebar.radio("Navegación", ["📊 Dashboard General", "📅 Gestión de Itinerarios", "🎫 Generación de Vouchers"])

st.sidebar.markdown("### Archivos de Origen")
desglose_file = st.sidebar.file_uploader("1. Desglose Servicios 2026 (.xlsx)", type=["xlsx", "xls"])
pasajeros_file = st.sidebar.file_uploader("2. Control de Pasajeros (.xlsx)", type=["xlsx", "xls"])

if desglose_file is not None:
    try:
        xls = pd.ExcelFile(desglose_file)
        dfs = []
        for sheet in xls.sheet_names:
            df_t = pd.read_excel(desglose_file, sheet_name=sheet)
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
        if pasajeros_file is not None:
            try:
                xls_pax = pd.ExcelFile(pasajeros_file)
                sheet_to_use = 'Listado Completo' if 'Listado Completo' in xls_pax.sheet_names else xls_pax.sheet_names[0]
                df_pasajeros = pd.read_excel(pasajeros_file, sheet_name=sheet_to_use)
                
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

        # --- VISTA 1: DASHBOARD GENERAL ---
        if menu == "📊 Dashboard General":
            st.title("🐺 Wolf Travel 2026 - Dashboard Operativo")
            st.markdown("Resumen de contratos y servicios turísticos cargados.")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Contratos en Desglose", len(df_desglose['CODIGO'].dropna().unique()) if 'CODIGO' in df_desglose.columns else 0)
            c2.metric("Pasajeros Totales (PAX)", int(df_desglose['PAX'].sum()) if 'PAX' in df_desglose.columns and pd.api.types.is_numeric_dtype(df_desglose['PAX']) else 0)
            c3.metric("Control de Pasajeros", "Cargado Correctamente" if df_pasajeros is not None else "Pendiente de Carga")
            
            st.dataframe(df_desglose, use_container_width=True)

        # --- VISTA 2: GESTIÓN DE ITINERARIOS ---
        elif menu == "📅 Gestión de Itinerarios":
            st.title("📅 Asignación de Fechas por Contrato")
            if 'CODIGO' in df_desglose.columns:
                contratos = df_desglose['CODIGO'].dropna().unique().tolist()
                selected_contract = st.selectbox("Seleccionar Código de Contrato:", contratos)
                
                row_data = df_desglose[df_desglose['CODIGO'] == selected_contract].iloc[0]
                st.info(f"**Colegio:** {row_data.get('COLEGIO', 'N/A')} | **Pasajeros:** {row_data.get('PAX', 'N/A')}")
                
                meta_cols = ['CODIGO', 'COLEGIO', 'PAX', 'HOTEL', 'DESTINO_ORIGEN', 'FECHA', 'DIAS', 'NTS']
                service_cols = [c for c in df_desglose.columns if c not in meta_cols and not str(c).startswith('Unnamed')]
                
                saved_df = get_saved_dates()
                for serv in service_cols:
                    col_a, col_b, col_c = st.columns([3, 2, 2])
                    col_a.markdown(f"**{serv}**")
                    
                    current_saved = saved_df[(saved_df['contract_code'] == selected_contract) & (saved_df['service_name'] == serv)]
                    default_date_val = datetime.today()
                    if not current_saved.empty:
                        try: default_date_val = datetime.strptime(current_saved.iloc[0]['service_date'], "%Y-%m-%d")
                        except: pass
                            
                    assigned_date = col_b.date_input(f"Fecha {serv}", value=default_date_val, key=f"date_{selected_contract}_{serv}")
                    if col_c.button("Guardar Fecha", key=f"btn_{selected_contract}_{serv}"):
                        save_service_date(selected_contract, serv, assigned_date.strftime("%Y-%m-%d"))
                        col_c.success("¡Guardado!")
            else:
                st.error("No se encontró la columna de códigos de contrato en el desglose.")

        # --- VISTA 3: GENERACIÓN DE VOUCHERS POR COCHE ---
        elif menu == "🎫 Generación de Vouchers":
            st.title("🎫 Generación de Vouchers por Coche")
            st.markdown("Seleccione el **Número de Coche** para generar los vouchers por excursión en su carpeta correspondiente.")
            
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
                    todos_los_servicios = st.checkbox("Generar confirmación de TODOS los servicios (excursiones) de este coche", value=True)
                
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
                        
                    if st.button("🚀 Generar y Guardar Vouchers en Carpeta", type="primary"):
                        # Crear carpeta con nombre dinámico: vouchers coche [número]
                        output_folder = f"vouchers coche {selected_coche}"
                        os.makedirs(output_folder, exist_ok=True)
                        
                        saved_dates_df = get_saved_dates()
                        generated_files = []
                        
                        servicios_a_procesar = service_cols if todos_los_servicios else [selected_service]
                        colegios_str = ", ".join(df_contratos_coche['COLEGIO'].dropna().unique().astype(str))
                        
                        for serv in servicios_a_procesar:
                            match_d = saved_dates_df[saved_dates_df['service_name'] == serv]
                            date_val = datetime.today()
                            if not match_d.empty:
                                try: date_val = datetime.strptime(match_d.iloc[0]['service_date'], "%Y-%m-%d")
                                except: pass
                            date_str = date_val.strftime("%d/%m/%Y")
                            
                            try: trip_val = int(tripulacion)
                            except: trip_val = 2
                            try: coord_val = int(coordinadores)
                            except: coord_val = 2
                            try: lib_val = int(liberados)
                            except: lib_val = 0
                            
                            total_personas = total_pax_coche + cant_acompanantes + trip_val + coord_val + lib_val
                            
                            excursion_first_word = str(serv).split()[0].upper()
                            voucher_code = f"ch#{selected_coche}-{excursion_first_word}"
                            
                            v_data = {
                                'voucher_code': voucher_code,
                                'coche': selected_coche,
                                'colegios_str': colegios_str,
                                'service': serv,
                                'date': date_str,
                                'coordinador': nombre_coord,
                                'pax': total_pax_coche,
                                'acompanantes': cant_acompanantes,
                                'tripulacion': str(trip_val),
                                'coordinadores': str(coord_val),
                                'liberados': str(lib_val),
                                'total_personas': total_personas
                            }
                            pdf_path = generate_pdf_voucher(v_data, output_folder)
                            generated_files.append(pdf_path)
                                
                        st.success(f"¡Se generaron con éxito {len(generated_files)} voucher(s) para el Coche #{selected_coche}!")
                        st.info(f"📁 Todos los archivos PDF fueron guardados automáticamente dentro de la carpeta: **{output_folder}/**")
                else:
                    st.warning("No se encontraron contratos asociados a este número de coche.")
            else:
                st.warning("⚠️ Cargue el archivo **Control de Pasajeros.xlsx** en la barra lateral para habilitar esta sección.")

    except Exception as e:
        st.error(f"Error procesando los archivos Excel: {e}")
else:
    st.sidebar.warning("Por favor, cargue los archivos Excel en la barra lateral.")
    st.title("🐺 Wolf Travel 2026 - Generador de Vouchers")
    st.info("Cargue los archivos de origen en la barra lateral para comenzar.")
