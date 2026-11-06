import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import datetime
import os

# Librerías para Mapa Interactivo con Clic Directo
import folium
from streamlit_folium import st_folium

# Librerías para PDF Profesional
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y BASE DE DATOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sistema de Bitácora y Control de Accidentes",
    page_icon="🚨",
    layout="wide"
)

def get_connection():
    return sqlite3.connect('bitacora_accidentes.db', check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS accidentes (
            id TEXT PRIMARY KEY,
            fecha TEXT,
            hora TEXT,
            carretera TEXT,
            latitud REAL,
            longitud REAL,
            prueba_piloto TEXT,
            conductor TEXT,
            gravedad TEXT,
            modo_conduccion TEXT,
            clima TEXT,
            detalles TEXT,
            video_path TEXT,
            telemetria_json TEXT
        )
    ''')
    conn.commit()

init_db()

# ---------------------------------------------------------
# 2. GENERADOR DE REPORTES PDF (ReportLab)
# ---------------------------------------------------------
def generar_pdf_profesional(data_dict, filename):
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=18, leading=22, textColor=colors.HexColor('#1E3A8A')
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#4B5563')
    )
    h2_style = ParagraphStyle(
        'SectionHeading', parent=styles['Heading2'],
        fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#1E3A8A'), spaceBefore=10, spaceAfter=5
    )
    cell_bold = ParagraphStyle('CellB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#1F2937'))
    cell_norm = ParagraphStyle('CellN', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#374151'))

    story.append(Paragraph("SISTEMA DE GESTIÓN VIAL Y TELEMETRÍA", title_style))
    story.append(Paragraph(f"INFORME TÉCNICO DE INCIDENTE EN CARRETERA | ID: {data_dict['id']}", subtitle_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1E3A8A'), spaceAfter=15))

    story.append(Paragraph("1. Información General del Incidente", h2_style))
    t1_data = [
        [Paragraph("ID Incidente", cell_bold), Paragraph(str(data_dict['id']), cell_norm), Paragraph("Fecha / Hora", cell_bold), Paragraph(f"{data_dict['fecha']} - {data_dict['hora']}", cell_norm)],
        [Paragraph("Carretera / Tramo", cell_bold), Paragraph(str(data_dict['carretera']), cell_norm), Paragraph("Coordenadas GPS", cell_bold), Paragraph(f"{data_dict['latitud']}, {data_dict['longitud']}", cell_norm)],
        [Paragraph("Conductor / Op.", cell_bold), Paragraph(str(data_dict['conductor']), cell_norm), Paragraph("Nivel Gravedad", cell_bold), Paragraph(str(data_dict['gravedad']), cell_norm)],
    ]
    t1 = Table(t1_data, colWidths=[110, 160, 110, 160])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t1)
    story.append(Spacer(1, 12))

    story.append(Paragraph("2. Condición Operativa y del Entorno", h2_style))
    t2_data = [
        [Paragraph("Prueba Piloto", cell_bold), Paragraph(str(data_dict['prueba_piloto']), cell_norm), Paragraph("Modo Conducción", cell_bold), Paragraph(str(data_dict['modo_conduccion']), cell_norm)],
        [Paragraph("Estado Clima", cell_bold), Paragraph(str(data_dict['clima']), cell_norm), Paragraph("Reg. Telemetría", cell_bold), Paragraph("Disponible", cell_norm)]
    ]
    t2 = Table(t2_data, colWidths=[110, 160, 110, 160])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t2)
    story.append(Spacer(1, 12))

    story.append(Paragraph("3. Descripción Detallada y Análisis Preliminar", h2_style))
    story.append(Paragraph(str(data_dict['detalles']), cell_norm))
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#9CA3AF'), spaceAfter=10))

    doc.build(story)
    return filename

# ---------------------------------------------------------
# 3. INTERFAZ PRINCIPAL
# ---------------------------------------------------------
st.title("🚨 Bitácora Inteligente de Accidentes Viales")
st.caption("Plataforma de Registro con Selección Geográfica en Mapa, Telemetría y PDF")

tab_reg, tab_cons, tab_dash = st.tabs(["📝 Registrar Incidente", "🔍 Consulta & Evidencia", "📊 Dashboard Global"])

# =========================================================
# PESTAÑA 1: REGISTRO CON SELECCIÓN DE MAPA
# =========================================================
with tab_reg:
    st.subheader("Captura de Incidente y Selección Geográfica")
    
    col_form, col_map = st.columns([1, 1.2])

    with col_map:
        st.markdown("##### 📍 Haz clic en el mapa para marcar el punto exacto:")
        
        # Coordenadas por defecto (Centro de mapa inicial)
        lat_inicial, lon_inicial = 14.3311, -89.4447
        
        # Crear mapa folium
        m = folium.Map(location=[lat_inicial, lon_inicial], zoom_start=11)

        # Capturar el evento de clic en el mapa
        map_data = st_folium(m, height=420, width="100%")

        # Extraer latitud y longitud tras el clic
        lat_seleccionada = None
        lon_seleccionada = None

        if map_data and map_data.get("last_clicked"):
            lat_seleccionada = round(map_data["last_clicked"]["lat"], 6)
            lon_seleccionada = round(map_data["last_clicked"]["lng"], 6)
            st.success(f"📍 Punto Seleccionado: **Lat:** {lat_seleccionada} | **Lon:** {lon_seleccionada}")
        else:
            st.info("👆 Haz clic sobre cualquier carretera en el mapa para capturar las coordenadas.")

    with col_form:
        with st.form("form_incidente", clear_on_submit=True):
            id_incidente = st.text_input("ID Incidente *", value=f"ACC-2026-00{datetime.datetime.now().second}")
            c1, c2 = st.columns(2)
            fecha = c1.date_input("Fecha *", datetime.date.today())
            hora = c2.time_input("Hora *", datetime.datetime.now().time())
            
            carretera = st.text_input("Carretera / Nombre del Tramo *", placeholder="Ej. Carretera Panamericana, Km 42")
            conductor = st.text_input("Conductor / Operador", placeholder="Nombre completo")
            
            c3, c4 = st.columns(2)
            gravedad = c3.selectbox("Gravedad", ["Leve", "Moderado", "Grave / Pérdida Total"])
            modo_conduccion = c4.selectbox("Modo", ["Manual", "Asistido (ADAS)", "Autónomo Completo"])
            
            prueba_piloto = st.selectbox("Prueba Piloto", ["N/A (Operación Normal)", "Piloto Radar v3", "Piloto Frenado Autónomo", "Piloto Sensores Lidar"])
            clima = st.selectbox("Clima", ["Despejado", "Lluvia Leve", "Lluvia Intensa", "Niebla"])
            
            video_file = st.file_uploader("Video Dashcam (.mp4)", type=["mp4", "mov"])
            telemetria_file = st.file_uploader("Telemetría (.csv)", type=["csv"])
            detalles = st.text_area("Descripción de los Hechos", placeholder="Detalles del impacto...")

            btn_guardar = st.form_submit_button("💾 Guardar Incidente")

            if btn_guardar:
                if not id_incidente or not carretera:
                    st.error("Por favor completa los campos obligatorios (*).")
                elif lat_seleccionada is None or lon_seleccionada is None:
                    st.error("⚠️ Debes seleccionar una ubicación en el mapa haciendo clic sobre él antes de guardar.")
                else:
                    video_path = ""
                    if video_file:
                        os.makedirs("uploads", exist_ok=True)
                        video_path = os.path.join("uploads", f"{id_incidente}_{video_file.name}")
                        with open(video_path, "wb") as f:
                            f.write(video_file.getbuffer())

                    telemetria_str = ""
                    if telemetria_file:
                        df_telem = pd.read_csv(telemetria_file)
                        telemetria_str = df_telem.to_json(orient="records")

                    conn = get_connection()
                    c = conn.cursor()
                    c.execute('''
                        INSERT INTO accidentes 
                        (id, fecha, hora, carretera, latitud, longitud, prueba_piloto, conductor, gravedad, modo_conduccion, clima, detalles, video_path, telemetria_json)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (id_incidente, str(fecha), str(hora), carretera, lat_seleccionada, lon_seleccionada, prueba_piloto, conductor, gravedad, modo_conduccion, clima, detalles, video_path, telemetria_str))
                    conn.commit()
                    st.success(f"✅ Registro {id_incidente} guardado correctamente con coordenadas ({lat_seleccionada}, {lon_seleccionada}).")

# =========================================================
# PESTAÑA 2: CONSULTA Y REPRODUCCIÓN EN MAPA
# =========================================================
with tab_cons:
    st.subheader("Bitácora de Siniestros")
    conn = get_connection()
    df_acc = pd.read_sql_query("SELECT * FROM accidentes", conn)

    if df_acc.empty:
        st.info("No hay registros almacenados todavía.")
    else:
        incidente_sel = st.selectbox("Seleccionar ID para Inspección:", df_acc['id'].tolist())
        row = df_acc[df_acc['id'] == incidente_sel].iloc[0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Fecha", row['fecha'])
        c2.metric("Gravedad", row['gravedad'])
        c3.metric("Modo", row['modo_conduccion'])
        c4.metric("Prueba", row['prueba_piloto'])

        st.markdown("---")
        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.markdown("### 📍 Ubicación Guardada")
            m_resumen = folium.Map(location=[row['latitud'], row['longitud']], zoom_start=14)
            folium.Marker(
                [row['latitud'], row['longitud']],
                popup=f"Incidente: {row['id']}",
                tooltip=row['carretera'],
                icon=folium.Icon(color="red", icon="info-sign")
            ).add_to(m_resumen)
            st_folium(m_resumen, height=320, width="100%", key=f"map_{row['id']}")

            if row['video_path'] and os.path.exists(row['video_path']):
                st.markdown("### 📽️ Evidencia de Video")
                st.video(row['video_path'])

        with col_right:
            st.markdown("### 📈 Análisis de Telemetría (Plotly)")
            if row['telemetria_json']:
                df_t = pd.read_json(row['telemetria_json'])
                fig = go.Figure()
                x_axis = df_t['tiempo_seg'] if 'tiempo_seg' in df_t.columns else df_t.index

                if 'velocidad_kmh' in df_t.columns:
                    fig.add_trace(go.Scatter(x=x_axis, y=df_t['velocidad_kmh'], name="Velocidad (km/h)", line=dict(color="blue", width=2)))
                if 'frenado_pct' in df_t.columns:
                    fig.add_trace(go.Scatter(x=x_axis, y=df_t['frenado_pct'], name="Presión Freno (%)", line=dict(color="red", width=2, dash='dash')))
                
                fig.update_layout(title="Datos Vehiculares", xaxis_title="Tiempo (s)", yaxis_title="Valor", height=320)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sin archivo de telemetría adjunto.")

        st.markdown("---")
        if st.button(f"Generar PDF de {row['id']}"):
            pdf_filename = f"Reporte_{row['id']}.pdf"
            generar_pdf_profesional(row.to_dict(), pdf_filename)
            with open(pdf_filename, "rb") as f:
                st.download_button("📥 Descargar PDF", data=f, file_name=pdf_filename, mime="application/pdf")

# =========================================================
# PESTAÑA 3: DASHBOARD STATS
# =========================================================
with tab_dash:
    st.subheader("Análisis Global de Siniestralidad")
    conn = get_connection()
    df_all = pd.read_sql_query("SELECT * FROM accidentes", conn)

    if not df_all.empty:
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Incidentes", len(df_all))
        kpi2.metric("En Prueba Piloto", len(df_all[df_all['prueba_piloto'] != 'N/A (Operación Normal)']))
        kpi3.metric("Casos Graves", len(df_all[df_all['gravedad'] == 'Grave / Pérdida Total']))

        st.markdown("---")
        g1, g2 = st.columns(2)
        with g1:
            fig_pie = px.pie(df_all, names='modo_conduccion', title="Distribución por Modo de Conducción", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
        with g2:
            fig_bar = px.bar(df_all, x='clima', color='gravedad', title="Siniestralidad según Clima", barmode="group")
            st.plotly_chart(fig_bar, use_container_width=True)