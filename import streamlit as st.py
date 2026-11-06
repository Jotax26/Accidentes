import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import datetime
import os

# Dependencias para el PDF profesional con ReportLab
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
    conn = sqlite3.connect('bitacora_accidentes.db', check_same_thread=False)
    return conn

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
# 2. GENERADOR DE REPORTES PDF PROFESIONAL (ReportLab)
# ---------------------------------------------------------
def generar_pdf_profesional(data_dict, filename):
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    # Estilos personalizados
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

    # Encabezado
    story.append(Paragraph("SISTEMA DE GESTIÓN VIAL Y TELEMETRÍA", title_style))
    story.append(Paragraph(f"INFORME TÉCNICO DE INCIDENTE EN CARRETERA | ID: {data_dict['id']}", subtitle_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1E3A8A'), spaceAfter=15))

    # Tabla 1: Datos Generales
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

    # Tabla 2: Prueba Piloto y Entorno
    story.append(Paragraph("2. Condición Operativa y del Entorno", h2_style))
    t2_data = [
        [Paragraph("Prueba Piloto", cell_bold), Paragraph(str(data_dict['prueba_piloto']), cell_norm), Paragraph("Modo Conducción", cell_bold), Paragraph(str(data_dict['modo_conduccion']), cell_norm)],
        [Paragraph("Estado Clima", cell_bold), Paragraph(str(data_dict['clima']), cell_norm), Paragraph("Reg. Telemetría", cell_bold), Paragraph("Disponible (.CSV / Plotly)", cell_norm)]
    ]
    t2 = Table(t2_data, colWidths=[110, 160, 110, 160])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t2)
    story.append(Spacer(1, 12))

    # Detalles Narrativos
    story.append(Paragraph("3. Descripción Detallada y Análisis Preliminar", h2_style))
    story.append(Paragraph(str(data_dict['detalles']), cell_norm))
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#9CA3AF'), spaceAfter=10))

    doc.build(story)
    return filename

# ---------------------------------------------------------
# 3. INTERFAZ STREAMLIT CON PESTAÑAS
# ---------------------------------------------------------
st.title("🚨 Bitácora Inteligente de Accidentes Viales")
st.caption("Plataforma de Registro, Análisis de Telemetría, Mapeo GPS y Exportación de Documentos")

tab_reg, tab_cons, tab_dash = st.tabs(["📝 Registrar Incidente", "🔍 Consulta & Evidencia", "📊 Dashboard Global"])

# =========================================================
# PESTAÑA 1: REGISTRO DE INCIDENTE
# =========================================================
with tab_reg:
    st.subheader("Formulario de Captura de Incidentes")
    
    with st.form("form_incidente", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            id_incidente = st.text_input("ID Incidente *", value=f"ACC-2026-00{datetime.datetime.now().second}")
            fecha = st.date_input("Fecha *", datetime.date.today())
            hora = st.time_input("Hora *", datetime.datetime.now().time())
        with col2:
            carretera = st.text_input("Carretera / Tramo *", placeholder="Ej. Panamericana Km 45")
            latitud = st.number_input("Latitud GPS", value=14.33110, format="%.5f")
            longitud = st.number_input("Longitud GPS", value=-89.44470, format="%.5f")
        with col3:
            conductor = st.text_input("Operador / Conductor", placeholder="Nombre completo")
            gravedad = st.selectbox("Gravedad del Daño", ["Leve", "Moderado", "Grave / Pérdida Total"])
            modo_conduccion = st.selectbox("Modo de Conducción", ["Manual", "Asistido (ADAS)", "Autónomo Completo"])

        col4, col5 = st.columns(2)
        with col4:
            prueba_piloto = st.selectbox("Prueba Piloto Asociada", ["N/A (Operación Normal)", "Piloto Radar v3", "Piloto Frenado Autónomo", "Piloto Sensores Lidar"])
            clima = st.selectbox("Condición Climática", ["Despejado", "Lluvia Leve", "Lluvia Intensa", "Niebla / Poca Visibilidad"])
        with col5:
            video_file = st.file_uploader("Adjuntar Video Dashcam (.mp4)", type=["mp4", "mov"])
            telemetria_file = st.file_uploader("Adjuntar Archivo de Telemetría (.csv)", type=["csv"])

        detalles = st.text_area("Descripción de los Hechos y Daños Visibles", placeholder="Describa cómo ocurrió el incidente...")

        btn_guardar = st.form_submit_button("💾 Guardar Registro en Base de Datos")

        if btn_guardar:
            if not id_incidente or not carretera:
                st.error("Por favor complete los campos obligatorios (*).")
            else:
                # Guardar video localmente si existe
                video_path = ""
                if video_file:
                    os.makedirs("uploads", exist_ok=True)
                    video_path = os.path.join("uploads", f"{id_incidente}_{video_file.name}")
                    with open(video_path, "wb") as f:
                        f.write(video_file.getbuffer())

                # Procesar telemetría CSV a JSON si existe
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
                ''', (id_incidente, str(fecha), str(hora), carretera, latitud, longitud, prueba_piloto, conductor, gravedad, modo_conduccion, clima, detalles, video_path, telemetria_str))
                conn.commit()
                st.success(f"✅ Incidente {id_incidente} guardado exitosamente.")

# =========================================================
# PESTAÑA 2: CONSULTA, MAPA, TELEMETRÍA Y EXPORTACIÓN PDF
# =========================================================
with tab_cons:
    st.subheader("Bitácora de Siniestros y Análisis Avanzado")
    
    conn = get_connection()
    df_acc = pd.read_sql_query("SELECT * FROM accidentes", conn)

    if df_acc.empty:
        st.info("No hay registros almacenados. Ingresa un registro en la pestaña anterior.")
    else:
        incidente_sel = st.selectbox("Seleccionar ID de Incidente para Inspección:", df_acc['id'].tolist())
        row = df_acc[df_acc['id'] == incidente_sel].iloc[0]

        # Resumen superior
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Fecha", row['fecha'])
        c2.metric("Gravedad", row['gravedad'])
        c3.metric("Modo", row['modo_conduccion'])
        c4.metric("Prueba", row['prueba_piloto'])

        st.markdown("---")
        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.markdown("### 📍 Ubicación GPS")
            df_mapa = pd.DataFrame({'lat': [row['latitud']], 'lon': [row['longitud']]})
            st.map(df_mapa, zoom=12)

            st.markdown("### 📽️ Evidencia de Video")
            if row['video_path'] and os.path.exists(row['video_path']):
                st.video(row['video_path'])
            else:
                st.info("Sin archivo de video disponible para este registro.")

        with col_right:
            st.markdown("### 📈 Análisis de Telemetría (Plotly)")
            if row['telemetria_json']:
                df_t = pd.read_json(row['telemetria_json'])
                
                # Gráfico interactivo con subplots / multilínea
                fig = go.Figure()
                if 'tiempo_seg' in df_t.columns:
                    x_axis = df_t['tiempo_seg']
                else:
                    x_axis = df_t.index

                if 'velocidad_kmh' in df_t.columns:
                    fig.add_trace(go.Scatter(x=x_axis, y=df_t['velocidad_kmh'], name="Velocidad (km/h)", line=dict(color="blue", width=2)))
                if 'frenado_pct' in df_t.columns:
                    fig.add_trace(go.Scatter(x=x_axis, y=df_t['frenado_pct'], name="Presión Freno (%)", line=dict(color="red", width=2, dash='dash')))
                
                fig.update_layout(title="Sensores Vehiculares Previos al Impacto", xaxis_title="Tiempo (segundos)", yaxis_title="Magnitud", height=320)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Este registro no incluye archivo de telemetría CSV. Mostrando gráfica simulada:")
                # Simulación de respaldo
                df_sim = pd.DataFrame({
                    'tiempo': [0, 1, 2, 3, 4, 5],
                    'velocidad': [80, 78, 75, 40, 10, 0],
                    'freno': [0, 0, 20, 100, 100, 0]
                })
                fig_sim = px.line(df_sim, x='tiempo', y=['velocidad', 'freno'], title="Simulación de Frenado de Emergencia")
                st.plotly_chart(fig_sim, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📄 Generación de Reporte Oficial")
        if st.button(f"Generar y Descargar PDF para {row['id']}"):
            pdf_filename = f"Reporte_{row['id']}.pdf"
            generar_pdf_profesional(row.to_dict(), pdf_filename)
            
            with open(pdf_filename, "rb") as f:
                st.download_button(
                    label="📥 Guardar PDF en Equipo",
                    data=f,
                    file_name=pdf_filename,
                    mime="application/pdf"
                )

# =========================================================
# PESTAÑA 3: DASHBOARD STATS
# =========================================================
with tab_dash:
    st.subheader("Métricas y Análisis Global de Siniestralidad")
    
    conn = get_connection()
    df_all = pd.read_sql_query("SELECT * FROM accidentes", conn)

    if df_all.empty:
        st.info("No hay suficientes datos para renderizar estadísticas globales.")
    else:
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Incidentes Registrados", len(df_all))
        kpi2.metric("En Prueba Piloto", len(df_all[df_all['prueba_piloto'] != 'N/A (Operación Normal)']))
        kpi3.metric("Casos Graves", len(df_all[df_all['gravedad'] == 'Grave / Pérdida Total']))

        st.markdown("---")
        g1, g2 = st.columns(2)
        
        with g1:
            fig_pie = px.pie(df_all, names='modo_conduccion', title="Distribución por Modo de Conducción", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

        with g2:
            fig_bar = px.bar(df_all, x='clima', color='gravedad', title="Siniestralidad según Condiciones Climáticas", barmode="group")
            st.plotly_chart(fig_bar, use_container_width=True)
