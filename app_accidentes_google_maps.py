#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Plataforma de Registro y Reporte de Accidentes con Mapa Interactivo de Google Maps
- Usa teselas oficiales de Google Maps (Calles, Satélite e Híbrido) sin requerir API Key.
- Permite seleccionar el punto exacto del accidente haciendo clic sobre el mapa.
- Guarda automáticamente la ubicación (Latitud, Longitud y Enlace directo a Google Maps).
- Incluye generación de reportes en PDF (ReportLab), Excel y gráficos interactivos (Plotly).
"""

import os
import io
import re
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
import folium
from streamlit_folium import st_folium
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Configuración general de la página
st.set_page_config(
    page_title="Gestión de Accidentes - Google Maps",
    page_icon="🚨",
    layout="wide"
)

ARCHIVO_DB = "registro_accidentes_guardados.csv"

def cargar_datos():
    """Carga el historial de accidentes o crea la estructura si no existe."""
    columnas = [
        "ID", "Fecha", "Hora", "Unidad", "Conductor", "Tipo_Accidente",
        "Severidad", "Descripcion", "Latitud", "Longitud", "Link_Google_Maps"
    ]
    if os.path.exists(ARCHIVO_DB):
        try:
            df = pd.read_csv(ARCHIVO_DB)
            for col in columnas:
                if col not in df.columns:
                    df[col] = ""
            return df
        except Exception:
            return pd.DataFrame(columns=columnas)
    return pd.DataFrame(columns=columnas)

def guardar_accidente(nuevo_registro):
    """Guarda un nuevo accidente en el archivo CSV y en la sesión."""
    df = cargar_datos()
    df_nuevo = pd.DataFrame([nuevo_registro])
    df_final = pd.concat([df, df_nuevo], ignore_index=True)
    df_final.to_csv(ARCHIVO_DB, index=False, encoding="utf-8-sig")
    return df_final

def generar_pdf_accidente(reg):
    """Genera un archivo PDF descargable con el detalle y link de ubicación del accidente."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # Encabezado
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "REPORTE OFICIAL DE INCIDENTE / ACCIDENTE VIAL")
    
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 68, f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    c.line(50, height - 75, width - 50, height - 75)

    # Datos del evento
    y = height - 105
    campos = [
        ("ID del Evento:", str(reg.get("ID", ""))),
        ("Fecha y Hora:", f"{reg.get('Fecha', '')} a las {reg.get('Hora', '')}"),
        ("Unidad / Equipo:", str(reg.get("Unidad", ""))),
        ("Conductor:", str(reg.get("Conductor", ""))),
        ("Tipo de Incidente:", str(reg.get("Tipo_Accidente", ""))),
        ("Severidad:", str(reg.get("Severidad", ""))),
        ("Coordenadas GPS:", f"Lat: {reg.get('Latitud', '')}, Lon: {reg.get('Longitud', '')}"),
        ("Enlace Google Maps:", str(reg.get("Link_Google_Maps", "")))
    ]

    for titulo, valor in campos:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y, titulo)
        c.setFont("Helvetica", 11)
        c.drawString(180, y, valor)
        y -= 24

    # Descripción
    y -= 10
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, y, "Descripción de los Hechos:")
    y -= 20
    c.setFont("Helvetica", 10)
    
    texto_desc = str(reg.get("Descripcion", "Sin descripción"))
    # Dividir texto largo en líneas de 80 caracteres
    lineas = [texto_desc[i:i+80] for i in range(0, len(texto_desc), 80)]
    for linea in lineas[:10]:
        c.drawString(50, y, linea)
        y -= 16

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# Inicializar coordenadas seleccionadas en session_state
if "lat_sel" not in st.session_state:
    st.session_state.lat_sel = 14.286111  # Ubicación inicial (El Ronco / Metapán, El Salvador)
if "lon_sel" not in st.session_state:
    st.session_state.lon_sel = -89.458611
if "punto_marcado" not in st.session_state:
    st.session_state.punto_marcado = False

st.title("🚨 Plataforma de Registro de Accidentes con Google Maps")
st.markdown("Selecciona directamente en el **Mapa de Google Maps** el punto exacto donde ocurrió el accidente para capturar y guardar su ubicación automáticamente.")

col_mapa, col_form = st.columns([1.35, 1])

with col_mapa:
    st.subheader("🗺️ 1. Selecciona el Punto en Google Maps")
    
    # Selector de estilo de mapa de Google
    tipo_mapa = st.radio(
        "Vista de Google Maps:",
        ["Google Maps (Calles)", "Google Satélite (Híbrido con Calles)", "Google Satélite Puro"],
        horizontal=True
    )

    # Configurar URL de teselas oficiales de Google Maps
    if tipo_mapa == "Google Maps (Calles)":
        tiles_url = "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}"
    elif tipo_mapa == "Google Satélite (Híbrido con Calles)":
        tiles_url = "https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}"
    else:
        tiles_url = "https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}"

    # Crear mapa centrado en la última ubicación seleccionada
    m = folium.Map(
        location=[st.session_state.lat_sel, st.session_state.lon_sel],
        zoom_start=15,
        tiles=tiles_url,
        attr="Google Maps"
    )

    # Control de coordenadas al hacer clic
    m.add_child(folium.LatLngPopup())

    # Si ya hay un punto seleccionado, dibujar el marcador rojo en el mapa
    link_actual = f"https://www.google.com/maps?q={st.session_state.lat_sel:.6f},{st.session_state.lon_sel:.6f}"
    folium.Marker(
        location=[st.session_state.lat_sel, st.session_state.lon_sel],
        popup=folium.Popup(f"<b>Punto del Accidente</b><br><a href='{link_actual}' target='_blank'>Abrir en Google Maps</a>", max_width=250),
        tooltip="Ubicación seleccionada del accidente",
        icon=folium.Icon(color="red", icon="exclamation-sign")
    ).add_to(m)

    # Renderizar el mapa interactivo y escuchar el clic del usuario
    mapa_output = st_folium(m, width="100%", height=480, key="mapa_google")

    # Si el usuario hace clic en un nuevo punto del mapa, actualizar estado inmediatamente
    if mapa_output and mapa_output.get("last_clicked"):
        nuevo_lat = mapa_output["last_clicked"]["lat"]
        nuevo_lon = mapa_output["last_clicked"]["lng"]
        if (abs(nuevo_lat - st.session_state.lat_sel) > 1e-6) or (abs(nuevo_lon - st.session_state.lon_sel) > 1e-6):
            st.session_state.lat_sel = nuevo_lat
            st.session_state.lon_sel = nuevo_lon
            st.session_state.punto_marcado = True
            st.rerun()

    # Mostrar tarjeta con el enlace de Google Maps generado automáticamente
    link_google = f"https://www.google.com/maps?q={st.session_state.lat_sel:.6f},{st.session_state.lon_sel:.6f}"
    st.info(
        f"📍 **Coordenadas capturadas:** `{st.session_state.lat_sel:.6f}, {st.session_state.lon_sel:.6f}`  \n"
        f"🔗 **Enlace directo:** [Abrir ubicación exacta en Google Maps]({link_google})"
    )

with col_form:
    st.subheader("📝 2. Datos del Accidente a Guardar")
    with st.form("formulario_accidente", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            fecha_acc = st.date_input("Fecha del Accidente", value=datetime.today())
            unidad = st.selectbox(
                "Unidad / Dispositivo",
                [
                    "C-11 (Harol Magaña)",
                    "C-12 (Juan Martínez)",
                    "C-13 (Walter Monroy)",
                    "C-14 (Marvin Osorio)",
                    "C-15 (Francisco Galdámez)",
                    "C-16",
                    "C-17 (Julio Monterroza)",
                    "Otro Equipo"
                ]
            )
            tipo_acc = st.selectbox(
                "Tipo de Incidente",
                ["Colisión vehicular", "Daño mecánico en ruta", "Salida de vía / Cuneta", "Incidente en planta", "Obstáculo / Derrumbe", "Otro"]
            )
        with c2:
            hora_acc = st.time_input("Hora del Accidente", value=datetime.now().time())
            conductor = st.text_input("Nombre del Conductor / Responsable", value="")
            severidad = st.selectbox("Nivel de Severidad", ["Leve", "Moderado", "Grave", "Crítico"])

        descripcion = st.text_area("Descripción detallada del accidente:", placeholder="Describe qué ocurrió, daños observados y acciones tomadas...")
        
        st.markdown("**Ubicación que se guardará automáticamente:**")
        st.code(link_google, language="text")

        btn_guardar = st.form_submit_button("💾 Guardar Accidente y Ubicación", use_container_width=True)

        if btn_guardar:
            df_actual = cargar_datos()
            nuevo_id = f"ACC-{len(df_actual) + 1:04d}"
            nuevo_registro = {
                "ID": nuevo_id,
                "Fecha": fecha_acc.strftime("%Y-%m-%d"),
                "Hora": hora_acc.strftime("%H:%M:%S"),
                "Unidad": unidad,
                "Conductor": conductor if conductor.strip() else "Por confirmar",
                "Tipo_Accidente": tipo_acc,
                "Severidad": severidad,
                "Descripcion": descripcion if descripcion.strip() else "Sin observaciones adicionales",
                "Latitud": round(st.session_state.lat_sel, 6),
                "Longitud": round(st.session_state.lon_sel, 6),
                "Link_Google_Maps": link_google
            }
            guardar_accidente(nuevo_registro)
            st.success(f"✅ ¡Accidente **{nuevo_id}** guardado exitosamente con su enlace de Google Maps!")

st.divider()

# SECCIÓN DE HISTORIAL, MAPA GENERAL Y DESCARGAS (EXCEL / PDF)
st.subheader("📊 3. Historial de Ubicaciones Guardadas y Reportes")
df_hist = cargar_datos()

if not df_hist.empty:
    tab1, tab2, tab3 = st.tabs(["📋 Tabla con Links a Google Maps", "🗺️ Mapa General de Accidentes", "📈 Estadísticas y Exportación"])
    
    with tab1:
        st.dataframe(
            df_hist,
            column_config={
                "Link_Google_Maps": st.column_config.LinkColumn(
                    "Ubicación en Google Maps",
                    help="Haz clic para abrir el punto exacto en Google Maps",
                    display_text="📍 Abrir en Google Maps"
                )
            },
            use_container_width=True,
            hide_index=True
        )

        # Generar PDF individual de cualquier accidente guardado
        c_pdf1, c_pdf2 = st.columns([2, 1])
        with c_pdf1:
            id_seleccionado = st.selectbox("Selecciona un registro para descargar su Ficha PDF:", df_hist["ID"].tolist())
        with c_pdf2:
            reg_sel = df_hist[df_hist["ID"] == id_seleccionado].iloc[0].to_dict()
            pdf_bytes = generar_pdf_accidente(reg_sel)
            st.write("")
            st.download_button(
                label=f"📄 Descargar Reporte PDF ({id_seleccionado})",
                data=pdf_bytes,
                file_name=f"Reporte_{id_seleccionado}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    with tab2:
        st.markdown("Vista consolidada de todos los accidentes registrados sobre **Google Maps**:")
        m_hist = folium.Map(
            location=[df_hist["Latitud"].astype(float).mean(), df_hist["Longitud"].astype(float).mean()],
            zoom_start=12,
            tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
            attr="Google Maps"
        )
        for _, row in df_hist.iterrows():
            folium.Marker(
                location=[float(row["Latitud"]), float(row["Longitud"])],
                popup=folium.Popup(
                    f"<b>{row['ID']} - {row['Unidad']}</b><br>"
                    f"Fecha: {row['Fecha']} {row['Hora']}<br>"
                    f"Tipo: {row['Tipo_Accidente']} ({row['Severidad']})<br>"
                    f"<a href='{row['Link_Google_Maps']}' target='_blank'>📍 Abrir en Google Maps</a>",
                    max_width=280
                ),
                tooltip=f"{row['ID']} | {row['Unidad']} ({row['Fecha']})",
                icon=folium.Icon(color="red" if row["Severidad"] in ["Grave", "Crítico"] else "orange", icon="info-sign")
            ).add_to(m_hist)
        st_folium(m_hist, width="100%", height=450, key="mapa_historial")

    with tab3:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig_sev = px.pie(df_hist, names="Severidad", title="Accidentes por Nivel de Severidad", hole=0.4)
            st.plotly_chart(fig_sev, use_container_width=True)
        with col_g2:
            fig_uni = px.bar(df_hist, x="Unidad", color="Tipo_Accidente", title="Incidentes por Unidad y Tipo")
            st.plotly_chart(fig_uni, use_container_width=True)

        # Botón para descargar todo el historial en Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_hist.to_excel(writer, index=False, sheet_name="Accidentes_GoogleMaps")
        excel_buffer.seek(0)

        st.download_button(
            label="📥 Descargar Base Completa en Excel (.xlsx)",
            data=excel_buffer,
            file_name=f"Historial_Accidentes_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Aún no hay accidentes guardados. Haz clic en el mapa de Google Maps arriba y presiona **'Guardar Accidente y Ubicación'** para crear el primer registro.")
