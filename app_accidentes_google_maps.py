#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Plataforma Integral de Registro de Accidentes, Gestión de Flota/Pilotos y Evidencias Multimedia
- Integra la base de datos oficial Base_Flota_Solidaristas_20261106.xlsx (61 Equipos y 37 Pilotos).
- Permite agregar nuevos Equipos y nuevos Pilotos directamente a la base de datos Excel.
- Selección del punto del accidente en Mapa Oficial de Google Maps (Calles / Satélite).
- Carga de evidencias multimedia (Fotos, Videos, Documentos) con galería interactiva.
- Generador de Informe Oficial en PDF (ReportLab) que incrusta los datos técnicos,
  el relato de la situación, el link de Google Maps y las fotografías de evidencia.
"""

import os
import io
import json
import textwrap
from datetime import datetime, date
import streamlit as st
import pandas as pd
import plotly.express as px
import folium
from streamlit_folium import st_folium
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

# Configuración general de la página
st.set_page_config(
    page_title="Gestión de Accidentes y Flota - Solidaristas",
    page_icon="🚨",
    layout="wide"
)

ARCHIVO_FLOTA = "Base_Flota_Solidaristas_20261106.xlsx"
ARCHIVO_ACCIDENTES = "registro_accidentes_solidaristas.csv"
CARPETA_EVIDENCIAS = "evidencias_accidentes"

os.makedirs(CARPETA_EVIDENCIAS, exist_ok=True)

# Semilla de respaldo con los 61 equipos y 37 pilotos del archivo Base_Flota_Solidaristas_20261106.xlsx
EQUIPOS_INICIALES = [
    {"id": 1, "codigo": "A01", "fecha_venc_rev": "", "placa": "C-125657", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 2, "codigo": "A02", "fecha_venc_rev": "", "placa": "C-125502", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 3, "codigo": "A03", "fecha_venc_rev": "", "placa": "C-125609", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 4, "codigo": "A04", "fecha_venc_rev": "", "placa": "C-125518", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 5, "codigo": "A05", "fecha_venc_rev": "", "placa": "C-125603", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 6, "codigo": "A06", "fecha_venc_rev": "", "placa": "C-125721", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 7, "codigo": "A07", "fecha_venc_rev": "", "placa": "C-125602", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 8, "codigo": "A08", "fecha_venc_rev": "", "placa": "C-125517", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 9, "codigo": "A09", "fecha_venc_rev": "", "placa": "C-125736", "anio": 2015, "marca": "INTERNATIONAL"},
    {"id": 10, "codigo": "A10", "fecha_venc_rev": "", "placa": "C-125501", "anio": 2015, "marca": "INTERNATIONAL"},
    {"id": 11, "codigo": "A11", "fecha_venc_rev": "", "placa": "C-66703", "anio": 2015, "marca": "INTERNATIONAL"},
    {"id": 12, "codigo": "A12", "fecha_venc_rev": "", "placa": "C-125742", "anio": 2015, "marca": "INTERNATIONAL"},
    {"id": 13, "codigo": "A13", "fecha_venc_rev": "", "placa": "C-125611", "anio": 2016, "marca": "INTERNATIONAL"},
    {"id": 14, "codigo": "A14", "fecha_venc_rev": "", "placa": "C-125720", "anio": 2016, "marca": "INTERNATIONAL"},
    {"id": 15, "codigo": "A15", "fecha_venc_rev": "", "placa": "C-125610", "anio": 2016, "marca": "INTERNATIONAL"},
    {"id": 16, "codigo": "A16", "fecha_venc_rev": "", "placa": "C-96846", "anio": 2016, "marca": "INTERNATIONAL"},
    {"id": 17, "codigo": "A17", "fecha_venc_rev": "", "placa": "C-118541", "anio": 2012, "marca": "FREIGHTLINER"},
    {"id": 18, "codigo": "A18", "fecha_venc_rev": "", "placa": "C-125745", "anio": 2014, "marca": "FREIGHTLINER"},
    {"id": 19, "codigo": "A19", "fecha_venc_rev": "", "placa": "C-119147", "anio": 2012, "marca": "INTERNATIONAL"},
    {"id": 20, "codigo": "A20", "fecha_venc_rev": "", "placa": "C-144118", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 21, "codigo": "A21", "fecha_venc_rev": "", "placa": "C-144119", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 22, "codigo": "C06", "fecha_venc_rev": "", "placa": "C-117153", "anio": 2013, "marca": "FREIGHTLINER CASCADIA"},
    {"id": 23, "codigo": "C07", "fecha_venc_rev": "", "placa": "C-117060", "anio": 2013, "marca": "FREIGHTLINER CASCADIA"},
    {"id": 24, "codigo": "C08", "fecha_venc_rev": "", "placa": "C-116995", "anio": 2013, "marca": "FREIGHTLINER CASCADIA"},
    {"id": 25, "codigo": "C09", "fecha_venc_rev": "", "placa": "C-119375", "anio": 2012, "marca": "FREIGHTLINER BLANCO"},
    {"id": 26, "codigo": "C10", "fecha_venc_rev": "", "placa": "C-143355", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 27, "codigo": "C11", "fecha_venc_rev": "", "placa": "C-143356", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 28, "codigo": "C12", "fecha_venc_rev": "", "placa": "C-143359", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 29, "codigo": "C13", "fecha_venc_rev": "", "placa": "C-143358", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 30, "codigo": "C14", "fecha_venc_rev": "", "placa": "C-143357", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 31, "codigo": "C15", "fecha_venc_rev": "", "placa": "C-143412", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 32, "codigo": "C16", "fecha_venc_rev": "", "placa": "C-143410", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 33, "codigo": "C17", "fecha_venc_rev": "", "placa": "C-143411", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 34, "codigo": "E03", "fecha_venc_rev": "", "placa": "CAT0324DCDFP00969", "anio": 2011, "marca": "CAT 324DL"},
    {"id": 35, "codigo": "E04", "fecha_venc_rev": "", "placa": "HHKHZ810HE0001476", "anio": 2016, "marca": "HYUNDAI ROBEX 300LC-9S"},
    {"id": 36, "codigo": "E05", "fecha_venc_rev": "", "placa": "HHKHZ810HE0001487", "anio": 2016, "marca": "HYUNDAI ROBEX 300LC-9S"},
    {"id": 37, "codigo": "G01", "fecha_venc_rev": "", "placa": "RE-1658", "anio": 2014, "marca": "CARMEX"},
    {"id": 38, "codigo": "G03", "fecha_venc_rev": "", "placa": "RE-10491", "anio": 2014, "marca": "CARMEX"},
    {"id": 39, "codigo": "G05", "fecha_venc_rev": "", "placa": "RE-10482", "anio": 2016, "marca": "CARMEX"},
    {"id": 40, "codigo": "L01", "fecha_venc_rev": "", "placa": "P-597222", "anio": 2008, "marca": "MAZDA BT-50"},
    {"id": 41, "codigo": "L02", "fecha_venc_rev": "", "placa": "P-339962", "anio": 2008, "marca": "NISSAN FRONTIER LCV"},
    {"id": 42, "codigo": "M01", "fecha_venc_rev": "", "placa": "P-199D3", "anio": 2011, "marca": "MERCEDES L0015/48"},
    {"id": 43, "codigo": "MB01", "fecha_venc_rev": "", "placa": "P-920975", "anio": 2019, "marca": "HYUNDAI COUNTY"},
    {"id": 44, "codigo": "MB02", "fecha_venc_rev": "", "placa": "P-11D28", "anio": 2019, "marca": "HYUNDAI COUNTY"},
    {"id": 45, "codigo": "MOTO", "fecha_venc_rev": "", "placa": "M-359043", "anio": 2016, "marca": "MOTO SUZUKI"},
    {"id": 46, "codigo": "P01", "fecha_venc_rev": "", "placa": "P-843791", "anio": 2019, "marca": "KIA K2700"},
    {"id": 47, "codigo": "P02", "fecha_venc_rev": "", "placa": "C-65671", "anio": 2015, "marca": "INTERNATIONAL 7600"},
    {"id": 48, "codigo": "P03", "fecha_venc_rev": "", "placa": "P-697022", "anio": 2016, "marca": "HYUNDAI H-100"},
    {"id": 49, "codigo": "PL01", "fecha_venc_rev": "", "placa": "RE-20889", "anio": 2007, "marca": "FONTAIN"},
    {"id": 50, "codigo": "PL02", "fecha_venc_rev": "", "placa": "RE-22070", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 51, "codigo": "PL03", "fecha_venc_rev": "", "placa": "RE-22071", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 52, "codigo": "PL04", "fecha_venc_rev": "", "placa": "RE-22074", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 53, "codigo": "PL05", "fecha_venc_rev": "", "placa": "RE-22116", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 54, "codigo": "PL06", "fecha_venc_rev": "", "placa": "RE-22118", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 55, "codigo": "PL07", "fecha_venc_rev": "", "placa": "RE-22119", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 56, "codigo": "PL08", "fecha_venc_rev": "", "placa": "RE-22150", "anio": 2025, "marca": "SINOTRUCK"},
    {"id": 60, "codigo": "A24", "fecha_venc_rev": "2027-03-24", "placa": "C-145231", "anio": 2026, "marca": "HOWO"},
    {"id": 61, "codigo": "A25", "fecha_venc_rev": "2027-03-24", "placa": "C-145233", "anio": 2026, "marca": "HOWO"},
    {"id": 63, "codigo": "A23", "fecha_venc_rev": "2027-03-24", "placa": "C-145243", "anio": 2026, "marca": "HOWO"},
    {"id": 64, "codigo": "A26", "fecha_venc_rev": "2027-03-24", "placa": "C-145169", "anio": 2026, "marca": "HOWO"},
    {"id": 65, "codigo": "A22", "fecha_venc_rev": "2027-03-24", "placa": "C-145246", "anio": 2026, "marca": "HOWO"}
]

PILOTOS_INICIALES = [
    {"id": 1, "nombre": "MARIO ENRIQUE FIGUEROA", "fecha_venc_lic": ""},
    {"id": 2, "nombre": "CHRISTIAN MEJIA", "fecha_venc_lic": ""},
    {"id": 3, "nombre": "ROBERTO GARCIA", "fecha_venc_lic": ""},
    {"id": 4, "nombre": "MAURICIO MUÑOZ", "fecha_venc_lic": ""},
    {"id": 5, "nombre": "OSCAR CALDERON", "fecha_venc_lic": ""},
    {"id": 6, "nombre": "RUBEN MURCIA", "fecha_venc_lic": ""},
    {"id": 7, "nombre": "DOUGLAS CLEMENTE", "fecha_venc_lic": ""},
    {"id": 8, "nombre": "HENRY ADALBERTO GARAY", "fecha_venc_lic": ""},
    {"id": 9, "nombre": "JOSE MANUEL ERAZO", "fecha_venc_lic": ""},
    {"id": 11, "nombre": "OVIDIO MAGAÑA", "fecha_venc_lic": ""},
    {"id": 12, "nombre": "SAMUEL HERNANDEZ", "fecha_venc_lic": "2027-10-08"},
    {"id": 13, "nombre": "FRANCIS HEREDIA", "fecha_venc_lic": ""},
    {"id": 14, "nombre": "WALTER MONROY", "fecha_venc_lic": ""},
    {"id": 15, "nombre": "JULIO MONTERROSA", "fecha_venc_lic": ""},
    {"id": 17, "nombre": "JOSE ROBERTO GUERRA", "fecha_venc_lic": ""},
    {"id": 18, "nombre": "OSCAR ORELLANA", "fecha_venc_lic": ""},
    {"id": 19, "nombre": "HAROLD MAGAÑA", "fecha_venc_lic": ""},
    {"id": 20, "nombre": "JUAN MARTINEZ", "fecha_venc_lic": ""},
    {"id": 22, "nombre": "MARVIN OSORIO", "fecha_venc_lic": ""},
    {"id": 23, "nombre": "MELVIN ARDON", "fecha_venc_lic": ""},
    {"id": 25, "nombre": "CHRISTIAN SINTIGO", "fecha_venc_lic": ""},
    {"id": 26, "nombre": "VICTOR GALDAMEZ", "fecha_venc_lic": ""},
    {"id": 27, "nombre": "SANTOS TEJADA", "fecha_venc_lic": ""},
    {"id": 28, "nombre": "RIGOBERTO DERAS", "fecha_venc_lic": ""},
    {"id": 29, "nombre": "RUDY FAJARDO", "fecha_venc_lic": ""},
    {"id": 30, "nombre": "GEOVANNY LONES", "fecha_venc_lic": "2027-09-12"},
    {"id": 32, "nombre": "REMIGIO ROBERTO MARADIAGA REYES", "fecha_venc_lic": "2027-09-15"},
    {"id": 34, "nombre": "GUILLERMO AYALA", "fecha_venc_lic": "2027-09-16"},
    {"id": 35, "nombre": "ADIEL ALBERTO VASQUEZ MIRANDA", "fecha_venc_lic": "2027-09-24"},
    {"id": 36, "nombre": "ROBERT ALEXI S RAMOS", "fecha_venc_lic": "2027-09-25"},
    {"id": 37, "nombre": "HENRY URYAS", "fecha_venc_lic": "2027-09-25"},
    {"id": 42, "nombre": "FRANCISCO VASQUEZ", "fecha_venc_lic": "2027-10-11"},
    {"id": 43, "nombre": "JOSE FRANCISCO FLORES REYES", "fecha_venc_lic": "2027-10-15"},
    {"id": 44, "nombre": "RAUL ANTONIO COCA ORELLANA", "fecha_venc_lic": "2027-10-15"},
    {"id": 45, "nombre": "FRANCISCO ANTONIO GALDAMEZ GOMES", "fecha_venc_lic": "2027-10-18"},
    {"id": 47, "nombre": "MARVIN ROALDO ALVARENGA CARDONA", "fecha_venc_lic": "2027-10-25"},
    {"id": 50, "nombre": "NAUN OCHOA", "fecha_venc_lic": "2027-10-28"}
]


def inicializar_base_flota():
    """Garantiza que el archivo Excel de Flota y Pilotos exista con las hojas 'Equipos' y 'Pilotos'."""
    if not os.path.exists(ARCHIVO_FLOTA):
        df_eq = pd.DataFrame(EQUIPOS_INICIALES)
        df_pi = pd.DataFrame(PILOTOS_INICIALES)
        with pd.ExcelWriter(ARCHIVO_FLOTA, engine="openpyxl") as writer:
            df_eq.to_excel(writer, sheet_name="Equipos", index=False)
            df_pi.to_excel(writer, sheet_name="Pilotos", index=False)


def cargar_flota_y_pilotos():
    """Lee las hojas Equipos y Pilotos desde el archivo Excel."""
    inicializar_base_flota()
    try:
        df_eq = pd.read_excel(ARCHIVO_FLOTA, sheet_name="Equipos")
        df_pi = pd.read_excel(ARCHIVO_FLOTA, sheet_name="Pilotos")
    except Exception:
        df_eq = pd.DataFrame(EQUIPOS_INICIALES)
        df_pi = pd.DataFrame(PILOTOS_INICIALES)

    df_eq["fecha_venc_rev"] = pd.to_datetime(df_eq["fecha_venc_rev"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    df_pi["fecha_venc_lic"] = pd.to_datetime(df_pi["fecha_venc_lic"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    return df_eq, df_pi


def guardar_flota_y_pilotos(df_eq, df_pi):
    """Guarda los DataFrames actualizados de Equipos y Pilotos en el archivo Excel."""
    with pd.ExcelWriter(ARCHIVO_FLOTA, engine="openpyxl") as writer:
        df_eq.to_excel(writer, sheet_name="Equipos", index=False)
        df_pi.to_excel(writer, sheet_name="Pilotos", index=False)


def cargar_accidentes():
    """Carga el historial de accidentes registrados."""
    columnas = [
        "ID", "Fecha", "Hora", "Codigo_Equipo", "Placa_Equipo", "Marca_Equipo", "Anio_Equipo",
        "Piloto", "Licencia_Venc", "Tipo_Accidente", "Severidad", "Lugar_Referencia",
        "Descripcion", "Acciones_Tomadas", "Latitud", "Longitud", "Link_Google_Maps", "Archivos_Evidencia"
    ]
    if os.path.exists(ARCHIVO_ACCIDENTES):
        try:
            df = pd.read_csv(ARCHIVO_ACCIDENTES)
            for col in columnas:
                if col not in df.columns:
                    df[col] = ""
            return df.fillna("")
        except Exception:
            return pd.DataFrame(columns=columnas)
    return pd.DataFrame(columns=columnas)


def guardar_accidente_db(nuevo_reg):
    """Agrega un nuevo registro de accidente al CSV."""
    df = cargar_accidentes()
    df_nuevo = pd.DataFrame([nuevo_reg])
    df_final = pd.concat([df, df_nuevo], ignore_index=True)
    df_final.to_csv(ARCHIVO_ACCIDENTES, index=False, encoding="utf-8-sig")
    return df_final


def guardar_archivos_evidencia(id_accidente, uploaded_files):
    """Guarda fotos y videos subidos en la carpeta del accidente y devuelve la lista de rutas."""
    if not uploaded_files:
        return []
    carpeta_evento = os.path.join(CARPETA_EVIDENCIAS, id_accidente)
    os.makedirs(carpeta_evento, exist_ok=True)
    rutas_guardadas = []
    for archivo in uploaded_files:
        nombre_limpio = os.path.basename(archivo.name)
        ruta_destino = os.path.join(carpeta_evento, nombre_limpio)
        with open(ruta_destino, "wb") as f:
            f.write(archivo.getbuffer())
        rutas_guardadas.append(ruta_destino)
    return rutas_guardadas


def generar_informe_pdf(reg):
    """Genera un Informe Ejecutivo en PDF con todos los detalles de la situación y fotos de evidencia."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # Banda superior institucional
    c.setFillColor(colors.HexColor("#0f172a"))
    c.rect(0, height - 75, width, 75, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(40, height - 35, "INFORME OFICIAL DE INCIDENTE / ACCIDENTE OPERATIVO")
    c.setFont("Helvetica", 10)
    c.drawString(40, height - 55, f"Folio: {reg.get('ID', '')}  |  Fecha de Emisión: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Etiqueta de Severidad
    sev = str(reg.get("Severidad", "Moderado"))
    color_sev = colors.HexColor("#dc2626") if sev in ["Grave", "Crítico"] else (
        colors.HexColor("#d97706") if sev == "Moderado" else colors.HexColor("#16a34a")
    )
    c.setFillColor(color_sev)
    c.roundRect(width - 150, height - 58, 110, 28, 5, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(width - 95, height - 48, f"SEVERIDAD: {sev.upper()}")

    y = height - 100
    c.setFillColor(colors.black)

    # SECCIÓN 1: DATOS GENERALES Y UNIDAD / PILOTO
    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(40, y - 5, width - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(48, y, "1. DATOS DEL EVENTO, UNIDAD Y PILOTO ASIGNADO")
    y -= 24

    datos_ficha = [
        ("Fecha y Hora del Suceso:", f"{reg.get('Fecha', '')} a las {reg.get('Hora', '')}"),
        ("Tipo de Situación / Incidente:", str(reg.get("Tipo_Accidente", ""))),
        ("Unidad / Equipo Involucrado:", f"Código: {reg.get('Codigo_Equipo', '')}  |  Placa: {reg.get('Placa_Equipo', '')}"),
        ("Marca y Año del Equipo:", f"{reg.get('Marca_Equipo', '')} (Año {reg.get('Anio_Equipo', '')})"),
        ("Piloto / Operador:", str(reg.get("Piloto", ""))),
        ("Vencimiento Licencia Piloto:", str(reg.get("Licencia_Venc", "No registrado")) or "No registrado"),
        ("Lugar / Referencia Vial:", str(reg.get("Lugar_Referencia", "Sin referencia")) or "Sin referencia"),
        ("Coordenadas Geográficas:", f"Latitud: {reg.get('Latitud', '')}, Longitud: {reg.get('Longitud', '')}"),
        ("Enlace Oficial Google Maps:", str(reg.get("Link_Google_Maps", "")))
    ]

    for etiqueta, valor in datos_ficha:
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9.5)
        c.drawString(48, y, etiqueta)
        if "Google Maps" in etiqueta and str(valor).startswith("http"):
            c.setFillColor(colors.HexColor("#1d4ed8"))
            c.setFont("Helvetica-Oblique", 9.5)
            c.drawString(210, y, valor)
            c.linkURL(valor, (210, y - 2, 540, y + 10), relative=0)
        else:
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 9.5)
            c.drawString(210, y, str(valor)[:75])
        y -= 17

    y -= 8

    # SECCIÓN 2: DESCRIPCIÓN DE LA SITUACIÓN SUCEDIDA
    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(40, y - 5, width - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(48, y, "2. DESCRIPCIÓN DETALLADA DE LA SITUACIÓN SUCEDIDA")
    y -= 22

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9.5)
    texto_desc = str(reg.get("Descripcion", "Sin descripción registrada.")).strip()
    lineas_desc = []
    for parrafo in texto_desc.splitlines():
        lineas_desc.extend(textwrap.wrap(parrafo, width=92) or [""])
    for linea in lineas_desc[:12]:
        c.drawString(48, y, linea)
        y -= 14

    y -= 8

    # SECCIÓN 3: ACCIONES TOMADAS Y SEGUIMIENTO
    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(40, y - 5, width - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(48, y, "3. ACCIONES INMEDIATAS Y MEDIDAS CORRECTIVAS")
    y -= 22

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9.5)
    texto_acc = str(reg.get("Acciones_Tomadas", "Sin acciones adicionales registradas.")).strip()
    lineas_acc = []
    for parrafo in texto_acc.splitlines():
        lineas_acc.extend(textwrap.wrap(parrafo, width=92) or [""])
    for linea in lineas_acc[:8]:
        c.drawString(48, y, linea)
        y -= 14

    y -= 10

    # SECCIÓN 4: EVIDENCIAS ADJUNTAS (FOTOS Y VIDEOS)
    archivos_str = str(reg.get("Archivos_Evidencia", "")).strip()
    lista_archivos = [a.strip() for a in archivos_str.split(";") if a.strip() and os.path.exists(a.strip())]

    extensiones_img = (".png", ".jpg", ".jpeg", ".webp", ".bmp")
    fotos = [f for f in lista_archivos if f.lower().endswith(extensiones_img)]
    videos_otros = [f for f in lista_archivos if not f.lower().endswith(extensiones_img)]

    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(40, y - 5, width - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(48, y, f"4. REGISTRO DE EVIDENCIAS ADJUNTAS ({len(lista_archivos)} ARCHIVOS)")
    y -= 20

    if not lista_archivos:
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica-Oblique", 9.5)
        c.drawString(48, y, "No se adjuntaron archivos multimedia en este reporte.")
        y -= 20
    else:
        if videos_otros:
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 9.5)
            c.drawString(48, y, "Videos / Documentos registrados en expediente digital:")
            y -= 14
            c.setFont("Helvetica", 9)
            for vo in videos_otros:
                c.drawString(60, y, f"• {os.path.basename(vo)}")
                y -= 13
            y -= 6

        if fotos:
            if y < 220:
                c.showPage()
                y = height - 50
                c.setFillColor(colors.HexColor("#0f172a"))
                c.setFont("Helvetica-Bold", 12)
                c.drawString(40, y, f"ANEXO FOTOGRÁFICO DE EVIDENCIAS - FOLIO {reg.get('ID', '')}")
                y -= 25

            x_positions = [45, 315]
            img_w, img_h = 245, 165
            col_idx = 0

            for idx, ruta_foto in enumerate(fotos):
                if y - img_h < 60:
                    c.showPage()
                    y = height - 50
                    c.setFillColor(colors.HexColor("#0f172a"))
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(40, y, f"ANEXO FOTOGRÁFICO DE EVIDENCIAS (CONTINUACIÓN) - {reg.get('ID', '')}")
                    y -= 25
                    col_idx = 0

                x_actual = x_positions[col_idx]
                try:
                    img_pil = Image.open(ruta_foto).convert("RGB")
                    img_buffer = io.BytesIO()
                    img_pil.save(img_buffer, format="JPEG", quality=85)
                    img_buffer.seek(0)
                    c.drawImage(ImageReader(img_buffer), x_actual, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, anchor='c')
                    c.setStrokeColor(colors.HexColor("#cbd5e1"))
                    c.rect(x_actual, y - img_h, img_w, img_h, fill=0, stroke=1)
                    c.setFillColor(colors.HexColor("#334155"))
                    c.setFont("Helvetica", 8)
                    c.drawString(x_actual, y - img_h - 11, f"Evidencia #{idx+1}: {os.path.basename(ruta_foto)[:38]}")
                except Exception:
                    c.drawString(x_actual, y - 20, f"[No se pudo renderizar imagen: {os.path.basename(ruta_foto)}]")

                col_idx += 1
                if col_idx >= 2:
                    col_idx = 0
                    y -= (img_h + 28)

            if col_idx == 1:
                y -= (img_h + 28)

    # Líneas de firma al pie
    if y < 90:
        c.showPage()
        y = height - 80

    c.setStrokeColor(colors.HexColor("#64748b"))
    c.line(70, 55, 250, 55)
    c.line(width - 250, 55, width - 70, 55)
    c.setFillColor(colors.HexColor("#334155"))
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(160, 42, "Firma del Piloto / Operador")
    c.drawCentredString(width - 160, 42, "Firma de Supervisor / Seguridad Vial")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


# Inicializar estado del mapa en session_state
if "lat_sel" not in st.session_state:
    st.session_state.lat_sel = 14.286111  # Planta El Ronco / Metapán
if "lon_sel" not in st.session_state:
    st.session_state.lon_sel = -89.458611

# Cargar datos maestros de Equipos y Pilotos
df_equipos, df_pilotos = cargar_flota_y_pilotos()

st.title("🚨 Sistema Integral de Reporte de Accidentes, Flota y Evidencias")
st.markdown(
    "Plataforma conectada a **`Base_Flota_Solidaristas_20261106.xlsx`** con selección geográfica en **Google Maps**, "
    "gestión de **Equipos y Pilotos**, carga de **Fotos/Videos de Evidencia** y emisión de **Informes en PDF**."
)

menu_tabs = st.tabs([
    "📝 1. Registrar Accidente y Evidencias",
    "📂 2. Historial, Galería Multimedia e Informe PDF",
    "🚛 3. Base de Datos: Pilotos y Equipos (Agregar / Editar)"
])

# ==============================================================================
# PESTAÑA 1: REGISTRAR ACCIDENTE CON MAPA GOOGLE MAPS Y EVIDENCIAS
# ==============================================================================
with menu_tabs[0]:
    col_mapa, col_form = st.columns([1.25, 1.15])

    with col_mapa:
        st.subheader("🗺️ 1. Ubicación del Accidente en Google Maps")
        st.caption("Haz clic sobre el punto exacto del mapa donde ocurrió el evento para fijar las coordenadas.")

        tipo_mapa = st.radio(
            "Capas de Google Maps:",
            ["Google Maps (Calles)", "Google Satélite (Híbrido)", "Google Satélite Puro"],
            horizontal=True
        )
        if tipo_mapa == "Google Maps (Calles)":
            tiles_url = "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}"
        elif tipo_mapa == "Google Satélite (Híbrido)":
            tiles_url = "https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}"
        else:
            tiles_url = "https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}"

        m = folium.Map(
            location=[st.session_state.lat_sel, st.session_state.lon_sel],
            zoom_start=15,
            tiles=tiles_url,
            attr="Google Maps"
        )
        m.add_child(folium.LatLngPopup())

        link_actual = f"https://www.google.com/maps?q={st.session_state.lat_sel:.6f},{st.session_state.lon_sel:.6f}"
        folium.Marker(
            location=[st.session_state.lat_sel, st.session_state.lon_sel],
            popup=folium.Popup(f"<b>Punto del Accidente</b><br><a href='{link_actual}' target='_blank'>Abrir en Google Maps</a>", max_width=250),
            tooltip="Ubicación seleccionada",
            icon=folium.Icon(color="red", icon="exclamation-sign")
        ).add_to(m)

        mapa_output = st_folium(m, width="100%", height=470, key="mapa_registro_principal")

        if mapa_output and mapa_output.get("last_clicked"):
            n_lat = mapa_output["last_clicked"]["lat"]
            n_lon = mapa_output["last_clicked"]["lng"]
            if (abs(n_lat - st.session_state.lat_sel) > 1e-6) or (abs(n_lon - st.session_state.lon_sel) > 1e-6):
                st.session_state.lat_sel = n_lat
                st.session_state.lon_sel = n_lon
                st.rerun()

        link_google = f"https://www.google.com/maps?q={st.session_state.lat_sel:.6f},{st.session_state.lon_sel:.6f}"
        st.info(
            f"📍 **Coordenadas seleccionadas:** `{st.session_state.lat_sel:.6f}, {st.session_state.lon_sel:.6f}`  \n"
            f"🔗 **Enlace Google Maps:** [Abrir punto exacto en Google Maps]({link_google})"
        )

    with col_form:
        st.subheader("📋 2. Formulario del Incidente y Evidencias")

        # Construir listas formateadas desde la base de datos Excel
        opciones_equipos = [
            f"{r['codigo']} | Placa: {r['placa']} | {r['marca']} ({r['anio']})"
            for _, r in df_equipos.sort_values("codigo").iterrows()
        ]
        opciones_pilotos = [
            f"{r['nombre']}" + (f" (Lic. vence: {r['fecha_venc_lic']})" if str(r['fecha_venc_lic']).strip() else "")
            for _, r in df_pilotos.sort_values("nombre").iterrows()
        ]

        with st.form("form_registro_accidente", clear_on_submit=True):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                fecha_ev = st.date_input("Fecha del Suceso", value=date.today())
                equipo_sel = st.selectbox("Unidad / Equipo (Base Solidaristas)", opciones_equipos)
                tipo_incidente = st.selectbox(
                    "Tipo de Situación / Accidente",
                    [
                        "Colisión vehicular en ruta",
                        "Incidente en maniobra dentro de planta",
                        "Daño mecánico / Falla de equipo",
                        "Salida de vía / Cuneta",
                        "Daño a terceros / Propiedad",
                        "Derrame o pérdida de carga",
                        "Otro incidente operativo"
                    ]
                )
            with f_col2:
                hora_ev = st.time_input("Hora del Suceso", value=datetime.now().time())
                piloto_sel = st.selectbox("Piloto Asignado (Base Solidaristas)", opciones_pilotos)
                severidad_ev = st.selectbox("Nivel de Severidad", ["Leve", "Moderado", "Grave", "Crítico"])

            lugar_ref = st.text_input(
                "Lugar o Punto de Referencia (Carretera, Km, Zona de Planta):",
                placeholder="Ej. Carretera a Metapán Km 104 / Zona de Enlonado Planta El Ronco"
            )
            descripcion_ev = st.text_area(
                "Descripción detallada de la situación ocurrida:",
                height=100,
                placeholder="Relata cómo sucedió la situación, condiciones de la vía, daños materiales o personales..."
            )
            acciones_ev = st.text_area(
                "Acciones inmediatas tomadas:",
                height=75,
                placeholder="Ej. Se dio aviso a supervisión, asistencia vial en sitio, traslado de unidad..."
            )

            st.markdown("#### 📷📹 Adjuntar Evidencias del Accidente (Fotos, Videos, Documentos)")
            archivos_subidos = st.file_uploader(
                "Selecciona una o varias fotos (.jpg, .png) o videos (.mp4, .mov) del suceso:",
                type=["jpg", "jpeg", "png", "webp", "mp4", "mov", "avi", "pdf"],
                accept_multiple_files=True
            )

            btn_registrar = st.form_submit_button("💾 Guardar Registro, Evidencias y Generar Informe PDF", use_container_width=True)

            if btn_registrar:
                df_acc_actual = cargar_accidentes()
                nuevo_id = f"ACC-{len(df_acc_actual) + 1:04d}"

                # Extraer datos técnicos del equipo seleccionado
                cod_eq = equipo_sel.split(" | ")[0].strip()
                fila_eq = df_equipos[df_equipos["codigo"].astype(str) == cod_eq].iloc[0]

                # Extraer datos del piloto seleccionado
                nom_pi = piloto_sel.split(" (Lic. vence:")[0].strip()
                filas_pi = df_pilotos[df_pilotos["nombre"].astype(str) == nom_pi]
                venc_lic = filas_pi.iloc[0]["fecha_venc_lic"] if not filas_pi.empty else ""

                # Guardar archivos multimedia en disco
                rutas_evidencia = guardar_archivos_evidencia(nuevo_id, archivos_subidos)

                nuevo_registro = {
                    "ID": nuevo_id,
                    "Fecha": fecha_ev.strftime("%Y-%m-%d"),
                    "Hora": hora_ev.strftime("%H:%M:%S"),
                    "Codigo_Equipo": str(fila_eq["codigo"]),
                    "Placa_Equipo": str(fila_eq["placa"]),
                    "Marca_Equipo": str(fila_eq["marca"]),
                    "Anio_Equipo": str(fila_eq["anio"]),
                    "Piloto": nom_pi,
                    "Licencia_Venc": str(venc_lic),
                    "Tipo_Accidente": tipo_incidente,
                    "Severidad": severidad_ev,
                    "Lugar_Referencia": lugar_ref if lugar_ref.strip() else "Ver enlace Google Maps",
                    "Descripcion": descripcion_ev if descripcion_ev.strip() else "Sin descripción detallada.",
                    "Acciones_Tomadas": acciones_ev if acciones_ev.strip() else "En seguimiento por supervisión.",
                    "Latitud": round(st.session_state.lat_sel, 6),
                    "Longitud": round(st.session_state.lon_sel, 6),
                    "Link_Google_Maps": link_google,
                    "Archivos_Evidencia": ";".join(rutas_evidencia)
                }

                guardar_accidente_db(nuevo_registro)
                st.session_state["ultimo_accidente_creado"] = nuevo_registro
                st.success(f"✅ ¡Incidente **{nuevo_id}** registrado con éxito con {len(rutas_evidencia)} archivo(s) de evidencia!")

        # Si acaba de crear un accidente, mostrar botón inmediato de descarga de PDF
        if "ultimo_accidente_creado" in st.session_state:
            reg_reciente = st.session_state["ultimo_accidente_creado"]
            pdf_inmediato = generar_informe_pdf(reg_reciente)
            st.download_button(
                label=f"📄 Descargar Informe Oficial en PDF ({reg_reciente['ID']})",
                data=pdf_inmediato,
                file_name=f"Informe_Situacion_{reg_reciente['ID']}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# ==============================================================================
# PESTAÑA 2: HISTORIAL, GALERÍA DE FOTOS/VIDEOS E INFORME PDF
# ==============================================================================
with menu_tabs[1]:
    st.subheader("📂 Expedientes de Accidentes, Evidencias Multimedia e Informes PDF")
    df_hist = cargar_accidentes()

    if df_hist.empty:
        st.info("Aún no hay accidentes registrados. Utiliza la Pestaña 1 para registrar el primer evento.")
    else:
        st.dataframe(
            df_hist[[
                "ID", "Fecha", "Hora", "Codigo_Equipo", "Placa_Equipo", "Marca_Equipo",
                "Piloto", "Tipo_Accidente", "Severidad", "Lugar_Referencia", "Link_Google_Maps"
            ]],
            column_config={
                "Link_Google_Maps": st.column_config.LinkColumn(
                    "Ubicación Google Maps",
                    display_text="📍 Abrir en Google Maps"
                )
            },
            use_container_width=True,
            hide_index=True
        )

        st.divider()
        col_det, col_media = st.columns([1.1, 1.3])

        with col_det:
            st.markdown("### 📄 Generar Informe PDF y Agregar Más Evidencias")
            id_sel = st.selectbox("Selecciona el Folio del Expediente:", df_hist["ID"].tolist()[::-1])
            reg_actual = df_hist[df_hist["ID"] == id_sel].iloc[0].to_dict()

            st.markdown(
                f"**Unidad:** `{reg_actual['Codigo_Equipo']}` (Placa `{reg_actual['Placa_Equipo']}` - {reg_actual['Marca_Equipo']})  \n"
                f"**Piloto:** {reg_actual['Piloto']}  \n"
                f"**Fecha y Hora:** {reg_actual['Fecha']} {reg_actual['Hora']}  \n"
                f"**Severidad:** `{reg_actual['Severidad']}` | **Tipo:** {reg_actual['Tipo_Accidente']}  \n"
                f"**Descripción:** {reg_actual['Descripcion']}  \n"
                f"**Ubicación:** [📍 Ver punto exacto en Google Maps]({reg_actual['Link_Google_Maps']})"
            )

            pdf_expediente = generar_informe_pdf(reg_actual)
            st.download_button(
                label=f"📥 Descargar Informe Completo en PDF ({id_sel})",
                data=pdf_expediente,
                file_name=f"Informe_Accidente_{id_sel}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

            with st.expander("➕ Subir fotos o videos adicionales a este expediente"):
                nuevos_archivos = st.file_uploader(
                    "Añadir más evidencias:",
                    type=["jpg", "jpeg", "png", "webp", "mp4", "mov", "avi", "pdf"],
                    accept_multiple_files=True,
                    key=f"up_extra_{id_sel}"
                )
                if st.button("Guardar Evidencias Adicionales", key=f"btn_extra_{id_sel}"):
                    if nuevos_archivos:
                        nuevas_rutas = guardar_archivos_evidencia(id_sel, nuevos_archivos)
                        previas = [x for x in str(reg_actual.get("Archivos_Evidencia", "")).split(";") if x.strip()]
                        todas = previas + nuevas_rutas
                        df_hist.loc[df_hist["ID"] == id_sel, "Archivos_Evidencia"] = ";".join(todas)
                        df_hist.to_csv(ARCHIVO_ACCIDENTES, index=False, encoding="utf-8-sig")
                        st.success("¡Evidencias añadidas al expediente y al PDF!")
                        st.rerun()

        with col_media:
            st.markdown(f"### 🖼️🎬 Galería Multimedia del Expediente `{id_sel}`")
            archivos_str = str(reg_actual.get("Archivos_Evidencia", "")).strip()
            lista_ev = [a.strip() for a in archivos_str.split(";") if a.strip() and os.path.exists(a.strip())]

            if not lista_ev:
                st.info("Este expediente aún no tiene fotos ni videos adjuntos.")
            else:
                fotos_ev = [f for f in lista_ev if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp"))]
                videos_ev = [f for f in lista_ev if f.lower().endswith((".mp4", ".mov", ".avi"))]

                if fotos_ev:
                    st.markdown(f"**📷 Fotografías ({len(fotos_ev)}):**")
                    cols_img = st.columns(2)
                    for i, ruta_img in enumerate(fotos_ev):
                        with cols_img[i % 2]:
                            st.image(ruta_img, caption=os.path.basename(ruta_img), use_container_width=True)

                if videos_ev:
                    st.markdown(f"**📹 Videos de Evidencia ({len(videos_ev)}):**")
                    for ruta_vid in videos_ev:
                        st.video(ruta_vid)

# ==============================================================================
# PESTAÑA 3: BASE DE DATOS DE PILOTOS Y EQUIPOS (AGREGAR NUEVOS Y EXPORTAR)
# ==============================================================================
with menu_tabs[2]:
    st.subheader("🚛 Administración de Base de Datos de Flota (`Equipos`) y `Pilotos`")
    st.markdown(
        f"Actualmente la base cuenta con **{len(df_equipos)} equipos** y **{len(df_pilotos)} pilotos** activos. "
        "Todo equipo o piloto que agregues aquí se guardará directamente en **`Base_Flota_Solidaristas_20261106.xlsx`**."
    )

    col_add_eq, col_add_pi = st.columns(2)

    with col_add_eq:
        st.markdown("### ➕ Agregar Nuevo Equipo a la Flota")
        with st.form("form_nuevo_equipo", clear_on_submit=True):
            c_eq1, c_eq2 = st.columns(2)
            with c_eq1:
                nuevo_codigo = st.text_input("Código del Equipo (Ej. C18, A27):", placeholder="C18").strip().upper()
                nueva_placa = st.text_input("Placa / Serie (Ej. C-145999):", placeholder="C-145999").strip().upper()
            with c_eq2:
                nuevo_anio = st.number_input("Año del Modelo:", min_value=1990, max_value=2030, value=2026, step=1)
                nueva_marca = st.text_input("Marca / Modelo (Ej. SINOTRUCK, HOWO):", placeholder="SINOTRUCK").strip().upper()

            incluir_venc_rev = st.checkbox("Registrar fecha de vencimiento de revisión / tarjeta")
            fecha_rev = st.date_input("Fecha Vencimiento Revisión:", value=date(2027, 3, 24))

            btn_add_eq = st.form_submit_button("💾 Guardar Nuevo Equipo", use_container_width=True)
            if btn_add_eq:
                if not nuevo_codigo or not nueva_placa:
                    st.error("Por favor ingresa al menos el Código y la Placa del equipo.")
                elif nuevo_codigo in df_equipos["codigo"].astype(str).values:
                    st.warning(f"El código {nuevo_codigo} ya existe en la base de equipos.")
                else:
                    nuevo_id_eq = int(df_equipos["id"].max()) + 1 if not df_equipos.empty else 1
                    fila_nueva_eq = pd.DataFrame([{
                        "id": nuevo_id_eq,
                        "codigo": nuevo_codigo,
                        "fecha_venc_rev": fecha_rev.strftime("%Y-%m-%d") if incluir_venc_rev else "",
                        "placa": nueva_placa,
                        "anio": int(nuevo_anio),
                        "marca": nueva_marca if nueva_marca else "POR DEFINIR"
                    }])
                    df_equipos_act = pd.concat([df_equipos, fila_nueva_eq], ignore_index=True)
                    guardar_flota_y_pilotos(df_equipos_act, df_pilotos)
                    st.success(f"✅ Equipo **{nuevo_codigo} ({nueva_placa})** agregado exitosamente a la base de datos.")
                    st.rerun()

    with col_add_pi:
        st.markdown("### ➕ Agregar Nuevo Piloto al Padrón")
        with st.form("form_nuevo_piloto", clear_on_submit=True):
            nuevo_nombre_pi = st.text_input(
                "Nombre Completo del Piloto:",
                placeholder="Ej. CARLOS ALBERTO MENDEZ"
            ).strip().upper()
            incluir_lic = st.checkbox("Registrar fecha de vencimiento de licencia", value=True)
            fecha_lic = st.date_input("Fecha de Vencimiento de Licencia:", value=date(2027, 10, 15))

            btn_add_pi = st.form_submit_button("💾 Guardar Nuevo Piloto", use_container_width=True)
            if btn_add_pi:
                if not nuevo_nombre_pi:
                    st.error("Por favor escribe el nombre completo del piloto.")
                elif nuevo_nombre_pi in df_pilotos["nombre"].astype(str).values:
                    st.warning(f"El piloto {nuevo_nombre_pi} ya está registrado en el padrón.")
                else:
                    nuevo_id_pi = int(df_pilotos["id"].max()) + 1 if not df_pilotos.empty else 1
                    fila_nueva_pi = pd.DataFrame([{
                        "id": nuevo_id_pi,
                        "nombre": nuevo_nombre_pi,
                        "fecha_venc_lic": fecha_lic.strftime("%Y-%m-%d") if incluir_lic else ""
                    }])
                    df_pilotos_act = pd.concat([df_pilotos, fila_nueva_pi], ignore_index=True)
                    guardar_flota_y_pilotos(df_equipos, df_pilotos_act)
                    st.success(f"✅ Piloto **{nuevo_nombre_pi}** agregado exitosamente al padrón.")
                    st.rerun()

    st.divider()
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.markdown(f"#### 📋 Listado Oficial de Equipos ({len(df_equipos)} unidades)")
        st.dataframe(df_equipos, use_container_width=True, hide_index=True, height=360)
    with col_t2:
        st.markdown(f"#### 👨‍✈️ Listado Oficial de Pilotos ({len(df_pilotos)} pilotos)")
        st.dataframe(df_pilotos, use_container_width=True, hide_index=True, height=360)

    # Botón para descargar la base de datos Excel actualizada
    with open(ARCHIVO_FLOTA, "rb") as f_xls:
        st.download_button(
            label="📥 Descargar Base de Flota y Pilotos Actualizada (Base_Flota_Solidaristas_20261106.xlsx)",
            data=f_xls.read(),
            file_name="Base_Flota_Solidaristas_Actualizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
