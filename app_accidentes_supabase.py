#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Plataforma Integral v4: Gestión de Accidentes, Flota, Evidencias Múltiples y Supabase Cloud
- Integración nativa con Supabase (Base de datos y Storage Bucket 'evidencias').
- Carga masiva de fotos y videos (sin límite de cantidad).
- Buscador geográfico + GPS en tiempo real (Mejora 5).
- Informe oficial en PDF con mapa satelital estático y código QR escaneable hacia Google Maps (Mejora 3).
- Cruce automático de telemetría GPS con regla de brecha >= 4 horas en horas y minutos (Mejora 6).
"""

import os
import io
import re
import math
import json
import zipfile
import textwrap
import urllib.request
import urllib.parse
from datetime import datetime, date
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import folium
from streamlit_folium import st_folium
from PIL import Image, ImageDraw
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas

# Configuración general de la página con diseño corporativo moderno (Enterprise UI)
st.set_page_config(
    page_title="Gestión de Accidentes y Supabase - Solidaristas",
    page_icon="🚨",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #090d16 0%, #0f172a 100%);
        color: #f8fafc;
    }
    
    h1, h2, h3 {
        color: #ffffff !important;
        font-weight: 800;
        letter-spacing: -0.025em;
    }
    
    div[data-testid="stMetric"] {
        background: rgba(30, 41, 59, 0.7);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 16px 20px;
        border-radius: 14px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    
    div[data-testid="stMetric"] label {
        color: #94a3b8 !important;
        font-weight: 600;
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #38bdf8 !important;
        font-weight: 800;
        font-size: 1.6rem;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: 700;
        padding: 0.6rem 1.2rem;
        box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
        transition: all 0.25s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
        box-shadow: 0 6px 20px rgba(59, 130, 246, 0.6);
        transform: translateY(-2px);
    }
    
    button[kind="primary"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.4) !important;
    }
    
    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #059669 100%, #047857 100%) !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.6) !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: rgba(15, 23, 42, 0.8);
        padding: 8px;
        border-radius: 14px;
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        color: #94a3b8;
        font-weight: 700;
        padding: 10px 20px;
        transition: all 0.2s ease;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
    }
    
    section[data-testid="stSidebar"] {
        background-color: #070b12;
        border-right: 1px solid rgba(255, 255, 255, 0.08);
    }
    
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] label {
        color: #cbd5e1 !important;
    }
</style>
""", unsafe_allow_html=True)

ARCHIVO_FLOTA = "Base_Flota_Solidaristas_20261106.xlsx"
ARCHIVO_ACCIDENTES = "registro_accidentes_solidaristas.csv"
ARCHIVO_TELEMETRIA_CACHE = "telemetria_actual_cache.xlsx"
CARPETA_EVIDENCIAS = "evidencias_accidentes"

os.makedirs(CARPETA_EVIDENCIAS, exist_ok=True)

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

@st.cache_resource
def obtener_cliente_supabase():
    try:
        from supabase import create_client
        if "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
            return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        pass
    return None

supabase_client = obtener_cliente_supabase()

def convertir_a_hm(horas_dec):
    if pd.isna(horas_dec):
        return "0h 00m"
    total_min = int(round(float(horas_dec) * 60))
    h = total_min // 60
    m = total_min % 60
    return f"{h}h {m:02d}m"

def normalizar_codigo_equipo(codigo_str):
    texto = str(codigo_str).upper().strip()
    match_c = re.search(r"\b([A-Z]+)-?(\d+)\b", texto)
    if match_c:
        return f"{match_c.group(1)}{match_c.group(2)}"
    return re.sub(r"[^A-Z0-9]", "", texto)

def inicializar_base_flota():
    if supabase_client:
        try:
            res_eq = supabase_client.table("equipos").select("id", count="exact").limit(1).execute()
            if res_eq.count == 0:
                supabase_client.table("equipos").upsert(EQUIPOS_INICIALES).execute()
            res_pi = supabase_client.table("pilotos").select("id", count="exact").limit(1).execute()
            if res_pi.count == 0:
                supabase_client.table("pilotos").upsert(PILOTOS_INICIALES).execute()
            return
        except Exception:
            pass

    if not os.path.exists(ARCHIVO_FLOTA):
        df_eq = pd.DataFrame(EQUIPOS_INICIALES)
        df_pi = pd.DataFrame(PILOTOS_INICIALES)
        with pd.ExcelWriter(ARCHIVO_FLOTA, engine="openpyxl") as writer:
            df_eq.to_excel(writer, sheet_name="Equipos", index=False)
            df_pi.to_excel(writer, sheet_name="Pilotos", index=False)

@st.cache_data(ttl=60)
def cargar_flota_y_pilotos():
    inicializar_base_flota()
    if supabase_client:
        try:
            res_eq = supabase_client.table("equipos").select("*").execute()
            res_pi = supabase_client.table("pilotos").select("*").execute()
            df_eq = pd.DataFrame(res_eq.data) if res_eq.data else pd.DataFrame(EQUIPOS_INICIALES)
            df_pi = pd.DataFrame(res_pi.data) if res_pi.data else pd.DataFrame(PILOTOS_INICIALES)
            df_eq["fecha_venc_rev"] = pd.to_datetime(df_eq["fecha_venc_rev"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
            df_pi["fecha_venc_lic"] = pd.to_datetime(df_pi["fecha_venc_lic"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
            return df_eq, df_pi
        except Exception:
            pass

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
    if supabase_client:
        try:
            supabase_client.table("equipos").upsert(df_eq.to_dict(orient="records")).execute()
            supabase_client.table("pilotos").upsert(df_pi.to_dict(orient="records")).execute()
        except Exception:
            pass
    with pd.ExcelWriter(ARCHIVO_FLOTA, engine="openpyxl") as writer:
        df_eq.to_excel(writer, sheet_name="Equipos", index=False)
        df_pi.to_excel(writer, sheet_name="Pilotos", index=False)

@st.cache_data(ttl=10)
def cargar_accidentes():
    columnas = [
        "ID", "Fecha", "Hora", "Codigo_Equipo", "Placa_Equipo", "Marca_Equipo", "Anio_Equipo",
        "Piloto", "Licencia_Venc", "Tipo_Accidente", "Severidad", "Lugar_Referencia",
        "Descripcion", "Acciones_Tomadas", "Latitud", "Longitud", "Link_Google_Maps",
        "Archivos_Evidencia", "Auditoria_Telemetria"
    ]
    if supabase_client:
        try:
            res = supabase_client.table("accidentes").select("*").execute()
            if res.data:
                df = pd.DataFrame(res.data)
                for col in columnas:
                    if col not in df.columns:
                        df[col] = ""
                return df.fillna("")
        except Exception:
            pass

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
    if supabase_client:
        try:
            supabase_client.table("accidentes").upsert(nuevo_reg).execute()
        except Exception:
            pass

    df = cargar_accidentes()
    df_nuevo = pd.DataFrame([nuevo_reg])
    df_final = pd.concat([df, df_nuevo], ignore_index=True)
    df_final.to_csv(ARCHIVO_ACCIDENTES, index=False, encoding="utf-8-sig")
    return df_final

def guardar_evidencias(id_accidente, lista_items_memoria):
    if not lista_items_memoria:
        return []
    rutas_guardadas = []
    carpeta_evento = os.path.join(CARPETA_EVIDENCIAS, id_accidente)
    os.makedirs(carpeta_evento, exist_ok=True)

    for idx, item in enumerate(lista_items_memoria, start=1):
        nombre_limpio = os.path.basename(item["nombre"])
        base, ext = os.path.splitext(nombre_limpio)
        nuevo_nombre = f"EV{idx:02d}_{base}{ext}"

        if supabase_client:
            try:
                ruta_nube = f"{id_accidente}/{nuevo_nombre}"
                supabase_client.storage.from_("evidencias").upload(
                    path=ruta_nube,
                    file=item["bytes"],
                    file_options={"upsert": "true"}
                )
                url_publica = supabase_client.storage.from_("evidencias").get_public_url(ruta_nube)
                rutas_guardadas.append(url_publica)
                continue
            except Exception:
                pass

        ruta_destino = os.path.join(carpeta_evento, nuevo_nombre)
        with open(ruta_destino, "wb") as f:
            f.write(item["bytes"])
        rutas_guardadas.append(ruta_destino)

    return rutas_guardadas

def clasificar_rutas_evidencia(archivos_str):
    lista = [a.strip() for a in str(archivos_str).split(";") if a.strip()]
    ext_img = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")
    ext_vid = (".mp4", ".mov", ".avi", ".mkv", ".webm", ".mpeg4")
    fotos = [f for f in lista if f.lower().endswith(ext_img) or "supabase.co" in f.lower()]
    videos = [f for f in lista if f.lower().endswith(ext_vid)]
    otros = [f for f in lista if f not in fotos and f not in videos]
    return lista, fotos, videos, otros

def buscar_direccion_nominatim(consulta):
    if not consulta or not consulta.strip():
        return []
    try:
        q_enc = urllib.parse.quote(consulta.strip())
        url = f"https://nominatim.openstreetmap.org/search?q={q_enc}&format=json&limit=5&countrycodes=sv,gt,hn"
        req = urllib.request.Request(url, headers={"User-Agent": "SolidaristasFlotaApp/4.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            datos = json.loads(resp.read().decode("utf-8"))
            return [{"nombre": d.get("display_name", ""), "lat": float(d.get("lat", 14.286111)), "lon": float(d.get("lon", -89.458611))} for d in datos]
    except Exception:
        return []

def extraer_coordenadas_de_texto_o_link(texto):
    if not texto:
        return None, None
    patrones = [r"@(-?\d+\.\d+),\s*(-?\d+\.\d+)", r"[?&]q=(-?\d+\.\d+),\s*(-?\d+\.\d+)", r"(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)"]
    for pat in patrones:
        m = re.search(pat, str(texto))
        if m:
            return float(m.group(1)), float(m.group(2))
    return None, None

def generar_imagen_mapa_estatico(lat, lon, zoom=15):
    try:
        lat_rad = math.radians(float(lat))
        n = 2.0 ** zoom
        xtile_f = (float(lon) + 180.0) / 360.0 * n
        ytile_f = (1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n
        x_center, y_center = int(math.floor(xtile_f)), int(math.floor(ytile_f))
        px_in_tile, py_in_tile = int((xtile_f - x_center) * 256), int((ytile_f - y_center) * 256)

        lienzo = Image.new("RGB", (256 * 3, 256 * 3), (226, 232, 240))
        for dx in [-1, 0, 1]:
            for dy in [-1, 0, 1]:
                tx, ty = x_center + dx, y_center + dy
                url_tile = f"https://mt1.google.com/vt/lyrs=y&x={tx}&y={ty}&z={zoom}"
                req = urllib.request.Request(url_tile, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=3) as r_tile:
                    tile_img = Image.open(io.BytesIO(r_tile.read())).convert("RGB")
                    lienzo.paste(tile_img, ((dx + 1) * 256, (dy + 1) * 256))

        cx, cy = 256 + px_in_tile, 256 + py_in_tile
        recorte = lienzo.crop((cx - 210, cy - 115, cx + 210, cy + 115))
        draw = ImageDraw.Draw(recorte)
        mx, my = 210, 115
        draw.ellipse((mx - 14, my - 14, mx + 14, my + 14), outline="white", width=3)
        draw.ellipse((mx - 11, my - 11, mx + 11, my + 11), fill="#dc2626", outline="#7f1d1d", width=2)
        draw.ellipse((mx - 4, my - 4, mx + 4, my + 4), fill="white")

        out_buf = io.BytesIO()
        recorte.save(out_buf, format="JPEG", quality=90)
        out_buf.seek(0)
        return out_buf
    except Exception:
        return None

def generar_informe_pdf(reg):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.setFillColor(colors.HexColor("#0f172a"))
    c.rect(0, height - 82, width, 82, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 14.5)
    c.drawString(38, height - 32, "INFORME OFICIAL DE INCIDENTE / ACCIDENTE OPERATIVO")
    c.setFont("Helvetica", 9.5)
    c.drawString(38, height - 50, f"Folio: {reg.get('ID', '')}  |  Emisión: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    sev = str(reg.get("Severidad", "Moderado"))
    color_sev = colors.HexColor("#dc2626") if sev in ["Grave", "Crítico"] else colors.HexColor("#16a34a")
    c.setFillColor(color_sev)
    c.roundRect(38, height - 74, 125, 18, 4, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString(100, height - 68, f"SEVERIDAD: {sev.upper()}")

    link_maps = str(reg.get("Link_Google_Maps", "https://www.google.com/maps"))
    try:
        c.setFillColor(colors.white)
        c.roundRect(width - 102, height - 77, 68, 72, 4, fill=1, stroke=0)
        qr_widget = qr.QrCodeWidget(link_maps)
        b = qr_widget.getBounds()
        d_qr = Drawing(58, 58, transform=[58.0 / (b[2] - b[0]), 0, 0, 58.0 / (b[3] - b[1]), 0, 0])
        d_qr.add(qr_widget)
        renderPDF.draw(d_qr, c, width - 97, height - 64)
        c.setFillColor(colors.HexColor("#0f172a"))
        c.setFont("Helvetica-Bold", 6.5)
        c.drawCentredString(width - 68, height - 74, "ESCANEAR MAPA")
    except Exception:
        pass

    y = height - 102
    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(38, y - 4, width - 76, 18, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 10)
    c.drawString(44, y, "1. FICHA TÉCNICA DEL EVENTO Y UBICACIÓN GEOGRÁFICA")
    y -= 20

    y_inicio_ficha = y
    datos_ficha = [
        ("Fecha y Hora:", f"{reg.get('Fecha', '')} a las {reg.get('Hora', '')}"),
        ("Tipo Incidente:", str(reg.get("Tipo_Accidente", ""))),
        ("Unidad / Equipo:", f"{reg.get('Codigo_Equipo', '')} | Placa: {reg.get('Placa_Equipo', '')}"),
        ("Marca y Año:", f"{reg.get('Marca_Equipo', '')} ({reg.get('Anio_Equipo', '')})"),
        ("Piloto Asignado:", str(reg.get("Piloto", ""))),
        ("Referencia Vial:", str(reg.get("Lugar_Referencia", "Sin referencia"))[:42]),
        ("Coordenadas:", f"{reg.get('Latitud', '')}, {reg.get('Longitud', '')}")
    ]

    for etiqueta, valor in datos_ficha:
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 8.8)
        c.drawString(44, y, etiqueta)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8.8)
        c.drawString(130, y, str(valor)[:45])
        y -= 15

    mapa_buf = generar_imagen_mapa_estatico(reg.get("Latitud", 14.286111), reg.get("Longitud", -89.458611), zoom=15)
    if mapa_buf:
        map_w, map_h = 215, 125
        map_x, map_y = width - 38 - map_w, y_inicio_ficha - map_h + 10
        try:
            c.drawImage(ImageReader(mapa_buf), map_x, map_y, width=map_w, height=map_h, preserveAspectRatio=False)
            c.setStrokeColor(colors.HexColor("#0f172a"))
            c.rect(map_x, map_y, map_w, map_h, fill=0, stroke=1)
        except Exception:
            pass

    y = min(y, y_inicio_ficha - 135)
    texto_tele = str(reg.get("Auditoria_Telemetria", "")).strip()
    if texto_tele:
        c.setFillColor(colors.HexColor("#dbeafe"))
        c.rect(38, y - 4, width - 76, 17, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#1e3a8a"))
        c.setFont("Helvetica-Bold", 9.5)
        c.drawString(44, y, "AUDITORÍA AUTOMÁTICA DE TELEMETRÍA GPS (PLANTA EL RONCO)")
        y -= 16
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8.8)
        for lin_t in textwrap.wrap(texto_tele, width=96)[:4]:
            c.drawString(44, y, lin_t)
            y -= 12
        y -= 4

    c.setFillColor(colors.HexColor("#e2e8f0"))
    c.rect(38, y - 4, width - 76, 18, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#0f172a"))
    c.setFont("Helvetica-Bold", 10)
    c.drawString(44, y, "2. DESCRIPCIÓN DETALLADA Y ACCIONES TOMADAS")
    y -= 18

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    for linea in textwrap.wrap(str(reg.get("Descripcion", "")), width=95)[:8]:
        c.drawString(44, y, linea)
        y -= 13

    lista_archivos, fotos, videos, otros = clasificar_rutas_evidencia(reg.get("Archivos_Evidencia", ""))
    if fotos:
        c.showPage()
        y = height - 50
        c.setFillColor(colors.HexColor("#0f172a"))
        c.setFont("Helvetica-Bold", 12)
        c.drawString(38, y, f"ANEXO FOTOGRÁFICO DE EVIDENCIAS ({len(fotos)} IMÁGENES) - FOLIO {reg.get('ID', '')}")
        y -= 25

        x_positions = [42, 312]
        img_w, img_h = 248, 162
        col_idx = 0

        for idx, ruta_foto in enumerate(fotos):
            if y - img_h < 60:
                c.showPage()
                y = height - 50
                col_idx = 0

            x_actual = x_positions[col_idx]
            try:
                if str(ruta_foto).startswith("http"):
                    req_img = urllib.request.Request(ruta_foto, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req_img, timeout=5) as r_img:
                        img_pil = Image.open(io.BytesIO(r_img.read())).convert("RGB")
                else:
                    img_pil = Image.open(ruta_foto).convert("RGB")

                img_buffer = io.BytesIO()
                img_pil.save(img_buffer, format="JPEG", quality=85)
                img_buffer.seek(0)
                c.drawImage(ImageReader(img_buffer), x_actual, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, anchor='c')
                c.setStrokeColor(colors.HexColor("#cbd5e1"))
                c.rect(x_actual, y - img_h, img_w, img_h, fill=0, stroke=1)
                c.setFillColor(colors.HexColor("#334155"))
                c.setFont("Helvetica-Bold", 8)
                c.drawString(x_actual, y - img_h - 11, f"Foto #{idx+1} de {len(fotos)}")
            except Exception:
                pass

            col_idx += 1
            if col_idx >= 2:
                col_idx = 0
                y -= (img_h + 26)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def generar_zip_expediente(reg):
    zip_buffer = io.BytesIO()
    id_folio = reg.get("ID", "ACC")
    pdf_bytes = generar_informe_pdf(reg).getvalue()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"Informe_Oficial_{id_folio}.pdf", pdf_bytes)
    zip_buffer.seek(0)
    return zip_buffer

def procesar_excel_telemetria(filepath_or_buffer):
    xls = pd.ExcelFile(filepath_or_buffer)
    hoja = "Report" if "Report" in xls.sheet_names else xls.sheet_names[0]
    raw_df = pd.read_excel(filepath_or_buffer, sheet_name=hoja, header=None)
    filas_enc = raw_df[raw_df.iloc[:, 0].astype(str).str.strip() == "Dispositivo"].index
    skip_rows = int(filas_enc[0]) if len(filas_enc) > 0 else 10

    df = pd.read_excel(filepath_or_buffer, sheet_name=hoja, skiprows=skip_rows)
    df = df.dropna(subset=["Dispositivo", "Llegó", "Salió"]).copy()
    df["Distancia hasta la llegada"] = pd.to_numeric(df["Distancia hasta la llegada"], errors="coerce").fillna(0.0)
    df["Llegó"] = pd.to_datetime(df["Llegó"])
    df["Salió"] = pd.to_datetime(df["Salió"])
    df["duracion_horas"] = (df["Salió"] - df["Llegó"]).dt.total_seconds() / 3600.0
    df["Fecha"] = df["Llegó"].dt.strftime("%Y-%m-%d")

    def limpiar_equipo_tele(row):
        disp = str(row["Dispositivo"]).strip()
        match_codigo = re.search(r"(C-\d+)", disp)
        return match_codigo.group(1) if match_codigo else disp

    df["Equipo"] = df.apply(limpiar_equipo_tele, axis=1)
    df["Codigo_Norm"] = df["Dispositivo"].apply(normalizar_codigo_equipo)
    df = df.sort_values(by=["Dispositivo", "Llegó"]).reset_index(drop=True)

    df["prev_salio"] = df.groupby("Dispositivo")["Salió"].shift(1)
    df["gap_horas"] = (df["Llegó"] - df["prev_salio"]).dt.total_seconds() / 3600.0
    df["nueva_ocasion"] = df["gap_horas"].isna() | (df["gap_horas"] >= 4.0)
    df["ocasion_id"] = df.groupby("Dispositivo")["nueva_ocasion"].cumsum()

    ocasion_df = df.groupby(["Equipo", "Dispositivo", "Codigo_Norm", "Fecha", "ocasion_id"], as_index=False).agg(
        fragmentos_gps=("Zona", "count"),
        primera_llegada=("Llegó", "min"),
        ultima_salida=("Salió", "max"),
        tiempo_neto_h=("duracion_horas", "sum"),
        distancia_inicial_km=("Distancia hasta la llegada", "first"),
        distancia_total_km=("Distancia hasta la llegada", "sum")
    )
    ocasion_df["tiempo_bruto_h"] = (ocasion_df["ultima_salida"] - ocasion_df["primera_llegada"]).dt.total_seconds() / 3600.0
    ocasion_df["Turno"] = np.where(ocasion_df["primera_llegada"].dt.hour < 12, "Carga AM", "Carga PM")
    ocasion_df["Tiempo_Neto_HM"] = ocasion_df["tiempo_neto_h"].apply(convertir_a_hm)
    ocasion_df["Hora_Inicio"] = ocasion_df["primera_llegada"].dt.strftime("%H:%M:%S")
    ocasion_df["Hora_Fin"] = ocasion_df["ultima_salida"].dt.strftime("%H:%M:%S")
    return df, ocasion_df

def auditar_accidente_contra_telemetria(codigo_eq, fecha_str, hora_str, ocasion_df):
    if ocasion_df is None or ocasion_df.empty:
        return "Sin archivo de telemetría cargado para cruce."
    cod_norm = normalizar_codigo_equipo(codigo_eq)
    sub = ocasion_df[ocasion_df["Codigo_Norm"] == cod_norm].copy()
    if sub.empty:
        return f"La unidad {codigo_eq} no registra entradas en telemetría."
    try:
        dt_acc = pd.to_datetime(f"{fecha_str} {hora_str}")
    except Exception:
        return "Fecha/hora no válida."

    en_planta = sub[(sub["primera_llegada"] <= dt_acc) & (dt_acc <= sub["ultima_salida"])]
    if not en_planta.empty:
        r = en_planta.iloc[0]
        return f"EN PLANTA: Ocurrió mientras la unidad estaba en PLANTA EL RONCO ({r['Turno']}, entrada {r['Hora_Inicio']} - salida {r['Hora_Fin']}, estancia {r['Tiempo_Neto_HM']})."
    return f"EN RUTA EXTERNA: La unidad operaba fuera de planta al momento del suceso."

if "lat_sel" not in st.session_state:
    st.session_state.lat_sel = 14.286111
if "lon_sel" not in st.session_state:
    st.session_state.lon_sel = -89.458611
if "cola_evidencias" not in st.session_state:
    st.session_state.cola_evidencias = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "resultados_busqueda_mapa" not in st.session_state:
    st.session_state.resultados_busqueda_mapa = []

if "ocasion_tele_df" not in st.session_state:
    if os.path.exists(ARCHIVO_TELEMETRIA_CACHE):
        try:
            _, st.session_state.ocasion_tele_df = procesar_excel_telemetria(ARCHIVO_TELEMETRIA_CACHE)
        except Exception:
            st.session_state.ocasion_tele_df = None
    else:
        st.session_state.ocasion_tele_df = None

def exportar_backup_maestro_zip():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(ARCHIVO_FLOTA):
            zf.write(ARCHIVO_FLOTA, arcname=ARCHIVO_FLOTA)
        if os.path.exists(ARCHIVO_ACCIDENTES):
            zf.write(ARCHIVO_ACCIDENTES, arcname=ARCHIVO_ACCIDENTES)
        if os.path.exists(ARCHIVO_TELEMETRIA_CACHE):
            zf.write(ARCHIVO_TELEMETRIA_CACHE, arcname=ARCHIVO_TELEMETRIA_CACHE)
        if os.path.exists(CARPETA_EVIDENCIAS):
            for root, _, files in os.walk(CARPETA_EVIDENCIAS):
                for file in files:
                    zf.write(os.path.join(root, file), arcname=os.path.relpath(os.path.join(root, file), "."))
    buffer.seek(0)
    return buffer

def restaurar_backup_maestro_zip(uploaded_zip):
    with zipfile.ZipFile(uploaded_zip, "r") as zf:
        zf.extractall(".")
    return True

def agregar_archivos_a_cola(archivos_nuevos, categoria):
    agregados = 0
    nombres_existentes = {(item["nombre"], item["tamano"]) for item in st.session_state.cola_evidencias}
    for arch in archivos_nuevos:
        contenido = arch.getvalue()
        clave = (arch.name, len(contenido))
        if clave not in nombres_existentes:
            st.session_state.cola_evidencias.append({"nombre": arch.name, "categoria": categoria, "tamano": len(contenido), "bytes": contenido})
            nombres_existentes.add(clave)
            agregados += 1
    return agregados

df_equipos, df_pilotos = cargar_flota_y_pilotos()

with st.sidebar:
    st.header("☁️ Supabase Cloud & Respaldo")
    estado_nube = "Supabase Cloud Conectado 🟢" if supabase_client else "Almacenamiento Local + ZIP 🟡"
    st.success(f"**Estado:** {estado_nube}")
    zip_maestro = exportar_backup_maestro_zip()
    st.download_button(
        label="💾 Descargar Respaldo Maestro (.ZIP)",
        data=zip_maestro
