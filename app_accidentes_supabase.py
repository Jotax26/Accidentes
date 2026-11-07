#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Plataforma Integral Definitiva: Gestión de Accidentes, Flota, Evidencias Múltiples, Supabase y Telemetría GPS
- Diseño UI/UX Moderno (Dark Mode & Neón Elegante con Glassmorphism).
- Conexión nativa con Supabase (Base de datos PostgreSQL y Storage 'evidencias').
- Carga masiva e ilimitada de evidencias (más de 5 fotos y 5 videos por accidente).
- Buscador geográfico y geolocalización GPS en tiempo real.
- Informe Ejecutivo en PDF con Mapa Satelital Estático y Código QR escaneable hacia Google Maps.
- Motor de Auditoría de Telemetría GPS con regla de brecha >= 4 horas en Formato Horas y Minutos 'Xh Ym'.
- Respaldo Maestro Portátil en archivo .ZIP.
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

# Configuración general de la página
st.set_page_config(
    page_title="Panel de Control de Accidentes y Flota - Solidaristas",
    page_icon="🚨",
    layout="wide"
)

# --- SISTEMA DE DISEÑO UI/UX: DARK MODE & NEÓN ELEGANTE ---
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
        background: rgba(30, 41, 59, 0.75);
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

# Semilla inicial con los 61 equipos y 37 pilotos del Excel oficial
EQUIPOS_INICIALES = [
    {"id": 1, "codigo": "A01", "fecha_venc_rev": "", "placa": "C-125657", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 2, "codigo": "A02", "fecha_venc_rev": "", "placa": "C-125502", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 3, "codigo": "A03", "fecha_venc_rev": "", "placa": "C-125609", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 4, "codigo": "A04", "fecha_venc_rev": "", "placa": "C-125518", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 5, "codigo": "A05", "fecha_venc_rev": "", "placa": "C-125603", "anio": 2013, "marca": "INTERNATIONAL"},
    {"id": 6, "codigo": "A06", "fecha_venc_rev": "", "placa": "C-125721", "anio": 2013, "marca":
