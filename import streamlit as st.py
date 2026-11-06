#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script Definitivo v7: Consolidación por Brecha >= 4h en Horas y Minutos
- Permite seleccionar el archivo Excel (.xlsx) mediante ventana flotante desde cualquier ubicación.
- Aplica la regla de brecha de 4 horas para separar turnos u ocasiones de espera independientes.
- Muestra todos los tiempos en formato legible "Xh Ym" (Ej. 3h 47m).
- Genera y abre automáticamente un Dashboard Interactivo en HTML con filtros en tiempo real.
"""

import os
import re
import sys
import json
import webbrowser
import pandas as pd
import numpy as np

def convertir_a_hm(horas_dec):
    if pd.isna(horas_dec):
        return "0h 00m"
    total_min = round(horas_dec * 60)
    h = total_min // 60
    m = total_min % 60
    return f"{h}h {m:02d}m"

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Interactivo de Cargas en Planta (Brecha >= 4h)</title>
    <script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
    <style>
        :root {
            --bg: #f4f6f9;
            --card: #ffffff;
            --primary: #1f77b4;
            --secondary: #ff7f0e;
            --dark: #1e293b;
            --border: #e2e8f0;
        }
        * { box-sizing: border-box; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
        body { margin: 0; padding: 20px; background-color: var(--bg); color: var(--dark); }
        .header {
            background: var(--card); padding: 18px 24px; border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.06); margin-bottom: 18px;
            display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;
        }
        .header h1 { margin: 0; font-size: 20px; color: #0f172a; }
        .header p { margin: 4px 0 0; font-size: 13px; color: #64748b; }
        .filters-bar {
            background: var(--card); padding: 16px 24px; border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.06); margin-bottom: 18px;
            display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 18px; align-items: start;
        }
        .filter-group { display: flex; flex-direction: column; gap: 6px; }
        .filter-title {
            font-size: 13px; font-weight: 700; color: #334155;
            display: flex; justify-content: space-between; align-items: center;
        }
        .filter-title button {
            background: none; border: none; color: var(--primary); font-size: 11px; cursor: pointer; font-weight: 600; padding: 0;
        }
        .filter-title button:hover { text-decoration: underline; }
        .checkbox-list {
            max-height: 125px; overflow-y: auto; border: 1px solid var(--border);
            border-radius: 6px; padding: 8px 10px; background: #f8fafc;
            display: flex; flex-direction: column; gap: 5px;
        }
        .checkbox-item { display: flex; align-items: center; gap: 8px; font-size: 12.5px; cursor: pointer; user-select: none; }
        .checkbox-item input { cursor: pointer; accent-color: var(--primary); }
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 18px; }
        .kpi-card {
            background: var(--card); padding: 14px 18px; border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05); border-left: 4px solid var(--primary);
        }
        .kpi-card.pm { border-left-color: var(--secondary); }
        .kpi-card.neutral { border-left-color: #10b981; }
        .kpi-label { font-size: 12px; color: #64748b; font-weight: 600; text-transform: uppercase; }
        .kpi-val { font-size: 22px; font-weight: 800; margin-top: 4px; color: #0f172a; }
        .kpi-sub { font-size: 11.5px; color: #64748b; margin-top: 2px; }
        .charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-bottom: 18px; }
        @media (max-width: 1100px) { .charts-grid { grid-template-columns: 1fr; } }
        .chart-card { background: var(--card); padding: 14px; border-radius: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); min-height: 480px; }
        .table-card { background: var(--card); padding: 18px 24px; border-radius: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); overflow-x: auto; }
        .table-card h3 { margin: 0 0 12px 0; font-size: 16px; }
        table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
        th, td { padding: 9px 12px; text-align: left; border-bottom: 1px solid var(--border); }
        th { background: #f1f5f9; font-weight: 700; color: #334155; position: sticky; top: 0; }
        tr:hover { background-color: #f8fafc; }
        .badge-am { background: #dbeafe; color: #1e40af; padding: 3px 8px; border-radius: 12px; font-weight: 700; font-size: 11px; }
        .badge-pm { background: #ffedd5; color: #9a3412; padding: 3px 8px; border-radius: 12px; font-weight: 700; font-size: 11px; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>Dashboard Interactivo: Cargas en Planta (Brecha >= 4h)</h1>
            <p>Zona: <strong>PLANTA EL RONCO SOLIDARISTAS</strong> | Tiempos en formato <strong>Horas y Minutos</strong></p>
        </div>
    </div>
    <div class="filters-bar">
        <div class="filter-group">
            <div class="filter-title"><span>📅 Fechas</span><button onclick="toggleAll('fecha', true)">Todas</button></div>
            <div class="checkbox-list" id="filter-fechas"></div>
        </div>
        <div class="filter-group">
            <div class="filter-title"><span>🚛 Dispositivos</span><button onclick="toggleAll('equipo', true)">Todos</button></div>
            <div class="checkbox-list" id="filter-equipos"></div>
        </div>
        <div class="filter-group">
            <div class="filter-title"><span>🕒 Turnos</span></div>
            <div class="checkbox-list" id="filter-turnos"></div>
        </div>
        <div class="filter-group">
            <div class="filter-title"><span>⚙️ Métrica</span></div>
            <div class="checkbox-list" style="justify-content: center; gap: 6px;">
                <label class="checkbox-item"><input type="radio" name="metrica" value="neto" checked onchange="actualizarDashboard()"><span>Tiempo Neto</span></label>
                <label class="checkbox-item"><input type="radio" name="metrica" value="bruto" onchange="actualizarDashboard()"><span>Tiempo Bruto</span></label>
            </div>
        </div>
    </div>
    <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Total en Planta</div><div class="kpi-val" id="kpi-total-h">0h 00m</div><div class="kpi-sub" id="kpi-total-cargas">0 ocasiones</div></div>
        <div class="kpi-card"><div class="kpi-label">Carga AM (Mañana)</div><div class="kpi-val" id="kpi-am-h" style="color: #1f77b4;">0h 00m</div><div class="kpi-sub" id="kpi-am-cargas">0 cargas</div></div>
        <div class="kpi-card pm"><div class="kpi-label">Carga PM (Tarde)</div><div class="kpi-val" id="kpi-pm-h" style="color: #d95f02;">0h 00m</div><div class="kpi-sub" id="kpi-pm-cargas">0 cargas</div></div>
        <div class="kpi-card neutral"><div class="kpi-label">Distancia Total</div><div class="kpi-val" id="kpi-dist">0.0 km</div><div class="kpi-sub" id="kpi-equipos-act">0 equipos</div></div>
    </div>
    <div class="charts-grid">
        <div class="chart-card"><div id="chart-timeline" style="width:100%; height:460px;"></div></div>
        <div class="chart-card"><div id="chart-bars" style="width:100%; height:460px;"></div></div>
    </div>
    <div class="table-card">
        <h3>📋 Detalle Consolidado (Brecha >= 4h)</h3>
        <table>
            <thead>
                <tr><th>Fecha</th><th>Equipo</th><th>Turno / Ocasión</th><th>GPS</th><th>Entrada</th><th>Salida</th><th>Neto (h y m)</th><th>Bruto (h y m)</th><th>Distancia (km)</th></tr>
            </thead>
            <tbody id="table-body"></tbody>
        </table>
    </div>
<script>
const RAW_DATA = __JSON_DATA__;

function horasA_HM(horas) {
    const totalMin = Math.round(horas * 60);
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return `${h}h ${m < 10 ? '0' : ''}${m}m`;
}

function initFilters() {
    const fechas = [...new Set(RAW_DATA.map(d => d.Fecha))].sort();
    const equipos = [...new Set(RAW_DATA.map(d => d.Equipo))].sort();
    const turnos = [...new Set(RAW_DATA.map(d => d.Turno))].sort();
    const fC = document.getElementById('filter-fechas');
    fechas.forEach(f => fC.innerHTML += `<label class="checkbox-item"><input type="checkbox" class="chk-fecha" value="${f}" checked onchange="actualizarDashboard()"><span>${f}</span></label>`);
    const eC = document.getElementById('filter-equipos');
    equipos.forEach(e => eC.innerHTML += `<label class="checkbox-item"><input type="checkbox" class="chk-equipo" value="${e}" checked onchange="actualizarDashboard()"><span>${e}</span></label>`);
    const tC = document.getElementById('filter-turnos');
    turnos.forEach(t => tC.innerHTML += `<label class="checkbox-item"><input type="checkbox" class="chk-turno" value="${t}" checked onchange="actualizarDashboard()"><span>${t}</span></label>`);
}
function toggleAll(tipo, estado) {
    document.querySelectorAll(`.chk-${tipo}`).forEach(chk => chk.checked = estado);
    actualizarDashboard();
}
function getSelected(tipo) {
    return Array.from(document.querySelectorAll(`.chk-${tipo}:checked`)).map(el => el.value);
}
function actualizarDashboard() {
    const selF = getSelected('fecha');
    const selE = getSelected('equipo');
    const selT = getSelected('turno');
    const metrica = document.querySelector('input[name="metrica"]:checked').value;
    const filtered = RAW_DATA.filter(d => selF.includes(d.Fecha) && selE.includes(d.Equipo) && selT.includes(d.Turno));
    const campoH = metrica === 'neto' ? 'tiempo_neto_h' : 'tiempo_bruto_h';
    
    const totalH = filtered.reduce((a, r) => a + r[campoH], 0);
    const amR = filtered.filter(r => r.Turno.includes('AM'));
    const pmR = filtered.filter(r => r.Turno.includes('PM'));
    const amH = amR.reduce((a, r) => a + r[campoH], 0);
    const pmH = pmR.reduce((a, r) => a + r[campoH], 0);
    const totalDist = filtered.reduce((a, r) => a + r.distancia_total_km, 0);

    document.getElementById('kpi-total-h').innerText = horasA_HM(totalH);
    document.getElementById('kpi-total-cargas').innerText = `${filtered.length} ocasiones`;
    document.getElementById('kpi-am-h').innerText = horasA_HM(amH);
    document.getElementById('kpi-am-cargas').innerText = `${amR.length} cargas`;
    document.getElementById('kpi-pm-h').innerText = horasA_HM(pmH);
    document.getElementById('kpi-pm-cargas').innerText = `${pmR.length} cargas`;
    document.getElementById('kpi-dist').innerText = `${totalDist.toFixed(1)} km`;

    const cats = [...new Set(filtered.map(r => r.Equipo_Fecha))].sort().reverse();
    const tAM = { type: 'bar', orientation: 'h', name: 'AM', y: [], x: [], base: [], text: [], marker: { color: '#1f77b4' } };
    const tPM = { type: 'bar', orientation: 'h', name: 'PM', y: [], x: [], base: [], text: [], marker: { color: '#ff7f0e' } };

    filtered.forEach(r => {
        const span = Math.max(r.end_hour_dec - r.start_hour_dec, 0.15);
        const target = r.Turno.includes('AM') ? tAM : tPM;
        target.y.push(r.Equipo_Fecha);
        target.base.push(r.start_hour_dec);
        target.x.push(span);
        target.text.push(horasA_HM(r[campoH]));
    });

    Plotly.newPlot('chart-timeline', [tAM, tPM], {
        title: { text: '<b>Línea de Tiempo (Brecha >= 4h)</b>' }, barmode: 'overlay',
        xaxis: { range: [0, 18.5], dtick: 2 }, yaxis: { categoryorder: 'array', categoryarray: cats },
        margin: { l: 180, r: 20, t: 40, b: 50 }
    }, { responsive: true });

    const eqMap = {};
    filtered.forEach(r => {
        if (!eqMap[r.Equipo]) eqMap[r.Equipo] = { am: 0, pm: 0, total: 0 };
        if (r.Turno.includes('AM')) eqMap[r.Equipo].am += r[campoH];
        else eqMap[r.Equipo].pm += r[campoH];
        eqMap[r.Equipo].total += r[campoH];
    });
    const sEq = Object.keys(eqMap).sort((a, b) => eqMap[a].total - eqMap[b].total);
    
    Plotly.newPlot('chart-bars', [
        { type: 'bar', orientation: 'h', name: 'AM', y: sEq, x: sEq.map(e => eqMap[e].am), text: sEq.map(e => horasA_HM(eqMap[e].am)), textposition: 'outside', marker: { color: '#1f77b4' } },
        { type: 'bar', orientation: 'h', name: 'PM', y: sEq, x: sEq.map(e => eqMap[e].pm), text: sEq.map(e => horasA_HM(eqMap[e].pm)), textposition: 'outside', marker: { color: '#ff7f0e' } }
    ], {
        title: { text: '<b>Acumulado por Equipo (Horas y Minutos)</b>' }, barmode: 'group',
        margin: { l: 170, r: 50, t: 40, b: 50 }
    }, { responsive: true });

    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';
    filtered.forEach(r => {
        const badge = r.Turno.includes('AM') ? `<span class="badge-am">AM</span>` : `<span class="badge-pm">PM</span>`;
        tbody.innerHTML += `<tr><td>${r.Fecha}</td><td>${r.Equipo}</td><td>${badge} (${r.ocasion_id})</td><td>${r.fragmentos_gps}</td><td>${r.hora_inicio}</td><td>${r.hora_fin}</td><td><strong>${horasA_HM(r.tiempo_neto_h)}</strong></td><td>${horasA_HM(r.tiempo_bruto_h)}</td><td>${r.distancia_inicial_km.toFixed(1)}km</td></tr>`;
    });
}
window.onload = () => { initFilters(); actualizarDashboard(); };
</script>
</body>
</html>
"""

def seleccionar_archivo():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        ruta = filedialog.askopenfilename(
            title="Seleccione el archivo Excel (.xlsx)",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        root.destroy()
        if ruta:
            return os.path.abspath(ruta)
    except Exception as e:
        print(f"Modo gráfico no disponible: {e}")
    ruta_manual = input("Ingrese la ruta completa de su archivo Excel: ").strip().strip('"').strip("'")
    if os.path.exists(ruta_manual):
        return os.path.abspath(ruta_manual)
    return None

def ejecutar_proceso():
    filepath = seleccionar_archivo()
    if not filepath:
        print("No se seleccionó ningún archivo.")
        return

    print(f"Procesando archivo: {filepath}")
    xls = pd.ExcelFile(filepath)
    hoja = 'Report' if 'Report' in xls.sheet_names else xls.sheet_names[0]

    raw_df = pd.read_excel(filepath, sheet_name=hoja, header=None)
    filas_enc = raw_df[raw_df.iloc[:, 0].astype(str).str.strip() == 'Dispositivo'].index
    skip_rows = int(filas_enc[0]) if len(filas_enc) > 0 else 10

    df = pd.read_excel(filepath, sheet_name=hoja, skiprows=skip_rows)
    df = df.dropna(subset=['Dispositivo', 'Llegó', 'Salió']).copy()

    df['Distancia hasta la llegada'] = pd.to_numeric(df['Distancia hasta la llegada'], errors='coerce').fillna(0.0)
    df['Llegó'] = pd.to_datetime(df['Llegó'])
    df['Salió'] = pd.to_datetime(df['Salió'])
    df['duracion_horas'] = (df['Salió'] - df['Llegó']).dt.total_seconds() / 3600.0
    df['Fecha'] = df['Llegó'].dt.strftime('%Y-%m-%d')

    def limpiar_equipo(row):
        disp = str(row['Dispositivo']).strip()
        nombre = str(row['Nombre']).strip() if pd.notna(row['Nombre']) else '0'
        apellido = str(row['Apellido']).strip() if pd.notna(row['Apellido']) else '0'
        match_codigo = re.search(r'(C-\d+)', disp)
        codigo = match_codigo.group(1) if match_codigo else disp
        if nombre not in ['0', '', 'nan', 'None']:
            return f"{codigo} ({nombre.split()[0]} {apellido.split()[0] if apellido not in ['0','','nan','None'] else ''})".strip()
        parentesis = re.findall(r'\(([^)]+)\)', disp)
        if len(parentesis) >= 2:
            return f"{codigo} ({parentesis[-1].strip()})"
        return f"{codigo} (Sin conductor)"

    df['Equipo'] = df.apply(limpiar_equipo, axis=1)
    df = df.sort_values(by=['Dispositivo', 'Llegó']).reset_index(drop=True)

    # REGLA DE BRECHA DE 4 HORAS (GAP >= 4h)
    df['prev_salio'] = df.groupby('Dispositivo')['Salió'].shift(1)
    df['gap_horas'] = (df['Llegó'] - df['prev_salio']).dt.total_seconds() / 3600.0
    df['nueva_ocasion'] = df['gap_horas'].isna() | (df['gap_horas'] >= 4.0)
    df['ocasion_id'] = df.groupby('Dispositivo')['nueva_ocasion'].cumsum()

    ocasion_df = df.groupby(['Equipo', 'Dispositivo', 'Fecha', 'ocasion_id'], as_index=False).agg(
        fragmentos_gps=('Zona', 'count'),
        primera_llegada=('Llegó', 'min'),
        ultima_salida=('Salió', 'max'),
        tiempo_neto_h=('duracion_horas', 'sum'),
        distancia_inicial_km=('Distancia hasta la llegada', 'first'),
        distancia_total_km=('Distancia hasta la llegada', 'sum')
    )

    ocasion_df['tiempo_bruto_h'] = (ocasion_df['ultima_salida'] - ocasion_df['primera_llegada']).dt.total_seconds() / 3600.0
    ocasion_df['Turno'] = np.where(ocasion_df['primera_llegada'].dt.hour < 12, 'Carga AM (Mañana)', 'Carga PM (Tarde)')

    ocasion_df['Equipo_Fecha'] = ocasion_df['Fecha'].str.slice(5, 10) + ' | ' + ocasion_df['Equipo']
    ocasion_df['start_hour_dec'] = ocasion_df['primera_llegada'].dt.hour + ocasion_df['primera_llegada'].dt.minute / 60.0 + ocasion_df['primera_llegada'].dt.second / 3600.0
    ocasion_df['end_hour_dec'] = ocasion_df['ultima_salida'].dt.hour + ocasion_df['ultima_salida'].dt.minute / 60.0 + ocasion_df['ultima_salida'].dt.second / 3600.0
    ocasion_df['hora_inicio'] = ocasion_df['primera_llegada'].dt.strftime('%H:%M:%S')
    ocasion_df['hora_fin'] = ocasion_df['ultima_salida'].dt.strftime('%H:%M:%S')
    ocasion_df = ocasion_df.sort_values(by=['Fecha', 'Equipo', 'primera_llegada']).reset_index(drop=True)

    for col in ocasion_df.select_dtypes(include=['datetime64[ns]', 'datetime64']).columns:
        ocasion_df[col] = ocasion_df[col].astype(str)

    carpeta = os.path.dirname(os.path.abspath(filepath))
    nombre = os.path.splitext(os.path.basename(filepath))[0]

    records = ocasion_df.to_dict(orient='records')
    ruta_html = os.path.join(carpeta, f"Dashboard_Brecha4h_HM_{nombre}.html")
    with open(ruta_html, 'w', encoding='utf-8') as f:
        f.write(HTML_TEMPLATE.replace('__JSON_DATA__', json.dumps(records, ensure_ascii=False)))

    print(f"\n[OK] Dashboard interactivo (Brecha >= 4h en Horas y Minutos) generado en:\n -> {ruta_html}")
    webbrowser.open('file://' + os.path.realpath(ruta_html))

if __name__ == '__main__':
    ejecutar_proceso()