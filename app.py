"""
CENACE PML Analyzer — Streamlit Cloud Edition  v5
======================================================
Recurrent Energy / Canadian Solar — SRF · Sebastian Roldan

v5 changes vs v4:
- Estado persistente con st.session_state (no se reinicia al cambiar scoring)
- Modo Completo: solo dashboard analítico (sin Excel)
- Modo Solo Datos: Centro de descargas con Excel datos + Excel análisis + KMZ opcional
- Workers fijo en 8 (sin slider)
- Optimizaciones de memoria y caching
- Leyenda de CCR en el mapa
- BESS scoring también descargable como Excel
"""

import streamlit as st
import pandas as pd
import numpy as np
import requests
import json
import urllib3
import time
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, date
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from difflib import SequenceMatcher

import plotly.express as px
import plotly.graph_objects as go

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ═══════════════════════════════════════════════════════════════════════
# CONFIG STREAMLIT
# ═══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="CENACE PML Analyzer · Recurrent Energy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Workers fijo (no más slider)
MAX_WORKERS = 8

# ═══════════════════════════════════════════════════════════════════════
# COLORES
# ═══════════════════════════════════════════════════════════════════════
RE_NAVY     = "#0e346b"
RE_RED      = "#a0090c"
RE_BLUE     = "#2777bd"
RE_ALT      = "#EBF3FB"
RE_INFO     = "#D9E2F3"
TEXT_DARK   = "#1a1a1a"
TEXT_TITLE  = "#0a2347"
GRID_LIGHT  = "#D8D8D8"
AXIS_LINE   = "#444444"

PALETTE = [
    "#0e346b", "#a0090c", "#d4a017", "#1a8a3a", "#7e57c2",
    "#ff5722", "#00897b", "#c2185b", "#3949ab", "#5d4037",
    "#1976d2", "#558b2f", "#f57c00", "#8e24aa", "#455a64",
]

C_HEADER = "0e346b"
C_SUB    = "2777bd"
C_RED    = "a0090c"
C_WHITE  = "FFFFFF"
C_ALT    = "EBF3FB"
C_INFO   = "D9E2F3"
C_GREEN  = "1a8a3a"
C_GOLD   = "d4a017"

st.markdown(f"""
<style>
    .main-header {{
        background: linear-gradient(135deg, {RE_NAVY} 0%, {RE_BLUE} 100%);
        padding: 2rem; border-radius: 8px; margin-bottom: 2rem; color: white;
    }}
    .main-header h1 {{ color: white !important; margin: 0; font-size: 2rem; }}
    .main-header p {{ color: rgba(255,255,255,0.85); margin: 0.5rem 0 0 0; }}
    .srf-badge {{
        background: {RE_RED}; color: white; padding: 0.3rem 0.8rem;
        border-radius: 4px; font-weight: bold; font-size: 0.85rem;
        display: inline-block; margin-left: 1rem;
    }}
    .stButton > button[kind="primary"] {{
        background: {RE_NAVY}; color: white; border: none;
    }}
    .stButton > button[kind="primary"]:hover {{ background: {RE_RED}; }}
    .mode-badge {{
        background: #f0f4f8; border-left: 4px solid {RE_NAVY};
        padding: 0.6rem 1rem; border-radius: 4px; margin-bottom: 1rem;
        color: {TEXT_DARK}; font-size: 0.9rem;
    }}
    .ccr-legend {{
        background: white; border: 1px solid #d0d0d0;
        border-radius: 6px; padding: 0.8rem 1rem;
        margin-top: 0.5rem;
    }}
    .ccr-legend-item {{
        display: inline-block; margin: 0.2rem 0.6rem 0.2rem 0;
        font-size: 0.88rem; color: {TEXT_DARK};
    }}
    .ccr-dot {{
        display: inline-block; width: 12px; height: 12px;
        border-radius: 50%; margin-right: 5px;
        border: 1px solid {TEXT_DARK};
        vertical-align: middle;
    }}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# CONSTANTES API
# ═══════════════════════════════════════════════════════════════════════
BASE_URL    = "https://ws01.cenace.gob.mx:8082/SWPML/SIM"
BLOQUE_MAX  = 7
FORMATO     = "JSON"
TIMEOUT     = 60

OSM_HEADERS = {
    'Accept': 'application/json',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'CENACE-PML-Analyzer/1.0 (Recurrent Energy MX, SRF)',
    'Referer': 'https://recurrent-energy.com/',
}


# ═══════════════════════════════════════════════════════════════════════
# CARGAR CATÁLOGO (cached)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data
def cargar_catalogo():
    catalog_path = os.path.join(os.path.dirname(__file__), 'data', 'catalogo_nodos.xlsx')
    if not os.path.exists(catalog_path):
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(catalog_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    catalogo = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 6 or not r[3]:
            continue
        clave = str(r[3]).strip()
        catalogo[clave] = {
            'sistema':   str(r[0]).strip() if r[0] else '',
            'ccr':       str(r[1]).strip() if r[1] else '',
            'zona':      str(r[2]).strip() if r[2] else '',
            'nombre':    str(r[4]).strip() if r[4] else '',
            'kv':        int(r[5]) if r[5] else 0,
            'estado':    str(r[15]).strip() if len(r) > 15 and r[15] else '',
            'municipio': str(r[17]).strip() if len(r) > 17 and r[17] else '',
        }
    wb.close()
    return catalogo


# ═══════════════════════════════════════════════════════════════════════
# OSM CACHE
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, ttl=86400 * 7)
def cargar_osm_subestaciones():
    osm_local = os.path.join(os.path.dirname(__file__), 'data', 'osm_subestaciones_mx.json')
    if os.path.exists(osm_local):
        with open(osm_local, 'r', encoding='utf-8') as f:
            return json.load(f)

    query = """
[out:json][timeout:180];
area["ISO3166-1"="MX"]->.mx;
(
  node["power"="substation"](area.mx);
  way["power"="substation"](area.mx);
);
out center tags;
"""
    try:
        r = requests.post('https://overpass-api.de/api/interpreter',
                         data={'data': query},
                         headers=OSM_HEADERS, timeout=200)
        if r.status_code != 200:
            return []
        data = r.json()
        elementos = data.get('elements', [])
    except Exception:
        return []

    def _parse_voltaje_max(v_raw):
        if not v_raw: return None
        v_clean = str(v_raw).replace('kV', '').replace('KV', '').replace('kv', '').strip()
        partes = [p.strip() for p in v_clean.split(';') if p.strip()]
        nums = []
        for p in partes:
            num_str = ''
            for ch in p:
                if ch.isdigit() or ch == '.':
                    num_str += ch
                else: break
            if num_str:
                try:
                    n = float(num_str)
                    nums.append(n / 1000 if n >= 1000 else n)
                except: pass
        return max(nums) if nums else None

    subs = []
    for el in elementos:
        tags = el.get('tags', {})
        sub_type = tags.get('substation', '')
        if sub_type in ('minor_distribution', 'traction', 'industrial',
                       'household', 'measurement'):
            continue
        if el['type'] == 'node':
            lat_e = el.get('lat', 0); lon_e = el.get('lon', 0)
        elif 'center' in el:
            lat_e = el['center'].get('lat', 0); lon_e = el['center'].get('lon', 0)
        else: continue
        if lat_e == 0 and lon_e == 0: continue
        v_kv = _parse_voltaje_max(tags.get('voltage', ''))
        subs.append({
            'osm_id': el['id'], 'osm_type': el['type'],
            'lat': round(lat_e, 6), 'lon': round(lon_e, 6),
            'name': tags.get('name', ''),
            'operator': tags.get('operator', ''),
            'voltage_kv': v_kv,
            'voltage_raw': tags.get('voltage', ''),
            'substation': sub_type,
            'rating': tags.get('rating', ''),
        })
    return subs


# ═══════════════════════════════════════════════════════════════════════
# MATCHING GEOGRÁFICO
# ═══════════════════════════════════════════════════════════════════════
BBOX_ESTADO = {
    'AGUASCALIENTES':       (21.62, -102.87, 22.45, -101.84),
    'BAJA CALIFORNIA':      (28.00, -117.30, 32.72, -112.75),
    'BAJA CALIFORNIA SUR':  (22.87, -115.30, 28.00, -109.40),
    'CAMPECHE':             (17.81, -92.47,  20.85, -89.10),
    'CHIAPAS':              (14.53, -94.14,  17.99, -90.37),
    'CHIHUAHUA':            (25.55, -109.07, 31.78, -103.30),
    'COAHUILA':             (24.55, -103.96, 29.88, -99.84),
    'COLIMA':               (18.69, -104.74, 19.52, -103.49),
    'DURANGO':              (22.36, -107.20, 26.83, -102.46),
    'GUANAJUATO':           (19.91, -102.10, 21.85, -99.69),
    'GUERRERO':             (16.31, -102.18, 18.92, -98.00),
    'HIDALGO':              (19.63, -99.91,  21.40, -97.97),
    'JALISCO':              (18.92, -105.69, 22.75, -101.50),
    'MEXICO':               (18.36, -100.59, 20.29, -98.62),
    'CIUDAD DE MEXICO':     (19.05, -99.36,  19.59, -98.94),
    'MICHOACAN':            (17.92, -103.74, 20.40, -100.07),
    'MORELOS':              (18.32, -99.50,  19.13, -98.62),
    'NAYARIT':              (20.60, -105.78, 23.09, -103.71),
    'NUEVO LEON':           (23.18, -101.21, 27.81, -98.40),
    'OAXACA':               (15.65, -98.51,  18.66, -93.86),
    'PUEBLA':               (17.86, -99.06,  20.84, -96.71),
    'QUERETARO':            (20.02, -100.59, 21.67, -99.04),
    'QUINTANA ROO':         (17.88, -89.30,  21.62, -86.69),
    'SAN LUIS POTOSI':      (21.16, -102.30, 24.55, -98.34),
    'SINALOA':              (22.50, -109.45, 27.04, -105.41),
    'SONORA':               (26.32, -115.07, 32.51, -108.40),
    'TABASCO':              (17.25, -94.13,  18.65, -91.00),
    'TAMAULIPAS':           (22.21, -100.16, 27.68, -97.13),
    'TLAXCALA':             (19.06, -98.71,  19.72, -97.62),
    'VERACRUZ':             (17.15, -98.65,  22.46, -93.61),
    'YUCATAN':              (19.55, -90.42,  21.62, -87.53),
    'ZACATECAS':            (21.04, -104.34, 25.13, -100.81),
}

PALABRAS_GENERICAS = {
    'maniobras', 'planta', 'entronque', 'subestacion', 'subestación',
    'central', 'switching', 'rectificadora', 'tap', 'industrial',
    'cementos', 'pemex', 'cuadro', 'parque', 'potencia', 'aeropuerto',
    'norte', 'sur', 'este', 'oeste', 'poniente', 'oriente',
    'centro', 'cerro', 'valle', 'laguna', 'puerto', 'villa',
    'cabo', 'monte', 'sierra', 'mesa', 'isla', 'rio',
    'nuevo', 'nueva', 'viejo', 'vieja', 'gran', 'grande',
    'antiguo', 'antigua', 'segundo', 'segunda', 'primero', 'primera',
    'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x',
    '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
    'uno', 'dos', 'tres', 'cuatro', 'cinco',
    'de', 'la', 'el', 'las', 'los', 'del', 'al', 'con', 'por',
    'mexico', 'santa', 'santo', 'san',
}

PALABRAS_SEMIGENERICAS = {
    'hermosillo', 'guadalajara', 'monterrey', 'merida', 'chihuahua',
    'culiacan', 'puebla', 'queretaro', 'leon', 'morelia', 'tepic',
    'durango', 'mazatlan', 'oaxaca', 'campeche', 'cancun',
    'chetumal', 'colima', 'manzanillo', 'veracruz', 'tampico',
    'reynosa', 'matamoros', 'juarez', 'mexicali', 'tijuana',
    'ensenada', 'cardenas', 'guadalupe', 'progreso', 'lazaro',
}


def en_estado(lat, lon, nombre_estado):
    if not nombre_estado: return None
    clave = nombre_estado.upper().strip()
    for variant in [clave, clave.split(' ')[0],
                    ' '.join(clave.split(' ')[:2]),
                    ' '.join(clave.split(' ')[:3])]:
        bb = BBOX_ESTADO.get(variant)
        if bb:
            return bb[0] <= lat <= bb[2] and bb[1] <= lon <= bb[3]
    return None


def normalizar_nombre(s):
    if not s: return ''
    s = s.lower().strip()
    repl = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n','ü':'u'}
    for k, v in repl.items():
        s = s.replace(k, v)
    for prefix in ['s/e ', 's/e. ', 'se ', 'sub ', 'subestacion ',
                   'central ', 'planta ', 'c.t. ', 'ct ',
                   'c.h. ', 'ch ', 'c.f.e. ', 'cfe ']:
        if s.startswith(prefix):
            s = s[len(prefix):]
    for art in ['los ', 'las ', 'el ', 'la ']:
        if s.startswith(art):
            resto = s[len(art):]
            if len(resto.split(' ', 1)[0]) >= 4:
                s = resto
                break
    for sufijo in [' maniobras', ' switching', ' tap',
                   ' planta', ' central', ' cfe']:
        if s.endswith(sufijo):
            s = s[:-len(sufijo)]
    roman_map = {' i': ' 1', ' ii': ' 2', ' iii': ' 3', ' iv': ' 4',
                 ' v': ' 5', ' vi': ' 6', ' vii': ' 7', ' viii': ' 8',
                 ' ix': ' 9', ' x': ' 10'}
    for r, a in roman_map.items():
        if s.endswith(r):
            s = s[:-len(r)] + a
            break
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def similitud(a, b):
    if not a or not b: return 0.0
    return SequenceMatcher(None, a, b).ratio()


def palabras_clave(nombre_norm):
    if not nombre_norm: return []
    return [p for p in nombre_norm.split()
            if len(p) >= 5 and p not in PALABRAS_GENERICAS and not p.isdigit()]


def buscar_match(nombre_cenace, voltaje_cenace_kv, estado_cat, candidatos):
    nombre_norm = normalizar_nombre(nombre_cenace)
    if not nombre_norm:
        return None, 0.0, 'sin nombre'

    exactos = [s for s in candidatos if s['name_norm'] == nombre_norm]
    if exactos:
        if estado_cat:
            exactos_en_est = [s for s in exactos
                              if en_estado(s['lat'], s['lon'], estado_cat) is not False]
            if exactos_en_est:
                exactos = exactos_en_est
            elif len(nombre_norm) < 18:
                exactos = []
        if exactos:
            if voltaje_cenace_kv > 0 and len(exactos) > 1:
                con_v = [s for s in exactos if s['voltage_kv'] and
                         abs(s['voltage_kv'] - voltaje_cenace_kv) < voltaje_cenace_kv * 0.15]
                if con_v:
                    return con_v[0], 1.0, 'match exacto + voltaje'
            return exactos[0], 1.0, 'match exacto'

    substring_matches = []
    for s in candidatos:
        n_osm = s['name_norm']
        if nombre_norm in n_osm and len(nombre_norm) >= 4:
            if n_osm.startswith(nombre_norm + ' ') or n_osm == nombre_norm:
                cob = len(nombre_norm) / len(n_osm)
                if cob >= 0.40:
                    substring_matches.append((0.92, s, 'OSM extiende CENACE'))
            elif (' ' + nombre_norm) in n_osm or n_osm.endswith(' ' + nombre_norm):
                cob = len(nombre_norm) / len(n_osm)
                if cob >= 0.40:
                    substring_matches.append((0.88, s, 'CENACE dentro de OSM'))
        elif n_osm in nombre_norm and len(n_osm) >= 4:
            if nombre_norm.startswith(n_osm + ' ') or nombre_norm == n_osm:
                cob = len(n_osm) / len(nombre_norm)
                if cob >= 0.40:
                    substring_matches.append((0.90, s, 'CENACE extiende OSM'))

    difusos = []
    for s in candidatos:
        sim = similitud(nombre_norm, s['name_norm'])
        if sim >= 0.80:
            difusos.append((sim, s))

    kw_cenace = palabras_clave(nombre_norm)
    keyword_matches = []
    if kw_cenace:
        for s in candidatos:
            kw_osm = palabras_clave(s.get('name_norm', ''))
            if not kw_osm: continue
            compartidas = set(kw_cenace) & set(kw_osm)
            if not compartidas: continue
            todas_semi = all(p in PALABRAS_SEMIGENERICAS for p in compartidas)
            largas_no_semi = sum(1 for p in compartidas
                                 if len(p) >= 6 and p not in PALABRAS_SEMIGENERICAS)
            if todas_semi: base_score = 0.62
            elif largas_no_semi >= 1: base_score = 0.78
            else: base_score = 0.70
            ratio = len(compartidas) / max((len(kw_cenace) + len(kw_osm)) / 2, 1)
            score = base_score + (ratio * 0.12)
            keyword_matches.append((score, s, list(compartidas)))

    scored = []; seen = set()
    for sc, s, t in substring_matches:
        if s['osm_id'] not in seen:
            scored.append((sc, s, 'substring', t))
            seen.add(s['osm_id'])
    for sim, s in difusos:
        if s['osm_id'] not in seen:
            scored.append((sim, s, 'difuso', ''))
            seen.add(s['osm_id'])
    for sc, s, kws in keyword_matches:
        if s['osm_id'] not in seen:
            scored.append((sc, s, 'keyword', f'comparte "{",".join(kws[:2])}"'))
            seen.add(s['osm_id'])

    if not scored:
        return None, 0.0, 'no match'

    if estado_cat:
        scored = [(sc, s, t, st) for sc, s, t, st in scored
                  if en_estado(s['lat'], s['lon'], estado_cat) is not False]

    if not scored:
        return None, 0.0, 'sin match (estado)'

    if voltaje_cenace_kv > 0:
        filt = []
        for sim, s, t, st in scored:
            v = s['voltage_kv']
            ok = (v is not None and abs(v - voltaje_cenace_kv) < voltaje_cenace_kv * 0.15)
            unk = v is None
            if t == 'difuso' and sim < 0.92 and not (ok or unk):
                continue
            if t == 'keyword' and not (ok or unk):
                continue
            if ok: filt.append((sim + 0.02, s, t, st))
            elif unk: filt.append((sim, s, t, st))
            else: filt.append((sim - 0.10, s, t, st))
        scored = filt

    if not scored:
        return None, 0.0, 'sin match (voltaje)'

    scored.sort(key=lambda x: -x[0])
    sim_f, mejor, tipo, sub_tipo = scored[0]
    if tipo == 'substring': razon = sub_tipo
    elif tipo == 'keyword': razon = f'palabra-clave: {sub_tipo}'
    elif sim_f >= 0.95: razon = 'match difuso alto'
    elif sim_f >= 0.85: razon = 'match difuso medio'
    else: razon = f'match difuso ({sim_f:.2f})'
    return mejor, sim_f, razon


def matchear_nodos(nodos, catalogo, osm_subs):
    for s in osm_subs:
        if 'name_norm' not in s:
            s['name_norm'] = normalizar_nombre(s.get('name', ''))
    candidatos = [s for s in osm_subs if s.get('name_norm')]
    resultados = []
    for clave in nodos:
        info = catalogo.get(clave, {})
        nombre = info.get('nombre', clave)
        kv = info.get('kv', 0)
        if kv == 0:
            try: kv = int(clave.split('-')[-1])
            except: kv = 0
        ccr = info.get('ccr', '')
        zona = info.get('zona', '')
        estado = info.get('estado', '')
        municipio = info.get('municipio', '')

        match, sim, razon = buscar_match(nombre, kv, estado, candidatos)

        if match:
            if 'palabra-clave' in razon: calidad = '🥉 Aceptable'
            elif sim >= 0.95: calidad = '🥇 Excelente'
            elif sim >= 0.90: calidad = '🥈 Bueno'
            else: calidad = '🥉 Aceptable'
        else:
            calidad = '❌ Sin match'

        resultados.append({
            'clave': clave, 'nombre_cenace': nombre, 'ccr': ccr, 'zona': zona,
            'kv_cenace': kv, 'estado': estado, 'municipio': municipio,
            'calidad': calidad,
            'similitud': round(sim, 3) if match else 0,
            'razon': razon,
            'nombre_osm': match['name'] if match else '',
            'kv_osm': match['voltage_kv'] if match and match['voltage_kv'] else None,
            'voltage_raw': match['voltage_raw'] if match else '',
            'operator_osm': match['operator'] if match else '',
            'subtipo_osm': match['substation'] if match else '',
            'lat': match['lat'] if match else None,
            'lon': match['lon'] if match else None,
            'osm_id': match['osm_id'] if match else None,
            'osm_type': match['osm_type'] if match else None,
        })
    return resultados


# ═══════════════════════════════════════════════════════════════════════
# CENACE — descarga
# ═══════════════════════════════════════════════════════════════════════
def parse_fecha(s): return datetime.strptime(s, "%Y/%m/%d")
def fmt(d): return d.strftime("%Y/%m/%d")

def generar_bloques(ini_str, fin_str, max_dias=BLOQUE_MAX):
    ini = parse_fecha(ini_str); fin = parse_fecha(fin_str)
    bloques = []; cursor = ini
    while cursor <= fin:
        bloque_fin = min(cursor + timedelta(days=max_dias - 1), fin)
        bloques.append((fmt(cursor), fmt(bloque_fin)))
        cursor = bloque_fin + timedelta(days=1)
    return bloques

def construir_url(nodos_lista, fecha_ini, fecha_fin, sistema, proceso):
    nodos_str = ",".join(nodos_lista)
    ai, mi, di = fecha_ini.split("/")
    af, mf, df_ = fecha_fin.split("/")
    return f"{BASE_URL}/{sistema}/{proceso}/{nodos_str}/{ai}/{mi}/{di}/{af}/{mf}/{df_}/{FORMATO}"

def consultar(nodos_lista, fecha_ini, fecha_fin, sistema, proceso,
              errores_lista, lock, max_reintentos=4):
    url = construir_url(nodos_lista, fecha_ini, fecha_fin, sistema, proceso)
    if len(url) > 2000:
        with lock:
            errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                   "nodos": len(nodos_lista),
                                   "error": "URL muy larga"})
        return None
    session = requests.Session(); session.verify = False
    for intento in range(max_reintentos):
        try:
            r = session.get(url, timeout=TIMEOUT)
            if r.status_code == 200: return r.text
            elif r.status_code == 204: return None
            elif r.status_code == 429:
                time.sleep(2 ** (intento + 2)); continue
            else:
                if intento < max_reintentos - 1:
                    time.sleep(2 ** (intento + 1)); continue
                with lock:
                    errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                           "nodos": len(nodos_lista),
                                           "error": f"HTTP {r.status_code}"})
                return None
        except requests.exceptions.Timeout:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento); continue
            with lock:
                errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                       "nodos": len(nodos_lista),
                                       "error": "Timeout"})
            return None
        except Exception as e:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento); continue
            with lock:
                errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                       "nodos": len(nodos_lista),
                                       "error": f"{type(e).__name__}"})
            return None
    return None

def _num(val):
    try: return float(val)
    except (TypeError, ValueError): return val

def parsear_json(texto):
    if not texto: return {}
    try: obj = json.loads(texto)
    except Exception: return {}
    datos = {}
    for nd in obj.get("Resultados", []):
        clv = nd.get("clv_nodo", "?")
        valores = nd.get("Valores", [])
        if not valores: continue
        registros = []
        for v in valores:
            registros.append({
                "fecha":   v.get("fecha", ""),
                "hora":    int(v.get("hora", 0)) if v.get("hora") else 0,
                "pml":     _num(v.get("pml", 0)),
                "pml_ene": _num(v.get("pml_ene", 0)),
                "pml_per": _num(v.get("pml_per", 0)),
                "pml_cng": _num(v.get("pml_cng", 0)),
            })
        datos[clv] = registros
    return datos

def descargar_pml(nodos, fecha_ini, fecha_fin, sistema, proceso, progress_cb=None):
    bloques = generar_bloques(fecha_ini, fecha_fin)
    LOTE = 10
    lotes_nodos = [nodos[i:i+LOTE] for i in range(0, len(nodos), LOTE)]
    consultas = []
    for lote in lotes_nodos:
        for bi, bf in bloques:
            consultas.append((lote, bi, bf))
    total = len(consultas)
    acumulado = {n: [] for n in nodos}
    errores_consulta = []; lock = Lock(); completed = 0
    def _job(c):
        lote, bi, bf = c
        return parsear_json(consultar(lote, bi, bf, sistema, proceso,
                                       errores_consulta, lock))
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_job, c): c for c in consultas}
        for fut in as_completed(futures):
            data = fut.result() or {}
            for nodo, regs in data.items():
                if nodo in acumulado:
                    acumulado[nodo].extend(regs)
            completed += 1
            if progress_cb:
                progress_cb(completed, total)
    acumulado = {n: regs for n, regs in acumulado.items() if regs}
    return acumulado, errores_consulta


# ═══════════════════════════════════════════════════════════════════════
# EXCEL DATOS (datos crudos)
# ═══════════════════════════════════════════════════════════════════════
def generar_excel_datos(acumulado, sistema, proceso, fecha_ini, fecha_fin):
    _NOW = datetime.now().strftime("%Y-%m-%d %H:%M")
    _side = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=_side, right=_side, top=_side, bottom=_side)

    def hdr(cell, bg=C_HEADER, fg=C_WHITE, size=11):
        cell.font = Font(bold=True, color=fg, size=size, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE

    _FONT_DATO = Font(name="Arial", size=10)
    _ALIGN_NUM = Alignment(horizontal="center", vertical="center")
    _ALIGN_TXT = Alignment(horizontal="left", vertical="center")
    _FILL_ALT  = PatternFill("solid", start_color=C_ALT)

    COLS   = ["Fecha", "Hora", "PML ($/MWh)", "Energía", "Pérdidas", "Congestión"]
    ANCHOS = [14, 8, 16, 16, 16, 18]
    ES_NUM = [False, True, True, True, True, True]

    wb = Workbook()
    ws = wb.active
    ws.title = "■ Portada"
    ws.sheet_properties.tabColor = C_HEADER
    ws.sheet_view.showGridLines = False

    ws.merge_cells("B2:H4")
    c = ws["B2"]
    c.value = "Precios Marginales Locales\nMercado Eléctrico Mayorista — CENACE"
    c.font = Font(bold=True, color=C_WHITE, size=22, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 5):
        ws.row_dimensions[r].height = 32

    ws.merge_cells("B5:H5")
    c = ws["B5"]; c.value = "Reporte de Datos"
    c.font = Font(bold=True, italic=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_RED)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("B7:H7")
    c = ws["B7"]; c.value = "PARÁMETROS DE LA CONSULTA"
    c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 22

    info_rows = [
        ("Sistema", sistema), ("Proceso", proceso),
        ("Fecha inicial", fecha_ini), ("Fecha final", fecha_fin),
        ("Nodos en reporte", f"{len(acumulado):,} nodos"),
        ("Total registros", f"{sum(len(f) for f in acumulado.values()):,} filas"),
        ("Fecha generación", _NOW),
    ]
    for ri, (lbl, val) in enumerate(info_rows, start=8):
        ws.merge_cells(f"B{ri}:D{ri}")
        c1 = ws[f"B{ri}"]; c1.value = lbl
        c1.font = Font(bold=True, color=C_HEADER, size=11, name="Arial")
        c1.fill = PatternFill("solid", start_color=C_INFO)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c1.border = BORDE
        ws.merge_cells(f"E{ri}:H{ri}")
        c2 = ws[f"E{ri}"]; c2.value = val
        c2.font = Font(size=11, name="Arial")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2.border = BORDE
        ws.row_dimensions[ri].height = 22

    ws.merge_cells("B17:H19")
    c = ws["B17"]; c.value = "SRF"
    c.font = Font(bold=True, color=C_HEADER, size=72, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(17, 20):
        ws.row_dimensions[r].height = 30

    ws.merge_cells("B21:H21")
    c = ws["B21"]; c.value = "Prepared by: Sebastian Roldan · Recurrent Energy"
    c.font = Font(italic=True, bold=True, color=C_HEADER, size=12, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=C_INFO)
    ws.row_dimensions[21].height = 22

    ws.merge_cells("B22:H22")
    c = ws["B22"]; c.value = "A subsidiary of Canadian Solar"
    c.font = Font(italic=True, color="555555", size=10, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[22].height = 18

    for col, w in [("A", 3), ("B", 14), ("C", 14), ("D", 14),
                   ("E", 14), ("F", 14), ("G", 14), ("H", 14), ("I", 3)]:
        ws.column_dimensions[col].width = w

    ws_res = wb.create_sheet("Resumen")
    ws_res.merge_cells("A1:F1")
    c = ws_res["A1"]; c.value = f"Resumen — {len(acumulado)} nodos · {sistema} · {proceso}"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 28

    res_cols = ["Nodo", "# Registros", "PML promedio", "PML máximo", "PML mínimo", "Período"]
    for ci, col in enumerate(res_cols, 1):
        hdr(ws_res.cell(row=2, column=ci, value=col), bg=C_SUB)
    ws_res.row_dimensions[2].height = 22

    for ri, (nodo, filas) in enumerate(acumulado.items(), start=3):
        if not filas: continue
        pmls = [f["pml"] for f in filas if isinstance(f["pml"], (int, float))]
        bg = C_ALT if ri % 2 == 0 else None
        valores = [
            nodo, len(filas),
            f"{sum(pmls)/len(pmls):.2f}" if pmls else "—",
            f"{max(pmls):.2f}" if pmls else "—",
            f"{min(pmls):.2f}" if pmls else "—",
            f"{filas[0]['fecha']} → {filas[-1]['fecha']}",
        ]
        for ci, v in enumerate(valores, 1):
            cc = ws_res.cell(row=ri, column=ci, value=v)
            cc.font = _FONT_DATO
            cc.alignment = _ALIGN_NUM if ci > 1 else _ALIGN_TXT
            cc.border = BORDE
            if bg: cc.fill = PatternFill("solid", start_color=bg)

    for ci, w in enumerate([14, 14, 14, 14, 14, 24], 1):
        ws_res.column_dimensions[get_column_letter(ci)].width = w
    ws_res.freeze_panes = "A3"

    for nodo, filas in acumulado.items():
        ws_n = wb.create_sheet(title=nodo[:31])
        ws_n.merge_cells("A1:F1")
        c = ws_n["A1"]; c.value = f"PML — Nodo: {nodo}"
        c.font = Font(bold=True, color=C_WHITE, size=13, name="Arial")
        c.fill = PatternFill("solid", start_color=C_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_n.row_dimensions[1].height = 28

        ws_n.merge_cells("A2:F2")
        c = ws_n["A2"]
        c.value = (f"Sistema: {sistema} | Proceso: {proceso} | "
                   f"Período: {fecha_ini} → {fecha_fin} | Total: {len(filas):,} | "
                   f"Sebastian Roldan (SRF) · Recurrent Energy")
        c.font = Font(italic=True, size=9, name="Arial", color="555555")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", start_color=C_INFO)
        ws_n.row_dimensions[2].height = 16

        for ci, (enc, ancho) in enumerate(zip(COLS, ANCHOS), 1):
            hdr(ws_n.cell(row=3, column=ci, value=enc), bg=C_SUB)
            ws_n.column_dimensions[get_column_letter(ci)].width = ancho
        ws_n.row_dimensions[3].height = 22

        for fila in filas:
            ws_n.append([fila["fecha"], fila["hora"], fila["pml"],
                         fila["pml_ene"], fila["pml_per"], fila["pml_cng"]])
        for i in range(len(filas)):
            row_idx = i + 4
            es_alt = (i % 2 == 0)
            for ci in range(1, 7):
                cc = ws_n.cell(row=row_idx, column=ci)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ES_NUM[ci-1] else _ALIGN_TXT
                if es_alt: cc.fill = _FILL_ALT
                if ES_NUM[ci-1]: cc.number_format = "#,##0.00"
        ws_n.freeze_panes = "A4"

    buffer = io.BytesIO()
    wb.save(buffer); buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# EXCEL ANÁLISIS (BESS Scoring + métricas)
# ═══════════════════════════════════════════════════════════════════════
def generar_excel_analisis(df_metricas, df_resumen, sistema, proceso, fecha_ini, fecha_fin):
    """Excel con BESS scoring para los 3 casos de uso + métricas."""
    _NOW = datetime.now().strftime("%Y-%m-%d %H:%M")
    _side = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=_side, right=_side, top=_side, bottom=_side)

    def hdr(cell, bg=C_HEADER, fg=C_WHITE, size=11):
        cell.font = Font(bold=True, color=fg, size=size, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE

    _FONT_DATO = Font(name="Arial", size=10)
    _ALIGN_NUM = Alignment(horizontal="center", vertical="center")
    _ALIGN_TXT = Alignment(horizontal="left", vertical="center")
    _FILL_ALT  = PatternFill("solid", start_color=C_ALT)

    wb = Workbook()
    ws = wb.active
    ws.title = "■ Portada"
    ws.sheet_properties.tabColor = C_RED
    ws.sheet_view.showGridLines = False

    ws.merge_cells("B2:H4")
    c = ws["B2"]
    c.value = "Análisis BESS — PML CENACE\nScoring por caso de uso"
    c.font = Font(bold=True, color=C_WHITE, size=22, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 5):
        ws.row_dimensions[r].height = 32

    ws.merge_cells("B5:H5")
    c = ws["B5"]; c.value = "Reporte de Análisis — Battery Energy Storage System"
    c.font = Font(bold=True, italic=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_RED)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("B7:H7")
    c = ws["B7"]; c.value = "PARÁMETROS"
    c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 22

    info_rows = [
        ("Sistema", sistema), ("Proceso", proceso),
        ("Período", f"{fecha_ini} → {fecha_fin}"),
        ("Nodos analizados", f"{len(df_metricas)} nodos"),
        ("Casos de uso", "Arbitraje · SSAA · Renewables Firming"),
        ("Metodología scoring", "Rank-percentile (0=peor, 100=mejor)"),
        ("Fecha generación", _NOW),
    ]
    for ri, (lbl, val) in enumerate(info_rows, start=8):
        ws.merge_cells(f"B{ri}:D{ri}")
        c1 = ws[f"B{ri}"]; c1.value = lbl
        c1.font = Font(bold=True, color=C_HEADER, size=11, name="Arial")
        c1.fill = PatternFill("solid", start_color=C_INFO)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c1.border = BORDE
        ws.merge_cells(f"E{ri}:H{ri}")
        c2 = ws[f"E{ri}"]; c2.value = val
        c2.font = Font(size=11, name="Arial")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2.border = BORDE
        ws.row_dimensions[ri].height = 22

    ws.merge_cells("B17:H19")
    c = ws["B17"]; c.value = "SRF"
    c.font = Font(bold=True, color=C_HEADER, size=72, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(17, 20):
        ws.row_dimensions[r].height = 30

    ws.merge_cells("B21:H21")
    c = ws["B21"]; c.value = "Prepared by: Sebastian Roldan · Recurrent Energy"
    c.font = Font(italic=True, bold=True, color=C_HEADER, size=12, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=C_INFO)
    ws.row_dimensions[21].height = 22

    ws.merge_cells("B22:H22")
    c = ws["B22"]; c.value = "A subsidiary of Canadian Solar"
    c.font = Font(italic=True, color="555555", size=10, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[22].height = 18

    for col, w in [("A", 3), ("B", 16), ("C", 16), ("D", 16),
                   ("E", 16), ("F", 16), ("G", 16), ("H", 16), ("I", 3)]:
        ws.column_dimensions[col].width = w

    # Hoja Top 5 por caso de uso
    ws_top = wb.create_sheet("🏆 Top 5 por Caso de Uso")
    ws_top.merge_cells("A1:F1")
    c = ws_top["A1"]; c.value = "Top 5 nodos recomendados por caso de uso BESS"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_top.row_dimensions[1].height = 28

    cur_row = 3
    for use_case in ['Arbitraje', 'Servicios Auxiliares', 'Renewables Firming']:
        df_score = calcular_score_bess(df_metricas, use_case)
        if df_score.empty: continue

        # Header del caso
        ws_top.merge_cells(f"A{cur_row}:F{cur_row}")
        c = ws_top.cell(row=cur_row, column=1)
        c.value = f"⚡ {use_case}"
        c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
        c.fill = PatternFill("solid", start_color=C_RED)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_top.row_dimensions[cur_row].height = 22
        cur_row += 1

        # Subheader
        cols_top = ["Rank", "Nodo", "Score", "PML promedio", "Volatilidad", "Spread P95-P5"]
        for ci, col in enumerate(cols_top, 1):
            hdr(ws_top.cell(row=cur_row, column=ci, value=col), bg=C_SUB)
        ws_top.row_dimensions[cur_row].height = 20
        cur_row += 1

        for i, (_, r) in enumerate(df_score.head(5).iterrows(), 1):
            bg = C_ALT if i % 2 == 0 else None
            medal = ['🥇', '🥈', '🥉', '4', '5'][i-1]
            valores = [medal, r['nodo'], f"{r['score']:.1f}",
                       f"${r['pml_promedio']:.2f}",
                       f"${r['volatilidad']:.2f}",
                       f"${r['spread_p95_p5']:.2f}"]
            for ci, v in enumerate(valores, 1):
                cc = ws_top.cell(row=cur_row, column=ci, value=v)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ci != 2 else _ALIGN_TXT
                if bg: cc.fill = PatternFill("solid", start_color=bg)
            cur_row += 1
        cur_row += 1

    for ci, w in enumerate([8, 16, 14, 16, 16, 18], 1):
        ws_top.column_dimensions[get_column_letter(ci)].width = w
    ws_top.freeze_panes = "A2"

    # Hoja BESS scoring completo (3 casos de uso)
    for use_case in ['Arbitraje', 'Servicios Auxiliares', 'Renewables Firming']:
        df_score = calcular_score_bess(df_metricas, use_case)
        if df_score.empty: continue

        sheet_name = f"BESS {use_case[:22]}"
        ws_bess = wb.create_sheet(title=sheet_name[:31])

        ws_bess.merge_cells("A1:J1")
        c = ws_bess["A1"]; c.value = f"BESS Scoring — {use_case}"
        c.font = Font(bold=True, color=C_WHITE, size=13, name="Arial")
        c.fill = PatternFill("solid", start_color=C_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_bess.row_dimensions[1].height = 26

        ws_bess.merge_cells("A2:J2")
        c = ws_bess["A2"]
        c.value = DESCRIPCIONES_USE_CASE_PLAIN.get(use_case, "")
        c.font = Font(italic=True, size=10, name="Arial", color="555555")
        c.fill = PatternFill("solid", start_color=C_INFO)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_bess.row_dimensions[2].height = 20

        cols = ["Rank", "Nodo", "Score (0-100)", "PML promedio",
                "Volatilidad", "Spread P95-P5", "Spread día prom",
                "Cambios bruscos", "Horas pico", "% horas neg"]
        for ci, col in enumerate(cols, 1):
            hdr(ws_bess.cell(row=3, column=ci, value=col), bg=C_SUB)
        ws_bess.row_dimensions[3].height = 32

        for i, (_, r) in enumerate(df_score.iterrows(), 1):
            bg = C_ALT if i % 2 == 0 else None
            valores = [i, r['nodo'], r['score'], r['pml_promedio'],
                       r['volatilidad'], r['spread_p95_p5'],
                       r['spread_avg_diario'], r['cambios_bruscos'],
                       r['horas_pico'], r['pct_horas_neg']]
            for ci, v in enumerate(valores, 1):
                cc = ws_bess.cell(row=i+3, column=ci, value=v)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ci != 2 else _ALIGN_TXT
                if bg: cc.fill = PatternFill("solid", start_color=bg)
                if ci in (3, 4, 5, 6, 7): cc.number_format = "#,##0.00"
                elif ci == 10: cc.number_format = "0.0\"%\""

        for ci, w in enumerate([8, 16, 14, 14, 14, 14, 14, 14, 12, 14], 1):
            ws_bess.column_dimensions[get_column_letter(ci)].width = w
        ws_bess.freeze_panes = "A4"

    # Hoja Resumen (estadísticas básicas)
    ws_res = wb.create_sheet("📊 Resumen Estadístico")
    ws_res.merge_cells("A1:K1")
    c = ws_res["A1"]; c.value = f"Resumen estadístico — {len(df_resumen)} nodos"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 26

    cols = ["Nodo", "Nombre", "CCR", "# Reg", "Promedio", "Mediana",
            "Máximo", "Mínimo", "Volatilidad", "P95", "% horas neg"]
    for ci, col in enumerate(cols, 1):
        hdr(ws_res.cell(row=2, column=ci, value=col), bg=C_SUB)
    ws_res.row_dimensions[2].height = 22

    for i, (_, r) in enumerate(df_resumen.iterrows(), 1):
        bg = C_ALT if i % 2 == 0 else None
        valores = [r['nodo'], r['nombre'], r['ccr'], r['registros'],
                   r['promedio'], r['mediana'], r['maximo'], r['minimo'],
                   r['std'], r['p95'], r['% horas neg']]
        for ci, v in enumerate(valores, 1):
            cc = ws_res.cell(row=i+2, column=ci, value=v)
            cc.font = _FONT_DATO
            cc.border = BORDE
            cc.alignment = _ALIGN_NUM if ci > 3 else _ALIGN_TXT
            if bg: cc.fill = PatternFill("solid", start_color=bg)
            if ci in (5, 6, 7, 8, 9, 10): cc.number_format = "#,##0.00"
            elif ci == 11: cc.number_format = "0.0\"%\""

    for ci, w in enumerate([14, 22, 14, 10, 12, 12, 12, 12, 12, 12, 12], 1):
        ws_res.column_dimensions[get_column_letter(ci)].width = w
    ws_res.freeze_panes = "A3"

    # Hoja Métricas Raw (todas las métricas calculadas)
    ws_m = wb.create_sheet("🔢 Métricas Calculadas")
    ws_m.merge_cells("A1:I1")
    c = ws_m["A1"]; c.value = "Métricas BESS calculadas (datos raw)"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_m.row_dimensions[1].height = 26

    cols_m = ["Nodo", "PML promedio", "Volatilidad", "Spread P95-P5",
              "Spread día prom", "Spread día/noche", "Cambios bruscos",
              "Horas pico", "% horas neg"]
    for ci, col in enumerate(cols_m, 1):
        hdr(ws_m.cell(row=2, column=ci, value=col), bg=C_SUB)
    ws_m.row_dimensions[2].height = 32

    for i, (_, r) in enumerate(df_metricas.iterrows(), 1):
        bg = C_ALT if i % 2 == 0 else None
        valores = [r['nodo'], r['pml_promedio'], r['volatilidad'],
                   r['spread_p95_p5'], r['spread_avg_diario'], r['spread_dia'],
                   r['cambios_bruscos'], r['horas_pico'], r['pct_horas_neg']]
        for ci, v in enumerate(valores, 1):
            cc = ws_m.cell(row=i+2, column=ci, value=v)
            cc.font = _FONT_DATO
            cc.border = BORDE
            cc.alignment = _ALIGN_NUM if ci != 1 else _ALIGN_TXT
            if bg: cc.fill = PatternFill("solid", start_color=bg)
            if ci in (2, 3, 4, 5, 6): cc.number_format = "#,##0.00"
            elif ci == 9: cc.number_format = "0.0\"%\""

    for ci, w in enumerate([14, 14, 14, 14, 14, 14, 14, 12, 12], 1):
        ws_m.column_dimensions[get_column_letter(ci)].width = w
    ws_m.freeze_panes = "A3"

    buffer = io.BytesIO()
    wb.save(buffer); buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# KMZ
# ═══════════════════════════════════════════════════════════════════════
def generar_kmz(matches_df):
    KML_NS = 'http://www.opengis.net/kml/2.2'
    def _se(parent, tag, text=None, **attrib):
        el = ET.SubElement(parent, tag, attrib)
        if text is not None: el.text = str(text)
        return el

    root = ET.Element('kml', xmlns=KML_NS)
    doc = _se(root, 'Document')
    _se(doc, 'name', f'PML CENACE Geo — {datetime.now().strftime("%Y-%m-%d")} · SRF')
    _se(doc, 'open', '1')

    def _style(parent, sid, color_abgr, scale=1.2):
        st_el = _se(parent, 'Style', id=sid)
        ic = _se(st_el, 'IconStyle')
        _se(ic, 'color', color_abgr); _se(ic, 'scale', str(scale))
        ico = _se(ic, 'Icon')
        _se(ico, 'href', 'http://maps.google.com/mapfiles/kml/shapes/square.png')
        lbl = _se(st_el, 'LabelStyle'); _se(lbl, 'scale', '0.85')
        return st_el

    _style(doc, 'styleExc',     'ff00aa00', 1.3)
    _style(doc, 'styleBueno',   'ff00bbff', 1.2)
    _style(doc, 'styleAcept',   'ff0088ff', 1.1)

    folder_main = _se(doc, 'Folder')
    _se(folder_main, 'name', '⚡ Nodos CENACE')
    _se(folder_main, 'open', '1')

    matches_df_ok = matches_df[matches_df['lat'].notna()].copy()
    regiones_count = {}
    for _, row in matches_df_ok.iterrows():
        ccr = str(row['ccr']) or 'SIN CLASIFICAR'
        regiones_count[ccr] = regiones_count.get(ccr, 0) + 1

    folders = {}
    for ccr in sorted(regiones_count, key=lambda r: -regiones_count[r]):
        f = _se(folder_main, 'Folder')
        _se(f, 'name', f'📍 {ccr} ({regiones_count[ccr]})')
        _se(f, 'open', '0')
        folders[ccr] = {}
        for cal in ['🥇 Excelente', '🥈 Bueno', '🥉 Aceptable']:
            sf = _se(f, 'Folder')
            _se(sf, 'name', cal)
            _se(sf, 'open', '1' if 'Excelente' in cal else '0')
            folders[ccr][cal] = sf

    for _, row in matches_df_ok.iterrows():
        cal = str(row['calidad'])
        ccr = str(row['ccr']) or 'SIN CLASIFICAR'
        folder = folders.get(ccr, {}).get(cal, folder_main)
        if cal.startswith('🥇'):    style = '#styleExc'
        elif cal.startswith('🥈'):  style = '#styleBueno'
        else:                      style = '#styleAcept'

        pm = _se(folder, 'Placemark')
        _se(pm, 'name', f"{row['clave']} · {str(row['nombre_osm'])[:25]}")
        _se(pm, 'styleUrl', style)
        desc = (
            f"<b>Clave CENACE:</b> {row['clave']}<br/>"
            f"<b>Nombre CENACE:</b> {row['nombre_cenace']}<br/>"
            f"<b>CCR / Zona:</b> {row['ccr']} · {row['zona']}<br/>"
            f"<b>Voltaje CENACE:</b> {row['kv_cenace']} kV<br/>"
            f"<b>Estado / Municipio:</b> {row['estado']} · {row['municipio']}<br/><br/>"
            f"<b>━━━ Match OSM ━━━</b><br/>"
            f"<b>Calidad:</b> {row['calidad']} (sim {row['similitud']})<br/>"
            f"<b>Nombre OSM:</b> {row['nombre_osm']}<br/>"
            f"<b>Voltaje OSM:</b> {row['kv_osm']} kV<br/>"
            f"<b>Operador:</b> {row['operator_osm']}<br/>"
            f"<i>SRF · Recurrent Energy</i>"
        )
        _se(pm, 'description', desc)
        pt = _se(pm, 'Point')
        _se(pt, 'coordinates', f"{float(row['lon'])},{float(row['lat'])},0")

    kml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding='unicode')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('doc.kml', kml_str.encode('utf-8'))
    buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# DATAFRAMES (cached para performance)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, max_entries=3)
def acumulado_a_dataframe_cached(acumulado_id, _acumulado, _catalogo):
    """Convierte acumulado a DF. acumulado_id es un hash para cache key."""
    rows = []
    for nodo, filas in _acumulado.items():
        info = _catalogo.get(nodo, {}) if _catalogo else {}
        for f in filas:
            try: pml_val = float(f["pml"])
            except (TypeError, ValueError): continue
            rows.append({
                "nodo":   nodo,
                "ccr":    info.get("ccr", "?"),
                "nombre": info.get("nombre", nodo),
                "estado": info.get("estado", ""),
                "fecha":  f["fecha"],
                "hora":   int(f["hora"]) if f["hora"] else 0,
                "pml":    pml_val,
            })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["fecha_dt"] = pd.to_datetime(df["fecha"], errors="coerce")
        df["mes"] = df["fecha_dt"].dt.month
        df["mes_nombre"] = df["fecha_dt"].dt.strftime("%b %Y")
    return df


def acumulado_a_dataframe(acumulado, catalogo):
    """Wrapper con hash key para cache."""
    if not acumulado: return pd.DataFrame()
    # Hash key: # nodos + total filas + primer nodo
    key = (len(acumulado),
           sum(len(v) for v in acumulado.values()),
           list(acumulado.keys())[0] if acumulado else "")
    return acumulado_a_dataframe_cached(str(key), acumulado, catalogo)


@st.cache_data(show_spinner=False, max_entries=3)
def calcular_resumen_cached(df_hash, _df):
    if _df.empty: return pd.DataFrame()
    summary = _df.groupby("nodo").agg(
        nombre=("nombre", "first"),
        ccr=("ccr", "first"),
        registros=("pml", "count"),
        promedio=("pml", "mean"),
        mediana=("pml", "median"),
        maximo=("pml", "max"),
        minimo=("pml", "min"),
        std=("pml", "std"),
    ).round(2)
    pct_neg = _df.groupby("nodo")["pml"].apply(
        lambda x: (x < 0).sum() / len(x) * 100 if len(x) > 0 else 0
    ).round(1)
    summary["% horas neg"] = pct_neg
    summary["p95"] = _df.groupby("nodo")["pml"].quantile(0.95).round(2)
    summary["p5"] = _df.groupby("nodo")["pml"].quantile(0.05).round(2)
    return summary.reset_index().sort_values("promedio", ascending=False).reset_index(drop=True)


def calcular_resumen(df):
    if df.empty: return pd.DataFrame()
    key = f"{len(df)}-{df['pml'].sum():.0f}"
    return calcular_resumen_cached(key, df)


# Layout estándar para gráficas
def _layout_estandar(titulo, height=480):
    return dict(
        title=dict(
            text=titulo,
            font=dict(family="Arial Black", size=16, color=TEXT_TITLE),
            x=0.02, xanchor="left",
        ),
        height=height,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Arial", size=12, color=TEXT_DARK),
        margin=dict(l=70, r=40, t=70, b=60),
        hoverlabel=dict(bgcolor="white", font_size=12, font_color=TEXT_DARK,
                        bordercolor=AXIS_LINE),
        uirevision='constant',  # mantiene zoom/pan al actualizar
    )


def _ejes_estandar(fig, x_title=None, y_title=None):
    fig.update_xaxes(
        showgrid=True, gridcolor=GRID_LIGHT,
        linecolor=AXIS_LINE, linewidth=1.5,
        tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
        title=dict(text=x_title or "", font=dict(family="Arial", size=13, color=TEXT_DARK)),
    )
    fig.update_yaxes(
        showgrid=True, gridcolor=GRID_LIGHT,
        linecolor=AXIS_LINE, linewidth=1.5,
        tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
        title=dict(text=y_title or "", font=dict(family="Arial", size=13, color=TEXT_DARK)),
        zeroline=True, zerolinecolor=RE_RED, zerolinewidth=1.5,
    )
    return fig


def grafica_lineas_tiempo(df, max_nodos=15):
    if df.empty: return None
    promedios = df.groupby("nodo")["pml"].mean().sort_values(ascending=False)
    nodos_mostrar = promedios.head(max_nodos).index.tolist()
    df_plot = df[df["nodo"].isin(nodos_mostrar)].copy()
    daily = df_plot.groupby(["nodo", "fecha_dt"])["pml"].mean().reset_index()
    fig = px.line(
        daily, x="fecha_dt", y="pml", color="nodo",
        color_discrete_sequence=PALETTE,
    )
    fig.update_traces(line=dict(width=2.5), mode='lines')
    fig.update_layout(**_layout_estandar(
        f"Evolución temporal · Top {len(nodos_mostrar)} nodos por PML promedio"))
    fig.update_layout(
        hovermode="x unified",
        legend=dict(
            orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor=AXIS_LINE, borderwidth=1,
            font=dict(family="Arial", size=11, color=TEXT_DARK),
            title=dict(text="<b>Nodo</b>", font=dict(color=TEXT_DARK)),
        ),
        margin=dict(l=70, r=200, t=70, b=60),
    )
    return _ejes_estandar(fig, "Fecha", "PML promedio diario ($/MWh)")


def grafica_heatmap_horario(df, nodo_seleccionado):
    if df.empty: return None
    sub = df[df["nodo"] == nodo_seleccionado].copy()
    if sub.empty: return None
    pivot = sub.pivot_table(values="pml", index="hora", columns="mes_nombre",
                              aggfunc="mean").round(1)
    if not pivot.empty:
        meses_orden = sub.sort_values("fecha_dt")["mes_nombre"].drop_duplicates().tolist()
        pivot = pivot.reindex(columns=meses_orden)

    colorscale = [
        [0.0, "#1a5490"], [0.25, "#5ba0d4"], [0.5, "#FFFFFF"],
        [0.75, "#f5a653"], [1.0, "#a0090c"],
    ]
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values, x=pivot.columns, y=pivot.index,
        colorscale=colorscale,
        zmid=pivot.values.mean() if pivot.size > 0 else 0,
        colorbar=dict(
            title=dict(text="$/MWh", font=dict(color=TEXT_DARK, size=12)),
            thickness=15, len=0.85,
            tickfont=dict(color=TEXT_DARK, size=11),
        ),
        hovertemplate="<b>Hora:</b> %{y}h<br><b>Mes:</b> %{x}<br><b>PML:</b> $%{z}<extra></extra>",
        text=pivot.values,
        texttemplate="%{text:.0f}",
        textfont=dict(size=10, color=TEXT_DARK, family="Arial Black"),
    ))
    fig.update_layout(**_layout_estandar(
        f"Heatmap PML hora × mes · {nodo_seleccionado}", height=540))
    fig.update_yaxes(autorange="reversed", dtick=2,
                      linecolor=AXIS_LINE, linewidth=1.5,
                      tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
                      title=dict(text="Hora del día", font=dict(family="Arial", size=13, color=TEXT_DARK)))
    fig.update_xaxes(linecolor=AXIS_LINE, linewidth=1.5,
                      tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
                      title=dict(text="Mes", font=dict(family="Arial", size=13, color=TEXT_DARK)))
    return fig


def grafica_barras_top(df, metrica="promedio", top_n=10):
    if df.empty: return None
    summary = df.groupby("nodo")["pml"].agg(["mean", "std", "max", "min"])
    summary["pct_neg"] = df.groupby("nodo")["pml"].apply(lambda x: (x < 0).sum() / len(x) * 100)
    summary = summary.reset_index()
    if metrica == "promedio":
        col, titulo, color_bar = "mean", f"Top {top_n} · Mayor PML promedio", RE_NAVY
    elif metrica == "volatilidad":
        col, titulo, color_bar = "std", f"Top {top_n} · Mayor volatilidad", RE_RED
    elif metrica == "negativos":
        col, titulo, color_bar = "pct_neg", f"Top {top_n} · Más horas negativas", RE_BLUE
    else: return None
    top = summary.nlargest(top_n, col).round(2)

    text_template = "%{x:.2f}" if metrica != "negativos" else "%{x:.1f}%"
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=top[col], y=top["nodo"],
        orientation='h',
        marker=dict(color=color_bar, line=dict(color=TEXT_DARK, width=1)),
        text=top[col],
        texttemplate=text_template,
        textposition='outside',
        textfont=dict(family="Arial Black", size=11, color=TEXT_DARK),
        hovertemplate=f"<b>%{{y}}</b><br>{metrica.capitalize()}: %{{x:.2f}}<extra></extra>",
    ))
    fig.update_layout(**_layout_estandar(titulo, height=440))
    fig.update_layout(yaxis=dict(autorange="reversed"))
    x_title = metrica.capitalize() + " ($/MWh)" if metrica != "negativos" else "% de horas"
    return _ejes_estandar(fig, x_title, "")


# ═══════════════════════════════════════════════════════════════════════
# BESS SCORING
# ═══════════════════════════════════════════════════════════════════════
PESOS_BESS = {
    'Arbitraje': {
        'spread_p95_p5':    0.30,
        'spread_avg_diario': 0.25,
        'volatilidad':      0.20,
        'horas_pico':       0.15,
        'pml_promedio':     0.10,
    },
    'Servicios Auxiliares': {
        'volatilidad':      0.40,
        'cambios_bruscos':  0.30,
        'pml_promedio':     0.15,
        'spread_p95_p5':    0.15,
    },
    'Renewables Firming': {
        'pct_horas_neg':    0.35,
        'spread_dia':       0.30,
        'spread_p95_p5':    0.20,
        'pml_promedio':     0.15,
    },
}

DESCRIPCIONES_USE_CASE = {
    'Arbitraje': '💰 **Arbitraje energético** — Comprar barato, vender caro. '
                  'Premia spread alto, volatilidad y picos de precio.',
    'Servicios Auxiliares': '⚙️ **Servicios Auxiliares (SSAA)** — Regulación de '
                             'frecuencia y voltaje. Premia volatilidad y cambios '
                             'bruscos en el precio.',
    'Renewables Firming': '🌅 **Firming de Renovables** — Acompañar generación solar/eólica. '
                           'Premia horas con precios negativos y diferencial día-noche.',
}

DESCRIPCIONES_USE_CASE_PLAIN = {
    'Arbitraje': 'Arbitraje energético — Premia spread alto, volatilidad y picos de precio.',
    'Servicios Auxiliares': 'Servicios Auxiliares — Premia volatilidad y cambios bruscos.',
    'Renewables Firming': 'Renewables Firming — Premia horas negativas y diferencial día-noche.',
}


@st.cache_data(show_spinner=False, max_entries=3)
def calcular_metricas_bess_cached(df_hash, _df):
    if _df.empty: return pd.DataFrame()
    metricas = []
    for nodo in _df["nodo"].unique():
        sub = _df[_df["nodo"] == nodo].copy()
        pml = sub["pml"].values
        if len(pml) == 0: continue

        prom = pml.mean()
        std_v = pml.std()
        p95 = np.percentile(pml, 95)
        p5  = np.percentile(pml, 5)
        spread_p95_p5 = p95 - p5

        sub["fecha_only"] = sub["fecha_dt"].dt.date
        daily = sub.groupby("fecha_only")["pml"].agg(['min', 'max', 'mean'])
        daily["spread"] = daily["max"] - daily["min"]
        spread_avg_diario = daily["spread"].mean() if not daily.empty else 0
        spread_dia = daily["max"].mean() - daily["min"].mean() if not daily.empty else 0

        sub_sorted = sub.sort_values(["fecha_dt", "hora"])
        diffs = sub_sorted["pml"].diff().abs()
        threshold = std_v * 1.5 if std_v > 0 else 0
        cambios_bruscos = (diffs > threshold).sum()

        threshold_pico = p95
        horas_pico = (pml > threshold_pico).sum()

        pct_neg = (pml < 0).mean() * 100

        metricas.append({
            'nodo':          nodo,
            'pml_promedio':  round(prom, 2),
            'volatilidad':   round(std_v, 2),
            'spread_p95_p5': round(spread_p95_p5, 2),
            'spread_avg_diario': round(spread_avg_diario, 2),
            'spread_dia':    round(spread_dia, 2),
            'cambios_bruscos': int(cambios_bruscos),
            'horas_pico':    int(horas_pico),
            'pct_horas_neg': round(pct_neg, 1),
        })

    return pd.DataFrame(metricas)


def calcular_metricas_bess(df):
    if df.empty: return pd.DataFrame()
    key = f"{len(df)}-{df['pml'].sum():.0f}"
    return calcular_metricas_bess_cached(key, df)


def calcular_score_bess(df_metricas, use_case, pesos=None):
    """Score con rank-percentile."""
    if df_metricas.empty: return pd.DataFrame()
    if pesos is None: pesos = PESOS_BESS.get(use_case, PESOS_BESS['Arbitraje'])

    df = df_metricas.copy()
    score = pd.Series(0.0, index=df.index)
    for metrica, peso in pesos.items():
        if metrica not in df.columns: continue
        ranks = df[metrica].rank(pct=True, method='average')
        score += ranks * peso

    df['score'] = (score * 100).round(1)
    return df.sort_values('score', ascending=False).reset_index(drop=True)


def grafica_bess_ranking(df_score, use_case, top_n=15):
    if df_score.empty: return None
    top = df_score.nlargest(top_n, 'score').copy()

    colors = []
    for s in top['score']:
        if s >= 75: colors.append("#1a8a3a")
        elif s >= 50: colors.append("#9bc24f")
        elif s >= 25: colors.append("#f5d017")
        else: colors.append("#a0090c")

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=top['score'], y=top['nodo'],
        orientation='h',
        marker=dict(color=colors, line=dict(color=TEXT_DARK, width=1)),
        text=top['score'],
        texttemplate="%{x:.1f}",
        textposition='outside',
        textfont=dict(family="Arial Black", size=12, color=TEXT_DARK),
        hovertemplate="<b>%{y}</b><br>Score BESS: <b>%{x:.1f}</b>/100<extra></extra>",
    ))
    fig.update_layout(**_layout_estandar(
        f"🏆 Top {top_n} nodos · Score BESS para {use_case}", height=520))
    fig.update_layout(yaxis=dict(autorange="reversed"))
    return _ejes_estandar(fig, "Score BESS (0-100)", "")


# ═══════════════════════════════════════════════════════════════════════
# MAPA INTERACTIVO con LEYENDA CCR
# ═══════════════════════════════════════════════════════════════════════
def _ccr_color_map(matches_df):
    """Asigna un color de PALETTE a cada CCR único en el dataframe."""
    ccrs_unique = sorted(matches_df['ccr'].fillna('?').unique())
    return {ccr: PALETTE[i % len(PALETTE)] for i, ccr in enumerate(ccrs_unique)}


def grafica_mapa(matches_df, df_pml=None, color_by='pml'):
    """color_by: 'pml' (con datos) o 'ccr' (modo solo mapa)."""
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty:
        return None

    if color_by == 'pml' and df_pml is not None and not df_pml.empty:
        pml_avg = df_pml.groupby('nodo')['pml'].agg(['mean', 'count', 'max', 'min']).reset_index()
        pml_avg.columns = ['clave', 'pml_avg', 'pml_count', 'pml_max', 'pml_min']
        map_df = matches_ok.merge(pml_avg, on='clave', how='left')

        map_df['hover'] = map_df.apply(
            lambda r: (
                f"<b>{r['clave']}</b><br>"
                f"<b>Nombre:</b> {r['nombre_cenace']}<br>"
                f"<b>CCR:</b> {r['ccr']} · {r['zona']}<br>"
                f"<b>Voltaje:</b> {r['kv_cenace']} kV<br>"
                f"<b>Estado:</b> {r['estado']}<br>"
                f"<b>Municipio:</b> {r['municipio']}<br>"
                f"━━━━━━━━━━━<br>"
                f"<b>PML promedio:</b> ${r['pml_avg']:.2f}<br>"
                f"<b>PML máx:</b> ${r['pml_max']:.2f}<br>"
                f"<b>PML mín:</b> ${r['pml_min']:.2f}<br>"
                f"━━━━━━━━━━━<br>"
                f"<b>Match:</b> {r['calidad']} (sim {r['similitud']})<br>"
                f"<b>OSM:</b> {r['nombre_osm']}"
            ), axis=1
        )
        marker_color = map_df['pml_avg']
        cmin = map_df['pml_avg'].min()
        cmax = map_df['pml_avg'].max()
        colorscale = [
            [0.0, "#1a8a3a"], [0.25, "#9bc24f"], [0.5, "#f5d017"],
            [0.75, "#f57f17"], [1.0, "#a0090c"],
        ]
        colorbar = dict(
            title=dict(text="<b>PML promedio<br>($/MWh)</b>",
                       font=dict(family="Arial", size=12, color=TEXT_DARK)),
            thickness=18, len=0.7, x=1.02,
            tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
        )
        showscale = True
    else:
        # Modo solo mapa: colorear por CCR
        map_df = matches_ok.copy()
        ccr_to_color = _ccr_color_map(map_df)
        map_df['hover'] = map_df.apply(
            lambda r: (
                f"<b>{r['clave']}</b><br>"
                f"<b>Nombre:</b> {r['nombre_cenace']}<br>"
                f"<b>CCR:</b> {r['ccr']} · {r['zona']}<br>"
                f"<b>Voltaje:</b> {r['kv_cenace']} kV<br>"
                f"<b>Estado:</b> {r['estado']}<br>"
                f"<b>Municipio:</b> {r['municipio']}<br>"
                f"━━━━━━━━━━━<br>"
                f"<b>Match:</b> {r['calidad']} (sim {r['similitud']})<br>"
                f"<b>OSM:</b> {r['nombre_osm']}<br>"
                f"<i>(modo solo mapa — sin datos PML)</i>"
            ), axis=1
        )
        map_df['color_ccr'] = map_df['ccr'].fillna('?').map(ccr_to_color)
        marker_color = map_df['color_ccr']
        cmin = cmax = None
        colorscale = None
        colorbar = None
        showscale = False

    # Auto-zoom
    lat_min, lat_max = map_df['lat'].min(), map_df['lat'].max()
    lon_min, lon_max = map_df['lon'].min(), map_df['lon'].max()
    lat_center = (lat_min + lat_max) / 2
    lon_center = (lon_min + lon_max) / 2
    spread = max(lat_max - lat_min, lon_max - lon_min)
    if spread > 15: zoom = 4
    elif spread > 8: zoom = 5
    elif spread > 4: zoom = 6
    elif spread > 2: zoom = 7
    elif spread > 1: zoom = 8
    else: zoom = 9

    fig = go.Figure()

    # Capa 1: halo blanco grande
    fig.add_trace(go.Scattermapbox(
        lat=map_df['lat'], lon=map_df['lon'],
        mode='markers',
        marker=dict(size=28, color='white', opacity=0.95),
        hoverinfo='skip', showlegend=False,
    ))

    # Capa 2: marker con color
    if color_by == 'pml' and showscale:
        fig.add_trace(go.Scattermapbox(
            lat=map_df['lat'], lon=map_df['lon'],
            mode='markers+text',
            marker=dict(
                size=20, color=marker_color,
                colorscale=colorscale, cmin=cmin, cmax=cmax,
                colorbar=colorbar, showscale=True,
                opacity=0.95,
            ),
            text=map_df['clave'], textposition="top right",
            textfont=dict(size=11, color=TEXT_TITLE, family="Arial Black"),
            hovertext=map_df['hover'], hoverinfo='text',
            name='Nodos', showlegend=False,
        ))
    else:
        fig.add_trace(go.Scattermapbox(
            lat=map_df['lat'], lon=map_df['lon'],
            mode='markers+text',
            marker=dict(size=20, color=marker_color, opacity=0.95),
            text=map_df['clave'], textposition="top right",
            textfont=dict(size=11, color=TEXT_TITLE, family="Arial Black"),
            hovertext=map_df['hover'], hoverinfo='text',
            name='Nodos', showlegend=False,
        ))

    # Capa 3: punto central pequeño negro
    fig.add_trace(go.Scattermapbox(
        lat=map_df['lat'], lon=map_df['lon'],
        mode='markers',
        marker=dict(size=4, color=TEXT_DARK, opacity=1.0),
        hoverinfo='skip', showlegend=False,
    ))

    titulo_mapa = (f"📍 Mapa interactivo · {len(map_df)} nodos · "
                   f"{'colores por PML' if color_by == 'pml' else 'colores por CCR'}")

    fig.update_layout(
        mapbox=dict(
            style='open-street-map',
            center=dict(lat=lat_center, lon=lon_center),
            zoom=zoom,
        ),
        height=680,
        margin=dict(l=0, r=0, t=50, b=0),
        title=dict(
            text=titulo_mapa,
            font=dict(family="Arial Black", size=16, color=TEXT_TITLE),
            x=0.02, xanchor="left",
        ),
        showlegend=False,
        paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=12, font_color=TEXT_DARK,
                        bordercolor=AXIS_LINE),
        uirevision='constant',
    )
    return fig


def render_leyenda_ccr(matches_df):
    """Leyenda de CCR debajo del mapa con colores y conteo."""
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty: return

    ccr_to_color = _ccr_color_map(matches_ok)
    counts = matches_ok['ccr'].fillna('?').value_counts()

    items_html = ""
    for ccr in sorted(ccr_to_color.keys()):
        color = ccr_to_color[ccr]
        count = counts.get(ccr, 0)
        items_html += (
            f'<span class="ccr-legend-item">'
            f'<span class="ccr-dot" style="background:{color}"></span>'
            f'<b>{ccr}</b> ({count})'
            f'</span>'
        )

    legend_html = (
        f'<div class="ccr-legend">'
        f'<b style="color:{TEXT_DARK};font-size:0.92rem;">📍 CCRs en esta consulta:</b><br>'
        f'{items_html}'
        f'</div>'
    )
    st.markdown(legend_html, unsafe_allow_html=True)


def render_panel_ccr(matches_df, df_pml=None):
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty: return

    if df_pml is not None and not df_pml.empty:
        pml_avg_nodo = df_pml.groupby('nodo')['pml'].mean().reset_index()
        pml_avg_nodo.columns = ['clave', 'pml_avg']
        merged = matches_ok.merge(pml_avg_nodo, on='clave', how='left')

        ccr_stats = merged.groupby('ccr').agg(
            nodos=('clave', 'count'),
            pml_promedio=('pml_avg', 'mean'),
            pml_max=('pml_avg', 'max'),
            pml_min=('pml_avg', 'min'),
        ).round(2).reset_index().sort_values('pml_promedio', ascending=False)

        st.markdown("##### 📊 Promedio por CCR (zonas analizadas)")
        st.dataframe(
            ccr_stats, use_container_width=True, hide_index=True,
            column_config={
                "ccr":          st.column_config.TextColumn("CCR", width="medium"),
                "nodos":        st.column_config.NumberColumn("# Nodos", format="%d"),
                "pml_promedio": st.column_config.NumberColumn("PML promedio", format="$%.2f"),
                "pml_max":      st.column_config.NumberColumn("Nodo más caro", format="$%.2f"),
                "pml_min":      st.column_config.NumberColumn("Nodo más barato", format="$%.2f"),
            },
        )
    else:
        ccr_stats = matches_ok.groupby('ccr').agg(
            nodos=('clave', 'count'),
        ).reset_index().sort_values('nodos', ascending=False)

        st.markdown("##### 📊 Conteo por CCR")
        st.dataframe(
            ccr_stats, use_container_width=True, hide_index=True,
            column_config={
                "ccr":   st.column_config.TextColumn("CCR", width="medium"),
                "nodos": st.column_config.NumberColumn("# Nodos", format="%d"),
            },
        )


# ═══════════════════════════════════════════════════════════════════════
# RENDER BESS SCORING
# ═══════════════════════════════════════════════════════════════════════
def render_bess_scoring(df, use_case_default='Arbitraje'):
    st.divider()
    st.markdown("## 🔋 BESS Scoring")
    st.caption("Ranking de nodos por idoneidad para distintos casos de uso. "
               "Score 0-100 usando rank-percentile.")

    use_case = st.selectbox(
        "Caso de uso",
        list(PESOS_BESS.keys()),
        index=list(PESOS_BESS.keys()).index(use_case_default),
        key="bess_use_case",
    )
    st.markdown(f"<div class='mode-badge'>{DESCRIPCIONES_USE_CASE[use_case]}</div>",
                unsafe_allow_html=True)

    df_metricas = calcular_metricas_bess(df)
    if df_metricas.empty:
        st.warning("No hay suficientes datos para calcular métricas BESS.")
        return

    pesos_default = PESOS_BESS[use_case].copy()
    pesos = pesos_default.copy()
    with st.expander("⚙️ Ajustar pesos del scoring (opcional)"):
        st.caption("Los pesos se normalizarán automáticamente.")
        col_p = st.columns(min(3, len(pesos_default)))
        for i, (metrica, valor) in enumerate(pesos_default.items()):
            with col_p[i % len(col_p)]:
                pesos[metrica] = st.slider(
                    metrica.replace('_', ' ').title(),
                    min_value=0.0, max_value=1.0, value=valor, step=0.05,
                    key=f"peso_{use_case}_{metrica}",
                )
        suma = sum(pesos.values())
        if abs(suma - 1.0) > 0.05:
            st.caption(f"⚠️ Pesos suman {suma:.2f}, se normalizará a 1.0")
            pesos = {k: v/suma for k, v in pesos.items()} if suma > 0 else pesos_default

    df_score = calcular_score_bess(df_metricas, use_case, pesos=pesos)

    st.markdown("### 🏆 Top 3 nodos recomendados")
    top3 = df_score.head(3)
    cols_t = st.columns(3)
    medals = ['🥇', '🥈', '🥉']
    for i, (idx, row) in enumerate(top3.iterrows()):
        with cols_t[i]:
            st.metric(
                f"{medals[i]} {row['nodo']}",
                f"{row['score']:.1f}",
                help=(f"PML promedio: ${row['pml_promedio']:.2f} | "
                      f"Volatilidad: ${row['volatilidad']:.2f} | "
                      f"Spread P95-P5: ${row['spread_p95_p5']:.2f}")
            )

    fig_rank = grafica_bess_ranking(df_score, use_case, top_n=min(15, len(df_score)))
    if fig_rank: st.plotly_chart(fig_rank, use_container_width=True)

    st.markdown("### 📋 Tabla completa con métricas")
    st.dataframe(
        df_score, use_container_width=True, hide_index=True,
        column_config={
            "nodo":           st.column_config.TextColumn("Clave", width="small"),
            "score":          st.column_config.NumberColumn("Score BESS", format="%.1f"),
            "pml_promedio":   st.column_config.NumberColumn("PML promedio", format="$%.2f"),
            "volatilidad":    st.column_config.NumberColumn("Volatilidad", format="$%.2f"),
            "spread_p95_p5":  st.column_config.NumberColumn("Spread P95-P5", format="$%.2f"),
            "spread_avg_diario": st.column_config.NumberColumn("Spread día prom", format="$%.2f"),
            "spread_dia":     st.column_config.NumberColumn("Spread día/noche", format="$%.2f"),
            "cambios_bruscos": st.column_config.NumberColumn("Cambios bruscos", format="%d"),
            "horas_pico":     st.column_config.NumberColumn("Horas pico", format="%d"),
            "pct_horas_neg":  st.column_config.NumberColumn("% horas neg", format="%.1f%%"),
        },
    )


# ═══════════════════════════════════════════════════════════════════════
# RENDER DASHBOARD COMPLETO (sin Excel - solo análisis)
# ═══════════════════════════════════════════════════════════════════════
def render_dashboard(acumulado, catalogo, matches_df=None):
    st.divider()
    st.markdown("## 📊 Dashboard analítico")
    st.caption("Pasa el mouse sobre los gráficos para ver valores específicos.")

    df = acumulado_a_dataframe(acumulado, catalogo)
    if df.empty:
        st.warning("No hay datos numéricos válidos.")
        return

    col1, col2, col3, col4 = st.columns(4)
    pml_global = df["pml"].mean()
    pml_max = df["pml"].max()
    pct_neg = (df["pml"] < 0).sum() / len(df) * 100
    nodo_top = df.groupby("nodo")["pml"].mean().idxmax()
    col1.metric("PML promedio global", f"${pml_global:.2f}")
    col2.metric("PML máximo", f"${pml_max:.2f}")
    col3.metric("% horas negativas", f"{pct_neg:.1f}%")
    col4.metric("Nodo top", nodo_top)

    # MAPA
    if matches_df is not None and not matches_df.empty:
        st.markdown("### 🗺️ Mapa interactivo de nodos")
        n_mapeados = matches_df['lat'].notna().sum()
        n_consultados = len(matches_df)
        if n_mapeados == 0:
            st.warning(f"❌ Ninguno de los {n_consultados} nodos pudo ser mapeado en OSM.")
        else:
            cm1, cm2, cm3 = st.columns(3)
            cm1.metric("Mapeados", f"{n_mapeados}/{n_consultados}")
            cc = matches_df['calidad'].value_counts()
            n_exc = cc.get('🥇 Excelente', 0)
            n_bue = cc.get('🥈 Bueno', 0)
            n_ace = cc.get('🥉 Aceptable', 0)
            cm2.metric("🥇 Excelente / 🥈 Bueno", f"{n_exc} / {n_bue}")
            cm3.metric("🥉 Aceptable", n_ace)

            fig_mapa = grafica_mapa(matches_df, df, color_by='pml')
            if fig_mapa: st.plotly_chart(fig_mapa, use_container_width=True)

            # Leyenda CCR
            render_leyenda_ccr(matches_df)

            render_panel_ccr(matches_df, df)

            sin_match = matches_df[matches_df['lat'].isna()]
            if not sin_match.empty:
                with st.expander(f"❌ {len(sin_match)} nodos sin match"):
                    st.dataframe(
                        sin_match[['clave', 'nombre_cenace', 'ccr', 'estado', 'razon']],
                        use_container_width=True, hide_index=True,
                    )

    # TABLA RESUMEN
    st.markdown("### 📋 Tabla resumen por nodo")
    summary = calcular_resumen(df)
    if not summary.empty:
        st.dataframe(
            summary, use_container_width=True, hide_index=True,
            column_config={
                "nodo":        st.column_config.TextColumn("Clave", width="small"),
                "nombre":      st.column_config.TextColumn("Nombre", width="medium"),
                "ccr":         st.column_config.TextColumn("CCR", width="small"),
                "registros":   st.column_config.NumberColumn("# Reg", format="%d"),
                "promedio":    st.column_config.NumberColumn("Promedio", format="$%.2f"),
                "mediana":     st.column_config.NumberColumn("Mediana", format="$%.2f"),
                "maximo":      st.column_config.NumberColumn("Máximo", format="$%.2f"),
                "minimo":      st.column_config.NumberColumn("Mínimo", format="$%.2f"),
                "std":         st.column_config.NumberColumn("Volatilidad", format="$%.2f"),
                "p95":         st.column_config.NumberColumn("P95", format="$%.2f"),
                "p5":          st.column_config.NumberColumn("P5", format="$%.2f"),
                "% horas neg": st.column_config.NumberColumn("% Neg", format="%.1f%%"),
            },
        )

    st.markdown("### 📈 Evolución temporal del PML")
    fig_lin = grafica_lineas_tiempo(df, max_nodos=min(15, len(acumulado)))
    if fig_lin: st.plotly_chart(fig_lin, use_container_width=True)

    st.markdown("### 🔥 Heatmap horario × mensual")
    nodos_disp = sorted(df["nodo"].unique())
    nodo_h = st.selectbox("Selecciona nodo:", nodos_disp, index=0, key="sel_heatmap")
    fig_heat = grafica_heatmap_horario(df, nodo_h)
    if fig_heat: st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("### 🏆 Rankings comparativos")
    tab1, tab2, tab3 = st.tabs(["Mayor PML", "Mayor volatilidad", "Más horas negativas"])
    with tab1:
        fig = grafica_barras_top(df, "promedio", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, use_container_width=True)
    with tab2:
        fig = grafica_barras_top(df, "volatilidad", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, use_container_width=True)
    with tab3:
        fig = grafica_barras_top(df, "negativos", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, use_container_width=True)

    # BESS scoring
    render_bess_scoring(df)


# ═══════════════════════════════════════════════════════════════════════
# RENDER CENTRO DE DESCARGAS (Modo Solo Datos)
# ═══════════════════════════════════════════════════════════════════════
def render_centro_descargas(acumulado, catalogo, sistema, proceso, fecha_ini, fecha_fin,
                              matches_df=None):
    st.divider()
    st.markdown("## 📥 Centro de descargas")
    st.caption("Selecciona los archivos que necesites descargar.")

    df = acumulado_a_dataframe(acumulado, catalogo)
    df_resumen = calcular_resumen(df) if not df.empty else pd.DataFrame()
    df_metricas = calcular_metricas_bess(df) if not df.empty else pd.DataFrame()
    ts = datetime.now().strftime("%Y%m%d_%H%M")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("##### 📊 Excel de Datos")
        st.caption("Datos PML completos: portada + resumen + 1 hoja por nodo con todos los registros horarios.")
        if "excel_datos_bytes" not in st.session_state:
            st.session_state["excel_datos_bytes"] = None
        if st.button("Generar Excel datos", key="btn_gen_datos", type="primary"):
            with st.spinner("Generando Excel de datos..."):
                st.session_state["excel_datos_bytes"] = generar_excel_datos(
                    acumulado, sistema, proceso, fecha_ini, fecha_fin)
        if st.session_state["excel_datos_bytes"]:
            st.download_button(
                label="📥 Descargar",
                data=st.session_state["excel_datos_bytes"],
                file_name=f"PML_CENACE_Datos_{sistema}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_datos",
            )

    with col2:
        st.markdown("##### 📈 Excel de Análisis")
        st.caption("BESS Scoring para los 3 casos de uso + métricas calculadas + resumen estadístico.")
        if "excel_analisis_bytes" not in st.session_state:
            st.session_state["excel_analisis_bytes"] = None
        if st.button("Generar Excel análisis", key="btn_gen_anal", type="primary"):
            if df_metricas.empty:
                st.error("No hay suficientes datos para generar análisis.")
            else:
                with st.spinner("Generando Excel de análisis..."):
                    st.session_state["excel_analisis_bytes"] = generar_excel_analisis(
                        df_metricas, df_resumen, sistema, proceso, fecha_ini, fecha_fin)
        if st.session_state["excel_analisis_bytes"]:
            st.download_button(
                label="📥 Descargar",
                data=st.session_state["excel_analisis_bytes"],
                file_name=f"PML_CENACE_Analisis_BESS_{sistema}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_anal",
            )

    with col3:
        st.markdown("##### 🌍 KMZ Geográfico")
        st.caption("Ubicación de los nodos en Google Earth Pro. Requiere geocodificación.")
        if matches_df is None or matches_df.empty:
            st.info("⚠️ La geocodificación no se hizo. Activa el toggle '🗺️ Incluir geocodificación' en el sidebar y vuelve a ejecutar.")
        else:
            n_mapeados = matches_df['lat'].notna().sum()
            if n_mapeados == 0:
                st.warning("Ningún nodo pudo ser mapeado.")
            else:
                if "kmz_bytes" not in st.session_state:
                    st.session_state["kmz_bytes"] = None
                if st.button(f"Generar KMZ ({n_mapeados} nodos)", key="btn_gen_kmz", type="primary"):
                    with st.spinner("Generando KMZ..."):
                        st.session_state["kmz_bytes"] = generar_kmz(matches_df)
                if st.session_state["kmz_bytes"]:
                    st.download_button(
                        label="📥 Descargar",
                        data=st.session_state["kmz_bytes"],
                        file_name=f"PML_CENACE_Geo_{ts}.kmz",
                        mime="application/vnd.google-earth.kmz",
                        key="dl_kmz",
                    )


# ═══════════════════════════════════════════════════════════════════════
# UI PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>⚡ CENACE PML Analyzer 
        <span class="srf-badge">SRF</span>
    </h1>
    <p>Descarga, análisis, mapeo y BESS scoring — Recurrent Energy / Canadian Solar</p>
</div>
""", unsafe_allow_html=True)

catalogo = cargar_catalogo()
if catalogo:
    st.info(f"📚 Catálogo CENACE cargado: **{len(catalogo):,} nodos** disponibles · "
            f"versión 2026-04-23")
else:
    st.warning("⚠️ Catálogo no cargado.")

# Inicializar session_state
if "consulta_ejecutada" not in st.session_state:
    st.session_state["consulta_ejecutada"] = False
if "acumulado" not in st.session_state:
    st.session_state["acumulado"] = {}
if "matches_df" not in st.session_state:
    st.session_state["matches_df"] = None
if "consulta_params" not in st.session_state:
    st.session_state["consulta_params"] = {}

with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    st.markdown("**🚀 Modo de uso**")
    modo = st.radio(
        "Selecciona qué hacer:",
        options=["🚀 Completo", "📊 Solo datos", "🗺️ Solo mapa"],
        index=0,
        help=(
            "**Completo**: descarga + dashboard + mapa + BESS scoring (sin Excel)\n\n"
            "**Solo datos**: descarga + Centro de descargas (Excel/KMZ)\n\n"
            "**Solo mapa**: solo geocodifica (~5 segundos)"
        ),
        label_visibility="collapsed",
    )

    es_solo_mapa = (modo == "🗺️ Solo mapa")
    es_solo_datos = (modo == "📊 Solo datos")
    es_completo = (modo == "🚀 Completo")
    necesita_datos = es_solo_datos or es_completo
    necesita_mapa = es_solo_mapa or es_completo

    st.markdown("---")
    st.markdown("**📍 Nodos CENACE**")
    nodos_input = st.text_area(
        "Lista de claves",
        height=150,
        placeholder="01VAJ-230\n01XAL-230\n06FUN-115",
        key="nodos_input",
    )

    if necesita_datos:
        col1, col2 = st.columns(2)
        with col1:
            sistema = st.selectbox("Sistema", ["SIN", "BCA", "BCS"], index=0, key="sistema_sel")
        with col2:
            proceso = st.selectbox("Proceso", ["MTR", "MDA"], index=0, key="proceso_sel")

        st.markdown("**📅 Período**")
        col_f1, col_f2 = st.columns(2)
        today = date.today()
        with col_f1:
            f_ini = st.date_input("Desde", value=today - timedelta(days=90), max_value=today, key="f_ini")
        with col_f2:
            f_fin = st.date_input("Hasta", value=today - timedelta(days=1), max_value=today, key="f_fin")
    else:
        sistema = "SIN"; proceso = "MTR"
        today = date.today()
        f_ini = today - timedelta(days=90)
        f_fin = today - timedelta(days=1)

    # Toggle de geocodificación en modo solo datos
    if es_solo_datos:
        st.markdown("**🗺️ Geocodificación**")
        incluir_geo_solo_datos = st.toggle("Incluir geocodificación",
                                            value=False,
                                            key="toggle_geo_datos",
                                            help="Activar si vas a descargar el KMZ. Tarda ~5s extra.")
    else:
        incluir_geo_solo_datos = False

    st.divider()
    st.caption("Sebastian Roldan (SRF)\nRecurrent Energy · Canadian Solar")


col_left, col_right = st.columns([2, 1])
with col_left:
    st.markdown("### 🚀 Ejecutar consulta")

    modo_descripcion = {
        "🚀 Completo": "Descargará datos PML + mapeará nodos + dashboard + BESS scoring (sin Excel).",
        "📊 Solo datos": "Descargará datos PML y mostrará un Centro de Descargas con Excel/KMZ.",
        "🗺️ Solo mapa": ("Solo localizará los nodos en OpenStreetMap. <b>Sin descargar PML.</b>"),
    }
    st.markdown(
        f"<div class='mode-badge'><b>Modo: {modo}</b><br>{modo_descripcion[modo]}</div>",
        unsafe_allow_html=True,
    )

    nodos = []
    if nodos_input.strip():
        nodos_raw = re.split(r"[,\n]+", nodos_input.strip())
        nodos = [n.strip().upper() for n in nodos_raw if n.strip()]
        seen = set()
        nodos = [n for n in nodos if not (n in seen or seen.add(n))]

    if nodos:
        cv = [n for n in nodos if n in catalogo] if catalogo else nodos
        ci = [n for n in nodos if catalogo and n not in catalogo]
        c_a, c_b, c_c = st.columns(3)
        c_a.metric("Solicitados", len(nodos))
        c_b.metric("✅ En catálogo", len(cv))
        c_c.metric("⚠️ Desconocidos", len(ci))
        if ci:
            with st.expander(f"⚠️ {len(ci)} nodos no encontrados"):
                st.code("\n".join(ci))

    if nodos and necesita_datos and f_ini < f_fin:
        n_dias = (f_fin - f_ini).days + 1
        n_bloques = (n_dias + BLOQUE_MAX - 1) // BLOQUE_MAX
        n_lotes = (len(nodos) + 9) // 10
        n_consultas = n_bloques * n_lotes
        tiempo = n_consultas * 1.5 / MAX_WORKERS
        st.caption(f"📊 **{n_dias} días · {n_consultas} consultas · ~{tiempo/60:.1f} min**")
    elif nodos and es_solo_mapa:
        st.caption(f"🗺️ Solo geocodificación de **{len(nodos)} nodos** (~5 segundos)")

    boton = st.button("⚡ Ejecutar", type="primary", disabled=(len(nodos) == 0))

with col_right:
    st.markdown("### 📋 Instrucciones")
    if es_solo_mapa:
        st.markdown("""
        1. Pega claves de nodos
        2. Click en **Ejecutar**
        3. Ve la ubicación geográfica en el mapa
        
        ⚡ **Modo rápido**: no descarga datos PML.
        """)
    elif es_solo_datos:
        st.markdown("""
        1. Pega claves de nodos
        2. Selecciona sistema, proceso, fechas
        3. (Opcional) Activa geocodificación si quieres KMZ
        4. Click en **Ejecutar**
        5. Genera y descarga lo que necesites del Centro de Descargas
        """)
    else:
        st.markdown("""
        1. Pega claves de nodos
        2. Selecciona sistema, proceso, fechas
        3. Click en **Ejecutar**
        4. Explora dashboard interactivo + BESS scoring
        
        💡 Para descargar archivos: usa modo **Solo datos**.
        """)


# ═══════════════════════════════════════════════════════════════════════
# EJECUCIÓN — guarda resultado en session_state
# ═══════════════════════════════════════════════════════════════════════
if boton and nodos:
    fecha_ini = f_ini.strftime("%Y/%m/%d")
    fecha_fin = f_fin.strftime("%Y/%m/%d")
    t0 = time.time()

    # Limpiar estado previo de descargas
    for k in ["excel_datos_bytes", "excel_analisis_bytes", "kmz_bytes"]:
        if k in st.session_state:
            st.session_state[k] = None

    try:
        acumulado = {}
        errores = []
        if necesita_datos:
            if f_ini >= f_fin:
                st.error("❌ La fecha inicial debe ser anterior a la final.")
                st.stop()

            progress = st.progress(0, text="Iniciando descarga...")
            def cb(done, total):
                progress.progress(done/total,
                                  text=f"Descargando: {done}/{total} ({done/total*100:.0f}%)")

            with st.spinner("Consultando CENACE..."):
                acumulado, errores = descargar_pml(
                    nodos, fecha_ini, fecha_fin, sistema, proceso, progress_cb=cb,
                )
            progress.progress(1.0, text="✅ Datos descargados")

            if not acumulado:
                st.error("❌ No se recibieron datos PML.")
                st.stop()

        # Geocodificación si necesita mapa O si en solo datos pidió geo
        matches_df = None
        if necesita_mapa or (es_solo_datos and incluir_geo_solo_datos):
            with st.spinner("Cargando OSM (primera vez ~30s)..."):
                osm_subs = cargar_osm_subestaciones()
            if osm_subs:
                with st.spinner("Matcheando nodos con OSM..."):
                    nodos_a_matchear = list(acumulado.keys()) if necesita_datos else nodos
                    resultados = matchear_nodos(nodos_a_matchear, catalogo, osm_subs)
                    matches_df = pd.DataFrame(resultados)

        # Guardar todo en session_state
        st.session_state["consulta_ejecutada"] = True
        st.session_state["acumulado"] = acumulado
        st.session_state["matches_df"] = matches_df
        st.session_state["consulta_params"] = {
            "sistema": sistema, "proceso": proceso,
            "fecha_ini": fecha_ini, "fecha_fin": fecha_fin,
            "modo": modo, "errores": errores,
            "tiempo": time.time() - t0,
        }

    except Exception as e:
        st.error(f"❌ Error: {type(e).__name__}: {e}")
        st.exception(e)


# ═══════════════════════════════════════════════════════════════════════
# RENDERIZAR RESULTADOS DESDE SESSION_STATE
# Esto se ejecuta cada vez que cambias algo (heatmap, BESS, etc)
# pero SIN volver a descargar datos
# ═══════════════════════════════════════════════════════════════════════
if st.session_state.get("consulta_ejecutada"):
    acumulado = st.session_state["acumulado"]
    matches_df = st.session_state["matches_df"]
    params = st.session_state["consulta_params"]
    modo_cur = params.get("modo", modo)

    if necesita_datos and acumulado:
        n_total = sum(len(f) for f in acumulado.values())
        cx, cy, cz = st.columns(3)
        cx.metric("Nodos con datos", f"{len(acumulado)}/{len(nodos) if nodos else len(acumulado)}")
        cy.metric("Total registros", f"{n_total:,}")
        cz.metric("Tiempo descarga", f"{params.get('tiempo', 0):.0f}s")

        if params.get("errores"):
            with st.expander(f"⚠️ {len(params['errores'])} errores en consultas"):
                st.dataframe(pd.DataFrame(params["errores"][:20]), use_container_width=True)

    # Render según modo
    if "🗺️" in modo_cur:
        # Solo mapa
        if matches_df is not None and not matches_df.empty:
            n_mapeados = matches_df['lat'].notna().sum()
            n_consultados = len(matches_df)
            cm1, cm2, cm3 = st.columns(3)
            cm1.metric("Mapeados", f"{n_mapeados}/{n_consultados}")
            cc = matches_df['calidad'].value_counts()
            cm2.metric("🥇 Excelente / 🥈 Bueno",
                        f"{cc.get('🥇 Excelente', 0)} / {cc.get('🥈 Bueno', 0)}")
            cm3.metric("🥉 Aceptable", cc.get('🥉 Aceptable', 0))

            if n_mapeados > 0:
                fig_mapa = grafica_mapa(matches_df, color_by='ccr')
                if fig_mapa: st.plotly_chart(fig_mapa, use_container_width=True)
                render_leyenda_ccr(matches_df)
                render_panel_ccr(matches_df, None)

                sin_match = matches_df[matches_df['lat'].isna()]
                if not sin_match.empty:
                    with st.expander(f"❌ {len(sin_match)} nodos sin match"):
                        st.dataframe(
                            sin_match[['clave', 'nombre_cenace', 'ccr', 'estado', 'razon']],
                            use_container_width=True, hide_index=True,
                        )

    elif "📊" in modo_cur:
        # Solo datos: Centro de descargas
        if acumulado:
            render_centro_descargas(
                acumulado, catalogo,
                params["sistema"], params["proceso"],
                params["fecha_ini"], params["fecha_fin"],
                matches_df=matches_df,
            )

    else:
        # Modo completo: dashboard sin Excel
        if acumulado:
            render_dashboard(acumulado, catalogo, matches_df)


st.divider()
st.caption("⚡ CENACE PML Analyzer · Sebastian Roldan (SRF) · Recurrent Energy / Canadian Solar")
