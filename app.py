"""
CENACE PML Analyzer — Streamlit Cloud Edition  v3
======================================================
Recurrent Energy / Canadian Solar — SRF · Sebastian Roldan

v3 changes vs v2:
- Mapa interactivo con Plotly Mapbox (OpenStreetMap)
- Matching geográfico OSM con filtro por estado (de v64 notebook)
- KMZ descargable como toggle opcional
- Panel de stats por CCR
- Colores corporativos mejorados para legibilidad
- Heatmap con escala más clara (azul-blanco-rojo)
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

# ═══════════════════════════════════════════════════════════════════════
# COLORES CORPORATIVOS RE
# ═══════════════════════════════════════════════════════════════════════
RE_NAVY  = "#0e346b"
RE_RED   = "#a0090c"
RE_BLUE  = "#2777bd"
RE_ALT   = "#EBF3FB"
RE_INFO  = "#D9E2F3"

# Paleta alta visibilidad para gráficas
PALETTE = [
    "#2777bd", "#a0090c", "#d4a017", "#4a8b3f", "#7e57c2",
    "#ff7043", "#26a69a", "#ec407a", "#5c6bc0", "#8d6e63",
    "#42a5f5", "#9ccc65", "#ffb74d", "#ab47bc", "#78909c",
]

# Colores para openpyxl (sin #)
C_HEADER = "0e346b"
C_SUB    = "2777bd"
C_RED    = "a0090c"
C_WHITE  = "FFFFFF"
C_ALT    = "EBF3FB"
C_INFO   = "D9E2F3"

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
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# CONSTANTES
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
# CARGAR CATÁLOGO CENACE
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
# OSM — DESCARGA + CACHE 7 DIAS
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, ttl=86400 * 7)
def cargar_osm_subestaciones():
    """Descarga subestaciones OSM México. Cache 7 días."""
    # Si ya está en repo, leerlo
    osm_local = os.path.join(os.path.dirname(__file__), 'data', 'osm_subestaciones_mx.json')
    if os.path.exists(osm_local):
        with open(osm_local, 'r', encoding='utf-8') as f:
            return json.load(f)

    # Descargar de Overpass
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
                else:
                    break
            if num_str:
                try:
                    n = float(num_str)
                    nums.append(n / 1000 if n >= 1000 else n)
                except:
                    pass
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
        else:
            continue
        if lat_e == 0 and lon_e == 0:
            continue
        v_kv = _parse_voltaje_max(tags.get('voltage', ''))
        subs.append({
            'osm_id':    el['id'],
            'osm_type':  el['type'],
            'lat':       round(lat_e, 6),
            'lon':       round(lon_e, 6),
            'name':      tags.get('name', ''),
            'operator':  tags.get('operator', ''),
            'voltage_kv': v_kv,
            'voltage_raw': tags.get('voltage', ''),
            'substation': sub_type,
            'rating':     tags.get('rating', ''),
        })
    return subs


# ═══════════════════════════════════════════════════════════════════════
# MATCHING GEOGRÁFICO (de v64)
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

    # 1. exacto
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

    # 2. substring
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

    # 3. difuso
    difusos = []
    for s in candidatos:
        sim = similitud(nombre_norm, s['name_norm'])
        if sim >= 0.80:
            difusos.append((sim, s))

    # 4. keyword
    kw_cenace = palabras_clave(nombre_norm)
    keyword_matches = []
    if kw_cenace:
        for s in candidatos:
            kw_osm = palabras_clave(s.get('name_norm', ''))
            if not kw_osm:
                continue
            compartidas = set(kw_cenace) & set(kw_osm)
            if not compartidas:
                continue
            todas_semi = all(p in PALABRAS_SEMIGENERICAS for p in compartidas)
            largas_no_semi = sum(1 for p in compartidas
                                 if len(p) >= 6 and p not in PALABRAS_SEMIGENERICAS)
            if todas_semi:
                base_score = 0.62
            elif largas_no_semi >= 1:
                base_score = 0.78
            else:
                base_score = 0.70
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
            if 'palabra-clave' in razon:
                calidad = '🥉 Aceptable'
            elif sim >= 0.95:
                calidad = '🥇 Excelente'
            elif sim >= 0.90:
                calidad = '🥈 Bueno'
            else:
                calidad = '🥉 Aceptable'
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
                                   "error": f"URL muy larga ({len(url)})"})
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

def descargar_pml(nodos, fecha_ini, fecha_fin, sistema, proceso,
                  max_workers=8, progress_cb=None):
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
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
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
# EXCEL DATOS
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
    c = ws["B5"]; c.value = "Reporte de Datos — Análisis BESS"
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

    # Resumen
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

    # Por nodo
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
# DASHBOARD COMPONENTS
# ═══════════════════════════════════════════════════════════════════════
def acumulado_a_dataframe(acumulado, catalogo):
    rows = []
    for nodo, filas in acumulado.items():
        info = catalogo.get(nodo, {}) if catalogo else {}
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


def calcular_resumen(df):
    if df.empty: return pd.DataFrame()
    summary = df.groupby("nodo").agg(
        nombre=("nombre", "first"),
        ccr=("ccr", "first"),
        registros=("pml", "count"),
        promedio=("pml", "mean"),
        mediana=("pml", "median"),
        maximo=("pml", "max"),
        minimo=("pml", "min"),
        std=("pml", "std"),
    ).round(2)
    pct_neg = df.groupby("nodo")["pml"].apply(
        lambda x: (x < 0).sum() / len(x) * 100 if len(x) > 0 else 0
    ).round(1)
    summary["% horas neg"] = pct_neg
    summary["p95"] = df.groupby("nodo")["pml"].quantile(0.95).round(2)
    summary["p5"] = df.groupby("nodo")["pml"].quantile(0.05).round(2)
    return summary.reset_index().sort_values("promedio", ascending=False).reset_index(drop=True)


def grafica_lineas_tiempo(df, max_nodos=15):
    if df.empty: return None
    promedios = df.groupby("nodo")["pml"].mean().sort_values(ascending=False)
    nodos_mostrar = promedios.head(max_nodos).index.tolist()
    df_plot = df[df["nodo"].isin(nodos_mostrar)].copy()
    daily = df_plot.groupby(["nodo", "fecha_dt"])["pml"].mean().reset_index()
    fig = px.line(
        daily, x="fecha_dt", y="pml", color="nodo",
        labels={"fecha_dt": "Fecha", "pml": "PML promedio diario ($/MWh)", "nodo": "Nodo"},
        title=f"PML promedio diario · Top {len(nodos_mostrar)} nodos",
        color_discrete_sequence=PALETTE,
    )
    fig.update_traces(line=dict(width=2.2))
    fig.update_layout(
        hovermode="x unified",
        height=480,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Arial", size=12, color="#222"),
        title_font_color=RE_NAVY, title_font_size=15,
        legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
                    bgcolor="rgba(255,255,255,0.9)", bordercolor="#cccccc", borderwidth=1),
        margin=dict(l=60, r=180, t=60, b=50),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#E0E0E0", linecolor="#888")
    fig.update_yaxes(showgrid=True, gridcolor="#E0E0E0", linecolor="#888",
                      zeroline=True, zerolinecolor="#a0090c", zerolinewidth=1.5)
    return fig


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
        z=pivot.values,
        x=pivot.columns,
        y=pivot.index,
        colorscale=colorscale,
        zmid=pivot.values.mean() if pivot.size > 0 else 0,
        colorbar=dict(title="$/MWh", thickness=15, len=0.8),
        hovertemplate="Hora: %{y}h<br>Mes: %{x}<br>PML: $%{z}<extra></extra>",
        text=pivot.values,
        texttemplate="%{text:.0f}",
        textfont=dict(size=9, color="#222"),
    ))
    fig.update_layout(
        title=f"Heatmap PML hora × mes · {nodo_seleccionado}",
        title_font_color=RE_NAVY, title_font_size=15,
        height=520,
        xaxis_title="Mes", yaxis_title="Hora del día",
        font=dict(family="Arial", size=12, color="#222"),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig.update_yaxes(autorange="reversed", dtick=2, linecolor="#888")
    fig.update_xaxes(linecolor="#888")
    return fig


def grafica_barras_top(df, metrica="promedio", top_n=10):
    if df.empty: return None
    summary = df.groupby("nodo")["pml"].agg(["mean", "std", "max", "min"])
    summary["pct_neg"] = df.groupby("nodo")["pml"].apply(lambda x: (x < 0).sum() / len(x) * 100)
    summary = summary.reset_index()
    if metrica == "promedio":
        col, titulo, color_bar = "mean", f"Top {top_n} · Mayor PML promedio", RE_NAVY
    elif metrica == "volatilidad":
        col, titulo, color_bar = "std", f"Top {top_n} · Mayor volatilidad (std)", RE_RED
    elif metrica == "negativos":
        col, titulo, color_bar = "pct_neg", f"Top {top_n} · Más horas negativas", RE_BLUE
    else: return None
    top = summary.nlargest(top_n, col)
    fig = px.bar(
        top, x=col, y="nodo", orientation="h",
        labels={col: metrica.capitalize() + " ($/MWh)" if metrica != "negativos" else "% horas",
                "nodo": ""},
        title=titulo,
    )
    fig.update_traces(marker_color=color_bar)
    fig.update_layout(
        height=420,
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial", size=12, color="#222"),
        title_font_color=RE_NAVY, title_font_size=15,
        yaxis=dict(autorange="reversed"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#E0E0E0", linecolor="#888")
    fig.update_yaxes(linecolor="#888")
    return fig


# ═══════════════════════════════════════════════════════════════════════
# MAPA INTERACTIVO
# ═══════════════════════════════════════════════════════════════════════
def grafica_mapa(matches_df, df_pml):
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty:
        return None

    pml_avg = df_pml.groupby('nodo')['pml'].agg(['mean', 'count', 'max', 'min']).reset_index()
    pml_avg.columns = ['clave', 'pml_avg', 'pml_count', 'pml_max', 'pml_min']
    map_df = matches_ok.merge(pml_avg, on='clave', how='left')

    # Hover
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

    # Auto-zoom
    lat_min, lat_max = map_df['lat'].min(), map_df['lat'].max()
    lon_min, lon_max = map_df['lon'].min(), map_df['lon'].max()
    lat_center = (lat_min + lat_max) / 2
    lon_center = (lon_min + lon_max) / 2
    spread = max(lat_max - lat_min, lon_max - lon_min)
    if spread > 15:    zoom = 4
    elif spread > 8:   zoom = 5
    elif spread > 4:   zoom = 6
    elif spread > 2:   zoom = 7
    elif spread > 1:   zoom = 8
    else:              zoom = 9

    fig = go.Figure()
    fig.add_trace(go.Scattermapbox(
        lat=map_df['lat'],
        lon=map_df['lon'],
        mode='markers+text',
        marker=dict(
            size=14,
            color=map_df['pml_avg'],
            colorscale=[
                [0.0, "#1a8a3a"],
                [0.25, "#9bc24f"],
                [0.5, "#f5d017"],
                [0.75, "#f57f17"],
                [1.0, "#a0090c"],
            ],
            cmin=map_df['pml_avg'].min(),
            cmax=map_df['pml_avg'].max(),
            colorbar=dict(
                title="PML<br>promedio<br>($/MWh)",
                thickness=15, len=0.7, x=1.02,
            ),
            opacity=0.92,
        ),
        text=map_df['clave'],
        textposition="top center",
        textfont=dict(size=10, color=RE_NAVY, family="Arial"),
        hovertext=map_df['hover'],
        hoverinfo='text',
        name='Nodos CENACE',
    ))

    fig.update_layout(
        mapbox=dict(
            style='open-street-map',
            center=dict(lat=lat_center, lon=lon_center),
            zoom=zoom,
        ),
        height=650,
        margin=dict(l=0, r=0, t=40, b=0),
        title=dict(
            text=f"📍 Mapa interactivo · {len(map_df)} nodos · pasa el mouse para ver detalles",
            font=dict(color=RE_NAVY, size=15, family="Arial"),
        ),
        showlegend=False,
        paper_bgcolor="white",
    )
    return fig


def render_panel_ccr(matches_df, df_pml):
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty: return

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


# ═══════════════════════════════════════════════════════════════════════
# RENDER DASHBOARD
# ═══════════════════════════════════════════════════════════════════════
def render_dashboard(acumulado, catalogo, matches_df=None, generar_kmz_flag=False):
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

            fig_mapa = grafica_mapa(matches_df, df)
            if fig_mapa:
                st.plotly_chart(fig_mapa, use_container_width=True)

            render_panel_ccr(matches_df, df)

            sin_match = matches_df[matches_df['lat'].isna()]
            if not sin_match.empty:
                with st.expander(f"❌ {len(sin_match)} nodos sin match"):
                    st.dataframe(
                        sin_match[['clave', 'nombre_cenace', 'ccr', 'estado', 'razon']],
                        use_container_width=True, hide_index=True,
                    )

            if generar_kmz_flag and n_mapeados > 0:
                with st.spinner("Generando KMZ..."):
                    kmz_bytes = generar_kmz(matches_df)
                ts = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(
                    label=f"🌍 Descargar KMZ para Google Earth Pro ({n_mapeados} nodos)",
                    data=kmz_bytes,
                    file_name=f"PML_CENACE_Geo_{ts}.kmz",
                    mime="application/vnd.google-earth.kmz",
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

    # LÍNEAS
    st.markdown("### 📈 Evolución temporal del PML")
    fig_lin = grafica_lineas_tiempo(df, max_nodos=min(15, len(acumulado)))
    if fig_lin: st.plotly_chart(fig_lin, use_container_width=True)

    # HEATMAP
    st.markdown("### 🔥 Heatmap horario × mensual")
    nodos_disp = sorted(df["nodo"].unique())
    nodo_h = st.selectbox("Selecciona nodo:", nodos_disp, index=0, key="sel_heatmap")
    fig_heat = grafica_heatmap_horario(df, nodo_h)
    if fig_heat: st.plotly_chart(fig_heat, use_container_width=True)

    # RANKINGS
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


# ═══════════════════════════════════════════════════════════════════════
# UI PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>⚡ CENACE PML Analyzer 
        <span class="srf-badge">SRF</span>
    </h1>
    <p>Descarga, análisis y mapeo geográfico de PML — Recurrent Energy / Canadian Solar</p>
</div>
""", unsafe_allow_html=True)

catalogo = cargar_catalogo()
if catalogo:
    st.info(f"📚 Catálogo CENACE cargado: **{len(catalogo):,} nodos** disponibles · "
            f"versión 2026-04-23")
else:
    st.warning("⚠️ Catálogo no cargado.")

with st.sidebar:
    st.markdown("### ⚙️ Parámetros de consulta")
    st.markdown("**📍 Nodos CENACE**")
    nodos_input = st.text_area(
        "Lista de claves",
        height=150,
        placeholder="01VAJ-230\n01XAL-230\n06FUN-115",
    )

    col1, col2 = st.columns(2)
    with col1:
        sistema = st.selectbox("Sistema", ["SIN", "BCA", "BCS"], index=0)
    with col2:
        proceso = st.selectbox("Proceso", ["MTR", "MDA"], index=0)

    st.markdown("**📅 Período**")
    col_f1, col_f2 = st.columns(2)
    today = date.today()
    with col_f1:
        f_ini = st.date_input("Desde", value=today - timedelta(days=90), max_value=today)
    with col_f2:
        f_fin = st.date_input("Hasta", value=today - timedelta(days=1), max_value=today)

    st.markdown("**🗺️ Geocodificación**")
    activar_geo = st.toggle("Activar mapa interactivo", value=True,
                             help="Hace matching con OpenStreetMap")
    generar_kmz_flag = False
    if activar_geo:
        generar_kmz_flag = st.toggle("Generar también KMZ descargable", value=False,
                                       help="Para abrir en Google Earth Pro")

    st.markdown("**⚡ Performance**")
    max_workers = st.slider("Workers paralelos", 3, 12, 8)

    st.divider()
    st.caption("Sebastian Roldan (SRF)\nRecurrent Energy · Canadian Solar")


col_left, col_right = st.columns([2, 1])
with col_left:
    st.markdown("### 🚀 Ejecutar consulta")
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

    if nodos and f_ini < f_fin:
        n_dias = (f_fin - f_ini).days + 1
        n_bloques = (n_dias + BLOQUE_MAX - 1) // BLOQUE_MAX
        n_lotes = (len(nodos) + 9) // 10
        n_consultas = n_bloques * n_lotes
        tiempo = n_consultas * 1.5 / max_workers
        st.caption(f"📊 **{n_dias} días · {n_consultas} consultas · ~{tiempo/60:.1f} min**")

    boton = st.button("⚡ Descargar datos", type="primary", disabled=(len(nodos) == 0))

with col_right:
    st.markdown("### 📋 Instrucciones")
    st.markdown("""
    1. Pega claves de nodos
    2. Selecciona sistema, proceso, fechas
    3. (Opcional) Activa mapa interactivo
    4. Click en **Descargar datos**
    5. Explora las visualizaciones
    """)


if boton and nodos and f_ini < f_fin:
    fecha_ini = f_ini.strftime("%Y/%m/%d")
    fecha_fin = f_fin.strftime("%Y/%m/%d")
    progress = st.progress(0, text="Iniciando...")
    t0 = time.time()

    def cb(done, total):
        progress.progress(done/total,
                          text=f"Descargando: {done}/{total} ({done/total*100:.0f}%)")

    try:
        with st.spinner("Consultando CENACE..."):
            acumulado, errores = descargar_pml(
                nodos, fecha_ini, fecha_fin, sistema, proceso,
                max_workers=max_workers, progress_cb=cb,
            )
        elapsed = time.time() - t0
        progress.progress(1.0, text=f"✅ Completo en {elapsed:.0f}s")

        if not acumulado:
            st.error("❌ No se recibieron datos.")
        else:
            n_total = sum(len(f) for f in acumulado.values())
            cx, cy, cz = st.columns(3)
            cx.metric("Nodos con datos", f"{len(acumulado)}/{len(nodos)}")
            cy.metric("Total registros", f"{n_total:,}")
            cz.metric("Tiempo", f"{elapsed:.0f}s")

            if errores:
                with st.expander(f"⚠️ {len(errores)} errores"):
                    st.dataframe(pd.DataFrame(errores[:20]), use_container_width=True)

            # Excel
            st.markdown("### 📥 Descargar Excel")
            with st.spinner("Generando Excel..."):
                excel_bytes = generar_excel_datos(acumulado, sistema, proceso, fecha_ini, fecha_fin)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label=f"📊 Descargar PML_CENACE_{sistema}_{ts}.xlsx",
                data=excel_bytes,
                file_name=f"PML_CENACE_{sistema}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

            # Geo
            matches_df = None
            if activar_geo:
                st.markdown("### 🗺️ Geocodificación de nodos")
                with st.spinner("Cargando OSM (primera vez ~30s)..."):
                    osm_subs = cargar_osm_subestaciones()
                if osm_subs:
                    st.caption(f"📍 {len(osm_subs):,} subestaciones OSM cargadas. Matcheando...")
                    with st.spinner("Matcheando nodos con OSM..."):
                        nodos_con_datos = list(acumulado.keys())
                        resultados = matchear_nodos(nodos_con_datos, catalogo, osm_subs)
                        matches_df = pd.DataFrame(resultados)
                else:
                    st.warning("⚠️ No se pudieron cargar las subestaciones OSM.")

            render_dashboard(acumulado, catalogo, matches_df, generar_kmz_flag)

    except Exception as e:
        st.error(f"❌ Error: {type(e).__name__}: {e}")
        st.exception(e)


st.divider()
st.caption("⚡ CENACE PML Analyzer · Sebastian Roldan (SRF) · Recurrent Energy / Canadian Solar")
