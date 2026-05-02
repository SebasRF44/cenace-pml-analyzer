"""
CENACE PML Analyzer — Streamlit Cloud Edition
================================================
Recurrent Energy / Canadian Solar — SRF
Sebastian Roldan
"""

import streamlit as st
import pandas as pd
import numpy as np
import requests
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import urllib3
import time
import io
import os
import re
from datetime import datetime, timedelta, date
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
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
RE_WHITE = "#FFFFFF"
RE_ALT   = "#EBF3FB"
RE_INFO  = "#D9E2F3"
RE_TOTAL = "#D6E4F0"

# Colores SIN prefijo # para openpyxl
C_HEADER = "0e346b"
C_SUB    = "2777bd"
C_RED    = "a0090c"
C_WHITE  = "FFFFFF"
C_ALT    = "EBF3FB"
C_TOTAL  = "D6E4F0"
C_INFO   = "D9E2F3"

# Estilos CSS para mejorar visualmente la app
st.markdown(f"""
<style>
    /* Header banner principal */
    .main-header {{
        background: linear-gradient(135deg, {RE_NAVY} 0%, {RE_BLUE} 100%);
        padding: 2rem;
        border-radius: 8px;
        margin-bottom: 2rem;
        color: white;
    }}
    .main-header h1 {{
        color: white !important;
        margin: 0;
        font-size: 2rem;
    }}
    .main-header p {{
        color: rgba(255,255,255,0.85);
        margin: 0.5rem 0 0 0;
    }}
    /* Marca SRF discreta */
    .srf-badge {{
        background: {RE_RED};
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 4px;
        font-weight: bold;
        font-size: 0.85rem;
        display: inline-block;
        margin-left: 1rem;
    }}
    /* Botones primary */
    .stButton > button[kind="primary"] {{
        background: {RE_NAVY};
        color: white;
        border: none;
    }}
    .stButton > button[kind="primary"]:hover {{
        background: {RE_RED};
    }}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# CONSTANTES CENACE API
# ═══════════════════════════════════════════════════════════════════════
BASE_URL    = "https://ws01.cenace.gob.mx:8082/SWPML/SIM"
BLOQUE_MAX  = 7
FORMATO     = "JSON"
TIMEOUT     = 60

# ═══════════════════════════════════════════════════════════════════════
# CARGAR CATÁLOGO CENACE (cached para no leer Excel cada vez)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data
def cargar_catalogo():
    """Carga el catálogo de nodos CENACE incluido en el repo."""
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
# FUNCIONES CENACE
# ═══════════════════════════════════════════════════════════════════════
def parse_fecha(s):
    return datetime.strptime(s, "%Y/%m/%d")

def fmt(d):
    return d.strftime("%Y/%m/%d")

def generar_bloques(ini_str, fin_str, max_dias=BLOQUE_MAX):
    ini = parse_fecha(ini_str)
    fin = parse_fecha(fin_str)
    bloques = []
    cursor = ini
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
    """Consulta CENACE con reintentos. Thread-safe."""
    url = construir_url(nodos_lista, fecha_ini, fecha_fin, sistema, proceso)

    if len(url) > 2000:
        with lock:
            errores_lista.append({
                "fecha": f"{fecha_ini} → {fecha_fin}",
                "nodos": len(nodos_lista),
                "error": f"URL muy larga ({len(url)} chars)"
            })
        return None

    session = requests.Session()
    session.verify = False

    for intento in range(max_reintentos):
        try:
            r = session.get(url, timeout=TIMEOUT)
            if r.status_code == 200:
                return r.text
            elif r.status_code == 204:
                return None
            elif r.status_code == 429:
                time.sleep(2 ** (intento + 2))
                continue
            else:
                if intento < max_reintentos - 1:
                    time.sleep(2 ** (intento + 1))
                    continue
                with lock:
                    errores_lista.append({
                        "fecha": f"{fecha_ini} → {fecha_fin}",
                        "nodos": len(nodos_lista),
                        "error": f"HTTP {r.status_code}"
                    })
                return None
        except requests.exceptions.Timeout:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento)
                continue
            with lock:
                errores_lista.append({
                    "fecha": f"{fecha_ini} → {fecha_fin}",
                    "nodos": len(nodos_lista),
                    "error": "Timeout"
                })
            return None
        except Exception as e:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento)
                continue
            with lock:
                errores_lista.append({
                    "fecha": f"{fecha_ini} → {fecha_fin}",
                    "nodos": len(nodos_lista),
                    "error": f"{type(e).__name__}: {str(e)[:80]}"
                })
            return None
    return None


def _num(val):
    try:
        return float(val)
    except (TypeError, ValueError):
        return val


def parsear_json(texto):
    """Parsea respuesta JSON CENACE. Retorna dict nodo → lista de registros."""
    if not texto:
        return {}
    try:
        obj = json.loads(texto)
    except Exception:
        return {}

    datos = {}
    for nd in obj.get("Resultados", []):
        clv = nd.get("clv_nodo", "?")
        valores = nd.get("Valores", [])
        if not valores:
            continue
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
    """Descarga datos PML del CENACE en paralelo. Retorna dict nodo → registros."""
    bloques = generar_bloques(fecha_ini, fecha_fin)

    # Lotes de hasta 10 nodos por consulta (límite CENACE de 20, dejamos margen)
    LOTE = 10
    lotes_nodos = [nodos[i:i+LOTE] for i in range(0, len(nodos), LOTE)]

    consultas = []
    for lote in lotes_nodos:
        for bi, bf in bloques:
            consultas.append((lote, bi, bf))

    total = len(consultas)
    acumulado = {n: [] for n in nodos}
    errores_consulta = []
    lock = Lock()
    completed = 0

    def _job(consulta):
        lote, bi, bf = consulta
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

    # Eliminar nodos sin datos
    acumulado = {n: regs for n, regs in acumulado.items() if regs}
    return acumulado, errores_consulta


# ═══════════════════════════════════════════════════════════════════════
# GENERAR EXCEL DE DATOS — Versión Streamlit (devuelve bytes)
# ═══════════════════════════════════════════════════════════════════════
def generar_excel_datos(acumulado, sistema, proceso, fecha_ini, fecha_fin):
    """Genera Excel con portada + resumen + hojas por nodo. Retorna bytes."""
    _NOW = datetime.now().strftime("%Y-%m-%d %H:%M")

    _side = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=_side, right=_side, top=_side, bottom=_side)

    def hdr(cell, bg=C_HEADER, fg=C_WHITE, size=11):
        cell.font = Font(bold=True, color=fg, size=size, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE

    # Estilos precomputados
    _FONT_DATO = Font(name="Arial", size=10)
    _ALIGN_NUM = Alignment(horizontal="center", vertical="center")
    _ALIGN_TXT = Alignment(horizontal="left", vertical="center")
    _FILL_ALT  = PatternFill("solid", start_color=C_ALT)
    _NUM_FMT   = "#,##0.00"

    COLS   = ["Fecha", "Hora", "PML ($/MWh)", "Energía ($/MWh)", "Pérdidas ($/MWh)", "Congestión ($/MWh)"]
    ANCHOS = [14, 8, 16, 16, 16, 18]
    ES_NUM = [False, True, True, True, True, True]

    wb = Workbook()

    # ── PORTADA ──
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
    c = ws["B5"]
    c.value = "Reporte de Datos — Análisis BESS"
    c.font = Font(bold=True, italic=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_RED)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("B7:H7")
    c = ws["B7"]
    c.value = "PARÁMETROS DE LA CONSULTA"
    c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 22

    info_rows = [
        ("Sistema",          sistema),
        ("Proceso",          proceso),
        ("Fecha inicial",    fecha_ini),
        ("Fecha final",      fecha_fin),
        ("Nodos en reporte", f"{len(acumulado):,} nodos"),
        ("Total registros",  f"{sum(len(f) for f in acumulado.values()):,} filas"),
        ("Fecha generación", _NOW),
    ]
    for ri, (lbl, val) in enumerate(info_rows, start=8):
        ws.merge_cells(f"B{ri}:D{ri}")
        c1 = ws[f"B{ri}"]
        c1.value = lbl
        c1.font = Font(bold=True, color=C_HEADER, size=11, name="Arial")
        c1.fill = PatternFill("solid", start_color=C_INFO)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c1.border = BORDE

        ws.merge_cells(f"E{ri}:H{ri}")
        c2 = ws[f"E{ri}"]
        c2.value = val
        c2.font = Font(size=11, name="Arial")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2.border = BORDE
        ws.row_dimensions[ri].height = 22

    # Watermark SRF
    ws.merge_cells("B17:H19")
    c = ws["B17"]
    c.value = "SRF"
    c.font = Font(bold=True, color=C_HEADER, size=72, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(17, 20):
        ws.row_dimensions[r].height = 30

    ws.merge_cells("B21:H21")
    c = ws["B21"]
    c.value = "Prepared by: Sebastian Roldan · Recurrent Energy"
    c.font = Font(italic=True, bold=True, color=C_HEADER, size=12, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=C_INFO)
    ws.row_dimensions[21].height = 22

    ws.merge_cells("B22:H22")
    c = ws["B22"]
    c.value = "A subsidiary of Canadian Solar"
    c.font = Font(italic=True, color="555555", size=10, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[22].height = 18

    for col_letter, w in [("A", 3), ("B", 14), ("C", 14), ("D", 14),
                          ("E", 14), ("F", 14), ("G", 14), ("H", 14), ("I", 3)]:
        ws.column_dimensions[col_letter].width = w

    # ── RESUMEN ──
    ws_res = wb.create_sheet("Resumen")
    ws_res.merge_cells("A1:F1")
    c = ws_res["A1"]
    c.value = f"Resumen — {len(acumulado)} nodos · {sistema} · {proceso}"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 28

    res_cols = ["Nodo", "# Registros", "PML promedio", "PML máximo", "PML mínimo", "Período"]
    for ci, col in enumerate(res_cols, 1):
        hdr(ws_res.cell(row=2, column=ci, value=col), bg=C_SUB)
    ws_res.row_dimensions[2].height = 22

    for ri, (nodo, filas) in enumerate(acumulado.items(), start=3):
        if not filas:
            continue
        pmls = [f["pml"] for f in filas if isinstance(f["pml"], (int, float))]
        bg = C_ALT if ri % 2 == 0 else None
        valores = [
            nodo,
            len(filas),
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
            if bg:
                cc.fill = PatternFill("solid", start_color=bg)

    for ci, w in enumerate([14, 14, 14, 14, 14, 24], 1):
        ws_res.column_dimensions[get_column_letter(ci)].width = w
    ws_res.freeze_panes = "A3"

    # ── HOJAS POR NODO ──
    for nodo, filas in acumulado.items():
        ws_n = wb.create_sheet(title=nodo[:31])

        # Título
        ws_n.merge_cells("A1:F1")
        c = ws_n["A1"]
        c.value = f"PML — Nodo: {nodo}"
        c.font = Font(bold=True, color=C_WHITE, size=13, name="Arial")
        c.fill = PatternFill("solid", start_color=C_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_n.row_dimensions[1].height = 28

        # Subtítulo SRF
        ws_n.merge_cells("A2:F2")
        c = ws_n["A2"]
        c.value = (f"Sistema: {sistema}  |  Proceso: {proceso}  |  "
                   f"Período: {fecha_ini} → {fecha_fin}  |  "
                   f"Total: {len(filas):,}  |  "
                   f"Sebastian Roldan (SRF) · Recurrent Energy")
        c.font = Font(italic=True, size=9, name="Arial", color="555555")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", start_color=C_INFO)
        ws_n.row_dimensions[2].height = 16

        # Headers
        for ci, (enc, ancho) in enumerate(zip(COLS, ANCHOS), 1):
            hdr(ws_n.cell(row=3, column=ci, value=enc), bg=C_SUB)
            ws_n.column_dimensions[get_column_letter(ci)].width = ancho
        ws_n.row_dimensions[3].height = 22

        # Datos
        for fila in filas:
            ws_n.append([fila["fecha"], fila["hora"], fila["pml"],
                         fila["pml_ene"], fila["pml_per"], fila["pml_cng"]])

        # Aplicar estilos por bloques
        for i in range(len(filas)):
            row_idx = i + 4
            es_alt = (i % 2 == 0)
            for ci in range(1, 7):
                cc = ws_n.cell(row=row_idx, column=ci)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ES_NUM[ci-1] else _ALIGN_TXT
                if es_alt:
                    cc.fill = _FILL_ALT
                if ES_NUM[ci-1]:
                    cc.number_format = _NUM_FMT

        ws_n.freeze_panes = "A4"

    # Guardar a bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# DASHBOARD ANALÍTICO INLINE — v2 con Plotly
# ═══════════════════════════════════════════════════════════════════════
def acumulado_a_dataframe(acumulado, catalogo):
    """Convierte el dict acumulado a un DataFrame plano para análisis."""
    rows = []
    for nodo, filas in acumulado.items():
        info = catalogo.get(nodo, {}) if catalogo else {}
        for f in filas:
            try:
                pml_val = float(f["pml"])
            except (TypeError, ValueError):
                continue
            rows.append({
                "nodo":      nodo,
                "ccr":       info.get("ccr", "?"),
                "nombre":    info.get("nombre", nodo),
                "estado":    info.get("estado", ""),
                "fecha":     f["fecha"],
                "hora":      int(f["hora"]) if f["hora"] else 0,
                "pml":       pml_val,
            })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["fecha_dt"] = pd.to_datetime(df["fecha"], errors="coerce")
        df["mes"] = df["fecha_dt"].dt.month
        df["mes_nombre"] = df["fecha_dt"].dt.strftime("%b %Y")
    return df


def calcular_resumen(df):
    """Tabla resumen estadística por nodo."""
    if df.empty:
        return pd.DataFrame()
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
    # % de horas con precio negativo
    pct_neg = df.groupby("nodo")["pml"].apply(
        lambda x: (x < 0).sum() / len(x) * 100 if len(x) > 0 else 0
    ).round(1)
    summary["% horas neg"] = pct_neg
    # P95 y P5
    summary["p95"] = df.groupby("nodo")["pml"].quantile(0.95).round(2)
    summary["p5"] = df.groupby("nodo")["pml"].quantile(0.05).round(2)
    summary = summary.reset_index()
    summary = summary.sort_values("promedio", ascending=False).reset_index(drop=True)
    return summary


def grafica_lineas_tiempo(df, max_nodos=15):
    """Gráfica de línea: PML promedio diario por nodo."""
    if df.empty:
        return None
    # Si hay muchos nodos, mostrar solo los top N por promedio
    promedios = df.groupby("nodo")["pml"].mean().sort_values(ascending=False)
    nodos_mostrar = promedios.head(max_nodos).index.tolist()
    df_plot = df[df["nodo"].isin(nodos_mostrar)].copy()

    # Agrupar por nodo y fecha → promedio diario
    daily = df_plot.groupby(["nodo", "fecha_dt"])["pml"].mean().reset_index()

    fig = px.line(
        daily, x="fecha_dt", y="pml", color="nodo",
        labels={"fecha_dt": "Fecha", "pml": "PML promedio diario ($/MWh)", "nodo": "Nodo"},
        title=f"PML promedio diario · Top {len(nodos_mostrar)} nodos por precio promedio",
    )
    fig.update_layout(
        hovermode="x unified",
        height=450,
        plot_bgcolor="#F8F9FA",
        paper_bgcolor="white",
        font=dict(family="Arial", size=11),
        title_font_color="#0e346b",
        legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#EBF3FB")
    fig.update_yaxes(showgrid=True, gridcolor="#EBF3FB", zeroline=True, zerolinecolor="#a0090c")
    return fig


def grafica_heatmap_horario(df, nodo_seleccionado=None):
    """Heatmap promedio PML por hora x mes para un nodo específico."""
    if df.empty:
        return None
    if nodo_seleccionado is None:
        # Tomar el nodo con mayor promedio
        nodo_seleccionado = df.groupby("nodo")["pml"].mean().idxmax()

    sub = df[df["nodo"] == nodo_seleccionado].copy()
    if sub.empty:
        return None

    pivot = sub.pivot_table(values="pml", index="hora", columns="mes_nombre",
                              aggfunc="mean").round(1)
    # Ordenar columnas cronológicamente
    if not pivot.empty:
        meses_orden = sub.sort_values("fecha_dt")["mes_nombre"].drop_duplicates().tolist()
        pivot = pivot.reindex(columns=meses_orden)

    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns,
        y=pivot.index,
        colorscale=[
            [0.0, "#0e346b"],
            [0.3, "#2777bd"],
            [0.5, "#EBF3FB"],
            [0.7, "#f5a653"],
            [1.0, "#a0090c"],
        ],
        colorbar=dict(title="$/MWh"),
        hovertemplate="Hora: %{y}<br>Mes: %{x}<br>PML: %{z} $/MWh<extra></extra>",
    ))
    fig.update_layout(
        title=f"Heatmap PML hora × mes · {nodo_seleccionado}",
        title_font_color="#0e346b",
        height=450,
        xaxis_title="Mes",
        yaxis_title="Hora del día",
        font=dict(family="Arial", size=11),
    )
    fig.update_yaxes(autorange="reversed", dtick=2)
    return fig


def grafica_barras_top(df, metrica="promedio", top_n=10):
    """Top N nodos por una métrica específica."""
    if df.empty:
        return None
    summary = df.groupby("nodo")["pml"].agg(["mean", "std", "max", "min"])
    summary["pct_neg"] = df.groupby("nodo")["pml"].apply(
        lambda x: (x < 0).sum() / len(x) * 100
    )
    summary = summary.reset_index()

    if metrica == "promedio":
        col = "mean"
        titulo = f"Top {top_n} nodos · Mayor PML promedio"
        color_bar = "#0e346b"
    elif metrica == "volatilidad":
        col = "std"
        titulo = f"Top {top_n} nodos · Mayor volatilidad (std)"
        color_bar = "#a0090c"
    elif metrica == "negativos":
        col = "pct_neg"
        titulo = f"Top {top_n} nodos · Mayor % horas negativas"
        color_bar = "#2777bd"
    else:
        return None

    top = summary.nlargest(top_n, col)
    fig = px.bar(
        top, x=col, y="nodo", orientation="h",
        labels={col: metrica.capitalize() + " ($/MWh)" if metrica != "negativos" else "% horas",
                "nodo": ""},
        title=titulo,
    )
    fig.update_traces(marker_color=color_bar)
    fig.update_layout(
        height=400,
        plot_bgcolor="#F8F9FA",
        paper_bgcolor="white",
        font=dict(family="Arial", size=11),
        title_font_color="#0e346b",
        yaxis=dict(autorange="reversed"),
    )
    return fig


def render_dashboard(acumulado, catalogo):
    """Renderiza todo el dashboard analítico inline."""
    st.divider()
    st.markdown("## 📊 Dashboard analítico")
    st.caption("Visualizaciones interactivas de los datos descargados — pasa el mouse sobre los gráficos para ver valores específicos.")

    df = acumulado_a_dataframe(acumulado, catalogo)

    if df.empty:
        st.warning("No hay datos numéricos válidos para graficar.")
        return

    # ─── MÉTRICAS GLOBALES ───
    col1, col2, col3, col4 = st.columns(4)
    pml_global_prom = df["pml"].mean()
    pml_global_max = df["pml"].max()
    pct_neg_global = (df["pml"] < 0).sum() / len(df) * 100
    nodo_top = df.groupby("nodo")["pml"].mean().idxmax()

    col1.metric("PML promedio global", f"${pml_global_prom:.2f}")
    col2.metric("PML máximo histórico", f"${pml_global_max:.2f}")
    col3.metric("% horas negativas", f"{pct_neg_global:.1f}%")
    col4.metric("Nodo con mayor PML", nodo_top)

    # ─── TABLA RESUMEN ───
    st.markdown("### 📋 Tabla resumen por nodo")
    summary = calcular_resumen(df)
    if not summary.empty:
        st.dataframe(
            summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "nodo": st.column_config.TextColumn("Clave", width="small"),
                "nombre": st.column_config.TextColumn("Nombre", width="medium"),
                "ccr": st.column_config.TextColumn("CCR", width="small"),
                "registros": st.column_config.NumberColumn("# Reg", format="%d"),
                "promedio": st.column_config.NumberColumn("Promedio", format="$%.2f"),
                "mediana":  st.column_config.NumberColumn("Mediana", format="$%.2f"),
                "maximo":   st.column_config.NumberColumn("Máximo", format="$%.2f"),
                "minimo":   st.column_config.NumberColumn("Mínimo", format="$%.2f"),
                "std":      st.column_config.NumberColumn("Volatilidad", format="$%.2f"),
                "p95":      st.column_config.NumberColumn("P95", format="$%.2f"),
                "p5":       st.column_config.NumberColumn("P5", format="$%.2f"),
                "% horas neg": st.column_config.NumberColumn("% Neg", format="%.1f%%"),
            },
        )

    # ─── GRÁFICA DE LÍNEAS ───
    st.markdown("### 📈 Evolución temporal del PML")
    max_lineas = min(15, len(acumulado))
    fig_lineas = grafica_lineas_tiempo(df, max_nodos=max_lineas)
    if fig_lineas:
        st.plotly_chart(fig_lineas, use_container_width=True)

    # ─── HEATMAP CON SELECTOR ───
    st.markdown("### 🔥 Heatmap horario × mensual")
    nodos_disponibles = sorted(df["nodo"].unique())
    nodo_heatmap = st.selectbox(
        "Selecciona un nodo para ver su patrón hora × mes:",
        nodos_disponibles,
        index=0,
    )
    fig_heat = grafica_heatmap_horario(df, nodo_heatmap)
    if fig_heat:
        st.plotly_chart(fig_heat, use_container_width=True)

    # ─── TOPS COMPARATIVOS ───
    st.markdown("### 🏆 Rankings comparativos")
    tab1, tab2, tab3 = st.tabs(["Mayor PML", "Mayor volatilidad", "Más horas negativas"])
    with tab1:
        fig_top1 = grafica_barras_top(df, "promedio", top_n=min(10, len(acumulado)))
        if fig_top1:
            st.plotly_chart(fig_top1, use_container_width=True)
    with tab2:
        fig_top2 = grafica_barras_top(df, "volatilidad", top_n=min(10, len(acumulado)))
        if fig_top2:
            st.plotly_chart(fig_top2, use_container_width=True)
    with tab3:
        fig_top3 = grafica_barras_top(df, "negativos", top_n=min(10, len(acumulado)))
        if fig_top3:
            st.plotly_chart(fig_top3, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════
# UI PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>⚡ CENACE PML Analyzer 
        <span class="srf-badge">SRF</span>
    </h1>
    <p>Descarga y análisis de Precios Marginales Locales del SEN — Recurrent Energy / Canadian Solar</p>
</div>
""", unsafe_allow_html=True)

# Cargar catálogo
catalogo = cargar_catalogo()
if catalogo:
    st.info(f"📚 Catálogo CENACE cargado: **{len(catalogo):,} nodos** disponibles · "
            f"versión 2026-04-23")
else:
    st.warning("⚠️ Catálogo no cargado. Revisa que `data/catalogo_nodos.xlsx` esté en el repo.")

# ─────── SIDEBAR: PARÁMETROS ───────
with st.sidebar:
    st.markdown(f"### ⚙️ Parámetros de consulta")

    # Nodos
    st.markdown("**📍 Nodos CENACE**")
    nodos_input = st.text_area(
        "Lista de claves (separadas por coma o salto de línea)",
        height=150,
        placeholder="01VAJ-230\n01XAL-230\n06FUN-115",
        help="Pega las claves de los nodos que quieres consultar. Ejemplo: 01VAJ-230, 01XAL-230"
    )

    # Sistema y proceso
    col1, col2 = st.columns(2)
    with col1:
        sistema = st.selectbox("Sistema", ["SIN", "BCA", "BCS"], index=0)
    with col2:
        proceso = st.selectbox("Proceso", ["MTR", "MDA"], index=0)

    # Fechas
    st.markdown("**📅 Período**")
    col_f1, col_f2 = st.columns(2)
    today = date.today()
    default_ini = today - timedelta(days=90)
    with col_f1:
        f_ini = st.date_input("Desde", value=default_ini, max_value=today)
    with col_f2:
        f_fin = st.date_input("Hasta", value=today - timedelta(days=1), max_value=today)

    # Workers
    st.markdown("**⚡ Performance**")
    max_workers = st.slider(
        "Workers paralelos", min_value=3, max_value=12, value=8,
        help="Más workers = más rápido pero arriesga rate limit (HTTP 429)"
    )

    st.divider()
    st.caption("Sebastian Roldan (SRF)\nRecurrent Energy · Canadian Solar")


# ─────── ÁREA PRINCIPAL: PROCESAMIENTO ───────
col_left, col_right = st.columns([2, 1])

with col_left:
    st.markdown("### 🚀 Ejecutar consulta")

    # Procesar input de nodos
    nodos = []
    if nodos_input.strip():
        nodos_raw = re.split(r"[,\n]+", nodos_input.strip())
        nodos = [n.strip().upper() for n in nodos_raw if n.strip()]
        # Quitar duplicados manteniendo orden
        seen = set()
        nodos = [n for n in nodos if not (n in seen or seen.add(n))]

    # Validaciones
    nodos_validos    = [n for n in nodos if n in catalogo] if catalogo else nodos
    nodos_invalidos  = [n for n in nodos if catalogo and n not in catalogo]

    if nodos:
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Nodos solicitados", len(nodos))
        col_b.metric("✅ En catálogo", len(nodos_validos))
        col_c.metric("⚠️ Desconocidos", len(nodos_invalidos))

        if nodos_invalidos:
            with st.expander(f"⚠️ {len(nodos_invalidos)} nodos no encontrados en catálogo"):
                st.code("\n".join(nodos_invalidos), language=None)
                st.caption("Estos nodos se consultarán igual, pero verifica que las claves sean correctas.")

    # Estimación
    if nodos and f_ini < f_fin:
        n_dias = (f_fin - f_ini).days + 1
        n_bloques = (n_dias + BLOQUE_MAX - 1) // BLOQUE_MAX
        n_lotes = (len(nodos) + 9) // 10
        n_consultas = n_bloques * n_lotes
        tiempo_seg = n_consultas * 1.5 / max_workers
        st.caption(f"📊 **{n_dias} días · {n_consultas} consultas · ~{tiempo_seg/60:.1f} min** "
                   f"(con {max_workers} workers)")

    boton = st.button("⚡ Descargar datos", type="primary", disabled=(len(nodos) == 0))

with col_right:
    st.markdown("### 📋 Instrucciones")
    st.markdown("""
    1. Pega las claves de los nodos en el panel izquierdo
    2. Selecciona sistema, proceso y fechas
    3. Ajusta el slider de workers (8 es óptimo)
    4. Click en **Descargar datos**
    5. Cuando termine, descarga el Excel generado
    """)
    st.info("💡 **Tip**: Para 50+ nodos en 3 meses, planea unos 3-5 minutos de descarga.")


# ─────── EJECUCIÓN ───────
if boton and nodos and f_ini < f_fin:
    fecha_ini = f_ini.strftime("%Y/%m/%d")
    fecha_fin = f_fin.strftime("%Y/%m/%d")

    progress = st.progress(0, text="Iniciando descarga...")
    status = st.empty()
    t0 = time.time()

    def cb(done, total):
        pct = done / total
        progress.progress(pct, text=f"Descargando: {done}/{total} consultas ({pct*100:.0f}%)")

    try:
        with st.spinner("Consultando CENACE..."):
            acumulado, errores = descargar_pml(
                nodos, fecha_ini, fecha_fin, sistema, proceso,
                max_workers=max_workers, progress_cb=cb
            )

        elapsed = time.time() - t0
        progress.progress(1.0, text=f"✅ Descarga completa en {elapsed:.0f}s")

        if not acumulado:
            st.error("❌ No se recibieron datos. Revisa las claves de nodos y el rango de fechas.")
        else:
            n_total = sum(len(f) for f in acumulado.values())
            col_x, col_y, col_z = st.columns(3)
            col_x.metric("Nodos con datos", f"{len(acumulado)}/{len(nodos)}")
            col_y.metric("Total registros", f"{n_total:,}")
            col_z.metric("Tiempo", f"{elapsed:.0f}s")

            if errores:
                with st.expander(f"⚠️ {len(errores)} errores en consultas"):
                    err_df = pd.DataFrame(errores[:20])
                    st.dataframe(err_df, use_container_width=True)

            st.markdown("### 📥 Descargar Excel")
            with st.spinner("Generando Excel..."):
                excel_bytes = generar_excel_datos(acumulado, sistema, proceso,
                                                    fecha_ini, fecha_fin)

            ts = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"PML_CENACE_{sistema}_{ts}.xlsx"

            st.download_button(
                label=f"📊 Descargar {filename}",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # Vista previa rápida
            with st.expander("👁️ Vista previa de datos crudos (primeras filas)"):
                for i, (nodo, filas) in enumerate(list(acumulado.items())[:3]):
                    st.markdown(f"**{nodo}** — {len(filas):,} registros")
                    df_prev = pd.DataFrame(filas[:10])
                    st.dataframe(df_prev, use_container_width=True)

            # ─── DASHBOARD ANALÍTICO v2 ───
            render_dashboard(acumulado, catalogo)

    except Exception as e:
        st.error(f"❌ Error inesperado: {type(e).__name__}: {str(e)}")
        st.exception(e)


# ─────── FOOTER ───────
st.divider()
st.caption(f"⚡ CENACE PML Analyzer · Sebastian Roldan (SRF) · Recurrent Energy / Canadian Solar")

v2 — Dashboard analítico con Plotly
