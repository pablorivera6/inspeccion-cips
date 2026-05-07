"""
PCC Integrity — Dashboard Unificado de Inspecciones
PAP · DCVG · CIPS
"""
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import msal
import requests
import io
import base64
import os
import sys
import tempfile
import datetime
import zipfile
import xml.etree.ElementTree as _ET
import gc

st.set_page_config(
    page_title="PCC – Inspecciones",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Tokens de color ────────────────────────────────────────────────────────────
C_RED   = "#D50032"
C_PROT  = "#374151"   # gris oscuro — institucional
C_SOBRE = "#6B7280"   # gris medio
C_SIN   = "#D50032"   # rojo — sin protección es crítico
C_DCVG  = "#374151"
C_GRAY  = "#F3F3F3"
C_LINE  = "#E0E0E0"

ESTADO_COLORS = {
    "Protegido":      "#374151",
    "Sobreprotegido": "#6B7280",
    "Sin protección": "#D50032",
    "Sin medición":   "#CBD5E1",
}

# ── Tokens semánticos CIPS ─────────────────────────────────────────────────────
CIPS_OK    = "#16A34A"   # Protegido
CIPS_WARN  = "#D97706"   # Sobreprotegido (precaución)
CIPS_CRIT  = "#D50032"   # Desprotegido (crítico)
CIPS_NONE  = "#475569"   # Sin dato

# Superficies modo oscuro CIPS
CIPS_BG    = "#0F172A"
CIPS_SURF1 = "#1E293B"
CIPS_SURF2 = "#273549"
CIPS_BORD  = "#334155"
CIPS_TXT1  = "#F1F5F9"
CIPS_TXT2  = "#94A3B8"

# ── CSS unificado ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* ── Design tokens ────────────────────────────────────────────────────── */
  :root {
    --brand:        #B91C1C;   /* chrome de marca: logo, botones          */
    --data-ok:      #16A34A;   /* PROTEGIDO (dato, nunca chrome)           */
    --data-warn:    #D97706;   /* SOBREPROTEGIDO                           */
    --data-crit:    #DC2626;   /* DESPROTEGIDO                             */
    --data-none:    #94A3B8;   /* sin dato                                 */
    --surf-0:       #F8FAFC;   /* fondo principal                          */
    --surf-1:       #FFFFFF;   /* cards                                    */
    --surf-2:       #F1F5F9;   /* sidebar, inputs, table header            */
    --bord:         #E2E8F0;   /* divisores                                */
    --text-1:       #0F172A;   /* títulos                                  */
    --text-2:       #475569;   /* labels                                   */
    --text-3:       #94A3B8;   /* metadatos                                */

    /* Escala tipográfica: 6 pasos, ratio 1.15 */
    --text-xs:   0.6875rem;  /* 11px */
    --text-sm:   0.8125rem;  /* 13px */
    --text-base: 0.875rem;   /* 14px */
    --text-md:   1rem;       /* 16px */
    --text-lg:   1.375rem;   /* 22px */
    --text-xl:   2rem;       /* 32px */

    /* Espaciado (escala 4px) */
    --sp-1: 4px;   --sp-2: 8px;  --sp-3: 12px;
    --sp-4: 16px;  --sp-6: 24px; --sp-8: 32px;
  }

  /* ── Animaciones: solo 3 ─────────────────────────────────────────────── */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  @keyframes barGrow {
    from { transform: scaleX(0); transform-origin: left; }
    to   { transform: scaleX(1); transform-origin: left; }
  }
  @keyframes shimmer {
    0%   { background-position: -400px 0; }
    100% { background-position:  400px 0; }
  }

  /* ── Base ────────────────────────────────────────────────────────────── */
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
  .stApp { background: var(--surf-0); }
  .block-container { padding: 1.8rem 2.5rem 1rem !important; max-width: 1600px; }

  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: var(--surf-2); }
  ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

  /* ── Sidebar ─────────────────────────────────────────────────────────── */
  [data-testid="stSidebar"] > div:first-child {
    background: var(--surf-1);
    border-right: 1px solid var(--bord);
  }
  [data-testid="stSidebar"] * { color: var(--text-1) !important; }
  [data-testid="stSidebar"] hr { border-color: var(--bord); }
  [data-testid="stSidebar"] [data-testid="stRadio"] label {
    font-size: var(--text-base) !important;
    font-weight: 500 !important;
  }
  [data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: var(--surf-2) !important; border-radius: 8px !important;
  }

  /* ── Títulos de sección ──────────────────────────────────────────────── */
  .pbi-title {
    font-size: var(--text-md); font-weight: 600; color: var(--text-1);
    margin: var(--sp-8) 0 var(--sp-4) 0;
  }

  /* ── Header principal (PAP/DCVG) ─────────────────────────────────────── */
  .main-header {
    background: var(--surf-1);
    padding: var(--sp-4) var(--sp-6);
    border-radius: 10px;
    border: 1px solid var(--bord);
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: var(--sp-6);
  }
  .main-header-title {
    font-size: var(--text-md); font-weight: 700; color: var(--text-1);
    margin: 0; display: flex; align-items: center; gap: var(--sp-2);
  }
  .main-header-meta {
    font-size: var(--text-sm); color: var(--text-2); font-weight: 500;
  }

  /* ── Stat cards (PAP/DCVG) ───────────────────────────────────────────── */
  .stat-container {
    background: var(--surf-1); border: 1px solid var(--bord);
    border-radius: 10px; padding: var(--sp-4);
    transition: box-shadow 0.15s ease;
  }
  .stat-container:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.07); }
  .stat-label {
    font-size: var(--text-xs); text-transform: uppercase; font-weight: 600;
    color: var(--text-3); letter-spacing: 0.08em; margin-bottom: var(--sp-2);
  }
  .stat-val {
    font-size: var(--text-lg); font-weight: 700; color: var(--text-1);
  }

  /* ── Estado list ─────────────────────────────────────────────────────── */
  .estado-item {
    display: flex; align-items: center; gap: var(--sp-2);
    font-size: var(--text-base); padding: 3px 0;
    transition: background 0.15s ease; border-radius: 6px;
  }
  .estado-item:hover { background: var(--surf-2); }
  .dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }

  /* ── Footer ──────────────────────────────────────────────────────────── */
  .pcc-footer {
    background: var(--brand);
    color: white; padding: var(--sp-4) var(--sp-6);
    margin-top: var(--sp-8); border-radius: 10px;
    display: flex; align-items: center; gap: var(--sp-4);
  }
  .pcc-footer-logo { font-size: var(--text-md); font-weight: 700; }
  .pcc-footer-text { font-size: var(--text-sm); opacity: 0.8; }

  /* ── Separadores ─────────────────────────────────────────────────────── */
  .sec-div {
    height: 1px; background: var(--bord);
    margin: var(--sp-6) 0 var(--sp-4) 0;
  }

  /* ── Bloques CIPS individuales ───────────────────────────────────────── */
  .bloque {
    background: var(--surf-1); border-radius: 10px;
    border: 1px solid var(--bord); padding: var(--sp-4);
    margin-bottom: var(--sp-4);
  }
  .bloque-titulo {
    font-weight: 600; font-size: var(--text-xs); color: var(--text-2);
    text-transform: uppercase; letter-spacing: 0.1em;
    margin-bottom: var(--sp-4); padding-bottom: var(--sp-2);
    border-bottom: 1px solid var(--bord);
  }

  /* ── Métricas Streamlit ──────────────────────────────────────────────── */
  [data-testid="stMetric"] {
    background: var(--surf-1); border-radius: 10px;
    border: 1px solid var(--bord);
    padding: var(--sp-4) !important;
  }
  [data-testid="stMetricLabel"] {
    font-size: var(--text-xs) !important; color: var(--text-3) !important;
    font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.08em;
  }
  [data-testid="stMetricValue"] {
    font-size: var(--text-lg) !important; font-weight: 700 !important;
    color: var(--text-1) !important;
  }

  /* ── Botones ─────────────────────────────────────────────────────────── */
  .stButton > button {
    background: var(--brand) !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important;
    font-size: var(--text-sm) !important;
    transition: opacity 0.15s ease !important;
  }
  .stButton > button:hover { opacity: 0.88 !important; }
  .stButton > button:active { opacity: 0.75 !important; }
  [data-testid="stDownloadButton"] > button {
    background: #166534 !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important;
    font-size: var(--text-sm) !important;
    transition: opacity 0.15s ease !important;
  }
  [data-testid="stDownloadButton"] > button:hover { opacity: 0.88 !important; }

  /* ── DataFrames ──────────────────────────────────────────────────────── */
  [data-testid="stDataFrame"] {
    border-radius: 10px; border: 1px solid var(--bord); overflow: hidden;
  }

  /* ── Plotly charts ───────────────────────────────────────────────────── */
  .js-plotly-plot { border-radius: 10px; }

  /* ── Expanders ───────────────────────────────────────────────────────── */
  [data-testid="stExpander"] {
    border-radius: 8px !important; border: 1px solid var(--bord) !important;
    overflow: hidden;
  }

  /* ── Shimmer (skeleton para sync SP) ─────────────────────────────────── */
  .sp-badge-shimmer {
    background: linear-gradient(90deg, #F0FFF4 25%, #dcffe8 50%, #F0FFF4 75%);
    background-size: 400px 100%;
    animation: shimmer 1.8s infinite;
  }

  #MainMenu, footer { visibility: hidden; }
  [data-testid="stToolbar"] { visibility: hidden; }

  /* ── CIPS: vista comparativa ─────────────────────────────────────────── */
  .cips-kpi-card {
    background: var(--surf-1);
    border: 1px solid var(--bord);
    border-radius: 10px;
    padding: var(--sp-4);
    transition: box-shadow 0.15s ease;
    height: 100%;
  }
  .cips-kpi-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.07); }

  .cips-label {
    font-size: var(--text-xs); font-weight: 600; color: var(--text-3);
    text-transform: uppercase; letter-spacing: 0.12em;
    margin-bottom: var(--sp-2);
  }
  .cips-value {
    font-size: var(--text-lg); font-weight: 700; color: var(--text-1);
    font-variant-numeric: tabular-nums; line-height: 1.1;
    animation: fadeUp 0.3s ease-out both;
  }
  .cips-sub {
    font-size: var(--text-xs); color: var(--text-3);
    margin-top: var(--sp-1); line-height: 1.4;
  }

  .cips-section-title {
    font-size: var(--text-xs); font-weight: 700; color: var(--text-2);
    text-transform: uppercase; letter-spacing: 0.16em;
    margin: var(--sp-8) 0 var(--sp-4);
  }

  .cips-bar-row { margin-bottom: var(--sp-6); }
  .cips-bar-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: var(--sp-2); gap: var(--sp-2);
  }
  .cips-bar-name {
    font-size: var(--text-sm); font-weight: 600; color: var(--text-1);
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  .cips-bar-right { display: flex; align-items: center; gap: var(--sp-2); flex-shrink: 0; }
  .cips-bar-score {
    font-size: var(--text-sm); font-weight: 700;
    font-variant-numeric: tabular-nums; min-width: 24px; text-align: right;
  }
  .cips-bar-track {
    height: 10px; border-radius: 5px; background: var(--surf-2);
    overflow: hidden; display: flex;
  }
  .cips-bar-seg {
    height: 100%;
    animation: barGrow 0.65s ease-out both;
  }
  .cips-bar-meta {
    font-size: var(--text-xs); color: var(--text-3);
    margin-top: var(--sp-1);
  }
  .cips-bar-legend {
    display: flex; gap: var(--sp-6); margin-top: var(--sp-4);
    font-size: var(--text-xs); color: var(--text-3); flex-wrap: wrap;
  }
  .cips-bar-legend span { display: flex; align-items: center; gap: var(--sp-1); }
  .cips-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex-shrink: 0; }

  .badge-critico  { display:inline-block; background:#FEE2E2; color:#991B1B; padding:2px 8px; border-radius:20px; font-size:var(--text-xs); font-weight:700; letter-spacing:0.04em; white-space:nowrap; }
  .badge-moderado { display:inline-block; background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:20px; font-size:var(--text-xs); font-weight:700; letter-spacing:0.04em; white-space:nowrap; }
  .badge-bajo     { display:inline-block; background:#DCFCE7; color:#166534; padding:2px 8px; border-radius:20px; font-size:var(--text-xs); font-weight:700; letter-spacing:0.04em; white-space:nowrap; }

  .cips-table { width:100%; border-collapse:separate; border-spacing:0; font-size:var(--text-sm); }
  .cips-table thead th {
    padding: var(--sp-2) var(--sp-3);
    font-size: var(--text-xs); font-weight: 700; color: var(--text-3);
    text-transform: uppercase; letter-spacing: 0.1em;
    border-bottom: 2px solid var(--bord);
    background: var(--surf-2);
    position: sticky; top: 0;
  }
  .cips-table thead th:first-child { border-radius: 8px 0 0 0; }
  .cips-table thead th:last-child  { border-radius: 0 8px 0 0; }
  .cips-table tbody tr { transition: background 0.15s ease; }
  .cips-table tbody tr:hover td { background: var(--surf-2); }
  .cips-table tbody td { padding: var(--sp-3) var(--sp-3); border-bottom: 1px solid var(--surf-2); color: var(--text-2); vertical-align: middle; }
  .cips-table .num { text-align: right; font-variant-numeric: tabular-nums; font-weight: 500; }
  .cips-table .center { text-align: center; }
  .cips-table tbody tr:last-child td { border-bottom: none; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CIPS — Colores y carga de datos procesados
# ══════════════════════════════════════════════════════════════════════════════

CIPS_COLORS = {
    "PROTEGIDO":      "#374151",   # gris oscuro — estado OK
    "DESPROTEGIDO":   "#D50032",   # rojo institucional — crítico
    "SOBREPROTEGIDO": "#7F1D1D",   # rojo oscuro — advertencia
}


def _estado_cp(v):
    if pd.isna(v):   return "DESPROTEGIDO"
    if v <= -1200:   return "SOBREPROTEGIDO"
    if v <= -850:    return "PROTEGIDO"
    return "DESPROTEGIDO"


def _meta_from_name(nombre):
    """Extrae tramo y fecha del nombre del archivo."""
    base = nombre.replace(".xlsx", "")
    parts = base.split("_")
    if len(parts) >= 3 and parts[0].upper() == "CIPS":
        fecha_raw = parts[-2] if len(parts) > 2 else ""
        try:
            fecha = datetime.datetime.strptime(fecha_raw, "%Y%m%d").strftime("%d/%m/%Y")
            tramo = " ".join(parts[1:-2])
            return tramo.replace("_", " "), fecha
        except Exception:
            pass
    return base.replace("_", " "), "—"


def _finalizar_df(df):
    """Normaliza voltajes, calcula Estado_CP y limpia coordenadas."""
    for col in ["On_mV_limpio", "Off_mV_limpio"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            if df[col].dropna().abs().median() < 5:
                df[col] = df[col] * 1000
    if "Estado_CP" not in df.columns and "Off_mV_limpio" in df.columns:
        df["Estado_CP"] = df["Off_mV_limpio"].apply(_estado_cp)
    for c in ["Lat_corr", "Long_corr"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # Filtrar coordenadas fuera del rango de Colombia
    if "Lat_corr" in df.columns:
        df.loc[~df["Lat_corr"].between(-5, 15), "Lat_corr"] = pd.NA
    if "Long_corr" in df.columns:
        df.loc[~df["Long_corr"].between(-82, -65), "Long_corr"] = pd.NA
    return df


def load_cips_processed(file, categoria="ACTUAL"):
    """Carga un Excel CIPS — soporta formato FastField (Survey Data) e Histórico (CIPS)."""
    nombre = getattr(file, "name", str(file))
    tramo, fecha = _meta_from_name(nombre)

    xl = pd.ExcelFile(file)
    getattr(file, "seek", lambda x: None)(0)

    # ── Formato histórico: hoja "CIPS" ───────────────────────────────────────
    if "CIPS" in xl.sheet_names and "Survey Data" not in xl.sheet_names:
        # Auto-detectar fila de cabecera: probar header=0 y 1
        _seek = getattr(file, "seek", lambda x: None)
        _seek(0)
        df0 = pd.read_excel(file, sheet_name="CIPS", header=0, nrows=1)
        _seek(0)
        header_row = 0 if "KILÓMETRO" in df0.columns else 1
        df = pd.read_excel(file, sheet_name="CIPS", header=header_row)
        # Todas las variantes conocidas de nombres de columna
        RENAME_H = {
            "KILÓMETRO":                  "PK_geom_m",
            "Von [V/CSE]":                "On_mV_limpio",
            "Voff [V/CSE]":               "Off_mV_limpio",
            "POTENCIAL ON [VCSE]":        "On_mV_limpio",
            "POTENCIAL INSTANT OFF [VCSE]":"Off_mV_limpio",
            "LATITUD":                    "Lat_corr",
            "LONGITUD":                   "Long_corr",
            "ALTITUD":                    "Altitud",
        }
        df = df.rename(columns={k: v for k, v in RENAME_H.items() if k in df.columns})
        if "PK_geom_m" in df.columns:
            df["PK_geom_m"] = pd.to_numeric(df["PK_geom_m"], errors="coerce") * 1000
        df = _finalizar_df(df)
        df = df.dropna(subset=["PK_geom_m", "Off_mV_limpio"], how="all")
        if fecha == "—":
            fecha = "Histórico"
        return {"df": df, "tramo": tramo, "fecha": fecha,
                "filename": nombre, "tipo": "CIPS", "categoria": categoria}

    # ── Formato FastField: hoja "Survey Data" ─────────────────────────────────
    getattr(file, "seek", lambda x: None)(0)
    df = pd.read_excel(file, sheet_name="Survey Data")
    RENAME_FF = {
        "Dist From Start":          "PK_geom_m",
        "On Voltage":               "On_mV_limpio",
        "Off Voltage":              "Off_mV_limpio",
        "Latitude":                 "Lat_corr",
        "Longitude":                "Long_corr",
        "Comment":                  "Comentario",
        "DCP/Feature/DCVG Anomaly": "Anomalia",
    }
    for src, dst in RENAME_FF.items():
        if src in df.columns and dst not in df.columns:
            df = df.rename(columns={src: dst})
    df = _finalizar_df(df)
    return {"df": df, "tramo": tramo, "fecha": fecha,
            "filename": nombre, "tipo": "CIPS", "categoria": categoria}


# ══════════════════════════════════════════════════════════════════════════════
# PAP / DCVG — Helpers
# ══════════════════════════════════════════════════════════════════════════════

def parse_abscisa(val):
    if pd.isna(val): return None
    s = str(val).strip().replace(" ", "")
    if "+" in s:
        try:
            km, m = s.split("+"); return float(int(km) * 1000 + int(m))
        except: return None
    try: return round(float(s) * 1000, 1)
    except: return None


def parse_gps(val):
    if pd.isna(val): return None, None
    try:
        p = str(val).split(","); return float(p[0].strip()), float(p[1].strip())
    except: return None, None


def clean_zeros(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").replace(0.0, float("nan"))
    return df


def calc_estado(v):
    if pd.isna(v): return "Sin medición"
    if v < -1200:  return "Sobreprotegido"
    if v <= -850:  return "Protegido"
    return "Sin protección"


def load_excel(file, pdf_urls=None):
    if pdf_urls is None: pdf_urls = {}
    xl   = pd.ExcelFile(file)
    root = pd.read_excel(xl, "Root")
    fc   = "Fecha " if "Fecha " in root.columns else "Fecha"
    meta = {
        "inspector": str(root.get("Personal", pd.Series(["—"])).iloc[0]),
        "cargo":     str(root.get("Cargo",    pd.Series(["—"])).iloc[0]),
        "fecha":     str(root[fc].iloc[0]) if fc in root.columns else "—",
        "form_name": str(root.get("Form Name", pd.Series(["—"])).iloc[0]),
    }
    tipo  = "DCVG" if "DCVG" in meta["form_name"].upper() else "PAP"
    sheet = "subform_1" if "subform_1" in xl.sheet_names else xl.sheet_names[-1]
    df    = pd.read_excel(xl, sheet).dropna(how="all")

    if "Localizacion GPS" in df.columns:
        c = df["Localizacion GPS"].apply(parse_gps)
        df["Latitud"]  = c.apply(lambda x: x[0])
        df["Longitud"] = c.apply(lambda x: x[1])

    if "Abscisa" in df.columns:
        df["Abscisa_m"] = df["Abscisa"].apply(parse_abscisa)

    tramo = "Sin tramo"
    if "Tramo" in df.columns:
        v = df["Tramo"].dropna().value_counts()
        if not v.empty: tramo = str(v.index[0])
    meta["tramo"] = tramo

    if tipo == "PAP":
        df = clean_zeros(df, ["On [mV]","Off [mV]","IR ON-OFF [mV]",
                               "Voltaje AC","Potencial Natural [mV]",
                               "Resistencia entre NEG1-NEG2 [ohm]"])
        df["Estado"] = df["Off [mV]"].apply(calc_estado)

    submission_id = None
    for s_name in xl.sheet_names:
        try:
            tmp = pd.read_excel(xl, s_name, nrows=2)
            if "Submission Id" in tmp.columns:
                submission_id = str(tmp["Submission Id"].dropna().iloc[0]); break
            elif "Receipt" in tmp.columns:
                submission_id = str(tmp["Receipt"].dropna().iloc[0]); break
        except Exception:
            pass

    pdf_url = None
    if submission_id:
        for p_name, p_url in pdf_urls.items():
            if submission_id in p_name:
                pdf_url = p_url; break

    return {"meta": meta, "df": df, "tipo": tipo, "pdf_url": pdf_url}


CHART = dict(
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(t=40, b=40, l=50, r=20),
    font=dict(size=12, family="'Inter', sans-serif", color="#475569"),
    transition=dict(duration=600, easing="cubic-in-out"),
)

def apply_chart(fig, h=280, xl="Abscisa (m)", yl="", title=""):
    fig.update_layout(
        **CHART, height=h,
        title=dict(text=title, font_size=14, font_color="#0F172A",
                   x=0, xanchor="left", pad=dict(l=0)) if title else {},
        xaxis_title=dict(text=xl, font=dict(size=11, color="#64748B")),
        yaxis_title=dict(text=yl, font=dict(size=11, color="#64748B")),
        legend=dict(orientation="h", y=-0.25, font_size=11, font_color="#475569"),
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Inter",
                        bordercolor="#E2E8F0"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False,
                     linecolor="#E2E8F0", linewidth=1)
    fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False,
                     linecolor="#E2E8F0", linewidth=1)
    # Animación de trazos al dibujar
    fig.update_traces(selector=dict(mode="lines+markers"),
                      line=dict(shape="spline"),
                      opacity=0.95)
    return fig

def _render_pdf(pdf_url, filename="reporte.pdf"):
    """Descarga el PDF y lo muestra embebido usando PDF.js para máxima compatibilidad."""
    try:
        with st.spinner("Cargando reporte..."):
            resp = requests.get(pdf_url, timeout=25)
        if not resp.ok:
            st.link_button("Abrir reporte", pdf_url, use_container_width=True)
            return
        b64 = base64.b64encode(resp.content).decode()
        col_dl, _ = st.columns([1, 4])
        with col_dl:
            st.download_button("Descargar reporte", data=resp.content,
                               file_name=filename, mime="application/pdf",
                               use_container_width=True)
        components.html(f"""
        <!DOCTYPE html>
        <html>
        <head>
          <style>
            body {{ margin:0; background:#e8eaed; }}
            #scroll-box {{
              width:100%; height:860px; overflow-y:auto;
              border-radius:10px; box-shadow:0 2px 12px rgba(0,0,0,0.1);
            }}
            .page-wrap {{
              display:flex; justify-content:center;
              padding:12px 0; background:#e8eaed;
            }}
            canvas {{
              box-shadow:0 2px 8px rgba(0,0,0,0.2);
              border-radius:2px;
              max-width:98%;
            }}
          </style>
        </head>
        <body>
          <div id="scroll-box"></div>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
          <script>
            pdfjsLib.GlobalWorkerOptions.workerSrc =
              'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

            var b64 = "{b64}";
            var bin = atob(b64);
            var arr = new Uint8Array(bin.length);
            for(var i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);

            pdfjsLib.getDocument({{data:arr}}).promise.then(function(doc) {{
              var box = document.getElementById('scroll-box');
              var total = doc.numPages;
              for(var n=1; n<=total; n++) {{
                (function(num) {{
                  doc.getPage(num).then(function(page) {{
                    var vp = page.getViewport({{scale:1.6}});
                    var canvas = document.createElement('canvas');
                    canvas.height = vp.height;
                    canvas.width  = vp.width;
                    var wrap = document.createElement('div');
                    wrap.className = 'page-wrap';
                    wrap.appendChild(canvas);
                    box.appendChild(wrap);
                    page.render({{canvasContext:canvas.getContext('2d'), viewport:vp}});
                  }});
                }})(n);
              }}
            }});
          </script>
        </body>
        </html>
        """, height=900)
    except Exception:
        st.link_button("Abrir reporte", pdf_url, use_container_width=True)


def pbi_title(text):
    st.markdown(f'<p class="pbi-title">{text}</p>', unsafe_allow_html=True)

def animated_kpi_row(items):
    """
    items: list de (label, value, color)
    Renderiza KPI cards con contador JS animado.
    """
    cards_html = ""
    js_counters = ""
    for i, (label, value, color) in enumerate(items):
        uid = f"kpi_{i}_{abs(hash(label)) % 9999}"
        # Solo animar si el valor es numérico entero
        if isinstance(value, int):
            display = f'<span id="{uid}">0</span>'
            js_counters += f"""
            animateCounter("{uid}", {value}, 900, {i * 120});
            """
        else:
            display = f'<span id="{uid}">{value}</span>'

        cards_html += f"""
        <div style="background:white;border:1px solid #E2E8F0;border-radius:10px;
                    padding:1rem 1.1rem;
                    animation:fadeUp 0.3s ease-out {i*0.06:.2f}s both;">
          <p style="font-size:var(--text-xs);text-transform:uppercase;font-weight:600;
                    color:#94A3B8;letter-spacing:0.1em;margin:0 0 6px 0;">{label}</p>
          <p style="font-size:var(--text-lg);font-weight:700;color:{color};margin:0;
                    font-variant-numeric:tabular-nums;">
            {display}
          </p>
        </div>
        """

    full_html = f"""
    <style>
      @keyframes fadeUp {{
        from {{ opacity:0; transform:translateY(10px); }}
        to   {{ opacity:1; transform:translateY(0); }}
      }}
    </style>
    <div style="display:grid;grid-template-columns:repeat({len(items)},1fr);gap:0.75rem;margin-bottom:1rem;">
      {cards_html}
    </div>
    """
    components.html(full_html, height=100)

def footer():
    st.markdown("""
    <div class="pcc-footer">
      <span class="pcc-footer-logo">Protección Catódica de Colombia</span>
      <span class="pcc-footer-text">· Inspección de Corriente Impresa</span>
    </div>
    """, unsafe_allow_html=True)

def divider():
    st.markdown('<div class="sec-div"></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAP Dashboard
# ══════════════════════════════════════════════════════════════════════════════

def render_pap(d):
    df   = d["df"].sort_values("Abscisa_m", na_position="last").copy()
    meta = d["meta"]
    total   = len(df)
    n_prot  = int((df["Estado"] == "Protegido").sum())
    n_sobre = int((df["Estado"] == "Sobreprotegido").sum())
    n_sin   = int((df["Estado"] == "Sin protección").sum())
    n_no    = int((df["Estado"] == "Sin medición").sum())

    st.markdown(f"""
    <div class="main-header">
      <div class="main-header-title">
        Inspección PAP <span style="color:#64748B;font-weight:400;margin-left:0.4rem;">| {meta['tramo']}</span>
      </div>
      <div class="main-header-meta">
        {meta['inspector']} • {meta['cargo']} • {meta['fecha']} • <b style="color:#0F172A;">{total} puntos</b>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col_tbl, col_map, col_right = st.columns([1.1, 2.0, 0.9])

    with col_tbl:
        pbi_title("Abscisa · P Off mV · P On mV · Observaciones")
        show = ["Abscisa","Off [mV]","On [mV]","Observaciones"]
        t = df[[c for c in show if c in df.columns]].copy()
        t.columns = [c.replace(" [mV]","") for c in t.columns]
        st.dataframe(t.reset_index(drop=True), use_container_width=True,
                     height=340, hide_index=True)

    with col_map:
        pbi_title("Distribución geográfica")
        mdf = df.dropna(subset=["Latitud","Longitud"])
        if not mdf.empty:
            fig = px.scatter_mapbox(
                mdf, lat="Latitud", lon="Longitud",
                color="Estado", color_discrete_map=ESTADO_COLORS,
                hover_data={"Abscisa":True,"Off [mV]":True,"On [mV]":True,
                             "Latitud":False,"Longitud":False},
                zoom=10, height=340, mapbox_style="open-street-map",
            )
            fig.update_traces(marker=dict(size=8, opacity=0.9))
            fig.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                               legend=dict(x=0.01,y=0.99,
                                           bgcolor="rgba(255,255,255,0.88)",
                                           borderwidth=1, font_size=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin GPS")

    with col_right:
        pbi_title("Estado de protección")
        st.markdown(f"""
        <div class="stat-container" style="margin-bottom:1rem;">
          <p class="stat-label">ESTADO DE PROTECCIÓN</p>
          <div class="estado-item">
            <span class="dot" style="background:{C_PROT};"></span>
            <span style="color:#334155;font-weight:500;">Protegido</span>
            <span style="margin-left:auto;font-weight:700;color:#0F172A;">{n_prot}</span>
          </div>
          <div class="estado-item">
            <span class="dot" style="background:{C_SOBRE};"></span>
            <span style="color:#334155;font-weight:500;">Sobreprotegido</span>
            <span style="margin-left:auto;font-weight:700;color:#0F172A;">{n_sobre}</span>
          </div>
          <div class="estado-item">
            <span class="dot" style="background:{C_SIN};"></span>
            <span style="color:#334155;font-weight:500;">Sin protección</span>
            <span style="margin-left:auto;font-weight:700;color:#0F172A;">{n_sin}</span>
          </div>
          <div class="estado-item" style="opacity:0.7;">
            <span class="dot" style="background:#BDBDBD;"></span>
            <span style="color:#64748B;font-weight:500;">Sin medición</span>
            <span style="margin-left:auto;font-weight:600;color:#64748B;">{n_no}</span>
          </div>
        </div>
        """, unsafe_allow_html=True)
        edo = df["Estado"].value_counts().reset_index()
        edo.columns = ["Estado","Count"]
        fig = px.pie(edo, values="Count", names="Estado",
                     color="Estado", color_discrete_map=ESTADO_COLORS, hole=0.5)
        fig.update_layout(height=200, margin=dict(t=0,b=0,l=0,r=0),
                           paper_bgcolor="white", showlegend=False)
        fig.update_traces(textposition="outside", textinfo="percent",
                           hovertemplate="%{label}<br>%{value} pts<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)

    divider()
    pbi_title("P On mV y P Off mV por Abscisa")
    pot = df.dropna(subset=["Abscisa_m"]).copy()
    pot = pot[pot["Off [mV]"].notna() | pot["On [mV]"].notna()] if "Off [mV]" in pot else pot
    if not pot.empty:
        fig = go.Figure()
        if "On [mV]" in pot.columns:
            fig.add_trace(go.Scatter(x=pot["Abscisa_m"], y=pot["On [mV]"],
                mode="lines+markers", name="P On mV",
                line=dict(color="#64B5F6", width=1.8), marker=dict(size=4, color="#64B5F6")))
        if "Off [mV]" in pot.columns:
            fig.add_trace(go.Scatter(x=pot["Abscisa_m"], y=pot["Off [mV]"],
                mode="lines+markers", name="P Off mV",
                line=dict(color=C_SOBRE, width=1.8), marker=dict(size=5, color=C_SOBRE)))
        fig.add_hrect(y0=-1200, y1=-850, fillcolor="rgba(21,101,192,0.05)", line_width=0)
        fig.add_hline(y=-850,  line=dict(color="#64B5F6", dash="dash", width=1.2),
                      annotation_text="-850", annotation_position="top left",
                      annotation_font=dict(size=9, color="#64B5F6"))
        fig.add_hline(y=-1200, line=dict(color="#EF5350", dash="dash", width=1.2),
                      annotation_text="-1200", annotation_position="top left",
                      annotation_font=dict(size=9, color="#EF5350"))
        st.plotly_chart(apply_chart(fig, 280, "Abscisa (m)", "mV"), use_container_width=True)

    has_ir = "IR ON-OFF [mV]" in df.columns and df["IR ON-OFF [mV]"].notna().any()
    has_ac = "Voltaje AC"     in df.columns and df["Voltaje AC"].notna().any()
    if has_ir or has_ac:
        divider()
        col_ir, col_ac = st.columns(2) if (has_ir and has_ac) else (st.container(), st.container())
        if has_ir:
            with col_ir:
                pbi_title("IR ON-OFF [mV] por Abscisa")
                sub = df.dropna(subset=["Abscisa_m","IR ON-OFF [mV]"])
                fig = go.Figure(go.Scatter(x=sub["Abscisa_m"], y=sub["IR ON-OFF [mV]"],
                    mode="lines+markers", name="IR ON-OFF",
                    line=dict(color="#5C6BC0", width=1.8), marker=dict(size=4)))
                st.plotly_chart(apply_chart(fig, 240, "Abscisa (m)", "mV"),
                                use_container_width=True)
        if has_ac:
            with col_ac:
                pbi_title("Voltaje AC por Abscisa")
                sub = df.dropna(subset=["Abscisa_m","Voltaje AC"])
                fig = go.Figure(go.Scatter(x=sub["Abscisa_m"], y=sub["Voltaje AC"],
                    mode="lines+markers", name="Voltaje AC",
                    line=dict(color="#7B1FA2", width=1.8), marker=dict(size=4)))
                st.plotly_chart(apply_chart(fig, 240, "Abscisa (m)", "V"),
                                use_container_width=True)

    divider()
    pbi_title("Mediciones eléctricas")
    cols_med = ["Abscisa","Off [mV]","On [mV]","Potencial Natural [mV]",
                "IR ON-OFF [mV]","Voltaje AC","Resistencia entre NEG1-NEG2 [ohm]",
                "Estado","Observaciones"]
    t = df[[c for c in cols_med if c in df.columns]].reset_index(drop=True)
    st.dataframe(t, use_container_width=True, height=280, hide_index=True)

    divider()
    pbi_title("Estado de infraestructura")
    cols_inf = ["Abscisa","Tramo","Tipo de tramo","Estado Pintura",
                "Estado Conexiones","Estado Verticalidad","Tipo mantenimiento","Observaciones"]
    t = df[[c for c in cols_inf if c in df.columns]].reset_index(drop=True)
    st.dataframe(t, use_container_width=True, height=280, hide_index=True)

    # ── Reporte PDF ────────────────────────────────────────────────────────────
    if d.get("pdf_url"):
        divider()
        pbi_title("Reporte de inspección")
        _render_pdf(d["pdf_url"], "Reporte_PAP.pdf")

    footer()


# ══════════════════════════════════════════════════════════════════════════════
# DCVG Dashboard
# ══════════════════════════════════════════════════════════════════════════════

def render_dcvg(d):
    df   = d["df"].sort_values("Abscisa_m", na_position="last").copy()
    meta = d["meta"]
    porc = next((c for c in ["PORC_IR","% IR"] if c in df.columns), None)
    car  = next((c for c in ["Caracter ON-OFF","Carácter ON-OFF","Caracter ON_OFF"] if c in df.columns), None)
    if porc: df[porc] = pd.to_numeric(df[porc], errors="coerce").replace(0.0, float("nan"))
    total = len(df)

    st.markdown(f"""
    <div class="main-header">
      <div class="main-header-title">
        Inspección DCVG <span style="color:#64748B;font-weight:400;margin-left:0.4rem;">| {meta['tramo']}</span>
      </div>
      <div class="main-header-meta">
        {meta['inspector']} • {meta['cargo']} • {meta['fecha']} • <b style="color:#0F172A;">{total} puntos</b>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col_bar1, col_bar2, col_tbl, col_donut = st.columns([1, 1, 1.2, 1.2])

    with col_bar1:
        if car and df[car].notna().any():
            pbi_title(f"Recuento de {car}")
            cdf = df[car].value_counts().reset_index()
            cdf.columns = ["Caracter","Count"]
            fig = px.bar(cdf, x="Caracter", y="Count", color_discrete_sequence=[C_PROT])
            fig.update_layout(**CHART, height=240, showlegend=False, xaxis_title="", yaxis_title="Recuento")
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor="#F0F0F0")
            st.plotly_chart(fig, use_container_width=True)

    with col_bar2:
        if "Clasificacion" in df.columns and df["Clasificacion"].notna().any():
            pbi_title("Recuento por Clasificación")
            cls_bar = df["Clasificacion"].value_counts().reset_index()
            cls_bar.columns = ["Clasificacion","Count"]
            fig = px.bar(cls_bar, x="Clasificacion", y="Count", color_discrete_sequence=[C_PROT])
            fig.update_layout(**CHART, height=240, showlegend=False, xaxis_title="", yaxis_title="Recuento")
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor="#F0F0F0")
            st.plotly_chart(fig, use_container_width=True)

    with col_tbl:
        pbi_title("Carácter · Clasificación · PORC_IR")
        show = [car, "Clasificacion", porc]
        show = [c for c in show if c and c in df.columns]
        if show:
            st.dataframe(df[show].reset_index(drop=True),
                         use_container_width=True, height=240, hide_index=True)

    with col_donut:
        if "Clasificacion" in df.columns and df["Clasificacion"].notna().any():
            pbi_title("Distribución por Clasificación")
            cls_df = df["Clasificacion"].value_counts().reset_index()
            cls_df.columns = ["Clasificacion","Count"]
            fig = px.pie(cls_df, values="Count", names="Clasificacion",
                         hole=0.5, color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                               paper_bgcolor="white",
                               legend=dict(font_size=9, x=1, y=0.5))
            fig.update_traces(textposition="inside", textinfo="percent",
                               hovertemplate="%{label}<br>%{value}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)

    divider()
    col_ir, col_map = st.columns([1.4, 1])

    with col_ir:
        if porc and df[porc].notna().any():
            pbi_title("% IR (PORC_IR) por Abscisa")
            sub = df.dropna(subset=["Abscisa_m", porc])
            fig = go.Figure(go.Scatter(x=sub["Abscisa_m"], y=sub[porc],
                mode="lines+markers", name="% IR",
                line=dict(color=C_PROT, width=1.8), marker=dict(size=4)))
            for val, color, label in [
                (15, "#4CAF50", "15%"), (35, "#FF9800", "35%"), (60, "#F44336", "60%")
            ]:
                fig.add_hline(y=val, line=dict(color=color, dash="dash", width=1.2),
                              annotation_text=label, annotation_position="top right",
                              annotation_font=dict(size=9, color=color))
            ymax = max(100.0, float(sub[porc].max()) * 1.1)
            fig.update_layout(yaxis=dict(range=[0, ymax]))
            st.plotly_chart(apply_chart(fig, 300, "Abscisa (m)", "% IR"), use_container_width=True)

    with col_map:
        pbi_title("Mapa de puntos DCVG")
        mdf = df.dropna(subset=["Latitud","Longitud"])
        if not mdf.empty:
            fig = px.scatter_mapbox(
                mdf, lat="Latitud", lon="Longitud",
                color="Clasificacion" if "Clasificacion" in mdf.columns else None,
                hover_data={"Abscisa":True, **({porc:True} if porc else {}),
                             "Latitud":False,"Longitud":False},
                zoom=10, height=300, mapbox_style="open-street-map",
            )
            fig.update_traces(marker=dict(size=8, opacity=0.9))
            fig.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                               legend=dict(bgcolor="rgba(255,255,255,0.88)",
                                           borderwidth=1, font_size=9))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin GPS.")

    divider()
    pbi_title("Datos de medición DCVG")
    show = ["Abscisa", porc, car, "Clasificacion",
            "Potencial ON mV","Potencial OFF mV","P_RE mV","OL_RE mV","Comentarios"]
    show = [c for c in show if c and c in df.columns]
    st.dataframe(df[show].reset_index(drop=True), use_container_width=True, height=300, hide_index=True)

    # ── Reporte PDF ────────────────────────────────────────────────────────────
    if d.get("pdf_url"):
        divider()
        pbi_title("Reporte de inspección")
        _render_pdf(d["pdf_url"], "Reporte_DCVG.pdf")

    footer()


# ══════════════════════════════════════════════════════════════════════════════
# Resumen Global
# ══════════════════════════════════════════════════════════════════════════════

def render_resumen(inspecciones):
    pap  = [d for d in inspecciones if d["tipo"]=="PAP"]
    dcvg = [d for d in inspecciones if d["tipo"]=="DCVG"]
    tramos    = len({d["meta"]["tramo"] for d in inspecciones})
    pts_pap   = sum(len(d["df"]) for d in pap)
    pts_dcvg  = sum(len(d["df"]) for d in dcvg)

    st.markdown(f"""
    <div class="main-header">
      <div class="main-header-title">
        Resumen Global <span style="color:#64748B;font-weight:400;margin-left:0.4rem;">| Todas las Inspecciones</span>
      </div>
      <div class="main-header-meta">
        <b>{len(inspecciones)} archivos</b> • {tramos} tramos • <b style="color:#0F172A;">{pts_pap+pts_dcvg} puntos totales</b>
      </div>
    </div>
    """, unsafe_allow_html=True)

    animated_kpi_row([
        ("Archivos totales", len(inspecciones),      "#0F172A"),
        ("Tramos / Zonas",   tramos,                 C_RED),
        ("Archivos PAP",     len(pap),               C_PROT),
        ("Archivos DCVG",    len(dcvg),              C_DCVG),
        ("Puntos totales",   pts_pap + pts_dcvg,     "#0F172A"),
    ])

    divider()
    pbi_title("Mapa global — todos los puntos")
    frames = []
    for d in inspecciones:
        sub = d["df"].dropna(subset=["Latitud","Longitud"]).copy()
        if sub.empty: continue
        sub["_tipo"]   = d["tipo"]
        sub["_tramo"]  = d["meta"]["tramo"]
        sub["_insp"]   = d["meta"]["inspector"]
        sub["_estado"] = sub.get("Estado", "—")
        frames.append(sub)
    if frames:
        all_pts = pd.concat(frames, ignore_index=True)
        fig = px.scatter_mapbox(
            all_pts, lat="Latitud", lon="Longitud", color="_tramo",
            hover_data={"_tramo":True,"_tipo":True,"_estado":True,
                        "_insp":True,"Latitud":False,"Longitud":False},
            labels={"_tramo":"Tramo","_tipo":"Tipo","_estado":"Estado","_insp":"Inspector"},
            zoom=7, height=420, mapbox_style="open-street-map",
        )
        fig.update_traces(marker=dict(size=7, opacity=0.85))
        fig.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                          legend=dict(x=0.01,y=0.99,
                                      bgcolor="rgba(255,255,255,0.88)",
                                      bordercolor="#ddd",borderwidth=1))
        st.plotly_chart(fig, use_container_width=True)

    divider()
    pbi_title("Inspecciones cargadas")
    rows = []
    for d in sorted(inspecciones, key=lambda x:(x["meta"]["tramo"],x["meta"]["fecha"])):
        m=d["meta"]; df2=d["df"]
        if d["tipo"]=="PAP" and "Estado" in df2.columns:
            n_p=int((df2["Estado"]=="Protegido").sum())
            n_s=int((df2["Estado"]=="Sobreprotegido").sum())
            res=f"{n_p} Prot. / {n_s} Sobre."
        else: res="—"
        rows.append({"Tipo":d["tipo"],"Tramo":m["tramo"],
                     "Inspector":m["inspector"],"Fecha":m["fecha"],
                     "Puntos":len(df2),"Resumen":res})
    st.dataframe(pd.DataFrame(rows), use_container_width=True,
                 hide_index=True, height=min(60+len(rows)*35, 360))

    if pap:
        divider()
        pbi_title("Estado de protección por tramo (PAP)")
        data=[]
        for d in pap:
            if "Estado" not in d["df"].columns: continue
            t=d["meta"]["tramo"]
            for e,c in ESTADO_COLORS.items():
                cnt=int((d["df"]["Estado"]==e).sum())
                if cnt>0: data.append({"Tramo":t,"Estado":e,"Count":cnt})
        if data:
            fig=px.bar(pd.DataFrame(data),x="Tramo",y="Count",color="Estado",
                       color_discrete_map=ESTADO_COLORS,barmode="stack",height=250)
            fig.update_layout(**CHART, height=250, xaxis_title="", yaxis_title="Puntos",
                               legend=dict(orientation="h",y=-0.28))
            fig.update_xaxes(showgrid=False)
            st.plotly_chart(fig, use_container_width=True)

    footer()


# ══════════════════════════════════════════════════════════════════════════════
# CIPS — Vista comparativa Actual vs Histórico
# ══════════════════════════════════════════════════════════════════════════════

def _criticidad_stats(todos):
    """Calcula métricas de criticidad por tramo."""
    rows = []
    for d in todos:
        df2 = d["df"]
        total = len(df2)
        if total == 0 or "Estado_CP" not in df2.columns:
            continue
        n_prot  = int((df2["Estado_CP"] == "PROTEGIDO").sum())
        n_desp  = int((df2["Estado_CP"] == "DESPROTEGIDO").sum())
        n_sobre = int((df2["Estado_CP"] == "SOBREPROTEGIDO").sum())
        pct_prot  = n_prot  / total * 100
        pct_desp  = n_desp  / total * 100
        pct_sobre = n_sobre / total * 100
        # Score: desprotegido pesa x2 (mayor riesgo de corrosión), sobreprotegido x1
        score = min(100, pct_desp * 2 + pct_sobre)
        rows.append({
            "tramo":     d["tramo"],
            "categoria": d["categoria"],
            "fecha":     d["fecha"],
            "total":     total,
            "n_prot":    n_prot,
            "n_desp":    n_desp,
            "n_sobre":   n_sobre,
            "pct_prot":  round(pct_prot,  1),
            "pct_desp":  round(pct_desp,  1),
            "pct_sobre": round(pct_sobre, 1),
            "score":     round(score, 1),
        })
    return sorted(rows, key=lambda r: r["score"], reverse=True)


def _badge(score):
    if score >= 50: return '<span class="badge-critico">CRÍTICO</span>'
    if score >= 20: return '<span class="badge-moderado">MODERADO</span>'
    return '<span class="badge-bajo">BAJO</span>'


def render_cips_comparativo(actual_list, historico_list):
    todos  = actual_list + historico_list
    n_act  = sum(len(d["df"]) for d in actual_list)
    n_his  = sum(len(d["df"]) for d in historico_list)
    stats  = _criticidad_stats(todos)

    n_total_pts       = sum(r["total"] for r in stats) if stats else 0
    n_criticos        = sum(1 for r in stats if r["score"] >= 50)
    pct_fuera_global  = (sum(r["n_desp"] + r["n_sobre"] for r in stats)
                         / n_total_pts * 100) if n_total_pts else 0
    tramo_top         = stats[0]["tramo"]  if stats else "—"
    score_top         = stats[0]["score"] if stats else 0

    # ── Canvas oscuro ──────────────────────────────────────────────────────────
    # ── Situación global: bloque narrativo + barra compuesta ──────────────────
    pct_ok     = 100 - pct_fuera_global
    crit_color = CIPS_CRIT if n_criticos > 0 else CIPS_OK
    range_color= CIPS_CRIT if pct_fuera_global > 30 else (CIPS_WARN if pct_fuera_global > 10 else CIPS_OK)

    # Barra compuesta (protegido | sobreprotegido | desprotegido) global
    pct_sobre_global = (sum(r["n_sobre"] for r in stats) / n_total_pts * 100) if n_total_pts else 0
    pct_desp_global  = (sum(r["n_desp"]  for r in stats) / n_total_pts * 100) if n_total_pts else 0
    pct_prot_global  = 100 - pct_sobre_global - pct_desp_global

    statement = (f'<span style="color:{crit_color};font-weight:700;">'
                 f'{n_criticos} tramo{"s" if n_criticos!=1 else ""} fuera de criterio</span>'
                 if n_criticos > 0
                 else '<span style="color:#16A34A;font-weight:700;">todos los tramos en criterio</span>')

    st.markdown(f"""
    <div style="background:white;border:1px solid #E2E8F0;border-radius:10px;
                padding:1.2rem 1.6rem;margin-bottom:1.2rem;">
      <div style="display:flex;align-items:baseline;justify-content:space-between;
                  flex-wrap:wrap;gap:0.5rem;margin-bottom:0.9rem;">
        <div>
          <span style="font-size:var(--text-xs);color:#94A3B8;font-weight:600;
                       text-transform:uppercase;letter-spacing:0.12em;">
            PCC Integrity · CIPS &nbsp;·&nbsp; {len(stats)} tramos &nbsp;·&nbsp; {n_total_pts:,} pts
          </span>
          <div style="font-size:var(--text-md);font-weight:600;color:#0F172A;margin-top:3px;">
            {statement},
            <span style="color:{range_color};font-weight:700;">{pct_fuera_global:.1f}%</span>
            del total fuera de rango
          </div>
        </div>
        <div style="display:flex;gap:1rem;font-size:var(--text-xs);color:#94A3B8;flex-shrink:0;">
          <span><span style="color:#16A34A;font-weight:700;">{pct_prot_global:.0f}%</span> prot.</span>
          <span><span style="color:#D97706;font-weight:700;">{pct_sobre_global:.0f}%</span> sobre.</span>
          <span><span style="color:#DC2626;font-weight:700;">{pct_desp_global:.0f}%</span> desp.</span>
        </div>
      </div>
      <div style="height:8px;border-radius:4px;background:#F1F5F9;overflow:hidden;display:flex;">
        <div style="width:{pct_prot_global:.1f}%;background:{CIPS_OK};
                    animation:barGrow 0.6s ease-out both;"></div>
        <div style="width:{pct_sobre_global:.1f}%;background:{CIPS_WARN};
                    animation:barGrow 0.6s ease-out 0.1s both;"></div>
        <div style="width:{pct_desp_global:.1f}%;background:{CIPS_CRIT};
                    animation:barGrow 0.6s ease-out 0.2s both;"></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Layout dos columnas: ranking | mapa ────────────────────────────────────
    col_rank, col_map = st.columns([4, 5], gap="medium")

    with col_rank:
        st.markdown('<div class="cips-section-title">Ranking de criticidad</div>',
                    unsafe_allow_html=True)
        if stats:
            bars_html = ""
            for i, r in enumerate(stats):
                row_delay  = f"{0.15 + i * 0.08:.2f}s"
                seg_delay  = f"{0.25 + i * 0.08:.2f}s"
                badge      = _badge(r["score"])
                score_col  = CIPS_CRIT if r["score"] >= 50 else (CIPS_WARN if r["score"] >= 20 else CIPS_OK)
                bars_html += f"""
                <div class="cips-bar-row" style="animation-delay:{row_delay};">
                  <div class="cips-bar-header">
                    <span class="cips-bar-name">{r['tramo']}</span>
                    <span class="cips-bar-right">
                      {badge}
                      <span class="cips-bar-score" style="color:{score_col};">{r['score']:.0f}</span>
                    </span>
                  </div>
                  <div class="cips-bar-track">
                    <div class="cips-bar-seg" style="width:{r['pct_prot']:.1f}%;background:{CIPS_OK};animation-delay:{seg_delay};"></div>
                    <div class="cips-bar-seg" style="width:{r['pct_sobre']:.1f}%;background:{CIPS_WARN};animation-delay:{seg_delay};"></div>
                    <div class="cips-bar-seg" style="width:{r['pct_desp']:.1f}%;background:{CIPS_CRIT};animation-delay:{seg_delay};"></div>
                  </div>
                  <div class="cips-bar-meta">
                    {r['pct_prot']:.0f}% prot &nbsp;·&nbsp;
                    {r['pct_sobre']:.0f}% sobre &nbsp;·&nbsp;
                    {r['pct_desp']:.0f}% desp &nbsp;·&nbsp;
                    {r['total']:,} pts
                  </div>
                </div>"""
            bars_html += f"""
            <div class="cips-bar-legend">
              <span><span class="cips-dot" style="background:{CIPS_OK};"></span>Protegido</span>
              <span><span class="cips-dot" style="background:{CIPS_WARN};"></span>Sobreprotegido</span>
              <span><span class="cips-dot" style="background:{CIPS_CRIT};"></span>Desprotegido</span>
              <span style="margin-left:auto;font-variant-numeric:tabular-nums;">
                Score = %desp × 2 + %sobre
              </span>
            </div>"""
            st.markdown(bars_html, unsafe_allow_html=True)

    with col_map:
        st.markdown('<div class="cips-section-title">Distribución geográfica</div>',
                    unsafe_allow_html=True)
        frames = []
        for d in todos:
            if "Lat_corr" not in d["df"].columns: continue
            sub = d["df"].dropna(subset=["Lat_corr","Long_corr"]).copy()
            if sub.empty: continue
            if "Estado_CP" not in sub.columns: sub["Estado_CP"] = "—"
            sub["_tramo"] = d["tramo"]
            if len(sub) > 3000:
                sub = sub.iloc[::max(1, len(sub)//3000)]
            frames.append(sub[["Lat_corr","Long_corr","Estado_CP","_tramo"]])
        if frames:
            all_pts = pd.concat(frames, ignore_index=True)
            # Etiquetas legibles para la leyenda
            LABEL_MAP = {"PROTEGIDO":"Protegido","SOBREPROTEGIDO":"Sobreprotegido",
                         "DESPROTEGIDO":"Desprotegido","—":"Sin dato"}
            all_pts["Estado"] = all_pts["Estado_CP"].map(LABEL_MAP).fillna("Sin dato")
            fig_map = px.scatter_mapbox(
                all_pts, lat="Lat_corr", lon="Long_corr",
                color="Estado",
                color_discrete_map={
                    "Protegido":      CIPS_OK,
                    "Sobreprotegido": CIPS_WARN,
                    "Desprotegido":   CIPS_CRIT,
                    "Sin dato":       CIPS_NONE,
                },
                hover_data={"_tramo": True, "Lat_corr": False, "Long_corr": False,
                            "Estado": False},
                zoom=6, height=440, mapbox_style="open-street-map",
                category_orders={"Estado": ["Desprotegido","Sobreprotegido","Protegido","Sin dato"]},
                labels={"Estado": "Estado CP", "_tramo": "Tramo"},
            )
            fig_map.update_traces(marker=dict(size=5, opacity=0.9))
            fig_map.update_layout(
                margin=dict(t=0, b=0, l=0, r=0),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(x=0.01, y=0.99,
                            bgcolor="rgba(255,255,255,0.95)", font_color="#1E293B",
                            bordercolor="#E2E8F0", borderwidth=1, font_size=11),
            )
            st.plotly_chart(fig_map, use_container_width=True)
        else:
            st.markdown('<p style="color:#475569;font-size:var(--text-base);">Sin coordenadas GPS disponibles.</p>',
                        unsafe_allow_html=True)

    # ── Perfil Off mV ──────────────────────────────────────────────────────────
    pk_key      = lambda d: next((c for c in ["PK_geom_m","PK_real_m"] if c in d["df"].columns), None)
    tiene_chart = any(pk_key(d) and "Off_mV_limpio" in d["df"].columns for d in todos)
    if tiene_chart:
        st.markdown('<div class="cips-section-title">Perfil Off mV por PK — todos los tramos</div>',
                    unsafe_allow_html=True)
        fig = go.Figure()

        # Paletas diferenciadas por estilo de línea + color
        DASH_STYLES = ["dot", "dash", "dashdot", "longdash", "longdashdot"]
        GRAY_PALETTE = ["#94A3B8","#64748B","#475569","#CBD5E1","#334155"]
        RED_PALETTE  = ["#D50032","#F87171","#9B0022","#EF4444","#B91C1C"]

        for i, d in enumerate(historico_list):
            pk = pk_key(d)
            if not pk or "Off_mV_limpio" not in d["df"].columns: continue
            sub = d["df"].dropna(subset=[pk]).sort_values(pk)
            if len(sub) > 2000: sub = sub.iloc[::max(1, len(sub)//2000)]
            r   = next((s for s in stats if s["tramo"] == d["tramo"]), None)
            col = (CIPS_CRIT if r and r["score"] >= 50
                   else CIPS_WARN if r and r["score"] >= 20
                   else GRAY_PALETTE[i % len(GRAY_PALETTE)])
            fig.add_trace(go.Scatter(
                x=sub[pk], y=sub["Off_mV_limpio"],
                mode="lines", name=d["tramo"][:24],
                line=dict(color=col, width=1.6, dash=DASH_STYLES[i % len(DASH_STYLES)]),
                opacity=0.85))

        for i, d in enumerate(actual_list):
            pk = pk_key(d)
            if not pk or "Off_mV_limpio" not in d["df"].columns: continue
            sub = d["df"].dropna(subset=[pk]).sort_values(pk)
            if len(sub) > 2000: sub = sub.iloc[::max(1, len(sub)//2000)]
            fig.add_trace(go.Scatter(
                x=sub[pk], y=sub["Off_mV_limpio"],
                mode="lines", name=f"[ACT] {d['tramo'][:20]}",
                line=dict(color=RED_PALETTE[i % len(RED_PALETTE)], width=2.2)))

        # Zonas de criterio
        fig.add_hrect(y0=-1200, y1=-850,
                      fillcolor="rgba(22,163,74,0.06)", line_width=0,
                      annotation_text="Zona protegida",
                      annotation_position="top left",
                      annotation_font=dict(size=9, color="#16A34A"))
        fig.add_hline(y=-850,
                      line=dict(color="#16A34A", dash="dash", width=1.2),
                      annotation_text="-850 mV",
                      annotation_position="top right",
                      annotation_font=dict(size=9, color="#16A34A"))
        fig.add_hline(y=-1200,
                      line=dict(color=CIPS_WARN, dash="dash", width=1.2),
                      annotation_text="-1 200 mV",
                      annotation_position="bottom right",
                      annotation_font=dict(size=9, color=CIPS_WARN))

        fig.update_layout(
            height=380,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(t=30, b=60, l=55, r=20),
            font=dict(size=11, family="'Inter', sans-serif", color="#475569"),
            xaxis_title=dict(text="PK (m)", font=dict(size=11, color="#64748B")),
            yaxis_title=dict(text="Off mV", font=dict(size=11, color="#64748B")),
            legend=dict(orientation="h", y=-0.3, font_size=10,
                        bgcolor="rgba(0,0,0,0)", font_color="#64748B"),
            hovermode="x unified",
            hoverlabel=dict(bgcolor="white", font_size=12,
                            font_family="Inter", bordercolor="#E2E8F0"),
        )
        fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False,
                         tickfont=dict(color="#94A3B8"), linecolor="#E2E8F0")
        fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False,
                         tickfont=dict(color="#94A3B8"))
        st.plotly_chart(fig, use_container_width=True)

    # ── Tabla detallada ────────────────────────────────────────────────────────
    if stats:
        st.markdown('<div class="cips-section-title">Detalle por tramo</div>',
                    unsafe_allow_html=True)
        rows_html = """
        <div style="overflow-x:auto;border:1px solid #E2E8F0;border-radius:12px;
                    box-shadow:0 1px 4px rgba(0,0,0,0.04);">
        <table class="cips-table">
          <thead>
            <tr>
              <th style="text-align:left;">Tramo</th>
              <th class="num">Puntos</th>
              <th class="num" style="color:#16A34A;">% Prot.</th>
              <th class="num" style="color:#D97706;">% Sobre.</th>
              <th class="num" style="color:#D50032;">% Desp.</th>
              <th class="num">Score</th>
              <th class="center">Nivel</th>
            </tr>
          </thead>
          <tbody>"""
        for i, r in enumerate(stats):
            score_col  = CIPS_CRIT if r["score"] >= 50 else (CIPS_WARN if r["score"] >= 20 else CIPS_OK)
            row_delay  = f"{0.1 + i * 0.07:.2f}s"
            rows_html += f"""
            <tr style="animation-delay:{row_delay};">
              <td style="font-weight:600;color:#0F172A;">{r['tramo']}</td>
              <td class="num" style="color:#64748B;">{r['total']:,}</td>
              <td class="num" style="color:#16A34A;">{r['pct_prot']:.1f}%</td>
              <td class="num" style="color:#D97706;">{r['pct_sobre']:.1f}%</td>
              <td class="num" style="color:#D50032;">{r['pct_desp']:.1f}%</td>
              <td class="num" style="font-weight:800;color:{score_col};">{r['score']:.0f}</td>
              <td class="center">{_badge(r['score'])}</td>
            </tr>"""
        rows_html += "</tbody></table></div>"
        st.markdown(rows_html, unsafe_allow_html=True)



# ══════════════════════════════════════════════════════════════════════════════
# CIPS — Dashboard histórico (estilo Power BI)
# ══════════════════════════════════════════════════════════════════════════════

def render_cips_dashboard(d):
    df_raw = d["df"].copy()
    tramo  = d["tramo"]
    fecha  = d["fecha"]
    total  = len(df_raw)

    pk_col = next((c for c in ["PK_geom_m","PK_real_m"] if c in df_raw.columns), None)
    if pk_col:
        df_raw[pk_col] = pd.to_numeric(df_raw[pk_col], errors="coerce")

    n_prot  = int((df_raw["Estado_CP"]=="PROTEGIDO").sum())      if "Estado_CP" in df_raw.columns else 0
    n_desp  = int((df_raw["Estado_CP"]=="DESPROTEGIDO").sum())   if "Estado_CP" in df_raw.columns else 0
    n_sobre = int((df_raw["Estado_CP"]=="SOBREPROTEGIDO").sum()) if "Estado_CP" in df_raw.columns else 0
    pct_p   = f"{n_prot/total*100:.1f}%" if total else "—"
    pct_d   = f"{n_desp/total*100:.1f}%" if total else "—"

    # ── Header ─────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:white;border:1px solid #E2E8F0;border-left:5px solid #D50032;
                padding:1.2rem 1.8rem;border-radius:12px;margin-bottom:1.2rem;
                box-shadow:0 4px 16px -4px rgba(0,0,0,0.06);
                display:flex;align-items:center;justify-content:space-between;">
      <div>
        <div style="font-size:var(--text-sm);color:#D50032;font-weight:700;
                    text-transform:uppercase;letter-spacing:0.12em;margin-bottom:4px;">
          Inspección CIPS
        </div>
        <div style="font-size:var(--text-lg);font-weight:800;color:#0F172A;letter-spacing:-0.02em;">
          {tramo}
        </div>
        <div style="font-size:var(--text-base);color:#64748B;margin-top:4px;">
          {fecha} &nbsp;·&nbsp; {total:,} puntos medidos
        </div>
      </div>
      <div style="display:flex;gap:1.2rem;text-align:center;">
        <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:0.65rem 1.1rem;">
          <div style="font-size:var(--text-lg);font-weight:800;color:#374151;">{pct_p}</div>
          <div style="font-size:var(--text-xs);color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-top:2px;">Protegido</div>
        </div>
        <div style="background:#FFF5F6;border:1px solid #FECDD3;border-radius:10px;padding:0.65rem 1.1rem;">
          <div style="font-size:var(--text-lg);font-weight:800;color:#D50032;">{pct_d}</div>
          <div style="font-size:var(--text-xs);color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-top:2px;">Desprotegido</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── KPI row ────────────────────────────────────────────────────────────────
    animated_kpi_row([
        ("Total puntos",     total,   "#0F172A"),
        ("Protegidos",       n_prot,  "#374151"),
        ("Desprotegidos",    n_desp,  "#D50032"),
        ("Sobreprotegidos",  n_sobre, "#7F1D1D"),
    ])

    # ── Leer filtros ANTES de renderizar tabla y mapa ──────────────────────────
    col_tbl, col_map, col_right = st.columns([1.0, 1.9, 0.95])

    with col_right:
        st.markdown("""
        <div style="background:white;border:1px solid #E2E8F0;border-radius:12px;
                    padding:1.2rem;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
          <p style="font-size:var(--text-sm);text-transform:uppercase;font-weight:700;
                    color:#64748B;letter-spacing:0.08em;margin:0 0 0.8rem 0;">
            Estado CP
          </p>
        """, unsafe_allow_html=True)

        if "Estado_CP" in df_raw.columns:
            estados_disp = sorted(df_raw["Estado_CP"].dropna().unique().tolist())
            est_sel = st.multiselect("Estado CP", estados_disp, default=estados_disp,
                                     label_visibility="collapsed",
                                     key=f"cips_est_{d['filename']}")
        else:
            est_sel = []

        st.markdown("""
          <p style="font-size:var(--text-sm);text-transform:uppercase;font-weight:700;
                    color:#64748B;letter-spacing:0.08em;margin:0.9rem 0 0.4rem 0;">
            Rango PK (m)
          </p>
        """, unsafe_allow_html=True)

        pk_min = float(df_raw[pk_col].min()) if pk_col and df_raw[pk_col].notna().any() else 0.0
        pk_max = float(df_raw[pk_col].max()) if pk_col and df_raw[pk_col].notna().any() else 1.0
        if pk_col and pk_max > pk_min:
            pk_sel = st.slider("PK", pk_min, pk_max, (pk_min, pk_max),
                               format="%.0f", label_visibility="collapsed",
                               key=f"cips_pk_{d['filename']}")
        else:
            pk_sel = (pk_min, pk_max)

        st.markdown("</div>", unsafe_allow_html=True)

        # Donut
        st.markdown("<div style='margin-top:0.8rem;'>", unsafe_allow_html=True)
        if "Estado_CP" in df_raw.columns:
            est_df = df_raw["Estado_CP"].value_counts().reset_index()
            est_df.columns = ["Estado_CP","Count"]
            fig = px.pie(est_df, values="Count", names="Estado_CP",
                         color="Estado_CP", color_discrete_map=CIPS_COLORS, hole=0.55)
            fig.update_layout(height=220, margin=dict(t=4,b=0,l=0,r=0),
                               paper_bgcolor="white",
                               showlegend=True,
                               legend=dict(font_size=10, orientation="h",
                                           y=-0.15, x=0.5, xanchor="center"))
            fig.update_traces(textposition="inside", textinfo="percent",
                               hovertemplate="%{label}<br>%{value} pts<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Aplicar filtros ────────────────────────────────────────────────────────
    df = df_raw.copy()
    if est_sel and "Estado_CP" in df.columns:
        df = df[df["Estado_CP"].isin(est_sel)]
    if pk_col:
        df = df[(df[pk_col] >= pk_sel[0]) & (df[pk_col] <= pk_sel[1])]

    # ── Tabla ──────────────────────────────────────────────────────────────────
    with col_tbl:
        n_fil = len(df)
        st.markdown(f"""
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.5rem;">
          <p class="pbi-title" style="margin:0;">Datos de medición</p>
          <span style="font-size:var(--text-sm);color:#64748B;font-weight:600;
                       background:#F1F5F9;padding:2px 10px;border-radius:20px;">
            {n_fil:,} pts
          </span>
        </div>
        """, unsafe_allow_html=True)

        col_labels = {
            pk_col:              "PK (m)",
            "On_mV_limpio":      "On mV",
            "Off_mV_limpio":     "Off mV",
            "IR_Drop_mV_limpio": "IR Drop mV",
            "Estado_CP":         "Estado",
        }
        show = [c for c in [pk_col, "On_mV_limpio", "Off_mV_limpio", "IR_Drop_mV_limpio", "Estado_CP"]
                if c and c in df.columns]
        tbl = df[show].rename(columns={k:v for k,v in col_labels.items() if k in show}).reset_index(drop=True)

        def color_estado(val):
            c = CIPS_COLORS.get(str(val), "")
            return f"color:{c};font-weight:700;" if c else ""

        try:
            if "Estado" in tbl.columns:
                styled = tbl.style.map(color_estado, subset=["Estado"])
            else:
                styled = tbl.style
            st.dataframe(styled, use_container_width=True, height=410, hide_index=True)
        except Exception:
            st.dataframe(tbl, use_container_width=True, height=410, hide_index=True)

    # ── Mapa ───────────────────────────────────────────────────────────────────
    with col_map:
        pbi_title("Trayectoria GPS")
        lat_c = "Lat_corr" if "Lat_corr" in df.columns else ("Latitude" if "Latitude" in df.columns else None)
        lon_c = "Long_corr" if "Long_corr" in df.columns else ("Longitude" if "Longitude" in df.columns else None)
        mdf   = df.dropna(subset=[lat_c, lon_c]) if lat_c and lon_c else pd.DataFrame()

        if not mdf.empty:
            hover_d = {lat_c: False, lon_c: False}
            if pk_col and pk_col in mdf.columns:        hover_d[pk_col]         = True
            if "Off_mV_limpio" in mdf.columns: hover_d["Off_mV_limpio"] = True
            if "Estado_CP" in mdf.columns:     hover_d["Estado_CP"]     = True

            fig = px.scatter_mapbox(
                mdf, lat=lat_c, lon=lon_c,
                color="Estado_CP" if "Estado_CP" in mdf.columns else None,
                color_discrete_map=CIPS_COLORS,
                hover_data=hover_d,
                zoom=10, height=410, mapbox_style="open-street-map",
            )
            fig.update_traces(marker=dict(size=5, opacity=0.92))
            fig.update_layout(
                margin=dict(t=0,b=0,l=0,r=0),
                legend=dict(x=0.01, y=0.99, bgcolor="rgba(255,255,255,0.92)",
                            borderwidth=1, font_size=10, bordercolor="#E2E8F0")
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.markdown("""
            <div style="height:410px;display:flex;align-items:center;justify-content:center;
                        background:#F8FAFC;border-radius:10px;border:1px dashed #CBD5E1;">
              <div style="text-align:center;color:#94A3B8;">
                <div style="font-size:2rem;margin-bottom:0.5rem;">📍</div>
                <div style="font-size:var(--text-base);font-weight:500;">Sin coordenadas GPS</div>
                <div style="font-size:var(--text-sm);margin-top:4px;">El archivo no tiene columnas Latitude/Longitude</div>
              </div>
            </div>""", unsafe_allow_html=True)

    # ── Gráfica On/Off vs PK ───────────────────────────────────────────────────
    if pk_col and ("On_mV_limpio" in df.columns or "Off_mV_limpio" in df.columns):
        divider()
        sub = df.dropna(subset=[pk_col]).sort_values(pk_col)
        n_sub = len(sub)

        # Reducir puntos para visualización si hay demasiados
        if n_sub > 3000:
            step = max(1, n_sub // 3000)
            sub  = sub.iloc[::step]

        pbi_title(f"On mV y Off mV por PK  ·  {n_sub:,} lecturas")
        fig = go.Figure()

        if "On_mV_limpio" in sub.columns:
            fig.add_trace(go.Scatter(
                x=sub[pk_col], y=sub["On_mV_limpio"],
                mode="lines", name="On mV",
                line=dict(color="#9CA3AF", width=1.4),
                fill="tozeroy", fillcolor="rgba(156,163,175,0.05)"))

        if "Off_mV_limpio" in sub.columns:
            fig.add_trace(go.Scatter(
                x=sub[pk_col], y=sub["Off_mV_limpio"],
                mode="lines", name="Off mV",
                line=dict(color="#D50032", width=2.0)))

        # Zona protegida sombreada
        fig.add_hrect(y0=-1200, y1=-850,
                      fillcolor="rgba(55,65,81,0.05)", line_width=0,
                      annotation_text="Zona protegida",
                      annotation_position="top left",
                      annotation_font=dict(size=9, color="#374151"))

        fig.add_hline(y=-850,
                      line=dict(color="#6B7280", dash="dash", width=1.3),
                      annotation_text="-850 mV",
                      annotation_position="top right",
                      annotation_font=dict(size=9, color="#6B7280"))
        fig.add_hline(y=-1200,
                      line=dict(color="#D50032", dash="dash", width=1.3),
                      annotation_text="-1.200 mV",
                      annotation_position="bottom right",
                      annotation_font=dict(size=9, color="#D50032"))

        fig.update_layout(
            **CHART, height=340,
            xaxis_title=dict(text="PK (m)", font=dict(size=11, color="#64748B")),
            yaxis_title=dict(text="mV",     font=dict(size=11, color="#64748B")),
            legend=dict(orientation="h", y=-0.2, font_size=12),
            hoverlabel=dict(bgcolor="white", font_size=12),
            hovermode="x unified",
        )
        fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
        fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
        st.plotly_chart(fig, use_container_width=True)

    footer()


# ══════════════════════════════════════════════════════════════════════════════
# SharePoint — carga automática de PAP/DCVG
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=600, show_spinner="Sincronizando con SharePoint...")
def fetch_sharepoint_files():
    if "sharepoint" not in st.secrets:
        return [], {}
    cfg = st.secrets["sharepoint"]
    client_id     = cfg.get("client_id")
    client_secret = cfg.get("client_secret")
    tenant_id     = cfg.get("tenant_id")
    tenant_name   = cfg.get("tenant_name")
    site_url      = cfg.get("site_url")
    folder_path   = cfg.get("folder_path")
    if not folder_path:
        return [], {}
    try:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app_obj   = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        result = app_obj.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            st.sidebar.error("Error autenticando con SharePoint.")
            return [], {}
        token   = result["access_token"]
        headers = {"Authorization": f"Bearer {token}"}
        hostname = f"{tenant_name}.sharepoint.com"
        site_path = site_url.replace(f"https://{hostname}", "")
        resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers)
        if not resp.ok:
            return [], {}
        site_id = resp.json().get("id")
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children",
            headers=headers
        )
        if not resp.ok:
            return [], {}
        files_data = resp.json().get("value", [])
        pdf_urls, downloaded = {}, []
        for item in files_data:
            name  = item.get("name", "")
            d_url = item.get("@microsoft.graph.downloadUrl")
            if not d_url or name.startswith("~"): continue
            if name.endswith(".pdf"):
                pdf_urls[name] = d_url
            elif name.endswith(".xlsx"):
                f_resp = requests.get(d_url)
                if f_resp.ok:
                    f_obj = io.BytesIO(f_resp.content)
                    f_obj.name = name
                    downloaded.append(f_obj)
        if downloaded:
            st.sidebar.success(f"{len(downloaded)} archivos cargados de SharePoint")
        return downloaded, pdf_urls
    except Exception as e:
        st.sidebar.error(f"Error SharePoint: {e}")
        return [], {}

def _repair_xlsx(data: bytes) -> io.BytesIO:
    """
    Parchea xlsx con XML inválido o con valores fuera de rango en styles.xml.
    Siempre reemplaza styles.xml con una versión mínima válida para evitar
    errores de openpyxl como 'Max value is 14' en font.family.
    """
    import re as _re
    _MINIMAL_STYLES = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    )
    def _clamp_family(b: bytes) -> bytes:
        """Clampea font family > 14 → 2; preserva todos los strings."""
        txt = b.decode("utf-8", errors="replace")
        return _re.sub(
            r'family val="(\d+)"',
            lambda m: f'family val="{min(int(m.group(1)), 14)}"',
            txt,
        ).encode("utf-8")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zin:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename == "xl/styles.xml":
                        content = _MINIMAL_STYLES
                    elif item.filename == "xl/sharedStrings.xml":
                        content = _clamp_family(content)
                    zout.writestr(item, content)
            buf.seek(0)
            return buf
    except Exception:
        return io.BytesIO(data)


@st.cache_data(ttl=600, show_spinner="Buscando archivos en SharePoint...")
def fetch_cips_metadata():
    """Solo obtiene lista de archivos (nombre + URL). No descarga datos."""
    if "sharepoint" not in st.secrets:
        return [], [], ["No hay configuración de SharePoint"]
    cfg = st.secrets["sharepoint"]
    errors = []
    try:
        app_obj = msal.ConfidentialClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
            client_credential=cfg["client_secret"],
        )
        token = app_obj.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        ).get("access_token")
        if not token:
            return [], [], ["No se pudo autenticar con SharePoint"]
        headers  = {"Authorization": f"Bearer {token}"}
        hostname = f"{cfg['tenant_name']}.sharepoint.com"
        site_path = cfg["site_url"].replace(f"https://{hostname}", "")
        site_id = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}",
            headers=headers
        ).json().get("id")
        if not site_id:
            return [], [], ["Site SharePoint no encontrado"]

        def _list(folder, categoria):
            resp = requests.get(
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder}:/children",
                headers=headers
            ).json()
            if "error" in resp:
                errors.append(f"{folder}: {resp['error'].get('message','')}")
                return []
            return [
                {"name": it["name"], "url": it["@microsoft.graph.downloadUrl"],
                 "size": it.get("size", 0), "categoria": categoria}
                for it in resp.get("value", [])
                if it.get("name","").endswith(".xlsx")
                   and not it.get("name","").startswith("~")
                   and it.get("@microsoft.graph.downloadUrl")
            ]

        actual_f = cfg.get("cips_actual_folder",     "Inspecciones Ocensa/CIPS ACTUAL")
        hist_f   = cfg.get("cips_historicos_folder", "Inspecciones Ocensa/CIPS HISTORICOS")
        return _list(actual_f, "ACTUAL"), _list(hist_f, "HISTÓRICO"), errors
    except Exception as e:
        return [], [], [str(e)]


def _cips_cache():
    if "cips_files" not in st.session_state:
        st.session_state.cips_files = {}
    return st.session_state.cips_files


def _load_one_cips(meta: dict):
    """Descarga y parsea un archivo CIPS. Caché en session_state por nombre."""
    cache = _cips_cache()
    name  = meta["name"]
    if name in cache:
        return cache[name], None
    try:
        r = requests.get(meta["url"], timeout=60)
        if not r.ok:
            return None, f"HTTP {r.status_code}"
        raw = r.content
        del r; gc.collect()
        f = _repair_xlsx(raw)
        del raw; gc.collect()
        f.name = name
        d = load_cips_processed(f, meta["categoria"])
        del f; gc.collect()
        cache[name] = d
        return d, None
    except MemoryError:
        gc.collect()
        return None, "Archivo demasiado grande"
    except Exception as e:
        return None, str(e)[:80]


def _sp_token():
    cfg = st.secrets.get("sharepoint", {})
    app_obj = msal.ConfidentialClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        client_credential=cfg["client_secret"],
    )
    return app_obj.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    ).get("access_token")


# ══════════════════════════════════════════════════════════════════════════════
# Sidebar
# ══════════════════════════════════════════════════════════════════════════════

def sidebar():
    with st.sidebar:
        # Logo
        if os.path.exists("logo-pcc-hd.png"):
            with open("logo-pcc-hd.png", "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            st.markdown(f'''
            <div style="padding:1rem 0 0 0;text-align:left;">
              <img src="data:image/png;base64,{b64}"
                   style="width:180px;max-width:100%;height:auto;">
            </div>
            ''', unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="padding:1.5rem 1rem 1rem;border-bottom:1px solid #E2E8F0;margin-bottom:1rem;">
              <div style="font-size:var(--text-lg);font-weight:800;color:{C_RED};letter-spacing:-0.5px;">
                Protección <span style="font-weight:400;color:#0F172A;">Catódica de Colombia</span>
              </div>
              <div style="font-size:var(--text-sm);color:#64748B;letter-spacing:0.08em;margin-top:4px;font-weight:600;">
                CATHODIC PROTECTION DASHBOARD
              </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<hr style="border-color:#E2E8F0;margin:0.5rem 0;">', unsafe_allow_html=True)

        # ── Navegación principal ──
        modo = st.radio(
            "Módulo",
            ["PAP / DCVG", "CIPS"],
            label_visibility="collapsed",
            key="nav_modo"
        )

        st.markdown('<hr style="border-color:#E2E8F0;margin:0.5rem 0;">', unsafe_allow_html=True)

        # ── Controles según modo ──────────────────────────────────────────────
        if modo == "PAP / DCVG":
            st.markdown('<p style="font-size:var(--text-xs);font-weight:600;color:#94A3B8;'
                        'text-transform:uppercase;letter-spacing:0.1em;margin:0.5rem 0 0.3rem;">Archivos</p>',
                        unsafe_allow_html=True)
            uploaded = st.file_uploader("Excel FastField", type=["xlsx"],
                                        accept_multiple_files=True,
                                        label_visibility="collapsed",
                                        key="pap_uploader")

            sp_result = fetch_sharepoint_files()
            sp_files  = sp_result[0] if isinstance(sp_result, tuple) else sp_result
            pdf_urls  = sp_result[1] if isinstance(sp_result, tuple) else {}
            all_files = (uploaded or []) + sp_files

            inspecciones = []
            if all_files:
                vistos = set()
                for f in all_files:
                    if f.name in vistos: continue
                    vistos.add(f.name)
                    try:
                        data = load_excel(f, pdf_urls); data["filename"] = f.name
                        inspecciones.append(data)
                    except Exception as e:
                        st.error(f"{f.name[:22]}…: {e}")

            if inspecciones:
                st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                by_tramo = {}
                for d in inspecciones:
                    by_tramo.setdefault(d["meta"]["tramo"], []).append(d)
                for tramo, items in by_tramo.items():
                    st.markdown(f'<p style="font-size:var(--text-xs);color:#999;text-transform:uppercase;'
                                f'letter-spacing:0.08em;margin:0.7rem 0 0.2rem;">{tramo}</p>',
                                unsafe_allow_html=True)
                    for d in items:
                        dot_color = "#374151" if d["tipo"]=="PAP" else "#6B7280"
                        st.markdown(f"""
                        <div style="display:flex;align-items:flex-start;gap:8px;
                                    padding:6px 4px;border-bottom:1px solid #F1F5F9;">
                          <span style="width:6px;height:6px;border-radius:50%;
                                       background:{dot_color};margin-top:5px;flex-shrink:0;"></span>
                          <div>
                            <div style="font-size:var(--text-sm);font-weight:600;color:#0F172A;">
                              {d['tipo']} — {d['meta']['fecha']}
                            </div>
                            <div style="font-size:var(--text-sm);color:#94A3B8;margin-top:1px;">
                              {d['meta']['inspector']} · {len(d['df'])} pts
                            </div>
                          </div>
                        </div>""", unsafe_allow_html=True)

            return modo, inspecciones, None, None, None, []

        else:  # CIPS
            st.markdown('<hr style="border-color:#E2E8F0;margin:0.8rem 0;">', unsafe_allow_html=True)

            # 1. Lista de archivos (rápido — solo nombres/URLs, sin descargar datos)
            actual_meta, hist_meta, meta_errs = fetch_cips_metadata()
            for err in meta_errs[:2]:
                st.warning(f"⚠ {err}", icon=None)

            all_meta = [(m, "ACTUAL")    for m in actual_meta] + \
                       [(m, "HISTÓRICO") for m in hist_meta]

            # 2. Cargar archivos uno a uno (session_state como caché)
            cache   = _cips_cache()
            pending = [m for m, _ in all_meta if m["name"] not in cache]

            if pending:
                prog = st.progress(0, text=f"Cargando {len(pending)} archivo(s)…")
                done = len(cache)
                total = done + len(pending)
                for m, _cat in all_meta:
                    if m["name"] in cache:
                        continue
                    prog.progress(done / max(total, 1),
                                  text=f"⬇ {m['name'][:32]}…")
                    _, err = _load_one_cips(m)
                    if err:
                        st.warning(f"⚠ {m['name'][:25]}: {err}", icon=None)
                    done += 1
                prog.empty()

            # 3. Construir listas desde caché
            actual_list, historico_list = [], []
            for m, cat in all_meta:
                d = cache.get(m["name"])
                if d:
                    (actual_list if cat == "ACTUAL" else historico_list).append(d)

            # 4. Archivos subidos manualmente
            st.markdown('<p style="font-size:var(--text-xs);font-weight:600;color:#94A3B8;'
                        'text-transform:uppercase;letter-spacing:0.1em;'
                        'margin:0.6rem 0 0.2rem;">Subir archivos</p>',
                        unsafe_allow_html=True)
            uploaded_cips = st.file_uploader("Excel CIPS", type=["xlsx"],
                                              accept_multiple_files=True,
                                              label_visibility="collapsed",
                                              key="cips_uploader")
            nombres_sp = {d["tramo"] for d in actual_list + historico_list}
            for f in (uploaded_cips or []):
                try:
                    f.seek(0)
                    d = load_cips_processed(f, "ACTUAL")
                    if d["tramo"] not in nombres_sp:
                        actual_list.append(d)
                except Exception:
                    pass

            # 5. Lista de archivos en sidebar
            if actual_list or historico_list:
                st.markdown('<hr style="border-color:#E0E0E0;margin:0.5rem 0;">', unsafe_allow_html=True)
                for label, lst, color in [("ACTUALES", actual_list, "#D50032"),
                                           ("HISTÓRICOS", historico_list, "#6B7280")]:
                    if not lst: continue
                    st.markdown(f'<p style="font-size:var(--text-xs);color:{color};font-weight:700;'
                                f'letter-spacing:0.08em;margin:0.5rem 0 0.2rem;">{label}</p>',
                                unsafe_allow_html=True)
                    for d in lst:
                        st.markdown(f"""
                        <div style="display:flex;align-items:flex-start;gap:8px;
                                    padding:5px 4px;border-bottom:1px solid #F1F5F9;">
                          <span style="width:6px;height:6px;border-radius:50%;
                                       background:{color};margin-top:5px;flex-shrink:0;"></span>
                          <div>
                            <div style="font-size:var(--text-sm);font-weight:600;color:#0F172A;">
                              {d['tramo'][:28]}
                            </div>
                            <div style="font-size:var(--text-sm);color:#94A3B8;margin-top:1px;">
                              {d['fecha']} · {len(d['df']):,} pts
                            </div>
                          </div>
                        </div>""", unsafe_allow_html=True)

            if st.button("Refrescar SharePoint", use_container_width=True, key="cips_refresh"):
                fetch_cips_metadata.clear()
                st.session_state.pop("cips_files", None)
                st.rerun()

            return modo, None, None, None, None, (actual_list, historico_list)


# ══════════════════════════════════════════════════════════════════════════════
# Loading animation
# ══════════════════════════════════════════════════════════════════════════════

def inject_loading_animation():
    """Reemplaza el spinner de Streamlit con el logo centrado y animado."""
    logo_b64 = ""
    if os.path.exists("logo-pcc-hd.png"):
        with open("logo-pcc-hd.png", "rb") as _f:
            logo_b64 = base64.b64encode(_f.read()).decode()

    st.markdown(f"""
    <style>
    @keyframes pcc-spin  {{ to {{ transform: rotate(360deg); }} }}
    @keyframes pcc-pulse {{ 0%,100% {{ transform:scale(1);   opacity:1;   }}
                            50%      {{ transform:scale(1.07);opacity:0.88;}} }}
    @keyframes pcc-fade  {{ from {{ opacity:0; }} to {{ opacity:1; }} }}

    /* Ocultar spinner nativo de Streamlit */
    [data-testid="stSpinnerContainer"],
    [data-testid="stSpinner"] {{ visibility: hidden !important; }}

    /* Overlay de carga */
    #pcc-loader {{
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(255,255,255,0.97);
        backdrop-filter: blur(6px);
        z-index: 99999;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        animation: pcc-fade 0.2s ease;
    }}
    #pcc-loader-ring {{
        position: absolute;
        width: 144px;
        height: 144px;
        border-radius: 50%;
        border: 3px solid #F1F5F9;
        border-top-color: #D50032;
        animation: pcc-spin 0.9s linear infinite;
    }}
    #pcc-loader-logo {{
        width: 108px;
        height: 108px;
        object-fit: contain;
        animation: pcc-pulse 1.6s ease-in-out infinite;
    }}
    #pcc-loader-text {{
        margin-top: 1.8rem;
        color: #94A3B8;
        font-size: 0.82rem;
        font-family: sans-serif;
        letter-spacing: 0.12em;
        text-transform: uppercase;
    }}
    </style>

    <div id="pcc-loader">
        <div style="position:relative;width:144px;height:144px;
                    display:flex;align-items:center;justify-content:center;">
            <div id="pcc-loader-ring"></div>
            <img id="pcc-loader-logo"
                 src="data:image/png;base64,{logo_b64}"
                 alt="PCC">
        </div>
        <p id="pcc-loader-text">Cargando&hellip;</p>
    </div>

    <script>
    (function() {{
        var overlay = document.getElementById('pcc-loader');
        if (!overlay) return;

        var hideTimer = null;

        function show() {{
            if (hideTimer) {{ clearTimeout(hideTimer); hideTimer = null; }}
            overlay.style.display = 'flex';
        }}
        function hide() {{
            // Pequeño delay para evitar parpadeos en transiciones rápidas
            hideTimer = setTimeout(function() {{
                overlay.style.display = 'none';
            }}, 150);
        }}

        function check() {{
            // 1. Spinner explícito (st.cache_data show_spinner, st.spinner)
            var hasSpinner = document.querySelector(
                '[data-testid="stSpinnerContainer"], [data-testid="stSpinner"]'
            );
            // 2. Estado "stale": Streamlit pone data-stale="true" en los elementos
            //    mientras está recalculando el script (es el estado del grisado)
            var isStale = document.querySelector('[data-stale="true"]');
            // 3. Indicador de "Running" en la toolbar de Streamlit
            var statusRunning = document.querySelector(
                '[data-testid="stStatusWidget"][aria-label="Running"],' +
                '[data-testid="stToolbarActions"] [aria-label="Stop"]'
            );

            if (hasSpinner || isStale || statusRunning) {{
                show();
            }} else {{
                hide();
            }}
        }}

        // Observar cambios de DOM y de atributos (para capturar data-stale)
        new MutationObserver(check).observe(document.body, {{
            childList: true, subtree: true,
            attributes: true, attributeFilter: ['data-stale', 'aria-label']
        }});

        // Verificar estado inicial
        check();
    }})();
    </script>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    inject_loading_animation()
    result = sidebar()
    modo   = result[0]

    if modo == "PAP / DCVG":
        inspecciones = result[1]

        if not inspecciones:
            st.markdown(f"""
            <div style="margin-top:4rem;text-align:center;padding:4rem;background:white;
                        border-radius:16px;border:1px dashed #CBD5E1;
                        box-shadow:0 4px 6px rgba(0,0,0,0.02);animation:fadeUp 0.5s ease-out forwards;">
              <h2 style="color:#0F172A;margin-bottom:0.8rem;font-weight:700;">Dashboard PAP / DCVG</h2>
              <p style="color:#64748B;font-size:var(--text-md);line-height:1.6;max-width:500px;margin:0 auto;">
                Sube los archivos <b>.xlsx</b> exportados desde FastField usando el panel lateral.<br>
                La app detectará automáticamente si es <b>PAP</b> o <b>DCVG</b> y organizará los datos por tramo.
              </p>
            </div>
            """, unsafe_allow_html=True)
            return

        pap_list  = [d for d in inspecciones if d["tipo"]=="PAP"]
        dcvg_list = [d for d in inspecciones if d["tipo"]=="DCVG"]

        if len(inspecciones) > 1:
            render_resumen(inspecciones)
            st.markdown('<div class="sec-div" style="margin:2rem 0;border-top:2px solid #EBEBEB;"></div>',
                        unsafe_allow_html=True)

        if pap_list:
            sel_pap = pap_list[0]
            if len(pap_list) > 1:
                opts = {f"{d['meta']['tramo']} — {d['meta']['inspector']} ({d['meta']['fecha']})": d
                        for d in pap_list}
                with st.sidebar:
                    st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                    st.markdown('<p style="font-size:var(--text-sm);color:#888;margin-bottom:0.2rem;">SELECCIONAR PAP</p>',
                                unsafe_allow_html=True)
                    sel_pap = opts[st.selectbox("PAP", list(opts.keys()),
                                                 label_visibility="collapsed")]
            render_pap(sel_pap)

        if dcvg_list:
            if pap_list:
                st.markdown('<div style="margin:1.5rem 0;border-top:2px solid #EBEBEB;"></div>',
                            unsafe_allow_html=True)
            sel_dcvg = dcvg_list[0]
            if len(dcvg_list) > 1:
                opts = {f"{d['meta']['tramo']} — {d['meta']['inspector']} ({d['meta']['fecha']})": d
                        for d in dcvg_list}
                with st.sidebar:
                    st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                    st.markdown('<p style="font-size:var(--text-sm);color:#888;margin-bottom:0.2rem;">SELECCIONAR DCVG</p>',
                                unsafe_allow_html=True)
                    sel_dcvg = opts[st.selectbox("DCVG", list(opts.keys()),
                                                  label_visibility="collapsed")]
            render_dcvg(sel_dcvg)

    else:  # CIPS
        actual_list, historico_list = result[5]
        todos = actual_list + historico_list

        if not todos:
            st.markdown("""
            <div style="margin-top:4rem;text-align:center;padding:4rem;background:white;
                        border-radius:16px;border:1px dashed #CBD5E1;animation:fadeUp 0.5s ease-out forwards;">
              <h2 style="color:#0F172A;margin-bottom:0.8rem;font-weight:700;">Dashboard CIPS</h2>
              <p style="color:#64748B;font-size:var(--text-md);line-height:1.6;max-width:520px;margin:0 auto;">
                Los archivos se sincronizan automáticamente desde SharePoint.<br>
                También puedes subir archivos <b>.xlsx</b> manualmente desde el panel lateral.
              </p>
            </div>""", unsafe_allow_html=True)
            return

        # Vista general comparativa (siempre visible si hay datos)
        render_cips_comparativo(actual_list, historico_list)

        # Detalle de inspección individual (selector)
        if todos:
            divider()
            st.markdown('<p style="font-size:var(--text-base);font-weight:600;color:#475569;margin-bottom:0.4rem;">VER DETALLE DE INSPECCIÓN</p>',
                        unsafe_allow_html=True)
            opts = {f"{'🔴' if d['categoria']=='ACTUAL' else '⬜'} {d['tramo']} — {d['fecha']}": d
                    for d in todos}
            sel_key = st.selectbox("Inspección", list(opts.keys()),
                                   label_visibility="collapsed", key="cips_detail_sel")
            render_cips_dashboard(opts[sel_key])


if __name__ == "__main__":
    main()
