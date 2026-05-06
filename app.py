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

st.set_page_config(
    page_title="PCC – Inspecciones",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Tokens de color ────────────────────────────────────────────────────────────
C_RED   = "#D50032"
C_PROT  = "#1565C0"
C_SOBRE = "#0D47A1"
C_SIN   = "#C62828"
C_DCVG  = "#1B5E20"
C_CIPS  = "#6A1B9A"
C_GRAY  = "#F3F3F3"
C_LINE  = "#E0E0E0"

ESTADO_COLORS = {
    "Protegido":      C_PROT,
    "Sobreprotegido": C_SOBRE,
    "Sin protección": C_SIN,
    "Sin medición":   "#BDBDBD",
}

# ── CSS unificado ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
  .stApp { background: #F8FAFC; }
  .block-container { padding: 2rem 2.5rem 1rem 2.5rem !important; max-width: 1400px; }

  /* ── Keyframes ─────────────────────────────────────────────────────────── */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(20px); }
    to   { opacity: 1; transform: translateY(0); }
  }
  @keyframes fadeIn {
    from { opacity: 0; }
    to   { opacity: 1; }
  }
  @keyframes slideInLeft {
    from { opacity: 0; transform: translateX(-24px); }
    to   { opacity: 1; transform: translateX(0); }
  }
  @keyframes slideInRight {
    from { opacity: 0; transform: translateX(24px); }
    to   { opacity: 1; transform: translateX(0); }
  }
  @keyframes scaleIn {
    from { opacity: 0; transform: scale(0.88); }
    to   { opacity: 1; transform: scale(1); }
  }
  @keyframes pulseGlow {
    0%, 100% { box-shadow: 0 4px 6px -1px rgba(0,0,0,0.04); }
    50%       { box-shadow: 0 8px 30px -4px rgba(213,0,50,0.18); }
  }
  @keyframes gradientShift {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
  }
  @keyframes dotPulse {
    0%, 100% { transform: scale(1); opacity: 1; }
    50%       { transform: scale(1.35); opacity: 0.75; }
  }
  @keyframes shimmer {
    0%   { background-position: -400px 0; }
    100% { background-position: 400px 0; }
  }
  @keyframes borderDraw {
    from { opacity: 0; transform: scaleX(0); transform-origin: left; }
    to   { opacity: 1; transform: scaleX(1); }
  }
  @keyframes countPop {
    0%   { transform: scale(0.6); opacity: 0; }
    70%  { transform: scale(1.08); }
    100% { transform: scale(1);   opacity: 1; }
  }
  @keyframes footerSlide {
    from { opacity: 0; transform: translateY(30px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  /* ── Sidebar ───────────────────────────────────────────────────────────── */
  [data-testid="stSidebar"] > div:first-child {
    background: #FFFFFF;
    border-right: 1px solid #E2E8F0;
    box-shadow: 2px 0 14px rgba(0,0,0,0.04);
    animation: slideInLeft 0.45s cubic-bezier(.22,.68,0,1.2) forwards;
  }
  [data-testid="stSidebar"] * { color: #1E293B !important; }
  [data-testid="stSidebar"] hr { border-color: #E2E8F0; }
  [data-testid="stSidebar"] [data-testid="stRadio"] label {
    font-size: 0.88rem !important; font-weight: 600 !important;
    transition: color 0.2s ease !important;
  }

  /* ── Títulos de sección ────────────────────────────────────────────────── */
  .pbi-title {
    font-size: 1rem; font-weight: 600; color: #0F172A;
    margin: 1.2rem 0 0.8rem 0; letter-spacing: -0.01em;
    animation: fadeIn 0.5s ease-out forwards;
    position: relative; padding-left: 10px;
  }
  .pbi-title::before {
    content: '';
    position: absolute; left: 0; top: 50%; transform: translateY(-50%);
    width: 3px; height: 70%; border-radius: 4px;
    background: linear-gradient(180deg, #D50032, #ff6b6b);
    animation: borderDraw 0.4s ease-out 0.2s both;
  }

  /* ── Header principal ──────────────────────────────────────────────────── */
  .main-header {
    background: linear-gradient(135deg, #ffffff 0%, #fafcff 100%);
    padding: 1.2rem 1.8rem;
    border-radius: 14px;
    border: 1px solid #E2E8F0;
    border-left: 4px solid #D50032;
    box-shadow: 0 4px 16px -4px rgba(0,0,0,0.06);
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 1.5rem;
    animation: fadeUp 0.5s cubic-bezier(.22,.68,0,1.2) forwards;
    transition: box-shadow 0.3s ease, transform 0.3s ease;
  }
  .main-header:hover {
    box-shadow: 0 8px 28px -6px rgba(213,0,50,0.12);
    transform: translateY(-1px);
  }
  .main-header-title {
    font-size: 1.3rem; font-weight: 700; color: #8B0000;
    margin: 0; display: flex; align-items: center; gap: 0.6rem;
  }
  .main-header-meta { font-size: 0.9rem; color: #64748B; font-weight: 500; }

  /* ── Cards / Stats ─────────────────────────────────────────────────────── */
  .stat-container {
    background: white; border: 1px solid #E2E8F0; border-radius: 12px;
    padding: 1.2rem; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    opacity: 0;
    animation: scaleIn 0.45s cubic-bezier(.22,.68,0,1.2) forwards;
    transition: transform 0.25s ease, box-shadow 0.25s ease;
    cursor: default;
  }
  .stat-container:hover {
    transform: translateY(-4px) scale(1.01);
    box-shadow: 0 12px 32px -8px rgba(0,0,0,0.12);
  }
  .stat-container:nth-child(1) { animation-delay: 0.05s; }
  .stat-container:nth-child(2) { animation-delay: 0.12s; }
  .stat-container:nth-child(3) { animation-delay: 0.19s; }
  .stat-container:nth-child(4) { animation-delay: 0.26s; }
  .stat-container:nth-child(5) { animation-delay: 0.33s; }

  .stat-label {
    font-size: 0.75rem; text-transform: uppercase; font-weight: 600;
    color: #64748B; letter-spacing: 0.05em; margin-bottom: 6px;
  }
  .stat-val {
    font-size: 1.6rem; font-weight: 700; color: #0F172A; margin-bottom: 0;
    animation: countPop 0.5s cubic-bezier(.22,.68,0,1.2) 0.3s both;
  }

  /* ── Estado list + dots ────────────────────────────────────────────────── */
  .estado-item {
    display: flex; align-items: center; gap: 8px;
    font-size: 0.85rem; margin: 6px 0; padding: 4px 0;
    animation: slideInLeft 0.35s ease-out forwards;
    transition: background 0.2s ease; border-radius: 6px; padding: 4px 6px;
  }
  .estado-item:hover { background: #F8FAFC; }
  .dot {
    width: 10px; height: 10px; border-radius: 50%; display: inline-block;
    animation: dotPulse 2.5s ease-in-out infinite;
  }

  /* ── Footer ────────────────────────────────────────────────────────────── */
  .pcc-footer {
    background: linear-gradient(135deg, #D50032 0%, #8B0000 50%, #A00025 100%);
    background-size: 200% 200%;
    color: white; padding: 1.4rem 2rem; margin-top: 3rem; border-radius: 14px;
    display: flex; align-items: center; gap: 1rem;
    box-shadow: 0 12px 32px -6px rgba(213,0,50,0.35);
    animation: footerSlide 0.6s cubic-bezier(.22,.68,0,1.2) 0.2s both,
               gradientShift 6s ease infinite;
    transition: box-shadow 0.3s ease;
  }
  .pcc-footer:hover {
    box-shadow: 0 16px 40px -6px rgba(213,0,50,0.45);
  }
  .pcc-footer-logo {
    font-size: 1.6rem; font-weight: 800; letter-spacing: -1px;
    animation: slideInLeft 0.5s cubic-bezier(.22,.68,0,1.2) 0.4s both;
  }
  .pcc-footer-text { font-size: 0.95rem; font-weight: 500; opacity: 0.9; }

  /* ── Separadores ───────────────────────────────────────────────────────── */
  .sec-div {
    border: none; height: 1px;
    background: linear-gradient(90deg, transparent, #E2E8F0 20%, #E2E8F0 80%, transparent);
    margin: 1.5rem 0 1rem 0;
    animation: borderDraw 0.5s ease-out forwards;
  }

  /* ── Bloques CIPS ──────────────────────────────────────────────────────── */
  .bloque {
    background: white; border-radius: 14px; padding: 1.5rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06); margin-bottom: 1rem;
    animation: fadeUp 0.45s cubic-bezier(.22,.68,0,1.2) forwards;
    transition: box-shadow 0.25s ease, transform 0.25s ease;
  }
  .bloque:hover {
    box-shadow: 0 8px 28px rgba(0,0,0,0.09);
    transform: translateY(-2px);
  }
  .bloque-titulo {
    font-weight: 700; font-size: 0.85rem; color: #8B0000;
    text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 1rem;
    padding-bottom: 0.5rem; border-bottom: 2px solid #F0F2F6;
  }

  /* ── Métricas CIPS ─────────────────────────────────────────────────────── */
  [data-testid="stMetric"] {
    background: white; border-radius: 14px;
    padding: 1.2rem 1.4rem !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    animation: scaleIn 0.45s cubic-bezier(.22,.68,0,1.2) forwards;
    transition: transform 0.25s ease, box-shadow 0.25s ease;
  }
  [data-testid="stMetric"]:hover {
    transform: translateY(-3px) scale(1.01);
    box-shadow: 0 10px 28px -6px rgba(0,0,0,0.12);
  }
  [data-testid="stMetricLabel"] {
    font-size: 0.75rem !important; color: #888 !important;
    font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.5px;
  }
  [data-testid="stMetricValue"] {
    font-size: 2.2rem !important; font-weight: 800 !important;
    animation: countPop 0.55s cubic-bezier(.22,.68,0,1.2) 0.15s both;
  }

  /* ── Botones ───────────────────────────────────────────────────────────── */
  .stButton > button {
    background: linear-gradient(135deg, #b8233e, #d42848) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(123,30,58,0.3) !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease !important;
  }
  .stButton > button:hover {
    transform: translateY(-2px) scale(1.02) !important;
    box-shadow: 0 8px 22px rgba(123,30,58,0.4) !important;
  }
  .stButton > button:active {
    transform: translateY(0) scale(0.98) !important;
  }
  [data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #1A7A4A, #22A06B) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease !important;
  }
  [data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) scale(1.02) !important;
    box-shadow: 0 8px 22px rgba(26,122,74,0.35) !important;
  }

  /* ── DataFrames ────────────────────────────────────────────────────────── */
  [data-testid="stDataFrame"] {
    animation: fadeUp 0.55s cubic-bezier(.22,.68,0,1.2) 0.1s both;
    border-radius: 10px; border: 1px solid #E2E8F0; overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    transition: box-shadow 0.25s ease;
  }
  [data-testid="stDataFrame"]:hover {
    box-shadow: 0 6px 20px rgba(0,0,0,0.08);
  }

  /* ── Plotly charts ─────────────────────────────────────────────────────── */
  .js-plotly-plot {
    animation: scaleIn 0.5s cubic-bezier(.22,.68,0,1.2) 0.1s both;
    border-radius: 10px;
  }

  /* ── Expanders ─────────────────────────────────────────────────────────── */
  [data-testid="stExpander"] {
    border-radius: 10px !important; border: 1px solid #E2E8F0 !important;
    overflow: hidden;
    transition: box-shadow 0.25s ease;
  }
  [data-testid="stExpander"]:hover {
    box-shadow: 0 4px 16px rgba(0,0,0,0.06);
  }

  /* ── Shimmer para SP sync badges ───────────────────────────────────────── */
  .sp-badge-shimmer {
    background: linear-gradient(90deg, #F0FFF4 25%, #dcffe8 50%, #F0FFF4 75%);
    background-size: 400px 100%;
    animation: shimmer 1.8s infinite;
  }

  #MainMenu, footer { visibility: hidden; }
  [data-testid="stToolbar"] { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Barra de progreso + animaciones JS globales ────────────────────────────────
components.html("""
<style>
  /* Barra de carga superior */
  #nprogress-bar {
    position: fixed; top: 0; left: 0; z-index: 99999;
    height: 3px; width: 0%;
    background: linear-gradient(90deg, #D50032, #ff6b6b, #D50032);
    background-size: 200% 100%;
    animation: barLoad 1.2s ease-out forwards, gradMove 1s linear infinite;
    border-radius: 0 3px 3px 0;
    box-shadow: 0 0 10px rgba(213,0,50,0.6), 0 0 5px rgba(213,0,50,0.4);
  }
  @keyframes barLoad {
    0%   { width: 0%; opacity: 1; }
    70%  { width: 85%; }
    100% { width: 100%; opacity: 0; }
  }
  @keyframes gradMove {
    0%   { background-position: 0% 50%; }
    100% { background-position: 200% 50%; }
  }

  /* Reveal on scroll */
  .reveal {
    opacity: 0;
    transform: translateY(28px);
    transition: opacity 0.6s cubic-bezier(.22,.68,0,1.2),
                transform 0.6s cubic-bezier(.22,.68,0,1.2);
  }
  .reveal.visible {
    opacity: 1;
    transform: translateY(0);
  }
</style>

<div id="nprogress-bar"></div>

<script>
(function() {
  // Barra de progreso al cargar
  var bar = document.getElementById('nprogress-bar');
  if (bar) {
    setTimeout(function() { bar.style.opacity = '0'; }, 1300);
  }

  // Scroll-reveal: aplica a elementos dentro del frame principal
  function applyReveal() {
    try {
      var doc = window.parent.document;

      // Elementos a animar
      var selectors = [
        '[data-testid="stDataFrame"]',
        '.js-plotly-plot',
        '[data-testid="stMetric"]',
        '[data-testid="stExpander"]',
        '[data-testid="stVerticalBlock"] > div > div',
      ];

      selectors.forEach(function(sel) {
        doc.querySelectorAll(sel).forEach(function(el) {
          if (!el.classList.contains('reveal')) {
            el.classList.add('reveal');
          }
        });
      });

      var observer = new IntersectionObserver(function(entries) {
        entries.forEach(function(e) {
          if (e.isIntersecting) {
            e.target.classList.add('visible');
            observer.unobserve(e.target);
          }
        });
      }, { threshold: 0.08 });

      doc.querySelectorAll('.reveal').forEach(function(el) {
        observer.observe(el);
      });

    } catch(e) {}
  }

  // Ejecutar en carga y en cada mutación del DOM
  applyReveal();
  setTimeout(applyReveal, 600);
  setTimeout(applyReveal, 1500);

  try {
    var mo = new MutationObserver(function() { applyReveal(); });
    mo.observe(window.parent.document.body, { childList: true, subtree: true });
  } catch(e) {}
})();
</script>
""", height=0)


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
        <div style="background:white;border:1px solid #E2E8F0;border-radius:12px;
                    padding:1.2rem;box-shadow:0 2px 8px rgba(0,0,0,0.04);
                    border-left:4px solid {color};
                    animation:scaleIn 0.45s cubic-bezier(.22,.68,0,1.2) {i*0.07:.2f}s both;
                    transition:transform 0.25s ease,box-shadow 0.25s ease;cursor:default;"
             onmouseover="this.style.transform='translateY(-4px) scale(1.01)';this.style.boxShadow='0 12px 32px -8px rgba(0,0,0,0.12)'"
             onmouseout="this.style.transform='';this.style.boxShadow='0 2px 8px rgba(0,0,0,0.04)'">
          <p style="font-size:0.72rem;text-transform:uppercase;font-weight:700;
                    color:#64748B;letter-spacing:0.06em;margin:0 0 6px 0;">{label}</p>
          <p style="font-size:1.8rem;font-weight:800;color:{color};margin:0;
                    animation:countPop 0.55s cubic-bezier(.22,.68,0,1.2) {i*0.07+0.2:.2f}s both;">
            {display}
          </p>
        </div>
        """

    full_html = f"""
    <style>
      @keyframes scaleIn {{
        from {{ opacity:0; transform:scale(0.88); }}
        to   {{ opacity:1; transform:scale(1); }}
      }}
      @keyframes countPop {{
        0%  {{ transform:scale(0.6); opacity:0; }}
        70% {{ transform:scale(1.08); }}
        100%{{ transform:scale(1); opacity:1; }}
      }}
    </style>
    <div style="display:grid;grid-template-columns:repeat({len(items)},1fr);gap:1rem;margin-bottom:1rem;">
      {cards_html}
    </div>
    <script>
    function animateCounter(id, target, duration, delay) {{
      setTimeout(function() {{
        var el = document.getElementById(id);
        if (!el) return;
        var start = 0, step = target / (duration / 16);
        var timer = setInterval(function() {{
          start += step;
          if (start >= target) {{ el.textContent = target.toLocaleString(); clearInterval(timer); }}
          else {{ el.textContent = Math.floor(start).toLocaleString(); }}
        }}, 16);
      }}, delay);
    }}
    {js_counters}
    </script>
    """
    components.html(full_html, height=110)

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


def render_cips_comparativo(actual_list, historico_list):
    todos = actual_list + historico_list
    n_act = sum(len(d["df"]) for d in actual_list)
    n_his = sum(len(d["df"]) for d in historico_list)

    # ── Header ─────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:white;border:1px solid #E2E8F0;border-radius:12px;
                padding:1.2rem 1.8rem;margin-bottom:1rem;
                box-shadow:0 4px 16px -4px rgba(0,0,0,0.06);
                display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.8rem;">
      <div>
        <div style="font-size:0.7rem;color:#D50032;font-weight:700;
                    text-transform:uppercase;letter-spacing:0.12em;margin-bottom:4px;">
          Inspecciones CIPS — Vista Comparativa
        </div>
        <div style="font-size:1.2rem;font-weight:800;color:#0F172A;">
          {len(todos)} tramo{'s' if len(todos)!=1 else ''} analizados
        </div>
      </div>
      <div style="display:flex;gap:1rem;text-align:center;flex-wrap:wrap;">
        <div style="background:#FFF5F6;border:1px solid #FECDD3;border-radius:10px;padding:0.6rem 1rem;">
          <div style="font-size:1.3rem;font-weight:800;color:#D50032;">{len(actual_list)}</div>
          <div style="font-size:0.68rem;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;">Actuales</div>
          <div style="font-size:0.72rem;color:#D50032;font-weight:600;">{n_act:,} pts</div>
        </div>
        <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:0.6rem 1rem;">
          <div style="font-size:1.3rem;font-weight:800;color:#6B7280;">{len(historico_list)}</div>
          <div style="font-size:0.68rem;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;">Históricos</div>
          <div style="font-size:0.72rem;color:#6B7280;font-weight:600;">{n_his:,} pts</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Análisis de criticidad ─────────────────────────────────────────────────
    stats = _criticidad_stats(todos)
    if stats:
        n_total_pts = sum(r["total"] for r in stats)
        pct_critico_global = sum(r["n_desp"] + r["n_sobre"] for r in stats) / n_total_pts * 100 if n_total_pts else 0
        tramo_mas_critico  = stats[0]["tramo"] if stats else "—"
        score_max          = stats[0]["score"] if stats else 0

        # KPIs globales
        kpi_css = ("background:white;border:1px solid #E2E8F0;border-radius:10px;"
                   "padding:0.9rem 1.1rem;box-shadow:0 2px 8px -2px rgba(0,0,0,0.05);")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div style="{kpi_css}">'
                        f'<div style="font-size:0.65rem;color:#94A3B8;text-transform:uppercase;letter-spacing:0.1em;">Tramos</div>'
                        f'<div style="font-size:1.6rem;font-weight:800;color:#0F172A;">{len(stats)}</div></div>',
                        unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div style="{kpi_css}">'
                        f'<div style="font-size:0.65rem;color:#94A3B8;text-transform:uppercase;letter-spacing:0.1em;">Puntos totales</div>'
                        f'<div style="font-size:1.6rem;font-weight:800;color:#0F172A;">{n_total_pts:,}</div></div>',
                        unsafe_allow_html=True)
        with c3:
            color_kpi = "#D50032" if pct_critico_global > 20 else "#374151"
            st.markdown(f'<div style="{kpi_css}">'
                        f'<div style="font-size:0.65rem;color:#94A3B8;text-transform:uppercase;letter-spacing:0.1em;">Fuera de rango global</div>'
                        f'<div style="font-size:1.6rem;font-weight:800;color:{color_kpi};">{pct_critico_global:.1f}%</div></div>',
                        unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div style="{kpi_css};border-left:4px solid #D50032;">'
                        f'<div style="font-size:0.65rem;color:#94A3B8;text-transform:uppercase;letter-spacing:0.1em;">Tramo más crítico</div>'
                        f'<div style="font-size:0.95rem;font-weight:800;color:#D50032;line-height:1.2;">'
                        f'{tramo_mas_critico[:30]}</div>'
                        f'<div style="font-size:0.72rem;color:#6B7280;">score {score_max:.0f}</div></div>',
                        unsafe_allow_html=True)

        st.markdown("<div style='margin:0.6rem 0;'></div>", unsafe_allow_html=True)

        # Gráfica horizontal de criticidad por tramo
        divider()
        pbi_title("Ranking de criticidad por tramo")

        bar_rows = []
        for r in stats:
            bar_rows.append({"Tramo": r["tramo"], "Estado": "PROTEGIDO",      "Pct": r["pct_prot"]})
            bar_rows.append({"Tramo": r["tramo"], "Estado": "DESPROTEGIDO",   "Pct": r["pct_desp"]})
            bar_rows.append({"Tramo": r["tramo"], "Estado": "SOBREPROTEGIDO", "Pct": r["pct_sobre"]})
        bar_df = pd.DataFrame(bar_rows)
        # Orden: más crítico arriba
        order = [r["tramo"] for r in reversed(stats)]

        fig_bar = px.bar(
            bar_df, x="Pct", y="Tramo", color="Estado", orientation="h",
            color_discrete_map={"PROTEGIDO": "#374151", "DESPROTEGIDO": "#D50032", "SOBREPROTEGIDO": "#7F1D1D"},
            category_orders={"Tramo": order},
            barmode="stack", height=max(260, len(stats) * 52),
            custom_data=["Estado"],
        )
        fig_bar.update_traces(
            hovertemplate="%{customdata[0]}: %{x:.1f}%<extra></extra>"
        )
        fig_bar.update_layout(
            **CHART,
            height=max(260, len(stats) * 52),
            xaxis=dict(title="% puntos", ticksuffix="%", showgrid=True,
                       gridcolor="#F1F5F9", range=[0, 100]),
            yaxis=dict(title="", showgrid=False),
            legend=dict(orientation="h", y=-0.18, font_size=11),
            bargap=0.25,
        )
        # Anotación con score en cada barra
        for r in stats:
            fig_bar.add_annotation(
                x=101, y=r["tramo"],
                text=f'<b style="color:#D50032">{r["score"]:.0f}</b>',
                showarrow=False, xanchor="left", font=dict(size=10, color="#D50032"),
                xref="x", yref="y",
            )
        st.plotly_chart(fig_bar, use_container_width=True)

        # Leyenda del score
        st.markdown(
            '<p style="font-size:0.72rem;color:#94A3B8;margin:-0.5rem 0 0.5rem;">'
            'Score de criticidad = % desprotegido × 2 + % sobreprotegido · '
            '<span style="color:#D50032;font-weight:600;">≥50 crítico</span> · '
            '<span style="color:#374151;font-weight:600;">&lt;20 bajo riesgo</span></p>',
            unsafe_allow_html=True,
        )

    # ── Mapa coloreado por Estado_CP ────────────────────────────────────────────
    frames = []
    for d in todos:
        lat = "Lat_corr" if "Lat_corr" in d["df"].columns else None
        lon = "Long_corr" if "Long_corr" in d["df"].columns else None
        if not lat: continue
        cols_need = [lat, lon]
        if "Estado_CP" in d["df"].columns: cols_need.append("Estado_CP")
        sub = d["df"].dropna(subset=[lat, lon]).copy()
        if sub.empty: continue
        if "Estado_CP" not in sub.columns:
            sub["Estado_CP"] = "—"
        sub["_tramo"]    = d["tramo"]
        sub["_categoria"]= d["categoria"]
        # Submuestreo para rendimiento
        if len(sub) > 3000:
            sub = sub.iloc[::max(1, len(sub)//3000)]
        frames.append(sub[[lat, lon, "Estado_CP", "_tramo", "_categoria"]])

    if frames:
        divider()
        pbi_title("Distribución geográfica — estado de protección")
        all_pts = pd.concat(frames, ignore_index=True)
        COLOR_MAP = {
            "PROTEGIDO":      "#6B7280",
            "DESPROTEGIDO":   "#D50032",
            "SOBREPROTEGIDO": "#7F1D1D",
            "—":              "#CBD5E1",
        }
        fig_map = px.scatter_mapbox(
            all_pts, lat="Lat_corr", lon="Long_corr",
            color="Estado_CP",
            color_discrete_map=COLOR_MAP,
            hover_data={"_tramo": True, "_categoria": True,
                        "Lat_corr": False, "Long_corr": False},
            zoom=7, height=500, mapbox_style="open-street-map",
            category_orders={"Estado_CP": ["DESPROTEGIDO","SOBREPROTEGIDO","PROTEGIDO","—"]},
        )
        fig_map.update_traces(marker=dict(size=5, opacity=0.9))
        fig_map.update_layout(
            margin=dict(t=0, b=0, l=0, r=0),
            legend=dict(x=0.01, y=0.99, bgcolor="rgba(255,255,255,0.92)",
                        borderwidth=1, font_size=11),
        )
        st.plotly_chart(fig_map, use_container_width=True)

    # ── Gráfica Off mV por PK ──────────────────────────────────────────────────
    pk_key = lambda d: next((c for c in ["PK_geom_m","PK_real_m"] if c in d["df"].columns), None)
    tiene_grafica = any(pk_key(d) and "Off_mV_limpio" in d["df"].columns for d in todos)
    if tiene_grafica:
        divider()
        pbi_title("Perfil Off mV por PK — todos los tramos")
        fig = go.Figure()

        COLORES_H = ["#9CA3AF","#6B7280","#4B5563","#374151","#D1D5DB"]
        COLORES_A = ["#D50032","#991B1B","#EF4444"]
        for i, d in enumerate(historico_list):
            pk = pk_key(d)
            if not pk or "Off_mV_limpio" not in d["df"].columns: continue
            sub = d["df"].dropna(subset=[pk]).sort_values(pk)
            if len(sub) > 2000: sub = sub.iloc[::max(1, len(sub)//2000)]
            fig.add_trace(go.Scatter(
                x=sub[pk], y=sub["Off_mV_limpio"],
                mode="lines", name=f"[H] {d['tramo'][:22]}",
                line=dict(color=COLORES_H[i % len(COLORES_H)], width=1.2, dash="dot"),
                opacity=0.75))

        for i, d in enumerate(actual_list):
            pk = pk_key(d)
            if not pk or "Off_mV_limpio" not in d["df"].columns: continue
            sub = d["df"].dropna(subset=[pk]).sort_values(pk)
            if len(sub) > 2000: sub = sub.iloc[::max(1, len(sub)//2000)]
            fig.add_trace(go.Scatter(
                x=sub[pk], y=sub["Off_mV_limpio"],
                mode="lines", name=f"[A] {d['tramo'][:22]}",
                line=dict(color=COLORES_A[i % len(COLORES_A)], width=2.0)))

        fig.add_hline(y=-850,  line=dict(color="#6B7280", dash="dash", width=1),
                      annotation_text="-850 mV", annotation_position="top right",
                      annotation_font=dict(size=9, color="#6B7280"))
        fig.add_hline(y=-1200, line=dict(color="#D50032", dash="dash", width=1),
                      annotation_text="-1.200 mV", annotation_position="bottom right",
                      annotation_font=dict(size=9, color="#D50032"))
        fig.add_hrect(y0=-1200, y1=-850, fillcolor="rgba(55,65,81,0.04)", line_width=0)
        fig.update_layout(
            **CHART, height=360,
            xaxis_title=dict(text="PK (m)", font=dict(size=11, color="#64748B")),
            yaxis_title=dict(text="Off mV",  font=dict(size=11, color="#64748B")),
            legend=dict(orientation="h", y=-0.28, font_size=10),
            hovermode="x unified",
        )
        fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
        fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
        st.plotly_chart(fig, use_container_width=True)

    # ── Tabla detallada ─────────────────────────────────────────────────────────
    if stats:
        divider()
        pbi_title("Detalle por tramo")
        tbl_rows = []
        for r in stats:
            nivel = ("🔴 CRÍTICO" if r["score"] >= 50
                     else "🟡 MODERADO" if r["score"] >= 20
                     else "🟢 BAJO")
            tbl_rows.append({
                "Tramo":         r["tramo"],
                "Cat.":          r["categoria"],
                "Fecha":         r["fecha"],
                "Pts":           r["total"],
                "% Prot.":       f"{r['pct_prot']:.1f}%",
                "% Desprot.":    f"{r['pct_desp']:.1f}%",
                "% Sobreprot.":  f"{r['pct_sobre']:.1f}%",
                "Score":         r["score"],
                "Nivel":         nivel,
            })
        tbl = pd.DataFrame(tbl_rows)
        def _style_score(val):
            if isinstance(val, float):
                if val >= 50: return "color:#D50032;font-weight:700;"
                if val >= 20: return "color:#B45309;font-weight:600;"
            return "color:#374151;"
        try:
            st.dataframe(
                tbl.style.map(_style_score, subset=["Score"]),
                use_container_width=True, hide_index=True,
                height=min(60 + len(tbl_rows)*38, 400),
            )
        except Exception:
            st.dataframe(tbl, use_container_width=True, hide_index=True)


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
        <div style="font-size:0.7rem;color:#D50032;font-weight:700;
                    text-transform:uppercase;letter-spacing:0.12em;margin-bottom:4px;">
          Inspección CIPS
        </div>
        <div style="font-size:1.3rem;font-weight:800;color:#0F172A;letter-spacing:-0.02em;">
          {tramo}
        </div>
        <div style="font-size:0.85rem;color:#64748B;margin-top:4px;">
          {fecha} &nbsp;·&nbsp; {total:,} puntos medidos
        </div>
      </div>
      <div style="display:flex;gap:1.2rem;text-align:center;">
        <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:0.65rem 1.1rem;">
          <div style="font-size:1.4rem;font-weight:800;color:#374151;">{pct_p}</div>
          <div style="font-size:0.68rem;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-top:2px;">Protegido</div>
        </div>
        <div style="background:#FFF5F6;border:1px solid #FECDD3;border-radius:10px;padding:0.65rem 1.1rem;">
          <div style="font-size:1.4rem;font-weight:800;color:#D50032;">{pct_d}</div>
          <div style="font-size:0.68rem;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-top:2px;">Desprotegido</div>
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
          <p style="font-size:0.72rem;text-transform:uppercase;font-weight:700;
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
          <p style="font-size:0.72rem;text-transform:uppercase;font-weight:700;
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
          <span style="font-size:0.75rem;color:#64748B;font-weight:600;
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
                <div style="font-size:0.9rem;font-weight:500;">Sin coordenadas GPS</div>
                <div style="font-size:0.8rem;margin-top:4px;">El archivo no tiene columnas Latitude/Longitude</div>
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


def _fetch_cips_folder(cfg, folder, categoria):
    """Descarga xlsx de una carpeta SharePoint, parsea y retorna lista de dicts cargados."""
    out = []
    errors = []
    try:
        app_obj = msal.ConfidentialClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
            client_credential=cfg["client_secret"],
        )
        token_resp = app_obj.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        token = token_resp.get("access_token")
        if not token:
            errors.append(f"No token: {token_resp.get('error_description','')}")
            return out, errors
        headers  = {"Authorization": f"Bearer {token}"}
        hostname = f"{cfg['tenant_name']}.sharepoint.com"
        site_path = cfg["site_url"].replace(f"https://{hostname}", "")
        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}",
            headers=headers
        ).json()
        site_id = site_resp.get("id")
        if not site_id:
            errors.append(f"Site no encontrado: {site_resp.get('error',{}).get('message','')}")
            return out, errors
        items_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder}:/children",
            headers=headers
        ).json()
        items = items_resp.get("value", [])
        if not items and "error" in items_resp:
            errors.append(f"Carpeta '{folder}': {items_resp['error'].get('message','')}")
        for item in items:
            name  = item.get("name", "")
            d_url = item.get("@microsoft.graph.downloadUrl")
            if not name.endswith(".xlsx") or name.startswith("~") or not d_url:
                continue
            r = requests.get(d_url)
            if not r.ok:
                errors.append(f"Descarga fallida {name}: HTTP {r.status_code}")
                continue
            # Reparar XML inválido antes de parsear
            f = _repair_xlsx(r.content)
            f.name = name
            try:
                f.seek(0)
                d = load_cips_processed(f, categoria)
                out.append(d)
            except Exception as e:
                errors.append(f"Error leyendo {name}: {e}")
    except Exception as e:
        errors.append(f"Error general: {e}")
    return out, errors


@st.cache_data(ttl=600, show_spinner="Sincronizando CIPS desde SharePoint...")
def fetch_cips_results():
    """Carga CIPS ACTUAL e HISTÓRICOS desde SharePoint. Retorna (actual_list, historico_list, errores)."""
    if "sharepoint" not in st.secrets:
        return [], [], ["No hay configuración de SharePoint"]
    cfg = st.secrets["sharepoint"]
    actual_folder     = cfg.get("cips_actual_folder",     "Inspecciones Ocensa/CIPS ACTUAL")
    historicos_folder = cfg.get("cips_historicos_folder", "Inspecciones Ocensa/CIPS HISTORICOS")
    actual,    errs_a = _fetch_cips_folder(cfg, actual_folder,     "ACTUAL")
    historicos, errs_h = _fetch_cips_folder(cfg, historicos_folder, "HISTÓRICO")
    return actual, historicos, errs_a + errs_h


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
              <div style="font-size:1.4rem;font-weight:800;color:{C_RED};letter-spacing:-0.5px;">
                Protección <span style="font-weight:400;color:#0F172A;">Catódica de Colombia</span>
              </div>
              <div style="font-size:0.75rem;color:#64748B;letter-spacing:0.08em;margin-top:4px;font-weight:600;">
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
            st.markdown('<p style="font-size:0.8rem;font-weight:600;color:#475569;margin:0.5rem 0;">CARGAR ARCHIVOS</p>',
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
                    st.markdown(f'<p style="font-size:0.67rem;color:#999;text-transform:uppercase;'
                                f'letter-spacing:0.08em;margin:0.7rem 0 0.2rem;">{tramo}</p>',
                                unsafe_allow_html=True)
                    for d in items:
                        color = C_PROT if d["tipo"]=="PAP" else C_DCVG
                        st.markdown(f"""
                        <div style="background:white;border:1px solid #E2E8F0;border-radius:6px;
                                    padding:0.6rem 0.8rem;margin:4px 0;border-left:4px solid {color};
                                    box-shadow:0 1px 2px rgba(0,0,0,0.02);">
                          <div style="font-size:0.85rem;font-weight:600;color:#0F172A;">
                            {d['tipo']} <span style="font-weight:400;color:#64748B;font-size:0.8rem;">— {d['meta']['fecha']}</span>
                          </div>
                          <div style="font-size:0.75rem;color:#64748B;margin-top:4px;">
                            {d['meta']['inspector']} • {len(d['df'])} pts
                          </div>
                        </div>""", unsafe_allow_html=True)

            return modo, inspecciones, None, None, None, []

        else:  # CIPS
            st.markdown('<hr style="border-color:#E2E8F0;margin:0.8rem 0;">', unsafe_allow_html=True)

            # SP sync automático — retorna datos ya cargados (no BytesIO)
            actual_list, historico_list, sp_errors = fetch_cips_results()

            # Mostrar errores de sync si existen
            if sp_errors:
                for err in sp_errors[:3]:
                    st.warning(f"⚠ {err}", icon=None)

            # Upload manual (opcional)
            st.markdown('<p style="font-size:0.75rem;font-weight:600;color:#475569;margin:0.3rem 0 0.3rem;">SUBIR ARCHIVOS ADICIONALES</p>',
                        unsafe_allow_html=True)
            uploaded_cips = st.file_uploader("Excel CIPS", type=["xlsx"],
                                              accept_multiple_files=True,
                                              label_visibility="collapsed",
                                              key="cips_uploader")

            # Agregar archivos subidos manualmente → ACTUAL por defecto
            nombres_sp = {d["tramo"] for d in actual_list + historico_list}
            for f in (uploaded_cips or []):
                try:
                    f.seek(0)
                    d = load_cips_processed(f, "ACTUAL")
                    if d["tramo"] not in nombres_sp:
                        actual_list.append(d)
                except Exception:
                    pass

            # Mostrar en sidebar agrupado
            if actual_list or historico_list:
                st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                for label, lst, color in [("ACTUALES", actual_list, "#D50032"),
                                           ("HISTÓRICOS", historico_list, "#6B7280")]:
                    if not lst: continue
                    st.markdown(f'<p style="font-size:0.67rem;color:{color};font-weight:700;'
                                f'letter-spacing:0.08em;margin:0.7rem 0 0.2rem;">{label}</p>',
                                unsafe_allow_html=True)
                    for d in lst:
                        st.markdown(f"""
                        <div style="background:white;border:1px solid #E2E8F0;border-radius:6px;
                                    padding:0.5rem 0.8rem;margin:3px 0;border-left:3px solid {color};">
                          <div style="font-size:0.82rem;font-weight:600;color:#0F172A;">
                            {d['tramo'][:28]}
                          </div>
                          <div style="font-size:0.72rem;color:#64748B;margin-top:2px;">
                            {d['fecha']} · {len(d['df']):,} pts
                          </div>
                        </div>""", unsafe_allow_html=True)

            if st.button("Refrescar SharePoint", use_container_width=True, key="cips_refresh"):
                fetch_cips_results.clear(); st.rerun()

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
              <p style="color:#64748B;font-size:1.1rem;line-height:1.6;max-width:500px;margin:0 auto;">
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
                    st.markdown('<p style="font-size:0.72rem;color:#888;margin-bottom:0.2rem;">SELECCIONAR PAP</p>',
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
                    st.markdown('<p style="font-size:0.72rem;color:#888;margin-bottom:0.2rem;">SELECCIONAR DCVG</p>',
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
              <p style="color:#64748B;font-size:1.1rem;line-height:1.6;max-width:520px;margin:0 auto;">
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
            st.markdown('<p style="font-size:0.85rem;font-weight:600;color:#475569;margin-bottom:0.4rem;">VER DETALLE DE INSPECCIÓN</p>',
                        unsafe_allow_html=True)
            opts = {f"{'🔴' if d['categoria']=='ACTUAL' else '⬜'} {d['tramo']} — {d['fecha']}": d
                    for d in todos}
            sel_key = st.selectbox("Inspección", list(opts.keys()),
                                   label_visibility="collapsed", key="cips_detail_sel")
            render_cips_dashboard(opts[sel_key])


if __name__ == "__main__":
    main()
