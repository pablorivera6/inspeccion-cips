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
# CIPS — Configuración y estado
# ══════════════════════════════════════════════════════════════════════════════

def _rp(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

EXCEL_INFRA = _rp(os.path.join("data", "Listado de Infraestructura para Cod Informes.xlsx"))
SHAPEFILES  = _rp("shapefiles")

for _k, _v in {
    "cips_res_df": None, "cips_res_bytes": None, "cips_res_name": None,
    "cips_sp_url": None,
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


@st.cache_data(ttl=0)
def cargar_infra():
    return pd.read_excel(EXCEL_INFRA)


def get_shp(distrito, linea):
    if not linea:
        return None
    try:
        df_lineas = cargar_infra()
        fila = df_lineas[(df_lineas["DISTRITO"] == distrito) & (df_lineas["TRAMO"] == linea)]
        if fila.empty:
            return None
        return os.path.join(SHAPEFILES, fila["ID TRAMO"].values[0] + ".shp")
    except Exception:
        return None


def _get_sp_token():
    """Token de cliente (sin login de usuario) para operaciones de escritura en SharePoint."""
    cfg = st.secrets.get("sharepoint", {})
    app_obj = msal.ConfidentialClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        client_credential=cfg["client_secret"],
    )
    result = app_obj.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")


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
    divider()
    pbi_title("Reporte de inspección (PDF)")
    col_pdf_up, col_pdf_sp = st.columns([1, 1])

    with col_pdf_up:
        pdf_file = st.file_uploader("Sube el PDF del reporte", type=["pdf"],
                                     key=f"pdf_up_{d.get('filename','')}", label_visibility="collapsed")

    with col_pdf_sp:
        if d.get("pdf_url"):
            st.link_button("Abrir reporte desde SharePoint", d["pdf_url"],
                           use_container_width=True)

    if pdf_file:
        pdf_bytes = pdf_file.read()
        st.download_button("Descargar reporte", data=pdf_bytes,
                           file_name=pdf_file.name, mime="application/pdf",
                           use_container_width=True)
        b64 = base64.b64encode(pdf_bytes).decode()
        components.html(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            f'width="100%" height="820" style="border:none;border-radius:10px;"></iframe>',
            height=840,
        )
    elif not d.get("pdf_url"):
        st.caption("Sube el PDF del reporte aquí, o asegúrate de que esté en la carpeta de SharePoint con el ID de la inspección en el nombre del archivo.")

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
    divider()
    pbi_title("Reporte de inspección (PDF)")
    col_pdf_up, col_pdf_sp = st.columns([1, 1])

    with col_pdf_up:
        pdf_file = st.file_uploader("Sube el PDF del reporte", type=["pdf"],
                                     key=f"pdf_dcvg_{d.get('filename','')}", label_visibility="collapsed")

    with col_pdf_sp:
        if d.get("pdf_url"):
            st.link_button("Abrir reporte desde SharePoint", d["pdf_url"],
                           use_container_width=True)

    if pdf_file:
        pdf_bytes = pdf_file.read()
        st.download_button("Descargar reporte", data=pdf_bytes,
                           file_name=pdf_file.name, mime="application/pdf",
                           use_container_width=True)
        b64 = base64.b64encode(pdf_bytes).decode()
        components.html(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            f'width="100%" height="820" style="border:none;border-radius:10px;"></iframe>',
            height=840,
        )
    elif not d.get("pdf_url"):
        st.caption("Sube el PDF del reporte aquí, o asegúrate de que esté en la carpeta de SharePoint.")

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
# CIPS — Dashboard de procesamiento
# ══════════════════════════════════════════════════════════════════════════════

def render_cips(distrito, linea, cliente, sp_files=None):
    shp    = get_shp(distrito, linea)
    shp_ok = bool(shp and os.path.exists(shp))
    sp_files = sp_files or []

    st.markdown(f"""
    <div class="main-header">
      <div class="main-header-title">
        Procesamiento CIPS
        <span style="color:#64748B;font-weight:400;margin-left:0.4rem;">| {linea or '—'}</span>
      </div>
      <div class="main-header-meta">
        {cliente} · {distrito} · Corriente Interrumpida
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Carga de archivos y parámetros ────────────────────────────────────────
    col_up, col_param = st.columns([3, 2])

    with col_up:
        st.markdown('<div class="bloque"><div class="bloque-titulo">Archivos de inspección</div>', unsafe_allow_html=True)

        # Archivos desde SharePoint (ya sincronizados)
        if sp_files:
            st.markdown(
                f'<p style="font-size:0.82rem;color:#1B5E20;font-weight:600;margin-bottom:6px;">'
                f'SharePoint — {len(sp_files)} archivo(s) sincronizado(s):</p>',
                unsafe_allow_html=True
            )
            for f in sp_files:
                st.markdown(
                    f'<div style="background:#F0FFF4;border:1px solid #C6F6D5;border-radius:6px;'
                    f'padding:0.35rem 0.7rem;margin:3px 0;font-size:0.8rem;color:#22543D;">'
                    f'{f.name}</div>',
                    unsafe_allow_html=True
                )
            st.markdown('<p style="font-size:0.78rem;color:#64748B;margin-top:8px;">O sube archivos adicionales:</p>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<p style="font-size:0.82rem;color:#64748B;margin-bottom:6px;">Sube los archivos Excel de la inspección:</p>',
                        unsafe_allow_html=True)

        archivos_subidos = st.file_uploader(
            "Archivos Excel CIPS",
            type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed",
            key="cips_uploader"
        )

        # Combinar SP + subidos (evitar duplicados por nombre)
        nombres_vistos = {f.name for f in sp_files}
        archivos_extra = []
        invalidos = []
        for a in (archivos_subidos or []):
            if a.name in nombres_vistos:
                continue
            nombres_vistos.add(a.name)
            try:
                if "Survey Data" in pd.ExcelFile(a).sheet_names:
                    archivos_extra.append(a)
                else:
                    invalidos.append(a.name)
            except:
                invalidos.append(a.name)
        if archivos_extra:
            st.success(f"{len(archivos_extra)} archivo(s) adicional(es) válido(s): {', '.join(a.name for a in archivos_extra)}")
        if invalidos:
            st.warning(f"Sin hoja 'Survey Data' (ignorados): {', '.join(invalidos)}")

        archivos = sp_files + archivos_extra
        st.markdown('</div>', unsafe_allow_html=True)

    with col_param:
        linea_display = linea if linea else "—"
        shp_label  = "Encontrado" if shp_ok else "No encontrado"
        n_archivos = len(archivos)
        sp_sync    = f"{len(sp_files)} de SP" if sp_files else "manual"
        _p = '<p style="margin:0.4rem 0;color:#1a1a1a">'
        rows_html = [f'{_p}<b>Cliente:</b> {cliente}</p>']
        if cliente == "TGI":
            rows_html.append(f'{_p}<b>Distrito:</b> {distrito}</p>')
        rows_html += [
            f'{_p}<b>Tramo:</b> {linea_display}</p>',
            f'{_p}<b>Shapefile:</b> {"✅" if shp_ok else "❌"} {shp_label}</p>',
            f'{_p}<b>Archivos:</b> {n_archivos} ({sp_sync})</p>',
        ]
        st.html(
            '<div class="bloque"><div class="bloque-titulo">Parámetros</div>'
            + "".join(rows_html) + "</div>"
        )

    # ── Botón procesar ─────────────────────────────────────────────────────────
    st.write("")
    _, col_btn, _ = st.columns([2, 1, 2])
    with col_btn:
        procesar = st.button("Procesar inspección", use_container_width=True)

    if procesar:
        archivos_validos = archivos
        if not archivos_validos:
            st.error("Sube al menos un archivo Excel con hoja 'Survey Data' antes de procesar.")
        elif not shp_ok:
            st.error(f"No se encontró el shapefile para el tramo **{linea}**. Verifica que el tramo esté en la lista de infraestructura.")
        else:
            st.session_state.cips_res_df = st.session_state.cips_res_bytes = st.session_state.cips_res_name = None
            st.session_state.cips_sp_url = None

            prog   = st.progress(0)
            estado = st.empty()

            def upd(p, msg):
                prog.progress(p, text=msg)
                estado.caption(msg)

            _ok = False
            with tempfile.TemporaryDirectory() as tmp:
                for a in archivos_validos:
                    a.seek(0)
                    with open(os.path.join(tmp, a.name), "wb") as f:
                        f.write(a.read())
                try:
                    upd(15, "Unificando archivos...")
                    from mod_unificar import ejecutar_unificar
                    unif = ejecutar_unificar(tmp)

                    upd(55, "Calculando PK geométrico (LRS)...")
                    from mod_cips_lrs import ejecutar_cips_lrs
                    salida = ejecutar_cips_lrs(tmp, unif, shp)

                    upd(85, "Cargando resultados...")
                    df_res = pd.read_excel(salida, sheet_name="Survey Data")
                    with open(salida, "rb") as f:
                        xbytes = f.read()

                    ts     = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    nombre = f"CIPS_{(linea or 'SIN_TRAMO').replace(' ','_')}_{ts}.xlsx"

                    st.session_state.cips_res_df    = df_res
                    st.session_state.cips_res_bytes = xbytes
                    st.session_state.cips_res_name  = nombre

                    upd(92, "Subiendo a SharePoint...")
                    try:
                        from mod_cips_sharepoint import subir_a_sharepoint
                        tok = _get_sp_token()
                        if tok:
                            tmp_sp = os.path.join(tmp, nombre)
                            with open(tmp_sp, "wb") as f:
                                f.write(xbytes)
                            sub = f"ocensa/{linea.replace(' ','_')}" if cliente == "OCENSA" \
                                  else datetime.datetime.now().strftime("%Y/%m")
                            url = subir_a_sharepoint(tmp_sp, tok, subcarpeta=sub)
                            st.session_state.cips_sp_url = url
                    except Exception as e_sp:
                        st.warning(f"Procesado OK, pero no se pudo subir a SharePoint: {e_sp}")

                    upd(100, "¡Proceso completado!")
                    _ok = True

                except Exception as e:
                    import traceback
                    prog.empty(); estado.empty()
                    st.error(f"Error en el procesamiento: {e}")
                    with st.expander("Ver detalle del error"):
                        st.code(traceback.format_exc())

            if _ok:
                prog.empty(); estado.empty()
                st.rerun()

    # ── Resultados ─────────────────────────────────────────────────────────────
    if st.session_state.cips_res_df is not None:
        df = st.session_state.cips_res_df
        divider()
        pbi_title("Resultados del procesamiento CIPS")

        total = len(df)
        prot  = int((df["Estado_CP"] == "PROTEGIDO").sum())   if "Estado_CP" in df.columns else 0
        desp  = int((df["Estado_CP"] == "DESPROTEGIDO").sum()) if "Estado_CP" in df.columns else 0
        sobre = int((df["Estado_CP"] == "SOBREPROTEGIDO").sum()) if "Estado_CP" in df.columns else 0
        pct   = f"{round(prot/total*100,1)}%" if total else "—"

        animated_kpi_row([
            ("Puntos totales",   total, "#0F172A"),
            ("Protegidos",       prot,  C_PROT),
            ("Desprotegidos",    desp,  C_SIN),
            ("Sobreprotegidos",  sobre, C_SOBRE),
        ])

        # Gráfica Off mV vs PK
        if "Off_mV_limpio" in df.columns and "PK_real_m" in df.columns:
            divider()
            pbi_title("Potencial OFF (mV) por PK")
            sub = df.dropna(subset=["PK_real_m","Off_mV_limpio"]).copy()
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=sub["PK_real_m"], y=sub["Off_mV_limpio"],
                mode="lines+markers", name="Off mV",
                line=dict(color=C_PROT, width=1.8), marker=dict(size=4)
            ))
            fig.add_hrect(y0=-1200, y1=-850, fillcolor="rgba(21,101,192,0.05)", line_width=0)
            fig.add_hline(y=-850,  line=dict(color="#64B5F6", dash="dash", width=1.2),
                          annotation_text="-850", annotation_position="top left",
                          annotation_font=dict(size=9, color="#64B5F6"))
            fig.add_hline(y=-1200, line=dict(color="#EF5350", dash="dash", width=1.2),
                          annotation_text="-1200", annotation_position="top left",
                          annotation_font=dict(size=9, color="#EF5350"))
            st.plotly_chart(apply_chart(fig, 300, "PK (m)", "mV"), use_container_width=True)

        # Mapa GPS corregido
        if "Lat_corr" in df.columns and "Long_corr" in df.columns:
            mdf = df.dropna(subset=["Lat_corr","Long_corr"])
            if not mdf.empty:
                divider()
                pbi_title("Mapa GPS corregido (sobre ducto)")
                color_col = "Estado_CP" if "Estado_CP" in mdf.columns else None
                color_map = {
                    "PROTEGIDO":      C_PROT,
                    "DESPROTEGIDO":   C_SIN,
                    "SOBREPROTEGIDO": C_SOBRE,
                }
                hover = {"PK_real_m": True, "Off_mV_limpio": True,
                         "Lat_corr": False, "Long_corr": False}
                fig = px.scatter_mapbox(
                    mdf, lat="Lat_corr", lon="Long_corr",
                    color=color_col,
                    color_discrete_map=color_map if color_col else None,
                    hover_data={k: v for k, v in hover.items() if k in mdf.columns},
                    zoom=10, height=380, mapbox_style="open-street-map",
                )
                fig.update_traces(marker=dict(size=6, opacity=0.85))
                fig.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                                   legend=dict(x=0.01,y=0.99,
                                               bgcolor="rgba(255,255,255,0.88)",
                                               borderwidth=1, font_size=10))
                st.plotly_chart(fig, use_container_width=True)

        with st.expander("Vista previa de datos procesados"):
            cols = [c for c in ["PK_real_m","Lat_corr","Long_corr",
                                 "On_mV_limpio","Off_mV_limpio",
                                 "IR_Drop_mV_limpio","Estado_CP",
                                 "Comentario","Anomalia"] if c in df.columns]
            st.dataframe(df[cols].head(300), use_container_width=True, height=300)

        divider()
        pbi_title("Descargar / Subir resultados")
        col_dl, col_sp = st.columns([1, 2])

        with col_dl:
            st.download_button(
                "Descargar Excel procesado",
                data=st.session_state.cips_res_bytes,
                file_name=st.session_state.cips_res_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_sp:
            if st.session_state.cips_sp_url:
                st.markdown(f"""
                <div style="background:linear-gradient(135deg,#1A7A4A,#22A06B);color:white;
                            border-radius:12px;padding:1.2rem 1.5rem;">
                  <b>Subido a SharePoint</b><br>
                  <a href="{st.session_state.cips_sp_url}" target="_blank"
                     style="color:rgba(255,255,255,0.9);font-size:0.82rem;word-break:break-all;">
                    {st.session_state.cips_sp_url}
                  </a>
                </div>
                """, unsafe_allow_html=True)
            else:
                if st.button("Subir a SharePoint", use_container_width=True):
                    with st.spinner("Subiendo..."):
                        try:
                            from mod_cips_sharepoint import subir_a_sharepoint
                            tok = _get_sp_token()
                            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
                                tf.write(st.session_state.cips_res_bytes)
                                tf_path = tf.name
                            sub = datetime.datetime.now().strftime("%Y/%m")
                            url = subir_a_sharepoint(tf_path, tok, subcarpeta=sub)
                            os.unlink(tf_path)
                            st.session_state.cips_sp_url = url
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

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


@st.cache_data(ttl=600, show_spinner="Sincronizando archivos CIPS desde SharePoint...")
def fetch_cips_sharepoint_files():
    if "sharepoint" not in st.secrets:
        return []
    cfg = st.secrets["sharepoint"]
    cips_folder = cfg.get("cips_folder_path", "")
    if not cips_folder:
        return []
    client_id     = cfg.get("client_id")
    client_secret = cfg.get("client_secret")
    tenant_id     = cfg.get("tenant_id")
    tenant_name   = cfg.get("tenant_name")
    site_url      = cfg.get("site_url")
    try:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app_obj   = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        result = app_obj.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            return []
        token    = result["access_token"]
        headers  = {"Authorization": f"Bearer {token}"}
        hostname = f"{tenant_name}.sharepoint.com"
        site_path = site_url.replace(f"https://{hostname}", "")
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}",
            headers=headers
        )
        if not resp.ok:
            return []
        site_id = resp.json().get("id")
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{cips_folder}:/children",
            headers=headers
        )
        if not resp.ok:
            return []
        downloaded = []
        for item in resp.json().get("value", []):
            name  = item.get("name", "")
            d_url = item.get("@microsoft.graph.downloadUrl")
            if not name.endswith(".xlsx") or name.startswith("~") or not d_url:
                continue
            f_resp = requests.get(d_url)
            if not f_resp.ok:
                continue
            f_obj = io.BytesIO(f_resp.content)
            f_obj.name = name
            # Solo incluir si tiene hoja Survey Data
            try:
                if "Survey Data" in pd.ExcelFile(f_obj).sheet_names:
                    f_obj.seek(0)
                    downloaded.append(f_obj)
            except Exception:
                pass
        return downloaded
    except Exception as e:
        st.sidebar.warning(f"SharePoint CIPS: {e}")
        return []


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
            try:
                df_lineas   = cargar_infra()
                st.markdown("**Cliente**")
                cliente = st.selectbox("Cliente", ["TGI","OCENSA"],
                                       label_visibility="collapsed", key="cips_cliente")
                if cliente == "OCENSA":
                    df_c  = df_lineas[df_lineas["DISTRITO"] == "OCENSA"]
                    dist  = "OCENSA"
                    st.markdown("**Tramo**")
                    linea = st.selectbox("Tramo", df_c["TRAMO"].tolist(),
                                         label_visibility="collapsed", key="cips_tramo_oc")
                else:
                    df_c   = df_lineas[df_lineas["DISTRITO"] != "OCENSA"]
                    dists  = sorted(df_c["DISTRITO"].unique())
                    st.markdown("**Distrito**")
                    dist  = st.selectbox("Distrito", dists,
                                         label_visibility="collapsed", key="cips_dist")
                    lineas = df_c[df_c["DISTRITO"] == dist]["TRAMO"].tolist()
                    st.markdown("**Línea**")
                    linea = st.selectbox("Línea", lineas,
                                         label_visibility="collapsed", key="cips_linea")
            except Exception:
                st.warning("No se encontró el archivo de infraestructura.")
                cliente, dist, linea = "TGI", "—", "—"

            st.markdown('<hr style="border-color:#E2E8F0;margin:0.8rem 0;">', unsafe_allow_html=True)

            # ── Sincronización automática desde SharePoint ────────────────────
            sp_cips_files = fetch_cips_sharepoint_files()
            if sp_cips_files:
                st.markdown(
                    f'<p style="font-size:0.78rem;color:#1B5E20;font-weight:600;">'
                    f'SharePoint: {len(sp_cips_files)} archivo(s) sincronizado(s)</p>',
                    unsafe_allow_html=True
                )
                for f in sp_cips_files:
                    st.markdown(
                        f'<div style="background:#F0FFF4;border:1px solid #C6F6D5;border-radius:6px;'
                        f'padding:0.4rem 0.7rem;margin:3px 0;font-size:0.78rem;color:#22543D;">'
                        f'{f.name}</div>',
                        unsafe_allow_html=True
                    )
                if st.button("Refrescar archivos", use_container_width=True, key="cips_refresh"):
                    fetch_cips_sharepoint_files.clear()
                    st.rerun()
            else:
                has_config = bool(st.secrets.get("sharepoint", {}).get("cips_folder_path", ""))
                if has_config:
                    st.caption("Sin archivos en la carpeta CIPS de SharePoint")
                    if st.button("Refrescar", use_container_width=True, key="cips_refresh"):
                        fetch_cips_sharepoint_files.clear()
                        st.rerun()
                else:
                    st.caption("Configura `cips_folder_path` en secrets.toml para sincronizar automáticamente.")

            return modo, None, cliente, dist, linea, sp_cips_files


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
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
        _, _, cliente, distrito, linea, sp_files = result
        render_cips(distrito, linea, cliente, sp_files=sp_files)


if __name__ == "__main__":
    main()
