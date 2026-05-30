import json
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components

from app import HTML, analyze, make_json_safe


st.set_page_config(
    page_title="Plan de Recibo",
    page_icon="PR",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .stApp { background:#eef4fb; }
      header[data-testid="stHeader"], div[data-testid="stToolbar"] { display:none; }
      .block-container { padding:0 !important; max-width:100% !important; }
      .streamlit-header {
        position:relative;
        margin:3px 18px 0;
        min-height:104px;
        border-radius:20px 20px 18px 18px;
        background:#061a38;
        color:#fff;
        box-shadow:0 4px 0 #12b981, 0 24px 50px rgba(15,23,42,.18);
        font-family:Segoe UI, Arial, sans-serif;
      }
      .streamlit-header .brand {
        position:absolute;
        left:28px;
        top:24px;
      }
      .streamlit-header h1 {
        margin:0;
        font-size:36px;
        line-height:1;
        font-weight:950;
        letter-spacing:0;
      }
      .streamlit-header .badge {
        display:inline-flex;
        align-items:center;
        margin-left:12px;
        padding:6px 10px;
        border-radius:999px;
        background:#dbeafe;
        color:#1768f2;
        font-size:14px;
        font-weight:950;
        vertical-align:middle;
      }
      .streamlit-header p {
        margin:12px 0 0;
        color:#e7eefb;
        font-size:15px;
        font-weight:850;
      }
      .fake-controls {
        position:absolute;
        right:14px;
        top:14px;
        display:flex;
        gap:12px;
        align-items:center;
      }
      .fake-file {
        width:360px;
        height:50px;
        border-radius:12px;
        background:#fff;
        color:#061a38;
        display:flex;
        align-items:center;
        padding:0 14px;
        font-weight:850;
        font-size:14px;
        overflow:hidden;
      }
      .fake-file span {
        display:inline-block;
        border:1px solid #1f2937;
        padding:5px 9px;
        margin-right:8px;
        font-weight:500;
        background:#f8fafc;
        color:#000;
      }
      .fake-action {
        height:50px;
        border-radius:11px;
        padding:0 20px;
        display:grid;
        place-items:center;
        font-size:15px;
        font-weight:950;
        white-space:nowrap;
      }
      .fake-action.primary { background:#1768f2; color:#fff; }
      .fake-action.secondary { background:#eaf2ff; color:#061a38; }
      .streamlit-meta {
        margin:16px 22px 0;
        color:#52657e;
        font:850 13px Segoe UI, Arial, sans-serif;
      }
      section[data-testid="stFileUploader"] {
        position:absolute;
        top:16px;
        right:536px;
        width:360px;
        height:50px;
        opacity:.01;
        z-index:20;
        overflow:hidden;
      }
      section[data-testid="stFileUploader"] * { cursor:pointer !important; }
      .dashboard-loaded section[data-testid="stFileUploader"] { display:none !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Seleccionar archivo PLAN DE RECIBO",
    type=["xlsm", "xlsx", "xls"],
    label_visibility="collapsed",
)

if uploaded:
    st.markdown('<div class="dashboard-loaded"></div>', unsafe_allow_html=True)
    try:
        dataset = analyze(uploaded, uploaded.name)
        safe_dataset = make_json_safe(dataset)
        html = HTML.replace("let dataset = null;", f"let dataset = {json.dumps(safe_dataset, ensure_ascii=False)};")
        html = html.replace(
            'document.getElementById("demo").onclick = () => upload(true);',
            'document.getElementById("demo").onclick = () => initializeDashboard();',
        )
        html = html.replace(
            'document.getElementById("analyze").onclick = () => upload(false);',
            'document.getElementById("analyze").onclick = () => initializeDashboard();',
        )
        html = html.replace(
            "Sube el archivo PLAN DE RECIBO o usa el archivo de OneDrive detectado en esta maquina.",
            "Archivo cargado en Streamlit. Usa los modulos de control.",
        )
        html = html.replace("</script>", "\nwindow.addEventListener('load', initializeDashboard);\n</script>", 1)
        components.html(html, height=1600, scrolling=True)
    except Exception as exc:
        st.error(f"No pude procesar el archivo: {exc}")
else:
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    st.markdown(
        f"""
        <section class="streamlit-header">
          <div class="brand">
            <h1>Plan de Recibo <span class="badge">2026.2</span></h1>
            <p>Generado: {generated} | Esperando archivo</p>
          </div>
          <div class="fake-controls">
            <div class="fake-file"><span>Seleccionar archivo</span> Sin archivo seleccionados</div>
            <div class="fake-action primary">Importar y generar dashboard</div>
            <div class="fake-action secondary">Usar archivo actual</div>
          </div>
        </section>
        <div class="streamlit-meta">PLAN DE RECIBO 2026.2.xlsm | carga tu archivo para generar los registros</div>
        """,
        unsafe_allow_html=True,
    )
