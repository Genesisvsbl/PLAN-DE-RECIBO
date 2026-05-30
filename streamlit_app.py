import json

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
      header[data-testid="stHeader"] { display:none; }
      div[data-testid="stToolbar"] { display:none; }
      .block-container { padding:0; max-width:100%; }
      .plan-shell {
        min-height:100vh;
        padding:20px 28px 28px;
        box-sizing:border-box;
        font-family:Segoe UI, Arial, sans-serif;
        color:#061a38;
      }
      .plan-top {
        display:flex;
        align-items:flex-start;
        justify-content:space-between;
        gap:24px;
        margin-bottom:18px;
      }
      .plan-brand h1 {
        margin:0 0 8px;
        font-size:34px;
        line-height:1;
        font-weight:950;
        letter-spacing:0;
      }
      .plan-brand span {
        display:inline-flex;
        margin-left:10px;
        padding:5px 9px;
        border-radius:999px;
        background:#dbeafe;
        color:#1768f2;
        font-size:13px;
        font-weight:950;
        vertical-align:middle;
      }
      .plan-brand p {
        margin:0;
        color:#52657e;
        font-size:14px;
        font-weight:700;
      }
      .plan-import-note {
        width:min(920px, 92vw);
        margin:180px auto 0;
        background:#ffffff;
        border:1px solid #d8e4f4;
        border-radius:18px;
        padding:34px 38px;
        box-shadow:0 26px 70px rgba(15,23,42,.12);
      }
      .plan-import-note strong {
        display:inline-flex;
        background:#eaf2ff;
        color:#1768f2;
        border-radius:999px;
        padding:8px 13px;
        font-size:13px;
        letter-spacing:.06em;
        font-weight:950;
        text-transform:uppercase;
        margin-bottom:18px;
      }
      .plan-import-note h2 {
        margin:0 0 12px;
        color:#061a38;
        font-size:38px;
        line-height:1.05;
        font-weight:950;
      }
      .plan-import-note p {
        margin:0;
        color:#52657e;
        font-size:16px;
        line-height:1.45;
        font-weight:650;
      }
      section[data-testid="stFileUploader"] {
        position:absolute;
        top:20px;
        right:28px;
        width:min(560px, calc(100vw - 56px));
        margin:0;
        background:#ffffff;
        border:1px solid #d8e4f4;
        border-radius:14px;
        padding:10px 12px;
        box-shadow:0 16px 36px rgba(15,23,42,.08);
        z-index:5;
      }
      section[data-testid="stFileUploader"] label,
      section[data-testid="stFileUploader"] small {
        display:none !important;
      }
      section[data-testid="stFileUploader"] button {
        background:#1768f2 !important;
        color:white !important;
        border-radius:10px !important;
        border:0 !important;
        font-weight:900 !important;
      }
      .dashboard-loaded section[data-testid="stFileUploader"] {
        display:none !important;
      }
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
    st.markdown(
        """
        <div class="plan-shell">
          <div class="plan-top">
            <div class="plan-brand">
              <h1>Plan de Recibo <span>2026.2</span></h1>
              <p>Importa el archivo para activar los modulos de control.</p>
            </div>
          </div>
          <section class="plan-import-note">
            <strong>Acceso al dashboard</strong>
            <h2>Sube tu archivo PLAN DE RECIBO</h2>
            <p>Cuando cargues el Excel, el sistema abre el dashboard profesional con Indicadores, Proveedores, Recibo, Conciliacion y Base.</p>
          </section>
        </div>
        """,
        unsafe_allow_html=True,
    )
