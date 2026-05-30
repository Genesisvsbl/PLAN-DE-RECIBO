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
      header[data-testid="stHeader"], div[data-testid="stToolbar"] { display:none; }
      .block-container { padding:0 !important; max-width:100% !important; }
      .landing {
        min-height:100vh;
        display:grid;
        place-items:center;
        font-family:Segoe UI, Arial, sans-serif;
        color:#061a38;
      }
      .landing-card {
        width:min(760px,92vw);
        background:white;
        border:1px solid #d8e4f4;
        border-radius:18px;
        padding:34px;
        box-shadow:0 24px 60px rgba(15,23,42,.12);
      }
      .landing-badge {
        display:inline-flex;
        align-items:center;
        gap:10px;
        background:#eaf2ff;
        color:#1768f2;
        border-radius:999px;
        padding:8px 12px;
        font-weight:900;
        font-size:13px;
      }
      .landing h1 {
        margin:18px 0 8px;
        font-size:34px;
        line-height:1.05;
        color:#061a38;
      }
      .landing p {
        font-size:16px;
        color:#52657e;
        line-height:1.55;
        margin:0;
      }
      section[data-testid="stFileUploader"] {
        background:#eef4fb;
        padding:24px 16px 18px;
        border-bottom:1px solid #d8e4f4;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_dashboard(dataset: dict) -> None:
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


if "dataset" in st.session_state:
    render_dashboard(st.session_state["dataset"])
else:
    uploaded = st.file_uploader(
        "Importar PLAN DE RECIBO",
        type=["xlsm", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    if uploaded:
        try:
            st.session_state["dataset"] = analyze(uploaded, uploaded.name)
            st.rerun()
        except Exception as exc:
            st.error(f"No pude procesar el archivo: {exc}")
    else:
        st.markdown(
            """
            <div class="landing">
              <section class="landing-card">
                <div class="landing-badge">PLAN DE RECIBO 2026.2</div>
                <h1>Importa tu archivo para generar el dashboard</h1>
                <p>Sube el Excel en el control superior. El sistema genera los modulos de Indicadores, Proveedores, Recibo, Conciliacion y Base.</p>
              </section>
            </div>
            """,
            unsafe_allow_html=True,
        )
