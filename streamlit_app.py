import json
import io
import urllib.parse
import urllib.request

import streamlit as st
import streamlit.components.v1 as components

from app import HTML, analyze, make_json_safe

DEFAULT_ONEDRIVE_URL = "https://anheuserbuschinbev-my.sharepoint.com/:x:/g/personal/gen_maz_co_win093_gmodelo_com_mx/IQBf1LJMe9tFSa7owOBKh3zIAT4cRN6NDEfwzZw2COA2cBE?email=genesisvsbl%40outlook.com&e=6a9CKG"


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
      div[data-testid="stTextInput"] {
        padding:0 16px;
      }
      .online-note {
        margin:10px 16px 0;
        color:#52657e;
        font:700 13px Segoe UI, Arial, sans-serif;
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


def candidate_download_urls(shared_url: str) -> list[str]:
    url = shared_url.strip()
    if not url:
        return []
    parsed = urllib.parse.urlparse(url)
    clean_url = urllib.parse.urlunparse(parsed._replace(query="", fragment=""))
    candidates = []
    for base in (url, clean_url):
        joiner = "&" if "?" in base else "?"
        candidates.append(f"{base}{joiner}download=1")
    encoded = urllib.parse.quote(url, safe="")
    candidates.append(f"{parsed.scheme}://{parsed.netloc}/download.aspx?SourceUrl={encoded}")
    seen = set()
    return [item for item in candidates if not (item in seen or seen.add(item))]


def fetch_online_workbook(shared_url: str) -> io.BytesIO:
    headers = {"User-Agent": "Mozilla/5.0"}
    errors = []
    for url in candidate_download_urls(shared_url):
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=45) as response:
                content = response.read()
                content_type = response.headers.get("Content-Type", "")
            if content[:2] == b"PK":
                return io.BytesIO(content)
            preview = content[:180].decode("utf-8", errors="ignore").lower()
            if "html" in content_type.lower() or "<html" in preview or "signin" in preview or "login" in preview:
                errors.append("SharePoint devolvio una pagina web/login, no el Excel.")
            else:
                errors.append(f"Respuesta no reconocida: {content_type or 'sin content-type'}")
        except Exception as exc:
            errors.append(str(exc))
    raise RuntimeError(
        "No pude descargar el archivo real desde SharePoint. "
        "El enlace debe permitir descarga directa sin iniciar sesion. "
        + " | ".join(errors[-2:])
    )


if "dataset" in st.session_state:
    render_dashboard(st.session_state["dataset"])
else:
    uploaded = st.file_uploader(
        "Importar PLAN DE RECIBO",
        type=["xlsm", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    remote_url = st.text_input(
        "URL compartida de OneDrive",
        value=DEFAULT_ONEDRIVE_URL,
        placeholder="Pega aqui el enlace compartido del archivo si no lo quieres subir manualmente",
        label_visibility="collapsed",
    )
    st.markdown('<div class="online-note">Archivo en linea: PLAN DE RECIBO 2026.2.xlsm</div>', unsafe_allow_html=True)
    refresh_from_url = st.button("Actualizar PLAN DE RECIBO en linea", type="primary", use_container_width=False)
    if uploaded:
        try:
            st.session_state["dataset"] = analyze(uploaded, uploaded.name)
            st.rerun()
        except Exception as exc:
            st.error(f"No pude procesar el archivo: {exc}")
    elif refresh_from_url and remote_url:
        try:
            file_obj = fetch_online_workbook(remote_url)
            file_name = "PLAN DE RECIBO 2026.2.xlsm"
            st.session_state["dataset"] = analyze(file_obj, file_name)
            st.rerun()
        except Exception as exc:
            st.error(f"No pude leer el archivo en linea. Detalle: {exc}")
    else:
        st.markdown(
            """
            <div class="landing">
              <section class="landing-card">
                <div class="landing-badge">PLAN DE RECIBO 2026.2</div>
                <h1>Importa tu archivo para generar el dashboard</h1>
                <p>Sube el Excel en el control superior o pega un enlace compartido de OneDrive. El sistema genera los modulos de Indicadores, Proveedores, Recibo, Conciliacion y Base.</p>
              </section>
            </div>
            """,
            unsafe_allow_html=True,
        )
