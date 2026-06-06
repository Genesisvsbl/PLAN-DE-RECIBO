import base64
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="Calculadora TRM",
    page_icon="favicon.ico",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================
# FUNCIONES
# =========================
def file_to_base64(path: str):
    file_path = Path(path)
    if not file_path.exists():
        return None
    return base64.b64encode(file_path.read_bytes()).decode("utf-8")


def money(value):
    if value is None:
        return "—"
    text = f"{value:,.2f}"
    text = text.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {text}"


def num(value):
    if value is None:
        return "—"
    text = f"{value:,.2f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def safe_div(a, b):
    if a is None or b is None or b == 0:
        return None
    return a / b


def clear_inputs():
    for key in [
        "trm",
        "unidad_sap",
        "valor_material_sap",
        "valor_unitario_fv",
        "valor_total_fv_factura",
        "unidad_medida",
        "cantidad_total_fv",
    ]:
        st.session_state[key] = None


def limpiar_si_cambia_modo(modo_actual):
    if "modo_anterior" not in st.session_state:
        st.session_state.modo_anterior = modo_actual

    if st.session_state.modo_anterior != modo_actual:
        clear_inputs()
        st.session_state.modo_anterior = modo_actual
        st.rerun()


def calcular_unidad(trm, unidad_sap, valor_unitario_fv, valor_total_fv_factura, unidad_medida, cantidad_total_fv):
    validacion_trm = safe_div(valor_unitario_fv, trm)

    valor_estiba = (
        safe_div(valor_unitario_fv * unidad_medida, unidad_sap)
        if valor_unitario_fv is not None and unidad_medida is not None
        else None
    )

    valor_total_calculado = (
        cantidad_total_fv * valor_estiba
        if cantidad_total_fv is not None and valor_estiba is not None
        else None
    )

    diferencia = (
        valor_total_fv_factura - valor_total_calculado
        if valor_total_fv_factura is not None and valor_total_calculado is not None
        else None
    )

    return validacion_trm, None, valor_estiba, valor_total_calculado, diferencia


def calcular_kg(trm, unidad_sap, valor_unitario_fv, valor_total_fv_factura, unidad_medida, cantidad_total_fv):
    validacion_trm = safe_div(valor_unitario_fv, trm)
    cantidad_recibir = safe_div(cantidad_total_fv, unidad_medida)

    valor_estiba = (
        safe_div(valor_unitario_fv * unidad_medida, unidad_sap)
        if valor_unitario_fv is not None and unidad_medida is not None
        else None
    )

    valor_total_calculado = (
        cantidad_total_fv * valor_unitario_fv
        if cantidad_total_fv is not None and valor_unitario_fv is not None
        else None
    )

    diferencia = (
        valor_total_fv_factura - valor_total_calculado
        if valor_total_fv_factura is not None and valor_total_calculado is not None
        else None
    )

    return validacion_trm, cantidad_recibir, valor_estiba, valor_total_calculado, diferencia


# =========================
# LOGO
# =========================
logo_b64 = file_to_base64("favicon.ico")

if logo_b64:
    logo_html = f'<img src="data:image/x-icon;base64,{logo_b64}" class="logo-img" alt="Logo">'
else:
    logo_html = '<div class="logo-fallback">TRM</div>'


# =========================
# CSS
# =========================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 8% 6%, rgba(227, 6, 19, 0.035), transparent 22%),
        radial-gradient(circle at 92% 10%, rgba(11, 95, 255, 0.035), transparent 22%),
        #ffffff;
}

.block-container {
    max-width: 1180px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}

header[data-testid="stHeader"] {
    background: transparent;
}

.main-shell {
    background: #ffffff;
    border: 1px solid #dbe3ed;
    border-radius: 22px;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.07);
    overflow: hidden;
    margin-bottom: 16px;
}

.topbar {
    height: 70px;
    border-bottom: 1px solid #dbe3ed;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 24px;
    background: #ffffff;
}

.brand {
    display: flex;
    align-items: center;
    gap: 14px;
}

.logo-img {
    width: 42px;
    height: 42px;
    object-fit: contain;
}

.logo-fallback {
    width: 42px;
    height: 42px;
    border-radius: 50%;
    border: 2px solid #e30613;
    color: #e30613;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 900;
    font-size: .7rem;
}

.brand-divider {
    width: 1px;
    height: 34px;
    background: #dbe3ed;
}

.brand-title {
    font-size: 1.08rem;
    font-weight: 900;
    color: #071a33;
    line-height: 1;
}

.brand-subtitle {
    color: #64748b;
    font-size: .74rem;
    margin-top: 4px;
    font-weight: 600;
}

.status {
    background: #ecfdf3;
    color: #027a48;
    border: 1px solid #86efac;
    padding: 8px 14px;
    border-radius: 999px;
    font-size: .72rem;
    font-weight: 900;
}

.hero {
    border: 1px solid #dbe3ed;
    background: linear-gradient(180deg, #ffffff 0%, #fbfdff 100%);
    border-radius: 20px;
    padding: 26px;
    margin: 22px 26px 22px 26px;
}

.eyebrow {
    color: #0b5fff;
    font-size: .68rem;
    font-weight: 900;
    letter-spacing: .18em;
    text-transform: uppercase;
}

.title {
    font-size: clamp(2rem, 4vw, 3rem);
    font-weight: 900;
    letter-spacing: -.06em;
    color: #102033;
    margin-top: 12px;
    margin-bottom: 0;
}

.copy {
    color: #64748b;
    margin-top: 16px;
    max-width: 780px;
    line-height: 1.6;
    font-size: .92rem;
}

.section-box {
    border: 1px solid #dbe3ed;
    border-radius: 16px;
    background: #ffffff;
    padding: 16px;
    margin-bottom: 16px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.035);
}

.section-title {
    font-size: .75rem;
    font-weight: 900;
    letter-spacing: .12em;
    text-transform: uppercase;
    color: #071a33;
    margin-bottom: 14px;
}

div[data-testid="stRadio"] > label {
    font-size: .78rem !important;
    font-weight: 800 !important;
    color: #334155 !important;
}

div[data-testid="stNumberInput"] label {
    font-size: .78rem !important;
    font-weight: 800 !important;
    color: #334155 !important;
}

div[data-testid="stNumberInput"] input {
    border-radius: 10px !important;
    border: 1px solid #dbe3ed !important;
    background: #ffffff !important;
    min-height: 42px;
    color: #e30613 !important;
    font-weight: 900 !important;
    font-size: .92rem !important;
}

div[data-testid="stNumberInput"] input:focus {
    border-color: #e30613 !important;
    box-shadow: 0 0 0 3px rgba(227, 6, 19, .08) !important;
}

.trm-help {
    margin-top: 10px;
    border: 1px solid #dbe3ed;
    background: #f8fafc;
    border-radius: 14px;
    padding: 14px;
    font-size: .82rem;
    line-height: 1.55;
    color: #334155;
}

.trm-help-title {
    color: #e30613;
    font-size: .72rem;
    font-weight: 900;
    text-transform: uppercase;
    letter-spacing: .1em;
    margin-bottom: 6px;
}

.results {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
}

.metric {
    border: 1px solid #dbe3ed;
    background: #ffffff;
    border-radius: 16px;
    padding: 18px;
    min-height: 118px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    box-shadow: 0 4px 12px rgba(15, 23, 42, .04);
}

.metric.ok {
    background: linear-gradient(180deg, #ecfdf3 0%, #d1fadf 100%);
    border: 1px solid #86efac;
}

.metric-label {
    color: #64748b;
    font-size: .66rem;
    font-weight: 900;
    letter-spacing: .14em;
    text-transform: uppercase;
}

.metric-value {
    color: #071a33;
    font-size: clamp(1.2rem, 2vw, 1.7rem);
    font-weight: 900;
    letter-spacing: -.05em;
}

.validation-ok {
    background: #ecfdf3;
    border: 1px solid #86efac;
    color: #027a48;
    padding: 15px 18px;
    border-radius: 14px;
    margin-top: 14px;
    font-weight: 900;
}

.validation-warn {
    background: #fff7ed;
    border: 1px solid #fdba74;
    color: #9a3412;
    padding: 15px 18px;
    border-radius: 14px;
    margin-top: 14px;
    font-weight: 900;
}

.empty-note {
    color: #64748b;
    font-size: .84rem;
    margin-top: 12px;
}

.footer {
    text-align: center;
    color: #94a3b8;
    font-size: .75rem;
    padding: 18px;
}

@media(max-width: 900px) {
    .results {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}

@media(max-width: 560px) {
    .results {
        grid-template-columns: 1fr;
    }

    .topbar {
        height: auto;
        padding: 16px;
        flex-direction: column;
        align-items: flex-start;
        gap: 12px;
    }

    .brand-divider {
        display: none;
    }

    .hero {
        margin: 16px;
        padding: 22px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================
# ENCABEZADO
# =========================
header_html = (
    '<div class="main-shell">'
    '<div class="topbar">'
    '<div class="brand">'
    f'{logo_html}'
    '<div class="brand-divider"></div>'
    '<div>'
    '<div class="brand-title">CalculadoraTRM</div>'
    '<div class="brand-subtitle">Validación financiera SAP / FV</div>'
    '</div>'
    '</div>'
    '<div class="status">Sistema operativo</div>'
    '</div>'
    '<div class="hero">'
    '<div class="eyebrow">Módulo financiero</div>'
    '<h1 class="title">Evaluación profesional de TRM</h1>'
    '<div class="copy">Calculadora para validar facturas contra TRM. Digite los valores de entrada '
    'y el sistema calculará automáticamente los campos de validación.</div>'
    '</div>'
    '</div>'
)

st.markdown(header_html, unsafe_allow_html=True)


# =========================
# TIPO DE VALIDACION
# =========================
with st.container():
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    modo = st.radio(
        "Tipo de validación",
        ["Por unidad", "Por KG"],
        horizontal=True,
        key="modo",
    )
    st.markdown("</div>", unsafe_allow_html=True)

limpiar_si_cambia_modo(modo)


# =========================
# DATOS SAP
# =========================
st.markdown('<div class="section-box"><div class="section-title">Datos SAP</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    trm = st.number_input(
        "TRM del día de la factura ⓘ",
        min_value=0.0,
        value=None,
        placeholder="Digite TRM",
        step=0.01,
        format="%.2f",
        key="trm",
        help=(
            "Transacción ME21N. "
            "Cabecera > Dat.org. "
            "Org compras: 5RPC. "
            "Grupo compras: 006. "
            "Sociedad: BA00. "
            "Entrega factura: fecha del día, moneda y Enter. "
            "Entrega de factura: Tipo de cambio = TRM correcta."
        ),
    )

    st.markdown(
        '<div class="trm-help">'
        '<div class="trm-help-title">Proceso SAP para obtener la TRM</div>'
        'En la transacción <strong>ME21N</strong>, diríjase a '
        '<strong>Cabecera &gt; Dat. org.</strong> y registre '
        '<strong>Org. compras 5RPC</strong>, <strong>Grupo compras 006</strong> '
        'y <strong>Sociedad BA00</strong>. Luego vaya a '
        '<strong>Entrega / Factura</strong>, seleccione la <strong>fecha del día</strong>, '
        'confirme la <strong>moneda</strong> y presione <strong>Enter</strong>. '
        'Finalmente, en <strong>Entrega de factura</strong>, el campo '
        '<strong>Tipo de cambio</strong> mostrará la <strong>TRM correcta</strong>.'
        '</div>',
        unsafe_allow_html=True,
    )

with c2:
    unidad_sap = st.number_input(
        "Unidad de medida SAP",
        min_value=0.0,
        value=None,
        placeholder="Digite unidad SAP",
        step=0.01,
        format="%.2f",
        key="unidad_sap",
    )

with c3:
    valor_material_sap = st.number_input(
        "Valor unidad material SAP",
        min_value=0.0,
        value=None,
        placeholder="Digite valor SAP",
        step=0.01,
        format="%.2f",
        key="valor_material_sap",
    )

st.markdown("</div>", unsafe_allow_html=True)


# =========================
# DATOS FACTURA
# =========================
st.markdown('<div class="section-box"><div class="section-title">Datos de factura</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    valor_unitario_fv = st.number_input(
        "Valor unitario FV",
        min_value=0.0,
        value=None,
        placeholder="Digite valor unitario",
        step=0.01,
        format="%.2f",
        key="valor_unitario_fv",
    )

with c2:
    valor_total_fv_factura = st.number_input(
        "Valor total FV factura",
        min_value=0.0,
        value=None,
        placeholder="Digite total factura",
        step=0.01,
        format="%.2f",
        key="valor_total_fv_factura",
    )

st.markdown("</div>", unsafe_allow_html=True)


# =========================
# DATOS VALIDACION
# =========================
st.markdown('<div class="section-box"><div class="section-title">Datos de validación</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    unidad_medida = st.number_input(
        "Unidad de medida",
        min_value=0.0,
        value=None,
        placeholder="Digite unidad",
        step=0.01,
        format="%.2f",
        key="unidad_medida",
    )

with c2:
    cantidad_total_fv = st.number_input(
        "Cantidad total FV",
        min_value=0.0,
        value=None,
        placeholder="Digite cantidad",
        step=0.01,
        format="%.2f",
        key="cantidad_total_fv",
    )

st.markdown("</div>", unsafe_allow_html=True)


# =========================
# CALCULOS
# =========================
if modo == "Por unidad":
    validacion_trm, cantidad_recibir, valor_estiba, valor_total_calculado, diferencia = calcular_unidad(
        trm,
        unidad_sap,
        valor_unitario_fv,
        valor_total_fv_factura,
        unidad_medida,
        cantidad_total_fv,
    )
else:
    validacion_trm, cantidad_recibir, valor_estiba, valor_total_calculado, diferencia = calcular_kg(
        trm,
        unidad_sap,
        valor_unitario_fv,
        valor_total_fv_factura,
        unidad_medida,
        cantidad_total_fv,
    )

resultado_ok = diferencia is not None and abs(diferencia) <= 1
clase_primera = "metric ok" if resultado_ok else "metric"


# =========================
# RESULTADOS
# =========================
st.markdown('<div class="section-box"><div class="section-title">Resultados calculados</div>', unsafe_allow_html=True)

if modo == "Por unidad":
    resultados_html = (
        '<div class="results">'
        f'<div class="{clase_primera}"><div class="metric-label">Validación FV sobre TRM</div><div class="metric-value">{num(validacion_trm)}</div></div>'
        f'<div class="metric"><div class="metric-label">Valor por estiba</div><div class="metric-value">{money(valor_estiba)}</div></div>'
        f'<div class="metric"><div class="metric-label">Valor total FV calculado</div><div class="metric-value">{money(valor_total_calculado)}</div></div>'
        f'<div class="metric"><div class="metric-label">Diferencia</div><div class="metric-value">{money(diferencia)}</div></div>'
        '</div>'
    )
else:
    resultados_html = (
        '<div class="results">'
        f'<div class="{clase_primera}"><div class="metric-label">Validación FV sobre TRM</div><div class="metric-value">{num(validacion_trm)}</div></div>'
        f'<div class="metric"><div class="metric-label">Cantidad total a recibir</div><div class="metric-value">{num(cantidad_recibir)}</div></div>'
        f'<div class="metric"><div class="metric-label">Valor por estiba</div><div class="metric-value">{money(valor_estiba)}</div></div>'
        f'<div class="metric"><div class="metric-label">Valor total FV calculado</div><div class="metric-value">{money(valor_total_calculado)}</div></div>'
        '</div>'
    )

st.markdown(resultados_html, unsafe_allow_html=True)

if diferencia is None:
    st.markdown(
        '<div class="empty-note">Complete los campos requeridos para generar la validación.</div>',
        unsafe_allow_html=True,
    )
elif abs(diferencia) <= 1:
    st.markdown(
        f'<div class="validation-ok">Perfecto: los valores coinciden. Diferencia: {money(diferencia)}</div>',
        unsafe_allow_html=True,
    )
elif diferencia > 0:
    st.markdown(
        f'<div class="validation-warn">Atención: el cálculo queda por debajo de la factura por {money(diferencia)}.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f'<div class="validation-warn">Atención: el cálculo queda por encima de la factura por {money(abs(diferencia))}.</div>',
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    '<div class="footer">CalculadoraTRM · Sistema profesional de validación</div>',
    unsafe_allow_html=True,
)
