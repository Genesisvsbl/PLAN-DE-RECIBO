from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import webbrowser
from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

import pandas as pd
import re

APP_HOST = "127.0.0.1"
APP_PORT = int(os.environ.get("PLAN_RECIBO_PORT", "8765"))
TODAY = date.today()
DEFAULT_FILE = Path(r"C:\Users\JOSUE\OneDrive\PLAN DE RECIBO  2026.2.xlsm")
MONTHS_ES = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}


SOURCE_HEADERS = [
    "CODIGO CITA",
    "CATEGORIA",
    "STATUS",
    "SEMANA",
    "FECHA PROGRAMADA",
    "FECHA ESTIMADA ENTREGA",
    "TIPO MATERIAL",
    "ACREEDOR",
    "PROVEEDOR",
    "SKU",
    "MATERIAL",
    "UMB",
    "CANTIDAD PROGRAMADA",
    "TELEFONO",
    "DOCUMENTO",
    "N DOCUMENTO",
    "TIPO ENTREGA",
    "TRANSPORTADORA",
    "REMESA TRANSPORTE",
    "PLACA",
    "FECHA ARRIBO",
    "HORA ARRIBO",
    "FECHA RECIBO",
    "HORA RECIBO",
    "CANTIDAD RECIBIDA",
    "TIQUETE",
    "PESO NETO",
    "NOVEDAD RECIBO",
    "CAUSAL NOVEDAD FISICA",
    "NOVEDAD DOCUMENTAL",
    "PLANEADOR",
    "OBSERVACIONES",
    "MODO",
    "TIPO",
]


def clean_value(value: Any) -> Any:
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        value = value.strip()
        if value.lower() == "(en blanco)":
            return ""
        if value in {"#NAME?", "#REF!", "#DIV/0!", "#VALUE!", "#N/A"}:
            return ""
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def as_date(value: Any) -> date | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        return value.date()
    return None


def json_value(value: Any) -> Any:
    if value is pd.NaT or pd.isna(value):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, float):
        if value != value:
            return ""
        if value.is_integer():
            return int(value)
    return value


def cita_base(value: Any) -> str:
    text = str(value or "").strip()
    match = re.match(r"^(\d{5})(?:-\d+)?$", text)
    return match.group(1) if match else ""


def cita_kpi_key(row: pd.Series) -> str:
    base = cita_base(row.get("CODIGO CITA", ""))
    if base:
        return base
    cita = str(row.get("CODIGO CITA", "") or "").strip().upper()
    if cita == "BTS":
        documento = clean_identifier(row.get("N DOCUMENTO", ""))
        sku = clean_identifier(row.get("SKU", ""))
        fecha = (
            as_date(row.get("FECHA CONTROL"))
            or as_date(row.get("FECHA ESTIMADA ENTREGA"))
            or as_date(row.get("FECHA PROGRAMADA"))
        )
        fallback = documento or sku or str(row.name)
        return f"BTS|{fallback}|{fecha.isoformat() if fecha else ''}"
    return ""


def clean_identifier(value: Any) -> str:
    value = clean_value(value)
    if value == "":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    if re.fullmatch(r"[\d.]+", text):
        text = text.replace(".", "")
    return text


def make_json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): make_json_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [make_json_safe(v) for v in value]
    if isinstance(value, tuple):
        return [make_json_safe(v) for v in value]
    if value is pd.NaT:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if hasattr(value, "item"):
        try:
            return make_json_safe(value.item())
        except Exception:
            pass
    return value


def read_plan(file_obj: io.BytesIO | str | Path) -> pd.DataFrame:
    raw = pd.read_excel(file_obj, sheet_name="PLAN DE RECIBO", header=1, engine="openpyxl")
    df = raw.iloc[:, : len(SOURCE_HEADERS)].copy()
    df.columns = SOURCE_HEADERS
    reprogram_cols = [col for col in raw.columns if "REPROGRAM" in str(col).upper()]
    df["FECHA REPROGRAMADA"] = raw[reprogram_cols[0]] if reprogram_cols else ""
    for col in df.columns:
        df[col] = df[col].map(clean_value)
    df = df[df.apply(lambda row: any(value != "" for value in row), axis=1)].copy()
    return df


def read_conciliation(file_obj: io.BytesIO | str | Path) -> dict[str, Any]:
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    try:
        raw = pd.read_excel(file_obj, sheet_name="CONCILIACION", header=None, engine="openpyxl")
    except Exception:
        return {"date": "", "summary": {"MP": 0, "GR": 0, "BTS": 0, "TOTAL DOC": 0}, "rows": []}

    receipt_date = clean_value(raw.iat[0, 1]) if raw.shape[0] > 0 and raw.shape[1] > 1 else ""
    receipt_date = json_value(receipt_date)
    rows: list[dict[str, Any]] = []
    for _, row in raw.iloc[3:, :4].iterrows():
        tipo = clean_value(row.iloc[0])
        proveedor = clean_value(row.iloc[1])
        documento = clean_value(row.iloc[2])
        cuenta = clean_value(row.iloc[3])
        if not tipo and not proveedor and not documento:
            continue
        rows.append(
            {
                "FECHA RECIBO": receipt_date,
                "TIPO": tipo,
                "PROVEEDOR": proveedor,
                "N DOCUMENTO": documento,
                "Cuenta de TIPO": cuenta if cuenta != "" else 0,
            }
        )

    summary = {"MP": 0, "GR": 0, "BTS": 0, "TOTAL DOC": 0}
    if raw.shape[0] > 3 and raw.shape[1] > 9:
        summary = {
            "MP": clean_value(raw.iat[3, 6]) or 0,
            "GR": clean_value(raw.iat[3, 7]) or 0,
            "BTS": clean_value(raw.iat[3, 8]) or 0,
            "TOTAL DOC": clean_value(raw.iat[3, 9]) or 0,
        }
    if not any(float(v or 0) for v in summary.values()):
        counts = Counter(str(item["TIPO"]) for item in rows)
        summary = {
            "MP": counts.get("MP", 0),
            "GR": counts.get("GR", 0),
            "BTS": counts.get("BTS", 0),
            "TOTAL DOC": len(rows),
        }
    return {"date": receipt_date, "summary": summary, "rows": rows}


def build_conciliation_from_plan(df: pd.DataFrame) -> dict[str, Any]:
    base = df.copy()
    base = base[base["FECHA RECIBO"].notna()].copy()
    base["N DOCUMENTO TXT"] = base["N DOCUMENTO"].map(clean_identifier)
    base = base[base["N DOCUMENTO TXT"].ne("")]
    if base.empty:
        return {"date": "", "summary": {"MP": 0, "GR": 0, "BTS": 0, "TOTAL DOC": 0}, "rows": []}

    def conc_type(row: pd.Series) -> str:
        raw = str(row.get("TIPO", "") or "").strip().upper()
        cita = str(row.get("CODIGO CITA", "") or "").strip().upper()
        if raw in {"MP", "GR", "BTS"}:
            return raw
        if cita == "BTS":
            return "BTS"
        return raw or "MP"

    base["TIPO CONCILIACION"] = base.apply(conc_type, axis=1)
    base["FECHA RECIBO TXT"] = base["FECHA RECIBO"].dt.strftime("%Y-%m-%d")
    grouped = (
        base.groupby(["FECHA RECIBO TXT", "TIPO CONCILIACION", "PROVEEDOR", "N DOCUMENTO TXT"], dropna=False)
        .agg(**{"Cuenta de TIPO": ("SKU", "size")})
        .reset_index()
        .sort_values(["FECHA RECIBO TXT", "TIPO CONCILIACION", "PROVEEDOR", "N DOCUMENTO TXT"])
    )
    rows = [
        {
            "FECHA RECIBO": row["FECHA RECIBO TXT"],
            "TIPO": row["TIPO CONCILIACION"],
            "PROVEEDOR": row["PROVEEDOR"],
            "N DOCUMENTO": row["N DOCUMENTO TXT"],
            "Cuenta de TIPO": int(row["Cuenta de TIPO"]),
        }
        for _, row in grouped.iterrows()
    ]
    summary = {
        "MP": int(grouped.loc[grouped["TIPO CONCILIACION"].eq("MP"), "Cuenta de TIPO"].sum()),
        "GR": int(grouped.loc[grouped["TIPO CONCILIACION"].eq("GR"), "Cuenta de TIPO"].sum()),
        "BTS": int(grouped.loc[grouped["TIPO CONCILIACION"].eq("BTS"), "Cuenta de TIPO"].sum()),
        "TOTAL DOC": int(grouped["N DOCUMENTO TXT"].nunique()),
    }
    return {"date": "", "summary": summary, "rows": rows}


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in ["FECHA PROGRAMADA", "FECHA ESTIMADA ENTREGA", "FECHA REPROGRAMADA", "FECHA ARRIBO", "FECHA RECIBO"]:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    def estado(row: pd.Series) -> str:
        entrega = as_date(row["FECHA ESTIMADA ENTREGA"])
        recibo = as_date(row["FECHA RECIBO"])
        modo = str(row.get("MODO", "")).upper()
        if modo == "CERRADO" or row.get("TIQUETE"):
            return "CERRADO"
        if entrega and entrega < TODAY:
            return "VENCIDO"
        if entrega and entrega == TODAY:
            return "HOY"
        return "ABIERTO"

    def alerta(row: pd.Series) -> str:
        alerts: list[str] = []
        fecha_control = as_date(row.get("FECHA CONTROL"))
        aplica_documentacion = bool(fecha_control and fecha_control <= TODAY)
        es_granel = str(row.get("TIPO", "")).strip().upper() == "GR"
        vehiculo_llego = bool(row.get("TIQUETE")) or as_date(row.get("FECHA RECIBO")) is not None
        if row["ESTADO CONTROL"] == "VENCIDO":
            alerts.append("Pendiente vencido")
        if aplica_documentacion and vehiculo_llego and not row.get("N DOCUMENTO"):
            alerts.append("Sin documento")
        if aplica_documentacion and vehiculo_llego and not row.get("PLACA"):
            alerts.append("Sin placa")
        if str(row.get("NOVEDAD RECIBO", "")).upper() == "SI":
            alerts.append("Con novedad")
        prog = pd.to_numeric(row.get("CANTIDAD PROGRAMADA"), errors="coerce")
        rec = pd.to_numeric(row.get("CANTIDAD RECIBIDA"), errors="coerce")
        if not es_granel and pd.notna(prog) and pd.notna(rec) and rec != 0 and abs(float(prog) - float(rec)) > 0.001:
            alerts.append("Diferencia cantidad")
        return " | ".join(alerts) if alerts else "OK"

    out["MES ENTREGA"] = out["FECHA ESTIMADA ENTREGA"].dt.strftime("%Y-%m").fillna("Sin fecha")
    out["DIA ENTREGA"] = out["FECHA ESTIMADA ENTREGA"].dt.strftime("%Y-%m-%d").fillna("")
    out["FECHA CONTROL"] = out["FECHA REPROGRAMADA"].fillna(out["FECHA ESTIMADA ENTREGA"])
    out["ANO CONTROL"] = out["FECHA CONTROL"].dt.year.fillna("").astype(str).str.replace(".0", "", regex=False)
    out["MES CONTROL"] = out["FECHA CONTROL"].apply(
        lambda value: f"{value.month:02d} - {MONTHS_ES.get(value.month, '')}" if pd.notna(value) else ""
    )
    out["SEMANA CONTROL"] = out["FECHA ESTIMADA ENTREGA"].apply(
        lambda value: f"SEM {value.isocalendar().week}" if pd.notna(value) else ""
    )
    out["SEMANA MATERIAL"] = out["FECHA CONTROL"].apply(
        lambda value: f"{value.year}-W{value.isocalendar().week:02d}" if pd.notna(value) else "Sin fecha"
    )
    out["ESTADO VEHICULO"] = out.apply(
        lambda row: "RECIBIDO"
        if row.get("TIQUETE")
        else ("EN ATENCION" if pd.notna(row.get("FECHA RECIBO")) else "PENDIENTE"),
        axis=1,
    )
    out["ESTADO CONTROL"] = out.apply(estado, axis=1)
    out["ALERTA"] = out.apply(alerta, axis=1)
    out["DOCUMENTACION"] = out.apply(
        lambda row: (
            "Pendiente futuro"
            if (as_date(row.get("FECHA CONTROL")) and as_date(row.get("FECHA CONTROL")) > TODAY)
            else (
                "Completa"
                if row.get("N DOCUMENTO") and (row.get("PLACA") or row.get("ESTADO VEHICULO") == "PENDIENTE")
                else "Incompleta"
            )
        ),
        axis=1,
    )
    out["DIF DIAS RECIBO"] = (out["FECHA RECIBO"] - out["FECHA ESTIMADA ENTREGA"]).dt.days
    out["CUMPLIMIENTO FECHA"] = out["DIF DIAS RECIBO"].apply(
        lambda value: "Pendiente" if pd.isna(value) else ("A tiempo" if value <= 0 else "Tarde")
    )
    out["CANTIDAD PROGRAMADA"] = pd.to_numeric(out["CANTIDAD PROGRAMADA"], errors="coerce").fillna(0)
    out["CANTIDAD RECIBIDA"] = pd.to_numeric(out["CANTIDAD RECIBIDA"], errors="coerce").fillna(0)
    out["ES GRANEL"] = out["TIPO"].astype(str).str.strip().str.upper().eq("GR")
    out["DIF CANTIDAD"] = out["CANTIDAD RECIBIDA"] - out["CANTIDAD PROGRAMADA"]
    out.loc[out["ES GRANEL"], "DIF CANTIDAD"] = 0
    out["CUMPLE CANTIDAD"] = out["DIF CANTIDAD"].abs().apply(lambda value: "Cumple" if value <= 0.001 else "No cumple")
    out.loc[out["ES GRANEL"], "CUMPLE CANTIDAD"] = "No aplica GR"
    out["CITA TRAFFIC"] = out["CODIGO CITA"].apply(lambda value: "Con codigo" if str(value).strip() else "Sin codigo")
    out["CITA BASE"] = out["CODIGO CITA"].apply(cita_base)
    out["CITA KPI"] = out.apply(cita_kpi_key, axis=1)
    return out


def counter_items(series: pd.Series, limit: int = 10, split_alerts: bool = False) -> list[dict[str, Any]]:
    counts: Counter[str] = Counter()
    for value in series.fillna("").astype(str):
        if split_alerts:
            if value == "OK":
                continue
            for part in value.split(" | "):
                if part:
                    counts[part] += 1
        else:
            counts[value or "Sin dato"] += 1
    return [{"name": name, "value": int(value)} for name, value in counts.most_common(limit)]


def qty_items(df: pd.DataFrame, group_col: str, limit: int = 10) -> list[dict[str, Any]]:
    grouped = (
        df.groupby(group_col, dropna=False)["CANTIDAD PROGRAMADA"]
        .sum()
        .sort_values(ascending=False)
        .head(limit)
    )
    return [{"name": str(index) if str(index) else "Sin dato", "value": float(value)} for index, value in grouped.items()]


def focus_material(value: Any, material_text: Any = "") -> str:
    text = f"{value or ''} {material_text or ''}".upper()
    if "AZUCAR" in text or "AZÚCAR" in text:
        return "AZUCAR"
    if "PREFORMA" in text:
        return "PREFORMAS"
    if "LATA" in text:
        return "LATAS"
    return ""


def cita_status_counts(df: pd.DataFrame) -> Counter[str]:
    priority = {"PENDIENTE": 3, "EN ATENCION": 2, "RECIBIDO": 1}
    counts: Counter[str] = Counter()
    valid = df["CITA KPI"].astype(str).str.strip().ne("")
    for _, group in df.loc[valid].groupby("CITA KPI", dropna=False):
        statuses = [str(value or "").strip().upper() for value in group["ESTADO VEHICULO"]]
        status = max(statuses, key=lambda value: priority.get(value, 0)) if statuses else ""
        counts[status or "SIN ESTADO"] += 1
    return counts


def analyze(file_obj: io.BytesIO | str | Path, file_name: str) -> dict[str, Any]:
    raw = read_plan(file_obj)
    df = enrich(raw)
    conciliation = build_conciliation_from_plan(df)
    if not conciliation["rows"]:
        conciliation = read_conciliation(file_obj)

    status_counts = cita_status_counts(df)
    total = int(sum(status_counts.values()))
    total_qty = float(df["CANTIDAD PROGRAMADA"].sum())
    received_qty = float(df["CANTIDAD RECIBIDA"].sum())
    closed = int((df["ESTADO CONTROL"] == "CERRADO").sum())
    open_count = int((df["ESTADO CONTROL"] == "ABIERTO").sum())
    overdue = int((df["ESTADO CONTROL"] == "VENCIDO").sum())
    today_count = int((df["ESTADO CONTROL"] == "HOY").sum())
    due_doc_mask = df["FECHA CONTROL"].dt.date.le(TODAY).fillna(False)
    arrived_mask = df["ESTADO VEHICULO"].ne("PENDIENTE")
    missing_doc = int((due_doc_mask & arrived_mask & df["N DOCUMENTO"].astype(str).str.strip().eq("")).sum())
    missing_plate = int((due_doc_mask & arrived_mask & df["PLACA"].astype(str).str.strip().eq("")).sum())
    due_doc_total = int(due_doc_mask.sum())
    complete_doc_total = int(
        (
            due_doc_mask
            & (df["N DOCUMENTO"].astype(str).str.strip().ne("") | df["ESTADO VEHICULO"].eq("PENDIENTE"))
            & (df["PLACA"].astype(str).str.strip().ne("") | df["ESTADO VEHICULO"].eq("PENDIENTE"))
        ).sum()
    )
    doc_rate = round(100 * complete_doc_total / max(due_doc_total, 1), 1)
    in_attention = int(status_counts.get("EN ATENCION", 0))
    received_vehicle = int(status_counts.get("RECIBIDO", 0))
    pending_vehicle = int(status_counts.get("PENDIENTE", 0))
    cancelled = int(
        df[["STATUS", "MODO", "CATEGORIA"]]
        .astype(str)
        .apply(lambda col: col.str.upper().str.contains("CANCEL", na=False))
        .any(axis=1)
        .sum()
    )
    qty_due = df[~df["ES GRANEL"]].copy()
    qty_programmed = float(qty_due["CANTIDAD PROGRAMADA"].sum())
    qty_received = float(qty_due["CANTIDAD RECIBIDA"].sum())
    qty_rate = round(100 * qty_received / max(qty_programmed, 1), 1)
    today_mask = df["FECHA CONTROL"].dt.date.eq(TODAY)
    today_total = int(df.loc[today_mask, "CITA KPI"].replace("", pd.NA).nunique())
    today_with_code = today_total
    today_code_rate = round(100 * today_with_code / max(today_total, 1), 1)

    month = (
        df.groupby("MES ENTREGA", dropna=False)
        .size()
        .sort_index()
        .rename("value")
        .reset_index()
        .rename(columns={"MES ENTREGA": "name"})
    )

    agenda_start = TODAY
    agenda_end = TODAY + timedelta(days=14)
    agenda_mask = (
        df["FECHA ESTIMADA ENTREGA"].dt.date.between(agenda_start, agenda_end)
        & (df["ESTADO CONTROL"] != "CERRADO")
    )
    agenda_cols = [
        "DIA ENTREGA",
        "CODIGO CITA",
        "ESTADO CONTROL",
        "ESTADO VEHICULO",
        "TIPO MATERIAL",
        "PROVEEDOR",
        "SKU",
        "MATERIAL",
        "CANTIDAD PROGRAMADA",
        "UMB",
        "TRANSPORTADORA",
        "PLACA",
        "N DOCUMENTO",
        "ALERTA",
    ]

    table_cols = [
        "FECHA PROGRAMADA",
        "FECHA ESTIMADA ENTREGA",
        "FECHA REPROGRAMADA",
        "DIA ENTREGA",
        "CODIGO CITA",
        "ESTADO CONTROL",
        "ESTADO VEHICULO",
        "TIPO MATERIAL",
        "PROVEEDOR",
        "SKU",
        "MATERIAL",
        "CANTIDAD PROGRAMADA",
        "UMB",
        "TRANSPORTADORA",
        "PLACA",
        "N DOCUMENTO",
        "ALERTA",
        "CUMPLE CANTIDAD",
        "DIF CANTIDAD",
        "CUMPLIMIENTO FECHA",
    ]

    qty_base = df[~df["ES GRANEL"]].copy()
    provider_qty = (
        qty_base.groupby("PROVEEDOR", dropna=False)
        .agg(
            citas=("PROVEEDOR", "size"),
            programada=("CANTIDAD PROGRAMADA", "sum"),
            recibida=("CANTIDAD RECIBIDA", "sum"),
            cumple=("CUMPLE CANTIDAD", lambda s: int((s == "Cumple").sum())),
        )
        .reset_index()
    )
    if provider_qty.empty:
        provider_qty = pd.DataFrame(columns=["PROVEEDOR", "citas", "programada", "recibida", "cumple"])
    provider_qty["cumplimiento"] = (provider_qty["cumple"] / provider_qty["citas"].clip(lower=1) * 100).round(1)
    provider_qty["diferencia"] = provider_qty["recibida"] - provider_qty["programada"]
    provider_qty = provider_qty.sort_values(["citas", "cumplimiento"], ascending=[False, True]).head(15)

    focus_df = df.copy()
    focus_df["GRUPO MATERIAL"] = focus_df.apply(lambda row: focus_material(row.get("TIPO MATERIAL"), row.get("MATERIAL")), axis=1)
    focus_df = focus_df[focus_df["GRUPO MATERIAL"].ne("")]
    weekly_focus = (
        focus_df.groupby(["SEMANA MATERIAL", "GRUPO MATERIAL"], dropna=False)
        .agg(programada=("CANTIDAD PROGRAMADA", "sum"), recibida=("CANTIDAD RECIBIDA", "sum"), citas=("GRUPO MATERIAL", "size"))
        .reset_index()
        .sort_values(["SEMANA MATERIAL", "GRUPO MATERIAL"])
    )
    weekly_focus["diferencia"] = weekly_focus["recibida"] - weekly_focus["programada"]

    df_for_json = df.copy()
    for col in df_for_json.columns:
        if pd.api.types.is_datetime64_any_dtype(df_for_json[col]):
            df_for_json[col] = df_for_json[col].dt.strftime("%Y-%m-%d").fillna("")
    df_for_json = df_for_json.fillna("")

    filters = {}
    for col in [
        "TIPO MATERIAL",
        "PROVEEDOR",
        "TRANSPORTADORA",
        "ANO CONTROL",
        "MES CONTROL",
        "SEMANA CONTROL",
        "MES ENTREGA",
        "ESTADO CONTROL",
        "ESTADO VEHICULO",
        "DOCUMENTACION",
        "CUMPLE CANTIDAD",
        "CITA TRAFFIC",
    ]:
        filters[col] = sorted([str(v) for v in df_for_json[col].unique() if str(v)])

    return {
        "fileName": file_name,
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "kpis": {
            "total": total,
            "totalQty": total_qty,
            "receivedQty": received_qty,
            "closed": closed,
            "open": open_count,
            "overdue": overdue,
            "today": today_count,
            "inAttention": in_attention,
            "receivedVehicle": received_vehicle,
            "pendingVehicle": pending_vehicle,
            "cancelled": cancelled,
            "qtyRate": qty_rate,
            "todayTrafficTotal": today_total,
            "todayTrafficWithCode": today_with_code,
            "todayTrafficRate": today_code_rate,
            "missingDocPlate": missing_doc + missing_plate,
            "docDueTotal": due_doc_total,
            "docCompleteTotal": complete_doc_total,
            "docRate": doc_rate,
            "alerts": int((df["ALERTA"] != "OK").sum()),
            "onTimeRate": round(100 * int((df["CUMPLIMIENTO FECHA"] == "A tiempo").sum()) / max(total, 1), 1),
        },
        "charts": {
            "month": month.to_dict(orient="records"),
            "status": counter_items(df["ESTADO CONTROL"], 10),
            "materialQty": qty_items(df, "TIPO MATERIAL", 10),
            "providers": counter_items(df["PROVEEDOR"], 10),
            "transport": counter_items(df["TRANSPORTADORA"], 10),
            "alerts": counter_items(df["ALERTA"], 10, split_alerts=True),
            "vehicle": counter_items(df["ESTADO VEHICULO"], 10),
            "providerQty": provider_qty.rename(columns={"PROVEEDOR": "name"}).to_dict(orient="records"),
            "weeklyFocus": weekly_focus.to_dict(orient="records"),
        },
        "filters": filters,
        "agenda": df_for_json.loc[agenda_mask, agenda_cols].head(300).to_dict(orient="records"),
        "conciliation": conciliation,
        "columns": table_cols,
        "rows": df_for_json.to_dict(orient="records"),
    }


def parse_multipart(body: bytes, content_type: str) -> tuple[str, io.BytesIO]:
    import cgi

    environ = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": content_type,
        "CONTENT_LENGTH": str(len(body)),
    }
    form = cgi.FieldStorage(fp=io.BytesIO(body), environ=environ, keep_blank_values=True)
    field = form["file"]
    file_name = Path(field.filename or "archivo.xlsx").name
    data = field.file.read()
    return file_name, io.BytesIO(data)


def csv_response(rows: list[dict[str, Any]], columns: list[str]) -> bytes:
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8-sig")


HTML = r"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Plan de Recibo | Dashboard Python</title>
  <style>
    :root {
      --ink:#0f1b32; --navy:#0b1627; --panel:#ffffff; --panel-2:#f8fbff;
      --slate:#334155; --muted:#64748b; --line:#d5e0ef;
      --blue:#2563eb; --blue-2:#60a5fa; --teal:#0f766e; --green:#15803d; --amber:#d97706; --red:#b91c1c;
      --bg:#eef3fa;
    }
    * { box-sizing: border-box; }
    body { margin:0; font-family:"Segoe UI", "Inter", Arial, sans-serif; color:var(--ink); background:#f6f8fc; overflow-x:hidden; }
    header { display:none; }
    h1 { margin:0; font-size:0; }
    h1::before { content:""; display:none; }
    h1::after { content:""; display:none; }
    header p { display:none; }
    #status { display:none; }
    main { padding:18px 18px 20px; width:100%; max-width:1900px; margin:0 auto; min-height:100vh; }
    .nav-tabs { display:flex; flex-direction:row; gap:8px; padding:8px; margin:10px 0 16px; background:#ffffff; border:1px solid #e2e9f3; border-radius:10px; box-shadow:0 12px 30px rgba(15,23,42,.055); }
    .nav-tabs.home-mode { display:flex; }
    .nav-tabs button { width:auto; background:#f8fbff; color:#273a55; box-shadow:none; border:1px solid #dbe6f4; padding:10px 14px 10px 38px; text-align:left; border-radius:7px; font-size:12px; font-weight:800; position:relative; }
    .nav-tabs button::before { content:""; position:absolute; left:12px; top:50%; width:16px; height:16px; transform:translateY(-50%); border-radius:5px; background:#eaf2ff; border:1px solid #bcd3f5; }
    .nav-tabs button::after { content:""; position:absolute; left:17px; top:50%; width:6px; height:6px; transform:translateY(-50%); border:2px solid #1768f2; border-radius:2px; }
    .nav-tabs button:nth-child(2)::before { background:#e8f1ff; border-color:#bcd3f5; }
    .nav-tabs button:nth-child(2)::after { border-color:#1768f2; border-radius:2px; }
    .nav-tabs button:nth-child(3)::before { background:#e6f8ee; border-color:#bbebcf; }
    .nav-tabs button:nth-child(3)::after { border-color:#16a34a; border-radius:50%; }
    .nav-tabs button:nth-child(4)::before { background:#e1f8fc; border-color:#ace7f0; }
    .nav-tabs button:nth-child(4)::after { border-color:#0891b2; border-radius:2px; }
    .nav-tabs button:nth-child(5)::before { background:#efe8ff; border-color:#d8c7ff; }
    .nav-tabs button:nth-child(5)::after { border-color:#7c3aed; border-radius:2px; }
    .nav-tabs button:nth-child(6)::before { background:#fff0e4; border-color:#fed7aa; }
    .nav-tabs button:nth-child(6)::after { border-color:#f97316; border-radius:2px; }
    .nav-tabs button.active { background:#1768f2; color:white; border-color:#1768f2; }
    .nav-tabs button.active::before { background:rgba(255,255,255,.18); border-color:rgba(255,255,255,.40); }
    .nav-tabs button.active::after { border-color:#ffffff; }
    .upload { position:absolute; right:18px; top:18px; width:min(650px, 48vw); background:transparent; border:0; border-radius:0; padding:0; display:grid; grid-template-columns:1fr auto auto; align-items:center; gap:12px; box-shadow:none; }
    .upload input { border:1px solid #dbe4f0; padding:11px 12px; border-radius:7px; background:#ffffff; width:100%; color:#36465d; font-size:12px; }
    button { border:0; border-radius:7px; padding:11px 16px; font-weight:800; cursor:pointer; background:#0f64e8; color:white; box-shadow:0 8px 18px rgba(37,99,235,.18); font-size:12px; }
    button.secondary { background:#e7eef8; color:var(--ink); box-shadow:none; }
    .module-toolbar { display:flex; justify-content:flex-end; margin:-6px 0 12px; }
    .module-toolbar.page-hidden { display:none; }
    .export-module { display:inline-flex; align-items:center; gap:8px; background:#061a38; color:#fff; border:1px solid #0a5df0; }
    button:disabled { opacity:.55; cursor:wait; }
    .meta { color:#5d6f88; font-size:11px; margin:4px 0 12px; font-weight:700; }
    .kpis { display:grid; grid-template-columns:repeat(6, minmax(0, 1fr)); gap:12px; margin:10px 0 12px; }
    .kpis.page-hidden, .grid.page-hidden { display:none !important; }
    .mini-kpis { grid-template-columns:repeat(4, minmax(120px, 1fr)); margin:6px 0 12px; }
    .mini-kpis .kpi { min-height:78px; padding-left:58px; }
    .kpi { position:relative; background:#ffffff; border:1px solid #e3eaf4; border-radius:8px; padding:15px 10px 12px 66px; min-height:86px; box-shadow:0 12px 24px rgba(15,23,42,.055); overflow:hidden; }
    .kpi.qty-wide { grid-column:auto; }
    .kpi::before, .kpi::after { content:none; }
    .kpi .kpi-icon { position:absolute; left:14px; top:16px; width:42px; height:42px; display:grid; place-items:center; border-radius:9px; background:#eaf2ff; color:#1768f2; }
    .kpi.teal .kpi-icon { background:#def7ec; color:#16a34a; }
    .kpi.green .kpi-icon { background:#dff7f5; color:#0891b2; }
    .kpi.amber .kpi-icon { background:#fff0d6; color:#f59e0b; }
    .kpi.red .kpi-icon { background:#fee7ea; color:#ef4444; }
    .kpi.purple .kpi-icon { background:#f1e8ff; color:#7c3aed; }
    .kpi svg { width:21px; height:21px; stroke:currentColor; fill:none; stroke-width:2.4; stroke-linecap:round; stroke-linejoin:round; }
    .kpi strong { display:block; font-size:clamp(17px, 1.05vw, 23px); line-height:1.05; margin-top:8px; white-space:nowrap; overflow:visible; text-overflow:clip; }
    .kpi.qty-wide strong { font-size:clamp(15px, .9vw, 20px); letter-spacing:-.01em; }
    .kpi span { color:#52657e; font-weight:900; font-size:clamp(8.2px, .52vw, 10px); text-transform:none; letter-spacing:0; line-height:1.18; }
    .kpi small { display:block; margin-top:6px; color:#64748b; font-size:8px; font-weight:800; }
    .filters { background:#ffffff; border:1px solid #e2e9f3; border-radius:8px; padding:13px; display:grid; grid-template-columns:repeat(9, minmax(110px, 1fr)); gap:9px; margin-bottom:12px; box-shadow:0 12px 26px rgba(15,23,42,.05); }
    label { font-size:10px; color:var(--slate); font-weight:800; display:block; margin-bottom:4px; }
    select, input[type="search"], input[type="date"] { width:100%; border:1px solid #c7d5e8; border-radius:7px; padding:8px 9px; background:white; color:var(--ink); font-size:11px; }
    .filter-actions { display:flex; align-items:end; justify-content:flex-end; }
    .clear-filters { width:100%; min-height:34px; border:1px solid #bfdbfe; background:#eef6ff; color:#1768f2; box-shadow:none; }
    .grid { display:grid; grid-template-columns:repeat(12, minmax(0, 1fr)); gap:12px; }
    .panel { background:#ffffff; border:1px solid #e2e9f3; border-radius:8px; padding:12px; min-height:238px; box-shadow:0 14px 28px rgba(15,23,42,.052); grid-column:span 4; }
    .dashboard-chart { grid-column:span 4; min-height:292px; }
    .panel.page-hidden { display:none; }
    .filters.page-hidden { display:none; }
    .home { display:block; margin-top:0; }
    .home.page-hidden { display:none; }
    .home-layout { display:block; }
    .home-card { background:#ffffff; border:1px solid #e2e9f3; border-radius:8px; box-shadow:0 16px 36px rgba(15,23,42,.06); padding:22px 24px; min-height:calc(100vh - 170px); }
    .home-eyebrow { color:#0b72c6; font-weight:800; letter-spacing:.07em; text-transform:none; font-size:13px; }
    .home-title { margin:0 0 10px; font-size:22px; line-height:1.1; color:#142038; font-weight:800; }
    .home-copy { color:#65758d; font-size:12px; margin:0 0 20px; line-height:1.55; }
    .module-grid { display:grid; grid-template-columns:repeat(3, minmax(240px, 1fr)); gap:16px; margin-top:18px; }
    .module-card { position:relative; border:1px solid #e0e7f2; border-radius:8px; padding:22px 20px 16px 76px; background:#ffffff; min-height:116px; cursor:pointer; transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; display:flex; flex-direction:column; }
    .module-card:hover { transform:translateY(-2px); box-shadow:0 16px 34px rgba(15,23,42,.08); border-color:#bfd2ee; }
    .module-icon { position:absolute; left:22px; top:22px; width:42px; height:42px; display:grid; place-items:center; border-radius:10px; background:#eaf2ff; color:#2563eb; font-weight:900; font-size:13px; margin:0; letter-spacing:.01em; line-height:1; }
    .module-icon::before { content:none; }
    .module-card:nth-child(2) .module-icon { background:#e4f8ed; color:#16a34a; }
    .module-card:nth-child(3) .module-icon { background:#e0f7fb; color:#0891b2; }
    .module-card:nth-child(4) .module-icon { background:#efe7ff; color:#7c3aed; }
    .module-card:nth-child(5) .module-icon { background:#fff0e4; color:#f97316; }
    .module-card h3 { margin:0 0 8px; font-size:15px; color:#142038; font-weight:800; }
    .module-card p { margin:0 0 14px; color:#65758d; font-size:12px; line-height:1.45; }
    .module-metric { display:none; }
    .module-link { color:#0b64d8; font-weight:800; font-size:12px; border:1px solid #d9e5f4; width:max-content; padding:6px 10px; border-radius:5px; }
    .side-list { display:none; }
    .side-card { position:relative; background:#ffffff; border:1px solid #e2e9f3; border-radius:8px; padding:28px 28px; box-shadow:0 16px 36px rgba(15,23,42,.06); min-height:218px; }
    .side-card h3 { margin:0 0 22px; font-size:18px; color:#142038; font-weight:800; padding-left:42px; }
    .side-card h3::before { content:""; position:absolute; left:28px; top:26px; width:28px; height:28px; border-radius:6px; background:#1768f2; }
    .side-card ul { margin:0; padding:0; list-style:none; display:grid; gap:18px; color:#304057; font-weight:700; font-size:12px; }
    .side-card li::before { content:""; display:inline-block; width:16px; height:16px; border-radius:50%; background:#1768f2; margin-right:12px; vertical-align:-3px; }
    .side-card .primary-action { display:inline-block; margin-top:18px; border-radius:7px; padding:11px 18px; color:white; background:#1768f2; font-weight:800; box-shadow:0 12px 22px rgba(37,99,235,.22); cursor:pointer; font-size:12px; }
    .top-title { min-height:50px; padding-right:min(700px, 52vw); }
    .top-title h2 { margin:0 0 6px; font-size:26px; line-height:1; font-weight:900; color:#12203a; letter-spacing:-.03em; }
    .top-title .badge { display:inline-block; margin-left:8px; padding:5px 8px; border-radius:999px; background:#dbeafe; color:#1768f2; font-weight:900; font-size:12px; vertical-align:middle; }
    .top-title p { margin:0; color:#64748b; font-size:13px; font-weight:500; }
    .panel h2 { margin:0 0 8px; font-size:13px; letter-spacing:.01em; color:#11203a; }
    .subhead { margin:14px 0 8px; font-size:13px; color:var(--slate); text-transform:uppercase; letter-spacing:.04em; }
    .wide { grid-column:1 / -1; }
    svg { width:100%; height:205px; overflow:visible; }
    .dashboard-chart svg { height:250px; }
    .table-wrap { overflow:auto; max-height:540px; border:1px solid var(--line); border-radius:10px; }
    table { border-collapse:collapse; width:100%; min-width:1100px; background:white; }
    th, td { border-bottom:1px solid #e2e8f0; padding:9px 10px; text-align:left; font-size:12px; vertical-align:top; }
    th { position:sticky; top:0; background:#0b1930; color:white; z-index:1; }
    tr:hover td { background:#f8fafc; }
    .pill { display:inline-block; padding:4px 8px; border-radius:999px; font-weight:700; font-size:11px; }
    .pill.CERRADO { background:#dcfce7; color:#166534; } .pill.ABIERTO { background:#dbeafe; color:#1d4ed8; }
    .pill.VENCIDO { background:#fee2e2; color:#991b1b; } .pill.HOY { background:#fef3c7; color:#92400e; }
    .pill.RECIBIDO { background:#dcfce7; color:#166534; } .pill.EN_ATENCION { background:#fef3c7; color:#92400e; } .pill.PENDIENTE { background:#e2e8f0; color:#334155; }
    .actions { display:flex; justify-content:space-between; align-items:center; margin:10px 0; gap:10px; }
    .mini-filters { display:flex; gap:10px; align-items:end; flex-wrap:wrap; }
    .mini-filters label { margin:0; min-width:160px; }
    .compact { max-height:300px; margin-top:12px; }
    .provider-cause-card { min-height:520px; grid-column:span 6; }
    .provider-cause-card h2 { display:flex; align-items:center; gap:10px; font-size:16px; margin:0 0 14px; }
    .provider-cause-card h2 .module-badge { width:34px; height:34px; border-radius:9px; display:grid; place-items:center; background:#eaf2ff; color:#1768f2; }
    .provider-cause-card h2 .module-badge svg { width:18px; height:18px; stroke:currentColor; fill:none; stroke-width:2.4; }
    .provider-cause-card svg { height:230px; margin-top:4px; }
    .cause-chart-scroll { height:230px; overflow-y:auto; overflow-x:hidden; padding-right:6px; }
    .cause-chart-scroll svg { height:auto; min-height:220px; }
    .provider-cause-card .compact { max-height:210px; margin-top:0; }
    .provider-cause-card table { min-width:0; table-layout:fixed; }
    .provider-cause-card th, .provider-cause-card td { padding:5px 6px; font-size:9px; white-space:normal; overflow-wrap:anywhere; line-height:1.18; }
    .provider-cause-card .actions { margin:0 0 8px; align-items:flex-start; display:block; }
    .provider-cause-card .mini-filters { display:grid; grid-template-columns:repeat(6, minmax(0, 1fr)); gap:8px; align-items:end; }
    .provider-cause-card .mini-filters label { min-width:0; font-size:9px; }
    .provider-cause-card .date-mode-strip { display:flex; align-items:center; gap:10px; margin:0 0 10px; padding:8px 10px; border:1px solid #bfdbfe; border-radius:12px; background:#eff6ff; color:#082552; }
    .provider-cause-card .date-mode-strip strong { font-size:11px; text-transform:uppercase; letter-spacing:.04em; }
    .provider-cause-card .date-mode-strip select { width:160px; height:34px; border:1px solid #9ec5fe; border-radius:9px; background:#fff; color:#082552; font-weight:900; padding:0 10px; }
    .provider-cause-card .date-mode-strip span { font-size:11px; font-weight:800; color:#355070; }
    .cause-filter-btn { height:32px; padding:8px 12px; border-radius:7px; }
    .cause-stats { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:10px; margin:14px 0; }
    .cause-stat { position:relative; min-height:62px; border:1px solid #e2e9f3; border-radius:8px; padding:10px 8px 8px 48px; background:#fff; box-shadow:0 8px 18px rgba(15,23,42,.035); }
    .cause-stat i { position:absolute; left:12px; top:15px; width:28px; height:28px; border-radius:8px; display:grid; place-items:center; background:#eaf2ff; color:#1768f2; }
    .cause-stat i svg { width:15px; height:15px; stroke:currentColor; fill:none; stroke-width:2.4; }
    .cause-stat strong { display:block; font-size:18px; line-height:1; color:#0f1b32; margin-top:4px; }
    .cause-stat span { color:#52657e; font-size:8px; font-weight:900; text-transform:uppercase; letter-spacing:.02em; }
    .cause-stat small { display:block; color:#64748b; font-size:9px; margin-top:4px; font-weight:700; }
    .cause-chart-box, .cause-table-box { border:1px solid #e2e9f3; border-radius:8px; padding:10px; background:#fff; margin-top:10px; }
    .cause-box-title { display:flex; justify-content:space-between; align-items:center; margin:0 0 6px; font-size:11px; font-weight:900; color:#142038; }
    .cause-footer { display:flex; justify-content:space-between; align-items:center; padding:7px 2px 0; color:#64748b; font-size:9px; font-weight:700; }
    .focus-section svg { height:340px; }
    .focus-top { display:grid; grid-template-columns:1fr 1.18fr; gap:12px; align-items:stretch; margin-top:8px; }
    .focus-chart-card, .focus-summary-card { border:1px solid #e2e9f3; border-radius:8px; background:#fff; padding:12px; box-shadow:0 10px 20px rgba(15,23,42,.035); }
    .focus-chart-card svg { height:300px; }
    .focus-summary-card .subhead { margin:0 0 8px; text-transform:none; letter-spacing:0; font-size:13px; color:#142038; }
    .focus-stats { display:grid; grid-template-columns:repeat(5, minmax(0, 1fr)); gap:8px; margin-top:10px; }
    .focus-stat { min-height:76px; border:1px solid #e2e9f3; border-radius:8px; background:#fff; padding:10px 8px; text-align:center; box-shadow:0 8px 18px rgba(15,23,42,.035); }
    .focus-stat i { width:30px; height:30px; border-radius:8px; display:grid; place-items:center; margin:0 auto 6px; background:#eaf2ff; color:#1768f2; }
    .focus-stat i svg { width:17px; height:17px; stroke:currentColor; fill:none; stroke-width:2.4; }
    .focus-stat span { display:block; color:#52657e; font-size:8px; font-weight:900; text-transform:uppercase; }
    .focus-stat strong { display:block; color:#0f1b32; font-size:17px; line-height:1.05; margin-top:5px; }
    .focus-stat.red strong { color:#dc2626; }
    .focus-stat.green strong { color:#15803d; }
    .focus-stat.amber strong { color:#d97706; }
    .focus-section .table-wrap { max-height:none; overflow:visible; }
    .focus-section table { min-width:0; table-layout:fixed; }
    .focus-section th, .focus-section td { padding:6px 7px; font-size:9.2px; line-height:1.18; white-space:normal; overflow-wrap:anywhere; vertical-align:middle; }
    .focus-section .compact { max-height:none; }
    .conciliation-panel .table-wrap { max-height:none; overflow:visible; }
    .conciliation-panel table { min-width:0; table-layout:fixed; }
    .conciliation-panel th, .conciliation-panel td { padding:7px 8px; font-size:10.5px; line-height:1.2; white-space:normal; overflow-wrap:anywhere; vertical-align:middle; }
    .dashboard-detail { grid-column:1 / -1; min-height:0; }
    .dashboard-detail .table-wrap { max-height:none; overflow:visible; }
    .dashboard-detail table { min-width:0; table-layout:fixed; }
    .dashboard-detail th, .dashboard-detail td { padding:5px 6px; font-size:8px; line-height:1.12; white-space:normal; overflow-wrap:anywhere; vertical-align:middle; }
    .dashboard-detail th:nth-child(1), .dashboard-detail td:nth-child(1) { width:72px; }
    .dashboard-detail th:nth-child(2), .dashboard-detail td:nth-child(2) { width:46px; text-align:center; }
    .dashboard-detail th:nth-child(3), .dashboard-detail td:nth-child(3),
    .dashboard-detail th:nth-child(4), .dashboard-detail td:nth-child(4) { width:54px; text-align:center; }
    .dashboard-detail th:nth-child(5), .dashboard-detail td:nth-child(5) { width:135px; }
    .dashboard-detail th:nth-child(6), .dashboard-detail td:nth-child(6) { width:72px; }
    .dashboard-detail th:nth-child(7), .dashboard-detail td:nth-child(7) { width:58px; }
    .dashboard-detail th:nth-child(9), .dashboard-detail td:nth-child(9) { width:62px; text-align:right; }
    .dashboard-detail th:nth-child(10), .dashboard-detail td:nth-child(10) { width:62px; text-align:right; }
    .dashboard-detail th:nth-child(11), .dashboard-detail td:nth-child(11) { width:50px; text-align:center; }
    .dashboard-detail th:nth-child(12), .dashboard-detail td:nth-child(12) { width:58px; text-align:center; }
    .dashboard-detail th:nth-child(13), .dashboard-detail td:nth-child(13) { width:52px; text-align:center; }
    .dashboard-detail th:nth-child(14), .dashboard-detail td:nth-child(14) { width:120px; }
    .dashboard-detail th:nth-child(9),
    .dashboard-detail th:nth-child(10),
    .dashboard-detail th:nth-child(11),
    .dashboard-detail th:nth-child(13) {
      white-space:normal;
      line-height:1.05;
      overflow-wrap:normal;
      word-break:normal;
    }
    .daily-detail { min-height:0; }
    .daily-detail .table-wrap { max-height:none; overflow:visible; }
    .daily-detail table { min-width:0; table-layout:fixed; }
    .daily-detail th, .daily-detail td { padding:5px 6px; font-size:8.2px; line-height:1.15; white-space:normal; overflow-wrap:anywhere; vertical-align:middle; }
    .status-chip { display:inline-block; padding:3px 6px; border-radius:999px; font-size:7.5px; font-weight:900; text-transform:uppercase; }
    .status-chip.CERRADO { background:#dcfce7; color:#15803d; border:1px solid #86efac; }
    .status-chip.ABIERTO { background:#dbeafe; color:#1d4ed8; border:1px solid #93c5fd; }
    .status-chip.HOY, .status-chip.EN_ATENCION { background:#fef3c7; color:#b45309; border:1px solid #fcd34d; }
    .status-chip.VENCIDO, .status-chip.CANCELADO { background:#fee2e2; color:#b91c1c; border:1px solid #fca5a5; }
    .doc-ok { display:inline-grid; place-items:center; width:18px; height:18px; border-radius:50%; color:#16a34a; border:1px solid #86efac; font-weight:900; }
    .doc-bad { display:inline-grid; place-items:center; width:18px; height:18px; border-radius:50%; color:#dc2626; border:1px solid #fca5a5; font-weight:900; }
    tr.qty-alert td { background:#fff1f2; border-bottom-color:#fecdd3; color:#7f1d1d; }
    tr.qty-alert td:first-child { border-left:3px solid #dc2626; }
    .detail-count { color:#1768f2; background:#eef5ff; border:1px solid #dbeafe; border-radius:6px; padding:6px 10px; font-size:10px; font-weight:900; }
    .empty { padding:42px; text-align:center; color:var(--muted); }
    .bar-label { font-size:14px; fill:#24364f; font-weight:850; }
    .bar-value { font-size:14px; fill:#0f1b32; font-weight:900; }
    .cause-axis-label { font-size:8px; fill:#475569; }
    .cause-value { font-size:8px; fill:#334155; }
    .axis { stroke:#cbd5e1; stroke-width:1; }
    /* ControlHub Blue BI visual pass */
    body {
      background:
        linear-gradient(180deg, #eaf2fb 0%, #f7faff 30%, #eef4fb 100%);
      color:#071a33;
    }
    main { max-width:1880px; padding:16px 18px 24px; }
    .top-title {
      min-height:86px;
      padding:18px min(720px, 50vw) 16px 22px;
      border-radius:16px;
      background:linear-gradient(135deg, #061833 0%, #082b54 58%, #0b5a94 100%);
      border:1px solid rgba(255,255,255,.10);
      box-shadow:0 18px 42px rgba(7, 26, 51, .22);
      position:relative;
      overflow:hidden;
    }
    .top-title::after {
      content:"";
      position:absolute;
      inset:auto 0 0 0;
      height:3px;
      background:linear-gradient(90deg, #1d75ff, #11c5e8, #18a058);
    }
    .top-title h2 { color:#fff; font-size:28px; letter-spacing:-.02em; }
    .top-title .badge { background:#dbeafe; color:#0b5de8; }
    .top-title p { color:#c8d8ee; font-weight:700; }
    .upload { top:26px; right:32px; }
    .upload input {
      border-color:#d8e4f5;
      box-shadow:0 10px 24px rgba(2, 8, 23, .10);
      font-weight:700;
    }
    button {
      border-radius:10px;
      background:linear-gradient(135deg, #0b63f6, #0047bd);
      box-shadow:0 12px 24px rgba(11,99,246,.22);
    }
    button.secondary { background:#eef4fb; color:#082145; border:1px solid #d8e4f5; }
    .meta { margin:10px 4px 12px; color:#526b8d; }
    .nav-tabs {
      margin:14px 0 14px;
      padding:10px;
      border-radius:15px;
      border:1px solid #d8e4f5;
      box-shadow:0 14px 34px rgba(7,26,51,.08);
      background:rgba(255,255,255,.92);
    }
    .nav-tabs button {
      border-radius:11px;
      padding:12px 18px 12px 42px;
      color:#19324f;
      background:#f8fbff;
      border-color:#d7e5f7;
      font-size:13px;
    }
    .nav-tabs button.active {
      background:linear-gradient(135deg, #0b63f6, #003f9e);
      border-color:#0b63f6;
      box-shadow:0 12px 22px rgba(11,99,246,.20);
    }
    .kpis { gap:12px; margin:12px 0 14px; }
    .kpi {
      min-height:96px;
      border-radius:15px;
      border:1px solid #dce7f5;
      background:linear-gradient(180deg, #fff 0%, #f8fbff 100%);
      box-shadow:0 16px 36px rgba(7,26,51,.08);
      padding:17px 12px 14px 74px;
    }
    .kpi::after {
      content:"";
      display:block;
      position:absolute;
      left:0;
      top:0;
      bottom:0;
      width:4px;
      background:#0b63f6;
    }
    .kpi.teal::after, .kpi.green::after { background:#16a34a; }
    .kpi.amber::after { background:#f59e0b; }
    .kpi.red::after { background:#ef4444; }
    .kpi.purple::after { background:#7c3aed; }
    .kpi .kpi-icon {
      left:16px;
      top:18px;
      width:44px;
      height:44px;
      border-radius:13px;
      background:#e8f1ff;
      color:#0b63f6;
      box-shadow:inset 0 0 0 1px rgba(11,99,246,.06);
    }
    .kpi strong { font-size:clamp(19px, 1.18vw, 26px); color:#071a33; }
    .kpi span { color:#435a78; font-size:clamp(8.8px, .55vw, 10.5px); text-transform:uppercase; }
    .kpi small { color:#6b7f9a; font-size:8.5px; }
    .filters {
      border-radius:15px;
      border:1px solid #dce7f5;
      box-shadow:0 16px 34px rgba(7,26,51,.06);
      background:#ffffff;
      padding:16px;
    }
    label { color:#173250; font-size:10.5px; }
    select, input[type="search"], input[type="date"] {
      border-radius:10px;
      border-color:#c8d8ec;
      background:#fbfdff;
      font-size:12px;
      font-weight:650;
    }
    select:focus, input[type="search"]:focus, input[type="date"]:focus {
      outline:2px solid rgba(11,99,246,.16);
      border-color:#5d95ef;
    }
    .grid { gap:14px; }
    .panel, .home-card, .module-card, .focus-chart-card, .focus-summary-card, .cause-chart-box, .cause-table-box, .side-card {
      border-radius:15px;
      border:1px solid #dce7f5;
      background:#ffffff;
      box-shadow:0 18px 38px rgba(7,26,51,.075);
    }
    .panel { padding:16px; }
    .panel h2 {
      font-size:15px;
      font-weight:900;
      color:#071a33;
      letter-spacing:0;
      margin-bottom:12px;
      display:flex;
      align-items:center;
      gap:8px;
    }
    .panel h2::before {
      content:"";
      width:8px;
      height:18px;
      border-radius:999px;
      background:linear-gradient(180deg, #0b63f6, #11c5e8);
      display:inline-block;
    }
    .home-card { min-height:calc(100vh - 206px); }
    .home-eyebrow {
      color:#0b63f6;
      font-size:12px;
      text-transform:uppercase;
      letter-spacing:.12em;
    }
    .home-title { font-size:30px; color:#071a33; }
    .module-card {
      min-height:138px;
      padding:24px 22px 18px 86px;
    }
    .module-card:hover {
      border-color:#9ec2f5;
      box-shadow:0 24px 48px rgba(7,26,51,.12);
      transform:translateY(-3px);
    }
    .module-icon {
      left:24px;
      top:24px;
      width:48px;
      height:48px;
      border-radius:14px;
      background:#e8f1ff;
      color:#0b63f6;
      box-shadow:inset 0 0 0 1px rgba(11,99,246,.08);
    }
    .module-card h3 { font-size:17px; color:#071a33; }
    .module-card p { color:#5b6f8c; font-size:13px; }
    .module-link {
      border-radius:9px;
      background:#eef5ff;
      border-color:#cfe0f8;
      color:#0b63f6;
    }
    .table-wrap {
      border-radius:13px;
      border-color:#d8e4f2;
      box-shadow:inset 0 0 0 1px rgba(255,255,255,.6);
    }
    table { background:#fff; }
    th {
      background:linear-gradient(180deg, #082145, #061833);
      color:#ffffff;
      font-size:10.5px;
      text-transform:none;
      letter-spacing:.01em;
    }
    td {
      color:#11243f;
      border-bottom-color:#e8eef7;
      font-weight:600;
    }
    tbody tr:nth-child(even) td { background:#fbfdff; }
    tr:hover td { background:#eef6ff; }
    tbody tr.qty-alert td,
    tbody tr.qty-alert:nth-child(even) td,
    tbody tr.qty-alert:hover td {
      background:#fff0f2 !important;
      border-bottom-color:#ffc9d0 !important;
      color:#8f0f16 !important;
    }
    tbody tr.qty-alert td:first-child {
      border-left:4px solid #dc2626 !important;
    }
    tbody tr.type-mp td { background:#f7fbff; }
    tbody tr.type-gr td { background:#f3f8ff; }
    tbody tr.type-bts td { background:#fff8ed; }
    tbody tr.type-mp td:first-child { border-left:4px solid #16a34a; }
    tbody tr.type-gr td:first-child { border-left:4px solid #0a5df0; }
    tbody tr.type-bts td:first-child { border-left:4px solid #f59e0b; }
    .status-chip, .pill {
      border-radius:999px;
      padding:4px 8px;
      font-weight:900;
      letter-spacing:.01em;
    }
    .doc-ok, .doc-bad { border-width:2px; }
    .bar-label { fill:#142a46; font-weight:900; }
    .bar-value { fill:#071a33; font-weight:950; }
    .axis { stroke:#d8e4f2; }
    .focus-stat, .cause-stat {
      border-radius:13px;
      border-color:#dce7f5;
      box-shadow:0 12px 26px rgba(7,26,51,.06);
      background:linear-gradient(180deg, #fff, #f9fbff);
    }
    .focus-stat strong, .cause-stat strong { font-size:20px; }
    .focus-stat i, .cause-stat i, .provider-cause-card h2 .module-badge {
      border-radius:12px;
      background:#e8f1ff;
      color:#0b63f6;
    }
    .cause-stat:nth-child(2) i, .focus-stat.green i { background:#dcfce7; color:#16a34a; }
    .focus-stat.amber i { background:#fff3d6; color:#f59e0b; }
    .focus-stat.red i { background:#ffe4e8; color:#ef4444; }
    @media (max-width: 1280px) {
      .home-card, .side-card { min-height:auto; }
      .module-grid { grid-template-columns:repeat(2, minmax(150px, 1fr)); }
      .upload { position:static; width:100%; margin:10px 0 10px; }
      .top-title { padding-right:0; }
      .filters { grid-template-columns:repeat(4, minmax(130px, 1fr)); }
      .kpis { grid-template-columns:repeat(3, minmax(0, 1fr)); }
      .panel { grid-column:span 6; }
    }
    @media (max-width: 1000px) {
      main { padding:14px; margin-left:0; width:100%; }
      .nav-tabs { position:static; width:100%; flex-direction:row; overflow:auto; margin-bottom:10px; }
      .upload, .filters, .grid, .kpis { grid-template-columns:1fr; }
      .panel { grid-column:1 / -1; }
    }
    @media (min-width: 1700px) {
      .panel { grid-column:span 4; }
      .dashboard-chart { grid-column:span 4; }
      .dashboard-detail { grid-column:1 / -1; }
      .panel.wide { grid-column:1 / -1; }
      svg { height:285px; }
    }

    /* Premium azul oscuro - opcion 1 */
    :root {
      --premium-navy:#061a38;
      --premium-navy-2:#082852;
      --premium-blue:#0a5df0;
      --premium-blue-2:#0f3f86;
      --premium-cyan:#11bde7;
      --premium-ink:#071a33;
      --premium-muted:#60718b;
      --premium-line:#d8e4f4;
      --premium-card:#ffffff;
    }
    body {
      font-family:"Segoe UI", "Inter", Arial, sans-serif;
      background:
        radial-gradient(circle at 18% 0%, rgba(15,93,240,.12), transparent 26%),
        linear-gradient(180deg, #eef4fb 0%, #f8fbff 42%, #eef4fb 100%);
      color:var(--premium-ink);
    }
    main { max-width:none; padding:16px 18px 22px; }
    .top-title {
      min-height:82px;
      padding:17px min(720px, 50vw) 14px 22px;
      border-radius:18px;
      background:linear-gradient(135deg, #041326 0%, var(--premium-navy) 42%, #073b68 100%);
      box-shadow:0 20px 48px rgba(4,19,38,.22);
      border:1px solid rgba(255,255,255,.12);
    }
    .top-title h2 { color:#fff; font-size:29px; font-weight:900; letter-spacing:0; }
    .top-title .badge { background:#dbeafe; color:#0a5df0; font-size:12px; }
    .top-title p { color:#c7d7ec; font-size:13px; font-weight:700; }
    .upload { top:26px; right:28px; width:min(700px, 49vw); gap:10px; }
    .upload input {
      height:42px;
      border-radius:11px;
      border-color:#d7e4f5;
      font-size:12px;
      font-weight:700;
      background:#fff;
      box-shadow:0 12px 26px rgba(4,19,38,.10);
    }
    button {
      min-height:40px;
      border-radius:11px;
      background:linear-gradient(135deg, #0a63ff, #0646b5);
      font-size:12px;
      font-weight:900;
      box-shadow:0 14px 28px rgba(10,93,240,.24);
    }
    button.secondary {
      color:#092245;
      background:#edf4fc;
      border:1px solid #d7e4f5;
      box-shadow:none;
    }
    .meta { color:#526985; font-size:11px; font-weight:800; margin:9px 4px 12px; }
    .nav-tabs {
      margin:12px 0 14px;
      padding:9px;
      border-radius:16px;
      background:rgba(255,255,255,.95);
      border:1px solid var(--premium-line);
      box-shadow:0 16px 34px rgba(4,19,38,.08);
    }
    .nav-tabs button {
      min-height:42px;
      border-radius:12px;
      padding:11px 16px 11px 44px;
      font-size:13px;
      font-weight:900;
      background:#f8fbff;
      color:#18314f;
      border:1px solid #d7e4f5;
      box-shadow:none;
    }
    .nav-tabs button.active {
      background:linear-gradient(135deg, #0a63ff, #073f9d);
      color:#fff;
      border-color:#0a63ff;
      box-shadow:0 13px 24px rgba(10,93,240,.22);
    }
    .nav-tabs button::before {
      left:13px;
      width:19px;
      height:19px;
      border-radius:7px;
      background:#e8f1ff;
      border-color:#bad3fb;
    }
    .nav-tabs button::after { left:19px; width:7px; height:7px; border-color:#0a5df0; }
    .kpis {
      grid-template-columns:repeat(6, minmax(0, 1fr));
      gap:12px;
      margin:12px 0 14px;
    }
    .kpi {
      min-height:94px;
      border-radius:16px;
      border:1px solid var(--premium-line);
      background:linear-gradient(180deg, #fff 0%, #f8fbff 100%);
      box-shadow:0 18px 38px rgba(4,19,38,.075);
      padding:17px 12px 13px 72px;
      overflow:hidden;
    }
    .kpi::after {
      display:block;
      content:"";
      position:absolute;
      left:0;
      top:0;
      bottom:0;
      width:4px;
      background:linear-gradient(180deg, #0a63ff, #0f3f86);
    }
    .kpi.teal::after, .kpi.green::after { background:linear-gradient(180deg, #16a35a, #0b7a58); }
    .kpi.amber::after { background:linear-gradient(180deg, #f59e0b, #cc6d00); }
    .kpi.red::after { background:linear-gradient(180deg, #ef4444, #b91c1c); }
    .kpi.purple::after { background:linear-gradient(180deg, #7c3aed, #5520bd); }
    .kpi .kpi-icon {
      left:16px;
      top:18px;
      width:43px;
      height:43px;
      border-radius:13px;
      background:#e8f1ff;
      color:#0a5df0;
      box-shadow:inset 0 0 0 1px rgba(10,93,240,.08);
    }
    .kpi.teal .kpi-icon, .kpi.green .kpi-icon { background:#dcf7ea; color:#139a55; }
    .kpi.amber .kpi-icon { background:#fff0d8; color:#e28706; }
    .kpi.red .kpi-icon { background:#ffe7eb; color:#ef4444; }
    .kpi.purple .kpi-icon { background:#efe8ff; color:#7c3aed; }
    .kpi svg { width:22px; height:22px; stroke-width:2.45; }
    .kpi span {
      display:block;
      color:#435979;
      font-size:10px;
      line-height:1.15;
      font-weight:950;
      text-transform:uppercase;
      letter-spacing:0;
      max-width:100%;
    }
    .kpi strong {
      margin-top:8px;
      color:#061a38;
      font-size:clamp(22px, 1.25vw, 28px);
      line-height:1;
      font-weight:950;
      white-space:nowrap;
      overflow:visible;
    }
    .kpi small { margin-top:6px; color:#6f819b; font-size:8.5px; font-weight:800; }
    .filters {
      border-radius:16px;
      border:1px solid var(--premium-line);
      background:#ffffff;
      box-shadow:0 16px 34px rgba(4,19,38,.065);
      padding:14px;
      gap:9px;
    }
    label { color:#18314f; font-size:10.5px; font-weight:950; }
    select, input[type="search"], input[type="date"] {
      height:38px;
      border-radius:10px;
      border-color:#c8d7ea;
      background:#fbfdff;
      color:#071a33;
      font-size:12px;
      font-weight:700;
    }
    .panel, .home-card, .module-card, .focus-chart-card, .focus-summary-card, .cause-chart-box, .cause-table-box {
      border-radius:16px;
      border:1px solid var(--premium-line);
      background:#fff;
      box-shadow:0 18px 38px rgba(4,19,38,.075);
    }
    .panel { padding:16px; min-height:310px; }
    .dashboard-chart { min-height:316px; }
    .panel h2 {
      color:#071a33;
      font-size:16px;
      font-weight:950;
      margin-bottom:12px;
    }
    .panel h2::before {
      width:9px;
      height:19px;
      background:linear-gradient(180deg, #0a63ff, #12b9e6);
    }
    .home-card {
      min-height:calc(100vh - 200px);
      padding:26px 28px;
      background:linear-gradient(180deg, #fff 0%, #fbfdff 100%);
    }
    .home-eyebrow {
      color:#0a5df0;
      font-size:12px;
      letter-spacing:.13em;
      text-transform:uppercase;
      font-weight:950;
    }
    .home-title { font-size:30px; color:#061a38; font-weight:950; letter-spacing:0; }
    .home-copy { color:#64758f; font-size:13px; font-weight:650; }
    .module-grid { grid-template-columns:repeat(3, minmax(250px, 1fr)); gap:16px; }
    .module-card {
      min-height:130px;
      padding:24px 22px 18px 88px;
      background:linear-gradient(180deg, #fff 0%, #f9fbff 100%);
    }
    .module-card:hover {
      transform:translateY(-3px);
      border-color:#9dbff2;
      box-shadow:0 24px 48px rgba(4,19,38,.13);
    }
    .module-icon {
      left:24px;
      top:24px;
      width:48px;
      height:48px;
      border-radius:15px;
      background:#e8f1ff;
      color:#0a5df0;
      font-size:0;
      box-shadow:inset 0 0 0 1px rgba(10,93,240,.08);
    }
    .module-icon svg { width:25px; height:25px; stroke:currentColor; fill:none; stroke-width:2.5; stroke-linecap:round; stroke-linejoin:round; }
    .module-card:nth-child(2) .module-icon { background:#e1f7ec; color:#16a35a; }
    .module-card:nth-child(3) .module-icon { background:#def7fb; color:#0891b2; }
    .module-card:nth-child(4) .module-icon { background:#efe8ff; color:#7c3aed; }
    .module-card:nth-child(5) .module-icon { background:#fff0e4; color:#f97316; }
    .module-card h3 { color:#071a33; font-size:18px; font-weight:950; margin-bottom:8px; }
    .module-card p { color:#5d6f8b; font-size:13px; font-weight:650; line-height:1.42; }
    .module-link {
      margin-top:auto;
      border-radius:10px;
      padding:8px 12px;
      background:#eef5ff;
      border-color:#cfe0f8;
      color:#0a5df0;
      font-size:12px;
      font-weight:950;
    }
    svg { height:248px; }
    .dashboard-chart svg { height:252px; }
    .bar-label { font-size:15px; fill:#142a46; font-weight:900; }
    .bar-value { font-size:15px; fill:#061a38; font-weight:950; }
    .cause-axis-label { font-size:12px; fill:#203955; font-weight:850; }
    .cause-value { font-size:12px; fill:#061a38; font-weight:950; }
    .axis { stroke:#d6e2f1; stroke-width:1.2; }
    .table-wrap { border-radius:14px; border-color:#d8e4f4; }
    th {
      background:linear-gradient(180deg, #08264f, #041326);
      color:#fff;
      font-size:10.5px;
      font-weight:950;
      letter-spacing:0;
    }
    td {
      color:#102542;
      font-size:10.2px;
      font-weight:650;
      border-bottom-color:#e7eef7;
    }
    .dashboard-detail th, .dashboard-detail td,
    .focus-section th, .focus-section td,
    .provider-cause-card th, .provider-cause-card td,
    .daily-detail th, .daily-detail td {
      font-size:9.4px;
      line-height:1.2;
      padding:6px 7px;
    }
    .status-chip {
      font-size:8.5px;
      padding:4px 8px;
      font-weight:950;
      white-space:nowrap;
    }
    .status-chip.ATENCION,
    .status-chip.EN_ATENCION { background:#fff3d6; color:#b66b00; border:1px solid #ffd580; }
    .status-chip.PEND,
    .status-chip.PENDIENTE { background:#e9f1fb; color:#24415f; border:1px solid #c8d7ea; }
    #statusDetailTable tbody tr.status-recibido:not(.qty-alert) td {
      background:#f2fbf6;
    }
    #statusDetailTable tbody tr.status-recibido:not(.qty-alert) td:first-child {
      border-left:4px solid #16a34a;
    }
    #statusDetailTable tbody tr.status-atencion:not(.qty-alert) td {
      background:#fff9ec;
      color:#7c4a03;
    }
    #statusDetailTable tbody tr.status-atencion:not(.qty-alert) td:first-child {
      border-left:4px solid #f59e0b;
    }
    #statusDetailTable tbody tr.status-pendiente:not(.qty-alert) td {
      background:#f3f8ff;
      color:#18395f;
    }
    #statusDetailTable tbody tr.status-pendiente:not(.qty-alert) td:first-child {
      border-left:4px solid #1768f2;
    }
    .doc-ok, .doc-bad { width:20px; height:20px; border-width:2px; }
    .provider-cause-card { min-height:620px; grid-column:span 6; }
    .provider-cause-card svg { height:auto; }
    .cause-chart-scroll { height:260px; }
    .cause-chart-box, .cause-table-box { padding:12px; }
    .cause-box-title { font-size:12px; color:#071a33; }
    .cause-stats { gap:9px; }
    .cause-stat {
      min-height:72px;
      border-radius:14px;
      padding-left:52px;
      background:linear-gradient(180deg, #fff, #f9fbff);
    }
    .cause-stat i { width:32px; height:32px; border-radius:11px; }
    .cause-stat strong { font-size:22px; color:#061a38; }
    .focus-top { grid-template-columns:1fr 1.18fr; gap:14px; }
    .focus-chart-card svg { height:400px; }
    .focus-stats { gap:9px; }
    .focus-stat { min-height:84px; border-radius:14px; }
    .focus-stat strong { font-size:20px; }
    .conciliation-panel .mini-kpis { grid-template-columns:repeat(4, minmax(0, 1fr)); }
    .conciliation-panel .kpi { min-height:88px; }
    .conciliation-panel .kpi {
      padding-left:82px;
    }
    .conciliation-panel .kpi .kpi-icon {
      left:20px;
      top:17px;
    }
    .conciliation-panel .kpi span {
      margin-top:2px;
    }
    .conciliation-panel .kpi strong {
      margin-top:10px;
      font-size:clamp(25px, 1.45vw, 32px);
    }
    .conciliation-table-card {
      border:1px solid #d8e4f4;
      border-radius:14px;
      background:#fff;
      overflow:hidden;
      box-shadow:0 12px 26px rgba(4,19,38,.055);
    }
    .conciliation-table-card h3 {
      margin:0;
      padding:10px 12px;
      color:#fff;
      background:linear-gradient(180deg, #08264f, #041326);
      font-size:13px;
      font-weight:950;
      display:flex;
      justify-content:space-between;
      align-items:center;
    }
    .conciliation-table-card h3 span:last-child {
      font-size:11px;
      color:#c7d7ec;
    }
    .conciliation-table-card table { min-width:0; table-layout:fixed; }
    .conciliation-table-card th, .conciliation-table-card td {
      font-size:10px;
      padding:7px 8px;
      line-height:1.18;
      white-space:normal;
      overflow-wrap:anywhere;
    }
    .conciliation-table-card th:nth-child(1), .conciliation-table-card td:nth-child(1) { width:84px; }
    .conciliation-table-card th:nth-child(2), .conciliation-table-card td:nth-child(2) { width:58px; text-align:center; }
    .conciliation-table-card th:nth-child(3), .conciliation-table-card td:nth-child(3) { width:40%; }
    .conciliation-table-card th:nth-child(5), .conciliation-table-card td:nth-child(5) { width:82px; text-align:center; }
    tr.conc-mp td { background:#f7fbff; }
    tr.conc-gr td { background:#f3f8ff; }
    tr.conc-bts td { background:#fff8ed; }
    tr.conc-mp td:first-child { border-left:4px solid #16a34a; }
    tr.conc-gr td:first-child { border-left:4px solid #0a5df0; }
    tr.conc-bts td:first-child { border-left:4px solid #f59e0b; }
    .type-chip {
      display:inline-block;
      min-width:34px;
      padding:4px 8px;
      border-radius:999px;
      font-size:9px;
      font-weight:950;
      text-align:center;
      white-space:nowrap;
      line-height:1;
      word-break:normal;
      overflow-wrap:normal;
    }
    .type-chip.MP { background:#dcf7ea; color:#08723e; border:1px solid #9ce3bd; }
    .type-chip.GR { background:#e8f1ff; color:#0748b2; border:1px solid #b8d2ff; }
    .type-chip.BTS { background:#fff0d8; color:#a85a00; border:1px solid #ffd390; }
    .focus-section .table-wrap { border-radius:12px; }
    .focus-section th, .focus-section td {
      font-size:8.8px;
      line-height:1.12;
      padding:6px 7px;
      font-weight:700;
    }
    #weeklyFocusDetailTable { table-layout:fixed; min-width:0; }
    #weeklyFocusDetailTable th,
    #weeklyFocusDetailTable td { vertical-align:middle; }
    #weeklyFocusDetailTable th:nth-child(1), #weeklyFocusDetailTable td:nth-child(1) { width:68px; }
    #weeklyFocusDetailTable th:nth-child(2), #weeklyFocusDetailTable td:nth-child(2) { width:68px; }
    #weeklyFocusDetailTable th:nth-child(3), #weeklyFocusDetailTable td:nth-child(3),
    #weeklyFocusDetailTable th:nth-child(4), #weeklyFocusDetailTable td:nth-child(4) { width:78px; }
    #weeklyFocusDetailTable th:nth-child(5), #weeklyFocusDetailTable td:nth-child(5) { width:132px; }
    #weeklyFocusDetailTable th:nth-child(6), #weeklyFocusDetailTable td:nth-child(6) { width:68px; }
    #weeklyFocusDetailTable th:nth-child(7), #weeklyFocusDetailTable td:nth-child(7) { width:auto; }
    #weeklyFocusDetailTable th:nth-child(8), #weeklyFocusDetailTable td:nth-child(8) { width:92px; }
    #weeklyFocusDetailTable th:nth-child(9), #weeklyFocusDetailTable td:nth-child(9) { width:66px; }
    #weeklyFocusDetailTable th:nth-child(10), #weeklyFocusDetailTable td:nth-child(10) { width:92px; text-align:right; }
    #weeklyFocusDetailTable th:nth-child(11), #weeklyFocusDetailTable td:nth-child(11) { width:88px; text-align:right; }
    #weeklyFocusDetailTable th:nth-child(12), #weeklyFocusDetailTable td:nth-child(12) { width:78px; text-align:right; }
    #weeklyFocusDetailTable th:nth-child(13), #weeklyFocusDetailTable td:nth-child(13) { width:84px; text-align:right; }
    #weeklyFocusDetailTable th:nth-child(10),
    #weeklyFocusDetailTable th:nth-child(11),
    #weeklyFocusDetailTable th:nth-child(12),
    #weeklyFocusDetailTable th:nth-child(13) {
      white-space:normal;
      word-break:normal;
      overflow-wrap:normal;
      line-height:1.05;
    }
    .focus-chart-card .bar-label { font-size:10px; fill:#082145; font-weight:850; }
    .focus-chart-card .bar-value { font-size:9.5px; fill:#061a38; font-weight:950; }
    .focus-chart-card .bar-value.inbar { fill:#ffffff; font-size:8.6px; font-weight:950; }
    .focus-chart-card .vertical-value { dominant-baseline:middle; writing-mode:initial; }
    .focus-chart-card .x-horizontal { font-size:9px; font-weight:900; dominant-baseline:middle; }
    .focus-chart-card .chart-title-label { font-size:10px; font-weight:950; }
    .focus-chart-card .chart-legend .bar-label { font-size:10px; }
    .focus-chart-card .pct-label { font-size:9px; fill:#061a38; font-weight:950; }
    #dashboard .dashboard-detail {
      grid-column:1 / -1 !important;
      width:100%;
    }
    #dashboard .dashboard-chart {
      min-height:210px !important;
      padding:10px 12px !important;
    }
    #dashboard .dashboard-chart svg {
      height:156px !important;
    }
    #dashboard .dashboard-chart h2 {
      margin-bottom:2px;
      font-size:14px;
    }
    #statusDetailTable {
      table-layout:fixed;
      width:100%;
      min-width:0;
    }
    #statusDetailTable th,
    #statusDetailTable td {
      font-size:8.2px !important;
      line-height:1.08 !important;
      padding:4px 5px !important;
      vertical-align:middle;
      text-align:center;
    }
    #statusDetailTable th:nth-child(1), #statusDetailTable td:nth-child(1) { width:5%; }
    #statusDetailTable th:nth-child(2), #statusDetailTable td:nth-child(2) { width:4%; }
    #statusDetailTable th:nth-child(3), #statusDetailTable td:nth-child(3) { width:5%; }
    #statusDetailTable th:nth-child(4), #statusDetailTable td:nth-child(4),
    #statusDetailTable th:nth-child(5), #statusDetailTable td:nth-child(5) { width:4%; }
    #statusDetailTable th:nth-child(6), #statusDetailTable td:nth-child(6) { width:12%; }
    #statusDetailTable th:nth-child(7), #statusDetailTable td:nth-child(7) { width:7%; }
    #statusDetailTable th:nth-child(8), #statusDetailTable td:nth-child(8) { width:5%; }
    #statusDetailTable th:nth-child(9), #statusDetailTable td:nth-child(9) { width:22%; }
    #statusDetailTable th:nth-child(10), #statusDetailTable td:nth-child(10),
    #statusDetailTable th:nth-child(11), #statusDetailTable td:nth-child(11) { width:6%; }
    #statusDetailTable th:nth-child(12), #statusDetailTable td:nth-child(12) { width:4%; }
    #statusDetailTable th:nth-child(13), #statusDetailTable td:nth-child(13) { width:5%; }
    #statusDetailTable th:nth-child(14), #statusDetailTable td:nth-child(14) { width:4%; }
    #statusDetailTable th:nth-child(15), #statusDetailTable td:nth-child(15) { width:7%; }
    #statusDetailTable .type-chip {
      display:inline-flex;
      align-items:center;
      justify-content:center;
      min-width:42px;
      height:22px;
    }
    @media (max-width: 1420px) {
      .kpis { grid-template-columns:repeat(3, minmax(0, 1fr)); }
      .upload { position:static; width:100%; margin:10px 0 8px; }
      .top-title { padding-right:22px; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Plan de Recibo Control Tower</h1>
      <p>Gestion operativa de citas, vehiculos, cumplimiento de proveedores y materiales criticos.</p>
    </div>
    <div id="status">Esperando archivo</div>
  </header>
  <main>
    <div class="top-title">
      <h2>Plan de Recibo <span class="badge">2026.2</span></h2>
      <p id="topMeta">Importa el archivo para activar los modulos de control.</p>
    </div>
    <section class="upload">
      <input id="file" type="file" accept=".xlsx,.xlsm,.xls" />
      <button id="analyze">Importar y generar dashboard</button>
      <button class="secondary" id="demo">Usar archivo actual</button>
    </section>
    <div class="meta" id="meta">Sube el archivo PLAN DE RECIBO o usa el archivo de OneDrive detectado en esta maquina.</div>

    <nav class="nav-tabs">
      <button class="active" data-page-target="home">Inicio</button>
      <button data-page-target="dashboard">Indicadores</button>
      <button data-page-target="proveedores">Proveedores</button>
      <button data-page-target="recibo">Recibo y materiales</button>
      <button data-page-target="conciliacion">Conciliacion</button>
      <button data-page-target="base">Base</button>
    </nav>
    <div class="module-toolbar page-hidden" id="moduleToolbar">
      <button class="export-module" id="exportModule" type="button">Exportar modulo</button>
    </div>

    <section class="home" id="home" data-page="home" hidden>
      <div class="home-layout">
        <div class="home-card">
          <div class="home-eyebrow">Acceso rapido</div>
          <h2 class="home-title">Modulos disponibles</h2>
          <p class="home-copy">Entra directo al control que necesitas revisar.</p>
          <div class="module-grid">
            <article class="module-card" data-open-page="dashboard">
              <div class="module-icon" data-module-icon="bars">KPI</div>
            <h3>Indicadores</h3>
            <p>Citas, estados, alertas y volumen programado.</p>
              <div class="module-metric" id="moduleIndicadores">--</div>
              <span class="module-link">Abrir modulo -></span>
            </article>
            <article class="module-card" data-open-page="proveedores">
              <div class="module-icon" data-module-icon="user">PV</div>
            <h3>Proveedores y causales</h3>
            <p>Causales, novedades documentales y cumplimiento.</p>
              <div class="module-metric" id="moduleProveedores">--</div>
              <span class="module-link">Abrir modulo -></span>
            </article>
            <article class="module-card" data-open-page="recibo">
              <div class="module-icon" data-module-icon="truck">RC</div>
              <h3>Recibo</h3>
              <p>Vehiculos, agenda y materiales criticos por semana.</p>
              <div class="module-metric" id="moduleRecibo">--</div>
              <span class="module-link">Abrir modulo -></span>
            </article>
            <article class="module-card" data-open-page="conciliacion">
              <div class="module-icon" data-module-icon="doc">DOC</div>
              <h3>Conciliacion</h3>
              <p>Documentos por fecha, tipo MP/GR/BTS y proveedor.</p>
              <div class="module-metric" id="moduleConciliacion">--</div>
              <span class="module-link">Abrir modulo -></span>
            </article>
            <article class="module-card" data-open-page="base">
              <div class="module-icon" data-module-icon="folder">BD</div>
              <h3>Base</h3>
              <p>Tabla filtrada para auditoria, exportacion y revision.</p>
              <div class="module-metric" id="moduleBase">--</div>
              <span class="module-link">Abrir modulo -></span>
            </article>
          </div>
        </div>
      </div>
    </section>

    <section class="filters page-hidden" id="filters" hidden>
      <div><label>Tipo material</label><select data-filter="TIPO MATERIAL"></select></div>
      <div><label>Proveedor</label><select data-filter="PROVEEDOR"></select></div>
      <div><label>Transportadora</label><select data-filter="TRANSPORTADORA"></select></div>
      <div><label>Año</label><select data-filter="ANO CONTROL"></select></div>
      <div><label>Mes</label><select data-filter="MES CONTROL"></select></div>
      <div><label>Semana</label><select data-filter="SEMANA CONTROL"></select></div>
      <div><label>Estado</label><select data-filter="ESTADO CONTROL"></select></div>
      <div><label>Estado vehiculo</label><select data-filter="ESTADO VEHICULO"></select></div>
      <div><label>Documentación</label><select data-filter="DOCUMENTACION"></select></div>
      <div><label>Cantidad</label><select data-filter="CUMPLE CANTIDAD"></select></div>
      <div><label>Cita Traffic</label><select data-filter="CITA TRAFFIC"></select></div>
      <div><label>Buscar</label><input id="search" type="search" placeholder="SKU, proveedor, placa..." /></div>
      <div><label>Programada desde</label><input id="progFrom" type="date" /></div>
      <div><label>Programada hasta</label><input id="progTo" type="date" /></div>
      <div><label>Estimada desde</label><input id="estFrom" type="date" /></div>
      <div><label>Estimada hasta</label><input id="estTo" type="date" /></div>
      <div><label>Reprogramada desde</label><input id="repFrom" type="date" /></div>
      <div><label>Reprogramada hasta</label><input id="repTo" type="date" /></div>
      <div class="filter-actions"><button class="clear-filters" id="clearFilters" type="button">Limpiar filtros</button></div>
    </section>

    <section class="kpis" id="kpis"></section>

    <section class="grid" id="dashboard" hidden>
      <div class="panel dashboard-chart" data-page="dashboard"><h2 id="citaChartTitle">Citas (cantidad)</h2><svg id="monthChart"></svg></div>
      <div class="panel dashboard-chart" data-page="dashboard"><h2>Estado de recibos</h2><svg id="statusChart"></svg></div>
      <div class="panel dashboard-chart" data-page="dashboard"><h2>Alertas principales</h2><svg id="alertChart"></svg></div>
      <div class="panel dashboard-detail" data-page="dashboard">
        <div class="actions"><h2 id="statusDetailTitle">Detalle de citas</h2><span class="detail-count" id="statusDetailCount"></span></div>
        <div class="table-wrap"><table id="statusDetailTable"></table></div>
      </div>
      <div class="panel provider-cause-card" id="proveedores" data-page="proveedores">
        <div class="actions">
          <h2><span class="module-badge" data-icon="doc"></span>Novedad documental</h2>
          <div class="mini-filters">
            <label>Desde<input id="docCauseFrom" type="date" /></label>
            <label>Hasta<input id="docCauseTo" type="date" /></label>
            <label>Proveedor<select id="docCauseProvider"><option value="">Todos</option></select></label>
            <label>Sector<select id="docCauseSector"><option value="">Todos</option></select></label>
            <button class="cause-filter-btn" type="button">Filtrar</button>
          </div>
        </div>
        <div class="cause-stats" id="docCauseStats"></div>
        <div class="cause-chart-box">
          <div class="cause-box-title"><span>Novedad documental por estado</span><span>Vista grafica</span></div>
          <div class="cause-chart-scroll"><svg id="docCauseChart"></svg></div>
        </div>
        <div class="cause-table-box">
          <div class="cause-box-title"><span>Detalle de novedades documentales</span><span>Detalle</span></div>
          <div class="table-wrap compact"><table id="docCauseTable"></table></div>
          <div class="cause-footer" id="docCauseFooter"></div>
        </div>
      </div>
      <div class="panel provider-cause-card" data-page="proveedores">
        <div class="actions">
          <h2><span class="module-badge" data-icon="alert"></span>Causal fisica</h2>
          <div class="date-mode-strip">
            <strong>Filtrar por fecha</strong>
            <select id="physicalCauseDateMode"><option value="recibo">Recibo</option><option value="arribo">Arribo</option></select>
            <span>Usa Arribo cuando la novedad no tenga fecha de recibo.</span>
          </div>
          <div class="mini-filters">
            <label>Desde<input id="physicalCauseFrom" type="date" /></label>
            <label>Hasta<input id="physicalCauseTo" type="date" /></label>
            <label>Proveedor<select id="physicalCauseProvider"><option value="">Todos</option></select></label>
            <label>Sector<select id="physicalCauseSector"><option value="">Todos</option></select></label>
            <button class="cause-filter-btn" type="button">Filtrar</button>
          </div>
        </div>
        <div class="cause-stats" id="physicalCauseStats"></div>
        <div class="cause-chart-box">
          <div class="cause-box-title"><span>Causal fisica por tipo</span><span>Vista grafica</span></div>
          <div class="cause-chart-scroll"><svg id="physicalCauseChart"></svg></div>
        </div>
        <div class="cause-table-box">
          <div class="cause-box-title"><span>Detalle de causales fisicas</span><span>Detalle</span></div>
          <div class="table-wrap compact"><table id="physicalCauseTable"></table></div>
          <div class="cause-footer" id="physicalCauseFooter"></div>
        </div>
      </div>
      <div class="panel" data-page="recibo-extra"><h2>Estado vehiculo</h2><svg id="vehicleChart"></svg></div>
      <div class="panel wide" data-page="proveedores-extra"><h2>Cumplimiento cantidad por proveedor</h2><svg id="providerQtyChart"></svg></div>
      <div class="panel wide focus-section" id="recibo" data-page="recibo">
        <div class="actions">
          <h2>Materiales por semana</h2>
          <div class="mini-filters">
            <label>Año<select id="focusYear"><option value="">Todos</option></select></label>
            <label>Mes<select id="focusMonth"><option value="">Todos</option></select></label>
            <label>Material<select id="focusMaterial"><option value="">Todos</option></select></label>
            <label>Proveedor<select id="focusProvider"><option value="">Todos</option></select></label>
            <label>Semana<select id="focusWeek"><option value="">Todas</option></select></label>
            <label>SKU<select id="focusSku"><option value="">Todos</option></select></label>
            <label>Consulta<input id="focusSearch" type="search" placeholder="SKU, material, proveedor..." /></label>
          </div>
        </div>
        <div class="focus-top">
          <div class="focus-chart-card">
            <svg id="weeklyFocusChart"></svg>
          </div>
          <div class="focus-summary-card">
            <h3 class="subhead">Resumen semanal</h3>
            <div class="table-wrap compact"><table id="weeklyFocusTable"></table></div>
            <section class="focus-stats" id="focusStats"></section>
          </div>
        </div>
        <h3 class="subhead">Detalle encontrado por texto en material</h3>
        <div class="table-wrap compact"><table id="weeklyFocusDetailTable"></table></div>
      </div>
      <div class="panel wide" data-page="recibo">
        <div class="actions">
          <h2>Agenda próximos 14 días</h2>
          <span id="agendaCount"></span>
        </div>
        <div class="table-wrap"><table id="agenda"></table></div>
      </div>
      <div class="panel wide conciliation-panel" id="conciliacion" data-page="conciliacion">
        <div class="actions">
          <h2>Conciliacion documental</h2>
          <div class="mini-filters">
            <label>Fecha desde<input id="concFrom" type="date" /></label>
            <label>Fecha hasta<input id="concTo" type="date" /></label>
            <label>Tipo<select id="concTipo"><option value="">Todos</option><option>MP</option><option>GR</option><option>BTS</option></select></label>
            <label>Proveedor<select id="concProvider"><option value="">Todos</option></select></label>
          </div>
        </div>
        <section class="kpis mini-kpis" id="conciliationKpis"></section>
        <div class="conciliation-table-card">
          <h3><span>Detalle por tipo</span><span id="conciliationTableMeta"></span></h3>
          <div class="table-wrap compact"><table id="conciliationTable"></table></div>
        </div>
      </div>
      <div class="panel wide" id="base" data-page="base">
        <div class="actions">
          <h2>Base filtrada</h2>
          <div><span id="rowCount"></span> <button class="secondary" id="download">Descargar CSV filtrado</button></div>
        </div>
        <div class="table-wrap"><table id="table"></table></div>
      </div>
    </section>
  </main>

<script>
let dataset = null;
let filteredRows = [];
let currentFocusRows = [];
let currentPage = "home";

const fmt = new Intl.NumberFormat("es-CO");
const money = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 2 });

function setBusy(text, busy=true) {
  document.getElementById("status").textContent = text;
  document.getElementById("analyze").disabled = busy;
  document.getElementById("demo").disabled = busy;
}

function showError(message) {
  setBusy("Error", false);
  document.getElementById("meta").textContent = message;
}

async function upload(useDemo=false) {
  try {
    setBusy("Procesando archivo...");
    let response;
    if (useDemo) {
      response = await fetch("/api/demo", { method: "POST" });
    } else {
      const file = document.getElementById("file").files[0];
      if (!file) return showError("Selecciona un archivo primero.");
      const body = new FormData();
      body.append("file", file);
      response = await fetch("/api/analyze", { method: "POST", body });
    }
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "No se pudo analizar el archivo.");
    dataset = data;
    initializeDashboard();
    setBusy("Dashboard listo", false);
  } catch (error) {
    showError(error.message);
  }
}

function initializeDashboard() {
  document.getElementById("home").hidden = false;
  document.getElementById("filters").hidden = false;
  document.getElementById("dashboard").hidden = false;
  document.getElementById("meta").textContent = `${dataset.fileName} | generado ${dataset.generatedAt} | ${fmt.format(dataset.kpis.total)} registros`;
  document.getElementById("topMeta").textContent = `Generado: ${dataset.generatedAt} | ${fmt.format(dataset.kpis.total)} registros`;
  renderKpis(dataset.kpis);
  renderModuleMetrics(dataset);
  buildFilters();
  applyFilters();
  setActivePage("home");
}

function renderModuleMetrics(data) {
  document.getElementById("moduleIndicadores").textContent = fmt.format(data.kpis.total);
  document.getElementById("moduleProveedores").textContent = fmt.format(data.filters.PROVEEDOR?.length || 0);
  document.getElementById("moduleRecibo").textContent = `${fmt.format(data.kpis.inAttention)} / ${fmt.format(data.kpis.receivedVehicle)}`;
  document.getElementById("moduleConciliacion").textContent = fmt.format(data.conciliation?.summary?.["TOTAL DOC"] || 0);
  document.getElementById("moduleBase").textContent = fmt.format(data.rows.length);
}

function renderKpis(k) {
  const items = [
    ["Total citas", k.total, "", "calendar", "vs dia anterior"],
    ["Recibidos", k.receivedVehicle, "teal", "doc", "vs dia anterior"],
    ["Atención", k.inAttention, "amber", "clock", "vs dia anterior"],
    ["Pendientes", k.pendingVehicle || 0, "", "folder", "vs dia anterior"],
    ["Cumplimiento doc.", `${k.docRate}%`, "purple", "doc", "vs dia anterior"],
    ["Citas hoy", k.todayTrafficTotal, "green", "user", "vs dia anterior"],
  ];
  document.getElementById("kpis").innerHTML = items.map(([label, value, cls, icon, note]) =>
    `<div class="kpi ${cls}"><i class="kpi-icon">${iconSvg(icon)}</i><span>${label}</span><strong>${typeof value === "number" ? money.format(value) : value}</strong><small>${note}</small></div>`
  ).join("");
}

function iconSvg(name) {
  const icons = {
    calendar: '<svg viewBox="0 0 24 24"><path d="M7 3v4M17 3v4M4 9h16M5 5h14a1 1 0 0 1 1 1v14H4V6a1 1 0 0 1 1-1z"/></svg>',
    bars: '<svg viewBox="0 0 24 24"><path d="M5 19V11M12 19V5M19 19v-8"/></svg>',
    truck: '<svg viewBox="0 0 24 24"><path d="M3 7h11v10H3zM14 11h4l3 3v3h-7z"/><path d="M7 19a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM17 19a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/></svg>',
    clock: '<svg viewBox="0 0 24 24"><path d="M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z"/><path d="M12 7v5l3 2"/></svg>',
    folder: '<svg viewBox="0 0 24 24"><path d="M3 7h7l2 2h9v10H3z"/></svg>',
    alert: '<svg viewBox="0 0 24 24"><path d="M12 3 22 20H2z"/><path d="M12 9v5M12 17h.01"/></svg>',
    percent: '<svg viewBox="0 0 24 24"><path d="M19 5 5 19"/><path d="M7 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM17 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/></svg>',
    bell: '<svg viewBox="0 0 24 24"><path d="M18 9a6 6 0 1 0-12 0c0 7-3 6-3 9h18c0-3-3-2-3-9z"/><path d="M10 21h4"/></svg>',
    doc: '<svg viewBox="0 0 24 24"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v5h5M10 13h6M10 17h6"/></svg>',
    x: '<svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/><path d="M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z"/></svg>',
    user: '<svg viewBox="0 0 24 24"><path d="M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8z"/><path d="M4 21a8 8 0 0 1 16 0"/></svg>',
  };
  return icons[name] || icons.calendar;
}

function buildFilters() {
  document.querySelectorAll("[data-filter]").forEach(select => {
    const col = select.dataset.filter;
    const values = dataset.filters[col] || [];
    select.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
    select.onchange = () => {
      if (col === "ANO CONTROL" || col === "MES CONTROL") syncWeekFilter();
      applyFilters();
    };
  });
  syncWeekFilter();
  document.getElementById("search").oninput = applyFilters;
  ["progFrom", "progTo", "estFrom", "estTo", "repFrom", "repTo"].forEach(id => {
    document.getElementById(id).onchange = applyFilters;
  });
  document.getElementById("clearFilters").onclick = clearDashboardFilters;
  ["focusYear", "focusMonth", "focusMaterial", "focusProvider", "focusWeek", "focusSku", "focusSearch"].forEach(id => {
    document.getElementById(id).onchange = () => renderWeeklyFocus(dataset.rows);
  });
  document.getElementById("focusSearch").oninput = () => renderWeeklyFocus(dataset.rows);
  ["docCauseFrom", "docCauseTo", "docCauseProvider", "docCauseSector", "physicalCauseFrom", "physicalCauseTo", "physicalCauseDateMode", "physicalCauseProvider", "physicalCauseSector"].forEach(id => {
    document.getElementById(id).onchange = () => renderProviderModule(dataset.rows);
  });
  ["concFrom", "concTo", "concTipo", "concProvider"].forEach(id => {
    document.getElementById(id).onchange = renderConciliation;
  });
  document.querySelectorAll(".nav-tabs button").forEach(button => {
    button.onclick = () => {
      setActivePage(button.dataset.pageTarget);
    };
  });
  document.querySelectorAll("[data-open-page]").forEach(card => {
    card.onclick = () => setActivePage(card.dataset.openPage);
  });
  document.querySelectorAll("[data-icon]").forEach(item => {
    item.innerHTML = iconSvg(item.dataset.icon);
  });
  document.querySelectorAll("[data-module-icon]").forEach(item => {
    item.innerHTML = iconSvg(item.dataset.moduleIcon);
  });
}

function syncWeekFilter() {
  const weekSelect = document.querySelector('[data-filter="SEMANA CONTROL"]');
  if (!weekSelect || !dataset?.rows) return;
  const selectedYear = selectedFilterValue("ANO CONTROL");
  const selectedMonth = selectedFilterValue("MES CONTROL");
  const currentWeek = weekSelect.value;
  const weeks = new Set();
  dataset.rows.forEach(row => {
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    const week = String(row["SEMANA CONTROL"] || "").trim();
    if (week) weeks.add(week);
  });
  const values = [...weeks].sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
  weekSelect.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  weekSelect.value = values.includes(currentWeek) ? currentWeek : "";
}

function clearDashboardFilters() {
  document.querySelectorAll("[data-filter]").forEach(select => {
    select.value = "";
  });
  ["search", "progFrom", "progTo", "estFrom", "estTo", "repFrom", "repTo"].forEach(id => {
    const field = document.getElementById(id);
    if (field) field.value = "";
  });
  syncWeekFilter();
  applyFilters();
}

function setActivePage(page) {
  currentPage = page;
  document.querySelectorAll(".nav-tabs button").forEach(button => {
    button.classList.toggle("active", button.dataset.pageTarget === page);
  });
  document.querySelector(".nav-tabs").classList.toggle("home-mode", page === "home");
  document.getElementById("moduleToolbar").classList.toggle("page-hidden", page === "home");
  document.getElementById("kpis").classList.toggle("page-hidden", page !== "dashboard");
  document.getElementById("dashboard").classList.toggle("page-hidden", page === "home");
  document.querySelectorAll("[data-page]").forEach(panel => {
    panel.classList.toggle("page-hidden", panel.dataset.page !== page);
  });
  document.getElementById("filters").classList.toggle("page-hidden", page === "home" || page === "proveedores" || page === "recibo" || page === "conciliacion");
  if (page === "proveedores") renderProviderModule(dataset.rows);
  if (page === "recibo") renderWeeklyFocus(dataset.rows);
  if (page === "conciliacion") renderConciliation();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function applyFilters() {
  const criteria = {};
  document.querySelectorAll("[data-filter]").forEach(select => criteria[select.dataset.filter] = select.value);
  const q = document.getElementById("search").value.trim().toLowerCase();
  const progFrom = document.getElementById("progFrom").value;
  const progTo = document.getElementById("progTo").value;
  const estFrom = document.getElementById("estFrom").value;
  const estTo = document.getElementById("estTo").value;
  const repFrom = document.getElementById("repFrom").value;
  const repTo = document.getElementById("repTo").value;
  filteredRows = dataset.rows.filter(row => {
    for (const [col, value] of Object.entries(criteria)) {
      if (value && String(row[col] || "") !== value) return false;
    }
    if (!dateInRange(row["FECHA PROGRAMADA"], progFrom, progTo)) return false;
    if (!dateInRange(row["FECHA ESTIMADA ENTREGA"], estFrom, estTo)) return false;
    if (!dateInRange(row["FECHA REPROGRAMADA"], repFrom, repTo)) return false;
    if (!q) return true;
    return Object.values(row).some(value => String(value || "").toLowerCase().includes(q));
  });
  renderFilteredKpis();
  renderChartsFromRows(filteredRows);
  renderTable("table", filteredRows.slice(0, 500), dataset.columns);
  document.getElementById("rowCount").textContent = `${fmt.format(filteredRows.length)} filas filtradas`;
  renderTable("agenda", dataset.agenda, Object.keys(dataset.agenda[0] || {}));
  document.getElementById("agendaCount").textContent = `${fmt.format(dataset.agenda.length)} pendientes`;
  renderConciliation();
}

function dateInRange(value, from, to) {
  const current = String(value || "").slice(0, 10);
  if (!current && (from || to)) return false;
  if (from && current < from) return false;
  if (to && current > to) return false;
  return true;
}

function renderFilteredKpis() {
  const today = new Date().toISOString().slice(0, 10);
  const dueRows = filteredRows.filter(row => String(row["FECHA CONTROL"] || "").slice(0, 10) <= today);
  const docComplete = dueRows.filter(row => String(row["N DOCUMENTO"] || "").trim() && (String(row.PLACA || "").trim() || row["ESTADO VEHICULO"] === "PENDIENTE")).length;
  const status = citaStatusSummary(filteredRows);
  const k = {
    total: uniqueCitas(filteredRows).size,
    receivedVehicle: status["RECIBIDO"] || 0,
    inAttention: status["EN ATENCION"] || 0,
    pendingVehicle: status["PENDIENTE"] || 0,
    docRate: Math.round((docComplete / Math.max(dueRows.length, 1)) * 1000) / 10,
    todayTrafficTotal: uniqueCitas(filteredRows.filter(row => String(row["FECHA CONTROL"] || "").slice(0, 10) === today)).size,
  };
  renderKpis(k);
  document.getElementById("status").textContent = `${fmt.format(filteredRows.length)} registros filtrados`;
}

function citaBase(row) {
  const kpi = String(row["CITA KPI"] || "").trim();
  if (kpi) return kpi;
  const direct = String(row["CITA BASE"] || "").trim();
  if (direct) return direct;
  const text = String(row["CODIGO CITA"] || "").trim();
  const match = text.match(/^(\d{5})(?:-\d+)?$/);
  if (match) return match[1];
  if (text.toUpperCase() === "BTS") {
    return `BTS|${cleanIdentifier(row["N DOCUMENTO"] || row.SKU || row.MATERIAL || row.PROVEEDOR || "")}`;
  }
  return "";
}

function citaStatusSummary(rows) {
  const priority = {"PENDIENTE": 3, "EN ATENCION": 2, "RECIBIDO": 1};
  const byCita = new Map();
  rows.forEach(row => {
    const cita = citaBase(row);
    if (!cita) return;
    const status = String(row["ESTADO VEHICULO"] || "").trim().toUpperCase();
    const current = byCita.get(cita);
    if (!current || (priority[status] || 0) > (priority[current] || 0)) byCita.set(cita, status);
  });
  const counts = {};
  byCita.forEach(status => {
    const key = status || "SIN ESTADO";
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function statusItemsFromRows(rows) {
  return Object.entries(citaStatusSummary(rows))
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => statusSortValue(a.name) - statusSortValue(b.name));
}

function uniqueCitas(rows) {
  const set = new Set();
  rows.forEach(row => {
    const base = citaBase(row);
    if (base) set.add(base);
  });
  return set;
}

function compactNumber(value) {
  const number = Number(value || 0);
  if (Math.abs(number) >= 1000000) return `${new Intl.NumberFormat("es-CO", { maximumFractionDigits: 2 }).format(number / 1000000)} M`;
  return money.format(number);
}

function countBy(rows, col, limit=10) {
  const map = new Map();
  const seen = new Set();
  rows.forEach(row => {
    const key = String(row[col] || `Sin ${col.toLowerCase()}`);
    const cita = citaBase(row);
    const uniqueKey = cita ? `${key}|${cita}` : "";
    if (uniqueKey) {
      if (seen.has(uniqueKey)) return;
      seen.add(uniqueKey);
    }
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map.entries()].sort((a,b) => b[1]-a[1]).slice(0, limit).map(([name,value]) => ({name, value}));
}

function sumBy(rows, col, valueCol, limit=10) {
  const map = new Map();
  rows.forEach(row => {
    const key = String(row[col] || `Sin ${col.toLowerCase()}`);
    map.set(key, (map.get(key) || 0) + Number(row[valueCol] || 0));
  });
  return [...map.entries()].sort((a,b) => b[1]-a[1]).slice(0, limit).map(([name,value]) => ({name, value}));
}

function alertItems(rows) {
  const map = new Map();
  rows.forEach(row => {
    if (row.ALERTA === "OK") return;
    String(row.ALERTA || "").split(" | ").forEach(part => {
      if (part) map.set(part, (map.get(part) || 0) + 1);
    });
  });
  return [...map.entries()].sort((a,b) => b[1]-a[1]).slice(0, 8).map(([name,value]) => ({name, value}));
}

function renderProviderModule(rows) {
  syncProviderCauseFilters(rows, "docCause");
  syncProviderCauseFilters(rows, "physicalCause");
  renderCauseBlock(rows, {
    field: "NOVEDAD DOCUMENTAL",
    chartId: "docCauseChart",
    tableId: "docCauseTable",
    statsId: "docCauseStats",
    footerId: "docCauseFooter",
    providerId: "docCauseProvider",
    sectorId: "docCauseSector",
    fromId: "docCauseFrom",
    toId: "docCauseTo",
    mode: "single",
    emptyName: "Sin novedad documental",
    label: "Novedad documental",
  });
  renderCauseBlock(rows, {
    field: "CAUSAL NOVEDAD FISICA",
    chartId: "physicalCauseChart",
    tableId: "physicalCauseTable",
    statsId: "physicalCauseStats",
    footerId: "physicalCauseFooter",
    providerId: "physicalCauseProvider",
    sectorId: "physicalCauseSector",
    dateModeId: "physicalCauseDateMode",
    fromId: "physicalCauseFrom",
    toId: "physicalCauseTo",
    mode: "single",
    emptyName: "Sin causal fisica",
    label: "Causal fisica",
  });
}

function syncProviderCauseFilters(rows, prefix) {
  const providerSelect = document.getElementById(`${prefix}Provider`);
  const sectorSelect = document.getElementById(`${prefix}Sector`);
  const currentProvider = providerSelect.value;
  const currentSector = sectorSelect.value;
  const providers = [...new Set(rows.map(row => row.PROVEEDOR).filter(Boolean))].sort();
  const sectors = [...new Set(rows.map(row => row["TIPO MATERIAL"]).filter(Boolean))].sort();
  providerSelect.innerHTML = `<option value="">Todos</option>` + providers.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  sectorSelect.innerHTML = `<option value="">Todos</option>` + sectors.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (providers.includes(currentProvider)) providerSelect.value = currentProvider;
  if (sectors.includes(currentSector)) sectorSelect.value = currentSector;
}

function renderCauseBlock(rows, config) {
  const provider = document.getElementById(config.providerId).value;
  const sectorFilter = document.getElementById(config.sectorId).value;
  const from = document.getElementById(config.fromId).value;
  const to = document.getElementById(config.toId).value;
  const dateMode = config.dateModeId ? document.getElementById(config.dateModeId).value : "recibo";
  const counts = new Map();
  const details = [];
  let withoutCause = 0;
  const docMap = new Map();
  const docCausePairs = new Set();
  const providers = new Set();
  const sectors = new Set();
  const materials = new Set();
  const docs = new Set();
  rows.forEach(row => {
    const sector = row["TIPO MATERIAL"] || "Sin sector";
    if (provider && row.PROVEEDOR !== provider) return;
    if (sectorFilter && sector !== sectorFilter) return;
    const causeDate = dateMode === "arribo" ? row["FECHA ARRIBO"] : row["FECHA RECIBO"];
    if (!dateInRange(causeDate, from, to)) return;
    const doc = String(row["N DOCUMENTO"] || "").trim();
    const docKey = doc || `${row["CODIGO CITA"] || ""}|${row.SKU || ""}|${row.PROVEEDOR || ""}|${row["FECHA RECIBO"] || ""}`;
    if (config.field === "NOVEDAD DOCUMENTAL" && docKey) {
      if (!docMap.has(docKey)) docMap.set(docKey, { hasCause: false });
    }
    const causes = extractCauses(row, config);
    if (!causes.length) {
      withoutCause += 1;
      return;
    }
    if (config.field === "NOVEDAD DOCUMENTAL" && docKey) {
      docMap.get(docKey).hasCause = true;
    }
    if (row.PROVEEDOR) providers.add(row.PROVEEDOR);
    if (sector) sectors.add(sector);
    if (row.MATERIAL) materials.add(row.MATERIAL);
    if (doc) docs.add(doc);
    causes.forEach(cause => {
      const countKey = config.field === "NOVEDAD DOCUMENTAL" ? `${docKey}|${cause}` : "";
      if (config.field !== "NOVEDAD DOCUMENTAL" || !docCausePairs.has(countKey)) {
        counts.set(cause, (counts.get(cause) || 0) + 1);
        if (countKey) docCausePairs.add(countKey);
      }
      details.push({
        [config.label]: cause,
        Proveedor: row.PROVEEDOR || "Sin proveedor",
        Material: row.MATERIAL || "",
        SKU: row.SKU || "",
        "Tipo material": sector,
        "N documento": row["N DOCUMENTO"] || "",
        "Fecha estimada": row["FECHA ESTIMADA ENTREGA"] || "",
        "Fecha recibo": row["FECHA RECIBO"] || "",
        "Fecha arribo": row["FECHA ARRIBO"] || "",
      });
    });
  });
  const chartData = [...counts.entries()]
    .sort((a,b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name, value]) => ({ name, value }));
  renderCauseBars(config.chartId, chartData);
  const docTotal = config.field === "NOVEDAD DOCUMENTAL" ? docMap.size : details.length + withoutCause;
  const docWithCause = config.field === "NOVEDAD DOCUMENTAL" ? [...docMap.values()].filter(item => item.hasCause).length : details.length;
  renderCauseStats(config, {
    total: details.length,
    withoutCause,
    docTotal,
    docWithCause,
    docWithoutCause: Math.max(docTotal - docWithCause, 0),
    providers: providers.size,
    sectors: sectors.size,
    materials: materials.size,
    docs: docs.size,
  });
  const rowsOut = details
    .sort((a,b) => String(a[config.label]).localeCompare(String(b[config.label])) || String(a.Proveedor).localeCompare(String(b.Proveedor)))
    .slice(0, 60);
  renderTable(config.tableId, rowsOut, [config.label, "Proveedor", "Material", "SKU", "Tipo material", "N documento", "Fecha estimada", "Fecha recibo", "Fecha arribo"]);
  const dateLabel = dateMode === "arribo" ? "fecha de arribo" : "fecha de recibo";
  document.getElementById(config.footerId).innerHTML = `<span>Mostrando ${rowsOut.length ? 1 : 0} a ${rowsOut.length} de ${details.length} registros</span><span>Filtrado por ${dateLabel}</span>`;
}

function renderCauseStats(config, stats) {
  const items = config.field === "NOVEDAD DOCUMENTAL"
    ? [
        ["Total documentos", stats.docTotal, "doc", "Filtrados"],
        ["Sin novedad", stats.docWithoutCause, "bars", "Documentos"],
        ["Con novedad", stats.docWithCause, "alert", "Documentos"],
        ["Proveedores", stats.providers, "bell", "Con novedad"],
      ]
    : [
        ["Total registros", stats.total, "percent", "100% del total"],
        ["Proveedores", stats.providers, "bell", "Involucrados"],
        ["Materiales", stats.materials, "bars", "Involucrados"],
        ["Sectores", stats.sectors, "folder", "Involucrados"],
      ];
  document.getElementById(config.statsId).innerHTML = items.map(([label, value, icon, note]) =>
    `<div class="cause-stat"><i>${iconSvg(icon)}</i><span>${label}</span><strong>${fmt.format(value || 0)}</strong><small>${note}</small></div>`
  ).join("");
}

function renderCauseBars(id, data) {
  const svg = document.getElementById(id);
  if (!data.length) {
    svg.setAttribute("viewBox", "0 0 760 250");
    svg.innerHTML = `<text x="24" y="116" fill="#64748b" font-size="13" font-weight="700">No hay novedades documentales para graficar en este filtro</text>`;
    return;
  }
  const w = 760, left = 170, top = 14, rowH = 24, right = 54;
  const h = Math.max(250, top + data.length * rowH + 18);
  const max = Math.max(...data.map(d => d.value), 1);
  const palette = ["#063f8f", "#0b63f6", "#1685e8", "#0f9f9a", "#7c3aed", "#f59e0b", "#ef4444"];
  const chartW = w - left - right;
  let html = "";
  data.forEach((d, i) => {
    const y = top + i * rowH;
    const bw = (d.value / max) * chartW;
    const color = palette[i % palette.length];
    html += `<text class="bar-label cause-axis-label" x="4" y="${y + 16}">${escapeHtml(short(d.name, 28))}</text>
      <rect x="${left}" y="${y}" width="${bw}" height="14" rx="4" fill="${color}"><title>${escapeHtml(d.name)}: ${fmt.format(d.value)}</title></rect>
      <text class="bar-label cause-value" x="${left + bw + 6}" y="${y + 11}">${fmt.format(Math.round(d.value))}</text>`;
  });
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.innerHTML = html;
}

function extractCauses(row, config) {
  if (config.mode === "alert") {
    return String(row.ALERTA || "").split(" | ").map(v => v.trim()).filter(v => v && v !== "OK");
  }
  const value = String(row[config.field] || "").trim();
  const normalized = value.toUpperCase();
  if (value && normalized !== "NO" && normalized !== "OK" && normalized !== "SIN NOVEDAD") return [value];
  if (config.field === "CAUSAL NOVEDAD FISICA") {
    const novelty = String(row["NOVEDAD RECIBO"] || "").trim().toUpperCase();
    const obs = String(row.OBSERVACIONES || "").trim();
    const programada = Number(row["CANTIDAD PROGRAMADA"] || 0);
    const recibida = Number(row["CANTIDAD RECIBIDA"] || 0);
    const esGr = row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR";
    const estado = displayLabel(row["ESTADO VEHICULO"] || "");
    if (!esGr && estado === "RECIBIDO" && programada > 0 && recibida === 0) return ["Recibido sin cantidad"];
    if (["SI", "SÍ", "YES", "1"].includes(novelty)) {
      if (obs && obs.toUpperCase() !== "OK") return [obs];
      return ["Con novedad fisica sin causal"];
    }
  }
  return [];
}

function renderCauseProvider(rows) {
  const field = document.getElementById("causeField").value;
  syncCauseSectors(rows);
  const selectedSector = document.getElementById("causeSector").value;
  if (field === "NOVEDAD DOCUMENTAL") {
    renderDocumentNovelty(rows, selectedSector);
    return;
  }
  const map = new Map();
  const chartMap = new Map();
  rows.forEach(row => {
    const sector = row["TIPO MATERIAL"] || "Sin sector";
    if (selectedSector && sector !== selectedSector) return;
    const provider = row.PROVEEDOR || "Sin proveedor";
    let causes = [];
    if (field === "ALERTA") {
      causes = String(row.ALERTA || "").split(" | ").filter(c => c && c !== "OK");
    } else {
      const value = String(row[field] || "").trim();
      if (value && value.toUpperCase() !== "NO") causes = [value];
    }
    causes.forEach(cause => {
      const key = `${provider}|${cause}`;
      if (!map.has(key)) map.set(key, { Proveedor: provider, Causal: cause, Casos: 0 });
      map.get(key).Casos += 1;
      const chartKey = `${cause}|${sector}`;
      if (!chartMap.has(chartKey)) chartMap.set(chartKey, { Causal: cause, Sector: sector, Casos: 0 });
      chartMap.get(chartKey).Casos += 1;
    });
  });
  const tableRows = [...map.values()].sort((a,b) => b.Casos - a.Casos).slice(0, 40);
  const chartRows = [...chartMap.values()].sort((a,b) => b.Casos - a.Casos).slice(0, 30);
  renderCauseMatrixChart("causeProviderChart", chartRows);
  renderTable("causeProviderTable", tableRows, ["Proveedor", "Causal", "Casos"]);
}

function renderDocumentNovelty(rows, selectedSector) {
  const noveltyRows = [];
  let withoutNovelty = 0;
  rows.forEach(row => {
    const sector = row["TIPO MATERIAL"] || "Sin sector";
    if (selectedSector && sector !== selectedSector) return;
    const novelty = String(row["NOVEDAD DOCUMENTAL"] || "").trim();
    const hasNovelty = novelty && novelty.toUpperCase() !== "NO" && novelty.toUpperCase() !== "SIN NOVEDAD";
    if (!hasNovelty) {
      withoutNovelty += 1;
      return;
    }
    noveltyRows.push({
      Proveedor: row.PROVEEDOR || "Sin proveedor",
      Material: row.MATERIAL || "",
      SKU: row.SKU || "",
      "Tipo material": sector,
      "Novedad documental": novelty,
      "N documento": row["N DOCUMENTO"] || "",
      "Fecha estimada": row["FECHA ESTIMADA ENTREGA"] || "",
      "Fecha recibo": row["FECHA RECIBO"] || "",
    });
  });
  const byProvider = new Map();
  noveltyRows.forEach(row => {
    const key = `${row.Proveedor}|${row["Novedad documental"]}`;
    if (!byProvider.has(key)) byProvider.set(key, { Proveedor: row.Proveedor, "Novedad documental": row["Novedad documental"], Casos: 0 });
    byProvider.get(key).Casos += 1;
  });
  const chartData = [...byProvider.values()]
    .sort((a,b) => b.Casos - a.Casos)
    .slice(0, 12)
    .map(row => ({ name: `${short(row.Proveedor, 22)} | ${short(row["Novedad documental"], 18)}`, value: row.Casos }));
  if (!chartData.length) {
    renderBars("causeProviderChart", [{ name: "Sin novedad documental", value: withoutNovelty }], true);
  } else {
    renderBars("causeProviderChart", chartData, true);
  }
  const summaryRows = [
    { Indicador: "Documentos sin novedad documental", Cantidad: withoutNovelty },
    { Indicador: "Documentos con novedad documental", Cantidad: noveltyRows.length },
  ];
  renderTable("causeProviderTable", summaryRows.concat(noveltyRows.slice(0, 80)), ["Indicador", "Cantidad", "Proveedor", "Material", "SKU", "Tipo material", "Novedad documental", "N documento", "Fecha estimada", "Fecha recibo"]);
}

function syncCauseSectors(rows) {
  const select = document.getElementById("causeSector");
  const current = select.value;
  const sectors = [...new Set(rows.map(row => row["TIPO MATERIAL"]).filter(Boolean))].sort();
  select.innerHTML = `<option value="">Todos</option>` + sectors.map(s => `<option>${escapeHtml(s)}</option>`).join("");
  if (sectors.includes(current)) select.value = current;
}

function renderCauseMatrixChart(id, rows) {
  const svg = document.getElementById(id);
  if (!rows.length) return svg.innerHTML = emptySvg();
  const w = 1180, h = 340, p = 52;
  const causes = [...new Set(rows.map(r => r.Causal))].slice(0, 8);
  const sectors = [...new Set(rows.map(r => r.Sector))].slice(0, 5);
  const colors = ["#0b63f6", "#0f9f9a", "#f59e0b", "#7c3aed", "#ef4444"];
  const byKey = new Map(rows.map(r => [`${r.Causal}|${r.Sector}`, r.Casos]));
  const max = Math.max(...rows.map(r => r.Casos), 1);
  const groupW = (w - p * 2) / causes.length;
  const barW = Math.max(Math.min(groupW / Math.max(sectors.length, 1) - 5, 24), 7);
  let html = `<line class="axis" x1="${p}" y1="${h-p}" x2="${w-p}" y2="${h-p}"></line>`;
  causes.forEach((cause, ci) => {
    const gx = p + ci * groupW + 8;
    sectors.forEach((sector, si) => {
      const value = byKey.get(`${cause}|${sector}`) || 0;
      const bh = (value / max) * (h - p * 2);
      const x = gx + si * (barW + 5);
      const y = h - p - bh;
      html += `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" rx="3" fill="${colors[si % colors.length]}"><title>${escapeHtml(cause)} | ${escapeHtml(sector)}: ${value}</title></rect>`;
      if (value) html += `<text class="bar-label" x="${x}" y="${Math.max(y - 5, 12)}">${value}</text>`;
    });
    html += `<text class="bar-label" transform="translate(${gx},${h-8}) rotate(-22)">${escapeHtml(short(cause, 18))}</text>`;
  });
  html += sectors.map((sector, i) => `<rect x="${w-250}" y="${18+i*22}" width="12" height="12" fill="${colors[i % colors.length]}"></rect><text class="bar-label" x="${w-232}" y="${29+i*22}">${escapeHtml(short(sector, 24))}</text>`).join("");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.innerHTML = html;
}

function renderConciliation() {
  if (!dataset || !dataset.conciliation) return;
  syncConciliationProvider();
  const from = document.getElementById("concFrom").value;
  const to = document.getElementById("concTo").value;
  const tipo = document.getElementById("concTipo").value;
  const provider = document.getElementById("concProvider").value;
  const rows = (dataset.conciliation.rows || []).filter(row => {
    if (!dateInRange(row["FECHA RECIBO"], from, to)) return false;
    if (tipo && String(row.TIPO || "") !== tipo) return false;
    if (provider && String(row.PROVEEDOR || "") !== provider) return false;
    return true;
  });
  const counts = { MP: 0, GR: 0, BTS: 0 };
  const uniqueDocs = new Set();
  rows.forEach(row => {
    const t = String(row.TIPO || "");
    const doc = String(row["N DOCUMENTO"] || "").trim();
    if (doc) uniqueDocs.add(doc);
    if (counts[t] != null && doc) counts[t] += 1;
  });
  const totalDoc = uniqueDocs.size;
  const kpis = [
    ["MP", counts.MP, "teal", "doc"],
    ["GR", counts.GR, "", "bars"],
    ["BTS", counts.BTS, "amber", "folder"],
    ["TOTAL DOC", totalDoc, "green", "doc"],
  ];
  document.getElementById("conciliationKpis").innerHTML = kpis.map(([label, value, cls, icon]) =>
    `<div class="kpi ${cls}"><i class="kpi-icon">${iconSvg(icon)}</i><span>${label}</span><strong>${fmt.format(value)}</strong></div>`
  ).join("");
  const sorted = rows.slice().sort((a,b) => {
    const order = { MP: 1, GR: 2, BTS: 3 };
    return (order[a.TIPO] || 9) - (order[b.TIPO] || 9)
      || String(a.PROVEEDOR || "").localeCompare(String(b.PROVEEDOR || ""))
      || String(a["N DOCUMENTO"] || "").localeCompare(String(b["N DOCUMENTO"] || ""), undefined, { numeric:true });
  }).map(row => ({
    _rowClass: `conc-${String(row.TIPO || "").toLowerCase()}`,
    "FECHA RECIBO": row["FECHA RECIBO"],
    TIPO: row.TIPO,
    PROVEEDOR: row.PROVEEDOR,
    "N DOCUMENTO": row["N DOCUMENTO"],
    "Cuenta de TIPO": row["Cuenta de TIPO"],
  }));
  document.getElementById("conciliationTableMeta").textContent = `${fmt.format(totalDoc)} docs | ${fmt.format(rows.length)} lineas`;
  renderTable("conciliationTable", sorted, ["FECHA RECIBO", "TIPO", "PROVEEDOR", "N DOCUMENTO", "Cuenta de TIPO"]);
}

function syncConciliationProvider() {
  const select = document.getElementById("concProvider");
  if (!select || !dataset?.conciliation) return;
  const current = select.value;
  const providers = [...new Set((dataset.conciliation.rows || []).map(row => String(row.PROVEEDOR || "")).filter(Boolean))].sort();
  select.innerHTML = `<option value="">Todos</option>` + providers.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (providers.includes(current)) select.value = current;
}

function renderChartsFromRows(rows) {
  const activeDay = getActiveDay(rows);
  document.getElementById("citaChartTitle").textContent = "Citas (cantidad)";
  renderBars("monthChart", countCitasBySelectedPeriod(rows), false);
  renderBars("statusChart", statusItemsFromRows(rows), true);
  renderBars("alertChart", alertItems(rows), true);
  renderStatusDetail(rows, activeDay);
  renderBars("vehicleChart", statusItemsFromRows(rows), true);
  renderProviderCompliance(rows);
}

function selectedFilterValue(column) {
  const select = document.querySelector(`[data-filter="${column}"]`);
  return select ? select.value : "";
}

function countCitasBySelectedPeriod(rows) {
  const year = selectedFilterValue("ANO CONTROL");
  const month = selectedFilterValue("MES CONTROL");
  const week = selectedFilterValue("SEMANA CONTROL");
  if (week) return countCitasByPeriod(rows, "day");
  if (month) return countCitasByPeriod(rows, "week");
  return countCitasByPeriod(rows, "month");
}

function countCitasByPeriod(rows, mode) {
  const map = new Map();
  const seen = new Set();
  rows.forEach(row => {
    const date = String(row["FECHA CONTROL"] || row["FECHA ESTIMADA ENTREGA"] || row["FECHA PROGRAMADA"] || "").slice(0, 10);
    const key = mode === "day"
      ? (date || "Sin fecha")
      : (mode === "week" ? String(row["SEMANA CONTROL"] || row["SEMANA MATERIAL"] || "Sin semana") : String(row["MES CONTROL"] || row["MES ENTREGA"] || "Sin mes"));
    const label = mode === "day" ? shortDate(key) : key;
    const cita = citaBase(row);
    const uniqueKey = cita ? `${key}|${cita}` : "";
    if (uniqueKey) {
      if (seen.has(uniqueKey)) return;
      seen.add(uniqueKey);
    }
    if (!map.has(key)) map.set(key, { name: label, value: 0, sort: key });
    map.get(key).value += 1;
  });
  return [...map.values()].sort((a,b) => String(a.sort).localeCompare(String(b.sort), undefined, { numeric:true }));
}

function countByDay(rows) {
  const map = new Map();
  const seen = new Set();
  rows.forEach(row => {
    const day = String(row["FECHA CONTROL"] || row["FECHA ESTIMADA ENTREGA"] || row["FECHA PROGRAMADA"] || "Sin fecha").slice(0, 10) || "Sin fecha";
    const cita = citaBase(row);
    const uniqueKey = cita ? `${day}|${cita}` : "";
    if (uniqueKey) {
      if (seen.has(uniqueKey)) return;
      seen.add(uniqueKey);
    }
    map.set(day, (map.get(day) || 0) + 1);
  });
  return [...map.entries()]
    .sort((a,b) => a[0].localeCompare(b[0]))
    .map(([name, value]) => ({ name: shortDate(name), value }));
}

function shortDate(value) {
  const text = String(value || "");
  const parts = text.split("-");
  if (parts.length === 3) return `${parts[2]}/${parts[1]}`;
  return text;
}

function displayLabel(text) {
  const raw = String(text || "");
  if (raw === "EN ATENCION") return "ATENCION";
  if (raw === "PENDIENTE") return "PEND";
  return raw;
}

function statusSortValue(text) {
  const raw = String(text || "").trim().toUpperCase();
  return ({ RECIBIDO: 1, ATENCION: 2, "EN ATENCION": 2, PEND: 3, PENDIENTE: 3 })[raw] || 9;
}

function statusClass(text) {
  const raw = String(text || "").trim().toUpperCase();
  if (raw === "EN ATENCION") return "status-atencion";
  if (raw === "ATENCION") return "status-atencion";
  if (raw === "RECIBIDO") return "status-recibido";
  if (raw === "PEND") return "status-pendiente";
  if (raw === "PENDIENTE") return "status-pendiente";
  return "";
}

function getActiveDay(rows) {
  const estFrom = document.getElementById("estFrom")?.value || "";
  const estTo = document.getElementById("estTo")?.value || "";
  const progFrom = document.getElementById("progFrom")?.value || "";
  const progTo = document.getElementById("progTo")?.value || "";
  if (estFrom && estFrom === estTo) return estFrom;
  if (progFrom && progFrom === progTo) return progFrom;
  const dates = new Set(rows.map(row => String(row["FECHA ESTIMADA ENTREGA"] || row["FECHA CONTROL"] || row["FECHA PROGRAMADA"] || "").slice(0, 10)).filter(Boolean));
  return dates.size === 1 ? [...dates][0] : "";
}

function countByHour(rows) {
  const map = new Map();
  const seen = new Set();
  rows.forEach(row => {
    const hour = hourLabel(row["HORA ARRIBO"] || row["HORA RECIBO"] || row["FECHA PROGRAMADA"]);
    const cita = citaBase(row);
    const uniqueKey = cita ? `${hour}|${cita}` : "";
    if (uniqueKey) {
      if (seen.has(uniqueKey)) return;
      seen.add(uniqueKey);
    }
    map.set(hour, (map.get(hour) || 0) + 1);
  });
  return [...map.entries()]
    .map(([name, value]) => ({ name, value }))
    .sort((a,b) => hourSortValue(a.name) - hourSortValue(b.name));
}

function hourLabel(value) {
  const text = String(value || "").trim();
  if (!text) return "Sin hora";
  const match = text.match(/(\d{1,2})(?::|\.)(\d{2})?/);
  if (match) return `${String(Math.min(Number(match[1]), 23)).padStart(2, "0")}h`;
  const dateMatch = text.match(/T?(\d{1,2}):(\d{2})/);
  if (dateMatch) return `${String(Math.min(Number(dateMatch[1]), 23)).padStart(2, "0")}h`;
  return "Sin hora";
}

function hourSortValue(label) {
  if (label === "Sin hora") return 99;
  const num = Number(String(label).replace(/\D/g, ""));
  return Number.isFinite(num) ? num : 99;
}

function renderStatusDetail(rows, activeDay="") {
  const detail = rows
    .slice()
    .sort((a,b) => statusSortValue(displayLabel(a["ESTADO VEHICULO"] || "")) - statusSortValue(displayLabel(b["ESTADO VEHICULO"] || ""))
      || typeSortValue(tipoCita(a)) - typeSortValue(tipoCita(b))
      || hourSortValue(hourLabel(a["HORA ARRIBO"] || a["HORA RECIBO"])) - hourSortValue(hourLabel(b["HORA ARRIBO"] || b["HORA RECIBO"]))
      || String(a.PROVEEDOR || "").localeCompare(String(b.PROVEEDOR || "")))
    .map(row => {
      const programada = Number(row["CANTIDAD PROGRAMADA"] || 0);
      const recibida = Number(row["CANTIDAD RECIBIDA"] || 0);
      const esGr = row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR";
      const type = tipoCita(row);
      const estado = displayLabel(row["ESTADO VEHICULO"] || "");
      const tieneDiferencia = estado === "RECIBIDO" && !esGr && Math.abs(programada - recibida) > 0.001;
      return {
        _rowClass: `${statusClass(estado)} ${typeClass(type)}${tieneDiferencia ? " qty-alert" : ""}`,
        Estado: estado,
        Tipo: type,
        "Cod cita": cleanIdentifier(row["CODIGO CITA"] || ""),
        "Hora prog.": shortTime(row["HORA ARRIBO"]),
        "Hora est.": shortTime(row["HORA RECIBO"]),
        Proveedor: row.PROVEEDOR || "",
        Material: row["TIPO MATERIAL"] || "",
        SKU: row.SKU || "",
        Descripcion: row.MATERIAL || "",
        "Cant. prog.": programada,
        "Cant. rec.": recibida,
        Unidad: row.UMB || "",
        Placa: row.PLACA || "",
        "Doc. OK": (row["N DOCUMENTO"] && (row.PLACA || row["ESTADO VEHICULO"] === "PENDIENTE")) ? "OK" : (row["ESTADO VEHICULO"] === "PENDIENTE" ? "PEND" : "FALTA"),
        Obs: row.ALERTA || row.OBSERVACIONES || "",
      };
    });
  const title = activeDay ? `Detalle de citas del día - ${formatDateEs(activeDay)}` : "Detalle de citas filtradas";
  document.getElementById("statusDetailTitle").textContent = title;
  document.getElementById("statusDetailCount").textContent = `Total registros: ${fmt.format(detail.length)}`;
  renderTable("statusDetailTable", detail, ["Estado", "Tipo", "Hora prog.", "Hora est.", "Proveedor", "Material", "SKU", "Descripcion", "Cant. prog.", "Cant. rec.", "Unidad", "Placa", "Doc. OK", "Obs"]);
}

function renderStatusDetail(rows, activeDay="") {
  const detail = rows
    .slice()
    .sort((a,b) => statusSortValue(displayLabel(a["ESTADO VEHICULO"] || "")) - statusSortValue(displayLabel(b["ESTADO VEHICULO"] || ""))
      || typeSortValue(tipoCita(a)) - typeSortValue(tipoCita(b))
      || hourSortValue(hourLabel(a["HORA ARRIBO"] || a["HORA RECIBO"])) - hourSortValue(hourLabel(b["HORA ARRIBO"] || b["HORA RECIBO"]))
      || String(a.PROVEEDOR || "").localeCompare(String(b.PROVEEDOR || "")))
    .map(row => {
      const programada = Number(row["CANTIDAD PROGRAMADA"] || 0);
      const recibida = Number(row["CANTIDAD RECIBIDA"] || 0);
      const esGr = row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR";
      const type = tipoCita(row);
      const estado = displayLabel(row["ESTADO VEHICULO"] || "");
      const tieneDiferencia = estado === "RECIBIDO" && !esGr && Math.abs(programada - recibida) > 0.001;
      return {
        _rowClass: `${statusClass(estado)} ${typeClass(type)}${tieneDiferencia ? " qty-alert" : ""}`,
        Estado: estado,
        Tipo: type,
        "Cod cita": cleanIdentifier(row["CODIGO CITA"] || ""),
        "Hora prog.": shortTime(row["HORA ARRIBO"]),
        "Hora est.": shortTime(row["HORA RECIBO"]),
        Proveedor: row.PROVEEDOR || "",
        Material: row["TIPO MATERIAL"] || "",
        SKU: row.SKU || "",
        Descripcion: row.MATERIAL || "",
        "Cant. prog.": programada,
        "Cant. rec.": recibida,
        Unidad: row.UMB || "",
        Placa: estado === "PENDIENTE" ? "" : (row.PLACA || ""),
        "Doc. OK": (row["N DOCUMENTO"] && (row.PLACA || estado === "PENDIENTE")) ? "OK" : (estado === "PENDIENTE" ? "PEND" : "FALTA"),
        Obs: row.ALERTA || row.OBSERVACIONES || "",
      };
    });
  const title = activeDay ? `Detalle de citas del dia - ${formatDateEs(activeDay)}` : "Detalle de citas filtradas";
  document.getElementById("statusDetailTitle").textContent = title;
  document.getElementById("statusDetailCount").textContent = `Total registros: ${fmt.format(detail.length)}`;
  renderTable("statusDetailTable", detail, ["Estado", "Tipo", "Cod cita", "Hora prog.", "Hora est.", "Proveedor", "Material", "SKU", "Descripcion", "Cant. prog.", "Cant. rec.", "Unidad", "Placa", "Doc. OK", "Obs"]);
}

function tipoCita(row) {
  const raw = String(row.TIPO || "").trim().toUpperCase();
  const cita = String(row["CODIGO CITA"] || "").trim().toUpperCase();
  if (raw === "MP" || raw === "GR" || raw === "BTS") return raw;
  if (cita === "BTS") return "BTS";
  return raw || "MP";
}

function typeClass(type) {
  const clean = String(type || "").toLowerCase();
  return clean ? `type-${clean}` : "";
}

function typeSortValue(type) {
  return ({ MP: 1, GR: 2, BTS: 3 })[String(type || "").toUpperCase()] || 9;
}

function shortTime(value) {
  const text = String(value || "").trim();
  const match = text.match(/(\d{1,2})(?::|\.)(\d{2})?/);
  if (!match) return "";
  return `${String(Number(match[1])).padStart(2, "0")}:${match[2] || "00"}`;
}

function formatDateEs(value) {
  const text = String(value || "").slice(0, 10);
  const parts = text.split("-");
  if (parts.length !== 3) return text;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function renderProviderCompliance(rows) {
  const map = new Map();
  rows.forEach(row => {
    if (row["CUMPLE CANTIDAD"] === "No aplica GR" || row["ES GRANEL"]) return;
    const key = String(row.PROVEEDOR || "Sin proveedor");
    if (!map.has(key)) map.set(key, { name: key, citas: 0, cumple: 0 });
    const item = map.get(key);
    item.citas += 1;
    if (row["CUMPLE CANTIDAD"] === "Cumple") item.cumple += 1;
  });
  const data = [...map.values()]
    .map(item => ({ name: item.name, value: item.citas ? item.cumple / item.citas * 100 : 0 }))
    .sort((a,b) => b.value - a.value)
    .slice(0, 10);
  renderBars("providerQtyChart", data, true);
}

function materialFocus(value, materialText="") {
  const base = String(value || "").trim();
  const text = `${value || ""} ${materialText || ""}`.toUpperCase();
  if (text.includes("AZUCAR") || text.includes("AZÚCAR")) return "AZUCAR";
  if (text.includes("PREFORMA")) return "PREFORMAS";
  if (text.includes("LATA")) return "LATAS";
  return base || "SIN TIPO";
}

function renderWeeklyFocus(rows) {
  rows = dataset ? dataset.rows : rows;
  const currentYear = document.getElementById("focusYear").value;
  const currentMonth = document.getElementById("focusMonth").value;
  const currentMaterial = document.getElementById("focusMaterial").value;
  const currentProvider = document.getElementById("focusProvider").value;
  const currentWeek = document.getElementById("focusWeek").value;
  const currentSku = document.getElementById("focusSku").value;
  const currentQuery = document.getElementById("focusSearch").value.trim().toLowerCase();
  const map = new Map();
  const focusRows = rows.filter(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    if (!group) return false;
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    if (currentYear && String(row["ANO CONTROL"] || "") !== currentYear) return false;
    if (currentMonth && String(row["MES CONTROL"] || "") !== currentMonth) return false;
    if (currentMaterial && group !== currentMaterial) return false;
    if (currentProvider && String(row.PROVEEDOR || "") !== currentProvider) return false;
    if (currentWeek && week !== currentWeek) return false;
    if (currentSku && String(row.SKU || "") !== currentSku) return false;
    if (currentQuery) {
      const haystack = [row.SKU, row.MATERIAL, row["TIPO MATERIAL"], row.PROVEEDOR, row["CODIGO CITA"], row["N DOCUMENTO"], row.PLACA]
        .map(value => String(value || "").toLowerCase())
        .join(" ");
      if (!haystack.includes(currentQuery)) return false;
    }
    return true;
  });
  focusRows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    const year = String(row["ANO CONTROL"] || "");
    const month = String(row["MES CONTROL"] || "");
    const key = `${week}|${group}`;
    if (!map.has(key)) {
      map.set(key, { semana: week, material: group, year, month, programada: 0, recibida: 0, citas: 0, graneles: 0, citaSet: new Set() });
    }
    const item = map.get(key);
    item.programada += Number(row["CANTIDAD PROGRAMADA"] || 0);
    item.recibida += Number(row["CANTIDAD RECIBIDA"] || 0);
    item.citas += 1;
    const base = citaBase(row);
    if (base) item.citaSet.add(base);
    if (row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR") item.graneles += 1;
  });
  let data = [...map.values()].sort((a,b) => `${a.semana}-${a.material}`.localeCompare(`${b.semana}-${b.material}`));
  syncFocusYears(data);
  syncFocusMonths(data, currentYear);
  syncFocusMaterials(data);
  syncFocusProviders(rows, currentYear, currentMonth, currentMaterial);
  syncFocusWeeks(data, currentYear, currentMonth);
  syncFocusSkus(rows, currentMaterial, currentWeek, currentYear, currentMonth, currentProvider, currentQuery);
  data = data.slice(-36);
  currentFocusRows = data.map(item => ({
    Semana: item.semana,
    Material: item.material,
    Citas: item.citaSet.size || item.citas,
    "Cantidad programada": item.programada,
    "Cantidad recibida": item.recibida,
    Diferencia: item.graneles === item.citas ? "N/A" : item.recibida - item.programada,
    "Cumplimiento %": item.graneles === item.citas ? "N/A" : compliancePct(item.programada, item.recibida)
  }));
  renderFocusProviderChart("weeklyFocusChart", focusRows, "", "", "", "", "");
  renderTable("weeklyFocusTable", currentFocusRows, ["Semana", "Material", "Citas", "Cantidad programada", "Cantidad recibida", "Diferencia", "Cumplimiento %"]);
  renderFocusStats(currentFocusRows);
  renderFocusDetailTable(focusRows, "", "", "", "", "");
  renderFocusAgenda(focusRows, "", "", "", "", "");
}

function renderFocusStats(rows) {
  const totalProgramada = rows.reduce((sum, row) => sum + Number(row["Cantidad programada"] || 0), 0);
  const totalRecibida = rows.reduce((sum, row) => sum + Number(row["Cantidad recibida"] || 0), 0);
  const diferencia = rows.some(row => row.Diferencia === "N/A") ? "N/A" : totalRecibida - totalProgramada;
  const pendientes = rows.reduce((sum, row) => sum + Math.max(Number(row["Cantidad programada"] || 0) - Number(row["Cantidad recibida"] || 0), 0), 0);
  const cumplimiento = diferencia === "N/A" ? "N/A" : compliancePct(totalProgramada, totalRecibida);
  const stats = [
    ["Programadas", totalProgramada, "green", "calendar"],
    ["Recibidas", totalRecibida, "teal", "truck"],
    ["Pendientes", pendientes, "amber", "clock"],
    ["Diferencia", diferencia, Number(diferencia) < 0 ? "red" : "", "x"],
    ["Cumplimiento", cumplimiento, "green", "percent"],
  ];
  document.getElementById("focusStats").innerHTML = stats.map(([label, value, cls, icon]) =>
    `<div class="focus-stat ${cls}"><i>${iconSvg(icon)}</i><span>${label}</span><strong>${typeof value === "number" ? money.format(value) : value}</strong></div>`
  ).join("");
}

function syncFocusMaterials(data) {
  const select = document.getElementById("focusMaterial");
  const current = select.value;
  const values = [...new Set(data.map(item => item.material).filter(Boolean))].sort((a,b) => a.localeCompare(b));
  select.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) select.value = current;
}

function syncFocusYears(data) {
  const select = document.getElementById("focusYear");
  const current = select.value;
  const values = [...new Set(data.map(item => item.year).filter(Boolean))].sort();
  select.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) select.value = current;
}

function syncFocusMonths(data, selectedYear) {
  const select = document.getElementById("focusMonth");
  const current = select.value;
  let values = data;
  if (selectedYear) values = values.filter(item => item.year === selectedYear);
  const months = [...new Set(values.map(item => item.month).filter(Boolean))].sort();
  select.innerHTML = `<option value="">Todos</option>` + months.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (months.includes(current)) select.value = current;
}

function syncFocusProviders(rows, selectedYear="", selectedMonth="", selectedMaterial="") {
  const select = document.getElementById("focusProvider");
  const current = select.value;
  const providers = new Set();
  rows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    if (!group) return;
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    if (selectedMaterial && group !== selectedMaterial) return;
    if (row.PROVEEDOR) providers.add(String(row.PROVEEDOR));
  });
  const values = [...providers].sort((a,b) => a.localeCompare(b));
  select.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) select.value = current;
}

function renderFocusDetailTable(rows, selectedMaterial, selectedWeek, selectedSku, selectedYear, selectedMonth) {
  const detail = [];
  rows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    if (!group) return;
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    if (selectedMaterial && group !== selectedMaterial) return;
    if (selectedWeek && week !== selectedWeek) return;
    if (selectedSku && String(row.SKU) !== selectedSku) return;
    const programada = Number(row["CANTIDAD PROGRAMADA"] || 0);
    const recibida = Number(row["CANTIDAD RECIBIDA"] || 0);
    detail.push({
      _rowClass: !(row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR") && Math.abs(recibida - programada) > 0.001 ? "qty-alert" : "",
      Semana: week,
      Grupo: group,
      "Fecha estimada entrega": row["FECHA ESTIMADA ENTREGA"],
      "Fecha recibo": row["FECHA RECIBO"],
      SKU: row.SKU,
      Material: row.MATERIAL,
      "Tipo material": row["TIPO MATERIAL"],
      Proveedor: row.PROVEEDOR,
      "Cod cita": row["CODIGO CITA"],
      "Cant programada": programada,
      "Cant recibida": recibida,
      Diferencia: row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR" ? "N/A" : recibida - programada,
      "Cumplimiento %": row["ES GRANEL"] || row["CUMPLE CANTIDAD"] === "No aplica GR" ? "N/A" : compliancePct(programada, recibida)
    });
  });
  detail.sort((a,b) => String(a["Fecha recibo"] || "").localeCompare(String(b["Fecha recibo"] || "")) || String(a.Proveedor).localeCompare(String(b.Proveedor)));
  renderTable("weeklyFocusDetailTable", detail.slice(0, 400), ["Semana", "Grupo", "Fecha estimada entrega", "Fecha recibo", "Proveedor", "SKU", "Material", "Tipo material", "Cod cita", "Cant programada", "Cant recibida", "Diferencia", "Cumplimiento %"]);
}

function renderFocusAgenda(rows, selectedMaterial, selectedWeek, selectedSku, selectedYear, selectedMonth) {
  const today = new Date().toISOString().slice(0, 10);
  const end = new Date();
  end.setDate(end.getDate() + 14);
  const endText = end.toISOString().slice(0, 10);
  const agenda = [];
  rows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    const date = String(row["FECHA CONTROL"] || row["FECHA ESTIMADA ENTREGA"] || "").slice(0, 10);
    if (!date || date < today || date > endText) return;
    if (row["ESTADO VEHICULO"] === "RECIBIDO") return;
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    if (selectedMaterial && group !== selectedMaterial) return;
    if (selectedWeek && week !== selectedWeek) return;
    if (selectedSku && String(row.SKU) !== selectedSku) return;
    const diffDays = Math.ceil((new Date(date) - new Date(today)) / 86400000);
    agenda.push({
      "Fecha estimada": date,
      Proveedor: row.PROVEEDOR || "",
      Material: row.MATERIAL || "",
      SKU: row.SKU || "",
      "Cod cita": row["CODIGO CITA"] || "",
      "Cantidad programada": row["CANTIDAD PROGRAMADA"] || 0,
      "Cantidad recibida": row["CANTIDAD RECIBIDA"] || 0,
      Estado: displayLabel(row["ESTADO VEHICULO"] || ""),
      "Dias restantes": diffDays === 0 ? "Hoy" : `${diffDays} dias`,
    });
  });
  agenda.sort((a,b) => String(a["Fecha estimada"]).localeCompare(String(b["Fecha estimada"])) || String(a.Proveedor).localeCompare(String(b.Proveedor)));
  renderTable("agenda", agenda.slice(0, 80), ["Fecha estimada", "Proveedor", "Material", "SKU", "Cod cita", "Cantidad programada", "Cantidad recibida", "Estado", "Dias restantes"]);
  document.getElementById("agendaCount").textContent = `${fmt.format(agenda.length)} pendientes`;
}

function renderFocusProviderChart(id, rows, selectedMaterial, selectedWeek, selectedSku, selectedYear, selectedMonth) {
  const map = new Map();
  const groupMode = selectedWeek ? "day" : (selectedMonth ? "week" : "month");
  rows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    if (!group) return;
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    if (selectedMaterial && group !== selectedMaterial) return;
    if (selectedWeek && week !== selectedWeek) return;
    if (selectedSku && String(row.SKU) !== selectedSku) return;
    const sku = row.SKU || "Sin SKU";
    const rawDate = String(row["FECHA CONTROL"] || row["FECHA ESTIMADA ENTREGA"] || row["FECHA RECIBO"] || "").slice(0, 10);
    const month = String(row["MES CONTROL"] || row["MES ENTREGA"] || "").trim();
    const period = groupMode === "day" ? shortDate(rawDate || "Sin fecha") : (groupMode === "week" ? week : (month || "Sin mes"));
    const sortKey = groupMode === "day" ? (rawDate || "9999-99-99") : (groupMode === "week" ? week : String(month || "99"));
    const key = `${sortKey}|${sku}`;
    if (!map.has(key)) {
      map.set(key, { name: `${period} | ${sku}`, period, sortKey, sku, programada: 0, recibida: 0, citas: 0 });
    }
    const item = map.get(key);
    item.programada += Number(row["CANTIDAD PROGRAMADA"] || 0);
    item.recibida += Number(row["CANTIDAD RECIBIDA"] || 0);
    item.citas += 1;
  });
  const data = [...map.values()]
    .sort((a,b) => String(a.sortKey).localeCompare(String(b.sortKey), undefined, { numeric:true }) || String(a.sku || "").localeCompare(String(b.sku || ""), undefined, { numeric:true }))
    .slice(-18);
  renderFocusComboChart(id, data, groupMode);
}

function renderFocusComboChart(id, data, groupMode) {
  const svg = document.getElementById(id);
  if (!data.length) return svg.innerHTML = emptySvg();
  const w = 920, h = 400, left = 58, right = 56, top = 118, bottom = 82;
  const chartW = w - left - right;
  const chartH = h - top - bottom;
  const maxQty = Math.max(...data.map(d => Math.max(d.programada, d.recibida)), 1);
  const maxPct = Math.max(100, ...data.map(d => Number(compliancePct(d.programada, d.recibida).replace("%", "")) || 0));
  const slot = chartW / Math.max(data.length, 1);
  const barW = Math.max(Math.min(slot * .26, 18), 7);
  const linePoints = [];
  let html = `
    <line class="axis" x1="${left}" y1="${h-bottom}" x2="${w-right}" y2="${h-bottom}"></line>
    <line class="axis" x1="${left}" y1="${top}" x2="${left}" y2="${h-bottom}"></line>
    <text class="bar-label chart-title-label" x="${left}" y="18">Programada vs recibida por ${groupMode === "day" ? "dia" : (groupMode === "week" ? "semana" : "mes")}</text>
    <g class="chart-legend" transform="translate(${left},34)">
      <rect x="0" y="0" width="10" height="10" rx="2" fill="#0a5df0"></rect><text class="bar-label" x="16" y="10">Programada</text>
      <rect x="112" y="0" width="10" height="10" rx="2" fill="#16a35a"></rect><text class="bar-label" x="128" y="10">Recibida</text>
      <path d="M212 6h24" stroke="#ef8a1a" stroke-width="3" fill="none"></path><text class="bar-label" x="244" y="10">%</text>
    </g>
  `;
  data.forEach((d, i) => {
    const cx = left + i * slot + slot / 2;
    const progH = (d.programada / maxQty) * chartH;
    const recH = (d.recibida / maxQty) * chartH;
    const progY = h - bottom - progH;
    const recY = h - bottom - recH;
    const pct = Number(compliancePct(d.programada, d.recibida).replace("%", "")) || 0;
    const pctY = h - bottom - (Math.min(pct, maxPct) / maxPct) * chartH;
    linePoints.push(`${cx},${pctY}`);
    html += `
      <rect x="${cx - barW - 2}" y="${progY}" width="${barW}" height="${progH}" rx="4" fill="#0a5df0">
        <title>${escapeHtml(d.name)} | Programada: ${fmt.format(d.programada)}</title>
      </rect>
      <rect x="${cx + 2}" y="${recY}" width="${barW}" height="${recH}" rx="4" fill="#16a35a">
        <title>${escapeHtml(d.name)} | Recibida: ${fmt.format(d.recibida)}</title>
      </rect>
      ${barInsideLabel(cx - barW - 2 + barW / 2, progY, progH, d.programada)}
      ${barInsideLabel(cx + 2 + barW / 2, recY, recH, d.recibida)}
      <text class="bar-label x-horizontal" x="${cx}" y="${h-54}" text-anchor="middle">${escapeHtml(short(d.sku, 9))}</text>
      <text class="bar-label x-horizontal" x="${cx}" y="${h-35}" text-anchor="middle">${escapeHtml(short(d.period, 9))}</text>
    `;
  });
  html += `<polyline points="${linePoints.join(" ")}" fill="none" stroke="#ef8a1a" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"></polyline>`;
  data.forEach((d, i) => {
    const cx = left + i * slot + slot / 2;
    const pct = Number(compliancePct(d.programada, d.recibida).replace("%", "")) || 0;
    const pctY = h - bottom - (Math.min(pct, maxPct) / maxPct) * chartH;
    const pctLabelY = Math.max(pctY - 8, top + 14);
    html += `<circle cx="${cx}" cy="${pctY}" r="4" fill="#ef8a1a"></circle><text class="bar-value pct-label" x="${cx + 7}" y="${pctLabelY}">${pct}%</text>`;
  });
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.innerHTML = html;
}

function barInsideLabel(x, y, height, value) {
  if (height < 18 || Number(value || 0) === 0) {
    return `<text class="bar-value vertical-value" transform="translate(${x},${Math.max(y - 6, 38)}) rotate(-90)" text-anchor="middle">${compactQty(value)}</text>`;
  }
  return `<text class="bar-value inbar vertical-value" transform="translate(${x},${y + height / 2}) rotate(-90)" text-anchor="middle">${compactQty(value)}</text>`;
}

function compactQty(value) {
  const n = Number(value || 0);
  if (Math.abs(n) >= 1000000) return `${money.format(n / 1000000)}M`;
  if (Math.abs(n) >= 1000) return `${money.format(n / 1000)}K`;
  return money.format(n);
}

function compliancePct(programada, recibida) {
  const p = Number(programada || 0);
  const r = Number(recibida || 0);
  if (!p) return "0%";
  return `${Math.round((r / p) * 1000) / 10}%`;
}

function syncFocusWeeks(data, selectedYear="", selectedMonth="") {
  const select = document.getElementById("focusWeek");
  const current = select.value;
  let values = data;
  if (selectedYear) values = values.filter(item => item.year === selectedYear);
  if (selectedMonth) values = values.filter(item => item.month === selectedMonth);
  const weeks = [...new Set(values.map(item => item.semana))].sort();
  select.innerHTML = `<option value="">Todas</option>` + weeks.map(w => `<option>${escapeHtml(w)}</option>`).join("");
  if (weeks.includes(current)) select.value = current;
}

function syncFocusSkus(rows, selectedMaterial, selectedWeek, selectedYear="", selectedMonth="", selectedProvider="", selectedQuery="") {
  const select = document.getElementById("focusSku");
  const current = select.value;
  const skus = new Set();
  rows.forEach(row => {
    const group = materialFocus(row["TIPO MATERIAL"], row["MATERIAL"]);
    const week = row["SEMANA MATERIAL"] || row["SEMANA CONTROL"] || "Sin semana";
    if (!group) return;
    if (selectedYear && String(row["ANO CONTROL"] || "") !== selectedYear) return;
    if (selectedMonth && String(row["MES CONTROL"] || "") !== selectedMonth) return;
    if (selectedMaterial && group !== selectedMaterial) return;
    if (selectedProvider && String(row.PROVEEDOR || "") !== selectedProvider) return;
    if (selectedWeek && week !== selectedWeek) return;
    if (selectedQuery) {
      const haystack = [row.SKU, row.MATERIAL, row["TIPO MATERIAL"], row.PROVEEDOR, row["CODIGO CITA"], row["N DOCUMENTO"], row.PLACA]
        .map(value => String(value || "").toLowerCase())
        .join(" ");
      if (!haystack.includes(selectedQuery)) return;
    }
    if (row.SKU !== "" && row.SKU != null) skus.add(String(row.SKU));
  });
  const values = [...skus].sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
  select.innerHTML = `<option value="">Todos</option>` + values.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (values.includes(current)) select.value = current;
}

function renderGroupedFocusChart(id, data) {
  const svg = document.getElementById(id);
  if (!data.length) return svg.innerHTML = emptySvg();
  const w = 1180, h = 340, p = 44;
  const colors = { AZUCAR: "#f59e0b", PREFORMAS: "#0b63f6", LATAS: "#0f9f9a" };
  const weeks = [...new Set(data.map(d => d.semana))].sort();
  const materials = ["AZUCAR", "PREFORMAS", "LATAS"].filter(m => data.some(d => d.material === m));
  const max = Math.max(...data.map(d => d.programada), 1);
  const groupW = (w - p * 2) / Math.max(weeks.length, 1);
  const barW = Math.max(Math.min(groupW / Math.max(materials.length, 1) - 6, 28), 7);
  const byKey = new Map(data.map(d => [`${d.semana}|${d.material}`, d]));
  let html = `<line class="axis" x1="${p}" y1="${h-p}" x2="${w-p}" y2="${h-p}"></line>`;
  weeks.forEach((week, wi) => {
    const gx = p + wi * groupW + 8;
    materials.forEach((mat, mi) => {
      const item = byKey.get(`${week}|${mat}`);
      if (!item) return;
      const bh = (item.programada / max) * (h - p * 2);
      const x = gx + mi * (barW + 5);
      const y = h - p - bh;
      html += `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" rx="3" fill="${colors[mat]}"><title>${week} ${mat}: ${fmt.format(item.programada)} programada | ${fmt.format(item.recibida)} recibida | ${item.citas} citas</title></rect>`;
      if (bh > 24) html += `<text class="bar-label" x="${x-2}" y="${y-5}">${fmt.format(Math.round(item.programada))}</text>`;
    });
    html += `<text class="bar-label" transform="translate(${gx},${h-9}) rotate(-25)">${escapeHtml(week)}</text>`;
  });
  html += materials.map((mat, i) => `<rect x="${w-210}" y="${20+i*22}" width="12" height="12" fill="${colors[mat]}"></rect><text class="bar-label" x="${w-192}" y="${31+i*22}">${mat}</text>`).join("");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.innerHTML = html;
}

function renderLine(id, data) {
  const svg = document.getElementById(id);
  if (!data.length) return svg.innerHTML = emptySvg();
  const w = 720, h = 250, p = 34;
  const max = Math.max(...data.map(d => d.value), 1);
  const points = data.map((d, i) => {
    const x = p + i * ((w - p*2) / Math.max(data.length - 1, 1));
    const y = h - p - (d.value / max) * (h - p*2);
    return {x, y, ...d};
  });
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.innerHTML = `<line class="axis" x1="${p}" y1="${h-p}" x2="${w-p}" y2="${h-p}"></line>
    <line class="axis" x1="${p}" y1="${p}" x2="${p}" y2="${h-p}"></line>
    <polyline points="${points.map(p => `${p.x},${p.y}`).join(" ")}" fill="none" stroke="#0b63f6" stroke-width="3"/>
    ${points.map(p => `<circle cx="${p.x}" cy="${p.y}" r="4" fill="#0b63f6"><title>${p.name}: ${fmt.format(p.value)}</title></circle>`).join("")}
    ${points.map((p,i) => i % Math.ceil(points.length / 7 || 1) === 0 ? `<text class="bar-label" x="${p.x-22}" y="${h-8}">${escapeHtml(p.name)}</text>` : "").join("")}`;
}

function renderBars(id, data, horizontal) {
  const svg = document.getElementById(id);
  if (!data.length) return svg.innerHTML = emptySvg();
  const w = 720, h = 300, p = 42;
  const max = Math.max(...data.map(d => d.value), 1);
  const palette = ["#0b3f86", "#0a5df0", "#16a35a", "#0ea5c6", "#ef8a1a", "#dc2f38", "#6d4be8", "#123d68"];
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  if (horizontal) {
    const labelW = 230;
    const rowH = (h - 24) / Math.max(data.length, 1);
    svg.innerHTML = data.map((d, i) => {
      const y = 14 + i * rowH;
      const bw = Math.max((d.value / max) * (w - labelW - 64), 5);
      const color = palette[i % palette.length];
      const barH = Math.max(Math.min(rowH - 8, 20), 12);
      return `<text class="bar-label" x="6" y="${y+barH-3}">${escapeHtml(short(displayLabel(d.name), 28))}</text>
        <rect x="${labelW}" y="${y}" width="${bw}" height="${barH}" rx="5" fill="${color}"><title>${escapeHtml(d.name)}: ${fmt.format(d.value)}</title></rect>
        <text class="bar-value" x="${labelW + bw + 9}" y="${y+barH-3}">${fmt.format(Math.round(d.value))}</text>`;
    }).join("");
  } else {
    const colW = (w - p*2) / data.length;
    svg.innerHTML = `<line class="axis" x1="${p}" y1="${h-p}" x2="${w-p}" y2="${h-p}"></line>` + data.map((d, i) => {
      const bh = (d.value / max) * (h - p*2);
      const x = p + i * colW + 6;
      const y = h - p - bh;
      const color = palette[i % palette.length];
      return `<rect x="${x}" y="${y}" width="${Math.max(colW-14, 10)}" height="${bh}" rx="5" fill="${color}"><title>${escapeHtml(d.name)}: ${fmt.format(d.value)}</title></rect>
        <text class="bar-value" x="${x}" y="${Math.max(y - 7, 14)}">${fmt.format(Math.round(d.value))}</text>
        <text class="bar-label" transform="translate(${x+2},${h-10}) rotate(-18)">${escapeHtml(short(displayLabel(d.name), 11))}</text>`;
    }).join("");
  }
}

function renderPie(id, data) {
  const svg = document.getElementById(id);
  if (!data.length) return svg.innerHTML = emptySvg();
  const total = data.reduce((s,d) => s+d.value, 0) || 1;
  const colors = ["#0b63f6","#ef4444","#16a34a","#7c3aed","#0f9f9a","#f59e0b","#063f8f"];
  let angle = -90;
  const cx = 170, cy = 125, r = 88;
  const slices = data.map((d, i) => {
    const delta = d.value / total * 360;
    const path = describeArc(cx, cy, r, angle, angle + delta);
    angle += delta;
    return `<path d="${path}" fill="${colors[i % colors.length]}"><title>${escapeHtml(d.name)}: ${fmt.format(d.value)}</title></path>`;
  }).join("");
  const legend = data.map((d,i) => `<rect x="310" y="${35+i*25}" width="12" height="12" fill="${colors[i % colors.length]}"></rect><text class="bar-label" x="330" y="${46+i*25}">${escapeHtml(short(d.name, 30))} (${Math.round(d.value/total*100)}%)</text>`).join("");
  svg.setAttribute("viewBox", "0 0 720 250");
  svg.innerHTML = slices + legend;
}

function describeArc(cx, cy, r, startAngle, endAngle) {
  const start = polar(cx, cy, r, endAngle), end = polar(cx, cy, r, startAngle);
  const large = endAngle - startAngle <= 180 ? "0" : "1";
  return `M ${cx} ${cy} L ${start.x} ${start.y} A ${r} ${r} 0 ${large} 0 ${end.x} ${end.y} Z`;
}
function polar(cx, cy, r, angle) {
  const rad = (angle - 90) * Math.PI / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function renderTable(id, rows, cols) {
  const table = document.getElementById(id);
  if (!rows.length || !cols.length) return table.innerHTML = `<tr><td class="empty">Sin registros para este filtro</td></tr>`;
  table.innerHTML = `<thead><tr>${cols.map(c => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead><tbody>` +
    rows.map(row => `<tr class="${escapeHtml(row._rowClass || "")}">${cols.map(c => `<td>${cellValue(c, row[c])}</td>`).join("")}</tr>`).join("") + `</tbody>`;
}

function cellValue(col, value) {
  if (col === "Estado") return `<span class="status-chip ${escapeHtml(String(value).replaceAll(" ", "_"))}">${escapeHtml(value)}</span>`;
  if (col === "TIPO" || col === "Tipo") return `<span class="type-chip ${escapeHtml(value)}">${escapeHtml(value)}</span>`;
  if (col === "Doc. OK") {
    const ok = String(value || "") === "OK";
    const pend = String(value || "") === "PEND";
    if (pend) return `<span class="status-chip PENDIENTE">PEND</span>`;
    return `<span class="${ok ? "doc-ok" : "doc-bad"}">${ok ? "✓" : "×"}</span>`;
  }
  if (col === "ESTADO CONTROL" || col === "ESTADO VEHICULO") return `<span class="pill ${escapeHtml(String(value).replaceAll(" ", "_"))}">${escapeHtml(value)}</span>`;
  if (isIdentifierColumn(col)) return escapeHtml(cleanIdentifier(value));
  if (typeof value === "number") return money.format(value);
  return escapeHtml(value ?? "");
}

function cellValue(col, value) {
  if (col === "Estado") return `<span class="status-chip ${escapeHtml(String(value).replaceAll(" ", "_"))}">${escapeHtml(value)}</span>`;
  if (col === "TIPO" || col === "Tipo") return `<span class="type-chip ${escapeHtml(value)}">${escapeHtml(value)}</span>`;
  if (col === "Doc. OK") {
    const ok = String(value || "") === "OK";
    const pend = String(value || "") === "PEND";
    if (pend) return `<span class="status-chip PENDIENTE">PEND</span>`;
    return `<span class="${ok ? "doc-ok" : "doc-bad"}">${ok ? "&#10003;" : "&#10005;"}</span>`;
  }
  if (col === "ESTADO CONTROL" || col === "ESTADO VEHICULO") return `<span class="pill ${escapeHtml(String(value).replaceAll(" ", "_"))}">${escapeHtml(value)}</span>`;
  if (isIdentifierColumn(col)) return escapeHtml(cleanIdentifier(value));
  if (typeof value === "number") return money.format(value);
  return escapeHtml(value ?? "");
}

function isIdentifierColumn(col) {
  const normalized = String(col || "").toUpperCase();
  return normalized === "SKU" || normalized.includes("COD CITA") || normalized.includes("CODIGO CITA") || normalized.includes("CITA TRAFFIC") || normalized.includes("N DOCUMENTO");
}

function cleanIdentifier(value) {
  if (value == null) return "";
  let text = String(value).trim();
  if (!text) return "";
  text = text.replace(/\\.0$/, "");
  if (/^[\\d.]+$/.test(text)) text = text.replaceAll(".", "");
  return text;
}

function downloadCsv() {
  const cols = dataset.columns;
  const lines = [cols.join(",")].concat(filteredRows.map(row => cols.map(col => csvCell(row[col])).join(",")));
  const blob = new Blob([lines.join("\n")], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "plan_recibo_filtrado.csv";
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportCurrentModule() {
  if (!dataset || currentPage === "home") return;
  const previousStatus = document.getElementById("status").textContent;
  document.getElementById("status").textContent = "Generando captura...";
  const titles = {
    dashboard: "Indicadores",
    proveedores: "Proveedores y causales",
    recibo: "Recibo y materiales",
    conciliacion: "Conciliacion",
    base: "Base filtrada",
  };
  const moduleTitle = titles[currentPage] || currentPage;
  const capture = document.createElement("section");
  capture.className = "module-capture";
  if (currentPage !== "dashboard") {
    const header = document.createElement("div");
    header.className = "capture-head";
    header.innerHTML = `<div><h1>${escapeHtml(moduleTitle)}</h1><p>${escapeHtml(dataset.fileName || "")} | generado ${escapeHtml(dataset.generatedAt || "")}</p></div><strong>Plan de Recibo 2026.2</strong>`;
    capture.appendChild(header);
  }
  if (currentPage === "dashboard") {
    capture.appendChild(prepCaptureNode(document.getElementById("kpis")));
  }
  document.querySelectorAll(`[data-page="${currentPage}"]`).forEach(node => {
    capture.appendChild(prepCaptureNode(node));
  });
  const css = [...document.querySelectorAll("style")].map(style => style.innerHTML).join("\n");
  const wrapper = document.createElement("div");
  wrapper.style.position = "fixed";
  wrapper.style.left = "-20000px";
  wrapper.style.top = "0";
  wrapper.style.width = "1860px";
  wrapper.style.background = "#eef4fb";
  wrapper.style.padding = "16px";
  wrapper.appendChild(capture);
  document.body.appendChild(wrapper);
  const width = Math.ceil(capture.scrollWidth + 32);
  const height = Math.ceil(capture.scrollHeight + 32);
  const scale = Math.max(1, Math.min(2, 12000 / Math.max(width, height)));
  const xhtml = new XMLSerializer().serializeToString(capture);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
    <foreignObject width="100%" height="100%">
      <div xmlns="http://www.w3.org/1999/xhtml">
        <style>${css}
          body{margin:0;background:#eef4fb;}
          .module-capture{width:100%;box-sizing:border-box;background:#eef4fb;color:#061a38;font-family:Segoe UI,Arial,sans-serif;}
          .capture-head{display:flex;justify-content:space-between;align-items:flex-start;margin:0 0 12px;padding:14px 18px;border-radius:16px;background:#061a38;color:white;}
          .capture-head h1{font-size:26px;line-height:1.05;margin:0 0 5px;}
          .capture-head p{margin:0;color:#c7d7ec;font-size:12px;font-weight:800;}
          .capture-head strong{font-size:13px;color:#dbeafe;}
          .page-hidden{display:block!important;}
          .module-toolbar,.module-link,button{display:none!important;}
          .grid{display:grid!important;grid-template-columns:repeat(12,minmax(0,1fr));gap:14px;}
          .kpis{display:grid!important;}
          .filters{display:grid!important;}
          .panel,.kpi,.filters{box-shadow:none!important;}
          .dashboard-detail{grid-column:1/-1!important;width:100%!important;}
        </style>${xhtml}
      </div>
    </foreignObject>
  </svg>`;
  const img = new Image();
  const svgBlob = new Blob([svg], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(svgBlob);
  let exported = false;
  const fileBase = `captura_${moduleTitle.toLowerCase().replaceAll(" ", "_")}`;
  const cleanup = () => {
    URL.revokeObjectURL(url);
    if (wrapper.parentNode) document.body.removeChild(wrapper);
    document.getElementById("status").textContent = previousStatus || "Dashboard listo";
  };
  const fallbackTimer = setTimeout(() => {
    if (exported) return;
    exported = true;
    downloadBlob(svgBlob, `${fileBase}.svg`);
    cleanup();
  }, 2500);
  img.onload = () => {
    if (exported) return;
    try {
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(width * scale);
      canvas.height = Math.ceil(height * scale);
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#eef4fb";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      canvas.toBlob(blob => {
        if (exported) return;
        clearTimeout(fallbackTimer);
        exported = true;
        downloadBlob(blob || svgBlob, `${fileBase}.${blob ? "png" : "svg"}`);
        cleanup();
      }, "image/png", 0.95);
    } catch (error) {
      clearTimeout(fallbackTimer);
      exported = true;
      downloadBlob(svgBlob, `${fileBase}.svg`);
      cleanup();
    }
  };
  img.onerror = () => {
    if (exported) return;
    clearTimeout(fallbackTimer);
    exported = true;
    downloadBlob(svgBlob, `${fileBase}.svg`);
    cleanup();
  };
  img.src = url;
}

function downloadBlob(blob, filename) {
  if (!blob) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 15000);
}

function prepCaptureNode(node) {
  const clone = node.cloneNode(true);
  clone.classList.remove("page-hidden");
  clone.querySelectorAll("input,select,textarea").forEach((field, index) => {
    const source = node.querySelectorAll("input,select,textarea")[index];
    if (!source) return;
    if (field.tagName === "SELECT") {
      [...field.options].forEach(option => option.removeAttribute("selected"));
      if (field.options[source.selectedIndex]) field.options[source.selectedIndex].setAttribute("selected", "selected");
    } else {
      field.setAttribute("value", source.value || "");
    }
  });
  return clone;
}

function csvCell(value) {
  const text = String(value ?? "").replaceAll('"', '""');
  return `"${text}"`;
}

function short(text, len) {
  text = String(text || "");
  return text.length > len ? text.slice(0, len-1) + "..." : text;
}
function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}
function emptySvg() {
  return `<text x="30" y="120" fill="#64748b">Sin datos para este filtro</text>`;
}

document.getElementById("analyze").onclick = () => upload(false);
document.getElementById("demo").onclick = () => upload(true);
document.getElementById("download").onclick = downloadCsv;
document.getElementById("exportModule").onclick = exportCurrentModule;
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        return

    def send_bytes(self, data: bytes, content_type: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        safe_payload = make_json_safe(payload)
        self.send_bytes(json.dumps(safe_payload, ensure_ascii=False, default=json_value).encode("utf-8"), "application/json; charset=utf-8", status)

    def do_GET(self) -> None:
        if self.path in {"/", "/index.html"}:
            self.send_bytes(HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        self.send_json({"error": "Ruta no encontrada"}, HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        try:
            if self.path == "/api/demo":
                if not DEFAULT_FILE.exists():
                    self.send_json({"error": f"No encontré el archivo demo: {DEFAULT_FILE}"}, HTTPStatus.NOT_FOUND)
                    return
                self.send_json(analyze(DEFAULT_FILE, DEFAULT_FILE.name))
                return

            if self.path == "/api/analyze":
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length)
                file_name, file_obj = parse_multipart(body, self.headers.get("Content-Type", ""))
                self.send_json(analyze(file_obj, file_name))
                return

            self.send_json({"error": "Ruta no encontrada"}, HTTPStatus.NOT_FOUND)
        except Exception as exc:
            self.send_json({"error": f"No pude procesar el archivo: {exc}"}, HTTPStatus.BAD_REQUEST)


def main() -> None:
    server = ThreadingHTTPServer((APP_HOST, APP_PORT), Handler)
    url = f"http://{APP_HOST}:{APP_PORT}"
    print(f"Dashboard disponible en {url}")
    print("Presiona Ctrl+C para detenerlo.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    server.serve_forever()


if __name__ == "__main__":
    main()
