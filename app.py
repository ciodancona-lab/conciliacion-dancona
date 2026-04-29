# app_conciliacion_dancona_v2.py
# Versión estable V3: conciliación continua, pendientes anteriores, regularizaciones, apertura y matching agregado MB-EXT.
# Streamlit Cloud requirements sugeridos:
# streamlit
# pandas
# numpy
# openpyxl
# xlrd
# xlwt

import io
import re
import uuid
import warnings
import json
import base64
import sys
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import xlwt
try:
    import requests
except Exception:
    requests = None
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# =============================================================================
# CONFIG
# =============================================================================

st.set_page_config(
    page_title="Conciliación Bancaria Dancona · V3",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOCAL_MAP = {
    "29841642": "Local 1",
    "29841627": "Local 2",
    "29841683": "Local 3",
    "29841670": "Local 5",
    "29841705": "Local 6",
    "31995644": "Local 7",
    "32032899": "Local 8",
}

IMP_CATS = {
    "IMPUESTO_LEY25413_DEB",
    "IMPUESTO_LEY25413_CRED",
    "RETENCION_IIBB",
    "IVA",
    "GASTO_BANCARIO",
    "DEBITO_LIQ_TARJETA",
    "DEBITO_PAGO_DIRECTO",
}

CAT_LABELS = {
    "QR": "QR / CR DEBIN SPOT",
    "LIQUIDACION_TARJETA": "Liquidación tarjeta",
    "DEBITO_LIQ_TARJETA": "Débito liquidación tarjeta",
    "TRANSFERENCIA_ENTRANTE": "Transferencia entrante",
    "TRANSFERENCIA_SALIENTE": "Transferencia saliente",
    "IMPUESTO_LEY25413_DEB": "Anticipo Ganancias Ley 25.413 débitos",
    "IMPUESTO_LEY25413_CRED": "Anticipo Ganancias Ley 25.413 créditos",
    "RETENCION_IIBB": "Retención Ingresos Brutos Mendoza",
    "IVA": "IVA base / IVA sobre gastos",
    "GASTO_BANCARIO": "Gasto bancario / comisión",
    "DEBITO_PAGO_DIRECTO": "Débito pago directo",
    "OTRO": "Otro",
}

# =============================================================================
# UTILIDADES
# =============================================================================

def parse_ar_num(value) -> float:
    """Convierte números argentinos/mixtos a float."""
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return 0.0
    s = s.replace("$", "").replace(" ", "").replace("\u00a0", "")
    # Negativos tipo (1.234,56)
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    # Si hay coma, asumimos formato AR: 1.234,56
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        out = float(s)
        return -out if neg else out
    except Exception:
        return 0.0


def norm_txt(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip()).upper()


def safe_date(x) -> Optional[pd.Timestamp]:
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.to_datetime(x, errors="coerce")
    s = str(x).strip()
    # En QR suele venir "dd/mm/yyyy hh:mm:ss"
    s10 = s[:10]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return pd.to_datetime(datetime.strptime(s10, fmt))
        except Exception:
            pass
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def fmt_date(ts) -> str:
    ts = safe_date(ts)
    if pd.isna(ts):
        return ""
    return ts.strftime("%d/%m/%Y")


def amount_equal(a: float, b: float, tol: float = 1.00) -> bool:
    return abs(abs(float(a)) - abs(float(b))) <= tol


def classify_bank(concepto: str) -> str:
    c = norm_txt(concepto)
    if "CR DEBIN SPOT" in c:
        return "QR"
    if "CR LIQ" in c or ("AMEX" in c and "CR" in c):
        return "LIQUIDACION_TARJETA"
    if "DEB LIQ" in c or "DEB.LIQ" in c:
        return "DEBITO_LIQ_TARJETA"
    if "GRAVAMEN LEY 25413 S/DEB" in c:
        return "IMPUESTO_LEY25413_DEB"
    if "GRAVAMEN LEY 25413 S/CRED" in c:
        return "IMPUESTO_LEY25413_CRED"
    if "INGRESOS BRUTOS" in c or "IIBB" in c:
        return "RETENCION_IIBB"
    if "I.V.A. BASE" in c or "IVA BASE" in c:
        return "IVA"
    if "COM TRANSFE" in c or "COM. TRANSFE" in c:
        return "GASTO_BANCARIO"
    if "C BE TR O/BCO" in c or "TRANSF.INT.DIST.TITULAR" in c or "CRED.TRAN" in c:
        return "TRANSFERENCIA_ENTRANTE"
    if "DB CREDIN" in c or "DEB.TRAN.INTERBMISMO" in c or "DEBITO TRANSF" in c:
        return "TRANSFERENCIA_SALIENTE"
    if "DEBITO PAGO DIRECTO" in c or "PAGO DIRECTO" in c:
        return "DEBITO_PAGO_DIRECTO"
    return "OTRO"


def category_compatible(pending_cat: str, actual_cat: str, pending_side: str, actual_type: str = "") -> bool:
    """Compatibilidad flexible para regularizar pendientes anteriores."""
    if pending_cat == actual_cat:
        return True
    # Banco egreso anterior suele aparecer ahora como MB-EXT, sin categoría fina del lado Flexxus.
    if pending_side == "BANCO_NO_FLEXXUS" and actual_type in {"MB-EXT", "MB-ENT-EX"}:
        return True
    # Banco ingreso anterior puede aparecer como PAV.
    if pending_side == "BANCO_NO_FLEXXUS" and actual_type == "PAV":
        return True
    # Flexxus PAV anterior puede acreditarse como QR, tarjeta o transferencia.
    if pending_side == "FLEXXUS_NO_BANCO" and actual_cat in {"QR", "LIQUIDACION_TARJETA", "TRANSFERENCIA_ENTRANTE", "OTRO"}:
        return True
    # Flexxus egreso anterior puede debitarse como transferencia o débito.
    if pending_side == "FLEXXUS_NO_BANCO" and actual_cat in {"TRANSFERENCIA_SALIENTE", "DEBITO_PAGO_DIRECTO", "OTRO"}:
        return True
    return False

# =============================================================================
# PARSERS
# =============================================================================

def parse_flexxus(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str, header=None)
    rows = []
    for _, row in df.iterrows():
        fecha = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
        if "/" not in fecha:
            continue
        tipo = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        numero = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
        movimiento = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
        debe = parse_ar_num(row.iloc[7]) if len(row) > 7 else 0.0
        haber = parse_ar_num(row.iloc[9]) if len(row) > 9 else 0.0
        saldo = parse_ar_num(row.iloc[10]) if len(row) > 10 else 0.0
        if tipo not in {"PAV", "MB-ENT-EX", "MB-EXT"}:
            continue
        # PAV suele ser Debe; egresos suelen estar en Haber. Tomamos el no-cero.
        monto = debe if abs(debe) > 0.005 else haber
        if abs(monto) <= 0.005:
            continue
        rows.append({
            "row_id": f"F-{uuid.uuid4().hex[:10]}",
            "FechaFlexxus": fecha,
            "FechaFlexxus_dt": safe_date(fecha),
            "Tipo": tipo,
            "Numero": numero,
            "Movimiento": movimiento,
            "ConceptoNorm": norm_txt(movimiento),
            "MontoFlexxus": abs(float(monto)),
            "SaldoFlexxus": saldo,
            "EsIngresoFlexxus": tipo == "PAV",
            "EsEgresoFlexxus": tipo in {"MB-ENT-EX", "MB-EXT"},
            "EsPedidosYa": "PEDIDOS YA" in norm_txt(movimiento),
            "Matched": False,
            "ConsumedPrev": False,
            "MatchStage": "",
            "MatchRef": "",
            "BancoFecha": "",
            "BancoConcepto": "",
            "BancoComprobante": "",
            "FechaAcreditacionUsada": "",
            "Diagnostico": "",
        })
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("No se encontraron movimientos PAV, MB-ENT-EX o MB-EXT en Flexxus.")
    return out


def parse_banco(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str, header=None)
    header_row = None
    for i, row in df.iterrows():
        if norm_txt(row.iloc[0] if len(row) > 0 else "") == "FECHA":
            header_row = i
            break
    if header_row is None:
        # fallback: usar fila 4 como en Banco Nación exportado
        header_row = 4
    raw = df.iloc[header_row + 1:].copy()
    raw = raw.iloc[:, :5]
    raw.columns = ["Fecha", "Comprobante", "Concepto", "Importe", "Saldo"]
    raw = raw.dropna(subset=["Fecha", "Concepto"])
    raw = raw[raw["Fecha"].astype(str).str.strip().ne("")]
    rows = []
    for source_order, (_, r) in enumerate(raw.iterrows()):
        imp = parse_ar_num(r["Importe"])
        if abs(imp) <= 0.005:
            continue
        rows.append({
            "row_id": f"B-{uuid.uuid4().hex[:10]}",
            "SourceOrder": source_order,
            "Fecha": str(r["Fecha"]).strip(),
            "Fecha_dt": safe_date(r["Fecha"]),
            "Comprobante": str(r["Comprobante"]).strip() if pd.notna(r["Comprobante"]) else "",
            "Concepto": str(r["Concepto"]).strip() if pd.notna(r["Concepto"]) else "",
            "ConceptoNorm": norm_txt(r["Concepto"]),
            "ImporteNum": imp,
            "ImporteAbs": abs(imp),
            "SaldoNum": parse_ar_num(r["Saldo"]),
            "EsIngreso": imp > 0,
            "EsEgreso": imp < 0,
            "Categoria": classify_bank(r["Concepto"]),
            "Matched": False,
            "ConsumedPrev": False,
            "MatchStage": "",
            "MatchRef": "",
            "FlexxusNumero": "",
            "Diagnostico": "",
        })
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("No se encontraron movimientos bancarios válidos.")
    return out.reset_index(drop=True)


def parse_qr(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    qr = pd.read_excel(file, dtype=str)
    if qr.empty:
        return qr
    if "Monto total" in qr.columns:
        qr["MontoTotal"] = qr["Monto total"].apply(parse_ar_num)
    else:
        qr["MontoTotal"] = 0.0
    qr["NetoQR"] = qr["MontoTotal"] * 0.99032
    qr["FechaQR"] = qr.get("Fecha", "").astype(str).str[:10]
    qr["FechaQR_dt"] = qr["FechaQR"].apply(safe_date)
    qr["CodComercio"] = qr.get("Cód. comercio", "").astype(str)
    qr["Cupon"] = qr.get("Ticket", qr.get("Id QR", "")).astype(str)
    return qr


def parse_trx(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    trx = pd.read_excel(file, dtype=str)
    if trx.empty:
        return trx
    if "IMPORTE NETO" in trx.columns:
        trx["MontoNeto"] = trx["IMPORTE NETO"].apply(parse_ar_num)
    elif "TOTAL LIQUIDACION" in trx.columns:
        trx["MontoNeto"] = trx["TOTAL LIQUIDACION"].apply(parse_ar_num)
    else:
        trx["MontoNeto"] = 0.0
    if "FECHA DE PAGO" in trx.columns:
        trx["FechaPago_dt"] = trx["FECHA DE PAGO"].apply(safe_date)
    if "COMERCIO" in trx.columns:
        trx["Local"] = trx["COMERCIO"].astype(str).map(LOCAL_MAP).fillna(trx["COMERCIO"].astype(str))
    return trx

# =============================================================================
# CONCILIACIÓN ANTERIOR / PENDIENTES
# =============================================================================

def _find_header_row(df: pd.DataFrame, required_any: List[str]) -> Optional[int]:
    for i in range(min(len(df), 20)):
        row_txt = " | ".join(norm_txt(x) for x in df.iloc[i].tolist())
        if any(norm_txt(r) in row_txt for r in required_any):
            return i
    return None


def parse_previous_conciliation(file) -> Tuple[pd.DataFrame, Dict[str, float], str]:
    """Lee conciliación anterior nueva o vieja.

    Devuelve pendientes abiertos con signo lógico y resumen anterior.
    Si el archivo no tiene hoja 'Pendientes abiertos', intenta leer hojas del reporte viejo.
    """
    if file is None:
        return pd.DataFrame(columns=[
            "pending_id", "FechaOrigen", "Origen", "TipoPendiente", "TipoMovimiento",
            "Numero", "Concepto", "Categoria", "Monto", "SignoCalculo", "Estado", "Fuente"
        ]), {}, "APERTURA"

    xl = pd.ExcelFile(file)
    pendings = []
    prev = {}
    mode = "ANTERIOR"

    # Resumen anterior
    if "Conciliacion semanal" in xl.sheet_names:
        ws = pd.read_excel(file, sheet_name="Conciliacion semanal", header=None, dtype=str)
        for _, row in ws.iterrows():
            label = norm_txt(row.iloc[0] if len(row) > 0 else "")
            val = parse_ar_num(row.iloc[1] if len(row) > 1 else 0)
            if "SALDO SEGÚN EXTRACTO" in label or "SALDO S/EXTRACTO" in label:
                prev["saldo_banco_anterior"] = val
            elif "SALDO S/FLEXXUS" in label and "PASANDO" not in label and "CALCULADO" not in label:
                prev["saldo_flexxus_anterior"] = val
            elif "DIFERENCIA FINAL" in label:
                prev["diferencia_final_anterior"] = val

    # Formato nuevo: Pendientes abiertos
    if "Pendientes abiertos" in xl.sheet_names:
        pa_raw = pd.read_excel(file, sheet_name="Pendientes abiertos", header=None, dtype=str)
        h = _find_header_row(pa_raw, ["Tipo pendiente", "Origen", "Importe", "Monto"])
        if h is not None:
            pa = pd.read_excel(file, sheet_name="Pendientes abiertos", header=h, dtype=str)
            for _, r in pa.iterrows():
                concepto = str(r.get("Concepto", r.get("Movimiento", ""))).strip()
                monto = parse_ar_num(r.get("Importe", r.get("Monto", 0)))
                if not concepto or abs(monto) <= 0.005:
                    continue
                origen = str(r.get("Origen", r.get("Tipo pendiente", ""))).strip().upper()
                tipo_p = str(r.get("Tipo pendiente", origen)).strip().upper()
                categoria = str(r.get("Categoría", r.get("Categoria", "OTRO"))).strip() or "OTRO"
                signo = infer_pending_sign(origen, tipo_p, categoria, concepto, monto)
                pendings.append({
                    "pending_id": str(r.get("ID", "")) or f"P-{uuid.uuid4().hex[:10]}",
                    "FechaOrigen": str(r.get("Fecha origen", r.get("Fecha", ""))).strip(),
                    "Origen": normalize_pending_origin(origen, tipo_p),
                    "TipoPendiente": tipo_p,
                    "TipoMovimiento": str(r.get("Tipo", "")).strip(),
                    "Numero": str(r.get("Número", r.get("Numero", ""))).strip(),
                    "Concepto": concepto,
                    "Categoria": categoria,
                    "Monto": abs(monto),
                    "SignoCalculo": signo,
                    "Estado": "ABIERTO",
                    "Fuente": "Pendientes abiertos anterior",
                })

    # Formato viejo: reconstruir desde hojas detalle
    if not pendings:
        # Flexxus no Banco
        if "Flexxus no Banco" in xl.sheet_names:
            raw = pd.read_excel(file, sheet_name="Flexxus no Banco", header=None, dtype=str)
            h = _find_header_row(raw, ["Fecha Flexxus", "Movimiento", "Monto"])
            if h is not None:
                fnb = pd.read_excel(file, sheet_name="Flexxus no Banco", header=h, dtype=str)
                for _, r in fnb.iterrows():
                    concepto = str(r.get("Movimiento", "")).strip()
                    monto = parse_ar_num(r.get("Monto", 0))
                    tipo = str(r.get("Tipo", "")).strip()
                    if not concepto or abs(monto) <= 0.005:
                        continue
                    signo = -1 if tipo == "PAV" else 1
                    pendings.append({
                        "pending_id": f"P-{uuid.uuid4().hex[:10]}",
                        "FechaOrigen": str(r.get("Fecha Flexxus", "")).strip(),
                        "Origen": "FLEXXUS_NO_BANCO",
                        "TipoPendiente": "FLEXXUS_NO_BANCO",
                        "TipoMovimiento": tipo,
                        "Numero": str(r.get("Número", "")).strip(),
                        "Concepto": concepto,
                        "Categoria": "PAV" if tipo == "PAV" else "EGRESO_FLEXXUS",
                        "Monto": abs(monto),
                        "SignoCalculo": signo,
                        "Estado": "ABIERTO",
                        "Fuente": "Flexxus no Banco anterior",
                    })
        # Banco ingresos no Flexxus
        if "Banco ingresos no Flexxus" in xl.sheet_names:
            raw = pd.read_excel(file, sheet_name="Banco ingresos no Flexxus", header=None, dtype=str)
            h = _find_header_row(raw, ["Fecha banco", "Concepto banco", "Importe"])
            if h is not None:
                bi = pd.read_excel(file, sheet_name="Banco ingresos no Flexxus", header=h, dtype=str)
                for _, r in bi.iterrows():
                    concepto = str(r.get("Concepto banco", "")).strip()
                    monto = parse_ar_num(r.get("Importe", 0))
                    if not concepto or abs(monto) <= 0.005:
                        continue
                    categoria = str(r.get("Categoría", r.get("Categoria", ""))).strip() or classify_bank(concepto)
                    pendings.append({
                        "pending_id": f"P-{uuid.uuid4().hex[:10]}",
                        "FechaOrigen": str(r.get("Fecha banco", "")).strip(),
                        "Origen": "BANCO_NO_FLEXXUS",
                        "TipoPendiente": "BANCO_INGRESO_NO_FLEXXUS",
                        "TipoMovimiento": "PAV",
                        "Numero": str(r.get("Comprobante banco", "")).strip(),
                        "Concepto": concepto,
                        "Categoria": categoria,
                        "Monto": abs(monto),
                        "SignoCalculo": 1,
                        "Estado": "ABIERTO",
                        "Fuente": "Banco ingresos no Flexxus anterior",
                    })
        # Banco egresos no Flexxus
        if "Banco egresos no Flexxus" in xl.sheet_names:
            raw = pd.read_excel(file, sheet_name="Banco egresos no Flexxus", header=None, dtype=str)
            h = _find_header_row(raw, ["Fecha banco", "Concepto banco", "Importe"])
            if h is not None:
                be = pd.read_excel(file, sheet_name="Banco egresos no Flexxus", header=h, dtype=str)
                for _, r in be.iterrows():
                    concepto = str(r.get("Concepto banco", "")).strip()
                    monto = parse_ar_num(r.get("Importe", 0))
                    if not concepto or abs(monto) <= 0.005:
                        continue
                    categoria = classify_bank(concepto)
                    pendings.append({
                        "pending_id": f"P-{uuid.uuid4().hex[:10]}",
                        "FechaOrigen": str(r.get("Fecha banco", "")).strip(),
                        "Origen": "BANCO_NO_FLEXXUS",
                        "TipoPendiente": "BANCO_EGRESO_NO_FLEXXUS",
                        "TipoMovimiento": "MB-EXT" if categoria in IMP_CATS else "MB-ENT-EX",
                        "Numero": str(r.get("Comprobante banco", "")).strip(),
                        "Concepto": concepto,
                        "Categoria": categoria,
                        "Monto": abs(monto),
                        "SignoCalculo": -1,
                        "Estado": "ABIERTO",
                        "Fuente": "Banco egresos no Flexxus anterior",
                    })

    out = pd.DataFrame(pendings)
    if out.empty:
        out = pd.DataFrame(columns=[
            "pending_id", "FechaOrigen", "Origen", "TipoPendiente", "TipoMovimiento",
            "Numero", "Concepto", "Categoria", "Monto", "SignoCalculo", "Estado", "Fuente"
        ])
    return out, prev, mode


def normalize_pending_origin(origen: str, tipo_p: str) -> str:
    txt = norm_txt(origen + " " + tipo_p)
    if "FLEXXUS" in txt and "BANCO" in txt:
        return "FLEXXUS_NO_BANCO" if txt.index("FLEXXUS") < txt.index("BANCO") else "BANCO_NO_FLEXXUS"
    if "BANCO" in txt and "FLEXXUS" in txt:
        return "BANCO_NO_FLEXXUS"
    if "FLEXXUS" in txt:
        return "FLEXXUS_NO_BANCO"
    return "BANCO_NO_FLEXXUS"


def infer_pending_sign(origen: str, tipo_p: str, categoria: str, concepto: str, monto: float) -> int:
    txt = norm_txt(origen + " " + tipo_p + " " + categoria + " " + concepto)
    if "BANCO" in txt and "INGRES" in txt:
        return 1
    if "BANCO" in txt and ("EGRES" in txt or "DEBIT" in txt or "RETENC" in txt or "IMPUEST" in txt or "IVA" in txt):
        return -1
    if "FLEXXUS" in txt and "INGRES" in txt:
        return -1
    if "FLEXXUS" in txt and "EGRES" in txt:
        return 1
    if categoria in {"QR", "LIQUIDACION_TARJETA", "TRANSFERENCIA_ENTRANTE"}:
        return 1
    if categoria in IMP_CATS or categoria in {"TRANSFERENCIA_SALIENTE"}:
        return -1
    return 1 if monto >= 0 else -1


def classify_mbext(movimiento: str) -> str:
    """Clasifica una fila MB-EXT de Flexxus para matcheo agregado contra banco."""
    c = norm_txt(movimiento)
    if "INGRESOS BRUTOS" in c or "IIBB" in c or "I.B." in c:
        return "RETENCION_IIBB"
    if "25413" in c and ("DEB" in c or "DEBIT" in c):
        return "IMPUESTO_LEY25413_DEB"
    if "25413" in c and ("CRED" in c or "CREDIT" in c):
        return "IMPUESTO_LEY25413_CRED"
    if "I.V.A" in c or "IVA" in c:
        return "IVA"
    if "COM TRANSFE" in c or "COM. TRANSFE" in c or "COMISION" in c or "COMISIÓN" in c:
        return "GASTO_BANCARIO"
    if "DEB LIQ" in c or "DEB.LIQ" in c or "LIQ MASTER" in c or "MASTERCARD" in c:
        return "DEBITO_LIQ_TARJETA"
    if "PAGO DIRECTO" in c:
        return "DEBITO_PAGO_DIRECTO"
    return "OTRO_MBEXT"


def amount_close(a: float, b: float, base_tol: float = 2.0, rel_tol: float = 0.00002) -> bool:
    """Tolerancia mixta: mínimo fijo + margen proporcional para grandes montos."""
    a = abs(float(a)); b = abs(float(b))
    return abs(a - b) <= max(base_tol, max(a, b) * rel_tol)


def get_saldo_banco_final(bank: pd.DataFrame) -> float:
    """Obtiene saldo final del extracto sin depender de bank.iloc[0]."""
    if bank.empty:
        return 0.0
    ordered = bank.sort_values("SourceOrder") if "SourceOrder" in bank.columns else bank.copy()
    first_date = ordered.iloc[0]["Fecha_dt"]
    last_date = ordered.iloc[-1]["Fecha_dt"]
    if pd.notna(first_date) and pd.notna(last_date) and first_date >= last_date:
        return float(ordered.iloc[0]["SaldoNum"])
    return float(ordered.iloc[-1]["SaldoNum"])


def get_saldo_banco_apertura(bank: pd.DataFrame) -> float:
    """Saldo inmediatamente anterior al primer movimiento cronológico del archivo."""
    if bank.empty:
        return 0.0
    tmp = bank.copy()
    tmp = tmp[pd.notna(tmp["Fecha_dt"])]
    if tmp.empty:
        return 0.0
    sort_cols = ["Fecha_dt"] + (["SourceOrder"] if "SourceOrder" in tmp.columns else ["row_id"])
    first = tmp.sort_values(sort_cols).iloc[0]
    return float(first["SaldoNum"]) - float(first["ImporteNum"])


def build_continuity_control(prev_summary: Dict[str, float], bank: pd.DataFrame) -> Dict:
    """Controla continuidad banco vs banco, no banco anterior vs Flexxus actual."""
    prev_bank = float(prev_summary.get("saldo_banco_anterior", 0) or 0)
    opening = get_saldo_banco_apertura(bank)
    if not prev_bank:
        return {"aplica": False, "mensaje": "Sin saldo banco anterior leído para validar continuidad."}
    diff = round(opening - prev_bank, 2)
    return {
        "aplica": True,
        "saldo_banco_cierre_anterior": prev_bank,
        "saldo_banco_apertura_actual": opening,
        "diferencia_continuidad_banco": diff,
        "ok": abs(diff) < 1.0,
        "mensaje": "OK: apertura bancaria actual coincide con cierre anterior." if abs(diff) < 1.0 else "Revisar: la apertura bancaria actual no coincide con el cierre anterior. Puede faltar un movimiento o el rango del extracto no es consecutivo.",
    }

# =============================================================================
# MOTOR DE MATCHING
# =============================================================================

def match_previous_pendings(prev_open: pd.DataFrame, flex: pd.DataFrame, bank: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Cancela pendientes anteriores contra movimientos actuales.

    V3 agrega matching agregado: una MB-EXT consolidada puede regularizar varias
    líneas pendientes anteriores de Banco no Flexxus de la misma categoría.
    """
    regs = []
    if prev_open.empty:
        return prev_open.copy(), flex, bank, pd.DataFrame(regs)

    prev = prev_open.copy().reset_index(drop=True)
    prev["Regularizado"] = False
    prev["RegularizadoCon"] = ""
    prev["FechaRegularizacion"] = ""
    prev["Diagnostico"] = "Pendiente anterior abierto"

    # 1) Matching agregado de MB-EXT actual contra múltiples pendientes bancarios anteriores.
    mbext_idx = flex[(~flex["Matched"]) & (~flex["ConsumedPrev"]) & (flex["Tipo"] == "MB-EXT")].index.tolist()
    for fidx in mbext_idx:
        monto = float(flex.loc[fidx, "MontoFlexxus"])
        cat = classify_mbext(flex.loc[fidx, "Movimiento"])
        if cat == "OTRO_MBEXT":
            continue
        candidates = prev[
            (~prev["Regularizado"]) &
            (prev["Origen"] == "BANCO_NO_FLEXXUS") &
            (prev["SignoCalculo"].astype(float) < 0) &
            (prev["Categoria"].astype(str).eq(cat))
        ].copy()
        if candidates.empty:
            continue
        total = float(candidates["Monto"].sum())
        if amount_close(total, monto):
            flex.loc[fidx, "ConsumedPrev"] = True
            flex.loc[fidx, "MatchStage"] = "REGULARIZACION_ANTERIOR_AGREGADA"
            flex.loc[fidx, "MatchRef"] = ",".join(candidates["pending_id"].astype(str).tolist())
            flex.loc[fidx, "Diagnostico"] = f"Regulariza {len(candidates)} pendientes bancarios anteriores categoría {cat} por suma agregada"
            flex.loc[fidx, "FechaAcreditacionUsada"] = flex.loc[fidx, "FechaFlexxus"]
            prev.loc[candidates.index, "Regularizado"] = True
            prev.loc[candidates.index, "RegularizadoCon"] = flex.loc[fidx, "row_id"]
            prev.loc[candidates.index, "FechaRegularizacion"] = flex.loc[fidx, "FechaFlexxus"]
            prev.loc[candidates.index, "Diagnostico"] = f"Regularizado por MB-EXT consolidada {flex.loc[fidx, 'Numero']}"
            regs.append({
                "ID pendiente": ",".join(candidates["pending_id"].astype(str).tolist()),
                "Fecha origen": "Varias",
                "Origen pendiente": "BANCO_EGRESO_NO_FLEXXUS_AGREGADO",
                "Concepto pendiente": f"{len(candidates)} líneas pendientes anteriores categoría {cat}",
                "Categoría pendiente": cat,
                "Monto pendiente": total,
                "Regularizado con": "Flexxus",
                "Fecha regularización": flex.loc[fidx, "FechaFlexxus"],
                "Tipo actual": "MB-EXT",
                "Referencia actual": flex.loc[fidx, "Numero"],
                "Concepto actual": flex.loc[fidx, "Movimiento"],
                "Monto actual": monto,
                "Diferencia": round(monto - total, 2),
                "Diagnóstico": f"Regularización agregada MB-EXT contra {len(candidates)} pendientes anteriores",
            })

    # 2) Matching 1:1 normal para lo que no fue agregado.
    prev = prev.sort_values(["Monto", "Categoria"], ascending=[False, True]).reset_index(drop=True)

    for pidx, p in prev.iterrows():
        if bool(p.get("Regularizado", False)):
            continue
        monto = float(p["Monto"])
        signo = int(p.get("SignoCalculo", 0))
        origen = str(p.get("Origen", ""))
        cat = str(p.get("Categoria", "OTRO"))

        if origen == "BANCO_NO_FLEXXUS":
            if signo > 0:
                candidates = flex[(~flex["Matched"]) & (~flex["ConsumedPrev"]) & (flex["Tipo"] == "PAV")].copy()
            else:
                candidates = flex[(~flex["Matched"]) & (~flex["ConsumedPrev"]) & (flex["Tipo"].isin(["MB-EXT", "MB-ENT-EX"]))].copy()
            candidates = candidates[candidates["MontoFlexxus"].apply(lambda x: amount_close(x, monto))]
            if not candidates.empty:
                candidates["date_diff"] = (candidates["FechaFlexxus_dt"] - safe_date(p.get("FechaOrigen", ""))).abs()
                best_idx = candidates.sort_values(["date_diff", "row_id"]).index[0]
                flex.loc[best_idx, "ConsumedPrev"] = True
                flex.loc[best_idx, "MatchStage"] = "REGULARIZACION_ANTERIOR"
                flex.loc[best_idx, "MatchRef"] = p["pending_id"]
                flex.loc[best_idx, "Diagnostico"] = "Regulariza Banco no Flexxus anterior"
                flex.loc[best_idx, "FechaAcreditacionUsada"] = flex.loc[best_idx, "FechaFlexxus"]
                prev.loc[pidx, "Regularizado"] = True
                prev.loc[pidx, "RegularizadoCon"] = flex.loc[best_idx, "row_id"]
                prev.loc[pidx, "FechaRegularizacion"] = flex.loc[best_idx, "FechaFlexxus"]
                prev.loc[pidx, "Diagnostico"] = "Regularizado con Flexxus actual"
                regs.append(build_reg_row(p, "Flexxus", flex.loc[best_idx].to_dict(), "Regulariza Banco no Flexxus anterior"))
                continue

        if origen == "FLEXXUS_NO_BANCO":
            if signo < 0:
                candidates = bank[(~bank["Matched"]) & (~bank["ConsumedPrev"]) & (bank["EsIngreso"])].copy()
            else:
                candidates = bank[(~bank["Matched"]) & (~bank["ConsumedPrev"]) & (bank["EsEgreso"])].copy()
            candidates = candidates[candidates["ImporteAbs"].apply(lambda x: amount_close(x, monto))]
            if not candidates.empty:
                comp = candidates[candidates["Categoria"].apply(lambda c: category_compatible(cat, c, origen))]
                if comp.empty:
                    comp = candidates
                comp["date_diff"] = (comp["Fecha_dt"] - safe_date(p.get("FechaOrigen", ""))).abs()
                best_idx = comp.sort_values(["date_diff", "row_id"]).index[0]
                bank.loc[best_idx, "ConsumedPrev"] = True
                bank.loc[best_idx, "MatchStage"] = "REGULARIZACION_ANTERIOR"
                bank.loc[best_idx, "MatchRef"] = p["pending_id"]
                bank.loc[best_idx, "Diagnostico"] = "Regulariza Flexxus no Banco anterior"
                prev.loc[pidx, "Regularizado"] = True
                prev.loc[pidx, "RegularizadoCon"] = bank.loc[best_idx, "row_id"]
                prev.loc[pidx, "FechaRegularizacion"] = bank.loc[best_idx, "Fecha"]
                prev.loc[pidx, "Diagnostico"] = "Regularizado con Banco actual"
                regs.append(build_reg_row(p, "Banco", bank.loc[best_idx].to_dict(), "Regulariza Flexxus no Banco anterior"))
                continue

    regs_df = pd.DataFrame(regs)
    return prev, flex, bank, regs_df

def build_reg_row(p, actual_source: str, actual: dict, diag: str) -> dict:
    if actual_source == "Flexxus":
        actual_fecha = actual.get("FechaFlexxus", "")
        actual_tipo = actual.get("Tipo", "")
        actual_concepto = actual.get("Movimiento", "")
        actual_monto = actual.get("MontoFlexxus", 0.0)
        actual_ref = actual.get("Numero", "")
    else:
        actual_fecha = actual.get("Fecha", "")
        actual_tipo = actual.get("Categoria", "")
        actual_concepto = actual.get("Concepto", "")
        actual_monto = actual.get("ImporteAbs", 0.0)
        actual_ref = actual.get("Comprobante", "")
    return {
        "ID pendiente": p.get("pending_id", ""),
        "Fecha origen": p.get("FechaOrigen", ""),
        "Origen pendiente": p.get("TipoPendiente", p.get("Origen", "")),
        "Concepto pendiente": p.get("Concepto", ""),
        "Categoría pendiente": p.get("Categoria", ""),
        "Monto pendiente": float(p.get("Monto", 0.0)),
        "Regularizado con": actual_source,
        "Fecha regularización": actual_fecha,
        "Tipo actual": actual_tipo,
        "Referencia actual": actual_ref,
        "Concepto actual": actual_concepto,
        "Monto actual": float(actual_monto),
        "Diferencia": round(float(actual_monto) - float(p.get("Monto", 0.0)), 2),
        "Diagnóstico": diag,
    }


def match_mbext_current_aggregated(flex: pd.DataFrame, bank: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Matchea MB-EXT corrientes consolidadas contra N líneas bancarias de la misma categoría.

    Esto resuelve el caso real Dancona: Flexxus consolida semanalmente IIBB/IVA/Ley 25413,
    pero Banco Nación lo muestra en muchas líneas diarias.
    """
    rows = []
    for fidx in flex[(~flex["Matched"]) & (~flex["ConsumedPrev"]) & (flex["Tipo"] == "MB-EXT")].index.tolist():
        monto = float(flex.loc[fidx, "MontoFlexxus"])
        cat = classify_mbext(flex.loc[fidx, "Movimiento"])
        if cat == "OTRO_MBEXT":
            continue
        candidates = bank[
            (~bank["Matched"]) & (~bank["ConsumedPrev"]) &
            (bank["EsEgreso"]) & (bank["Categoria"] == cat)
        ].copy()
        if candidates.empty:
            continue

        fecha_f = flex.loc[fidx, "FechaFlexxus_dt"]
        if pd.notna(fecha_f):
            win = candidates[(candidates["Fecha_dt"] <= fecha_f + pd.Timedelta(days=1)) & (candidates["Fecha_dt"] >= fecha_f - pd.Timedelta(days=21))]
        else:
            win = candidates
        if win.empty:
            win = candidates

        total_win = float(win["ImporteAbs"].sum())
        used = win
        if not amount_close(total_win, monto):
            total_all = float(candidates["ImporteAbs"].sum())
            if amount_close(total_all, monto):
                used = candidates
                total_win = total_all
            else:
                continue

        flex.loc[fidx, "Matched"] = True
        flex.loc[fidx, "MatchStage"] = "CORRIENTE_AGREGADO_MBEXT"
        flex.loc[fidx, "MatchRef"] = ",".join(used["row_id"].astype(str).tolist())
        flex.loc[fidx, "BancoFecha"] = fmt_date(used["Fecha_dt"].max())
        flex.loc[fidx, "BancoConcepto"] = f"Consolidado banco {cat}: {len(used)} líneas"
        flex.loc[fidx, "BancoComprobante"] = "AGREGADO"
        flex.loc[fidx, "FechaAcreditacionUsada"] = flex.loc[fidx, "FechaFlexxus"]
        flex.loc[fidx, "Diagnostico"] = f"OK - MB-EXT agregado contra {len(used)} líneas bancarias categoría {cat}"

        bank.loc[used.index, "Matched"] = True
        bank.loc[used.index, "MatchStage"] = "CORRIENTE_AGREGADO_MBEXT"
        bank.loc[used.index, "MatchRef"] = flex.loc[fidx, "row_id"]
        bank.loc[used.index, "FlexxusNumero"] = flex.loc[fidx, "Numero"]
        bank.loc[used.index, "Diagnostico"] = f"Incluido en MB-EXT agregado {flex.loc[fidx, 'Numero']}"
        rows.append({
            "Fecha Flexxus": flex.loc[fidx, "FechaFlexxus"],
            "Número Flexxus": flex.loc[fidx, "Numero"],
            "Categoría": cat,
            "Concepto Flexxus": flex.loc[fidx, "Movimiento"],
            "Monto Flexxus": monto,
            "Cantidad líneas banco": len(used),
            "Total banco": total_win,
            "Diferencia": round(monto - total_win, 2),
            "Diagnóstico": "OK agregado MB-EXT",
        })
    return flex, bank, pd.DataFrame(rows)


def match_current(flex: pd.DataFrame, bank: pd.DataFrame, qr: pd.DataFrame, trx: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Matchea movimientos corrientes Flexxus vs banco, sin tocar regularizaciones anteriores."""
    flex, bank, mbext_ag = match_mbext_current_aggregated(flex, bank)
    flex.attrs["mbext_agregado"] = mbext_ag
    # Lookups auxiliares
    qr_lookup = {}
    if not qr.empty and "NetoQR" in qr.columns:
        for _, r in qr.iterrows():
            qr_lookup.setdefault(round(float(r.get("NetoQR", 0)), 2), []).append(r)
    trx_lookup = {}
    if not trx.empty and "MontoNeto" in trx.columns:
        for _, r in trx.iterrows():
            trx_lookup.setdefault(round(float(r.get("MontoNeto", 0)), 2), []).append(r)

    def do_match(tipo: str, bank_side: str):
        flex_idx = flex[(~flex["Matched"]) & (~flex["ConsumedPrev"]) & (flex["Tipo"] == tipo)].index.tolist()
        for idx in flex_idx:
            monto = float(flex.loc[idx, "MontoFlexxus"])
            fecha_f = flex.loc[idx, "FechaFlexxus_dt"]
            if bank_side == "INGRESO":
                candidates = bank[(~bank["Matched"]) & (~bank["ConsumedPrev"]) & bank["EsIngreso"]].copy()
            else:
                candidates = bank[(~bank["Matched"]) & (~bank["ConsumedPrev"]) & bank["EsEgreso"]].copy()
            candidates = candidates[candidates["ImporteAbs"].apply(lambda x: amount_close(x, monto))]
            if candidates.empty:
                continue
            candidates["date_diff"] = (candidates["Fecha_dt"] - fecha_f).abs()
            best = candidates.sort_values(["date_diff", "row_id"]).iloc[0]
            bidx = best.name

            flex.loc[idx, "Matched"] = True
            flex.loc[idx, "MatchStage"] = "CORRIENTE"
            flex.loc[idx, "MatchRef"] = bank.loc[bidx, "row_id"]
            flex.loc[idx, "BancoFecha"] = bank.loc[bidx, "Fecha"]
            flex.loc[idx, "BancoConcepto"] = bank.loc[bidx, "Concepto"]
            flex.loc[idx, "BancoComprobante"] = bank.loc[bidx, "Comprobante"]
            flex.loc[idx, "Diagnostico"] = "OK corriente - match exacto por importe"
            # Regla fecha acreditación
            if bank.loc[bidx, "Fecha_dt"] < fecha_f:
                flex.loc[idx, "FechaAcreditacionUsada"] = flex.loc[idx, "FechaFlexxus"]
            else:
                flex.loc[idx, "FechaAcreditacionUsada"] = bank.loc[bidx, "Fecha"]

            bank.loc[bidx, "Matched"] = True
            bank.loc[bidx, "MatchStage"] = "CORRIENTE"
            bank.loc[bidx, "MatchRef"] = flex.loc[idx, "row_id"]
            bank.loc[bidx, "FlexxusNumero"] = flex.loc[idx, "Numero"]
            bank.loc[bidx, "Diagnostico"] = "OK corriente - match exacto por importe"

    do_match("PAV", "INGRESO")
    do_match("MB-ENT-EX", "EGRESO")
    do_match("MB-EXT", "EGRESO")

    return flex, bank

# =============================================================================
# RESULTADOS
# =============================================================================

def get_saldo_flexxus(flex: pd.DataFrame) -> float:
    # Mantiene criterio del sistema anterior: saldo de la serie PAV si existe; si no, último saldo general.
    pav = flex[(flex["Tipo"] == "PAV") & (flex["SaldoFlexxus"].abs() > 0)]
    if not pav.empty:
        return float(pav.sort_values("FechaFlexxus_dt").iloc[-1]["SaldoFlexxus"])
    nonzero = flex[flex["SaldoFlexxus"].abs() > 0]
    if not nonzero.empty:
        return float(nonzero.sort_values("FechaFlexxus_dt").iloc[-1]["SaldoFlexxus"])
    return 0.0


def compute_results(flex: pd.DataFrame, bank: pd.DataFrame, prev_status: pd.DataFrame, regs: pd.DataFrame, mode: str, prev_summary: Optional[Dict[str, float]] = None) -> Dict:
    saldo_f = get_saldo_flexxus(flex)
    saldo_b = get_saldo_banco_final(bank)
    continuity = build_continuity_control(prev_summary or {}, bank)

    current_unmatched_f = flex[(~flex["Matched"]) & (~flex["ConsumedPrev"])]
    current_unmatched_b = bank[(~bank["Matched"]) & (~bank["ConsumedPrev"])]

    C1 = current_unmatched_f[current_unmatched_f["Tipo"] == "PAV"]["MontoFlexxus"].sum()
    C2 = current_unmatched_f[current_unmatched_f["Tipo"].isin(["MB-ENT-EX", "MB-EXT"])]["MontoFlexxus"].sum()
    C3 = current_unmatched_b[current_unmatched_b["EsIngreso"]]["ImporteAbs"].sum()
    C4 = current_unmatched_b[current_unmatched_b["EsEgreso"]]["ImporteAbs"].sum()

    if prev_status is None or prev_status.empty:
        prev_open = pd.DataFrame()
        prev_open_effect = 0.0
    else:
        prev_open = prev_status[~prev_status.get("Regularizado", False)].copy()
        prev_open_effect = float((prev_open["Monto"] * prev_open["SignoCalculo"]).sum()) if not prev_open.empty else 0.0

    calc = saldo_f - C1 + C2 + C3 - C4 + prev_open_effect
    diff = round(saldo_b - calc, 2)

    # Generar pendientes abiertos para próxima conciliación: anteriores no regularizados + nuevos corrientes.
    next_pendings = []
    if not prev_open.empty:
        for _, p in prev_open.iterrows():
            d = p.to_dict()
            d["Estado"] = "ABIERTO_ANTERIOR"
            d["Fuente"] = d.get("Fuente", "") or "Pendiente anterior aún abierto"
            next_pendings.append(d)

    for _, f in current_unmatched_f.iterrows():
        sign = -1 if f["Tipo"] == "PAV" else 1
        next_pendings.append({
            "pending_id": f"P-{uuid.uuid4().hex[:10]}",
            "FechaOrigen": f["FechaFlexxus"],
            "Origen": "FLEXXUS_NO_BANCO",
            "TipoPendiente": "FLEXXUS_INGRESO_NO_BANCO" if f["Tipo"] == "PAV" else "FLEXXUS_EGRESO_NO_BANCO",
            "TipoMovimiento": f["Tipo"],
            "Numero": f["Numero"],
            "Concepto": f["Movimiento"],
            "Categoria": "PAV" if f["Tipo"] == "PAV" else f["Tipo"],
            "Monto": float(f["MontoFlexxus"]),
            "SignoCalculo": sign,
            "Estado": "ABIERTO_NUEVO",
            "Fuente": "Corriente actual",
        })
    for _, b in current_unmatched_b.iterrows():
        sign = 1 if b["EsIngreso"] else -1
        next_pendings.append({
            "pending_id": f"P-{uuid.uuid4().hex[:10]}",
            "FechaOrigen": b["Fecha"],
            "Origen": "BANCO_NO_FLEXXUS",
            "TipoPendiente": "BANCO_INGRESO_NO_FLEXXUS" if b["EsIngreso"] else "BANCO_EGRESO_NO_FLEXXUS",
            "TipoMovimiento": "PAV" if b["EsIngreso"] else ("MB-EXT" if b["Categoria"] in IMP_CATS else "MB-ENT-EX"),
            "Numero": b["Comprobante"],
            "Concepto": b["Concepto"],
            "Categoria": b["Categoria"],
            "Monto": float(b["ImporteAbs"]),
            "SignoCalculo": sign,
            "Estado": "ABIERTO_NUEVO",
            "Fuente": "Corriente actual",
        })

    pend_next = pd.DataFrame(next_pendings)

    return {
        "mode": mode,
        "saldo_flexxus": saldo_f,
        "saldo_banco": saldo_b,
        "C1": float(C1),
        "C2": float(C2),
        "C3": float(C3),
        "C4": float(C4),
        "prev_open_effect": float(prev_open_effect),
        "calc_final": float(calc),
        "diferencia": float(diff),
        "current_unmatched_f": current_unmatched_f.copy(),
        "current_unmatched_b": current_unmatched_b.copy(),
        "prev_open": prev_open.copy() if not prev_open.empty else pd.DataFrame(),
        "regularizaciones": regs.copy() if regs is not None else pd.DataFrame(),
        "mbext_agregado": flex.attrs.get("mbext_agregado", pd.DataFrame()),
        "continuity": continuity,
        "pendientes_proxima": pend_next,
        "matched_flex": flex[flex["Matched"]].copy(),
        "matched_bank": bank[bank["Matched"]].copy(),
        "consumed_prev_flex": flex[flex["ConsumedPrev"]].copy(),
        "consumed_prev_bank": bank[bank["ConsumedPrev"]].copy(),
    }

# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def write_df(ws, start_row: int, start_col: int, df: pd.DataFrame, money_cols: Optional[List[str]] = None, date_cols: Optional[List[str]] = None):
    money_cols = money_cols or []
    date_cols = date_cols or []
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for j, col in enumerate(df.columns, start_col):
        c = ws.cell(start_row, j, col)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    for i, (_, row) in enumerate(df.iterrows(), start_row + 1):
        for j, col in enumerate(df.columns, start_col):
            val = row[col]
            if isinstance(val, (np.integer, np.floating)):
                val = float(val)
            c = ws.cell(i, j, val)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if col in money_cols:
                c.number_format = '#,##0.00'
            if col in date_cols:
                c.number_format = "dd/mm/yyyy"
    for col_idx in range(start_col, start_col + len(df.columns)):
        letter = get_column_letter(col_idx)
        max_len = 10
        for cell in ws[letter]:
            max_len = max(max_len, len(str(cell.value or "")))
        ws.column_dimensions[letter].width = min(max_len + 2, 42)
    ws.freeze_panes = ws.cell(start_row + 1, start_col).coordinate
    try:
        ws.auto_filter.ref = ws.dimensions
    except Exception:
        pass


def build_excel_report(flex: pd.DataFrame, bank: pd.DataFrame, qr: pd.DataFrame, trx: pd.DataFrame, res: Dict) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen ejecutivo"

    title_font = Font(bold=True, size=15, color="1F4E78")
    sub_font = Font(bold=True, size=11)
    money_fmt = '#,##0.00'
    green_fill = PatternFill("solid", fgColor="E2F0D9")
    red_fill = PatternFill("solid", fgColor="FCE4D6")

    ws["A1"] = "DANCONA ALIMENTOS - CONCILIACIÓN BANCARIA CONTINUA V3"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")
    periodo = ""
    if not bank.empty:
        periodo = f"Período banco {fmt_date(bank['Fecha_dt'].min())} al {fmt_date(bank['Fecha_dt'].max())}"
    ws["A2"] = periodo
    ws["A3"] = "Modo: " + ("CONCILIACIÓN DE APERTURA (sin archivo anterior)" if res["mode"] == "APERTURA" else "CONCILIACIÓN CONTINUA CON ARCHIVO ANTERIOR")

    summary = [
        ("Saldo S/Flexxus actual", res["saldo_flexxus"]),
        ("Menos C1 corriente: ingresos Flexxus no Banco", -res["C1"]),
        ("Más C2 corriente: egresos Flexxus no Banco", res["C2"]),
        ("Más C3 corriente: ingresos Banco no Flexxus", res["C3"]),
        ("Menos C4 corriente: egresos Banco no Flexxus", -res["C4"]),
        ("Efecto de pendientes anteriores aún abiertos", res["prev_open_effect"]),
        ("SALDO BANCO CALCULADO", res["calc_final"]),
        ("SALDO SEGÚN EXTRACTO BANCO", res["saldo_banco"]),
        ("DIFERENCIA FINAL", res["diferencia"]),
    ]
    ws["A5"] = "Concepto"
    ws["B5"] = "Importe"
    ws["A5"].font = sub_font
    ws["B5"].font = sub_font
    for i, (label, val) in enumerate(summary, 6):
        ws.cell(i, 1, label)
        c = ws.cell(i, 2, float(val))
        c.number_format = money_fmt
        if "DIFERENCIA FINAL" in label:
            c.fill = green_fill if abs(float(val)) < 1.0 else red_fill
            c.font = Font(bold=True, color="006100" if abs(float(val)) < 1.0 else "9C0006")
        if "SALDO" in label:
            ws.cell(i, 1).font = sub_font
            c.font = sub_font

    ws["D5"] = "Control auditoría"
    ws["D5"].font = sub_font
    controls = [
        ("Movimientos corrientes matcheados", len(res["matched_flex"])),
        ("Regularizaciones anteriores detectadas", len(res["regularizaciones"])),
        ("Pendientes anteriores aún abiertos", len(res["prev_open"])),
        ("Pendientes para próxima conciliación", len(res["pendientes_proxima"])),
        ("MB-EXT agregados", len(res.get("mbext_agregado", pd.DataFrame()))),
    ]
    for i, (label, val) in enumerate(controls, 6):
        ws.cell(i, 4, label)
        ws.cell(i, 5, val)

    ws["A17"] = "Regla: no se usa ajuste heredado genérico. Todo arrastre queda abierto movimiento por movimiento."
    ws["A17"].font = Font(italic=True, color="666666")
    ws.merge_cells("A17:E17")
    for col in ["A", "B", "D", "E"]:
        ws.column_dimensions[col].width = 34 if col in {"A", "D"} else 18

    # Regularizaciones
    ws_reg = wb.create_sheet("Regularizaciones anteriores")
    regs = res["regularizaciones"]
    if regs.empty:
        regs = pd.DataFrame([{"Diagnóstico": "Sin regularizaciones de períodos anteriores detectadas"}])
    write_df(ws_reg, 1, 1, regs, money_cols=["Monto pendiente", "Monto actual", "Diferencia"])

    # MB-EXT agregado
    ws_ag = wb.create_sheet("MB-EXT agregado")
    ag = res.get("mbext_agregado", pd.DataFrame())
    if ag.empty:
        ag = pd.DataFrame([{"Diagnóstico": "Sin MB-EXT corrientes agregados contra banco"}])
    write_df(ws_ag, 1, 1, ag, money_cols=["Monto Flexxus", "Total banco", "Diferencia"])

    # Pendientes abiertos
    ws_pa = wb.create_sheet("Pendientes abiertos")
    pend = res["pendientes_proxima"]
    if pend.empty:
        pend = pd.DataFrame([{"Estado": "SIN PENDIENTES", "Concepto": "No quedan pendientes abiertos", "Monto": 0.0}])
    cols = [c for c in ["pending_id", "FechaOrigen", "Origen", "TipoPendiente", "TipoMovimiento", "Numero", "Concepto", "Categoria", "Monto", "SignoCalculo", "Estado", "Fuente"] if c in pend.columns]
    write_df(ws_pa, 1, 1, pend[cols], money_cols=["Monto"])

    # Corriente Flexxus no Banco
    ws_fnb = wb.create_sheet("Flexxus no Banco corriente")
    fnb = res["current_unmatched_f"].copy()
    if fnb.empty:
        fnb = pd.DataFrame([{"Diagnóstico": "Sin movimientos corrientes Flexxus no Banco"}])
    else:
        fnb = fnb[["FechaFlexxus", "Tipo", "Numero", "Movimiento", "MontoFlexxus", "Diagnostico"]]
    write_df(ws_fnb, 1, 1, fnb, money_cols=["MontoFlexxus"])

    # Banco no Flexxus corriente
    ws_bnf = wb.create_sheet("Banco no Flexxus corriente")
    bnf = res["current_unmatched_b"].copy()
    if bnf.empty:
        bnf = pd.DataFrame([{"Diagnóstico": "Sin movimientos corrientes Banco no Flexxus"}])
    else:
        bnf = bnf[["Fecha", "Comprobante", "Concepto", "Categoria", "ImporteNum", "SaldoNum", "Diagnostico"]]
    write_df(ws_bnf, 1, 1, bnf, money_cols=["ImporteNum", "SaldoNum"])

    # Impuestos y gastos
    ws_imp = wb.create_sheet("Impuestos y gastos")
    imp_rows = []
    # Corrientes banco no Flexxus
    cur_b = res["current_unmatched_b"]
    if not cur_b.empty:
        for _, b in cur_b[cur_b["Categoria"].isin(IMP_CATS)].iterrows():
            imp_rows.append({
                "Estado": "Pendiente nuevo corriente",
                "Fecha": b["Fecha"],
                "Categoría": CAT_LABELS.get(b["Categoria"], b["Categoria"]),
                "Concepto": b["Concepto"],
                "Importe": abs(float(b["ImporteNum"])),
            })
    # Regularizados
    if not res["regularizaciones"].empty:
        for _, r in res["regularizaciones"].iterrows():
            cat = str(r.get("Categoría pendiente", ""))
            if cat in IMP_CATS or any(k in norm_txt(r.get("Concepto pendiente", "")) for k in ["IIBB", "INGRESOS BRUTOS", "GRAVAMEN", "I.V.A", "IVA", "COM TRANSFE"]):
                imp_rows.append({
                    "Estado": "Regularizado desde período anterior",
                    "Fecha": r.get("Fecha regularización", ""),
                    "Categoría": CAT_LABELS.get(cat, cat),
                    "Concepto": r.get("Concepto pendiente", ""),
                    "Importe": float(r.get("Monto pendiente", 0.0)),
                })
    imp_df = pd.DataFrame(imp_rows) if imp_rows else pd.DataFrame([{"Estado": "Sin impuestos/gastos pendientes ni regularizados", "Importe": 0.0}])
    write_df(ws_imp, 1, 1, imp_df, money_cols=["Importe"])

    # Carga Flexxus: solo matcheados corrientes, no regularizaciones anteriores ni pendientes.
    ws_carga = wb.create_sheet("Carga Flexxus")
    matched = res["matched_flex"].copy()
    if matched.empty:
        carga = pd.DataFrame([{"Aviso": "No hay movimientos corrientes para importar"}])
    else:
        carga = matched[["FechaFlexxus", "Tipo", "Numero", "MontoFlexxus", "BancoFecha", "FechaAcreditacionUsada", "BancoConcepto", "BancoComprobante"]].copy()
    write_df(ws_carga, 1, 1, carga, money_cols=["MontoFlexxus"])

    # Auditorías QR/TRX
    ws_qr = wb.create_sheet("Auditoría QR PCT")
    if qr.empty:
        qr_out = pd.DataFrame([{"Aviso": "No se cargó archivo QR"}])
    else:
        qr_out = qr.copy()
        keep = [c for c in ["FechaQR", "Estado", "CodComercio", "Cupon", "MontoTotal", "NetoQR", "Id QR"] if c in qr_out.columns]
        qr_out = qr_out[keep].head(1000)
    write_df(ws_qr, 1, 1, qr_out, money_cols=["MontoTotal", "NetoQR"])

    ws_trx = wb.create_sheet("Auditoría TRX Merchant")
    if trx.empty:
        trx_out = pd.DataFrame([{"Aviso": "No se cargó archivo TRX"}])
    else:
        keep = [c for c in ["FECHA DE PAGO", "NUMERO LIQUIDACION", "COMERCIO", "Local", "TARJETA", "MontoNeto"] if c in trx.columns]
        trx_out = trx[keep].copy().head(1000)
    write_df(ws_trx, 1, 1, trx_out, money_cols=["MontoNeto"])

    # Notas proceso
    ws_notas = wb.create_sheet("Notas proceso")
    notas = pd.DataFrame([
        {"Tema": "Principio", "Detalle": "La conciliación anterior es obligatoria desde la segunda corrida. Si no se adjunta, el sistema trabaja en modo apertura."},
        {"Tema": "Prohibición", "Detalle": "No se permite cerrar con ajuste heredado genérico. Todo saldo heredado debe abrirse movimiento por movimiento."},
        {"Tema": "Regularizaciones", "Detalle": "Un MB-EXT actual puede cancelar egresos bancarios anteriores por IIBB, IVA, Ley 25.413, comisiones o débitos de liquidación."},
        {"Tema": "MB-EXT agregado", "Detalle": "Si Flexxus consolida impuestos/gastos y Banco los muestra en N líneas, el sistema suma las líneas bancarias por categoría y las matchea contra la MB-EXT."},
        {"Tema": "Carga Flexxus", "Detalle": "El archivo de importación incluye solo movimientos corrientes matcheados contra banco actual. Las regularizaciones anteriores se auditan aparte."},
        {"Tema": "QR", "Detalle": "Neto QR = Monto total × 0,99032."},
        {"Tema": "Pedidos Ya", "Detalle": "Pedidos Ya no se trata como QR ni tarjeta; no se fuerza cupón ni liquidación."},
    ])
    write_df(ws_notas, 1, 1, notas)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def build_import_xls(res: Dict) -> io.BytesIO:
    wb = xlwt.Workbook()
    ws = wb.add_sheet("ASIENTOS")
    header_style = xlwt.easyxf("font: bold on; align: horiz center")
    date_style = xlwt.easyxf(num_format_str="DD/MM/YYYY")
    money_style = xlwt.easyxf(num_format_str="#,##0.00")
    headers = ["FECHAMOVIMIENTO", "TIPOMOVIMIENTO", "MONTO", "FECHAACREDITACION"]
    for c, h in enumerate(headers):
        ws.write(0, c, h, header_style)
    matched = res["matched_flex"].copy()
    # Solo movimientos corrientes. No incluye ConsumedPrev.
    for i, (_, fr) in enumerate(matched.iterrows(), 1):
        fm = safe_date(fr["FechaFlexxus"])
        fa = safe_date(fr.get("FechaAcreditacionUsada", "") or fr.get("BancoFecha", "") or fr["FechaFlexxus"])
        if pd.notna(fm):
            ws.write(i, 0, fm.to_pydatetime(), date_style)
        else:
            ws.write(i, 0, str(fr["FechaFlexxus"]))
        ws.write(i, 1, str(fr["Tipo"]))
        ws.write(i, 2, float(abs(fr["MontoFlexxus"])), money_style)
        if pd.notna(fa):
            ws.write(i, 3, fa.to_pydatetime(), date_style)
        else:
            ws.write(i, 3, str(fr.get("FechaAcreditacionUsada", "")))
    for c in range(4):
        ws.col(c).width = 5000
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =============================================================================
# UI STREAMLIT V4 INTEGRADA
# =============================================================================

APP_VERSION = "V4-INTEGRADA-LOGIN-HISTORIAL-MBEXT-2026-04-29"
HISTORIAL_FILE = "historico.json"

def secret_get(key: str, default: str = "") -> str:
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

def render_header_v4():
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1F4E78,#0F243E);padding:28px;border-radius:14px;margin-bottom:20px;color:white">
      <div style="font-size:13px;opacity:.75;letter-spacing:.08em">GRUPO DANCONA · CONTROL BANCARIO</div>
      <div style="font-size:30px;font-weight:700">Conciliación Bancaria Continua V4</div>
      <div style="font-size:15px;opacity:.85;margin-top:6px">Login · Historial · Botón Comenzar · Pendientes anteriores · Matching agregado MB-EXT · Sin ajustes genéricos.</div>
      <div style="font-size:11px;opacity:.70;margin-top:10px;font-family:monospace">{APP_VERSION}</div>
    </div>
    """, unsafe_allow_html=True)

def login_gate():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return
    render_header_v4()
    st.subheader("🔐 Iniciar sesión")
    valid_user = secret_get("APP_USER", "dancona2016@gmail.com")
    valid_password = secret_get("APP_PASSWORD", "Dancona2026*")
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        usuario = st.text_input("Usuario", placeholder="email")
        password = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", type="primary", use_container_width=True):
            if usuario == valid_user and password == valid_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
    st.stop()

def reset_app_state():
    keep = {"authenticated"}
    for k in list(st.session_state.keys()):
        if k not in keep:
            del st.session_state[k]
    try:
        st.cache_data.clear()
        st.cache_resource.clear()
    except Exception:
        pass

def github_get_file(filename: str = HISTORIAL_FILE):
    token = secret_get("GITHUB_TOKEN", "")
    repo = secret_get("GITHUB_REPO", "")
    if not token or not repo or requests is None:
        return None, None, "GitHub no configurado o requests no disponible"
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200:
            data = r.json()
            content = json.loads(base64.b64decode(data["content"]).decode("utf-8"))
            return content, data.get("sha"), "OK"
        if r.status_code == 404:
            return {"semanas": []}, None, "Historial no existe todavía; se creará al guardar."
        return None, None, f"GitHub HTTP {r.status_code}: {r.text[:250]}"
    except Exception as e:
        return None, None, f"Error leyendo GitHub: {e}"

def github_save_file(content_dict: dict, sha=None, filename: str = HISTORIAL_FILE):
    token = secret_get("GITHUB_TOKEN", "")
    repo = secret_get("GITHUB_REPO", "")
    if not token or not repo or requests is None:
        return False, "GitHub no configurado. La conciliación se generó, pero no se guardó historial."
    url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    encoded = base64.b64encode(json.dumps(content_dict, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
    body = {"message": f"Update {filename} - conciliacion V4", "content": encoded}
    if sha:
        body["sha"] = sha
    try:
        r = requests.put(url, headers=headers, json=body, timeout=20)
        if r.status_code in (200, 201):
            return True, "Historial guardado en GitHub."
        return False, f"GitHub HTTP {r.status_code}: {r.text[:300]}"
    except Exception as e:
        return False, f"Error guardando GitHub: {e}"

def guardar_resumen_historial(res: Dict, mode: str):
    hist, sha, msg = github_get_file()
    if hist is None:
        return False, msg
    hist.setdefault("semanas", [])
    item = {
        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
        "fecha_proceso": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "version_app": APP_VERSION,
        "modo": mode,
        "saldo_flexxus": round(float(res.get("saldo_flexxus", 0)), 2),
        "saldo_banco": round(float(res.get("saldo_banco", 0)), 2),
        "banco_calculado": round(float(res.get("calc_final", 0)), 2),
        "diferencia": round(float(res.get("diferencia", 0)), 2),
        "C1_flexxus_ingresos_no_banco": round(float(res.get("C1", 0)), 2),
        "C2_flexxus_egresos_no_banco": round(float(res.get("C2", 0)), 2),
        "C3_banco_ingresos_no_flexxus": round(float(res.get("C3", 0)), 2),
        "C4_banco_egresos_no_flexxus": round(float(res.get("C4", 0)), 2),
        "pendientes_anteriores_abiertos": int(len(res.get("prev_open", pd.DataFrame()))),
        "regularizaciones_anteriores": int(len(res.get("regularizaciones", pd.DataFrame()))),
        "pendientes_proxima": int(len(res.get("pendientes_proxima", pd.DataFrame()))),
        "mbext_agregados": int(len(res.get("mbext_agregado", pd.DataFrame()))),
    }
    hist["semanas"].append(item)
    return github_save_file(hist, sha)

def render_historial_tab():
    st.subheader("📊 Historial")
    hist, sha, msg = github_get_file()
    if hist is None:
        st.warning(msg)
        st.info("La app puede conciliar igual. Para guardar historial configurá GITHUB_TOKEN y GITHUB_REPO en Secrets.")
        return
    semanas = hist.get("semanas", [])
    if not semanas:
        st.info("Todavía no hay conciliaciones guardadas en el historial.")
        st.caption(msg)
        return
    df = pd.DataFrame(semanas)
    if "fecha_proceso" in df.columns:
        df = df.sort_values("fecha_proceso", ascending=False)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Descargar historico.json", data=json.dumps(hist, ensure_ascii=False, indent=2).encode("utf-8"), file_name="historico.json", mime="application/json", use_container_width=True)

def render_diagnostico_tab():
    st.subheader("🧪 Diagnóstico técnico")
    st.code(APP_VERSION)
    st.write("Python")
    st.code(sys.version)
    versions = {
        "streamlit": getattr(st, "__version__", ""),
        "pandas": pd.__version__,
        "numpy": np.__version__,
        "requests": getattr(requests, "__version__", "NO DISPONIBLE") if requests is not None else "NO DISPONIBLE",
        "xlwt": getattr(xlwt, "__VERSION__", "instalado"),
    }
    for mod in ["openpyxl", "xlrd"]:
        try:
            m = __import__(mod)
            versions[mod] = getattr(m, "__version__", "instalado")
        except Exception as e:
            versions[mod] = f"ERROR: {e}"
    st.json(versions)
    st.write("Estado de sesión")
    st.json({k: str(v)[:120] for k, v in st.session_state.items() if k not in {"password"}})
    st.write("Secrets detectados")
    st.json({"APP_USER": bool(secret_get("APP_USER", "")), "APP_PASSWORD": bool(secret_get("APP_PASSWORD", "")), "GITHUB_TOKEN": bool(secret_get("GITHUB_TOKEN", "")), "GITHUB_REPO": secret_get("GITHUB_REPO", "")})

def render_conciliacion_tab():
    st.subheader("🏦 Conciliación")
    st.info("Subí Flexxus y Banco como mínimo. TRX, QR y conciliación anterior mejoran la auditoría. Desde la segunda semana, subí siempre la conciliación anterior para cancelar pendientes.")
    col_a, col_b = st.columns(2)
    with col_a:
        f_flex = st.file_uploader("1 · Conciliación Bancaria Flexxus actual", type=["xls", "xlsx"], key="f_flex_v4")
        f_trx = st.file_uploader("3 · TRX Merchant Center", type=["xlsx", "xls"], key="f_trx_v4")
        f_prev = st.file_uploader("5 · Conciliación anterior del sistema", type=["xlsx"], key="f_prev_v4")
    with col_b:
        f_bank = st.file_uploader("2 · Últimos movimientos Banco Nación actual", type=["xls", "xlsx"], key="f_bank_v4")
        f_qr = st.file_uploader("4 · Transacciones QR", type=["xlsx", "xls"], key="f_qr_v4")
    with st.expander("🧪 Diagnóstico de archivos cargados"):
        files = {"Flexxus": f_flex, "Banco": f_bank, "TRX": f_trx, "QR": f_qr, "Conciliación anterior": f_prev}
        st.dataframe(pd.DataFrame([{"Archivo": k, "Cargado": v is not None, "Nombre": getattr(v, "name", ""), "Tamaño bytes": getattr(v, "size", 0) if v is not None else 0} for k, v in files.items()]), use_container_width=True, hide_index=True)

    if not f_flex or not f_bank:
        st.warning("Subí al menos Flexxus actual y Banco Nación actual para habilitar el procesamiento.")
        return

    run = st.button("▶️ Comenzar conciliación", type="primary", use_container_width=True)
    if st.button("🧹 Limpiar resultado", use_container_width=True):
        for k in ["last_result_v4", "last_error_v4", "last_xlsx_v4", "last_xls_v4"]:
            st.session_state.pop(k, None)
        st.rerun()

    if run:
        try:
            with st.spinner("Leyendo archivos..."):
                flex = parse_flexxus(f_flex)
                bank = parse_banco(f_bank)
                trx = parse_trx(f_trx) if f_trx else pd.DataFrame()
                qr = parse_qr(f_qr) if f_qr else pd.DataFrame()
                prev_open, prev_summary, mode = parse_previous_conciliation(f_prev)
            with st.spinner("Cancelando pendientes anteriores contra movimientos actuales..."):
                prev_status, flex, bank, regs = match_previous_pendings(prev_open, flex, bank)
            with st.spinner("Conciliando movimientos corrientes..."):
                flex, bank = match_current(flex, bank, qr, trx)
                res = compute_results(flex, bank, prev_status, regs, mode, prev_summary)
            with st.spinner("Generando entregables..."):
                xlsx = build_excel_report(flex, bank, qr, trx, res)
                xls = build_import_xls(res)
            st.session_state.last_result_v4 = res
            st.session_state.last_xlsx_v4 = xlsx.getvalue()
            st.session_state.last_xls_v4 = xls.getvalue()
            st.session_state.last_error_v4 = None
            st.success("Conciliación procesada.")
        except Exception as e:
            err = traceback.format_exc()
            st.session_state.last_error_v4 = err
            st.exception(e)
            st.error("No se pudo procesar. Descargá el diagnóstico y pasámelo.")
            st.download_button("Descargar diagnóstico del error", data=err.encode("utf-8"), file_name="diagnostico_error_v4.txt", mime="text/plain")
            return

    if "last_result_v4" not in st.session_state:
        st.caption("Esperando que presiones **Comenzar conciliación**.")
        return

    res = st.session_state.last_result_v4
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Saldo Flexxus", f"${res['saldo_flexxus']:,.2f}")
    c2.metric("Saldo Banco", f"${res['saldo_banco']:,.2f}")
    c3.metric("Banco calculado", f"${res['calc_final']:,.2f}")
    c4.metric("Diferencia", f"${res['diferencia']:,.2f}")

    cont = res.get("continuity", {})
    if cont.get("aplica"):
        st.subheader("🔗 Trazabilidad con semana anterior")
        cc1, cc2, cc3 = st.columns(3)
        cc1.metric("Saldo banco cierre anterior", f"${cont.get('saldo_banco_cierre_anterior',0):,.2f}")
        cc2.metric("Apertura banco actual", f"${cont.get('saldo_banco_apertura_actual',0):,.2f}")
        cc3.metric("Diferencia continuidad banco", f"${cont.get('diferencia_continuidad_banco',0):,.2f}")
        if cont.get("ok"):
            st.success(cont.get("mensaje"))
        else:
            st.warning(cont.get("mensaje"))
    else:
        st.info(cont.get("mensaje", "Sin control de continuidad bancaria."))

    st.subheader("Prueba explícita")
    proof = pd.DataFrame([
        {"Concepto": "Saldo Flexxus actual", "Importe": res["saldo_flexxus"]},
        {"Concepto": "- C1 corriente: ingresos Flexxus no Banco", "Importe": -res["C1"]},
        {"Concepto": "+ C2 corriente: egresos Flexxus no Banco", "Importe": res["C2"]},
        {"Concepto": "+ C3 corriente: ingresos Banco no Flexxus", "Importe": res["C3"]},
        {"Concepto": "- C4 corriente: egresos Banco no Flexxus", "Importe": -res["C4"]},
        {"Concepto": "+/- pendientes anteriores aún abiertos", "Importe": res["prev_open_effect"]},
        {"Concepto": "Banco calculado", "Importe": res["calc_final"]},
        {"Concepto": "Banco real", "Importe": res["saldo_banco"]},
        {"Concepto": "Diferencia final", "Importe": res["diferencia"]},
    ])
    st.dataframe(proof, use_container_width=True, hide_index=True)
    if abs(res["diferencia"]) >= 1:
        st.error("La diferencia final no da 0. El sistema NO la forzó. Revisá pendientes abiertos y movimientos no conciliados.")
    else:
        st.success("Diferencia final conciliada en 0,00 o dentro de tolerancia de redondeo.")

    st.subheader("Regularizaciones de períodos anteriores")
    if res["regularizaciones"].empty:
        st.write("Sin regularizaciones anteriores detectadas.")
    else:
        st.dataframe(res["regularizaciones"], use_container_width=True, hide_index=True)

    st.subheader("MB-EXT agregados contra banco")
    ag = res.get("mbext_agregado", pd.DataFrame())
    if ag.empty:
        st.write("Sin MB-EXT agregados en esta corrida.")
    else:
        st.dataframe(ag, use_container_width=True, hide_index=True)

    st.subheader("Pendientes abiertos para próxima conciliación")
    pend = res["pendientes_proxima"]
    st.dataframe(pend if not pend.empty else pd.DataFrame([{"Estado": "Sin pendientes abiertos"}]), use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button("Descargar Excel de conciliación V4", data=st.session_state.last_xlsx_v4, file_name="Conciliacion_Semanal_Dancona_V4.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with d2:
        st.download_button("Descargar .xls para importar a Flexxus", data=st.session_state.last_xls_v4, file_name="ASIENTOS_FLEXXUS_V4.xls", mime="application/vnd.ms-excel", use_container_width=True)

    if st.button("💾 Guardar resumen en historial", use_container_width=True):
        ok, msg = guardar_resumen_historial(res, res.get("mode", ""))
        if ok:
            st.success(msg)
        else:
            st.warning(msg)

def main():
    login_gate()
    with st.sidebar:
        st.markdown("### Sesión")
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.authenticated = False
            reset_app_state()
            st.rerun()
        if st.button("🔄 Resetear app / limpiar estado", use_container_width=True):
            reset_app_state()
            st.rerun()
        st.caption(APP_VERSION)
    render_header_v4()
    tab1, tab2, tab3 = st.tabs(["🏦 Conciliación", "📊 Historial", "🧪 Diagnóstico"])
    with tab1:
        render_conciliacion_tab()
    with tab2:
        render_historial_tab()
    with tab3:
        render_diagnostico_tab()

if __name__ == "__main__":
    main()
