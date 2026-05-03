"""
qr_ledger_regla3_v2.py - Regla 3 sobre PENDIENTES finales.

A diferencia de v1 que operaba sobre bank_df antes del motor (con riesgo de
falsos positivos), esta versión opera sobre `res['pendientes_proxima']`
después de que el motor V5.9.11 hizo todo lo posible.

Filosofía: la regla 3 es la última red. Solo bloquea pendientes que
quedaron sin regularizar y tienen un cupón terminal histórico equivalente.
"""

from __future__ import annotations

from datetime import datetime
from typing import Dict, List, Set, Tuple

import pandas as pd

from qr_ledger import QRLedger


def aplicar_regla3_sobre_pendientes(
    res: Dict,
    ledger: QRLedger,
    tol: float = 0.05,
    max_days_after: int = 7,
) -> Tuple[List[Dict], Set[str], pd.DataFrame]:
    """
    Bloquea pendientes QR de Banco-no-Flexxus que tienen cupón terminal histórico.

    res['pendientes_proxima'] se modifica IN-PLACE. Las filas bloqueadas se
    marcan con Estado='REGULARIZADO_REGLA3' y se DEJAN en el DataFrame para
    auditoría, pero NO se suman a B12 ni a C3.

    Devuelve (bloqueos, cupones_consumidos, pend_modificado)
    """
    bloqueos: List[Dict] = []
    cupones_usados: Set[str] = set()

    pend = res.get("pendientes_proxima", pd.DataFrame())
    if pend is None or pend.empty:
        return bloqueos, cupones_usados, pend

    # Filtrar pendientes QR de banco no Flexxus
    mask = (
        pend["Origen"].astype(str).eq("BANCO_NO_FLEXXUS") &
        pend["SignoCalculo"].astype(float).gt(0) &
        pend["Categoria"].astype(str).str.contains("QR", na=False)
    )
    qr_pend = pend[mask].copy()
    if qr_pend.empty:
        return bloqueos, cupones_usados, pend

    # Cupones terminales con bank_importe definido
    terminales = [
        e for e in ledger.by_cupon.values()
        if e.es_terminal() and e.bank_importe > 0
    ]
    if not terminales:
        return bloqueos, cupones_usados, pend

    # Ordenar pendientes por importe desc para procesamiento determinístico
    qr_pend = qr_pend.sort_values(
        by=["Monto", "Numero"],
        ascending=[False, True],
    )

    indices_a_quitar = []
    for idx, row in qr_pend.iterrows():
        importe = float(row["Monto"])
        fecha_str = str(row.get("FechaOrigen", "") or "")
        try:
            if "/" in fecha_str:
                d_pend = datetime.strptime(fecha_str[:10], "%d/%m/%Y").date()
            else:
                d_pend = datetime.strptime(fecha_str[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            continue

        candidatos = []
        for cupon in terminales:
            if cupon.cupon_qr in cupones_usados:
                continue
            if abs(cupon.bank_importe - importe) > tol:
                continue
            if cupon.bank_fecha:
                try:
                    d_hist = datetime.strptime(cupon.bank_fecha, "%Y-%m-%d").date()
                    delta = (d_pend - d_hist).days
                    if delta < 0 or delta > max_days_after:
                        continue
                except ValueError:
                    continue
            else:
                continue
            candidatos.append(cupon)

        if len(candidatos) == 1:
            cupon = candidatos[0]
            cupones_usados.add(cupon.cupon_qr)
            indices_a_quitar.append(idx)
            bloqueos.append({
                "ID Pendiente": str(row.get("pending_id", "")),
                "Fecha pend": fecha_str,
                "Comp pend": str(row.get("Numero", "")),
                "Importe pend": importe,
                "Cupon historico": cupon.cupon_id or cupon.cupon_qr[:12],
                "Local cupon": cupon.local,
                "Bank importe historico": cupon.bank_importe,
                "Bank fecha historico": cupon.bank_fecha,
                "Bank comprobante historico": cupon.bank_comprobante,
                "Periodo cierre cupon": cupon.periodo_cierre or cupon.periodo_origen,
                "Diferencia": round(importe - cupon.bank_importe, 2),
                "Diagnostico": "Bloqueado por regla 3: equivalencia QR histórica",
            })

    # Marcar los pendientes IN-PLACE con info del cupón histórico,
    # PERO DEJARLOS en B12. No se sacan del pending. Esto preserva la
    # continuidad contable entre períodos (el motor V5.9.11 del próximo
    # período los procesa normalmente). El valor de regla 3 es la
    # TRAZABILIDAD: el conciliador ve qué pendientes ya tienen cupón
    # identificado y no requieren investigación manual.
    if indices_a_quitar:
        # Construir mapping pending_id -> info del cupón
        info_por_pid = {}
        for b in bloqueos:
            info_por_pid[str(b["ID Pendiente"])] = {
                "cupon": b["Cupon historico"],
                "local": b["Local cupon"],
                "fecha_hist": b["Bank fecha historico"],
                "comp_hist": b["Bank comprobante historico"],
                "periodo": b["Periodo cierre cupon"],
            }
        # Aplicar a las filas afectadas
        for idx in indices_a_quitar:
            pid = str(pend.loc[idx, "pending_id"]) if "pending_id" in pend.columns else ""
            info = info_por_pid.get(pid)
            if info:
                # Anotar en columnas existentes
                if "Diagnostico" in pend.columns:
                    pend.at[idx, "Diagnostico"] = (
                        f"REGLA3: Identificado con cupón histórico {info['cupon']} "
                        f"({info['local']}, {info['fecha_hist']}, comp {info['comp_hist']}, "
                        f"cerrado en {info['periodo']})"
                    )
                if "Estado" in pend.columns:
                    pend.at[idx, "Estado"] = "ABIERTO_IDENTIFICADO_REGLA3"
        res["pendientes_proxima"] = pend
        # NO recompute: B12 no cambia, REG3 = 0, diferencia sigue cerrando
        # como lo hace el motor V5.9.11 puro.
        # Aseguramos REG3 = 0 explícitamente para el Excel.
        res["REG3"] = 0.0

    return bloqueos, cupones_usados, res.get("pendientes_proxima", pend)


def validar(bloqueos, cupones_usados, ledger):
    problemas = []
    if len(bloqueos) != len(cupones_usados):
        problemas.append(f"len(bloqueos)={len(bloqueos)} != len(cupones)={len(cupones_usados)}")
    for c in cupones_usados:
        e = ledger.cupon(c)
        if e is None or not e.es_terminal():
            problemas.append(f"cupon {c} no terminal")
    suma_b = sum(b["Importe pend"] for b in bloqueos)
    suma_c = sum(ledger.cupon(c).bank_importe for c in cupones_usados if ledger.cupon(c))
    if abs(suma_b - suma_c) > 0.10:
        problemas.append(f"suma {suma_b:.2f} != cupones {suma_c:.2f}")
    return problemas
