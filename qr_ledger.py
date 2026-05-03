"""
qr_ledger.py - Ledger persistente de cupones QR para conciliación bancaria Dancona.

Pieza central de V5.9.12. Reemplaza el matching por importe que arrastraban V5.9.x
con un ledger acumulativo donde cada fila es un CUPÓN QR identificado por su clave
estable (cupón + terminal + comercio + fecha + bruto). Cada cupón referencia
opcionalmente un bank_uid (línea de extracto bancario) y un flex_uid (línea
Flexxus). Una vez que un bank_uid o flex_uid queda ligado a un cupón, no puede
ser usado por otro cupón ni reaparecer como pendiente nuevo en períodos
posteriores.

Reglas duras (no negociables):
- 1 bank_uid solo puede pertenecer a 1 cupón.
- 1 flex_uid solo puede pertenecer a 1 cupón (no se reutilizan PAV).
- El ledger es ACUMULATIVO. Períodos anteriores no se reescriben salvo que
  intencionalmente se ejecute un "reseed" controlado.
- Estados terminales (CERRADO_TRIPLE, CERRADO_PAR_SIN_CUPON) no se reabren.

Estructura de cada fila del ledger (dict, persistido como JSON o parquet):

    cupon_qr:           PK lógica - hash estable (cupon+terminal+comercio+fecha+bruto)
    cupon_id:           Cupón tal cual viene en Transacciones QR
    fecha_qr:           ISO date (YYYY-MM-DD)
    local:              Local 1..8 si se identificó por CodComercio
    cod_comercio:       código bruto
    terminal:           terminal QR
    bruto:              monto total (Transacciones)
    neto_calc:          bruto * 0.99032

    bank_uid:           hash(fecha_banco|comprobante|importe) - opcional
    bank_fecha:         fecha de acreditación bancaria
    bank_importe:       monto que efectivamente acreditó banco
    bank_comprobante:   comprobante de extracto

    flex_uid:           hash(tipo|numero|fecha|monto) - opcional
    flex_fecha:         fecha del PAV
    flex_numero:        nro Flexxus
    flex_monto:         monto Flexxus
    flex_local:         local de origen del PAV

    estado:             SOLO_QR | QR_BANCO | QR_FLEXXUS | CERRADO_TRIPLE
                        BANCO_HUERFANO | FLEXXUS_HUERFANO | AMBIGUO_REVISAR
                        CERRADO_PAR_SIN_CUPON (estado de siembra histórica)
    confianza:          ALTA | MEDIA | BAJA
    motivo:             texto libre con explicación
    periodo_origen:     período donde apareció por primera vez (YYYY-MM-DD..YYYY-MM-DD)
    periodo_cierre:     período donde alcanzó estado terminal (vacío si abierto)
    creado:             ISO datetime
    actualizado:        ISO datetime
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# UIDs estables
# ---------------------------------------------------------------------------

def _norm(x: Any) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    return re.sub(r"\s+", " ", s).upper()


def _to_iso_date(x: Any) -> str:
    """Normaliza cualquier representación de fecha a 'YYYY-MM-DD' o ''."""
    if x is None or x == "":
        return ""
    if isinstance(x, (pd.Timestamp, datetime)):
        try:
            return pd.Timestamp(x).strftime("%Y-%m-%d")
        except Exception:
            return ""
    s = str(x).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    try:
        return pd.to_datetime(s, errors="coerce").strftime("%Y-%m-%d")
    except Exception:
        return ""


def _money(x: Any) -> float:
    if x is None or x == "":
        return 0.0
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        try:
            return round(float(str(x).replace(".", "").replace(",", ".")), 2)
        except Exception:
            return 0.0


def _hash(*parts: Any) -> str:
    raw = "|".join(_norm(p) for p in parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def cupon_key(cupon: Any, terminal: Any, comercio: Any, fecha: Any, bruto: Any) -> str:
    """Clave estable de cupón. Si falta cupón explícito, cae a la tupla
    (terminal, comercio, fecha, bruto) que también es única en la práctica."""
    return _hash("CUPON", cupon, terminal, comercio, _to_iso_date(fecha), f"{_money(bruto):.2f}")


def bank_uid(fecha: Any, comprobante: Any, importe: Any) -> str:
    return _hash("BANK", _to_iso_date(fecha), comprobante, f"{_money(importe):.2f}")


def flex_uid(tipo: Any, numero: Any, fecha: Any, monto: Any) -> str:
    return _hash("FLEX", tipo, numero, _to_iso_date(fecha), f"{_money(monto):.2f}")


# ---------------------------------------------------------------------------
# Estados
# ---------------------------------------------------------------------------

class Estado:
    SOLO_QR = "SOLO_QR"
    QR_BANCO = "QR_BANCO"
    QR_FLEXXUS = "QR_FLEXXUS"
    CERRADO_TRIPLE = "CERRADO_TRIPLE"
    BANCO_HUERFANO = "BANCO_HUERFANO"
    FLEXXUS_HUERFANO = "FLEXXUS_HUERFANO"
    AMBIGUO_REVISAR = "AMBIGUO_REVISAR"
    CERRADO_PAR_SIN_CUPON = "CERRADO_PAR_SIN_CUPON"

    TERMINALES = {CERRADO_TRIPLE, CERRADO_PAR_SIN_CUPON}


# ---------------------------------------------------------------------------
# Entrada del ledger
# ---------------------------------------------------------------------------

@dataclass
class LedgerEntry:
    cupon_qr: str
    cupon_id: str = ""
    fecha_qr: str = ""
    local: str = ""
    cod_comercio: str = ""
    terminal: str = ""
    bruto: float = 0.0
    neto_calc: float = 0.0

    bank_uid: str = ""
    bank_fecha: str = ""
    bank_importe: float = 0.0
    bank_comprobante: str = ""

    flex_uid: str = ""
    flex_fecha: str = ""
    flex_numero: str = ""
    flex_monto: float = 0.0
    flex_local: str = ""

    estado: str = Estado.SOLO_QR
    confianza: str = "MEDIA"
    motivo: str = ""
    periodo_origen: str = ""
    periodo_cierre: str = ""
    creado: str = ""
    actualizado: str = ""

    def es_terminal(self) -> bool:
        return self.estado in Estado.TERMINALES

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# Ledger - colección con índices secundarios
# ---------------------------------------------------------------------------

class QRLedger:
    """Colección indexada de LedgerEntry. Garantiza unicidad de bank_uid y flex_uid.

    Índices:
      - by_cupon[cupon_qr] -> LedgerEntry
      - by_bank[bank_uid]  -> cupon_qr
      - by_flex[flex_uid]  -> cupon_qr
    """

    def __init__(self) -> None:
        self.by_cupon: Dict[str, LedgerEntry] = {}
        self.by_bank: Dict[str, str] = {}
        self.by_flex: Dict[str, str] = {}

    # ---- carga / persistencia ----------------------------------------------

    @classmethod
    def from_json(cls, data: Any) -> "QRLedger":
        led = cls()
        if data is None:
            return led
        if isinstance(data, str):
            data = json.loads(data)
        for d in data:
            e = LedgerEntry(**{k: v for k, v in d.items() if k in LedgerEntry.__annotations__})
            led._index(e, replace=True)
        return led

    def to_json(self) -> str:
        return json.dumps([e.to_dict() for e in self.by_cupon.values()], ensure_ascii=False, indent=2)

    def to_dataframe(self) -> pd.DataFrame:
        if not self.by_cupon:
            return pd.DataFrame(columns=list(LedgerEntry.__annotations__.keys()))
        return pd.DataFrame([e.to_dict() for e in self.by_cupon.values()])

    # ---- índices internos --------------------------------------------------

    def _index(self, e: LedgerEntry, replace: bool = False) -> None:
        if e.cupon_qr in self.by_cupon and not replace:
            raise ValueError(f"cupon_qr ya existe: {e.cupon_qr}")
        # validar unicidad de bank_uid / flex_uid antes de indexar
        if e.bank_uid:
            owner = self.by_bank.get(e.bank_uid)
            if owner and owner != e.cupon_qr:
                raise ValueError(
                    f"bank_uid {e.bank_uid} ya pertenece a cupon {owner}, "
                    f"no puede asignarse a {e.cupon_qr}"
                )
        if e.flex_uid:
            owner = self.by_flex.get(e.flex_uid)
            if owner and owner != e.cupon_qr:
                raise ValueError(
                    f"flex_uid {e.flex_uid} ya pertenece a cupon {owner}, "
                    f"no puede asignarse a {e.cupon_qr}"
                )
        self.by_cupon[e.cupon_qr] = e
        if e.bank_uid:
            self.by_bank[e.bank_uid] = e.cupon_qr
        if e.flex_uid:
            self.by_flex[e.flex_uid] = e.cupon_qr

    # ---- consultas ---------------------------------------------------------

    def cupon(self, cupon_qr: str) -> Optional[LedgerEntry]:
        return self.by_cupon.get(cupon_qr)

    def bank_locked(self, uid: str) -> bool:
        owner = self.by_bank.get(uid)
        if not owner:
            return False
        e = self.by_cupon.get(owner)
        return bool(e and e.es_terminal())

    def flex_locked(self, uid: str) -> bool:
        owner = self.by_flex.get(uid)
        if not owner:
            return False
        e = self.by_cupon.get(owner)
        return bool(e and e.es_terminal())

    def bank_owner(self, uid: str) -> Optional[LedgerEntry]:
        owner = self.by_bank.get(uid)
        return self.by_cupon.get(owner) if owner else None

    def flex_owner(self, uid: str) -> Optional[LedgerEntry]:
        owner = self.by_flex.get(uid)
        return self.by_cupon.get(owner) if owner else None

    # ---- mutaciones --------------------------------------------------------

    def upsert_cupon(
        self,
        cupon_qr: str,
        defaults: Dict[str, Any],
    ) -> LedgerEntry:
        """Crea el cupón si no existe; si existe, no pisa los campos ya seteados."""
        now = datetime.now().isoformat(timespec="seconds")
        e = self.by_cupon.get(cupon_qr)
        if e is None:
            kwargs = {k: v for k, v in defaults.items() if k in LedgerEntry.__annotations__}
            kwargs["cupon_qr"] = cupon_qr
            kwargs.setdefault("creado", now)
            kwargs.setdefault("actualizado", now)
            e = LedgerEntry(**kwargs)
            self._index(e)
            return e
        # update no destructivo de campos vacíos
        for k, v in defaults.items():
            if k not in LedgerEntry.__annotations__:
                continue
            cur = getattr(e, k, None)
            if cur in (None, "", 0, 0.0) and v not in (None, "", 0, 0.0):
                setattr(e, k, v)
        e.actualizado = now
        return e

    def attach_bank(
        self,
        cupon_qr: str,
        b_uid: str,
        bank_fecha: str,
        bank_importe: float,
        bank_comprobante: str,
        motivo: str = "",
    ) -> None:
        e = self.by_cupon[cupon_qr]
        if e.bank_uid and e.bank_uid != b_uid:
            raise ValueError(
                f"cupon {cupon_qr} ya tiene bank_uid {e.bank_uid}, "
                f"no se puede sobrescribir con {b_uid}"
            )
        owner = self.by_bank.get(b_uid)
        if owner and owner != cupon_qr:
            raise ValueError(
                f"bank_uid {b_uid} ya pertenece a cupon {owner}"
            )
        e.bank_uid = b_uid
        e.bank_fecha = bank_fecha
        e.bank_importe = float(bank_importe)
        e.bank_comprobante = bank_comprobante
        if motivo:
            e.motivo = (e.motivo + " | " if e.motivo else "") + motivo
        self.by_bank[b_uid] = cupon_qr
        self._recompute_estado(e)
        e.actualizado = datetime.now().isoformat(timespec="seconds")

    def attach_flex(
        self,
        cupon_qr: str,
        f_uid: str,
        flex_fecha: str,
        flex_numero: str,
        flex_monto: float,
        flex_local: str = "",
        motivo: str = "",
    ) -> None:
        e = self.by_cupon[cupon_qr]
        if e.flex_uid and e.flex_uid != f_uid:
            raise ValueError(
                f"cupon {cupon_qr} ya tiene flex_uid {e.flex_uid}, "
                f"no se puede sobrescribir con {f_uid}"
            )
        owner = self.by_flex.get(f_uid)
        if owner and owner != cupon_qr:
            raise ValueError(
                f"flex_uid {f_uid} ya pertenece a cupon {owner}"
            )
        e.flex_uid = f_uid
        e.flex_fecha = flex_fecha
        e.flex_numero = flex_numero
        e.flex_monto = float(flex_monto)
        e.flex_local = flex_local
        if motivo:
            e.motivo = (e.motivo + " | " if e.motivo else "") + motivo
        self.by_flex[f_uid] = cupon_qr
        self._recompute_estado(e)
        e.actualizado = datetime.now().isoformat(timespec="seconds")

    def mark_estado(self, cupon_qr: str, estado: str, motivo: str = "", periodo_cierre: str = "") -> None:
        e = self.by_cupon[cupon_qr]
        if e.es_terminal() and estado != e.estado:
            raise ValueError(
                f"cupon {cupon_qr} está en estado terminal {e.estado}; "
                f"no se puede cambiar a {estado}"
            )
        e.estado = estado
        if motivo:
            e.motivo = (e.motivo + " | " if e.motivo else "") + motivo
        if periodo_cierre and estado in Estado.TERMINALES:
            e.periodo_cierre = periodo_cierre
        e.actualizado = datetime.now().isoformat(timespec="seconds")

    def _recompute_estado(self, e: LedgerEntry) -> None:
        """Recalcula estado a partir de los campos. Nunca degrada un estado terminal."""
        if e.es_terminal():
            return
        b = bool(e.bank_uid)
        f = bool(e.flex_uid)
        if b and f:
            e.estado = Estado.CERRADO_TRIPLE
            e.confianza = "ALTA"
        elif b and not f:
            e.estado = Estado.QR_BANCO
        elif f and not b:
            e.estado = Estado.QR_FLEXXUS
        else:
            e.estado = Estado.SOLO_QR

    # ---- invariantes -------------------------------------------------------

    def assert_invariants(self) -> List[str]:
        """Devuelve lista de violaciones encontradas. Vacío = ok."""
        problems: List[str] = []
        # 1) ningún bank_uid en más de un cupón
        seen_b: Dict[str, str] = {}
        for c, e in self.by_cupon.items():
            if e.bank_uid:
                if e.bank_uid in seen_b:
                    problems.append(
                        f"bank_uid {e.bank_uid} en cupones {seen_b[e.bank_uid]} y {c}"
                    )
                else:
                    seen_b[e.bank_uid] = c
        # 2) ningún flex_uid en más de un cupón
        seen_f: Dict[str, str] = {}
        for c, e in self.by_cupon.items():
            if e.flex_uid:
                if e.flex_uid in seen_f:
                    problems.append(
                        f"flex_uid {e.flex_uid} en cupones {seen_f[e.flex_uid]} y {c}"
                    )
                else:
                    seen_f[e.flex_uid] = c
        # 3) índices coherentes
        for uid, c in self.by_bank.items():
            e = self.by_cupon.get(c)
            if e is None or e.bank_uid != uid:
                problems.append(f"by_bank[{uid}]={c} pero cupon no existe o bank_uid distinto")
        for uid, c in self.by_flex.items():
            e = self.by_cupon.get(c)
            if e is None or e.flex_uid != uid:
                problems.append(f"by_flex[{uid}]={c} pero cupon no existe o flex_uid distinto")
        return problems


# ---------------------------------------------------------------------------
# Ingesta y matching cupón-céntrico
# ---------------------------------------------------------------------------

LOCAL_MAP_DEFAULT = {
    "29841642": "Local 1",
    "29841627": "Local 2",
    "29841683": "Local 3",
    "29841670": "Local 5",
    "29841705": "Local 6",
    "31995644": "Local 7",
    "32032899": "Local 8",
}


def ingest_qr_transactions(
    led: QRLedger,
    qr_df: pd.DataFrame,
    periodo: str,
    local_map: Optional[Dict[str, str]] = None,
) -> Tuple[QRLedger, int]:
    """Carga cupones de Transacciones QR PCT al ledger.

    qr_df se asume con las columnas que produce parse_qr() de la app:
      Cupon, TerminalQR, CodComercio, FechaQR, MontoTotal, NetoQR

    Devuelve (ledger, nuevos_creados).
    """
    if qr_df is None or qr_df.empty:
        return led, 0
    local_map = local_map or LOCAL_MAP_DEFAULT
    nuevos = 0
    for _, r in qr_df.iterrows():
        cupon = str(r.get("Cupon", "") or "").strip()
        terminal = str(r.get("TerminalQR", "") or "").strip()
        comercio = re.sub(r"[^0-9]", "", str(r.get("CodComercio", "") or ""))
        fecha = _to_iso_date(r.get("FechaQR", ""))
        bruto = _money(r.get("MontoTotal", 0))
        neto = _money(r.get("NetoQR", 0)) or round(bruto * 0.99032, 2)
        if bruto <= 0:
            continue
        key = cupon_key(cupon, terminal, comercio, fecha, bruto)
        existed = key in led.by_cupon
        led.upsert_cupon(
            key,
            {
                "cupon_id": cupon,
                "fecha_qr": fecha,
                "cod_comercio": comercio,
                "local": local_map.get(comercio, comercio),
                "terminal": terminal,
                "bruto": bruto,
                "neto_calc": neto,
                "periodo_origen": periodo,
            },
        )
        if not existed:
            nuevos += 1
    return led, nuevos


def attach_bank_to_cupons(
    led: QRLedger,
    bank_df: pd.DataFrame,
    periodo: str,
    tol: float = 0.05,
    max_days_after: int = 5,
) -> Dict[str, Any]:
    """Recorre líneas QR del banco y las liga al cupón correcto.

    bank_df debe tener: Fecha, Comprobante, ImporteAbs, Categoria, EsIngreso, Fecha_dt.

    Reglas:
      - sólo categoria QR / EsIngreso=True
      - bank_uid bloqueado si ya está en el ledger atado a otro cupón
      - match por (importe ≈ neto_calc, fecha_banco entre fecha_qr y fecha_qr+5d)
      - si hay >1 cupón candidato → AMBIGUO_REVISAR (no se ata)
    """
    stats = {"asignados": 0, "huerfanos": 0, "ambiguos": 0, "ya_asignados": 0}
    if bank_df is None or bank_df.empty:
        return stats

    qr_lines = bank_df[
        bank_df.get("EsIngreso", False) & bank_df.get("Categoria", "").astype(str).eq("QR")
    ].copy()
    if qr_lines.empty:
        return stats

    # set de cupones que todavía no tienen bank_uid (se actualiza vivo en el loop)
    cupones_sin_bank = {e.cupon_qr for e in led.by_cupon.values() if not e.bank_uid}

    for _, b in qr_lines.iterrows():
        fecha_b = _to_iso_date(b.get("Fecha", b.get("Fecha_dt", "")))
        comprob = str(b.get("Comprobante", "") or "").strip()
        importe = _money(b.get("ImporteAbs", 0))
        b_uid = bank_uid(fecha_b, comprob, importe)

        # ya asignado en alguna corrida previa
        if b_uid in led.by_bank:
            stats["ya_asignados"] += 1
            continue

        # candidatos: cupones AÚN sin bank_uid, importe ≈, fecha en ventana
        candidatos = []
        for cupon_qr_id in cupones_sin_bank:
            e = led.by_cupon.get(cupon_qr_id)
            if e is None or e.bank_uid:  # defensivo
                continue
            if abs(e.neto_calc - importe) > tol:
                continue
            if e.fecha_qr and fecha_b:
                try:
                    d_qr = datetime.strptime(e.fecha_qr, "%Y-%m-%d").date()
                    d_b = datetime.strptime(fecha_b, "%Y-%m-%d").date()
                    delta = (d_b - d_qr).days
                    if delta < -1 or delta > max_days_after:
                        continue
                except ValueError:
                    pass
            candidatos.append(e)

        if len(candidatos) == 1:
            e = candidatos[0]
            led.attach_bank(
                e.cupon_qr, b_uid,
                bank_fecha=fecha_b, bank_importe=importe, bank_comprobante=comprob,
                motivo=f"banco atado en {periodo} por neto y fecha",
            )
            stats["asignados"] += 1
            cupones_sin_bank.discard(e.cupon_qr)
        elif len(candidatos) > 1:
            # Desempate: el cupón con fecha_qr MÁS CERCANA a la fecha banco gana.
            # Si hay empate de proximidad, marcar como ambiguo y registrar huérfano.
            try:
                d_b = datetime.strptime(fecha_b, "%Y-%m-%d").date()
                cand_con_dist = []
                for c in candidatos:
                    try:
                        d_q = datetime.strptime(c.fecha_qr, "%Y-%m-%d").date()
                        cand_con_dist.append((c, abs((d_b - d_q).days)))
                    except (ValueError, TypeError):
                        cand_con_dist.append((c, 9999))
                cand_con_dist.sort(key=lambda x: x[1])
                if len(cand_con_dist) >= 2 and cand_con_dist[0][1] == cand_con_dist[1][1]:
                    # empate real: no podemos decidir
                    _registrar_huerfano_banco(
                        led, b_uid, fecha_b, importe, comprob, periodo,
                        ambiguos=[c.cupon_qr for c in candidatos],
                    )
                    stats["ambiguos"] += 1
                else:
                    # ganador único por proximidad
                    e = cand_con_dist[0][0]
                    led.attach_bank(
                        e.cupon_qr, b_uid,
                        bank_fecha=fecha_b, bank_importe=importe, bank_comprobante=comprob,
                        motivo=f"banco atado en {periodo} por neto+fecha (desempate proximidad)",
                    )
                    stats["asignados"] += 1
                    cupones_sin_bank.discard(e.cupon_qr)
            except (ValueError, TypeError):
                _registrar_huerfano_banco(
                    led, b_uid, fecha_b, importe, comprob, periodo,
                    ambiguos=[c.cupon_qr for c in candidatos],
                )
                stats["ambiguos"] += 1
        else:
            _registrar_huerfano_banco(led, b_uid, fecha_b, importe, comprob, periodo)
            stats["huerfanos"] += 1
    return stats


def _registrar_huerfano_banco(
    led: QRLedger, b_uid: str, fecha: str, importe: float, comprob: str,
    periodo: str, ambiguos: Optional[List[str]] = None,
) -> None:
    """Crea una entrada BANCO_HUERFANO con cupon ficticio para mantener trazabilidad."""
    if b_uid in led.by_bank:
        return
    fake_cupon = _hash("BANCO_HUERFANO", b_uid)
    if fake_cupon in led.by_cupon:
        return
    motivo = f"CR DEBIN sin cupón identificado en {periodo}"
    estado = Estado.AMBIGUO_REVISAR if ambiguos else Estado.BANCO_HUERFANO
    if ambiguos:
        motivo += f" - cupones candidatos por importe: {','.join(ambiguos[:5])}"
    led.upsert_cupon(
        fake_cupon,
        {
            "cupon_id": "",
            "fecha_qr": "",
            "bank_uid": b_uid,
            "bank_fecha": fecha,
            "bank_importe": importe,
            "bank_comprobante": comprob,
            "estado": estado,
            "confianza": "BAJA",
            "motivo": motivo,
            "periodo_origen": periodo,
        },
    )
    led.by_bank[b_uid] = fake_cupon


def attach_flexxus_to_cupons(
    led: QRLedger,
    flex_df: pd.DataFrame,
    periodo: str,
    tol: float = 1.00,
    max_days_after: int = 3,
    max_days_before: int = 2,
) -> Dict[str, Any]:
    """Liga PAV de Flexxus a cupones del ledger. Solo PAV no-PedidosYa.

    flex_df debe traer: Tipo, Numero, FechaFlexxus, MontoFlexxus, EsPedidosYa, Movimiento.
    
    Ventana temporal: delta = fecha_PAV - fecha_QR ∈ [-max_days_before, max_days_after].
    Default [-2, +3] cubre el 100% de los casos observados en datos reales Dancona.
    """
    stats = {"asignados": 0, "ya_asignados": 0, "huerfanos": 0, "ambiguos": 0}
    if flex_df is None or flex_df.empty:
        return stats

    pav = flex_df[
        flex_df.get("Tipo", "").astype(str).eq("PAV") & ~flex_df.get("EsPedidosYa", False)
    ].copy()
    if pav.empty:
        return stats

    cupones_sin_flex = {e.cupon_qr for e in led.by_cupon.values() if not e.flex_uid and e.bruto > 0}

    for _, f in pav.iterrows():
        tipo = "PAV"
        numero = str(f.get("Numero", "") or "").strip()
        fecha_f = _to_iso_date(f.get("FechaFlexxus", ""))
        monto = _money(f.get("MontoFlexxus", 0))
        f_uid = flex_uid(tipo, numero, fecha_f, monto)

        if f_uid in led.by_flex:
            stats["ya_asignados"] += 1
            continue

        # match por neto (PAV registra el neto que el banco va a acreditar,
        # NO el bruto del cupón QR)
        candidatos = []
        for cupon_qr_id in cupones_sin_flex:
            e = led.by_cupon.get(cupon_qr_id)
            if e is None or e.flex_uid:
                continue
            if abs(e.neto_calc - monto) > tol:
                continue
            if e.fecha_qr and fecha_f:
                try:
                    d_qr = datetime.strptime(e.fecha_qr, "%Y-%m-%d").date()
                    d_f = datetime.strptime(fecha_f, "%Y-%m-%d").date()
                    delta = (d_f - d_qr).days
                    if delta < -max_days_before or delta > max_days_after:
                        continue
                except ValueError:
                    pass
            candidatos.append(e)

        if len(candidatos) == 1:
            e = candidatos[0]
            led.attach_flex(
                e.cupon_qr, f_uid,
                flex_fecha=fecha_f, flex_numero=numero, flex_monto=monto,
                flex_local=e.local,
                motivo=f"flex atado en {periodo} por bruto y fecha",
            )
            stats["asignados"] += 1
            cupones_sin_flex.discard(e.cupon_qr)
        elif len(candidatos) > 1:
            # Desempate por proximidad de fecha. Si empate, queda ambiguo.
            try:
                d_f = datetime.strptime(fecha_f, "%Y-%m-%d").date()
                cand_dist = []
                for c in candidatos:
                    try:
                        d_q = datetime.strptime(c.fecha_qr, "%Y-%m-%d").date()
                        cand_dist.append((c, abs((d_f - d_q).days)))
                    except (ValueError, TypeError):
                        cand_dist.append((c, 9999))
                cand_dist.sort(key=lambda x: x[1])
                if len(cand_dist) >= 2 and cand_dist[0][1] == cand_dist[1][1]:
                    stats["ambiguos"] += 1
                else:
                    e = cand_dist[0][0]
                    led.attach_flex(
                        e.cupon_qr, f_uid,
                        flex_fecha=fecha_f, flex_numero=numero, flex_monto=monto,
                        flex_local=e.local,
                        motivo=f"flex atado en {periodo} por neto+fecha (desempate proximidad)",
                    )
                    stats["asignados"] += 1
                    cupones_sin_flex.discard(e.cupon_qr)
            except (ValueError, TypeError):
                stats["ambiguos"] += 1
        else:
            stats["huerfanos"] += 1
    return stats


# ---------------------------------------------------------------------------
# Lookup helpers para que el motor de matching sepa qué está bloqueado
# ---------------------------------------------------------------------------

def bank_uids_locked(led: QRLedger) -> set:
    """bank_uids que pertenecen a cupones en estado terminal."""
    return {uid for uid, c in led.by_bank.items()
            if (e := led.by_cupon.get(c)) is not None and e.es_terminal()}


def flex_uids_locked(led: QRLedger) -> set:
    """flex_uids que pertenecen a cupones en estado terminal."""
    return {uid for uid, c in led.by_flex.items()
            if (e := led.by_cupon.get(c)) is not None and e.es_terminal()}


def bank_uids_assigned(led: QRLedger) -> set:
    """bank_uids ya atados a algún cupón (terminal o no)."""
    return set(led.by_bank.keys())


def flex_uids_assigned(led: QRLedger) -> set:
    """flex_uids ya atados a algún cupón (terminal o no)."""
    return set(led.by_flex.keys())


# ---------------------------------------------------------------------------
# Resumen para reporting
# ---------------------------------------------------------------------------

def resumen_estados(led: QRLedger) -> pd.DataFrame:
    rows = []
    by_estado: Dict[str, List[LedgerEntry]] = {}
    for e in led.by_cupon.values():
        by_estado.setdefault(e.estado, []).append(e)
    for est, items in sorted(by_estado.items()):
        rows.append({
            "Estado": est,
            "Cantidad": len(items),
            "Suma bruto": round(sum(i.bruto for i in items), 2),
            "Suma neto": round(sum(i.neto_calc for i in items), 2),
            "Suma banco": round(sum(i.bank_importe for i in items), 2),
            "Suma flex": round(sum(i.flex_monto for i in items), 2),
        })
    return pd.DataFrame(rows)
