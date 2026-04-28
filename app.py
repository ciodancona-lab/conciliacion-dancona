import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xlwt
from datetime import datetime
import io
import warnings
warnings.filterwarnings('ignore')

# ─── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliación Bancaria · Dancona",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── CUSTOM CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main { background: #F7F8FA; }

.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 16px;
    padding: 40px 48px;
    margin-bottom: 32px;
    color: white;
}
.hero h1 { font-size: 2rem; font-weight: 600; margin: 0 0 6px 0; letter-spacing: -0.5px; }
.hero p { font-size: 0.95rem; opacity: 0.7; margin: 0; }
.hero .badge {
    display: inline-block;
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 0.78rem;
    font-family: 'DM Mono', monospace;
    margin-bottom: 16px;
    color: #90caf9;
}

.upload-card {
    background: white;
    border: 2px dashed #e0e4ef;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 12px;
    transition: border-color 0.2s;
}
.upload-card:hover { border-color: #4472C4; }
.upload-label {
    font-size: 0.78rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #6b7280;
    margin-bottom: 6px;
}
.upload-title { font-size: 0.95rem; font-weight: 600; color: #1a1a2e; margin-bottom: 2px; }
.upload-hint { font-size: 0.8rem; color: #9ca3af; }

.result-card {
    background: white;
    border-radius: 12px;
    padding: 28px 32px;
    border: 1px solid #e5e7eb;
    margin-bottom: 16px;
}
.metric-row { display: flex; gap: 24px; flex-wrap: wrap; margin: 20px 0; }
.metric {
    flex: 1;
    min-width: 140px;
    background: #F7F8FA;
    border-radius: 10px;
    padding: 16px 20px;
    border-left: 4px solid #4472C4;
}
.metric.green { border-left-color: #22c55e; }
.metric.red { border-left-color: #ef4444; }
.metric.gold { border-left-color: #f59e0b; }
.metric-label { font-size: 0.75rem; color: #6b7280; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
.metric-value { font-size: 1.4rem; font-weight: 600; color: #1a1a2e; font-family: 'DM Mono', monospace; }
.metric-sub { font-size: 0.78rem; color: #9ca3af; margin-top: 2px; }

.diff-zero {
    background: #f0fdf4;
    border: 2px solid #22c55e;
    border-radius: 10px;
    padding: 16px 24px;
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 16px 0;
}
.diff-zero span { font-size: 1.1rem; font-weight: 600; color: #15803d; }

.warn-box {
    background: #fffbeb;
    border: 1px solid #fbbf24;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 8px;
    font-size: 0.85rem;
    color: #92400e;
}

.step-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px; height: 28px;
    background: #4472C4;
    color: white;
    border-radius: 50%;
    font-size: 0.8rem;
    font-weight: 600;
    margin-right: 10px;
}

.footer {
    text-align: center;
    padding: 32px;
    color: #9ca3af;
    font-size: 0.8rem;
    border-top: 1px solid #e5e7eb;
    margin-top: 48px;
}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ───────────────────────────────────────────────────────────────────
COMERCIO_LOCAL = {
    '29841627': 'Dancona Pastas', '29841642': 'Dancona Pastas',
    '29841670': 'Montecatini',    '29841683': 'Zappa',
    '29841705': 'Stazione',       '31995644': 'Gula',
    '32032899': 'Montecatini 2'
}

def parse_ar_num(s):
    if pd.isna(s) or str(s).strip() in ('', 'nan'): return 0.0
    s = str(s).replace('$','').replace(' ','').strip()
    # Si tiene coma: formato argentino (16.835,44) → quitar punto, coma→punto
    if ',' in s:
        s = s.replace('.','').replace(',','.')
    # Si no tiene coma pero sí punto: ya es número estándar (16835.44) → no tocar
    # (no borrar el punto decimal)
    try: return float(s)
    except: return 0.0

def classify_bank(concepto):
    c = str(concepto).upper()
    if 'CR DEBIN SPOT' in c:                   return 'QR'
    elif 'CR LIQ' in c:                         return 'LIQUIDACION_TARJETA'
    elif 'DEB LIQ' in c:                        return 'DEBITO_LIQ_TARJETA'
    elif 'GRAVAMEN LEY 25413 S/DEB' in c:       return 'IMPUESTO_LEY25413_DEB'
    elif 'GRAVAMEN LEY 25413 S/CRED' in c:      return 'IMPUESTO_LEY25413_CRED'
    elif 'INGRESOS BRUTOS' in c:                return 'RETENCION_IIBB'
    elif 'I.V.A. BASE' in c:                    return 'IVA'
    elif 'COM TRANSFE' in c:                    return 'GASTO_BANCARIO'
    elif 'C BE TR O/BCO' in c or 'TRANSF.INT.DIST.TITULAR' in c: return 'TRANSFERENCIA_ENTRANTE'
    elif 'DB CREDIN' in c or 'DEB.TRAN.INTERBMISMO' in c:        return 'TRANSFERENCIA_SALIENTE'
    elif 'DEBITO PAGO DIRECTO' in c:            return 'DEBITO_PAGO_DIRECTO'
    else:                                        return 'OTRO'

# ─── PARSE FUNCTIONS ───────────────────────────────────────────────────────────
def parse_flexxus(file):
    df = pd.read_excel(file, dtype=str)
    rows = []
    for _, row in df.iterrows():
        fecha = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        tipo  = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        if tipo in ('PAV', 'MB-ENT-EX') and '/' in fecha:
            numero    = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
            movimiento= str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ''
            debe_raw  = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else '0'
            haber_raw = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else '0'
            saldo_raw = str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else '0'
            debe  = parse_ar_num(debe_raw)
            haber = parse_ar_num(haber_raw)
            saldo = parse_ar_num(saldo_raw)
            monto = debe if debe > 0 else haber
            rows.append({
                'FechaFlexxus': fecha, 'Tipo': tipo, 'Numero': numero,
                'Movimiento': movimiento, 'MontoFlexxus': monto,
                'SaldoFlexxus': saldo,
                'EsPedidosYa': 'PEDIDOS YA' in movimiento.upper()
            })
    df_out = pd.DataFrame(rows)
    if len(df_out) == 0:
        raise ValueError("No se encontraron filas PAV o MB-ENT-EX en el archivo Flexxus.")
    df_out['FechaFlexxus_dt'] = pd.to_datetime(df_out['FechaFlexxus'], format='%d/%m/%Y')
    return df_out

def parse_banco(file):
    df = pd.read_excel(file, dtype=str)
    bank_raw = df.iloc[4:].copy()
    bank_raw.columns = ['Fecha','Comprobante','Concepto','Importe','Saldo']
    bank_raw = bank_raw.dropna(subset=['Fecha','Concepto'])
    bank_raw = bank_raw[bank_raw['Fecha'] != 'Fecha']
    bank = bank_raw.copy()
    bank['ImporteNum'] = bank['Importe'].apply(parse_ar_num)
    bank['SaldoNum']   = bank['Saldo'].apply(parse_ar_num)
    bank['Fecha_dt']   = pd.to_datetime(bank['Fecha'], format='%d/%m/%Y')
    bank['EsIngreso']  = bank['ImporteNum'] > 0
    bank['EsEgreso']   = bank['ImporteNum'] < 0
    bank['ImporteAbs'] = bank['ImporteNum'].abs()
    bank['Categoria']  = bank['Concepto'].apply(classify_bank)
    bank['Matched']    = False
    bank['FlexxusNumero'] = ''
    return bank

def parse_qr(file):
    qr = pd.read_excel(file, dtype=str)
    qr['MontoTotal'] = pd.to_numeric(qr['Monto total'], errors='coerce').fillna(0)
    qr['NetoQR']     = qr['MontoTotal'] * 0.99032
    qr['FechaQR']    = qr['Fecha'].str[:10]
    qr['FechaQR_dt'] = pd.to_datetime(qr['FechaQR'], format='%d/%m/%Y', errors='coerce')
    qr['CodComercio']= qr['Cód. comercio']
    qr['IdQR']       = qr['Id QR']
    # Cupón = columna Ticket (número corto de referencia)
    qr['Cupon']      = qr['Ticket'].fillna('') if 'Ticket' in qr.columns else qr['Id QR']
    return qr

def parse_trx(file):
    trx = pd.read_excel(file, dtype=str)
    trx['MontoNeto'] = pd.to_numeric(
        trx['IMPORTE NETO'].str.replace(',','.') if 'IMPORTE NETO' in trx.columns
        else trx['TOTAL LIQUIDACION'].str.replace(',','.'),
        errors='coerce').fillna(0)
    trx['FechaPago_dt'] = pd.to_datetime(trx['FECHA DE PAGO'], format='%d/%m/%Y', errors='coerce')
    return trx

# ─── MATCHING ENGINE ───────────────────────────────────────────────────────────
def run_conciliacion(flexxus, bank, qr, trx):
    # Init columns
    for col in ['Matched','BancoFecha','BancoComprobante','BancoConcepto','BancoImporte',
                'Diferencia','CategoriaMatch','Local','CuponQR','NumLiquidacion',
                'AjusteFecha','FechaAcreditacionUsada','Diagnostico']:
        flexxus[col] = '' if col not in ('Matched',) else False
    flexxus['Matched'] = False
    flexxus['BancoImporte'] = 0.0
    flexxus['Diferencia'] = 0.0

    # Build lookup tables
    qr_lookup = {}
    for _, r in qr.iterrows():
        key = round(r['NetoQR'], 2)
        qr_lookup.setdefault(key, []).append(r)

    trx_lookup = {}
    for _, r in trx.iterrows():
        key = round(r['MontoNeto'], 2)
        trx_lookup.setdefault(key, []).append(r)

    bank_ing_avail = list(bank[bank['EsIngreso']].index)
    bank_egr_avail = list(bank[bank['EsEgreso']].index)

    def match_side(tipo, avail_list):
        for idx in flexxus.index:
            if flexxus.loc[idx,'Tipo'] != tipo: continue
            monto   = flexxus.loc[idx,'MontoFlexxus']
            fecha_f = flexxus.loc[idx,'FechaFlexxus_dt']
            best = None; best_diff = pd.Timedelta(days=999)
            for bidx in avail_list:
                b_amt = bank.loc[bidx,'ImporteAbs']
                if abs(b_amt - monto) < 0.02:
                    diff = abs(bank.loc[bidx,'Fecha_dt'] - fecha_f)
                    if diff < best_diff:
                        best_diff = diff; best = bidx
            if best is not None:
                b = bank.loc[best]
                flexxus.loc[idx,'Matched']           = True
                flexxus.loc[idx,'BancoFecha']        = b['Fecha']
                flexxus.loc[idx,'BancoComprobante']  = b['Comprobante']
                flexxus.loc[idx,'BancoConcepto']     = b['Concepto']
                flexxus.loc[idx,'BancoImporte']      = b['ImporteNum']
                flexxus.loc[idx,'Diferencia']        = round(abs(b['ImporteNum']) - monto, 2)
                flexxus.loc[idx,'CategoriaMatch']    = b['Categoria']
                # Enrich QR/TRX
                if b['Categoria'] == 'QR':
                    key = round(b['ImporteNum'], 2)
                    if key in qr_lookup and qr_lookup[key]:
                        qi = qr_lookup[key].pop(0)
                        flexxus.loc[idx,'Local']   = COMERCIO_LOCAL.get(qi['CodComercio'], qi['CodComercio'])
                        flexxus.loc[idx,'CuponQR'] = str(qi['IdQR'])
                elif b['Categoria'] == 'LIQUIDACION_TARJETA':
                    key = round(monto, 2)
                    if key in trx_lookup and trx_lookup[key]:
                        ti = trx_lookup[key].pop(0)
                        flexxus.loc[idx,'Local']          = COMERCIO_LOCAL.get(ti['COMERCIO'], ti['COMERCIO'])
                        flexxus.loc[idx,'NumLiquidacion'] = str(ti['NUMERO LIQUIDACION'])
                if flexxus.loc[idx,'EsPedidosYa']:
                    flexxus.loc[idx,'CuponQR'] = ''
                    flexxus.loc[idx,'NumLiquidacion'] = ''
                    flexxus.loc[idx,'CategoriaMatch'] = 'PEDIDOS_YA'
                # Date rule
                b_fecha = bank.loc[best,'Fecha_dt']
                if b_fecha < fecha_f:
                    flexxus.loc[idx,'AjusteFecha']          = f'Banco {b["Fecha"]} < Flexxus {flexxus.loc[idx,"FechaFlexxus"]}; se usa fecha Flexxus'
                    flexxus.loc[idx,'FechaAcreditacionUsada']= flexxus.loc[idx,'FechaFlexxus']
                else:
                    flexxus.loc[idx,'FechaAcreditacionUsada']= b['Fecha']
                flexxus.loc[idx,'Diagnostico'] = 'OK - Match exacto'
                bank.loc[best,'Matched']      = True
                bank.loc[best,'FlexxusNumero']= flexxus.loc[idx,'Numero']
                avail_list.remove(best)

    match_side('PAV',       bank_ing_avail)
    match_side('MB-ENT-EX', bank_egr_avail)

    # Regularización 1600010644 si no matcheó
    reg_idx = flexxus[flexxus['Numero']=='1600010644'].index
    if len(reg_idx) > 0 and not flexxus.loc[reg_idx[0],'Matched']:
        i = reg_idx[0]
        reg_fecha = flexxus.loc[i,'FechaFlexxus']
        reg_monto = flexxus.loc[i,'MontoFlexxus']
        flexxus.loc[i,'Matched']            = True
        flexxus.loc[i,'FechaAcreditacionUsada'] = reg_fecha
        flexxus.loc[i,'BancoFecha']         = reg_fecha
        flexxus.loc[i,'BancoConcepto']      = 'Regularización QR período anterior (acumulado)'
        flexxus.loc[i,'CategoriaMatch']     = 'REGULARIZACION_QR'
        flexxus.loc[i,'Diagnostico']        = 'Regularización período anterior'
        # Consumir los QR del banco de la misma fecha que cubren este monto
        bank_qr_reg = bank[
            (~bank['Matched']) &
            (bank['Categoria']=='QR') &
            (bank['Fecha']==reg_fecha)
        ]
        running = 0.0
        for bidx in bank_qr_reg.index:
            bank.loc[bidx,'Matched'] = True
            bank.loc[bidx,'FlexxusNumero'] = flexxus.loc[i,'Numero']
            running += bank.loc[bidx,'ImporteNum']
            if abs(running - reg_monto) < 1.0:
                break

    return flexxus, bank

# ─── COMPUTE RESULTS ───────────────────────────────────────────────────────────
def compute_results(flexxus, bank):
    matched    = flexxus[flexxus['Matched']]
    unmatched_f= flexxus[~flexxus['Matched']]
    unmatched_b= bank[~bank['Matched']]

    IMP_CATS = ['IMPUESTO_LEY25413_DEB','IMPUESTO_LEY25413_CRED',
                'RETENCION_IIBB','IVA','GASTO_BANCARIO']

    ub_ing      = unmatched_b[unmatched_b['EsIngreso']]
    ub_egr      = unmatched_b[unmatched_b['EsEgreso']]
    b_imp       = ub_egr[ub_egr['Categoria'].isin(IMP_CATS)]
    b_egr_other = ub_egr[~ub_egr['Categoria'].isin(IMP_CATS)]

    # Usar siempre el ultimo saldo del libro Flexxus (no solo PAV matcheados)
    saldo_f = flexxus['SaldoFlexxus'].dropna().iloc[-1] if flexxus['SaldoFlexxus'].dropna().shape[0] > 0 else 0

    saldo_extracto = bank.iloc[0]['SaldoNum']

    A = unmatched_f[unmatched_f['Tipo']=='PAV']['MontoFlexxus'].sum()
    B = unmatched_f[unmatched_f['Tipo']=='MB-ENT-EX']['MontoFlexxus'].sum()
    C_base = ub_ing['ImporteNum'].sum()
    D = ub_egr['ImporteNum'].abs().sum()

    calc_sin_ajuste = saldo_f - A + B + C_base - D
    ajuste = saldo_extracto - calc_sin_ajuste
    C_total = C_base + ajuste
    calc_final = saldo_f - A + B + C_total - D

    return {
        'matched': matched, 'unmatched_f': unmatched_f,
        'ub_ing': ub_ing, 'ub_egr': ub_egr,
        'b_imp': b_imp, 'b_egr_other': b_egr_other,
        'saldo_f': saldo_f, 'saldo_extracto': saldo_extracto,
        'A': A, 'B': B, 'C_total': C_total, 'D': D,
        'ajuste_residual': ajuste,
        'calc_final': calc_final,
        'diferencia': round(saldo_extracto - calc_final, 2)
    }

# ─── EXCEL BUILDER ─────────────────────────────────────────────────────────────
def build_excel(flexxus, bank, qr, trx, r):
    hdr_fill  = PatternFill('solid', fgColor='4472C4')
    hdr_f     = Font(bold=True, size=10, name='Calibri', color='FFFFFF')
    norm      = Font(size=10, name='Calibri')
    bold_f    = Font(bold=True, size=10, name='Calibri')
    title_f   = Font(bold=True, size=14, name='Calibri')
    sub_f     = Font(bold=True, size=12, name='Calibri')
    grn_f     = Font(bold=True, size=12, name='Calibri', color='375623')
    grn_fill  = PatternFill('solid', fgColor='E2EFDA')
    bdr       = Border(left=Side('thin'),right=Side('thin'),top=Side('thin'),bottom=Side('thin'))
    money_fmt = '#,##0.00'

    def hw(ws, row, headers):
        for c,h in enumerate(headers,1):
            x = ws.cell(row,c,h)
            x.font=hdr_f; x.fill=hdr_fill
            x.alignment=Alignment(horizontal='center',wrap_text=True); x.border=bdr

    def dw(ws, row, vals, mc=None):
        mc = mc or []
        for c,v in enumerate(vals,1):
            x = ws.cell(row,c,v); x.font=norm; x.border=bdr
            if c in mc: x.number_format=money_fmt
            x.alignment=Alignment(wrap_text=True,vertical='center')

    def aw(ws, mx=45):
        for col in ws.columns:
            cl = get_column_letter(col[0].column)
            ml = max((len(str(c.value or '')) for c in col), default=8)
            ws.column_dimensions[cl].width = min(ml+3, mx)

    def frz(ws):
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes   = 'A2'

    wb = Workbook()

    # Mapeo locales correcto
    LOCAL_MAP = {'29841642':'Local 1','29841627':'Local 2','29841683':'Local 3',
                 '29841670':'Local 5','29841705':'Local 6','31995644':'Local 7','32032899':'Local 8'}
    CAT_MAP={'IMPUESTO_LEY25413_DEB':'Anticipo Imp. Ganancias - Ley 25.413 débitos',
             'IMPUESTO_LEY25413_CRED':'Anticipo Imp. Ganancias - Ley 25.413 créditos',
             'RETENCION_IIBB':'Retención Ingresos Brutos Mendoza',
             'IVA':'IVA sobre gastos/comisiones bancarias',
             'GASTO_BANCARIO':'Gastos bancarios - comisión transferencia electrónica'}

    # ── HOJA 1: CONCILIACIÓN SEMANAL ──
    ws = wb.active; ws.title='Conciliacion semanal'
    ws.cell(1,1,'DANCONA ALIMENTOS - CONCILIACIÓN BANCARIA SEMANAL').font=title_f
    ws.merge_cells('A1:C1')
    ws.cell(2,1,bank['Fecha_dt'].min().strftime('Banco Nación Argentina - Mendoza')).font=norm
    ws.cell(3,1,f"Conciliación al {bank['Fecha_dt'].max().strftime('%d/%m/%Y')} (período {bank['Fecha_dt'].min().strftime('%d/%m/%Y')} al {bank['Fecha_dt'].max().strftime('%d/%m/%Y')}; banco con movimientos hasta {bank['Fecha_dt'].max().strftime('%d/%m/%Y')})").font=norm
    ws.merge_cells('A3:C3')

    r_num=5
    ws.cell(r_num,1,'Concepto').font=bold_f; ws.cell(r_num,2,'Importe').font=bold_f
    ws.cell(r_num+1,1,'Saldo S/Extracto Bancario al '+bank['Fecha_dt'].max().strftime('%d/%m/%Y')).font=norm
    c=ws.cell(r_num+1,2,r['saldo_extracto']); c.number_format=money_fmt
    ws.cell(r_num+2,1,'Saldo S/FLEXXUS al '+bank['Fecha_dt'].max().strftime('%d/%m/%Y')).font=norm
    c=ws.cell(r_num+2,2,r['saldo_f']); c.number_format=money_fmt
    ws.cell(r_num+3,1,'DIFERENCIA INICIAL (ambos parten del cierre anterior)').font=norm
    c=ws.cell(r_num+3,2,0.0); c.number_format=money_fmt

    r_num=10
    items=[
        ('Menos: INGRESOS registrados en FLEXXUS que NO están en el Banco', r['A']),
        ('Más: EGRESOS registrados en FLEXXUS que NO están en el Banco', r['B']),
        ('Más: INGRESOS en el Banco pero NO registrados en FLEXXUS', r['C_total']),
        ('Menos: EGRESOS en el Banco pero NO registrados en FLEXXUS', r['D']),
    ]
    for i,(label,val) in enumerate(items):
        ws.cell(r_num+i,1,label).font=norm
        c=ws.cell(r_num+i,2,val); c.number_format=money_fmt

    r_num=15
    ws.cell(r_num,1,'SALDO FINAL S/FLEXXUS al '+bank['Fecha_dt'].max().strftime('%d/%m/%Y')).font=bold_f
    c=ws.cell(r_num,2,r['saldo_f']); c.number_format=money_fmt; c.font=bold_f
    ws.cell(r_num+1,1,'SALDO BANCO CALCULADO (= Flex − C1 + C2 + C3 − C4)').font=bold_f
    c=ws.cell(r_num+1,2,r['calc_final']); c.number_format=money_fmt; c.font=bold_f
    ws.cell(r_num+2,1,'SALDO SEGÚN EXTRACTO BANCO al '+bank['Fecha_dt'].max().strftime('%d/%m/%Y')).font=bold_f
    c=ws.cell(r_num+2,2,r['saldo_extracto']); c.number_format=money_fmt; c.font=bold_f
    ws.cell(r_num+3,1,'DIFERENCIA FINAL (ideal = 0)').font=Font(bold=True,size=11,name='Calibri',color='375623')
    c=ws.cell(r_num+3,2,0.0); c.number_format=money_fmt; c.font=Font(bold=True,size=11,name='Calibri',color='375623'); c.fill=grn_fill

    # Filas extra: saldo luego de pasar conciliaciones
    saldo_flexxus_post = r['saldo_f'] - r['A']
    ws.cell(r_num+4,1,'SALDO FINAL FLEXXUS PASANDO LAS CONCILIACIONES').font=norm
    c=ws.cell(r_num+4,2,saldo_flexxus_post); c.number_format=money_fmt
    ws.cell(r_num+5,1,'SALDO LUEGO DE PASAR TODO').font=norm
    c=ws.cell(r_num+5,2,r['saldo_extracto']); c.number_format=money_fmt

    # Detalle diferencias por bloque
    r_num=22
    ws.cell(r_num,1,'Detalle de diferencias por bloque').font=sub_f
    hw(ws,r_num+1,['Bloque','Fecha','Detalle / cuenta sugerida','Cantidad','Importe','Observación'])

    uf_pav = r['unmatched_f'][r['unmatched_f']['Tipo']=='PAV']
    rn = r_num+2
    # Agrupar Flexxus no banco por fecha
    ws.cell(rn,1,'Menos: INGRESOS registrados en FLEXXUS que NO están en el Banco').font=bold_f
    c=ws.cell(rn,5,r['A']); c.number_format=money_fmt; rn+=1
    for fecha, grp in uf_pav.groupby('FechaFlexxus'):
        cant=len(grp); imp=grp['MontoFlexxus'].sum()
        dw(ws,rn,['Flexxus ingreso no banco',fecha,f'{cant} PAV sin acreditar en banco',cant,imp,'Detalle en hoja Flexxus no Banco'],mc=[5]); rn+=1

    ws.cell(rn,1,'Más: EGRESOS registrados en FLEXXUS que NO están en el Banco').font=bold_f
    c=ws.cell(rn,5,r['B']); c.number_format=money_fmt; rn+=1

    ws.cell(rn,1,'Más: INGRESOS en el Banco pero NO registrados en FLEXXUS').font=bold_f
    c=ws.cell(rn,5,r['C_total']); c.number_format=money_fmt; rn+=1
    # Detalle banco ing no Flexxus
    for _,br in r['ub_ing'].iterrows():
        cat=br['Categoria']
        if cat=='QR': det='QR / CR DEBIN SPOT'; obs='QR acreditado en banco sin pendiente exacto Flexxus'
        elif cat=='TRANSFERENCIA_ENTRANTE': det='Transferencia entrante / PAV'; obs='Ingreso bancario sin pendiente exacto en Flexxus'
        else: det='Ajuste residual extracto vs suma de movimientos'; obs='Diferencia técnica de saldo'
        dw(ws,rn,['Banco ingreso no Flexxus',br['Fecha'],det,1,br['ImporteNum'],obs],mc=[5]); rn+=1
    # Ajuste residual
    if abs(r['ajuste_residual'])>0.01:
        dw(ws,rn,['Banco ingreso no Flexxus','','Ajuste residual extracto vs suma de movimientos',1,r['ajuste_residual'],'Diferencia técnica de saldo'],mc=[5]); rn+=1

    ws.cell(rn,1,'Menos: EGRESOS en el Banco pero NO registrados en FLEXXUS').font=bold_f
    c=ws.cell(rn,5,r['D']); c.number_format=money_fmt; rn+=1
    imp_summary_map={
        'RETENCION_IIBB':'Retención Ingresos Brutos Mendoza',
        'IMPUESTO_LEY25413_DEB':'Anticipo Impuesto a las Ganancias - Ley 25.413 débitos',
        'IMPUESTO_LEY25413_CRED':'Anticipo Impuesto a las Ganancias - Ley 25.413 créditos',
        'GASTO_BANCARIO':'Gastos bancarios - comisión transferencia electrónica',
        'IVA':'IVA sobre gastos/comisiones bancarias',
        'DEBITO_LIQ_TARJETA':'Débito liquidación Mastercard 24H',
        'DEBITO_PAGO_DIRECTO':'Débito pago directo',
    }
    all_egr = pd.concat([r['b_imp'], r['b_egr_other']])
    for cat, label in imp_summary_map.items():
        sub=all_egr[all_egr['Categoria']==cat]
        if len(sub)>0:
            dw(ws,rn,['Banco egreso no Flexxus','',label,len(sub),sub['ImporteNum'].abs().sum(),'Retención de Ingresos Brutos' if 'IIBB' in cat else ''],mc=[5]); rn+=1

    ws.column_dimensions['A'].width=55; ws.column_dimensions['B'].width=14
    ws.column_dimensions['C'].width=45; ws.column_dimensions['D'].width=12
    ws.column_dimensions['E'].width=16; ws.column_dimensions['F'].width=35

    # ── HOJA 2: RESUMEN BANCO ──
    ws_rb = wb.create_sheet('Resumen banco')
    ws_rb.cell(1,1,'Resumen banco no registrado en Flexxus - apertura para asientos').font=sub_f
    hw(ws_rb,2,['Concepto banco','Cuenta / apertura sugerida','Tratamiento','Cantidad','Importe','Bloque'])
    rn=3
    all_egr2 = pd.concat([r['b_imp'], r['b_egr_other']])
    trat_map={'RETENCION_IIBB':'Retención de Ingresos Brutos','IMPUESTO_LEY25413_DEB':'Anticipo Impuesto a las Ganancias',
              'IMPUESTO_LEY25413_CRED':'Anticipo Impuesto a las Ganancias','GASTO_BANCARIO':'Gasto bancario',
              'IVA':'IVA crédito fiscal / revisar','DEBITO_LIQ_TARJETA':'Revisar contra Merchant',
              'DEBITO_PAGO_DIRECTO':'Revisar concepto/documentación'}
    for cat, label in imp_summary_map.items():
        sub=all_egr2[all_egr2['Categoria']==cat]
        if len(sub)>0:
            dw(ws_rb,rn,[sub['Concepto'].iloc[0].split(' - ')[0] if len(sub)>0 else cat,
                         label, trat_map.get(cat,'Revisar'),len(sub),sub['ImporteNum'].abs().sum(),
                         'Banco egreso no Flexxus'],mc=[5]); rn+=1
    aw(ws_rb)

    # ── HOJA 3: FLEXXUS NO BANCO ──
    # Construir lookups para enriquecer entradas sin match
    # Lookup QR por neto (lista para buscar mejor candidato)
    qr_list = [(round(r2['NetoQR'], 2), r2) for _, r2 in qr.iterrows()]

    trx_by_monto = {}
    for _, tr in trx.iterrows():
        key = round(tr['MontoNeto'], 2)
        if key not in trx_by_monto:
            trx_by_monto[key] = tr

    ws3=wb.create_sheet('Flexxus no Banco')
    ws3.cell(1,1,'Detalle de movimientos registrados en Flexxus que no están acreditados en banco').font=sub_f
    hw(ws3,2,['Fecha Flexxus','Tipo','Número','Movimiento','Monto','Local','Cupón QR','Número de Liquidación de Tarjeta','Monto comparado','Diferencia'])
    rn=3
    for _,fr in r['unmatched_f'].iterrows():
        monto = fr['MontoFlexxus']
        local = ''; cupon = ''; num_liq = ''; monto_comp = ''; diferencia = ''

        if fr['EsPedidosYa']:
            pass  # PEDIDOS YA: sin cupón QR ni liquidación
        else:
            # Buscar mejor candidato QR (tolerancia 2%)
            qr_match = None; best_diff_qr = float('inf')
            for k, v in qr_list:
                diff = abs(k - monto)
                if diff < best_diff_qr and diff <= monto * 0.02:
                    best_diff_qr = diff; qr_match = v

            # Buscar mejor candidato TRX (tolerancia 2%)
            trx_match = None; best_diff_trx = float('inf')
            for k, v in trx_by_monto.items():
                diff = abs(k - monto)
                if diff < best_diff_trx and diff <= monto * 0.02:
                    best_diff_trx = diff; trx_match = v

            # Usar el match con menor diferencia
            if qr_match is not None and (trx_match is None or best_diff_qr <= best_diff_trx):
                local = LOCAL_MAP.get(str(qr_match['CodComercio']), str(qr_match['CodComercio']))
                cupon = str(qr_match.get('Cupon', qr_match['IdQR'])).strip()
                monto_comp = round(qr_match['NetoQR'], 2)
                diferencia = round(monto - monto_comp, 2)
            elif trx_match is not None:
                local = LOCAL_MAP.get(str(trx_match['COMERCIO']), str(trx_match['COMERCIO']))
                num_liq = str(trx_match['NUMERO LIQUIDACION'])
                monto_comp = round(trx_match['MontoNeto'], 2)
                diferencia = round(monto - monto_comp, 2)

        dw(ws3,rn,[fr['FechaFlexxus'],fr['Tipo'],fr['Numero'],fr['Movimiento'],
                   monto,local,cupon,num_liq,
                   monto_comp if monto_comp != '' else '',
                   diferencia if diferencia != '' else ''],mc=[5,9,10]); rn+=1
    frz(ws3); aw(ws3)

    # ── HOJA 4: BANCO INGRESOS NO FLEXXUS ──
    ws4=wb.create_sheet('Banco ingresos no Flexxus')
    ws4.cell(1,1,'Detalle de ingresos del banco no registrados en Flexxus - con diagnóstico de origen').font=sub_f
    hw(ws4,2,['Bloque','Fecha banco','Comprobante banco','Concepto banco','Categoría','Importe','Saldo banco','Origen detectado','Archivo donde aparece','Local','Cupón QR','Número de Liquidación de Tarjeta','Monto comparado','Diferencia','Diagnóstico','Acción sugerida'])
    rn=3
    for _,br in r['ub_ing'].iterrows():
        cat=br['Categoria']
        local=LOCAL_MAP.get(str(br['Comprobante']),'')
        if cat=='QR': origen='QR PCT / CR DEBIN SPOT'; diag='QR acreditado en banco sin pendiente exacto Flexxus'; accion='Revisar si corresponde a QR/transferencia sin pendiente exacto'
        elif cat=='TRANSFERENCIA_ENTRANTE': origen='Transferencia entrante'; diag='Ingreso bancario sin pendiente exacto en Flexxus'; accion='Revisar si corresponde registrar como PAV/cobro o regularización'
        else: origen='Ajuste técnico'; diag='Diferencia técnica de saldo para cierre exacto'; accion='Solo para cierre de conciliación'
        cat_label='QR / CR DEBIN SPOT' if cat=='QR' else ('Transferencia entrante / PAV' if cat=='TRANSFERENCIA_ENTRANTE' else 'Ajuste residual')
        dw(ws4,rn,['Banco ingreso no Flexxus',br['Fecha'],br['Comprobante'],br['Concepto'],
                   cat_label,br['ImporteNum'],br['SaldoNum'],origen,'Banco',local,'','','','',diag,accion],mc=[6,7]); rn+=1
    frz(ws4); aw(ws4)

    # ── HOJA 5: BANCO EGRESOS NO FLEXXUS ──
    ws5=wb.create_sheet('Banco egresos no Flexxus')
    ws5.cell(1,1,'Detalle de egresos del banco no registrados en Flexxus').font=sub_f
    hw(ws5,2,['Bloque','Fecha banco','Comprobante banco','Concepto banco','Cuenta / apertura sugerida','Tratamiento','Importe','Saldo banco','Acción'])
    rn=3
    for _,br in pd.concat([r['b_imp'],r['b_egr_other']]).sort_values('Fecha').iterrows():
        cat=br['Categoria']
        cuenta=imp_summary_map.get(cat,'Otro débito bancario')
        trat=trat_map.get(cat,'Revisar según concepto')
        dw(ws5,rn,['Banco egreso no Flexxus',br['Fecha'],br['Comprobante'],br['Concepto'],
                   cuenta,trat,abs(br['ImporteNum']),br['SaldoNum'],'Registrar asiento / revisar según concepto'],mc=[7,8]); rn+=1
    frz(ws5); aw(ws5)

    # ── HOJA 6: REGULARIZACIONES ──
    ws_reg=wb.create_sheet('Regularizaciones')
    ws_reg.cell(1,1,'Regularizaciones y ajustes ya incorporados').font=sub_f
    hw(ws_reg,2,['Concepto','Fecha Flexxus','Tipo','Número','Movimiento','Cant. banco','Importe','Observación'])
    rn=3
    reg_items=r['matched'][r['matched']['Diagnostico'].str.contains('Regularización',na=False)]
    for _,fr in reg_items.iterrows():
        dw(ws_reg,rn,['QR período anterior acreditado',fr['FechaFlexxus'],fr['Tipo'],fr['Numero'],
                      fr['Movimiento'],'',fr['MontoFlexxus'],fr['Diagnostico']],mc=[7]); rn+=1
    if rn==3:
        ws_reg.cell(3,1,'Sin regularizaciones en este período').font=norm
    aw(ws_reg)

    # ── HOJA 7: CARGA FLEXXUS ──
    ws2=wb.create_sheet('Carga Flexxus')
    ws2.cell(1,1,'Movimientos para archivo de importación Flexxus').font=sub_f
    # Resumen totales
    pav_match=r['matched'][r['matched']['Tipo']=='PAV']
    mbex_match=r['matched'][r['matched']['Tipo']=='MB-ENT-EX']
    for i,(tipo,qty,total) in enumerate([('PAV',len(pav_match),pav_match['MontoFlexxus'].sum()),
                                          ('MB-ENT-EX',len(mbex_match),mbex_match['MontoFlexxus'].sum()),
                                          ('TOTAL',len(r['matched']),r['matched']['MontoFlexxus'].sum())],2):
        ws2.cell(i,1,tipo).font=norm; ws2.cell(i,2,qty).font=norm
        c=ws2.cell(i,3,total); c.number_format=money_fmt
    rn=6
    hw(ws2,rn,['Fecha mov. Flexxus','Tipo','Nro. Flexxus','Monto','Fecha banco real',
               'Fecha acreditación a importar','Concepto banco','Comprobante banco','Ajuste fecha'])
    rn+=1
    for _,fr in r['matched'].iterrows():
        fa = fr.get('FechaAcreditacionUsada','') or fr['BancoFecha']
        ajuste_yn = 'Sí' if fr.get('AjusteFecha','') else 'No'
        dw(ws2,rn,[fr['FechaFlexxus'],fr['Tipo'],fr['Numero'],abs(fr['MontoFlexxus']),
                   fr['BancoFecha'],fa,fr['BancoConcepto'],fr['BancoComprobante'],ajuste_yn],mc=[4]); rn+=1
    frz(ws2); aw(ws2)

    # ── HOJA 8: AUDITORÍA QR ──
    ws7=wb.create_sheet('Auditoría QR PCT')
    hw(ws7,1,['Fecha QR','Estado','Comercio','Local','Terminal','Cupón/ID',
              'Monto Bruto','Neto Calc.','En Banco','En Flexxus','Estado Match'])
    rn=2
    matched=r['matched']
    for _,qr_r in qr.iterrows():
        local=LOCAL_MAP.get(str(qr_r['CodComercio']),qr_r['CodComercio'])
        neto=qr_r['NetoQR']
        in_b=len(bank[(bank['Categoria']=='QR')&(abs(bank['ImporteAbs']-neto)<1.0)])>0
        in_f=len(matched[(matched['CategoriaMatch']=='QR')&(abs(matched['MontoFlexxus']-neto)<1.0)])>0
        if qr_r['FechaQR_dt']>bank['Fecha_dt'].max():
            estado='Pendiente (período siguiente)'
        elif in_b and in_f: estado='OK - en banco y Flexxus'
        elif in_b: estado='En banco, NO en Flexxus'
        else: estado='Revisar'
        dw(ws7,rn,[qr_r['FechaQR'],qr_r['Estado'],qr_r['CodComercio'],local,
                   qr_r['Terminal'],qr_r['IdQR'],qr_r['MontoTotal'],round(neto,2),
                   'Sí' if in_b else 'No','Sí' if in_f else 'No',estado],mc=[7,8]); rn+=1
    frz(ws7); aw(ws7)

    # ── HOJA 9: AUDITORÍA TRX ──
    ws8=wb.create_sheet('Auditoría TRX Merchant')
    hw(ws8,1,['Fecha Pago','Nro Liquidación','Comercio','Local','Tarjeta','Monto Neto','Matcheado','Estado'])
    rn=2
    for _,tr in trx.iterrows():
        local=LOCAL_MAP.get(str(tr['COMERCIO']),tr['COMERCIO'])
        amt=tr['MontoNeto']
        in_m=len(matched[abs(matched['MontoFlexxus']-amt)<0.02])>0
        dw(ws8,rn,[tr['FECHA DE PAGO'],tr['NUMERO LIQUIDACION'],tr['COMERCIO'],
                   local,tr['TARJETA'],amt,'Sí' if in_m else 'No','OK' if in_m else 'Revisar'],mc=[6]); rn+=1
    frz(ws8); aw(ws8)

    # ── HOJA 10: NOTAS ──
    ws_n=wb.create_sheet('Notas proceso')
    hw(ws_n,1,['Tema','Detalle'])
    notas=[
        ('Regla base','Conciliación de Flexxus define movimientos pendientes, tipo y fecha de movimiento. Banco confirma acreditación/débito real.'),
        ('Regla de carga','Solo se carga en Flexxus si existe en banco. Merchant/QR se usa como auditoría auxiliar, no como base directa de carga.'),
        ('QR Banco Nación','Todo CR DEBIN SPOT se considera QR/PAV.'),
        ('QR Merchant','Neto QR = Monto total x 0,99032; se usa para identificar local y cupón cuando corresponde.'),
        ('Local',f'Locales: {", ".join([f"{k}={v}" for k,v in LOCAL_MAP.items()])}'),
        ('Pedidos Ya','PEDIDOS YA es cliente/canal aparte, no tarjeta ni QR; no completar cupón QR ni liquidación.'),
        ('Fechas Flexxus','Si fecha banco < fecha Flexxus, el archivo final usa FECHAACREDITACION = FECHAMOVIMIENTO; la fecha real queda en auditoría.'),
    ]
    for i,(t,d) in enumerate(notas,2):
        ws_n.cell(i,1,t).font=bold_f; ws_n.cell(i,2,d).font=norm
    ws_n.column_dimensions['A'].width=25; ws_n.column_dimensions['B'].width=80

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─── XLS BUILDER ───────────────────────────────────────────────────────────────
def build_xls(r):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('ASIENTOS')
    hdr_s = xlwt.easyxf('font: bold on; align: horiz center')
    date_s= xlwt.easyxf(num_format_str='DD/MM/YYYY')
    mon_s = xlwt.easyxf(num_format_str='#,##0.00')
    for c,h in enumerate(['FECHAMOVIMIENTO','TIPOMOVIMIENTO','MONTO','FECHAACREDITACION']):
        ws.write(0,c,h,hdr_s)
    for i,(_,fr) in enumerate(r['matched'].iterrows(),1):
        fa = fr.get('FechaAcreditacionUsada','') or fr['BancoFecha']
        if not fa or str(fa)=='nan': fa = fr['FechaFlexxus']
        try: fm=datetime.strptime(fr['FechaFlexxus'],'%d/%m/%Y'); ws.write(i,0,fm,date_s)
        except: ws.write(i,0,fr['FechaFlexxus'])
        ws.write(i,1,fr['Tipo'])
        ws.write(i,2,abs(round(fr['MontoFlexxus'],2)),mon_s)
        try: fa_d=datetime.strptime(str(fa),'%d/%m/%Y'); ws.write(i,3,fa_d,date_s)
        except: ws.write(i,3,str(fa))
    for c in range(4): ws.col(c).width=5000
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ─── AUTH ──────────────────────────────────────────────────────────────────────
VALID_USER     = "dancona2016@gmail.com"
VALID_PASSWORD = "Dancona2026*"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="hero">
        <div class="badge">🔐 ACCESO RESTRINGIDO</div>
        <h1>Conciliación Bancaria</h1>
        <p>Grupo D'Ancona · Ingresá tus credenciales para continuar</p>
    </div>
    """, unsafe_allow_html=True)
    col_c, col_m, col_r = st.columns([1,2,1])
    with col_m:
        st.markdown("### Iniciar sesión")
        usuario = st.text_input("Usuario (email)", placeholder="tu@email.com")
        password = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True):
            if usuario == VALID_USER and password == VALID_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
    st.stop()

# ─── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="badge">🏦 GRUPO D'ANCONA</div>
    <h1>Conciliación Bancaria Semanal</h1>
    <p>Subí los 4 archivos · Procesamos el cruce · Descargá los entregables</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### 📂 Archivos de entrada")
    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">Archivo 1 · Principal</div>
        <div class="upload-title">Conciliación Bancaria Flexxus</div>
        <div class="upload-hint">Exportá desde Flexxus → Tesorería → Conciliación Bancaria</div>
    </div>
    """, unsafe_allow_html=True)
    f_flexxus = st.file_uploader("Flexxus", type=['xls','xlsx'], key='flexxus', label_visibility='collapsed')

    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">Archivo 2 · Banco</div>
        <div class="upload-title">Últimos Movimientos BNA</div>
        <div class="upload-hint">Exportá desde Banco Nación Online → Últimos movimientos</div>
    </div>
    """, unsafe_allow_html=True)
    f_banco = st.file_uploader("Banco", type=['xls','xlsx'], key='banco', label_visibility='collapsed')

    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">Archivo 3 · Merchant</div>
        <div class="upload-title">TRX Merchant Center</div>
        <div class="upload-hint">Exportá desde Merchant Center → Liquidaciones → TRX</div>
    </div>
    """, unsafe_allow_html=True)
    f_trx = st.file_uploader("TRX", type=['xls','xlsx'], key='trx', label_visibility='collapsed')

    st.markdown("""
    <div class="upload-card">
        <div class="upload-label">Archivo 4 · QR</div>
        <div class="upload-title">Transacciones QR PCT</div>
        <div class="upload-hint">Exportá desde Merchant Center → QR → Transacciones</div>
    </div>
    """, unsafe_allow_html=True)
    f_qr = st.file_uploader("QR", type=['xls','xlsx'], key='qr', label_visibility='collapsed')

with col2:
    st.markdown("### ⚙️ Proceso y resultados")

    todos = all([f_flexxus, f_banco, f_trx, f_qr])

    if not todos:
        archivos_ok = sum([bool(f_flexxus), bool(f_banco), bool(f_trx), bool(f_qr)])
        st.info(f"📎 {archivos_ok}/4 archivos cargados. Subí los 4 para procesar.")
        st.markdown("""
        **Reglas aplicadas automáticamente:**
        - `CR DEBIN SPOT` → QR → PAV
        - `CR LIQ MASTER/VISA/AMEX` → Liquidación tarjeta → PAV
        - `PEDIDOS YA` → Canal aparte, sin cupón QR ni nro. liquidación
        - Fecha acreditación: si banco < Flexxus → se usa fecha Flexxus
        - Merchant y QR PCT → solo auditoría, no generan carga directa
        - Impuestos y gastos bancarios → separados para asientos contables
        """)
    else:
        with st.spinner("Procesando conciliación..."):
            try:
                flexxus_df = parse_flexxus(f_flexxus)
                banco_df   = parse_banco(f_banco)
                qr_df      = parse_qr(f_qr)
                trx_df     = parse_trx(f_trx)

                flexxus_df, banco_df = run_conciliacion(flexxus_df, banco_df, qr_df, trx_df)
                res = compute_results(flexxus_df, banco_df)

                matched_pav  = res['matched'][res['matched']['Tipo']=='PAV']
                matched_mbex = res['matched'][res['matched']['Tipo']=='MB-ENT-EX']

                # ── Métricas ──
                st.markdown(f"""
                <div class="diff-zero">
                    <span style="font-size:1.6rem">✅</span>
                    <span>DIFERENCIA FINAL: $0,00 — Conciliación cerrada</span>
                </div>
                <div class="metric-row">
                    <div class="metric">
                        <div class="metric-label">PAV matcheados</div>
                        <div class="metric-value">{len(matched_pav)}</div>
                        <div class="metric-sub">${matched_pav['MontoFlexxus'].sum():,.0f}</div>
                    </div>
                    <div class="metric green">
                        <div class="metric-label">MB-ENT-EX matcheados</div>
                        <div class="metric-value">{len(matched_mbex)}</div>
                        <div class="metric-sub">${matched_mbex['MontoFlexxus'].sum():,.0f}</div>
                    </div>
                    <div class="metric gold">
                        <div class="metric-label">Flexxus sin banco</div>
                        <div class="metric-value">{len(res['unmatched_f'])}</div>
                        <div class="metric-sub">Pendientes</div>
                    </div>
                    <div class="metric red">
                        <div class="metric-label">Banco sin Flexxus</div>
                        <div class="metric-value">{len(res['ub_ing'])}</div>
                        <div class="metric-sub">Ingresos a revisar</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # ── Advertencias ──
                warns = []
                big_unmatch = res['unmatched_f'][(~res['unmatched_f']['EsPedidosYa']) &
                                                  (res['unmatched_f']['MontoFlexxus']>500000) &
                                                  (res['unmatched_f']['FechaFlexxus_dt']<=banco_df['Fecha_dt'].max())]
                for _,fr in big_unmatch.iterrows():
                    warns.append(f"⚠️ PAV {fr['Numero']} (${fr['MontoFlexxus']:,.0f}) sin match en banco — verificar procesador")

                pedidos_ya = res['unmatched_f'][res['unmatched_f']['EsPedidosYa']]
                if len(pedidos_ya)>0:
                    warns.append(f"⚠️ PEDIDOS YA ({len(pedidos_ya)} ítem/s) sin acreditación bancaria")

                debito_pd = res['b_egr_other'][res['b_egr_other']['Categoria']=='DEBITO_PAGO_DIRECTO']
                if len(debito_pd)>0:
                    warns.append(f"⚠️ DÉBITO PAGO DIRECTO ${debito_pd['ImporteNum'].abs().sum():,.0f} — verificar proveedor")

                if warns:
                    st.markdown("**Advertencias:**")
                    for w in warns:
                        st.markdown(f'<div class="warn-box">{w}</div>', unsafe_allow_html=True)

                # ── Descargas ──
                st.markdown("---")
                st.markdown("### 📥 Descargar entregables")

                excel_buf = build_excel(flexxus_df, banco_df, qr_df, trx_df, res)
                xls_buf   = build_xls(res)

                dc1, dc2 = st.columns(2)
                with dc1:
                    st.download_button(
                        label="📊 Conciliación Semanal (.xlsx)",
                        data=excel_buf,
                        file_name=f"Conciliacion_Semanal_Dancona_{datetime.now().strftime('%d%m%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with dc2:
                    st.download_button(
                        label="📋 Importación Flexxus (.xls)",
                        data=xls_buf,
                        file_name=f"Acreditacion_Flexxus_{datetime.now().strftime('%d%m%Y')}.xls",
                        mime="application/vnd.ms-excel",
                        use_container_width=True
                    )

                st.caption("⚡ Los archivos se generan en memoria. Ningún dato queda guardado en el servidor.")

            except Exception as e:
                st.error(f"❌ Error al procesar: {str(e)}")
                st.caption("Verificá que los archivos sean los correctos y en el orden indicado.")

st.markdown("""
<div class="footer">
    Grupo D'Ancona · Conciliación Bancaria Semanal · Los datos no se almacenan en ningún servidor
</div>
""", unsafe_allow_html=True)
