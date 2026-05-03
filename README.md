# Conciliación Bancaria Dancona V5.9.13

App Streamlit para conciliación bancaria semanal con ledger QR persistente embebido.

## Archivos del paquete (subir a Streamlit)

- `app.py` — entrypoint Streamlit
- `qr_ledger.py` — motor del ledger QR
- `qr_ledger_regla3_v2.py` — regla 3 (equivalencia QR histórica)
- `requirements.txt`

## Workflow semanal (igual que V5.9.11)

Subir los 5 archivos de siempre:
1. Conciliación bancaria Flexxus actual
2. Últimos movimientos Banco Nación
3. TRX Merchant Center
4. Transacciones QR acumuladas
5. Conciliación anterior del sistema

Apretar "Comenzar conciliación".

El ledger QR persistente viaja **embebido** en una hoja oculta `_QR_LEDGER`
del Excel de conciliación anterior (archivo 5). Se carga automáticamente
cuando subís el archivo. La conciliación nueva que descargás trae el
ledger actualizado embebido para la próxima corrida.

## Qué cambia respecto a V5.9.11

1. La hoja "Banco ingresos no Flexxus" trae los pendientes igual que antes.
2. **Los pendientes que tienen cupón QR histórico identificado** (gracias al
   ledger acumulado) llevan en su columna "Diagnóstico" el detalle:
   `REGLA3: Identificado con cupón histórico XXX (Local Y, fecha Z, comp W)`.
   El conciliador ve cuáles son seguros sin investigar.
3. Hojas nuevas en el Excel:
   - `Regularizaciones Ledger QR` — auditoría de qué se identificó por regla 3
   - `Ledger QR snapshot` — foto del ledger persistente al cierre
   - `_QR_LEDGER` (oculta) — el ledger en JSON para la próxima corrida

## Primera corrida en producción

Si subís un Excel del sistema viejo V5.9.11 que NO tiene el ledger embebido,
la app corre en modo V5.9.11 puro (sin regla 3 esa primera vez). El Excel
descargado SÍ trae el ledger embebido para que las próximas corridas tengan
todo el contexto histórico.

## Cómo correr local

```bash
pip install -r requirements.txt
streamlit run app.py
```
