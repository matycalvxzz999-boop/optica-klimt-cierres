import io
import os
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from extractor import extraer_cierre_gemini, CierreDia
import pandas as pd
import altair as alt

load_dotenv()
APP_PASSWORD = os.getenv("APP_PASSWORD", "optica2026")
ARCHIVO_MAESTRO = "cierre_mensual_optica.xlsx"

st.set_page_config(
    page_title="Optica Klimt - Tablero de Comando Gerencial",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS PROFESIONALES ---
st.markdown("""
    <style>
    .main {
        background-color: #F8F9FA;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 500;
    }
    .stDownloadButton>button {
        border-radius: 6px;
        font-weight: 500;
        background-color: #1F4E78;
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.5rem;
        color: #1F4E78;
    }
    </style>
""", unsafe_allow_html=True)

# --- ESTILOS EXCEL ---
BORDE_SUAVE = Border(
    bottom=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
)

BORDE_TOTAL = Border(
    top=Side(style="thin", color="1F4E78"),
    bottom=Side(style="double", color="1F4E78"),
)

COLUMNAS_EXCEL = [
    "Fecha",
    "Efectivo",
    "MP Neto",
    "MP IVA (21%)",
    "BBVA Neto",
    "BBVA IVA (21%)",
    "Provincia",
    "Gastos",
    "Celia / Toti",
    "Observaciones / Detalle",
]

def estilizar_encabezado(ws, row_idx, num_cols, color_hex="1F4E78"):
    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
    font = Font(color="FFFFFF", bold=True, size=11)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE_SUAVE

def ajustar_columnas(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.row == 1:
                continue
            val = str(cell.value or "")
            if val.startswith("="):
                val = "$999,999.00"
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

def procesar_datos_dia(res: CierreDia) -> dict:
    efectivo_ventas = 0.0
    mp_total = 0.0
    bbva_total = 0.0
    provincia_ventas = 0.0

    for op in res.operaciones:
        monto = float(op.monto or 0)
        if op.medio_pago == "qr_mp":
            mp_total += monto
        elif op.medio_pago == "payway_bbva":
            bbva_total += monto
        elif op.medio_pago == "posnet_provincia":
            provincia_ventas += monto
        elif op.medio_pago == "efectivo":
            efectivo_ventas += monto
        else:
            bbva_total += monto

    if efectivo_ventas == 0 and res.total_efectivo_cuaderno:
        efectivo_ventas = float(res.total_efectivo_cuaderno)
    if bbva_total == 0 and res.total_tarjetas_cuaderno:
        bbva_total = float(res.total_tarjetas_cuaderno)

    # Discriminación exacta de IVA redondeada a 2 decimales para evitar decimales infinitos
    mp_neto = round(mp_total / 1.21, 2)
    mp_iva = round(mp_total - mp_neto, 2)

    bbva_neto = round(bbva_total / 1.21, 2)
    bbva_iva = round(bbva_total - bbva_neto, 2)

    gastos_total = 0.0
    celia_toti_total = 0.0
    detalles_gastos = []

    for g in res.gastos_retiros:
        monto_g = float(g.monto or 0)
        concepto_lower = g.concepto.lower()
        tipo_lower = g.tipo.lower()

        if "cel" in concepto_lower or "toti" in concepto_lower or "celia" in concepto_lower or tipo_lower == "retiro":
            celia_toti_total += monto_g
            detalles_gastos.append(f"{g.concepto} (${monto_g:,.0f})")
        else:
            gastos_total += monto_g
            detalles_gastos.append(f"{g.concepto} (${monto_g:,.0
