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
            detalles_gastos.append(f"{g.concepto} (${monto_g:,.0f})")

    if gastos_total == 0 and celia_toti_total == 0 and res.total_gastos_cuaderno:
        gastos_total = float(res.total_gastos_cuaderno)

    if res.observaciones:
        detalles_gastos.append(res.observaciones)

    return {
        "efectivo": efectivo_ventas,
        "mp_neto": mp_neto,
        "mp_iva": mp_iva,
        "bbva_neto": bbva_neto,
        "bbva_iva": bbva_iva,
        "provincia": provincia_ventas,
        "gastos": gastos_total,
        "celia_toti": celia_toti_total,
        "detalle_obs": " | ".join(detalles_gastos),
    }

def construir_fila_dia(res: CierreDia) -> list:
    fecha_str = str(res.fecha) if res.fecha else "Sin fecha"
    d = procesar_datos_dia(res)
    return [
        fecha_str,
        d["efectivo"],
        d["mp_neto"],
        d["mp_iva"],
        d["bbva_neto"],
        d["bbva_iva"],
        d["provincia"],
        d["gastos"],
        d["celia_toti"],
        d["detalle_obs"],
    ]

def crear_libro_base(filas_datos: list) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cierre Mensual"

    ws["A1"] = "OPTICA KLIMT - CIERRE DE CAJA MENSUAL (CON IVA Y TRIBUTOS)"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")

    ws.append([])
    ws.append(COLUMNAS_EXCEL)
    estilizar_encabezado(ws, 3, len(COLUMNAS_EXCEL), "1F4E78")

    filas_ordenadas = sorted(filas_datos, key=lambda x: str(x[0]))

    for fila in filas_ordenadas:
        ws.append(fila)
        r = ws.max_row
        for col in range(1, len(COLUMNAS_EXCEL) + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDE_SUAVE
            if 2 <= col <= 9:
                cell.number_format = "$#,##0.00"
                if col in (2, 4, 6, 8, 9):
                    cell.font = Font(bold=True)

    if filas_ordenadas:
        fila_inicio = 4
        fila_fin = ws.max_row
        fila_tot = fila_fin + 1
        ws.cell(row=fila_tot, column=1, value="TOTAL MES").font = Font(bold=True, size=11, color="1F4E78")
        ws.cell(row=fila_tot, column=1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.cell(row=fila_tot, column=1).border = BORDE_TOTAL

        letras = ["B", "C", "D", "E", "F", "G", "H", "I"]
        for idx, letra in enumerate(letras, start=2):
            c_tot = ws.cell(row=fila_tot, column=idx, value=f"=SUM({letra}{fila_inicio}:{letra}{fila_fin})")
            c_tot.font = Font(bold=True, size=11, color="1F4E78")
            c_tot.number_format = "$#,##0.00"
            c_tot.border = BORDE_TOTAL
            c_tot.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ajustar_columnas(ws)
    return wb

def asegurar_archivo_maestro():
    if not os.path.exists(ARCHIVO_MAESTRO):
        wb_vacio = crear_libro_base([])
        wb_vacio.save(ARCHIVO_MAESTRO)

def generar_excel_solo_dia_bytes(res: CierreDia) -> bytes:
    wb = crear_libro_base([construir_fila_dia(res)])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def guardar_y_ordenar_en_mensual(res: CierreDia):
    nueva_fila = construir_fila_dia(res)
    fecha_nueva = nueva_fila[0]
    filas_existentes = []

    if os.path.exists(ARCHIVO_MAESTRO):
        try:
            wb_viejo = load_workbook(ARCHIVO_MAESTRO)
            ws_viejo = wb_viejo.active
            for row in ws_viejo.iter_rows(min_row=4, values_only=True):
                if not row or not row[0] or str(row[0]).startswith("TOTAL") or str(row[0]).startswith("OPTICA"):
                    continue
                if str(row[0]) != fecha_nueva:
                    filas_existentes.append(list(row[:len(COLUMNAS_EXCEL)]))
        except Exception:
            pass

    filas_existentes.append(nueva_fila)
    wb_nuevo = crear_libro_base(filas_existentes)
    wb_nuevo.save(ARCHIVO_MAESTRO)

asegurar_archivo_maestro()

# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

app_password = st.secrets.get("APP_PASSWORD", APP_PASSWORD)

if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)
        elif os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=180)
        
        st.markdown("## Acceso Restringido - Optica Klimt")
        st.write("Ingrese la contraseña institucional para continuar.")
        pwd = st.text_input("Contraseña:", type="password")
        if st.button("Ingresar", type="primary", use_container_width=True):
            if pwd == app_password:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
    st.stop()

# --- MENÚ LATERAL ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=200)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=200)
    else:
        st.markdown("### Optica Klimt")
    
    st.write("Tablero de Comando Gerencial")
    st.divider()
    
    menu_opcion = st.radio(
        "Navegación",
        ["Tablero de Comando (Resumen)", "Procesar Cierre Diario"],
        label_visibility="collapsed"
    )
    
    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()
