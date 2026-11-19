import io
import os
import re
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from extractor import CierreDia, GastoRetiro, Operacion, extraer_cierre_gemini

load_dotenv()

APP_PASSWORD = os.getenv("APP_PASSWORD", "optica2026")
ARCHIVO_MAESTRO = "cierre_mensual_optica.xlsx"

st.set_page_config(
    page_title="Óptica Klimt - Tablero de Comando Gerencial",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- ESTILOS UI ---
st.markdown(
    """
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
        font-weight: 600;
        background-color: #1F4E78;
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.45rem;
        color: #1F4E78;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- CONSTANTES DE DISEÑO EXCEL ---
COLOR_PRIMARIO = "1F4E78"
COLOR_SECUNDARIO = "D9E1F2"
COLOR_ZEBRA = "F2F5F9"
COLOR_VERDE_BG = "E2EFDA"
COLOR_VERDE_TXT = "274E13"
COLOR_ROJO_BG = "FCE4D6"
COLOR_ROJO_TXT = "9C0006"

FORMATO_MONEDA_ARS = '$ #,##0.00;-$ #,##0.00;"$ 0.00"'

BORDE_SUAVE = Border(
    bottom=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
)

BORDE_TOTAL = Border(
    top=Side(style="thin", color=COLOR_PRIMARIO),
    bottom=Side(style="double", color=COLOR_PRIMARIO),
)

COLUMNAS_HISTORICO = [
    "Fecha",
    "Efectivo",
    "MP Neto",
    "MP IVA (21%)",
    "BBVA Neto",
    "BBVA IVA (21%)",
    "Provincia",
    "Gastos",
    "Celia / Toti",
    "Diferencia Caja",
    "Observaciones / Detalle",
]

ETIQUETAS_MEDIO_PAGO = {
    "efectivo": "Efectivo",
    "qr_mp": "Mercado Pago (QR / Débito / Transf.)",
    "payway_bbva": "Payway BBVA",
    "posnet_provincia": "Posnet Provincia",
    "transferencia": "Transferencia Bancaria",
    "otro": "Otro Medio",
}

DENOMINACIONES_ARS = [20000, 10000, 2000, 1000, 500, 200, 100, 50, 20, 10]


# --- UTILIDADES FINANCIERAS EXACTAS (DECIMAL) ---
def to_dec(val: Decimal | float | int | str | None) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    if isinstance(val, Decimal):
        return val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    try:
        if isinstance(val, float) and pd.isna(val):
            return Decimal("0.00")
        return Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0.00")


def desglosar_iva_21(bruto: Decimal) -> tuple[Decimal, Decimal]:
    """Aplica divisor estricto 1.21 con redondeo contable ROUND_HALF_UP."""
    bruto = to_dec(bruto)
    if bruto == Decimal("0.00"):
        return Decimal("0.00"), Decimal("0.00")
    neto = (bruto / Decimal("1.21")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    iva = (bruto - neto).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return neto, iva


def sanitizar_nombre_archivo(texto: str) -> str:
    limpio = re.sub(r"[^A-Za-z0-9_-]", "_", texto.strip())
    return re.sub(r"_+", "_", limpio).strip("_") or "SinDato"


# --- CÁLCULOS DE CIERRE ---
def procesar_datos_dia(res: CierreDia) -> dict:
    efectivo_ventas = Decimal("0.00")
    mp_total = Decimal("0.00")
    bbva_total = Decimal("0.00")
    provincia_ventas = Decimal("0.00")
    transf_ventas = Decimal("0.00")

    for op in res.operaciones:
        monto = to_dec(op.monto)
        if op.medio_pago == "qr_mp":
            mp_total += monto
        elif op.medio_pago == "payway_bbva":
            bbva_total += monto
        elif op.medio_pago == "posnet_provincia":
            provincia_ventas += monto
        elif op.medio_pago == "efectivo":
            efectivo_ventas += monto
        elif op.medio_pago == "transferencia":
            transf_ventas += monto
        else:
            bbva_total += monto

    # Respaldo si la planilla solo tenía subtotales escritos sin detalle línea a línea
    if efectivo_ventas == Decimal("0.00") and res.total_efectivo_cuaderno:
        efectivo_ventas = to_dec(res.total_efectivo_cuaderno)
    if (
        mp_total == Decimal("0.00")
        and bbva_total == Decimal("0.00")
        and provincia_ventas == Decimal("0.00")
        and res.total_tarjetas_cuaderno
    ):
        bbva_total = to_dec(res.total_tarjetas_cuaderno)

    mp_neto, mp_iva = desglosar_iva_21(mp_total)
    bbva_neto, bbva_iva = desglosar_iva_21(bbva_total + transf_ventas)

    gastos_total = Decimal("0.00")
    celia_toti_total = Decimal("0.00")
    detalles_gastos: list[str] = []

    for g in res.gastos_retiros:
        monto_g = to_dec(g.monto)
        concepto_lower = g.concepto.lower()
        tipo_lower = g.tipo.lower()

        if (
            "cel" in concepto_lower
            or "toti" in concepto_lower
            or "celia" in concepto_lower
            or tipo_lower == "retiro"
        ):
            celia_toti_total += monto_g
            detalles_gastos.append(f"[Retiro] {g.concepto} (${monto_g:,.2f})")
        else:
            gastos_total += monto_g
            detalles_gastos.append(f"[Gasto] {g.concepto} (${monto_g:,.2f})")

    if gastos_total == Decimal("0.00") and celia_toti_total == Decimal("0.00") and res.total_gastos_cuaderno:
        gastos_total = to_dec(res.total_gastos_cuaderno)

    if res.observaciones:
        detalles_gastos.append(res.observaciones)

    saldo_inicial = to_dec(res.saldo_inicial)
    total_ingresos_brutos = efectivo_ventas + mp_total + bbva_total + transf_ventas + provincia_ventas
    total_egresos = gastos_total + celia_toti_total

    # Efectivo físico esperado en caja = Saldo inicial + Ventas efectivo - Gastos - Retiros
    efectivo_esperado_caja = saldo_inicial + efectivo_ventas - total_egresos
    # Total general esperado (incluyendo saldos electrónicos + efectivo en caja)
    total_general_esperado = saldo_inicial + total_ingresos_brutos - total_egresos

    # Si Celia anotó "CAJA TOTAL" en el cuaderno, se compara contra el total general esperado; si no, contra efectivo esperado
    caja_cuaderno = to_dec(res.caja_total_cuaderno)
    diferencia_cuaderno = (
        caja_cuaderno - total_general_esperado if res.caja_total_cuaderno is not None else Decimal("0.00")
    )

    return {
        "saldo_inicial": saldo_inicial,
        "efectivo": efectivo_ventas,
        "mp_bruto": mp_total,
        "mp_neto": mp_neto,
        "mp_iva": mp_iva,
        "bbva_bruto": bbva_total + transf_ventas,
        "bbva_neto": bbva_neto,
        "bbva_iva": bbva_iva,
        "provincia": provincia_ventas,
        "gastos": gastos_total,
        "celia_toti": celia_toti_total,
        "total_egresos": total_egresos,
        "total_ingresos_brutos": total_ingresos_brutos,
        "efectivo_esperado_caja": efectivo_esperado_caja,
        "total_general_esperado": total_general_esperado,
        "caja_cuaderno": caja_cuaderno,
        "diferencia_cuaderno": diferencia_cuaderno,
        "detalle_obs": " | ".join(detalles_gastos),
    }


def construir_fila_historico(res: CierreDia) -> list:
    fecha_str = res.fecha.strftime("%Y-%m-%d") if res.fecha else datetime.now().strftime("%Y-%m-%d")
    d = procesar_datos_dia(res)
    return [
        fecha_str,
        float(d["efectivo"]),
        float(d["mp_neto"]),
        float(d["mp_iva"]),
        float(d["bbva_neto"]),
        float(d["bbva_iva"]),
        float(d["provincia"]),
        float(d["gastos"]),
        float(d["celia_toti"]),
        float(d["diferencia_cuaderno"]),
        d["detalle_obs"],
    ]


# --- MOTOR DE EXCEL PROFESIONAL ---
def _configurar_impresion(ws, landscape: bool = True):
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if landscape else ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.center.text = "ÓPTICA KLIMT - AUDITORÍA DE CIERRE DE CAJA"
    ws.oddHeader.center.size = 10
    ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws.oddFooter.right.text = "Página &P de &N"


def _estilizar_encabezado(ws, row_idx: int, num_cols: int, color_hex: str = COLOR_PRIMARIO):
    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
    font = Font(color="FFFFFF", bold=True, size=11)
    ws.row_dimensions[row_idx].height = 26
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE_SUAVE


def _ajustar_columnas(ws, min_width: int = 14, max_width: int = 45):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            val = str(cell.value or "")
            if val.startswith("="):
                val = "$ 999,999.00"
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, min_width), max_width)


def _construir_hoja_historico(ws, filas_datos: list):
    ws.title = "Historico_Mensual"
    ws["A1"] = "ÓPTICA KLIMT - CONSOLIDADO MENSUAL DE CAJA (CON DISCRIMINACIÓN DE IVA 21%)"
    ws["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    ws.row_dimensions[1].height = 24

    ws.append([])
    ws.append(COLUMNAS_HISTORICO)
    _estilizar_encabezado(ws, 3, len(COLUMNAS_HISTORICO))
    ws.freeze_panes = "A4"

    filas_ordenadas = sorted(filas_datos, key=lambda x: str(x[0]))

    for idx_f, fila in enumerate(filas_ordenadas):
        ws.append(fila)
        r = ws.max_row
        fill_zebra = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
        for col in range(1, len(COLUMNAS_HISTORICO) + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDE_SUAVE
            if idx_f % 2 == 1:
                cell.fill = fill_zebra
            if col == 1:
                cell.alignment = Alignment(horizontal="center")
            elif 2 <= col <= 10:
                cell.number_format = FORMATO_MONEDA_ARS
                if col in (2, 4, 6, 8, 9, 10):
                    cell.font = Font(bold=True)
            else:
                cell.alignment = Alignment(wrap_text=True)

    if filas_ordenadas:
        fila_inicio = 4
        fila_fin = ws.max_row
        fila_tot = fila_fin + 1
        ws.row_dimensions[fila_tot].height = 24
        c_lbl = ws.cell(row=fila_tot, column=1, value="TOTAL MES")
        c_lbl.font = Font(bold=True, size=11, color=COLOR_PRIMARIO)
        c_lbl.fill = PatternFill(start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid")
        c_lbl.border = BORDE_TOTAL
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in range(2, 11):
            letra = get_column_letter(col_idx)
            c_tot = ws.cell(row=fila_tot, column=col_idx, value=f"=SUM({letra}{fila_inicio}:{letra}{fila_fin})")
            c_tot.font = Font(bold=True, size=11, color=COLOR_PRIMARIO)
            c_tot.number_format = FORMATO_MONEDA_ARS
            c_tot.border = BORDE_TOTAL
            c_tot.fill = PatternFill(start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid")

    _ajustar_columnas(ws)
    _configurar_impresion(ws, landscape=True)


def generar_excel_completo_cierre_bytes(
    res: CierreDia,
    arqueo_conteo: dict[int, int] | None = None,
    supervisor: str = "Martín",
    comentario_auditor: str = "",
) -> bytes:
    """
    Genera UN SOLO archivo .xlsx auditable con 5 hojas conectadas mediante fórmulas reales:
    1. Resumen
    2. Ingresos
    3. Egresos
    4. Arqueo_Denominaciones
    5. Diferencias_Observaciones
    """
    wb = Workbook()
    d = procesar_datos_dia(res)
    fecha_display = res.fecha.strftime("%d/%m/%Y") if res.fecha else datetime.now().strftime("%d/%m/%Y")
    turno_display = res.turno or "Único"
    cajero_display = res.cajero or "Celia"

    # ---------------------------------------------------------
    # HOJA 1: RESUMEN EJECUTIVO (Con fórmulas vinculadas a hojas 2, 3 y 4)
    # ---------------------------------------------------------
    ws_res = wb.active
    ws_res.title = "Resumen"
    ws_res["A1"] = "ÓPTICA KLIMT — ACTA DE CIERRE DE CAJA DIARIO"
    ws_res["A1"].font = Font(bold=True, size=15, color=COLOR_PRIMARIO)
    ws_res.row_dimensions[1].height = 28

    meta = [
        ("Fecha Contable:", fecha_display, "Cajero / Responsable:", cajero_display),
        ("Turno:", turno_display, "Supervisor:", supervisor),
        ("Fecha Generación:", datetime.now().strftime("%d/%m/%Y %H:%M"), "Estado:", "CONFIRMADO / AUDITADO"),
    ]
    for idx, (k1, v1, k2, v2) in enumerate(meta, start=3):
        ws_res.cell(row=idx, column=1, value=k1).font = Font(bold=True, color=COLOR_PRIMARIO)
        ws_res.cell(row=idx, column=2, value=v1)
        ws_res.cell(row=idx, column=3, value=k2).font = Font(bold=True, color=COLOR_PRIMARIO)
        ws_res.cell(row=idx, column=4, value=v2)
        for c in range(1, 5):
            ws_res.cell(row=idx, column=c).border = BORDE_SUAVE
            ws_res.cell(row=idx, column=c).fill = PatternFill(
                start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid"
            )

    ws_res.cell(row=7, column=1, value="CONCEPTO CONTABLE").font = Font(bold=True, color="FFFFFF")
    ws_res.cell(row=7, column=2, value="IMPORTE (ARS)").font = Font(bold=True, color="FFFFFF")
    ws_res.cell(row=7, column=3, value="ORIGEN / FÓRMULA AUDITORÍA").font = Font(bold=True, color="FFFFFF")
    _estilizar_encabezado(ws_res, 7, 3)
    ws_res.freeze_panes = "A8"

    # Filas dinámicas conectadas a Ingresos, Egresos y Arqueo
    num_ops = max(len(res.operaciones), 1)
    fila_tot_ingresos = 4 + num_ops
    num_egr = max(len(res.gastos_retiros), 1)
    fila_tot_egresos = 4 + num_egr
    fila_tot_arqueo = 4 + len(DENOMINACIONES_ARS)

    filas_resumen = [
        ("1. Fondo Inicial de Caja", float(d["saldo_inicial"]), "Saldo inicial declarado en libreta"),
        ("2. Total Ingresos Brutos (Todos los medios)", f"=Ingresos!E{fila_tot_ingresos}", "Suma hoja Ingresos"),
        ("   - Ventas en Efectivo", f"=SUMIF(Ingresos!C4:C{fila_tot_ingresos-1},\"Efectivo\",Ingresos!E4:E{fila_tot_ingresos-1})", "Subtotal Efectivo"),
        ("   - Ventas Mercado Pago (Bruto)", f"=SUMIF(Ingresos!C4:C{fila_tot_ingresos-1},\"Mercado Pago*\",Ingresos!E4:E{fila_tot_ingresos-1})", "Incluye IVA 21%"),
        ("   - Ventas Payway BBVA (Bruto)", f"=SUMIF(Ingresos!C4:C{fila_tot_ingresos-1},\"Payway BBVA\",Ingresos!E4:E{fila_tot_ingresos-1})", "Incluye IVA 21%"),
        ("   - Ventas Posnet Provincia", f"=SUMIF(Ingresos!C4:C{fila_tot_ingresos-1},\"Posnet Provincia\",Ingresos!E4:E{fila_tot_ingresos-1})", "Posnet Banco Provincia"),
        ("3. Total IVA Discriminado (21% sobre MP y BBVA)", f"=Ingresos!G{fila_tot_ingresos}", "IVA Crédito/Débito Fiscal"),
        ("4. Total Egresos (Gastos y Retiros)", f"=Egresos!D{fila_tot_egresos}", "Suma hoja Egresos"),
        ("5. Total Esperado en Efectivo (Caja Física)", "=B8+B10-B15", "Fondo Inicial + Efectivo - Egresos"),
        ("6. Total Contado en Arqueo de Billetes", f"=Arqueo_Denominaciones!C{fila_tot_arqueo}", "Suma hoja Arqueo"),
        ("7. Diferencia de Arqueo (Contado - Esperado)", "=B17-B16", "Positivo = Sobrante | Negativo = Faltante"),
        ("8. Total Declarado en Cuaderno (Cierre Celia)", float(d["caja_cuaderno"]), "Valor total anotado en cuaderno"),
        ("9. Total General Esperado (Efectivo + Tarjetas/QR)", "=B8+B9-B15", "Fondo Inicial + Ingresos Totales - Egresos"),
        ("10. Diferencia vs. Cuaderno", "=B19-B20", "Comparación cuaderno vs sistema"),
    ]

    for offset, (concepto, formula_o_val, nota) in enumerate(filas_resumen):
        r = 8 + offset
        c1 = ws_res.cell(row=r, column=1, value=concepto)
        c2 = ws_res.cell(row=r, column=2, value=formula_o_val)
        c3 = ws_res.cell(row=r, column=3, value=nota)

        c1.border = BORDE_SUAVE
        c2.border = BORDE_SUAVE
        c3.border = BORDE_SUAVE
        c2.number_format = FORMATO_MONEDA_ARS
        c3.font = Font(italic=True, color="595959", size=10)

        if concepto.startswith(("1.", "2.", "4.", "5.", "6.", "8.", "9.")):
            c1.font = Font(bold=True)
            c2.font = Font(bold=True)

        # Destacado visual de diferencias (filas 18 y 21)
        if concepto.startswith(("7.", "10.")):
            dif_ref = (
                (to_dec(sum(k * v for k, v in (arqueo_conteo or {}).items())) - d["efectivo_esperado_caja"])
                if concepto.startswith("7.")
                else d["diferencia_cuaderno"]
            )
            bg = COLOR_ROJO_BG if dif_ref < Decimal("0.00") else COLOR_VERDE_BG
            fg = COLOR_ROJO_TXT if dif_ref < Decimal("0.00") else COLOR_VERDE_TXT
            fill_dif = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            for col_i in (1, 2, 3):
                cell_d = ws_res.cell(row=r, column=col_i)
                cell_d.fill = fill_dif
                cell_d.font = Font(bold=True, color=fg, size=11)
                cell_d.border = BORDE_TOTAL

    _ajustar_columnas(ws_res, min_width=18, max_width=52)
    _configurar_impresion(ws_res, landscape=False)

    # ---------------------------------------------------------
    # HOJA 2: INGRESOS (Con fórmulas de Neto e IVA 21%)
    # ---------------------------------------------------------
    ws_ing = wb.create_sheet(title="Ingresos")
    ws_ing["A1"] = f"DETALLE DE INGRESOS — FECHA: {fecha_display} — TURNO: {turno_display}"
    ws_ing["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)

    headers_ing = ["#", "Detalle / Producto", "Medio de Pago", "Confianza IA", "Importe Bruto", "Neto Gravado", "IVA (21%)"]
    ws_ing.append([])
    ws_ing.append(headers_ing)
    _estilizar_encabezado(ws_ing, 3, len(headers_ing))
    ws_ing.freeze_panes = "A4"

    ops_a_volcar = res.operaciones if res.operaciones else [
        Operacion(detalle="Sin movimientos registrados", monto=Decimal("0.00"), medio_pago="efectivo", confianza="alta")
    ]

    for idx, op in enumerate(ops_a_volcar, start=1):
        r = 3 + idx
        medio_txt = ETIQUETAS_MEDIO_PAGO.get(op.medio_pago, op.medio_pago)
        aplica_iva = op.medio_pago in ("qr_mp", "payway_bbva", "transferencia")

        ws_ing.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws_ing.cell(row=r, column=2, value=op.detalle)
        ws_ing.cell(row=r, column=3, value=medio_txt)
        ws_ing.cell(row=r, column=4, value=op.confianza.upper()).alignment = Alignment(horizontal="center")
        ws_ing.cell(row=r, column=5, value=float(to_dec(op.monto))).number_format = FORMATO_MONEDA_ARS

        if aplica_iva:
            ws_ing.cell(row=r, column=6, value=f"=ROUND(E{r}/1.21, 2)").number_format = FORMATO_MONEDA_ARS
            ws_ing.cell(row=r, column=7, value=f"=ROUND(E{r}-F{r}, 2)").number_format = FORMATO_MONEDA_ARS
        else:
            ws_ing.cell(row=r, column=6, value=f"=E{r}").number_format = FORMATO_MONEDA_ARS
            ws_ing.cell(row=r, column=7, value=0.0).number_format = FORMATO_MONEDA_ARS

        for c in range(1, 8):
            ws_ing.cell(row=r, column=c).border = BORDE_SUAVE

    r_tot_ing = 4 + len(ops_a_volcar)
    ws_ing.cell(row=r_tot_ing, column=2, value="TOTAL INGRESOS").font = Font(bold=True, color=COLOR_PRIMARIO)
    for col_i in (5, 6, 7):
        letra = get_column_letter(col_i)
        c_t = ws_ing.cell(row=r_tot_ing, column=col_i, value=f"=SUM({letra}4:{letra}{r_tot_ing-1})")
        c_t.font = Font(bold=True, color=COLOR_PRIMARIO)
        c_t.number_format = FORMATO_MONEDA_ARS
    for c in range(1, 8):
        ws_ing.cell(row=r_tot_ing, column=c).border = BORDE_TOTAL
        ws_ing.cell(row=r_tot_ing, column=c).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )

    _ajustar_columnas(ws_ing)
    _configurar_impresion(ws_ing, landscape=True)

    # ---------------------------------------------------------
    # HOJA 3: EGRESOS (Gastos y Retiros)
    # ---------------------------------------------------------
    ws_egr = wb.create_sheet(title="Egresos")
    ws_egr["A1"] = f"DETALLE DE EGRESOS Y RETIROS — FECHA: {fecha_display}"
    ws_egr["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)

    headers_egr = ["#", "Tipo", "Concepto / Destinatario", "Monto (ARS)"]
    ws_egr.append([])
    ws_egr.append(headers_egr)
    _estilizar_encabezado(ws_egr, 3, len(headers_egr))
    ws_egr.freeze_panes = "A4"

    egr_a_volcar = res.gastos_retiros if res.gastos_retiros else [
        GastoRetiro(tipo="gasto", concepto="Sin egresos ni retiros registrados", monto=Decimal("0.00"))
    ]

    for idx, eg in enumerate(egr_a_volcar, start=1):
        r = 3 + idx
        ws_egr.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws_egr.cell(row=r, column=2, value=eg.tipo.upper()).alignment = Alignment(horizontal="center")
        ws_egr.cell(row=r, column=3, value=eg.concepto)
        c_m = ws_egr.cell(row=r, column=4, value=float(to_dec(eg.monto)))
        c_m.number_format = FORMATO_MONEDA_ARS
        for c in range(1, 5):
            ws_egr.cell(row=r, column=c).border = BORDE_SUAVE

    r_tot_egr = 4 + len(egr_a_volcar)
    ws_egr.cell(row=r_tot_egr, column=3, value="TOTAL EGRESOS Y RETIROS").font = Font(bold=True, color=COLOR_PRIMARIO)
    c_te = ws_egr.cell(row=r_tot_egr, column=4, value=f"=SUM(D4:D{r_tot_egr-1})")
    c_te.font = Font(bold=True, color=COLOR_PRIMARIO)
    c_te.number_format = FORMATO_MONEDA_ARS
    for c in range(1, 5):
        ws_egr.cell(row=r_tot_egr, column=c).border = BORDE_TOTAL
        ws_egr.cell(row=r_tot_egr, column=c).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )

    _ajustar_columnas(ws_egr)
    _configurar_impresion(ws_egr, landscape=False)

    # ---------------------------------------------------------
    # HOJA 4: ARQUEO POR DENOMINACIONES (Billetes ARS)
    # ---------------------------------------------------------
    ws_arq = wb.create_sheet(title="Arqueo_Denominaciones")
    ws_arq["A1"] = f"ARQUEO DE EFECTIVO POR DENOMINACIÓN — FECHA: {fecha_display}"
    ws_arq["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)

    headers_arq = ["Denominación (ARS)", "Cantidad de Billetes", "Subtotal (ARS)"]
    ws_arq.append([])
    ws_arq.append(headers_arq)
    _estilizar_encabezado(ws_arq, 3, len(headers_arq))
    ws_arq.freeze_panes = "A4"

    conteo = arqueo_conteo or {}
    for idx, denom in enumerate(DENOMINACIONES_ARS):
        r = 4 + idx
        cant = int(conteo.get(denom, 0))
        c_d = ws_arq.cell(row=r, column=1, value=denom)
        c_d.number_format = FORMATO_MONEDA_ARS
        c_c = ws_arq.cell(row=r, column=2, value=cant)
        c_c.alignment = Alignment(horizontal="center")
        c_s = ws_arq.cell(row=r, column=3, value=f"=A{r}*B{r}")
        c_s.number_format = FORMATO_MONEDA_ARS
        for c in range(1, 4):
            ws_arq.cell(row=r, column=c).border = BORDE_SUAVE

    r_tot_arq = 4 + len(DENOMINACIONES_ARS)
    ws_arq.cell(row=r_tot_arq, column=1, value="TOTAL EFECTIVO CONTADO").font = Font(bold=True, color=COLOR_PRIMARIO)
    c_tc = ws_arq.cell(row=r_tot_arq, column=2, value=f"=SUM(B4:B{r_tot_arq-1})")
    c_tc.alignment = Alignment(horizontal="center")
    c_tc.font = Font(bold=True, color=COLOR_PRIMARIO)
    c_ts = ws_arq.cell(row=r_tot_arq, column=3, value=f"=SUM(C4:C{r_tot_arq-1})")
    c_ts.font = Font(bold=True, color=COLOR_PRIMARIO)
    c_ts.number_format = FORMATO_MONEDA_ARS
    for c in range(1, 4):
        ws_arq.cell(row=r_tot_arq, column=c).border = BORDE_TOTAL
        ws_arq.cell(row=r_tot_arq, column=c).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )

    _ajustar_columnas(ws_arq)
    _configurar_impresion(ws_arq, landscape=False)

    # ---------------------------------------------------------
    # HOJA 5: DIFERENCIAS Y OBSERVACIONES DE AUDITORÍA
    # ---------------------------------------------------------
    ws_obs = wb.create_sheet(title="Diferencias_Observaciones")
    ws_obs["A1"] = "OBSERVACIONES, CAMPOS DUDOSOS Y JUSTIFICACIÓN DE DIFERENCIAS"
    ws_obs["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)

    ws_obs.append([])
    ws_obs.append(["Sección", "Detalle"])
    _estilizar_encabezado(ws_obs, 3, 2)

    notas = [
        ("Observaciones del Cuaderno", res.observaciones or "Sin observaciones manuscritas."),
        (
            "Campos Marcados como Dudosos por IA",
            ", ".join(res.campos_dudosos) if res.campos_dudosos else "Ninguno. Lectura con alta confianza.",
        ),
        ("Comentarios del Cajero / Supervisor", comentario_auditor or "Sin comentarios adicionales."),
        ("Auditoría Técnica", f"Cierre procesado por usuario autenticado el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}."),
    ]
    for idx, (sec, det) in enumerate(notas, start=4):
        ws_obs.cell(row=idx, column=1, value=sec).font = Font(bold=True)
        ws_obs.cell(row=idx, column=2, value=det).alignment = Alignment(wrap_text=True)
        ws_obs.cell(row=idx, column=1).border = BORDE_SUAVE
        ws_obs.cell(row=idx, column=2).border = BORDE_SUAVE
        ws_obs.row_dimensions[idx].height = 32

    _ajustar_columnas(ws_obs, min_width=25, max_width=80)
    _configurar_impresion(ws_obs, landscape=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def asegurar_archivo_maestro():
    if not os.path.exists(ARCHIVO_MAESTRO):
        wb = Workbook()
        ws = wb.active
        _construir_hoja_historico(ws, [])
        tmp_path = f"{ARCHIVO_MAESTRO}.tmp"
        wb.save(tmp_path)
        os.replace(tmp_path, ARCHIVO_MAESTRO)


def guardar_y_ordenar_en_mensual(res: CierreDia):
    """Guarda de forma atómica en el consolidado mensual evitando duplicados por fecha."""
    nueva_fila = construir_fila_historico(res)
    fecha_nueva = nueva_fila[0]
    filas_existentes = []

    if os.path.exists(ARCHIVO_MAESTRO):
        try:
            wb_viejo = load_workbook(ARCHIVO_MAESTRO)
            ws_viejo = wb_viejo.active
            for row in ws_viejo.iter_rows(min_row=4, values_only=True):
                if not row or not row[0] or str(row[0]).startswith(("TOTAL", "ÓPTICA", "OPTICA")):
                    continue
                if str(row[0]) != fecha_nueva:
                    fila_limpia = list(row[: len(COLUMNAS_HISTORICO)])
                    while len(fila_limpia) < len(COLUMNAS_HISTORICO):
                        fila_limpia.append(0.0)
                    filas_existentes.append(fila_limpia)
        except Exception:
            pass

    filas_existentes.append(nueva_fila)

    wb_nuevo = Workbook()
    ws_nuevo = wb_nuevo.active
    _construir_hoja_historico(ws_nuevo, filas_existentes)

    tmp_path = f"{ARCHIVO_MAESTRO}.tmp"
    wb_nuevo.save(tmp_path)
    os.replace(tmp_path, ARCHIVO_MAESTRO)


asegurar_archivo_maestro()

# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

try:
    app_password = st.secrets.get("APP_PASSWORD", APP_PASSWORD)
except Exception:
    app_password = APP_PASSWORD

if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)
        elif os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=180)

        st.markdown("## Acceso Restringido — Óptica Klimt")
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
        st.markdown("### Óptica Klimt")

    st.write("Tablero de Comando Gerencial")
    st.divider()

    menu_opcion = st.radio(
        "Navegación",
        ["Tablero de Comando (Resumen)", "Procesar Cierre Diario"],
        label_visibility="collapsed",
    )

    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# --- VISTA 1: TABLERO DE COMANDO (PANTALLA PRINCIPAL DIRECTA PARA MARTÍN) ---
if menu_opcion == "Tablero de Comando (Resumen)":
    st.title("Óptica Klimt — Tablero de Comando Gerencial")
    st.write("Vista ejecutiva del consolidado mensual, discriminación de IVA (21%) y distribución por medio de pago.")

    try:
        df_master = pd.read_excel(ARCHIVO_MAESTRO, skiprows=2)
        df_master = df_master.dropna(subset=["Fecha"])
        df_master = df_master[~df_master["Fecha"].astype(str).str.contains("TOTAL|OPTICA|ÓPTICA", case=False)]
    except Exception:
        df_master = pd.DataFrame(columns=COLUMNAS_HISTORICO)

    if not df_master.empty:
        for col_num in COLUMNAS_HISTORICO[1:-1]:
            if col_num in df_master.columns:
                df_master[col_num] = pd.to_numeric(df_master[col_num], errors="coerce").fillna(0.0)
            else:
                df_master[col_num] = 0.0

        t_efectivo = Decimal(str(df_master["Efectivo"].sum())).quantize(Decimal("0.01"))
        t_mp_neto = Decimal(str(df_master["MP Neto"].sum())).quantize(Decimal("0.01"))
        t_mp_iva = Decimal(str(df_master["MP IVA (21%)"].sum())).quantize(Decimal("0.01"))
        t_bbva_neto = Decimal(str(df_master["BBVA Neto"].sum())).quantize(Decimal("0.01"))
        t_bbva_iva = Decimal(str(df_master["BBVA IVA (21%)"].sum())).quantize(Decimal("0.01"))
        t_provincia = Decimal(str(df_master["Provincia"].sum())).quantize(Decimal("0.01"))
        t_celia = Decimal(str(df_master["Celia / Toti"].sum())).quantize(Decimal("0.01"))

        t_mp_total = t_mp_neto + t_mp_iva
        t_bbva_total = t_bbva_neto + t_bbva_iva
        total_iva_mes = t_mp_iva + t_bbva_iva
        total_general_ventas = t_efectivo + t_mp_total + t_bbva_total + t_provincia

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ventas Totales Mes", f"$ {total_general_ventas:,.2f}")
        m2.metric("Total IVA Acumulado (21%)", f"$ {total_iva_mes:,.2f}")
        m3.metric("Tarjetas / QR Totales", f"$ {(t_mp_total + t_bbva_total + t_provincia):,.2f}")
        m4.metric("Retiros (Celia / Toti)", f"$ {t_celia:,.2f}")

        st.divider()
        st.subheader("Distribución de Ingresos por Medio de Pago (Bruto con IVA)")

        df_pie = pd.DataFrame(
            {
                "Medio de Pago": ["Efectivo", "Mercado Pago", "Payway BBVA", "Posnet Provincia"],
                "Monto": [float(t_efectivo), float(t_mp_total), float(t_bbva_total), float(t_provincia)],
            }
        )
        df_pie = df_pie[df_pie["Monto"] > 0]

        if not df_pie.empty:
            base_chart = (
                alt.Chart(df_pie)
                .mark_arc(innerRadius=55, outerRadius=125)
                .encode(
                    theta=alt.Theta(field="Monto", type="quantitative"),
                    color=alt.Color(
                        field="Medio de Pago",
                        type="nominal",
                        scale=alt.Scale(range=["#2ECC71", "#3498DB", "#9B59B6", "#E67E22"]),
                    ),
                    tooltip=["Medio de Pago", alt.Tooltip("Monto:Q", format="$,.2f")],
                )
                .properties(height=360)
            )
            st.altair_chart(base_chart, use_container_width=True)
    else:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ventas Totales Mes", "$ 0.00")
        m2.metric("Total IVA Acumulado (21%)", "$ 0.00")
        m3.metric("Tarjetas / QR Totales", "$ 0.00")
        m4.metric("Retiros (Celia / Toti)", "$ 0.00")
        st.info(
            "La planilla consolidada mensual está inicializada en cero y lista para usar. "
            "Podés descargarla debajo o cargar un cierre diario desde la pestaña 'Procesar Cierre Diario'."
        )

    st.divider()
    st.subheader("Detalle Consolidado Mensual")
    st.dataframe(df_master, use_container_width=True)

    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        with open(ARCHIVO_MAESTRO, "rb") as f:
            st.download_button(
                label="Descargar Excel Consolidado Mensual (.xlsx)",
                data=f.read(),
                file_name=ARCHIVO_MAESTRO,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
    with col_ex2:
        with st.expander("Importar o restaurar Excel mensual existente (.xlsx)"):
            excel_subido = st.file_uploader("Seleccionar planilla .xlsx", type=["xlsx"], key="importar_maestro")
            if excel_subido is not None:
                try:
                    wb_test = load_workbook(io.BytesIO(excel_subido.getvalue()))
                    if "Cierre Mensual" not in wb_test.sheetnames and "Historico_Mensual" not in wb_test.sheetnames:
                        st.error("El archivo subido no corresponde al formato oficial de Óptica Klimt.")
                    else:
                        tmp_imp = f"{ARCHIVO_MAESTRO}.tmp"
                        with open(tmp_imp, "wb") as f_out:
                            f_out.write(excel_subido.getbuffer())
                        os.replace(tmp_imp, ARCHIVO_MAESTRO)
                        st.success("Planilla consolidada restaurada correctamente.")
                        st.rerun()
                except Exception as err:
                    st.error(f"Archivo inválido o corrupto: {err}")

# --- VISTA 2: PROCESAR CIERRE DIARIO ---
elif menu_opcion == "Procesar Cierre Diario":
    st.title("Óptica Klimt — Procesamiento y Auditoría de Cierre Diario")
    st.write("Suba la fotografía de la libreta para extraer operaciones, discriminar IVA y generar el acta de cierre.")

    uploaded_file = st.file_uploader("Seleccionar imagen de la planilla (JPG, PNG)", type=["jpg", "jpeg", "png"])

    if uploaded_file is not None:
        col_img, col_act = st.columns([1, 1.5], gap="large")
        with col_img:
            st.image(uploaded_file, caption="Planilla cargada", use_container_width=True)
        with col_act:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Procesar Cierre con IA", type="primary", use_container_width=True):
                try:
                    with st.spinner("Analizando libreta y validando importes..."):
                        resultado = extraer_cierre_gemini(uploaded_file.getvalue())
                        st.session_state.resultado = resultado
                        st.session_state.guardado_maestro = False
                        st.success("Planilla procesada y auditada correctamente.")
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

    if "resultado" in st.session_state:
        res: CierreDia = st.session_state.resultado
        d = procesar_datos_dia(res)

        st.divider()
        fecha_titulo = res.fecha.strftime("%d/%m/%Y") if res.fecha else "Sin fecha"
        st.subheader(f"Resumen Auditado — Fecha: {fecha_titulo}")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Efectivo Ventas", f"$ {d['efectivo']:,.2f}")
        c2.metric("Mercado Pago", f"$ {d['mp_bruto']:,.2f}", f"IVA: $ {d['mp_iva']:,.2f}")
        c3.metric("Payway BBVA", f"$ {d['bbva_bruto']:,.2f}", f"IVA: $ {d['bbva_iva']:,.2f}")
        c4.metric("Provincia", f"$ {d['provincia']:,.2f}")
        c5.metric("Retiros (Celia/Toti)", f"$ {d['celia_toti']:,.2f}", f"Gastos: $ {d['gastos']:,.2f}")

        if res.campos_dudosos:
            st.warning(f"Revisar valores marcados como poco claros: {', '.join(res.campos_dudosos)}")

        # --- ARQUEO DE BILLETES Y DATOS DE AUDITORÍA OPCIONALES ---
        with st.expander("Arqueo de Billetes y Datos de Auditoría (Opcional para el Excel)", expanded=False):
            col_m1, col_m2, col_m3 = st.columns(3)
            turno_in = col_m1.text_input("Turno:", value=res.turno or "Único")
            cajero_in = col_m2.text_input("Cajero:", value=res.cajero or "Celia")
            sup_in = col_m3.text_input("Supervisor:", value="Martín")
            comentario_in = st.text_input("Observaciones del cierre:", value=res.observaciones or "")

            res.turno = turno_in
            res.cajero = cajero_in

            st.markdown("**Conteo de Billetes (ARS):**")
            cols_billetes = st.columns(5)
            arqueo_dict: dict[int, int] = {}
            for idx_b, denom in enumerate(DENOMINACIONES_ARS):
                with cols_billetes[idx_b % 5]:
                    arqueo_dict[denom] = st.number_input(
                        f"$ {denom:,}", min_value=0, step=1, value=0, key=f"bill_{denom}"
                    )

        st.divider()
        st.subheader("Exportación y Confirmación")

        fecha_slug = res.fecha.strftime("%Y-%m-%d") if res.fecha else datetime.now().strftime("%Y-%m-%d")
        nombre_archivo_cierre = (
            f"Cierre_Caja_{fecha_slug}_{sanitizar_nombre_archivo(res.turno or 'Unico')}_"
            f"{sanitizar_nombre_archivo(res.cajero or 'Celia')}.xlsx"
        )

        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            if not st.session_state.get("guardado_maestro", False):
                if st.button("Confirmar y Sumar al Consolidado Mensual", type="primary", use_container_width=True):
                    guardar_y_ordenar_en_mensual(res)
                    st.session_state.guardado_maestro = True
                    st.success("Cierre confirmado y registrado en el consolidado mensual.")
            else:
                st.info("Este cierre ya fue confirmado en el consolidado mensual.")

        with col_dl2:
            excel_cierre_unico = generar_excel_completo_cierre_bytes(
                res=res,
                arqueo_conteo=arqueo_dict,
                supervisor=sup_in,
                comentario_auditor=comentario_in,
            )
            st.download_button(
                label="Descargar Excel Completo de Este Cierre (.xlsx)",
                data=excel_cierre_unico,
                file_name=nombre_archivo_cierre,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_dl3:
            if os.path.exists(ARCHIVO_MAESTRO):
                with open(ARCHIVO_MAESTRO, "rb") as f:
                    st.download_button(
                        label="Descargar Consolidado Mensual (.xlsx)",
                        data=f.read(),
                        file_name=ARCHIVO_MAESTRO,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
