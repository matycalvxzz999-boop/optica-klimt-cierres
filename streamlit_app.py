import io
import os
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from extractor import extraer_cierre_gemini, CierreDia

load_dotenv()
APP_PASSWORD = os.getenv("APP_PASSWORD", "optica2026")
ARCHIVO_MAESTRO = "cierre_mensual_optica.xlsx"

st.set_page_config(
    page_title="Optica Klimt - Cierres de Caja",
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
    "Ventas en Efectivo",
    "Ventas Mercado Pago",
    "Ventas Payway BBVA",
    "Ventas Payway Provincia",
    "Gastos",
    "Celia o Toti",
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
    mp_ventas = 0.0
    bbva_ventas = 0.0
    provincia_ventas = 0.0

    for op in res.operaciones:
        monto = float(op.monto or 0)
        if op.medio_pago == "qr_mp":
            mp_ventas += monto
        elif op.medio_pago == "payway_bbva":
            bbva_ventas += monto
        elif op.medio_pago == "posnet_provincia":
            provincia_ventas += monto
        elif op.medio_pago == "efectivo":
            efectivo_ventas += monto
        else:
            bbva_ventas += monto

    if efectivo_ventas == 0 and res.total_efectivo_cuaderno:
        efectivo_ventas = float(res.total_efectivo_cuaderno)
    if bbva_ventas == 0 and res.total_tarjetas_cuaderno:
        bbva_ventas = float(res.total_tarjetas_cuaderno)

    gastos_total = 0.0
    celia_toti_total = 0.0
    detalles_gastos = []

    for g in res.gastos_retiros:
        monto_g = float(g.monto or 0)
        concepto_lower = g.concepto.lower()
        tipo_lower = g.tipo.lower()

        if "celia" in concepto_lower or "toti" in concepto_lower or "retiro" in concepto_lower or tipo_lower == "retiro":
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
        "efectivo_ventas": efectivo_ventas,
        "mp_ventas": mp_ventas,
        "bbva_ventas": bbva_ventas,
        "provincia_ventas": provincia_ventas,
        "gastos": gastos_total,
        "celia_toti": celia_toti_total,
        "detalle_obs": " | ".join(detalles_gastos),
    }

def construir_fila_dia(res: CierreDia) -> list:
    fecha_str = str(res.fecha) if res.fecha else "Sin fecha"
    d = procesar_datos_dia(res)
    return [
        fecha_str,
        d["efectivo_ventas"],
        d["mp_ventas"],
        d["bbva_ventas"],
        d["provincia_ventas"],
        d["gastos"],
        d["celia_toti"],
        d["detalle_obs"],
    ]

def crear_libro_base(filas_datos: list) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cierre Mensual"

    ws["A1"] = "OPTICA KLIMT - CIERRE DE CAJA MENSUAL"
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
            if 2 <= col <= 7:
                cell.number_format = "$#,##0.00"
                if col in (2, 6, 7):
                    cell.font = Font(bold=True)

    if filas_ordenadas:
        fila_inicio = 4
        fila_fin = ws.max_row
        fila_tot = fila_fin + 1
        ws.cell(row=fila_tot, column=1, value="TOTAL MES").font = Font(bold=True, size=11, color="1F4E78")
        ws.cell(row=fila_tot, column=1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.cell(row=fila_tot, column=1).border = BORDE_TOTAL

        letras = ["B", "C", "D", "E", "F", "G"]
        for idx, letra in enumerate(letras, start=2):
            c_tot = ws.cell(row=fila_tot, column=idx, value=f"=SUM({letra}{fila_inicio}:{letra}{fila_fin})")
            c_tot.font = Font(bold=True, size=11, color="1F4E78")
            c_tot.number_format = "$#,##0.00"
            c_tot.border = BORDE_TOTAL
            c_tot.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ajustar_columnas(ws)
    return wb

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
        wb_viejo = load_workbook(ARCHIVO_MAESTRO)
        ws_viejo = wb_viejo.active
        for row in ws_viejo.iter_rows(min_row=4, values_only=True):
            if not row or not row[0] or str(row[0]).startswith("TOTAL"):
                continue
            if str(row[0]) != fecha_nueva:
                filas_existentes.append(list(row[:len(COLUMNAS_EXCEL)]))

    filas_existentes.append(nueva_fila)
    wb_nuevo = crear_libro_base(filas_existentes)
    wb_nuevo.save(ARCHIVO_MAESTRO)

# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

# Leer contraseña institucional desde los Secrets de Streamlit con fallback a variable local
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

# --- MENÚ LATERAL PROFESIONAL ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=200)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=200)
    else:
        st.markdown("### Optica Klimt")
    
    st.write("Sistema de Gestión de Cierres")
    st.divider()
    
    menu_opcion = st.radio(
        "Navegación",
        ["Procesar Cierre Diario", "Resumen Mensual"],
        label_visibility="collapsed"
    )
    
    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# --- VISTA 1: PROCESAR CIERRE ---
if menu_opcion == "Procesar Cierre Diario":
    st.title("Optica Klimt - Cierres de Caja")
    st.write("Suba la imagen del cuaderno de cierre diario para procesar los registros.")

    uploaded_file = st.file_uploader("Seleccionar imagen de la planilla (JPG, PNG)", type=["jpg", "jpeg", "png"])

    if uploaded_file is not None:
        col_img, col_act = st.columns([1, 1.5], gap="large")
        with col_img:
            st.image(uploaded_file, caption="Planilla cargada", use_column_width=True)
        with col_act:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Procesar Cierre con IA", type="primary", use_container_width=True):
                try:
                    with st.spinner("Procesando datos de la libreta..."):
                        resultado = extraer_cierre_gemini(uploaded_file.getvalue())
                        st.session_state.resultado = resultado
                        st.session_state.guardado_maestro = False
                        st.success("Datos procesados correctamente.")
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

    if "resultado" in st.session_state:
        res: CierreDia = st.session_state.resultado
        d = procesar_datos_dia(res)

        st.divider()
        st.subheader(f"Resumen del Día — Fecha: {res.fecha or 'No detectada'}")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Efectivo Ventas", f"${d['efectivo_ventas']:,.0f}")
        c2.metric("Mercado Pago", f"${d['mp_ventas']:,.0f}")
        c3.metric("Payway BBVA", f"${d['bbva_ventas']:,.0f}")
        c4.metric("Payway Provincia", f"${d['provincia_ventas']:,.0f}")
        c5.metric("Gastos / Retiros", f"${d['gastos']:,.0f}", f"Celia/Toti: ${d['celia_toti']:,.0f}")

        if res.campos_dudosos:
            st.warning(f"Revisar valores poco claros: {', '.join(res.campos_dudosos)}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Exportar Registros")
        
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            if not st.session_state.get("guardado_maestro", False):
                if st.button("Sumar al Excel Mensual", type="primary", use_container_width=True):
                    guardar_y_ordenar_en_mensual(res)
                    st.session_state.guardado_maestro = True
                    st.success("Cierre agregado al acumulado mensual.")
            else:
                st.info("El cierre ya se encuentra registrado en el archivo mensual.")

        with col_dl2:
            if os.path.exists(ARCHIVO_MAESTRO):
                with open(ARCHIVO_MAESTRO, "rb") as f:
                    st.download_button(
                        label="Descargar Excel Mensual",
                        data=f.read(),
                        file_name=ARCHIVO_MAESTRO,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        with col_dl3:
            st.download_button(
                label="Descargar Solo Este Día",
                data=generar_excel_solo_dia_bytes(res),
                file_name=f"Cierre_Optica_{res.fecha or 'dia'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# --- VISTA 2: RESUMEN MENSUAL ---
elif menu_opcion == "Resumen Mensual":
    st.title("Optica Klimt - Planilla Mensual")
    st.write("Acceso directo al archivo consolidado del mes en curso.")

    if os.path.exists(ARCHIVO_MAESTRO):
        st.success("Archivo maestro disponible para descarga.")
        with open(ARCHIVO_MAESTRO, "rb") as f:
            st.download_button(
                label="Descargar Planilla Consolidada (.xlsx)",
                data=f.read(),
                file_name=ARCHIVO_MAESTRO,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    else:
        st.info("Aún no hay registros cargados en el archivo mensual.")
