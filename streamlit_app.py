import io
import os
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from extractor import extraer_cierre_gemini, CierreDia
import pandas as pd

load_dotenv()
APP_PASSWORD = os.getenv("APP_PASSWORD", "optica2026")
ARCHIVO_MAESTRO = "cierre_mensual_optica.xlsx"

st.set_page_config(
    page_title="Optica Klimt - Tablero de Comando",
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

        # REQUERIMIENTO: Separar retiros de Celia / Toti (ignorar como gasto común)
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
        ["Procesar Cierre Diario", "Tablero de Comando (Resumen)"],
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
        c5.metric("Retiros (Celia/Toti)", f"${d['celia_toti']:,.0f}", f"Gastos: ${d['gastos']:,.0f}")

        # REQUERIMIENTO: Discriminación de IVA (21%) para Mercado Pago y BBVA
        st.markdown("### 📊 Discriminación Fiscal (IVA 21%)")
        iva_mp_neto = d['mp_ventas'] / 1.21
        iva_mp_val = d['mp_ventas'] - iva_mp_neto
        iva_bbva_neto = d['bbva_ventas'] / 1.21
        iva_bbva_val = d['bbva_ventas'] - iva_bbva_neto
        total_iva_dia = iva_mp_val + iva_bbva_val

        col_iva1, col_iva2, col_iva3 = st.columns(3)
        col_iva1.metric("Neto Gravado (MP + BBVA)", f"${(iva_mp_neto + iva_bbva_neto):,.0f}")
        col_iva2.metric("IVA Debito (21%)", f"${total_iva_dia:,.0f}")
        col_iva3.metric("Efectivo / Provincia (Sin IVA)", f"${(d['efectivo_ventas'] + d['provincia_ventas']):,.0f}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Exportar Registros")
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if not st.session_state.get("guardado_maestro", False):
                if st.button("Sumar al Excel Mensual", type="primary", use_container_width=True):
                    guardar_y_ordenar_en_mensual(res)
                    st.session_state.guardado_maestro = True
                    st.success("¡Cierre agregado y guardado en el maestro mensual!")
            else:
                st.info("El cierre ya se encuentra registrado en el archivo mensual.")

        with col_dl2:
            if os.path.exists(ARCHIVO_MAESTRO):
                with open(ARCHIVO_MAESTRO, "rb") as f:
                    st.download_button(
                        label="Descargar Planilla Consolidada (.xlsx)",
                        data=f.read(),
                        file_name=ARCHIVO_MAESTRO,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

# --- VISTA 2: TABLERO DE COMANDO (RESUMEN MENSUAL Y GRÁFICOS) ---
elif menu_opcion == "Tablero de Comando (Resumen)":
    st.title("Optica Klimt - Tablero de Comando Gerencial")
    st.write("Vista consolidada de resultados, evolución de ingresos y distribución fiscal.")

    if os.path.exists(ARCHIVO_MAESTRO):
        df_master = pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Cierre Mensual", skiprows=2)
        # Limpiar filas vacías o de totales
        df_master = df_master.dropna(subset=["Fecha"])
        df_master = df_master[~df_master["Fecha"].astype(str).str.contains("TOTAL|OPTICA", case=False)]

        if not df_master.empty:
            # Métricas principales
            t_efectivo = df_master["Ventas en Efectivo"].sum()
            t_mp = df_master["Ventas Mercado Pago"].sum()
            t_bbva = df_master["Ventas Payway BBVA"].sum()
            t_provincia = df_master["Ventas Payway Provincia"].sum()
            t_gastos = df_master["Gastos"].sum()
            t_celia = df_master["Celia o Toti"].sum()

            total_general_ventas = t_efectivo + t_mp + t_bbva + t_provincia

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ventas Totales Mes", f"${total_general_ventas:,.0f}")
            m2.metric("Efectivo Acumulado", f"${t_efectivo:,.0f}")
            m3.metric("Tarjetas / QR Totales", f"${(t_mp + t_bbva + t_provincia):,.0f}")
            m4.metric("Retiros (Celia/Toti)", f"${t_celia:,.0f}")

            st.divider()

            # Discriminación de IVA acumulado del mes
            st.subheader("💼 Discriminación Fiscal Mensual (IVA)")
            iva_acum_mp = t_mp - (t_mp / 1.21)
            iva_acum_bbva = t_bbva - (t_bbva / 1.21)
            total_iva_mes = iva_acum_mp + iva_acum_bbva
            neto_gravado_mes = (t_mp / 1.21) + (t_bbva / 1.21)

            i1, i2, i3 = st.columns(3)
            i1.metric("Total Neto Gravado", f"${neto_gravado_mes:,.0f}")
            i2.metric("IVA Acumulado Total (21%)", f"${total_iva_mes:,.0f}")
            i3.metric("Total No Gravado (Efectivo + Prov.)", f"${(t_efectivo + t_provincia):,.0f}")

            st.divider()

            # GRÁFICO DE TORTA (Distribución de ingresos por medio de pago + IVA)
            st.subheader("📊 Distribución de Ingresos y Medios de Pago")
            
            import matplotlib.pyplot as plt
            fig, ax = plt.subplots(figsize=(8, 5))
            
            labels = ['Efectivo', 'Mercado Pago', 'Payway BBVA', 'Payway Provincia']
            sizes = [t_efectivo, t_mp, t_bbva, t_provincia]
            colors = ['#2ECC71', '#3498DB', '#9B59B6', '#E67E22']
            
            # Filtrar ceros para que el gráfico quede prolijo
            datos_filtrados = [(l, s, c) for l, s, c in zip(labels, sizes, colors) if s > 0]
            if datos_filtrados:
                labs, vals, cols = zip(*datos_filtrados)
                ax.pie(vals, labels=labs, autopct='%1.1f%%', startangle=140, colors=cols, textprops={'fontsize': 10, 'weight': 'bold'})
                ax.axis('equal')
                st.pyplot(fig)
            else:
                st.info("Aún no hay suficientes montos registrados para graficar.")

            st.divider()
            st.subheader("📋 Detalle de Cierres Registrados")
            st.dataframe(df_master, use_container_width=True)

            with open(ARCHIVO_MAESTRO, "rb") as f:
                st.download_button(
                    label="📥 Descargar Planilla Consolidada Completa (.xlsx)",
                    data=f.read(),
                    file_name=ARCHIVO_MAESTRO,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.info("El archivo maestro está vacío. Cargá tu primer cierre desde la pestaña anterior.")
    else:
        st.info("Aún no hay registros guardados. Procesá un cierre para iniciar el Tablero de Comando.")
