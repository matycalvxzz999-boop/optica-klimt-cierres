import io
import os
import zipfile
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Optional
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from extractor_facturas import (
    FacturaExtraida,
    extraer_factura,
    generar_nombre_archivo,
    solo_digitos,
    to_dec,
)
from planilla_writer import ErrorPlanilla, PlanillaMemoria

load_dotenv()

ZONA_ARG = ZoneInfo("America/Argentina/Buenos_Aires")

st.set_page_config(
    page_title="3CICON — Facturas IVA Compras",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- ESTILOS MODERNOS Y DE ALTA VISIBILIDAD ---
st.markdown(
    """
    <style>
    /* Fondo general suave */
    .stApp {
        background-color: #F1F5F9;
        color: #0F172A;
    }
    /* Ocultar sidebar para pantalla más limpia */
    [data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none;
    }
    /* Encabezado superior */
    .top-banner {
        background: linear-gradient(135deg, #0F766E 0%, #115E59 100%);
        padding: 1.4rem 1.8rem;
        border-radius: 14px;
        color: #FFFFFF;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 14px rgba(15, 118, 110, 0.18);
    }
    .top-banner h1 {
        margin: 0;
        font-size: 1.75rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: -0.3px;
    }
    .top-banner p {
        margin: 0.3rem 0 0 0;
        font-size: 1rem;
        color: #CCFBF1;
    }
    /* Tarjetas contenedoras */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important;
        padding: 1.1rem 1.25rem !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    /* Etiqueta de paso */
    .paso-badge {
        display: inline-block;
        background-color: #CCFBF1;
        color: #0F766E;
        font-weight: 700;
        font-size: 0.82rem;
        padding: 0.22rem 0.65rem;
        border-radius: 999px;
        margin-bottom: 0.4rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .paso-titulo {
        font-size: 1.2rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0 0 0.6rem 0;
    }
    /* Pestañas grandes */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 0.65rem 1.4rem;
        font-size: 1rem;
        font-weight: 600;
        color: #334155;
        border: 1px solid #CBD5E1;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
        border-color: #0F766E !important;
    }
    /* Botones más grandes y visibles */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
        font-size: 1rem;
        padding: 0.65rem 1.2rem;
        border: 1px solid #CBD5E1;
    }
    .stButton > button[kind="primary"] {
        background-color: #0F766E;
        border-color: #0F766E;
        color: #FFFFFF;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #115E59;
        border-color: #115E59;
    }
    .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 700;
        font-size: 1rem;
        padding: 0.75rem 1.2rem;
        background-color: #0F766E;
        color: #FFFFFF;
        border: none;
    }
    .stDownloadButton > button:hover {
        background-color: #115E59;
        color: #FFFFFF;
    }
    /* Métricas */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F766E;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.92rem;
        font-weight: 600;
        color: #475569;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _get_secret_or_env(key: str) -> Optional[str]:
    try:
        val = st.secrets.get(key)
        if val is not None and str(val).strip():
            return str(val).strip()
    except Exception:
        pass
    env_val = os.getenv(key)
    return env_val.strip() if env_val and env_val.strip() else None


# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

app_password = _get_secret_or_env("APP_PASSWORD")
if app_password and not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col_login, _ = st.columns([1.2, 1.6, 1.2])
    with col_login:
        with st.container(border=True):
            st.markdown("### 3CICON — IVA Compras")
            st.write("Ingresá la contraseña para entrar al sistema.")
            pwd = st.text_input("Contraseña", type="password", placeholder="Escribí tu clave...")
            if st.button("Entrar", type="primary", use_container_width=True):
                if pwd == app_password:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
    st.stop()


# --- UTILIDADES (MISMA LÓGICA INTACTA) ---
def _parse_fecha_celda(val: Any) -> Optional[date]:
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    txt = str(val or "").strip()
    if not txt:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:10], fmt).date()
        except ValueError:
            pass
    return None


def _fila_a_factura(row: pd.Series) -> FacturaExtraida:
    alic_raw = row.get("alicuota_iva")
    alic_val = None
    if pd.notna(alic_raw) and str(alic_raw).strip() != "":
        try:
            d_al = Decimal(str(alic_raw).replace("%", "").strip())
            alic_val = d_al if d_al > 0 else None
        except Exception:
            alic_val = None

    cod_raw = row.get("codigo_tipo")
    cod_val = int(cod_raw) if pd.notna(cod_raw) and str(cod_raw).strip() != "" else 1

    pv_raw = row.get("punto_venta")
    pv_val = int(pv_raw) if pd.notna(pv_raw) and str(pv_raw).strip() != "" else None

    nro_raw = row.get("numero")
    nro_val = int(nro_raw) if pd.notna(nro_raw) and str(nro_raw).strip() != "" else None

    return FacturaExtraida(
        fecha_emision=_parse_fecha_celda(row.get("fecha_emision")),
        codigo_tipo=cod_val,
        tipo_letra=str(row.get("tipo_letra") or "A"),
        punto_venta=pv_val,
        numero=nro_val,
        cuit_emisor=solo_digitos(str(row.get("cuit_emisor") or "")),
        razon_social_emisor=str(row.get("razon_social_emisor") or "").strip().upper(),
        neto_gravado=to_dec(row.get("neto_gravado")),
        alicuota_iva=alic_val,
        importe_iva=to_dec(row.get("importe_iva")),
        imp_internos_no_gravado=to_dec(row.get("imp_internos_no_gravado")),
        perc_iibb_caba=to_dec(row.get("perc_iibb_caba")),
        perc_iibb_pba=to_dec(row.get("perc_iibb_pba")),
        perc_iva=to_dec(row.get("perc_iva")),
        total=to_dec(row.get("total")),
        metodo=str(row.get("metodo") or "MANUAL"),
        observaciones=str(row.get("observaciones") or ""),
    )


def _crear_zip_renombrado(items_zip: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    usados: dict[str, int] = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in items_zip:
            base, ext = os.path.splitext(nombre)
            if nombre in usados:
                usados[nombre] += 1
                nombre_final = f"{base}_{usados[nombre]}{ext}"
            else:
                usados[nombre] = 0
                nombre_final = nombre
            z.writestr(nombre_final, contenido)
    return buf.getvalue()


# --- ENCABEZADO SUPERIOR + ACCIONES RÁPIDAS ---
col_hdr, col_actions = st.columns([4, 1.4], vertical_alignment="center")
with col_hdr:
    st.markdown(
        """
        <div class="top-banner">
            <h1>3CICON — Carga de Facturas</h1>
            <p>Subí tu planilla del mes y las facturas. El sistema lee los datos, los ordena y renombra los archivos.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with col_actions:
    if st.button("Limpiar todo", use_container_width=True):
        for k in ("lote_df", "archivos_raw", "excel_salida", "zip_salida", "reporte_escritura"):
            st.session_state.pop(k, None)
        st.rerun()
    if app_password:
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.autenticado = False
            st.rerun()


# --- NAVEGACIÓN POR PESTAÑAS ---
tab_carga, tab_tablero = st.tabs(["Cargar Facturas", "Tablero de Control"])


# ==============================================================================
# PESTAÑA 1: CARGAR Y PROCESAR
# ==============================================================================
with tab_carga:
    col_p1, col_p2 = st.columns(2, gap="medium")

    with col_p1:
        with st.container(border=True):
            st.markdown('<span class="paso-badge">Paso 1</span>', unsafe_allow_html=True)
            st.markdown('<div class="paso-titulo">Subí la planilla Excel del mes</div>', unsafe_allow_html=True)
            excel_up = st.file_uploader(
                "Archivo .xlsx oficial de 3CICON",
                type=["xlsx"],
                key="planilla_up",
                label_visibility="collapsed",
            )
            if excel_up is not None:
                try:
                    pm_check = PlanillaMemoria(excel_up.getvalue())
                    est = pm_check.resumen_estado()
                    st.success(
                        f"Planilla lista — Cargadas: **{est['cantidad_cargadas']}** | "
                        f"Próxima fila: **{est['proxima_fila']}** | "
                        f"Libres: **{est['filas_libres']}**"
                    )
                except ErrorPlanilla as err_p:
                    st.error(str(err_p))
            else:
                st.caption("Elegí o arrastrá el archivo Excel de IVA Compras.")

    with col_p2:
        with st.container(border=True):
            st.markdown('<span class="paso-badge">Paso 2</span>', unsafe_allow_html=True)
            st.markdown('<div class="paso-titulo">Subí las facturas (PDF o fotos)</div>', unsafe_allow_html=True)
            facturas_up = st.file_uploader(
                "Podés subir varias facturas juntas",
                type=["pdf", "jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="facturas_up",
                label_visibility="collapsed",
            )
            if facturas_up:
                st.info(f"**{len(facturas_up)} archivo(s)** seleccionado(s) para leer.")
            else:
                st.caption("Podés arrastrar varios PDF, JPG o PNG al mismo tiempo.")

    if facturas_up:
        st.markdown("")
        if st.button(f"Leer {len(facturas_up)} factura(s) automáticamente", type="primary", use_container_width=True):
            filas_lote = []
            archivos_raw: dict[str, bytes] = {}
            progreso = st.progress(0)
            estado_txt = st.empty()

            for idx, archivo in enumerate(facturas_up, start=1):
                estado_txt.info(f"Leyendo factura {idx} de {len(facturas_up)}: **{archivo.name}** ...")
                contenido = archivo.getvalue()
                archivos_raw[archivo.name] = contenido
                ext = os.path.splitext(archivo.name)[1].lower()
                try:
                    fac = extraer_factura(contenido, archivo.name)
                    filas_lote.append(
                        {
                            "incluir": True,
                            "archivo_original": archivo.name,
                            "fecha_emision": fac.fecha_emision.strftime("%Y-%m-%d") if fac.fecha_emision else "",
                            "razon_social
