import contextlib
import io
import json
import os
import re
import time
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Optional
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from supabase import Client, create_client

from extractor import CierreDia, GastoRetiro, Operacion, extraer_cierre_gemini

load_dotenv()

ZONA_ARG = ZoneInfo("America/Argentina/Buenos_Aires")
ARCHIVO_MAESTRO_LOCAL = "cierre_mensual_optica.xlsx"
ARCHIVO_AUDITORIA_LOCAL = "auditoria_cierres_local.json"
ARCHIVO_LOCK = f"{ARCHIVO_MAESTRO_LOCAL}.lock"
LOGO_CANDIDATOS = ("logo.png", "logo.jpg")

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
    "Turno",
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


# --- TIEMPO, SEGURIDAD Y ARCHIVOS ---
def ahora_arg() -> datetime:
    return datetime.now(ZONA_ARG)


def _get_secret_or_env(key: str) -> Optional[str]:
    try:
        val = st.secrets.get(key)
        if val is not None and str(val).strip():
            return str(val).strip()
    except Exception:
        pass
    env_val = os.getenv(key)
    if env_val and env_val.strip():
        return env_val.strip()
    return None


def obtener_password_configurada() -> Optional[str]:
    return _get_secret_or_env("APP_PASSWORD")


def obtener_cliente_supabase() -> Optional[Client]:
    url = _get_secret_or_env("SUPABASE_URL")
    key = _get_secret_or_env("SUPABASE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


def mostrar_logo(width: int = 180) -> bool:
    for ruta in LOGO_CANDIDATOS:
        if os.path.exists(ruta):
            st.image(ruta, width=width)
            return True
    return False


def sanitizar_nombre_archivo(texto: str) -> str:
    limpio = re.sub(r"[^A-Za-z0-9_-]", "_", (texto or "").strip())
    return re.sub(r"_+", "_", limpio).strip("_") or "SinDato"


def sanitizar_celda_excel(valor: Any) -> str:
    texto = "" if valor is None else str(valor)
    if texto.startswith(("=", "+", "-", "@", "\t", "\r")):
        return "'" + texto
    return texto


@contextlib.contextmanager
def bloqueo_archivo(lock_path: str = ARCHIVO_LOCK, timeout_seg: float = 10.0):
    inicio = time.monotonic()
    fd = None
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("utf-8"))
            break
        except FileExistsError:
            with contextlib.suppress(OSError):
                if time.time() - os.path.getmtime(lock_path) > 60:
                    os.remove(lock_path)
                    continue
            if time.monotonic() - inicio >= timeout_seg:
                raise TimeoutError("El archivo local está en uso por otra operación. Reintentá.")
            time.sleep(0.15)
    try:
        yield
    finally:
        if fd is not None:
            with contextlib.suppress(OSError):
                os.close(fd)
        with contextlib.suppress(OSError):
            os.remove(lock_path)


def guardar_bytes_atomico(ruta_destino: str, contenido_bytes: bytes):
    tmp_path = f"{ruta_destino}.tmp.{os.getpid()}"
    with open(tmp_path, "wb") as f_out:
        f_out.write(contenido_bytes)
        f_out.flush()
        os.fsync(f_out.fileno())
    os.replace(tmp_path, ruta_destino)


# --- UTILIDADES FINANCIERAS EXACTAS (DECIMAL) ---
def to_dec(val: Decimal | float | int | str | None) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    if isinstance(val, Decimal):
        if val.is_nan() or val.is_infinite():
            return Decimal("0.00")
        return val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    try:
        if isinstance(val, float) and pd.isna(val):
            return Decimal("0.00")
        limpio = str(val).replace("$", "").strip()
        return Decimal(limpio).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def desglosar_iva_21(bruto: Decimal) -> tuple[Decimal, Decimal]:
    """Aplica divisor estricto 1.21 con redondeo contable ROUND_HALF_UP."""
    bruto = to_dec(bruto)
    if bruto == Decimal("0.00"):
        return Decimal("0.00"), Decimal("0.00")
    neto = (bruto / Decimal("1.21")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    iva = (bruto - neto).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return neto, iva


def sumar_serie_decimal(df: pd.DataFrame, columna: str) -> Decimal:
    if columna not in df.columns or df.empty:
        return Decimal("0.00")
    total = Decimal("0.00")
    for raw_val in df[columna].tolist():
        total += to_dec(raw_val)
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# --- CÁLCULOS DE CIERRE ---
def procesar_datos_dia(res: CierreDia, arqueo_conteo: Optional[dict[int, int]] = None) -> dict[str, Any]:
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

    if efectivo_ventas == Decimal("0.00") and res.total_efectivo_cuaderno is not None:
        efectivo_ventas = to_dec(res.total_efectivo_cuaderno)
    if (
        mp_total == Decimal("0.00")
        and bbva_total == Decimal("0.00")
        and provincia_ventas == Decimal("0.00")
        and res.total_tarjetas_cuaderno is not None
    ):
        bbva_total = to_dec(res.total_tarjetas_cuaderno)

    mp_neto, mp_iva = desglosar_iva_21(mp_total)
    bbva_neto, bbva_iva = desglosar_iva_21(bbva_total + transf_ventas)

    gastos_total = Decimal("0.00")
    celia_toti_total = Decimal("0.00")
    detalles_gastos: list[str] = []

    for g in res.gastos_retiros:
        monto_g = to_dec(g.monto)
        concepto_lower = (g.concepto or "").lower()
        tipo_lower = (g.tipo or "").lower()
        es_retiro = bool(
            re.search(r"\b(cel|celia|toti)\b", concepto_lower) or tipo_lower == "retiro"
        )
        if es_retiro:
            celia_toti_total += monto_g
            detalles_gastos.append(f"[Retiro] {g.concepto} (${monto_g:,.2f})")
        else:
            gastos_total += monto_g
            detalles_gastos.append(f"[Gasto] {g.concepto} (${monto_g:,.2f})")

    if gastos_total == Decimal("0.00") and celia_toti_total == Decimal("0.00") and res.total_gastos_cuaderno is not None:
        gastos_total = to_dec(res.total_gastos_cuaderno)

    if res.observaciones:
        detalles_gastos.append(res.observaciones)

    saldo_inicial = to_dec(res.saldo_inicial)
    total_ingresos_brutos = efectivo_ventas + mp_total + bbva_total + transf_ventas + provincia_ventas
    total_egresos = gastos_total + celia_toti_total

    efectivo_esperado_caja = saldo_inicial + efectivo_ventas - total_egresos
    total_general_esperado = saldo_inicial + total_ingresos_brutos - total_egresos

    conteo = arqueo_conteo or {}
    total_contado_billetes = sum(
        (Decimal(denom) * Decimal(max(0, int(cant))) for denom, cant in conteo.items()),
        Decimal("0.00"),
    ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    diferencia_arqueo = (total_contado_billetes - efectivo_esperado_caja).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )

    tiene_caja_cuaderno = res.caja_total_cuaderno is not None
    caja_cuaderno = to_dec(res.caja_total_cuaderno)
    diferencia_cuaderno = (
        (caja_cuaderno - total_general_esperado).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if tiene_caja_cuaderno
        else Decimal("0.00")
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
        "total_contado_billetes": total_contado_billetes,
        "diferencia_arqueo": diferencia_arqueo,
        "tiene_caja_cuaderno": tiene_caja_cuaderno,
        "caja_cuaderno": caja_cuaderno,
        "diferencia_cuaderno": diferencia_cuaderno,
        "detalle_obs": " | ".join(detalles_gastos),
    }


def normalizar_turno(turno: Optional[str]) -> str:
    limpio = re.sub(r"\s+", " ", (turno or "").strip())
    return limpio if limpio else "Único"


def clave_cierre(fecha_iso: str, turno: str) -> str:
    return f"{fecha_iso}::{normalizar_turno(turno).lower()}"


# --- FUENTE DE VERDAD PERSISTENTE (SUPABASE + FALLBACK LOCAL) ---
def _leer_auditoria_local() -> list[dict[str, Any]]:
    if not os.path.exists(ARCHIVO_AUDITORIA_LOCAL):
        return []
    try:
        with open(ARCHIVO_AUDITORIA_LOCAL, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def _guardar_auditoria_local(eventos: list[dict[str, Any]]):
    with bloqueo_archivo():
        guardar_bytes_atomico(
            ARCHIVO_AUDITORIA_LOCAL,
            json.dumps(eventos, ensure_ascii=False, indent=2).encode("utf-8"),
        )


def _cierre_a_registro_db(
    res: CierreDia,
    arqueo_conteo: Optional[dict[int, int]] = None,
    supervisor: str = "Martín",
    comentario_auditor: str = "",
    usuario_actor: str = "Usuario",
) -> dict[str, Any]:
    fecha_iso = res.fecha.strftime("%Y-%m-%d") if res.fecha else ahora_arg().strftime("%Y-%m-%d")
    turno_norm = normalizar_turno(res.turno)
    cajero_norm = (res.cajero or "Celia").strip() or "Celia"
    d = procesar_datos_dia(res, arqueo_conteo=arqueo_conteo)

    payload_cierre = {
        "fecha": fecha_iso,
        "turno": turno_norm,
        "cajero": cajero_norm,
        "supervisor": supervisor,
        "comentario_auditor": comentario_auditor,
        "saldo_inicial": str(d["saldo_inicial"]),
        "total_efectivo_cuaderno": str(res.total_efectivo_cuaderno) if res.total_efectivo_cuaderno is not None else None,
        "total_tarjetas_cuaderno": str(res.total_tarjetas_cuaderno) if res.total_tarjetas_cuaderno is not None else None,
        "total_gastos_cuaderno": str(res.total_gastos_cuaderno) if res.total_gastos_cuaderno is not None else None,
        "caja_total_cuaderno": str(res.caja_total_cuaderno) if res.caja_total_cuaderno is not None else None,
        "campos_dudosos": list(res.campos_dudosos or []),
        "observaciones": res.observaciones or "",
        "operaciones": [
            {
                "detalle": op.detalle,
                "monto": str(to_dec(op.monto)),
                "medio_pago": op.medio_pago,
                "confianza": op.confianza,
            }
            for op in (res.operaciones or [])
        ],
        "gastos_retiros": [
            {
                "tipo": g.tipo,
                "concepto": g.concepto,
                "monto": str(to_dec(g.monto)),
            }
            for g in (res.gastos_retiros or [])
        ],
        "arqueo_conteo": {str(k): int(v) for k, v in (arqueo_conteo or {}).items()},
        "calculos": {
            "efectivo": str(d["efectivo"]),
            "mp_bruto": str(d["mp_bruto"]),
            "mp_neto": str(d["mp_neto"]),
            "mp_iva": str(d["mp_iva"]),
            "bbva_bruto": str(d["bbva_bruto"]),
            "bbva_neto": str(d["bbva_neto"]),
            "bbva_iva": str(d["bbva_iva"]),
            "provincia": str(d["provincia"]),
            "gastos": str(d["gastos"]),
            "celia_toti": str(d["celia_toti"]),
            "total_egresos": str(d["total_egresos"]),
            "total_ingresos_brutos": str(d["total_ingresos_brutos"]),
            "efectivo_esperado_caja": str(d["efectivo_esperado_caja"]),
            "total_general_esperado": str(d["total_general_esperado"]),
            "total_contado_billetes": str(d["total_contado_billetes"]),
            "diferencia_arqueo": str(d["diferencia_arqueo"]),
            "caja_cuaderno": str(d["caja_cuaderno"]),
            "diferencia_cuaderno": str(d["diferencia_cuaderno"]),
            "detalle_obs": d["detalle_obs"],
        },
        "actualizado_por": usuario_actor,
        "actualizado_en": ahora_arg().isoformat(),
    }

    return {
        "clave_cierre": clave_cierre(fecha_iso, turno_norm),
        "fecha": fecha_iso,
        "turno": turno_norm,
        "cajero": cajero_norm,
        "supervisor": supervisor,
        "efectivo": float(d["efectivo"]),
        "mp_neto": float(d["mp_neto"]),
        "mp_iva": float(d["mp_iva"]),
        "bbva_neto": float(d["bbva_neto"]),
        "bbva_iva": float(d["bbva_iva"]),
        "provincia": float(d["provincia"]),
        "gastos": float(d["gastos"]),
        "celia_toti": float(d["celia_toti"]),
        "diferencia_caja": float(d["diferencia_cuaderno"]),
        "detalle_obs": d["detalle_obs"],
        "payload_json": payload_cierre,
        "actualizado_por": usuario_actor,
        "actualizado_en": ahora_arg().isoformat(),
    }


def _registro_db_a_cierre(reg: dict[str, Any]) -> tuple[CierreDia, dict[int, int], str, str]:
    p = reg.get("payload_json") or {}
    fecha_val: Optional[date] = None
    fecha_str = p.get("fecha") or reg.get("fecha")
    if fecha_str:
        with contextlib.suppress(ValueError):
            fecha_val = date.fromisoformat(str(fecha_str)[:10])

    ops = [
        Operacion(
            detalle=str(item.get("detalle", "")),
            monto=to_dec(item.get("monto")),
            medio_pago=item.get("medio_pago", "efectivo"),
            confianza=item.get("confianza", "alta"),
        )
        for item in (p.get("operaciones") or [])
    ]
    gastos = [
        GastoRetiro(
            tipo=item.get("tipo", "gasto"),
            concepto=str(item.get("concepto", "")),
            monto=to_dec(item.get("monto")),
        )
        for item in (p.get("gastos_retiros") or [])
    ]

    cierre = CierreDia(
        fecha=fecha_val,
        turno=normalizar_turno(p.get("turno") or reg.get("turno")),
        cajero=str(p.get("cajero") or reg.get("cajero") or "Celia"),
        saldo_inicial=to_dec(p.get("saldo_inicial")) if p.get("saldo_inicial") is not None else Decimal("0.00"),
        operaciones=ops,
        gastos_retiros=gastos,
        total_efectivo_cuaderno=to_dec(p["total_efectivo_cuaderno"])
        if p.get("total_efectivo_cuaderno") is not None
        else None,
        total_tarjetas_cuaderno=to_dec(p["total_tarjetas_cuaderno"])
        if p.get("total_tarjetas_cuaderno") is not None
        else None,
        total_gastos_cuaderno=to_dec(p["total_gastos_cuaderno"])
        if p.get("total_gastos_cuaderno") is not None
        else None,
        caja_total_cuaderno=to_dec(p["caja_total_cuaderno"]) if p.get("caja_total_cuaderno") is not None else None,
        campos_dudosos=list(p.get("campos_dudosos") or []),
        observaciones=str(p.get("observaciones") or ""),
    )

    arqueo_raw = p.get("arqueo_conteo") or {}
    arqueo_conteo = {denom: int(arqueo_raw.get(str(denom), arqueo_raw.get(denom, 0)) or 0) for denom in DENOMINACIONES_ARS}
    supervisor = str(p.get("supervisor") or reg.get("supervisor") or "Martín")
    comentario = str(p.get("comentario_auditor") or "")
    return cierre, arqueo_conteo, supervisor, comentario


def _leer_cierres_desde_local_excel() -> list[dict[str, Any]]:
    if not os.path.exists(ARCHIVO_MAESTRO_LOCAL):
        return []
    try:
        df = pd.read_excel(ARCHIVO_MAESTRO_LOCAL, skiprows=2)
        if df.empty or "Fecha" not in df.columns:
            return []
        df = df.dropna(subset=["Fecha"])
        df = df[~df["Fecha"].astype(str).str.contains("TOTAL|OPTICA|ÓPTICA", case=False, na=False)]
        registros: list[dict[str, Any]] = []
        for _, row in df.iterrows():
            fecha_str = str(row["Fecha"])[:10]
            turno_str = normalizar_turno(row["Turno"]) if "Turno" in df.columns else "Único"
            ef = to_dec(row.get("Efectivo"))
            mp_n = to_dec(row.get("MP Neto"))
            mp_i = to_dec(row.get("MP IVA (21%)"))
            bb_n = to_dec(row.get("BBVA Neto"))
            bb_i = to_dec(row.get("BBVA IVA (21%)"))
            prov = to_dec(row.get("Provincia"))
            gast = to_dec(row.get("Gastos"))
            ret = to_dec(row.get("Celia / Toti"))
            dif = to_dec(row.get("Diferencia Caja"))
            obs = str(row.get("Observaciones / Detalle") or "")

            cierre = CierreDia(
                fecha=date.fromisoformat(fecha_str),
                turno=turno_str,
                cajero="Celia",
                saldo_inicial=Decimal("0.00"),
                operaciones=[],
                gastos_retiros=[],
                observaciones=obs,
            )
            reg = _cierre_a_registro_db(cierre, usuario_actor="Migración Local")
            reg["efectivo"] = float(ef)
            reg["mp_neto"] = float(mp_n)
            reg["mp_iva"] = float(mp_i)
            reg["bbva_neto"] = float(bb_n)
            reg["bbva_iva"] = float(bb_i)
            reg["provincia"] = float(prov)
            reg["gastos"] = float(gast)
            reg["celia_toti"] = float(ret)
            reg["diferencia_caja"] = float(dif)
            reg["detalle_obs"] = obs
            registros.append(reg)
        return registros
    except Exception:
        return []


def obtener_todos_los_cierres() -> list[dict[str, Any]]:
    sb = obtener_cliente_supabase()
    if sb is not None:
        resp = sb.table("cierres_caja").select("*").order("fecha", desc=False).order("turno", desc=False).execute()
        datos = resp.data or []
        return sorted(datos, key=lambda r: (str(r.get("fecha", "")), str(r.get("turno", ""))))
    return sorted(_leer_cierres_desde_local_excel(), key=lambda r: (str(r.get("fecha", "")), str(r.get("turno", ""))))


def obtener_auditoria() -> list[dict[str, Any]]:
    sb = obtener_cliente_supabase()
    if sb is not None:
        resp = sb.table("auditoria_cierres").select("*").order("creado_en", desc=True).limit(500).execute()
        return resp.data or []
    return list(reversed(_leer_auditoria_local()))


def buscar_cierre_existente(fecha_iso: str, turno: str) -> Optional[dict[str, Any]]:
    clave = clave_cierre(fecha_iso, turno)
    for reg in obtener_todos_los_cierres():
        if str(reg.get("clave_cierre") or clave_cierre(str(reg.get("fecha", "")), str(reg.get("turno", "")))) == clave:
            return reg
    return None


def guardar_cierre_persistente(
    res: CierreDia,
    arqueo_conteo: Optional[dict[int, int]] = None,
    supervisor: str = "Martín",
    comentario_auditor: str = "",
    usuario_actor: str = "Usuario",
    confirmar_reemplazo: bool = False,
) -> tuple[bool, str]:
    reg_nuevo = _cierre_a_registro_db(
        res,
        arqueo_conteo=arqueo_conteo,
        supervisor=supervisor,
        comentario_auditor=comentario_auditor,
        usuario_actor=usuario_actor,
    )
    existente = buscar_cierre_existente(reg_nuevo["fecha"], reg_nuevo["turno"])
    if existente is not None and not confirmar_reemplazo:
        return False, f"Ya existe un cierre para la fecha {reg_nuevo['fecha']} y turno '{reg_nuevo['turno']}'."

    evento_auditoria = {
        "fecha_cierre": reg_nuevo["fecha"],
        "turno": reg_nuevo["turno"],
        "accion": "REEMPLAZO" if existente is not None else "ALTA",
        "usuario": usuario_actor,
        "creado_en": ahora_arg().isoformat(),
        "detalle": (
            f"Reemplazó cierre previo de {usuario_actor} ({existente.get('actualizado_en')})"
            if existente is not None
            else f"Alta de cierre ({reg_nuevo['cajero']})"
        ),
        "payload_anterior": existente.get("payload_json") if existente is not None else None,
        "payload_nuevo": reg_nuevo["payload_json"],
    }

    sb = obtener_cliente_supabase()
    if sb is not None:
        sb.table("cierres_caja").upsert(reg_nuevo, on_conflict="clave_cierre").execute()
        sb.table("auditoria_cierres").insert(evento_auditoria).execute()
    else:
        actuales = [
            r
            for r in _leer_cierres_desde_local_excel()
            if clave_cierre(str(r.get("fecha", "")), str(r.get("turno", ""))) != reg_nuevo["clave_cierre"]
        ]
        actuales.append(reg_nuevo)
        actuales = sorted(actuales, key=lambda r: (str(r.get("fecha", "")), str(r.get("turno", ""))))
        _guardar_historico_local_desde_registros(actuales)

        eventos = _leer_auditoria_local()
        eventos.append(evento_auditoria)
        _guardar_auditoria_local(eventos)

    return True, "Cierre guardado y registrado en auditoría."


# --- MOTOR DE EXCEL ÚNICO CONSOLIDADO (6 HOJAS) ---
def _configurar_impresion(ws, landscape: bool = True):
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if landscape else ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.center.text = "ÓPTICA KLIMT - CONSOLIDADO DE CIERRES DE CAJA"
    ws.oddHeader.center.size = 10
    ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = f"Generado: {ahora_arg().strftime('%d/%m/%Y %H:%M')}"
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


def _ajustar_columnas(ws, min_width: int = 14, max_width: int = 48):
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


def _guardar_historico_local_desde_registros(registros: list[dict[str, Any]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Historico_Mensual"
    ws["A1"] = "ÓPTICA KLIMT - CONSOLIDADO MENSUAL DE CAJA (CON DISCRIMINACIÓN DE IVA 21%)"
    ws["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    ws.row_dimensions[1].height = 24

    ws.append([])
    ws.append(COLUMNAS_HISTORICO)
    _estilizar_encabezado(ws, 3, len(COLUMNAS_HISTORICO))
    ws.freeze_panes = "A4"

    for idx_f, r_data in enumerate(registros):
        fila = [
            str(r_data.get("fecha", "")),
            sanitizar_celda_excel(r_data.get("turno", "Único")),
            float(to_dec(r_data.get("efectivo"))),
            float(to_dec(r_data.get("mp_neto"))),
            float(to_dec(r_data.get("mp_iva"))),
            float(to_dec(r_data.get("bbva_neto"))),
            float(to_dec(r_data.get("bbva_iva"))),
            float(to_dec(r_data.get("provincia"))),
            float(to_dec(r_data.get("gastos"))),
            float(to_dec(r_data.get("celia_toti"))),
            float(to_dec(r_data.get("diferencia_caja"))),
            sanitizar_celda_excel(r_data.get("detalle_obs", "")),
        ]
        ws.append(fila)
        r = ws.max_row
        fill_zebra = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
        for col in range(1, len(COLUMNAS_HISTORICO) + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDE_SUAVE
            if idx_f % 2 == 1:
                cell.fill = fill_zebra
            if col in (1, 2):
                cell.alignment = Alignment(horizontal="center")
            elif 3 <= col <= 11:
                cell.number_format = FORMATO_MONEDA_ARS
                if col in (3, 5, 7, 9, 10, 11):
                    cell.font = Font(bold=True)
            else:
                cell.alignment = Alignment(wrap_text=True)

    if registros:
        fila_inicio = 4
        fila_fin = ws.max_row
        fila_tot = fila_fin + 1
        ws.row_dimensions[fila_tot].height = 24
        c_lbl = ws.cell(row=fila_tot, column=1, value="TOTAL MES")
        c_lbl.font = Font(bold=True, size=11, color=COLOR_PRIMARIO)
        c_lbl.fill = PatternFill(start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid")
        c_lbl.border = BORDE_TOTAL
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in range(3, 12):
            letra = get_column_letter(col_idx)
            c_tot = ws.cell(row=fila_tot, column=col_idx, value=f"=SUM({letra}{fila_inicio}:{letra}{fila_fin})")
            c_tot.font = Font(bold=True, size=11, color=COLOR_PRIMARIO)
            c_tot.number_format = FORMATO_MONEDA_ARS
            c_tot.border = BORDE_TOTAL
            c_tot.fill = PatternFill(start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid")

    _ajustar_columnas(ws)
    _configurar_impresion(ws, landscape=True)

    buf = io.BytesIO()
    wb.save(buf)
    with bloqueo_archivo():
        guardar_bytes_atomico(ARCHIVO_MAESTRO_LOCAL, buf.getvalue())


def generar_excel_unico_bytes(
    registros: list[dict[str, Any]],
    eventos_auditoria: list[dict[str, Any]],
    etiqueta_rango: str,
) -> bytes:
    wb = Workbook()
    ws_res = wb.active
    ws_res.title = "Resumen_Mensual"

    ws_cierres = wb.create_sheet("Cierres")
    ws_ing = wb.create_sheet("Detalle_Ingresos")
    ws_egr = wb.create_sheet("Detalle_Egresos")
    ws_arq = wb.create_sheet("Arqueo")
    ws_aud = wb.create_sheet("Auditoria")

    # --- HOJA 2: CIERRES ---
    ws_cierres["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — CIERRES CONFIRMADOS ({etiqueta_rango})")
    ws_cierres["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    headers_cierres = [
        "Fecha",
        "Turno",
        "Cajero",
        "Supervisor",
        "Efectivo",
        "MP Neto",
        "MP IVA (21%)",
        "BBVA Neto",
        "BBVA IVA (21%)",
        "Provincia",
        "Gastos",
        "Celia / Toti",
        "Esperado Total",
        "Cuaderno Total",
        "Diferencia Caja",
        "Observaciones",
    ]
    ws_cierres.append([])
    ws_cierres.append(headers_cierres)
    _estilizar_encabezado(ws_cierres, 3, len(headers_cierres))
    ws_cierres.freeze_panes = "A4"

    for idx, r_data in enumerate(registros):
        cierre, arqueo_c, sup_c, _ = _registro_db_a_cierre(r_data)
        d = procesar_datos_dia(cierre, arqueo_conteo=arqueo_c)
        row_num = 4 + idx
        vals = [
            str(r_data.get("fecha", "")),
            sanitizar_celda_excel(r_data.get("turno", "Único")),
            sanitizar_celda_excel(r_data.get("cajero", "Celia")),
            sanitizar_celda_excel(r_data.get("supervisor") or sup_c),
            float(d["efectivo"]),
            float(d["mp_neto"]),
            float(d["mp_iva"]),
            float(d["bbva_neto"]),
            float(d["bbva_iva"]),
            float(d["provincia"]),
            float(d["gastos"]),
            float(d["celia_toti"]),
            float(d["total_general_esperado"]),
            float(d["caja_cuaderno"]),
            float(d["diferencia_cuaderno"]),
            sanitizar_celda_excel(d["detalle_obs"]),
        ]
        ws_cierres.append(vals)
        for col_i in range(1, len(headers_cierres) + 1):
            cell = ws_cierres.cell(row=row_num, column=col_i)
            cell.border = BORDE_SUAVE
            if idx % 2 == 1:
                cell.fill = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
            if 5 <= col_i <= 15:
                cell.number_format = FORMATO_MONEDA_ARS

    r_fin_cierres = 3 + max(len(registros), 1)
    r_tot_cierres = r_fin_cierres + 1
    ws_cierres.cell(row=r_tot_cierres, column=1, value="TOTAL").font = Font(bold=True, color=COLOR_PRIMARIO)
    for col_i in range(5, 16):
        letra = get_column_letter(col_i)
        c_t = ws_cierres.cell(row=r_tot_cierres, column=col_i, value=f"=SUM({letra}4:{letra}{r_fin_cierres})")
        c_t.font = Font(bold=True, color=COLOR_PRIMARIO)
        c_t.number_format = FORMATO_MONEDA_ARS
    for col_i in range(1, len(headers_cierres) + 1):
        ws_cierres.cell(row=r_tot_cierres, column=col_i).border = BORDE_TOTAL
        ws_cierres.cell(row=r_tot_cierres, column=col_i).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )
    ws_cierres.auto_filter.ref = f"A3:{get_column_letter(len(headers_cierres))}{r_fin_cierres}"
    _ajustar_columnas(ws_cierres)
    _configurar_impresion(ws_cierres, landscape=True)

    # --- HOJA 3: DETALLE_INGRESOS ---
    ws_ing["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — DETALLE DE INGRESOS ({etiqueta_rango})")
    ws_ing["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    headers_ing = ["Fecha", "Turno", "Detalle / Producto", "Medio de Pago", "Confianza", "Bruto", "Neto", "IVA (21%)"]
    ws_ing.append([])
    ws_ing.append(headers_ing)
    _estilizar_encabezado(ws_ing, 3, len(headers_ing))
    ws_ing.freeze_panes = "A4"

    ing_row_idx = 4
    for r_data in registros:
        cierre, _, _, _ = _registro_db_a_cierre(r_data)
        f_str = str(r_data.get("fecha", ""))
        t_str = normalizar_turno(r_data.get("turno"))
        for op in cierre.operaciones:
            aplica_iva = op.medio_pago in ("qr_mp", "payway_bbva", "transferencia")
            ws_ing.cell(row=ing_row_idx, column=1, value=f_str).alignment = Alignment(horizontal="center")
            ws_ing.cell(row=ing_row_idx, column=2, value=sanitizar_celda_excel(t_str)).alignment = Alignment(horizontal="center")
            ws_ing.cell(row=ing_row_idx, column=3, value=sanitizar_celda_excel(op.detalle))
            ws_ing.cell(row=ing_row_idx, column=4, value=sanitizar_celda_excel(ETIQUETAS_MEDIO_PAGO.get(op.medio_pago, op.medio_pago)))
            ws_ing.cell(row=ing_row_idx, column=5, value=op.confianza.upper()).alignment = Alignment(horizontal="center")
            ws_ing.cell(row=ing_row_idx, column=6, value=float(to_dec(op.monto))).number_format = FORMATO_MONEDA_ARS
            if aplica_iva:
                ws_ing.cell(row=ing_row_idx, column=7, value=f"=ROUND(F{ing_row_idx}/1.21,2)").number_format = FORMATO_MONEDA_ARS
                ws_ing.cell(row=ing_row_idx, column=8, value=f"=ROUND(F{ing_row_idx}-G{ing_row_idx},2)").number_format = FORMATO_MONEDA_ARS
            else:
                ws_ing.cell(row=ing_row_idx, column=7, value=f"=F{ing_row_idx}").number_format = FORMATO_MONEDA_ARS
                ws_ing.cell(row=ing_row_idx, column=8, value=0.0).number_format = FORMATO_MONEDA_ARS
            for col_i in range(1, 9):
                ws_ing.cell(row=ing_row_idx, column=col_i).border = BORDE_SUAVE
            ing_row_idx += 1

    r_fin_ing = max(ing_row_idx - 1, 4)
    r_tot_ing = r_fin_ing + 1
    ws_ing.cell(row=r_tot_ing, column=3, value="TOTAL INGRESOS").font = Font(bold=True, color=COLOR_PRIMARIO)
    for col_i in (6, 7, 8):
        letra = get_column_letter(col_i)
        c_t = ws_ing.cell(row=r_tot_ing, column=col_i, value=f"=SUM({letra}4:{letra}{r_fin_ing})")
        c_t.font = Font(bold=True, color=COLOR_PRIMARIO)
        c_t.number_format = FORMATO_MONEDA_ARS
    for col_i in range(1, 9):
        ws_ing.cell(row=r_tot_ing, column=col_i).border = BORDE_TOTAL
        ws_ing.cell(row=r_tot_ing, column=col_i).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )
    _ajustar_columnas(ws_ing)
    _configurar_impresion(ws_ing, landscape=True)

    # --- HOJA 4: DETALLE_EGRESOS ---
    ws_egr["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — DETALLE DE EGRESOS Y RETIROS ({etiqueta_rango})")
    ws_egr["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    headers_egr = ["Fecha", "Turno", "Tipo", "Concepto", "Monto (ARS)"]
    ws_egr.append([])
    ws_egr.append(headers_egr)
    _estilizar_encabezado(ws_egr, 3, len(headers_egr))
    ws_egr.freeze_panes = "A4"

    egr_row_idx = 4
    for r_data in registros:
        cierre, _, _, _ = _registro_db_a_cierre(r_data)
        f_str = str(r_data.get("fecha", ""))
        t_str = normalizar_turno(r_data.get("turno"))
        for g in cierre.gastos_retiros:
            ws_egr.cell(row=egr_row_idx, column=1, value=f_str).alignment = Alignment(horizontal="center")
            ws_egr.cell(row=egr_row_idx, column=2, value=sanitizar_celda_excel(t_str)).alignment = Alignment(horizontal="center")
            ws_egr.cell(row=egr_row_idx, column=3, value=g.tipo.upper()).alignment = Alignment(horizontal="center")
            ws_egr.cell(row=egr_row_idx, column=4, value=sanitizar_celda_excel(g.concepto))
            ws_egr.cell(row=egr_row_idx, column=5, value=float(to_dec(g.monto))).number_format = FORMATO_MONEDA_ARS
            for col_i in range(1, 6):
                ws_egr.cell(row=egr_row_idx, column=col_i).border = BORDE_SUAVE
            egr_row_idx += 1

    r_fin_egr = max(egr_row_idx - 1, 4)
    r_tot_egr = r_fin_egr + 1
    ws_egr.cell(row=r_tot_egr, column=4, value="TOTAL EGRESOS").font = Font(bold=True, color=COLOR_PRIMARIO)
    c_te = ws_egr.cell(row=r_tot_egr, column=5, value=f"=SUM(E4:E{r_fin_egr})")
    c_te.font = Font(bold=True, color=COLOR_PRIMARIO)
    c_te.number_format = FORMATO_MONEDA_ARS
    for col_i in range(1, 6):
        ws_egr.cell(row=r_tot_egr, column=col_i).border = BORDE_TOTAL
        ws_egr.cell(row=r_tot_egr, column=col_i).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )
    _ajustar_columnas(ws_egr)
    _configurar_impresion(ws_egr, landscape=False)

    # --- HOJA 5: ARQUEO ---
    ws_arq["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — ARQUEO DE BILLETES POR CIERRE ({etiqueta_rango})")
    ws_arq["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    headers_arq = ["Fecha", "Turno", "Denominación", "Cantidad", "Subtotal"]
    ws_arq.append([])
    ws_arq.append(headers_arq)
    _estilizar_encabezado(ws_arq, 3, len(headers_arq))
    ws_arq.freeze_panes = "A4"

    arq_row_idx = 4
    for r_data in registros:
        _, arqueo_c, _, _ = _registro_db_a_cierre(r_data)
        f_str = str(r_data.get("fecha", ""))
        t_str = normalizar_turno(r_data.get("turno"))
        for denom in DENOMINACIONES_ARS:
            cant = max(0, int(arqueo_c.get(denom, 0)))
            if cant == 0 and len(registros) > 0:
                continue
            ws_arq.cell(row=arq_row_idx, column=1, value=f_str).alignment = Alignment(horizontal="center")
            ws_arq.cell(row=arq_row_idx, column=2, value=sanitizar_celda_excel(t_str)).alignment = Alignment(horizontal="center")
            ws_arq.cell(row=arq_row_idx, column=3, value=denom).number_format = FORMATO_MONEDA_ARS
            ws_arq.cell(row=arq_row_idx, column=4, value=cant).alignment = Alignment(horizontal="center")
            ws_arq.cell(row=arq_row_idx, column=5, value=f"=C{arq_row_idx}*D{arq_row_idx}").number_format = FORMATO_MONEDA_ARS
            for col_i in range(1, 6):
                ws_arq.cell(row=arq_row_idx, column=col_i).border = BORDE_SUAVE
            arq_row_idx += 1

    r_fin_arq = max(arq_row_idx - 1, 4)
    r_tot_arq = r_fin_arq + 1
    ws_arq.cell(row=r_tot_arq, column=3, value="TOTAL ARQUEO").font = Font(bold=True, color=COLOR_PRIMARIO)
    ws_arq.cell(row=r_tot_arq, column=4, value=f"=SUM(D4:D{r_fin_arq})").font = Font(bold=True, color=COLOR_PRIMARIO)
    c_tarq = ws_arq.cell(row=r_tot_arq, column=5, value=f"=SUM(E4:E{r_fin_arq})")
    c_tarq.font = Font(bold=True, color=COLOR_PRIMARIO)
    c_tarq.number_format = FORMATO_MONEDA_ARS
    for col_i in range(1, 6):
        ws_arq.cell(row=r_tot_arq, column=col_i).border = BORDE_TOTAL
        ws_arq.cell(row=r_tot_arq, column=col_i).fill = PatternFill(
            start_color=COLOR_SECUNDARIO, end_color=COLOR_SECUNDARIO, fill_type="solid"
        )
    _ajustar_columnas(ws_arq)
    _configurar_impresion(ws_arq, landscape=False)

    # --- HOJA 6: AUDITORIA ---
    ws_aud["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — BITÁCORA DE AUDITORÍA ({etiqueta_rango})")
    ws_aud["A1"].font = Font(bold=True, size=13, color=COLOR_PRIMARIO)
    headers_aud = ["Fecha/Hora (ART)", "Usuario", "Acción", "Fecha Cierre", "Turno", "Detalle"]
    ws_aud.append([])
    ws_aud.append(headers_aud)
    _estilizar_encabezado(ws_aud, 3, len(headers_aud))
    ws_aud.freeze_panes = "A4"

    for idx_a, ev in enumerate(eventos_auditoria):
        r_a = 4 + idx_a
        ws_aud.cell(row=r_a, column=1, value=sanitizar_celda_excel(ev.get("creado_en", "")))
        ws_aud.cell(row=r_a, column=2, value=sanitizar_celda_excel(ev.get("usuario", "")))
        ws_aud.cell(row=r_a, column=3, value=sanitizar_celda_excel(ev.get("accion", "")))
        ws_aud.cell(row=r_a, column=4, value=sanitizar_celda_excel(ev.get("fecha_cierre", "")))
        ws_aud.cell(row=r_a, column=5, value=sanitizar_celda_excel(ev.get("turno", "")))
        ws_aud.cell(row=r_a, column=6, value=sanitizar_celda_excel(ev.get("detalle", ""))).alignment = Alignment(wrap_text=True)
        for col_i in range(1, 7):
            ws_aud.cell(row=r_a, column=col_i).border = BORDE_SUAVE
    _ajustar_columnas(ws_aud, min_width=16, max_width=70)
    _configurar_impresion(ws_aud, landscape=True)

    # --- HOJA 1: RESUMEN_MENSUAL ---
    ws_res["A1"] = sanitizar_celda_excel(f"ÓPTICA KLIMT — RESUMEN CONSOLIDADO ({etiqueta_rango})")
    ws_res["A1"].font = Font(bold=True, size=15, color=COLOR_PRIMARIO)
    ws_res.row_dimensions[1].height = 28

    ws_res.cell(row=3, column=1, value="CONCEPTO").font = Font(bold=True, color="FFFFFF")
    ws_res.cell(row=3, column=2, value="IMPORTE (ARS)").font = Font(bold=True, color="FFFFFF")
    _estilizar_encabezado(ws_res, 3, 2)

    resumen_filas = [
        ("Efectivo", f"=Cierres!E{r_tot_cierres}"),
        ("Mercado Pago Neto", f"=Cierres!F{r_tot_cierres}"),
        ("Mercado Pago IVA (21%)", f"=Cierres!G{r_tot_cierres}"),
        ("Payway BBVA Neto", f"=Cierres!H{r_tot_cierres}"),
        ("Payway BBVA IVA (21%)", f"=Cierres!I{r_tot_cierres}"),
        ("Posnet Provincia", f"=Cierres!J{r_tot_cierres}"),
        ("Total IVA (21%)", f"=B6+B8"),
        ("Gastos", f"=Cierres!K{r_tot_cierres}"),
        ("Retiros (Celia / Toti)", f"=Cierres!L{r_tot_cierres}"),
        ("Diferencia Caja (Cuaderno vs Esperado)", f"=Cierres!O{r_tot_cierres}"),
    ]

    for idx_r, (concepto, formula_str) in enumerate(resumen_filas):
        r = 4 + idx_r
        c1 = ws_res.cell(row=r, column=1, value=concepto)
        c2 = ws_res.cell(row=r, column=2, value=formula_str)
        c1.border = BORDE_SUAVE
        c2.border = BORDE_SUAVE
        c2.number_format = FORMATO_MONEDA_ARS
        if concepto in ("Total IVA (21%)", "Diferencia Caja (Cuaderno vs Esperado)"):
            c1.font = Font(bold=True, color=COLOR_PRIMARIO)
            c2.font = Font(bold=True, color=COLOR_PRIMARIO)

    ws_res.cell(row=16, column=1, value="RUBRO PARA GRÁFICO").font = Font(bold=True, color="FFFFFF")
    ws_res.cell(row=16, column=2, value="MONTO BRUTO (ARS)").font = Font(bold=True, color="FFFFFF")
    _estilizar_encabezado(ws_res, 16, 2)

    grafico_filas = [
        ("Efectivo", "=B4"),
        ("Mercado Pago", "=B5+B6"),
        ("Payway BBVA", "=B7+B8"),
        ("Provincia", "=B9"),
        ("Gastos + Retiros", "=B11+B12"),
    ]
    for idx_g, (rubro, f_rubro) in enumerate(grafico_filas):
        rg = 17 + idx_g
        ws_res.cell(row=rg, column=1, value=rubro).border = BORDE_SUAVE
        cg = ws_res.cell(row=rg, column=2, value=f_rubro)
        cg.border = BORDE_SUAVE
        cg.number_format = FORMATO_MONEDA_ARS

    chart = BarChart()
    chart.title = "Totales por Rubro"
    chart.y_axis.title = "ARS"
    chart.x_axis.title = "Rubro"
    data_ref = Reference(ws_res, min_col=2, min_row=16, max_row=21)
    cats_ref = Reference(ws_res, min_col=1, min_row=17, max_row=21)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend = None
    chart.width = 16
    chart.height = 9
    ws_res.add_chart(chart, "D4")

    _ajustar_columnas(ws_res, min_width=22, max_width=45)
    _configurar_impresion(ws_res, landscape=True)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def migrar_excel_anterior(file_bytes: bytes, usuario_actor: str = "Admin") -> tuple[int, list[str]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    hoja_objetivo = None
    for nombre in ("Historico_Mensual", "Cierre Mensual", wb.sheetnames[0]):
        if nombre in wb.sheetnames:
            hoja_objetivo = nombre
            break
    if not hoja_objetivo:
        raise ValueError("No se encontró ninguna hoja válida en el archivo Excel.")

    ws = wb[hoja_objetivo]
    filas_importadas = 0
    advertencias: list[str] = []

    for idx_row, row in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
        if not row or row[0] is None:
            continue
        primera_col = str(row[0]).strip()
        if not primera_col or primera_col.upper().startswith(("TOTAL", "ÓPTICA", "OPTICA")):
            continue

        fecha_norm: Optional[str] = None
        if isinstance(row[0], datetime):
            fecha_norm = row[0].strftime("%Y-%m-%d")
        elif isinstance(row[0], date):
            fecha_norm = row[0].strftime("%Y-%m-%d")
        else:
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                with contextlib.suppress(ValueError):
                    fecha_norm = datetime.strptime(primera_col[:10], fmt).strftime("%Y-%m-%d")
                    break

        if not fecha_norm:
            advertencias.append(f"Fila {idx_row}: fecha inválida '{primera_col}', omitida.")
            continue

        tiene_turno_col = len(row) >= 12 and isinstance(row[1], str) and not row[1].replace(".", "", 1).isdigit()
        offset = 1 if tiene_turno_col else 0
        turno_val = normalizar_turno(str(row[1])) if tiene_turno_col else "Único"

        ef = to_dec(row[1 + offset] if len(row) > 1 + offset else 0)
        mp_n = to_dec(row[2 + offset] if len(row) > 2 + offset else 0)
        mp_i = to_dec(row[3 + offset] if len(row) > 3 + offset else 0)
        bb_n = to_dec(row[4 + offset] if len(row) > 4 + offset else 0)
        bb_i = to_dec(row[5 + offset] if len(row) > 5 + offset else 0)
        prov = to_dec(row[6 + offset] if len(row) > 6 + offset else 0)
        gast = to_dec(row[7 + offset] if len(row) > 7 + offset else 0)
        ret = to_dec(row[8 + offset] if len(row) > 8 + offset else 0)
        dif = to_dec(row[9 + offset] if len(row) > 9 + offset else 0)
        obs = str(row[10 + offset] or "") if len(row) > 10 + offset else ""

        cierre_mig = CierreDia(
            fecha=date.fromisoformat(fecha_norm),
            turno=turno_val,
            cajero="Celia",
            saldo_inicial=Decimal("0.00"),
            operaciones=[],
            gastos_retiros=[],
            observaciones=obs,
        )
        reg = _cierre_a_registro_db(cierre_mig, usuario_actor=f"Migración ({usuario_actor})")
        reg["efectivo"] = float(ef)
        reg["mp_neto"] = float(mp_n)
        reg["mp_iva"] = float(mp_i)
        reg["bbva_neto"] = float(bb_n)
        reg["bbva_iva"] = float(bb_i)
        reg["provincia"] = float(prov)
        reg["gastos"] = float(gast)
        reg["celia_toti"] = float(ret)
        reg["diferencia_caja"] = float(dif)
        reg["detalle_obs"] = obs

        sb = obtener_cliente_supabase()
        if sb is not None:
            sb.table("cierres_caja").upsert(reg, on_conflict="clave_cierre").execute()
            sb.table("auditoria_cierres").insert(
                {
                    "fecha_cierre": fecha_norm,
                    "turno": turno_val,
                    "accion": "MIGRACION",
                    "usuario": usuario_actor,
                    "creado_en": ahora_arg().isoformat(),
                    "detalle": f"Importado desde hoja '{hoja_objetivo}' (fila {idx_row}).",
                    "payload_anterior": None,
                    "payload_nuevo": reg["payload_json"],
                }
            ).execute()
        else:
            actuales = [
                r
                for r in _leer_cierres_desde_local_excel()
                if clave_cierre(str(r.get("fecha", "")), str(r.get("turno", ""))) != reg["clave_cierre"]
            ]
            actuales.append(reg)
            actuales = sorted(actuales, key=lambda r: (str(r.get("fecha", "")), str(r.get("turno", ""))))
            _guardar_historico_local_desde_registros(actuales)
            eventos = _leer_auditoria_local()
            eventos.append(
                {
                    "fecha_cierre": fecha_norm,
                    "turno": turno_val,
                    "accion": "MIGRACION",
                    "usuario": usuario_actor,
                    "creado_en": ahora_arg().isoformat(),
                    "detalle": f"Importado desde hoja '{hoja_objetivo}' (fila {idx_row}).",
                    "payload_anterior": None,
                    "payload_nuevo": reg["payload_json"],
                }
            )
            _guardar_auditoria_local(eventos)

        filas_importadas += 1

    return filas_importadas, advertencias


# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

app_password = obtener_password_configurada()
if not app_password:
    st.error("Configuración incompleta: definí APP_PASSWORD en Streamlit Secrets o en variables de entorno.")
    st.stop()

if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        mostrar_logo(width=180)
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

# --- ESTADO DE PERSISTENCIA ---
usa_supabase = obtener_cliente_supabase() is not None

# --- MENÚ LATERAL ---
with st.sidebar:
    if not mostrar_logo(width=200):
        st.markdown("### Óptica Klimt")

    st.write("Tablero de Comando Gerencial")
    if usa_supabase:
        st.caption("Persistencia: Supabase (Activa)")
    else:
        st.caption("Persistencia: Local temporal (configurá Supabase para persistencia en nube)")
    st.divider()

    menu_opcion = st.radio(
        "Navegación",
        ["Tablero de Comando (Resumen)", "Procesar Cierre Diario"],
        label_visibility="collapsed",
    )

    st.divider()
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.pop("resultado", None)
        st.session_state.pop("guardado_maestro", None)
        st.rerun()

# --- VISTA 1: TABLERO DE COMANDO ---
if menu_opcion == "Tablero de Comando (Resumen)":
    st.title("Óptica Klimt — Tablero de Comando Gerencial")
    st.write("Vista ejecutiva del consolidado, discriminación de IVA (21%) y exportación del Excel único.")

    try:
        todos_cierres = obtener_todos_los_cierres()
    except Exception as e_db:
        st.error(f"Error al consultar la fuente de verdad: {e_db}")
        todos_cierres = []

    fechas_disponibles = sorted({str(r.get("fecha", "")) for r in todos_cierres if r.get("fecha")})
    meses_disponibles = sorted({f[:7] for f in fechas_disponibles if len(f) >= 7}, reverse=True)

    st.subheader("Filtro de Período para Tablero y Excel Único")
    modo_filtro = st.radio(
        "Seleccionar período:",
        ["Todo el histórico", "Por mes", "Rango de fechas"],
        horizontal=True,
    )

    cierres_filtrados = todos_cierres
    etiqueta_rango = "Todo el histórico"

    if modo_filtro == "Por mes" and meses_disponibles:
        mes_sel = st.selectbox("Mes (AAAA-MM):", meses_disponibles)
        cierres_filtrados = [r for r in todos_cierres if str(r.get("fecha", "")).startswith(mes_sel)]
        etiqueta_rango = f"Mes {mes_sel}"
    elif modo_filtro == "Rango de fechas" and fechas_disponibles:
        f_min = date.fromisoformat(fechas_disponibles[0][:10])
        f_max = date.fromisoformat(fechas_disponibles[-1][:10])
        c_f1, c_f2 = st.columns(2)
        desde_f = c_f1.date_input("Desde:", value=f_min)
        hasta_f = c_f2.date_input("Hasta:", value=f_max)
        d_iso = desde_f.strftime("%Y-%m-%d")
        h_iso = hasta_f.strftime("%Y-%m-%d")
        cierres_filtrados = [r for r in todos_cierres if d_iso <= str(r.get("fecha", ""))[:10] <= h_iso]
        etiqueta_rango = f"{d_iso} a {h_iso}"

    if cierres_filtrados:
        df_view = pd.DataFrame(
            [
                {
                    "Fecha": str(r.get("fecha", "")),
                    "Turno": r.get("turno", "Único"),
                    "Efectivo": r.get("efectivo", 0.0),
                    "MP Neto": r.get("mp_neto", 0.0),
                    "MP IVA (21%)": r.get("mp_iva", 0.0),
                    "BBVA Neto": r.get("bbva_neto", 0.0),
                    "BBVA IVA (21%)": r.get("bbva_iva", 0.0),
                    "Provincia": r.get("provincia", 0.0),
                    "Gastos": r.get("gastos", 0.0),
                    "Celia / Toti": r.get("celia_toti", 0.0),
                    "Diferencia Caja": r.get("diferencia_caja", 0.0),
                    "Observaciones / Detalle": r.get("detalle_obs", ""),
                }
                for r in cierres_filtrados
            ]
        )

        t_efectivo = sumar_serie_decimal(df_view, "Efectivo")
        t_mp_neto = sumar_serie_decimal(df_view, "MP Neto")
        t_mp_iva = sumar_serie_decimal(df_view, "MP IVA (21%)")
        t_bbva_neto = sumar_serie_decimal(df_view, "BBVA Neto")
        t_bbva_iva = sumar_serie_decimal(df_view, "BBVA IVA (21%)")
        t_provincia = sumar_serie_decimal(df_view, "Provincia")
        t_celia = sumar_serie_decimal(df_view, "Celia / Toti")

        t_mp_total = t_mp_neto + t_mp_iva
        t_bbva_total = t_bbva_neto + t_bbva_iva
        total_iva_mes = t_mp_iva + t_bbva_iva
        total_general_ventas = t_efectivo + t_mp_total + t_bbva_total + t_provincia

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ventas Totales", f"$ {total_general_ventas:,.2f}")
        m2.metric("Total IVA Acumulado (21%)", f"$ {total_iva_mes:,.2f}")
        m3.metric("Tarjetas / QR / Prov. Totales", f"$ {(t_mp_total + t_bbva_total + t_provincia):,.2f}")
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
        df_view = pd.DataFrame(columns=COLUMNAS_HISTORICO)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ventas Totales", "$ 0.00")
        m2.metric("Total IVA Acumulado (21%)", "$ 0.00")
        m3.metric("Tarjetas / QR / Prov. Totales", "$ 0.00")
        m4.metric("Retiros (Celia / Toti)", "$ 0.00")
        st.info("No hay cierres confirmados para el período seleccionado.")

    st.divider()
    st.subheader("Detalle Consolidado de Cierres")
    st.dataframe(df_view, use_container_width=True)

    eventos_aud = obtener_auditoria()
    excel_unico_bytes = generar_excel_unico_bytes(
        registros=cierres_filtrados,
        eventos_auditoria=eventos_aud,
        etiqueta_rango=etiqueta_rango,
    )
    nombre_unico = f"Optica_Klimt_Consolidado_{sanitizar_nombre_archivo(etiqueta_rango)}.xlsx"

    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1:
        st.download_button(
            label="Descargar Excel único (.xlsx)",
            data=excel_unico_bytes,
            file_name=nombre_unico,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_ex2:
        backup_bytes = generar_excel_unico_bytes(
            registros=todos_cierres,
            eventos_auditoria=eventos_aud,
            etiqueta_rango="Backup_Completo",
        )
        st.download_button(
            label="Descargar copia de seguridad (.xlsx)",
            data=backup_bytes,
            file_name=f"Backup_Optica_Klimt_{ahora_arg().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_ex3:
        with st.expander("Migrar Excel anterior (.xlsx) a la fuente de verdad"):
            excel_mig = st.file_uploader("Seleccionar cierre_mensual_optica.xlsx", type=["xlsx"], key="mig_xlsx")
            confirm_mig = st.checkbox("Confirmo importar este archivo a la fuente de verdad", key="confirm_mig")
            if excel_mig is not None and confirm_mig:
                try:
                    cant_ok, advs = migrar_excel_anterior(excel_mig.getvalue(), usuario_actor="Admin")
                    st.success(f"Migración completada: {cant_ok} cierre(s) importado(s).")
                    for a in advs:
                        st.warning(a)
                    st.rerun()
                except Exception as err_mig:
                    st.error(f"Error al migrar archivo: {err_mig}")

# --- VISTA 2: PROCESAR CIERRE DIARIO ---
elif menu_opcion == "Procesar Cierre Diario":
    st.title("Óptica Klimt — Procesamiento y Auditoría de Cierre Diario")
    st.write("Suba la fotografía de la libreta para extraer operaciones, discriminar IVA y guardar en el consolidado.")

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
                        st.success("Planilla procesada correctamente.")
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

    if "resultado" in st.session_state:
        res: CierreDia = st.session_state.resultado
        if res.fecha is None:
            res.fecha = ahora_arg().date()
        res.turno = normalizar_turno(res.turno)

        arqueo_dict: dict[int, int] = {denom: 0 for denom in DENOMINACIONES_ARS}
        sup_in = "Martín"
        comentario_in = res.observaciones or ""

        with st.expander("Turno, Arqueo de Billetes y Auditoría", expanded=True):
            col_m1, col_m2, col_m3 = st.columns(3)
            fecha_in = col_m1.date_input("Fecha contable:", value=res.fecha)
            turno_in = col_m2.text_input("Turno:", value=res.turno)
            cajero_in = col_m3.text_input("Cajero:", value=res.cajero or "Celia")

            col_s1, col_s2 = st.columns(2)
            sup_in = col_s1.text_input("Supervisor:", value=sup_in)
            comentario_in = col_s2.text_input("Observaciones del cierre:", value=comentario_in)

            res.fecha = fecha_in
            res.turno = normalizar_turno(turno_in)
            res.cajero = (cajero_in or "Celia").strip() or "Celia"
            res.observaciones = comentario_in

            st.markdown("**Conteo de Billetes (ARS):**")
            cols_billetes = st.columns(5)
            for idx_b, denom in enumerate(DENOMINACIONES_ARS):
                with cols_billetes[idx_b % 5]:
                    arqueo_dict[denom] = st.number_input(
                        f"$ {denom:,}", min_value=0, step=1, value=0, key=f"bill_{denom}"
                    )

        d = procesar_datos_dia(res, arqueo_conteo=arqueo_dict)

        st.divider()
        st.subheader(f"Resumen Auditado — Fecha: {res.fecha.strftime('%d/%m/%Y')} | Turno: {res.turno}")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Efectivo Ventas", f"$ {d['efectivo']:,.2f}")
        c2.metric("Mercado Pago", f"$ {d['mp_bruto']:,.2f}", f"IVA: $ {d['mp_iva']:,.2f}")
        c3.metric("Payway BBVA", f"$ {d['bbva_bruto']:,.2f}", f"IVA: $ {d['bbva_iva']:,.2f}")
        c4.metric("Provincia", f"$ {d['provincia']:,.2f}")
        c5.metric("Retiros (Celia/Toti)", f"$ {d['celia_toti']:,.2f}", f"Gastos: $ {d['gastos']:,.2f}")

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Esperado Efectivo Caja", f"$ {d['efectivo_esperado_caja']:,.2f}")
        k2.metric("Total Billetes Contados", f"$ {d['total_contado_billetes']:,.2f}", f"Dif: $ {d['diferencia_arqueo']:,.2f}")
        k3.metric("Total General Esperado", f"$ {d['total_general_esperado']:,.2f}")
        k4.metric(
            "Diferencia vs. Cuaderno",
            f"$ {d['diferencia_cuaderno']:,.2f}" if d["tiene_caja_cuaderno"] else "Sin dato en cuaderno",
        )

        if res.campos_dudosos:
            st.warning(f"Revisar valores marcados como poco claros: {', '.join(res.campos_dudosos)}")

        fecha_iso_actual = res.fecha.strftime("%Y-%m-%d")
        existente_colision = buscar_cierre_existente(fecha_iso_actual, res.turno)
        confirmar_colision = True
        if existente_colision is not None and not st.session_state.get("guardado_maestro", False):
            st.warning(
                f"Ya existe un cierre confirmado para **{fecha_iso_actual}** y turno **{res.turno}** "
                f"(actualizado: {existente_colision.get('actualizado_en', 's/d')})."
            )
            confirmar_colision = st.checkbox(
                "Confirmo reemplazar el cierre existente para esta fecha y turno (el anterior quedará guardado en Auditoría)",
                key="confirmar_reemplazo_cierre",
            )

        st.divider()
        st.subheader("Confirmación y Exportación")

        todos_act = obtener_todos_los_cierres()
        aud_act = obtener_auditoria()
        excel_consolidado_actual = generar_excel_unico_bytes(
            registros=todos_act,
            eventos_auditoria=aud_act,
            etiqueta_rango="Consolidado_Actual",
        )

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if not st.session_state.get("guardado_maestro", False):
                if st.button("Confirmar y Guardar en Fuente de Verdad", type="primary", use_container_width=True):
                    try:
                        ok, msg = guardar_cierre_persistente(
                            res=res,
                            arqueo_conteo=arqueo_dict,
                            supervisor=sup_in,
                            comentario_auditor=comentario_in,
                            usuario_actor="Usuario",
                            confirmar_reemplazo=confirmar_colision,
                        )
                        if ok:
                            st.session_state.guardado_maestro = True
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    except Exception as e_save:
                        st.error(f"Error al guardar el cierre: {e_save}")
            else:
                st.info("Este cierre ya fue confirmado y guardado en la fuente de verdad.")

        with col_dl2:
            nombre_unico_cierre = (
                f"Optica_Klimt_Cierres_{fecha_iso_actual}_{sanitizar_nombre_archivo(res.turno)}.xlsx"
            )
            st.download_button(
                label="Descargar Excel único (.xlsx)",
                data=excel_consolidado_actual,
                file_name=nombre_unico_cierre,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
