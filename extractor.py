import io
import logging
import os
import random
import time
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Literal, Optional
from zoneinfo import ZoneInfo

from google import genai
from google.genai import types
from PIL import Image
from pydantic import BaseModel, Field, field_validator

logger = logging.getLogger("optica_klimt.extractor")

MedioPago = Literal[
    "efectivo",
    "qr_mp",
    "payway_bbva",
    "posnet_provincia",
    "transferencia",
    "otro",
]


class Operacion(BaseModel):
    detalle: str = Field(description="Código de producto, número de boleta o descripción de la venta")
    monto: Optional[Decimal] = Field(default=None, description="Monto exacto en ARS sin símbolos")
    medio_pago: MedioPago
    confianza: Literal["alta", "media", "baja"] = "alta"

    @field_validator("monto", mode="before")
    @classmethod
    def validar_monto_op(cls, v):
        if v is None or v == "":
            return None
        try:
            dec = Decimal(str(v).replace("$", "").strip())
            return dec
        except (InvalidOperation, ValueError):
            return None


class GastoRetiro(BaseModel):
    tipo: Literal["gasto", "retiro"]
    concepto: str
    monto: Optional[Decimal] = None

    @field_validator("monto", mode="before")
    @classmethod
    def validar_monto_gasto(cls, v):
        if v is None or v == "":
            return None
        try:
            dec = Decimal(str(v).replace("$", "").strip())
            return dec
        except (InvalidOperation, ValueError):
            return None


class CierreDia(BaseModel):
    fecha: Optional[date] = None
    turno: Optional[str] = Field(default="Único", description="Turno si figura en la hoja, por defecto Único")
    cajero: Optional[str] = Field(default="Celia", description="Nombre de quien confecciona la planilla")
    saldo_inicial: Optional[Decimal] = None
    operaciones: list[Operacion] = Field(default_factory=list)
    gastos_retiros: list[GastoRetiro] = Field(default_factory=list)
    total_efectivo_cuaderno: Optional[Decimal] = None
    total_tarjetas_cuaderno: Optional[Decimal] = None
    total_gastos_cuaderno: Optional[Decimal] = None
    caja_total_cuaderno: Optional[Decimal] = None
    campos_dudosos: list[str] = Field(default_factory=list)
    observaciones: Optional[str] = None


SYSTEM_PROMPT = """
Sos un auditor contable experto que transcribe cierres de caja manuscritos de "Óptica Klimt" en Argentina.
Analizá la fotografía de la libreta fila por fila, de arriba hacia abajo y de izquierda a derecha:
Columnas habituales: [Productos / Boleta | EFECTIVO | TARJETA | MODO | ENTRADA | SALIDA].

REGLAS ESTRICTAS DE EXTRACCIÓN:
1. Transcribí ÚNICAMENTE los valores visibles en la imagen. Si un número está tachado, borroneado o es ilegible, devolvé null y describí la fila en "campos_dudosos". JAMÁS inventes ni estimes cifras.
2. Formato numérico argentino: en la libreta el punto separa miles (ej: $37.000 = 37000, $165.000 = 165000). Devolvé números decimales puros sin puntos de miles ni signo $.
3. Fecha contable en formato ISO (AAAA-MM-DD). Revisá el margen izquierdo vertical (ej: "17-11-26" -> "2026-11-17"). Si falta el año, asumí {ANIO_ACTUAL}.
4. Saldo inicial ("CAJA $13.000" en el margen izquierdo): cargalo en "saldo_inicial".
5. CLASIFICACIÓN OBLIGATORIA DE MEDIO DE PAGO (mirar columna TARJETA y columna MODO):
  - Si el monto está en la columna EFECTIVO -> medio_pago = "efectivo".
  - Si el monto está en la columna TARJETA, leé la aclaración en MODO:
    * "QR MP", "QR", "MP", "DEBITO MP", "MP TRANS" -> "qr_mp".
    * "Payway", "BBVA", "Visa", "Master" -> "payway_bbva".
    * "posnet B PCIA", "B PCIA", "PCIA", "Provincia" -> "posnet_provincia".
    * "Transferencia" (bancaria directa) -> "transferencia".
    * Si está en TARJETA y no especifica modo -> "payway_bbva".
6. SALIDAS (columna derecha "SALIDA"):
  - Registrá cada ítem en "gastos_retiros".
  - Si dice "RETIRO DE CAJA", "Celia", "Toti" -> tipo = "retiro".
  - Si dice "Comida", "Kiosco", "Limp", etc. -> tipo = "gasto".
7. SUBTOTALES DEL CUADERNO:
  - Copiá en total_efectivo_cuaderno, total_tarjetas_cuaderno y caja_total_cuaderno los totales escritos al pie de la hoja (si hay un número tachado y corregido abajo, tomá el valor final vigente, ej: $330.000).
8. Devolvé ÚNICAMENTE JSON válido ajustado al esquema CierreDia.
"""


def _obtener_api_key() -> Optional[str]:
    api_key = os.getenv("GEMINI_API_KEY")
    if api_key:
        return api_key.strip()
    try:
        import streamlit as st

        secret_key = st.secrets.get("GEMINI_API_KEY")
        if secret_key:
            return str(secret_key).strip()
    except Exception:
        pass
    return None


def _preparar_imagen(img_path_or_bytes: bytes | str) -> Image.Image:
    if isinstance(img_path_or_bytes, bytes):
        with Image.open(io.BytesIO(img_path_or_bytes)) as raw_img:
            img = raw_img.copy()
    else:
        with Image.open(img_path_or_bytes) as raw_img:
            img = raw_img.copy()

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    img.thumbnail((1800, 1800), Image.Resampling.LANCZOS)
    return img


def extraer_cierre_gemini(img_path_or_bytes: bytes | str) -> CierreDia:
    api_key = _obtener_api_key()
    if not api_key:
        raise ValueError("Falta configurar GEMINI_API_KEY en los Secrets de Streamlit o archivo .env.")

    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=45000),
    )

    img = _preparar_imagen(img_path_or_bytes)
    anio_actual = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).year
    prompt = SYSTEM_PROMPT.replace("{ANIO_ACTUAL}", str(anio_actual))

    modelos_candidatos = [
        "gemini-3.8-flash",
        "gemini-3.1-pro-preview",
    ]

    ultimo_error: Optional[Exception] = None
    hubo_error_transitorio = False

    for modelo in modelos_candidatos:
        max_reintentos = 3
        for intento in range(max_reintentos):
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=[
                        img,
                        prompt
                        + "\nDevolvé el resultado estrictamente en formato JSON bajo el esquema CierreDia.",
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=CierreDia,
                        temperature=0.0,
                    ),
                )
                if not response.text:
                    raise ValueError(f"El modelo {modelo} devolvió una respuesta vacía.")
                return CierreDia.model_validate_json(response.text)

            except Exception as e:
                ultimo_error = e
                err_str = str(e)
                if any(
                    code in err_str
                    for code in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "504", "DEADLINE")
                ):
                    hubo_error_transitorio = True
                    jitter = random.uniform(0.1, 0.7)
                    espera = (2.0 * (intento + 1)) + jitter
                    logger.warning("Reintento %s/%s para modelo %s tras error transitorio: %s", intento + 1, max_reintentos, modelo, err_str)
                    time.sleep(espera)
                    continue
                if "404" in err_str or "NOT_FOUND" in err_str:
                    logger.warning("Modelo %s no disponible (404), pasando al siguiente candidato.", modelo)
                    break
                raise RuntimeError(f"Error al interpretar la planilla con {modelo}: {err_str}") from e

    if hubo_error_transitorio:
        raise RuntimeError(
            f"Los servidores de IA están experimentando alta demanda temporal (503). "
            f"Aguardá 15 segundos y volvé a presionar 'Procesar Cierre con IA'. Detalle: {ultimo_error}"
        )
    raise RuntimeError(
        f"Ningún modelo respondió ({', '.join(modelos_candidatos)}). Verificá GEMINI_API_KEY, "
        f"sus permisos y los nombres de modelo. Detalle: {ultimo_error}"
    )
