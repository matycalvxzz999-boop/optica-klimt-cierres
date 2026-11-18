import os
import io
import time
from decimal import Decimal
from datetime import date
from typing import Literal, Optional
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from PIL import Image

MedioPago = Literal["efectivo", "qr_mp", "payway_bbva", "posnet_provincia", "transferencia", "otro"]

class Operacion(BaseModel):
    detalle: str
    monto: Optional[Decimal] = None
    medio_pago: MedioPago
    confianza: Literal["alta", "media", "baja"] = "alta"

class GastoRetiro(BaseModel):
    tipo: Literal["gasto", "retiro"]
    concepto: str
    monto: Optional[Decimal] = None

class CierreDia(BaseModel):
    fecha: Optional[date] = None
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
Sos un asistente contable que transcribe cierres de caja manuscritos de una óptica en Argentina.
Analiza la foto de la libreta fila por fila de izquierda a derecha (Productos | EFECTIVO | TARJETA | MODO | ENTRADA | SALIDA).
REGLAS ESTRICTAS:
- Transcribí SOLO lo que se ve. Si un número es ilegible, poné null y marcalo en "campos_dudosos". NUNCA inventes números.
- Formato numérico argentino: punto para miles, coma para decimales. Devolvé los montos como números decimales puros (ej: 167000).
- Fechas en formato ISO (AAAA-MM-DD). El encabezado lateral izquierdo suele tener el día y fecha (ej: "Lunes 16/11") y la caja inicial (ej: "CAJA $ 16.000"). Si falta el año, asumí 2026.
- CLASIFICACIÓN DE MEDIO DE PAGO (MIRAR SIEMPRE LA COLUMNA "MODO"):
  - Si el monto está en columna EFECTIVO -> "efectivo".
  - Si el monto está en columna TARJETA, mirá qué dice al lado en columna MODO:
    - Si dice "QR MP", "QR" o "MP" -> "qr_mp".
    - Si dice "Payway" -> "payway_bbva".
    - Si dice "posnet B PCIA", "B PCIA", "PCIA" o "Provincia" -> "posnet_provincia".
    - Si no aclara modo pero está en TARJETA -> "payway_bbva".
- SALIDAS (columna derecha "SALIDA"): incluí cada gasto (ej: Limp, Kiosco) y cada retiro ("Retiro de caja") en "gastos_retiros".
- Copiá los subtotales inferiores (total efectivo, total tarjetas/MP y CAJA TOTAL) tal como Celia los escribió.
Devolvé la respuesta ÚNICAMENTE en formato JSON válido que cumpla con el esquema solicitado.
"""

def extraer_cierre_gemini(img_path_or_bytes) -> CierreDia:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Falta configurar GEMINI_API_KEY en el archivo .env")
    
    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=25000)
    )
    
    if isinstance(img_path_or_bytes, bytes):
        img = Image.open(io.BytesIO(img_path_or_bytes))
    else:
        img = Image.open(img_path_or_bytes)

    img.thumbnail((1600, 1600))

    # Actualizamos al modelo estable actual de Google
    modelos_candidatos = [
        "gemini-2.5-flash",
    ]

    ultimo_error = None
    for modelo in modelos_candidatos:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=[img, SYSTEM_PROMPT + "\nDevolvé el resultado estrictamente en formato JSON esquema CierreDia."],
                config={
                    "response_mime_type": "application/json",
                    "response_schema": CierreDia,
                },
            )
            return CierreDia.model_validate_json(response.text)
        except Exception as e:
            ultimo_error = e
            if "404" in str(e) or "NOT_FOUND" in str(e):
                continue
            raise e
            
    raise RuntimeError(f"No se pudo procesar con el modelo de Gemini. Último error: {ultimo_error}")
