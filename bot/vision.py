"""Lee boletas y facturas con Gemini."""

from __future__ import annotations

import json
import os
import re
from datetime import date, datetime

import httpx

from bot.catalogo import ALIASES, TIPOS, categorias_de, categoria_respaldo

MODELO = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash")
RESPALDO = ("gemini-3.5-flash", "gemini-3.7-flash", "gemini-3.8-flash")


class VisionError(RuntimeError):
    pass


def _clave() -> str:
    clave = os.environ.get("GEMINI_API_KEY", "").strip()
    if not clave:
        raise VisionError(
            "Falta GEMINI_API_KEY en el archivo .env. "
            "La sacas gratis en https://aistudio.google.com/apikey"
        )
    return clave


def _entero(valor) -> int | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return int(round(valor))
    texto = str(valor).strip()
    texto = texto.replace("$", "").replace(" ", "")
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", texto):
        texto = texto.replace(".", "")
    else:
        texto = texto.replace(".", "").replace(",", "")
    if not texto or not re.fullmatch(r"-?\d+", texto):
        return None
    return int(texto)


def _fecha(valor) -> date | None:
    if not valor:
        return None
    texto = str(valor).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def _elegir_categoria(tipo: str, candidata: str | None) -> str:
    opciones = categorias_de(tipo)
    if not candidata:
        return categoria_respaldo(tipo)
    alias = ALIASES.get(candidata.strip().lower())
    if alias in opciones:
        return alias
    if candidata in opciones:
        return candidata
    baja = candidata.lower()
    for op in opciones:
        if op.lower() == baja or baja in op.lower() or op.lower() in baja:
            return op
    return categoria_respaldo(tipo)


def normalizar(datos: dict) -> dict:
    tipo = datos.get("tipo") or "Gasto variable"
    if tipo not in TIPOS:
        tipo = "Gasto variable"
    neto = _entero(datos.get("neto"))
    total = _entero(datos.get("total"))
    iva = _entero(datos.get("iva"))
    afecto = datos.get("afecto_iva")
    if afecto not in ("Sí", "No"):
        afecto = "Sí" if (iva or 0) > 0 else None
    if neto is None and total is not None:
        if afecto == "Sí":
            neto = int(round(total / 1.19))
        else:
            neto = total
            afecto = afecto or "No"
    if afecto not in ("Sí", "No"):
        afecto = "Sí"
    return {
        "fecha": _fecha(datos.get("fecha")) or date.today(),
        "tipo": tipo,
        "categoria": _elegir_categoria(tipo, datos.get("categoria")),
        "contraparte": (datos.get("contraparte") or "").strip(),
        "documento": str(datos.get("documento") or "").strip(),
        "descripcion": (datos.get("descripcion") or "").strip(),
        "neto": neto,
        "afecto_iva": afecto,
        "marca": (datos.get("marca") or "").strip(),
        "modelo": (datos.get("modelo") or "").strip(),
        "anio_auto": _entero(datos.get("anio_auto")),
        "patente": (datos.get("patente") or "").strip().upper(),
        "precio_venta": _entero(datos.get("precio_venta")) or total,
        "costo_neto": _entero(datos.get("costo_neto")),
        "confianza": datos.get("confianza") or "media",
        "notas": (datos.get("notas") or "").strip(),
    }


def _prompt(pista: str = "") -> str:
    extra = f"\nEl usuario aclara: {pista}" if pista else ""
    return f"""Eres el asistente contable de una automotora en Chile.
Lee la imagen o el PDF de una boleta, factura o comprobante de pago.
Responde SOLO un JSON con estas claves:
fecha (YYYY-MM-DD),
tipo (uno de: Gasto fijo, Gasto variable, Otro ingreso, Taller, Repuestos, Venta auto),
categoria (un concepto corto en español, por ejemplo Electricidad, Combustible, Arriendo local / patio, Gestión de financiamiento, Usado),
contraparte (quien emite o el cliente),
documento (número de boleta o factura),
descripcion (una línea),
neto (entero, pesos chilenos SIN IVA),
iva (entero o null),
total (entero, con IVA si corresponde),
afecto_iva ("Sí" si el documento muestra IVA recuperable de factura, "No" si es boleta de honorarios, sueldo, arriendo de inmueble, particular o no se ve IVA),
marca, modelo, anio_auto, patente, precio_venta, costo_neto (solo si es un auto; si no, null),
confianza ("alta", "media" o "baja"),
notas (qué no pudiste leer).
Si solo ves el total con IVA, calcula neto = round(total/1.19).
No inventes montos: si no se lee el número, deja neto en null.{extra}"""


def _modelos() -> list[str]:
    vistos = []
    for nombre in (MODELO, *RESPALDO):
        if nombre and nombre not in vistos:
            vistos.append(nombre)
    return vistos


async def leer_documento(data: bytes, mime: str, pista: str = "") -> dict:
    cuerpo = {
        "contents": [
            {
                "parts": [
                    {"text": _prompt(pista)},
                    {"inlineData": {"mimeType": mime, "data": _b64(data)}},
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }
    clave = _clave()
    async with httpx.AsyncClient(timeout=90) as client:
        for modelo in _modelos():
            url = (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                f"{modelo}:generateContent?key={clave}"
            )
            resp = await client.post(url, json=cuerpo)
            if resp.status_code == 200:
                break
            if resp.status_code not in (429, 503):
                raise VisionError(
                    "No pude leer el documento. Mándalo de nuevo en un momento."
                )
        else:
            raise VisionError(
                "El lector está saturado en este momento. Espera un minuto y manda la misma foto otra vez."
            )
    payload = resp.json()
    try:
        texto = payload["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError) as exc:
        raise VisionError("Gemini no devolvió una lectura del documento.") from exc
    texto = texto.strip()
    if texto.startswith("```"):
        texto = re.sub(r"^```(?:json)?|```$", "", texto, flags=re.MULTILINE).strip()
    try:
        datos = json.loads(texto)
    except json.JSONDecodeError as exc:
        raise VisionError("No pude interpretar la lectura del documento.") from exc
    if isinstance(datos, list):
        datos = datos[0] if datos else {}
    return normalizar(datos)


def _b64(data: bytes) -> str:
    import base64

    return base64.b64encode(data).decode("ascii")
