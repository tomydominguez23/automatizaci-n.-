"""Responde preguntas sobre la planilla de Drive."""

from __future__ import annotations

import httpx

from bot.nube import _libro, configurada
from bot.planilla import PlanillaError
from bot.vision import MODELO, RESPALDO, VisionError, _clave


def _filas(ws, rango: str, limite: int = 40) -> list[str]:
    valores = ws.get(rango, value_render_option="FORMATTED_VALUE")
    lineas = []
    for fila in valores:
        if not any(str(c).strip() for c in fila):
            continue
        lineas.append(" | ".join(str(c) for c in fila))
        if len(lineas) >= limite:
            break
    return lineas


def contexto() -> str:
    if not configurada():
        raise PlanillaError("La planilla de Drive todavía no está vinculada.")
    libro = _libro()
    partes = ["Configuración", *_filas(libro.worksheet("Configuracion"), "A4:C16", 16)]
    partes.append("Estado de resultados (columnas: concepto y meses enero a diciembre, más el total)")
    partes.extend(_filas(libro.worksheet("Estado_Resultados"), "A4:N32", 40))
    partes.append("Últimos comprobantes")
    partes.extend(_filas(libro.worksheet("Comprobantes"), "A5:O60", 45))
    return "\n".join(partes)


async def responder(pregunta: str) -> str:
    datos = contexto()
    prompt = (
        "Eres el asistente de la automotora. Responde en español, breve, "
        "solo con estos datos de la planilla. Los montos son pesos chilenos, netos de IVA salvo que diga lo contrario. "
        "Si el dato no está, dilo. No inventes cifras.\n\n"
        f"PREGUNTA: {pregunta}\n\nDATOS:\n{datos}"
    )
    clave = _clave()
    modelos = []
    for nombre in (MODELO, *RESPALDO):
        if nombre not in modelos:
            modelos.append(nombre)
    async with httpx.AsyncClient(timeout=60) as client:
        for modelo in modelos:
            url = (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                f"{modelo}:generateContent?key={clave}"
            )
            resp = await client.post(
                url,
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.2},
                },
            )
            if resp.status_code == 200:
                break
            if resp.status_code not in (429, 503):
                raise VisionError("No pude consultar la planilla. Inténtalo en un momento.")
        else:
            raise VisionError("El lector está saturado. Pregunta de nuevo en un minuto.")
    try:
        return resp.json()["candidates"][0]["content"]["parts"][0]["text"].replace("**", "").strip()
    except (KeyError, IndexError) as exc:
        raise VisionError("No obtuve una respuesta de la planilla.") from exc
