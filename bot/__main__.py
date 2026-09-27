"""Bot de Telegram: foto del comprobante → fila en la planilla."""

from __future__ import annotations

import json
import os
from datetime import date, datetime
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from bot.catalogo import TIPOS, categorias_de
from bot.consulta import responder
from bot.nube import configurada as drive_listo
from bot.nube import huella_ya_cargada as huella_en_drive
from bot.nube import registrar as registrar_en_drive
from bot.planilla import (
    Movimiento,
    PlanillaAbierta,
    PlanillaError,
    afecto_sugerido,
    hash_archivo,
    huella_ya_cargada,
    preparar_planilla,
    registrar,
)
from bot.vision import VisionError, leer_documento

ROOT = Path(__file__).resolve().parents[1]


def _directorio_datos() -> Path:
    explicito = os.environ.get("DATA_DIR", "").strip()
    if explicito:
        return Path(explicito)
    persistente = Path("/data")
    if persistente.is_dir() and os.access(persistente, os.W_OK):
        return persistente
    return ROOT / "data"


DATA = _directorio_datos()
OWNER = DATA / "dueno.json"
POR_PAGINA = 6


def _dueno_permitido(user_id: int) -> bool:
    crudos = os.environ.get("TELEGRAM_ALLOWED_IDS", "").strip()
    if crudos:
        ids = {p.strip() for p in crudos.split(",") if p.strip()}
        return str(user_id) in ids
    if not OWNER.exists():
        return True
    guardado = json.loads(OWNER.read_text(encoding="utf-8")).get("user_id")
    return guardado == user_id


def _fijar_dueno(user_id: int) -> None:
    if os.environ.get("TELEGRAM_ALLOWED_IDS", "").strip():
        return
    if OWNER.exists():
        return
    DATA.mkdir(parents=True, exist_ok=True)
    OWNER.write_text(json.dumps({"user_id": user_id}), encoding="utf-8")


def _clp(valor) -> str:
    if valor is None:
        return "no se leyó"
    return "$" + f"{int(valor):,}".replace(",", ".")


def _fecha_txt(valor: date) -> str:
    return valor.strftime("%d/%m/%Y")


def _texto_borrador(b: dict) -> str:
    lineas = [
        "Así leí el comprobante:",
        f"Fecha: {_fecha_txt(b['fecha'])}",
        f"Tipo: {b['tipo']}",
        f"Categoría: {b['categoria']}",
        f"Quién: {b['contraparte'] or '—'}",
        f"Documento: {b['documento'] or '—'}",
        f"Detalle: {b['descripcion'] or '—'}",
        f"Neto (sin IVA): {_clp(b['neto'])}",
        f"Afecto a IVA: {b['afecto_iva']}",
    ]
    if b["tipo"] == "Venta auto":
        lineas.append(
            f"Auto: {b.get('marca') or ''} {b.get('modelo') or ''} {b.get('patente') or ''}".strip()
        )
        lineas.append(f"Precio con IVA: {_clp(b.get('precio_venta'))}")
    if b.get("notas"):
        lineas.append(f"Ojo: {b['notas']}")
    if b.get("neto") is None:
        lineas.append("Escribe el neto sin IVA (solo el número) para poder cargarlo.")
    else:
        lineas.append("Si el neto está mal, escribe el número correcto. Si está bien, confírmalo.")
    return "\n".join(lineas)


def _teclado(b: dict) -> InlineKeyboardMarkup:
    filas = []
    if b.get("neto") is not None:
        filas.append([InlineKeyboardButton("Cargarlo a la planilla", callback_data="ok")])
    filas.append(
        [
            InlineKeyboardButton("Cambiar tipo", callback_data="tipo"),
            InlineKeyboardButton("Cambiar categoría", callback_data="cat:0"),
        ]
    )
    iva = "Marcar sin IVA" if b.get("afecto_iva") == "Sí" else "Marcar con IVA"
    filas.append(
        [
            InlineKeyboardButton(iva, callback_data="iva"),
            InlineKeyboardButton("Cancelar", callback_data="no"),
        ]
    )
    return InlineKeyboardMarkup(filas)


def _teclado_tipos() -> InlineKeyboardMarkup:
    filas = []
    for i in range(0, len(TIPOS), 2):
        fila = []
        for tipo in TIPOS[i : i + 2]:
            fila.append(InlineKeyboardButton(tipo, callback_data=f"t:{TIPOS.index(tipo)}"))
        filas.append(fila)
    filas.append([InlineKeyboardButton("Volver", callback_data="back")])
    return InlineKeyboardMarkup(filas)


def _teclado_categorias(tipo: str, pagina: int) -> InlineKeyboardMarkup:
    opciones = categorias_de(tipo)
    inicio = pagina * POR_PAGINA
    trozo = opciones[inicio : inicio + POR_PAGINA]
    filas = [[InlineKeyboardButton(nombre, callback_data=f"c:{inicio + i}")] for i, nombre in enumerate(trozo)]
    nav = []
    if inicio > 0:
        nav.append(InlineKeyboardButton("Anterior", callback_data=f"cat:{pagina - 1}"))
    if inicio + POR_PAGINA < len(opciones):
        nav.append(InlineKeyboardButton("Siguiente", callback_data=f"cat:{pagina + 1}"))
    if nav:
        filas.append(nav)
    filas.append([InlineKeyboardButton("Volver", callback_data="back")])
    return InlineKeyboardMarkup(filas)


def _serializar(b: dict) -> dict:
    copia = dict(b)
    if isinstance(copia.get("fecha"), date):
        copia["fecha"] = copia["fecha"].isoformat()
    return copia


def _borrador(context: ContextTypes.DEFAULT_TYPE) -> dict | None:
    crudo = context.user_data.get("borrador")
    if not crudo:
        return None
    if isinstance(crudo.get("fecha"), str):
        crudo["fecha"] = datetime.strptime(crudo["fecha"], "%Y-%m-%d").date()
    return crudo


async def _guardar_borrador(context, borrador: dict) -> None:
    context.user_data["borrador"] = _serializar(borrador)


async def _rechazar(update: Update) -> bool:
    user = update.effective_user
    if user and _dueno_permitido(user.id):
        return False
    if update.message:
        await update.message.reply_text("Este bot solo carga la planilla de la automotora.")
    elif update.callback_query:
        await update.callback_query.answer("No autorizado", show_alert=True)
    return True


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if await _rechazar(update):
        return
    _fijar_dueno(update.effective_user.id)
    await update.message.reply_text(
        "Mándame una foto o el PDF de la boleta, la factura o el comprobante.\n"
        "Te muestro lo que leí y, cuando confirmes, lo dejo en la planilla.\n"
        "Gastos fijos y variables se suman solos en el mes. "
        "También puedes preguntarme por la planilla, por ejemplo: "
        "¿cuánto llevo de gastos en septiembre? o ¿cuál fue la utilidad de agosto?"
    )


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if await _rechazar(update):
        return
    await update.message.reply_text(
        "1. Saca la foto del documento, derecha y con el total visible.\n"
        "2. Revisa neto, categoría y si tiene IVA.\n"
        "3. Toca «Cargarlo a la planilla».\n"
        "Si el monto está mal, escribe solo el neto, por ejemplo 45000.\n"
        "También puedes escribir una pregunta sobre ingresos, gastos o utilidad.\n"
        "Cierra el Excel local solo si el bot está guardando en el computador."
    )


async def cmd_cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("borrador", None)
    await update.message.reply_text("Listo, descarté el comprobante que estaba pendiente.")


def _aplicar_lectura(lectura: dict, huella: str, archivo: str) -> dict:
    lectura["huella"] = huella
    lectura["archivo"] = archivo
    lectura["afecto_iva"] = afecto_sugerido(lectura["categoria"], lectura.get("afecto_iva"))
    return lectura


async def _recibir_archivo(update: Update, context: ContextTypes.DEFAULT_TYPE, data: bytes, mime: str, nombre: str) -> None:
    if await _rechazar(update):
        return
    _fijar_dueno(update.effective_user.id)
    espera = await update.message.reply_text("Estoy leyendo el documento…")
    huella = hash_archivo(data)
    try:
        repetida = (
            huella_en_drive(huella)
            if drive_listo()
            else huella_ya_cargada(context.bot_data["planilla"], huella)
        )
        if repetida:
            await espera.edit_text("Ese archivo ya está en la planilla. No lo cargué de nuevo.")
            return
        pista = ""
        lectura = await leer_documento(data, mime, pista)
    except (VisionError, PlanillaError) as exc:
        await espera.edit_text(str(exc))
        return
    except Exception:
        await espera.edit_text(
            "No pude revisar si esa foto ya estaba cargada. La vuelvo a leer: mándala otra vez."
        )
        return
    carpeta = DATA / "comprobantes" / lectura["fecha"].strftime("%Y-%m")
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"{huella}{Path(nombre).suffix.lower() or '.jpg'}"
    destino.write_bytes(data)
    borrador = _aplicar_lectura(lectura, huella, str(destino.relative_to(ROOT)))
    await _guardar_borrador(context, borrador)
    vivo = _borrador(context)
    await espera.edit_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))


async def on_foto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    foto = update.message.photo[-1]
    archivo = await context.bot.get_file(foto.file_id)
    data = bytes(await archivo.download_as_bytearray())
    await _recibir_archivo(update, context, data, "image/jpeg", f"{foto.file_unique_id}.jpg")


async def on_documento(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    doc = update.message.document
    mime = doc.mime_type or ""
    if mime not in ("application/pdf", "image/jpeg", "image/png", "image/webp"):
        await update.message.reply_text("Mándame una foto o un PDF.")
        return
    archivo = await context.bot.get_file(doc.file_id)
    data = bytes(await archivo.download_as_bytearray())
    await _recibir_archivo(update, context, data, mime, doc.file_name or "comprobante")


async def on_texto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if await _rechazar(update):
        return
    original = (update.message.text or "").strip()
    borrador = _borrador(context)
    limpio = original.replace(".", "").replace("$", "").replace(" ", "")
    if borrador and limpio.isdigit():
        borrador["neto"] = int(limpio)
        await _guardar_borrador(context, borrador)
        vivo = _borrador(context)
        await update.message.reply_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))
        return
    espera = await update.message.reply_text("Estoy mirando la planilla…")
    try:
        respuesta = await responder(original)
    except (VisionError, PlanillaError) as exc:
        await espera.edit_text(str(exc))
        return
    if borrador:
        respuesta += "\n\nSigo con el comprobante pendiente. Si el neto está mal, escribe solo el número."
    await espera.edit_text(respuesta)


async def on_boton(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    if await _rechazar(update):
        return
    borrador = _borrador(context)
    dato = query.data or ""
    if dato == "no":
        context.user_data.pop("borrador", None)
        await query.edit_message_text("Descartado. Cuando quieras, manda otro comprobante.")
        return
    if not borrador:
        await query.edit_message_text("Ese comprobante ya no está pendiente. Manda la foto de nuevo.")
        return
    if dato == "tipo":
        await query.edit_message_text("¿Qué es este documento?", reply_markup=_teclado_tipos())
        return
    if dato.startswith("t:"):
        idx = int(dato.split(":")[1])
        borrador["tipo"] = TIPOS[idx]
        opciones = categorias_de(borrador["tipo"])
        borrador["categoria"] = opciones[0] if opciones else borrador["categoria"]
        borrador["afecto_iva"] = afecto_sugerido(borrador["categoria"], None)
        await _guardar_borrador(context, borrador)
        vivo = _borrador(context)
        await query.edit_message_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))
        return
    if dato.startswith("cat:"):
        pagina = int(dato.split(":")[1])
        await query.edit_message_text(
            f"Categoría para {borrador['tipo']}:",
            reply_markup=_teclado_categorias(borrador["tipo"], pagina),
        )
        return
    if dato.startswith("c:"):
        idx = int(dato.split(":")[1])
        opciones = categorias_de(borrador["tipo"])
        if 0 <= idx < len(opciones):
            borrador["categoria"] = opciones[idx]
            borrador["afecto_iva"] = afecto_sugerido(borrador["categoria"], None)
        await _guardar_borrador(context, borrador)
        vivo = _borrador(context)
        await query.edit_message_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))
        return
    if dato == "iva":
        borrador["afecto_iva"] = "No" if borrador.get("afecto_iva") == "Sí" else "Sí"
        await _guardar_borrador(context, borrador)
        vivo = _borrador(context)
        await query.edit_message_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))
        return
    if dato == "back":
        vivo = _borrador(context)
        await query.edit_message_text(_texto_borrador(vivo), reply_markup=_teclado(vivo))
        return
    if dato == "ok":
        if borrador.get("neto") is None:
            await query.edit_message_text("Falta el neto. Escríbelo como número, por ejemplo 45000.")
            return
        mov = Movimiento(
            fecha=borrador["fecha"],
            tipo=borrador["tipo"],
            categoria=borrador["categoria"],
            contraparte=borrador.get("contraparte") or "",
            documento=borrador.get("documento") or "",
            descripcion=borrador.get("descripcion") or "",
            neto=int(borrador["neto"]),
            afecto_iva=borrador.get("afecto_iva") or "Sí",
            origen="Telegram",
            archivo=borrador.get("archivo") or "",
            huella=borrador.get("huella") or "",
            marca=borrador.get("marca") or "",
            modelo=borrador.get("modelo") or "",
            anio_auto=borrador.get("anio_auto"),
            patente=borrador.get("patente") or "",
            precio_venta=borrador.get("precio_venta"),
            costo_neto=borrador.get("costo_neto"),
            notas=borrador.get("notas") or "",
        )
        try:
            donde = (
                registrar_en_drive(mov)
                if drive_listo()
                else registrar(mov, context.bot_data["planilla"])
            )
        except PlanillaAbierta as exc:
            await query.edit_message_text(str(exc))
            return
        except PlanillaError as exc:
            await query.edit_message_text(str(exc))
            return
        context.user_data.pop("borrador", None)
        await query.edit_message_text(
            f"Quedó cargado.\n{donde}\n"
            f"{mov.tipo} · {mov.categoria} · {_clp(mov.neto)} neto · { _fecha_txt(mov.fecha) }"
        )


def main() -> None:
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        raise SystemExit(
            "Falta TELEGRAM_BOT_TOKEN. Crea el bot con @BotFather y pon el token en .env"
        )
    planilla = ROOT / os.environ.get("PLANILLA", "Control_Financiero_Automotora.xlsx")
    estado = preparar_planilla(planilla)
    if drive_listo():
        print("Escribe en Google Drive")
    else:
        print(f"Escribe en el Excel local: {planilla.name} ({estado})")
    app = Application.builder().token(token).build()
    app.bot_data["planilla"] = planilla
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("ayuda", cmd_ayuda))
    app.add_handler(CommandHandler("cancelar", cmd_cancelar))
    app.add_handler(MessageHandler(filters.PHOTO, on_foto))
    app.add_handler(MessageHandler(filters.Document.ALL, on_documento))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_texto))
    app.add_handler(CallbackQueryHandler(on_boton))
    print("Bot escuchando. En Telegram abre el chat y manda /start", flush=True)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
