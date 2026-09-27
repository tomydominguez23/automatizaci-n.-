"""Escribe los comprobantes en la planilla de Google Drive."""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path

from bot.catalogo import categorias_de
from bot.formulas import COMP, COMP_FIRST, COMP_LAST
from bot.planilla import Movimiento, PlanillaError

ROOT = Path(__file__).resolve().parents[1]
SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
)
TOKEN = ROOT / "credenciales" / "token.json"
CLIENTE = ROOT / "credenciales" / "cliente.json"


def ruta_token() -> Path:
    explicita = os.environ.get("GOOGLE_TOKEN_PATH", "").strip()
    if explicita:
        return Path(explicita)
    persistente = Path("/data")
    if persistente.is_dir():
        return persistente / "token.json"
    return TOKEN


def asegurar_token() -> Path:
    ruta = ruta_token()
    crudo = os.environ.get("GOOGLE_TOKEN_JSON", "").strip()
    if crudo and not ruta.exists():
        ruta.parent.mkdir(parents=True, exist_ok=True)
        ruta.write_text(crudo, encoding="utf-8")
    return ruta


def configurada() -> bool:
    return bool(os.environ.get("GOOGLE_SHEETS_ID", "").strip()) and asegurar_token().exists()


def _serial(dia: date) -> int:
    return (dia - date(1899, 12, 30)).days


def _cliente():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    import gspread

    ruta = asegurar_token()
    cred = Credentials.from_authorized_user_file(str(ruta), list(SCOPES))
    if cred.expired and cred.refresh_token:
        cred.refresh(Request())
        ruta.write_text(cred.to_json(), encoding="utf-8")
    return gspread.authorize(cred)


def _libro():
    hoja_id = os.environ.get("GOOGLE_SHEETS_ID", "").strip()
    if not hoja_id:
        raise PlanillaError("Falta GOOGLE_SHEETS_ID en .env.")
    if not asegurar_token().exists():
        raise PlanillaError("Falta autorizar Google. Corre: python3 -m bot.vincular")
    return _cliente().open_by_key(hoja_id)


def _siguiente(ws, columna: int, inicio: int, fin: int) -> int:
    valores = ws.col_values(columna)
    for fila in range(inicio, fin + 1):
        if fila - 1 >= len(valores) or str(valores[fila - 1]).strip() == "":
            return fila
    raise PlanillaError("No quedan filas libres en la planilla de Drive.")


def _poner(ws, fila: int, columna: int, valor) -> None:
    if isinstance(valor, datetime):
        valor = _serial(valor.date())
    elif isinstance(valor, date):
        valor = _serial(valor)
    ws.update_cell(fila, columna, valor)


def huella_ya_cargada(huella: str) -> bool:
    if not huella or not configurada():
        return False
    ws = _libro().worksheet(COMP)
    for valor in ws.col_values(17)[COMP_FIRST - 1 : COMP_LAST]:
        if valor == huella:
            return True
    return False


def registrar(mov: Movimiento) -> str:
    libro = _libro()
    ws = libro.worksheet(COMP)
    fila = _siguiente(ws, 2, COMP_FIRST, COMP_LAST)
    pares = {
        1: fila - COMP_FIRST + 1,
        2: mov.fecha,
        3: mov.tipo,
        4: mov.categoria,
        5: mov.contraparte,
        6: mov.documento,
        7: mov.descripcion,
        8: int(mov.neto),
        9: mov.afecto_iva if mov.afecto_iva in ("Sí", "No") else "Sí",
        14: "Telegram",
        15: "Confirmado",
        16: mov.archivo,
        17: mov.huella,
    }
    for col, valor in pares.items():
        _poner(ws, fila, col, valor)
    detalle = f"Drive · Comprobantes fila {fila}"

    if mov.tipo == "Otro ingreso":
        hoja = libro.worksheet("Otros_Ingresos")
        r = _siguiente(hoja, 1, 6, 105)
        cat = mov.categoria if mov.categoria in categorias_de("Otro ingreso") else "Otros"
        for col, valor in (
            (1, mov.fecha),
            (2, cat),
            (3, mov.descripcion or mov.contraparte),
            (4, int(mov.neto)),
            (5, mov.afecto_iva if mov.afecto_iva in ("Sí", "No") else "Sí"),
        ):
            _poner(hoja, r, col, valor)
        detalle += f" · Otros ingresos fila {r}"

    elif mov.tipo in ("Taller", "Repuestos"):
        hoja = libro.worksheet("Taller_Repuestos")
        r = _siguiente(hoja, 1, 6, 205)
        datos = {
            1: mov.fecha,
            2: "Repuestos" if mov.tipo == "Repuestos" else "Taller",
            3: mov.documento,
            4: mov.contraparte,
            5: mov.descripcion,
            6: int(mov.neto),
            10: "Sí" if mov.afecto_iva == "Sí" else "No",
        }
        if mov.costo_neto is not None:
            datos[9] = int(mov.costo_neto)
        for col, valor in datos.items():
            _poner(hoja, r, col, valor)
        detalle += f" · Taller fila {r}"

    elif mov.tipo == "Venta auto":
        hoja = libro.worksheet("Ventas_Autos")
        r = _siguiente(hoja, 4, 6, 205)
        tipo_auto = mov.categoria if mov.categoria in ("Nuevo", "Usado", "Consignación") else "Usado"
        datos = {
            4: tipo_auto,
            5: mov.marca,
            6: mov.modelo,
            8: mov.patente,
            9: mov.contraparte,
            10: "Sí" if tipo_auto == "Nuevo" else "No",
            17: "Vendido",
            18: mov.descripcion or mov.documento,
        }
        if mov.anio_auto:
            datos[7] = int(mov.anio_auto)
        if mov.costo_neto:
            datos[2] = mov.fecha
            datos[11] = int(mov.costo_neto)
        precio = mov.precio_venta if mov.precio_venta is not None else int(mov.neto)
        if tipo_auto == "Consignación":
            datos[15] = int(precio)
        else:
            datos[3] = mov.fecha
            datos[14] = int(precio)
        if mov.comision:
            datos[16] = int(mov.comision)
        for col, valor in datos.items():
            _poner(hoja, r, col, valor)
        detalle += f" · Ventas fila {r}"

    return detalle
