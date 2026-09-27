"""Lee y escribe Control_Financiero_Automotora.xlsx."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from bot.catalogo import AFECTO_POR_DEFECTO, categorias_de
from bot.formulas import (
    COMP,
    COMP_FIRST,
    COMP_LAST,
    suma_iva_concepto,
    suma_iva_tipo,
    suma_neto,
)

ROOT = Path(__file__).resolve().parents[1]
PLANILLA = ROOT / "Control_Financiero_Automotora.xlsx"

GF = "Gastos_Fijos"
GV = "Gastos_Variables"
OI = "Otros_Ingresos"
TR = "Taller_Repuestos"
VA = "Ventas_Autos"
SH_IVA = "IVA"

NAVY = "1A365D"
INPUT_BG = "FFF8DC"
FORMULA_BG = "EDF2F7"
LINE = "CBD5E0"
WHITE = "FFFFFF"
TEAL = "234E52"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
font_h = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_input = Font(name="Calibri", size=10, color="744210")
font_formula = Font(name="Calibri", size=10, color="2A4365")
font_muted = Font(name="Calibri", size=9, italic=True, color="718096")
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_input = PatternFill("solid", fgColor=INPUT_BG)
fill_formula = PatternFill("solid", fgColor=FORMULA_BG)
fill_teal = PatternFill("solid", fgColor=TEAL)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
FMT_CLP = '"$"#,##0'


class PlanillaAbierta(RuntimeError):
    """El Excel está abierto y Windows/Excel no deja guardarlo."""


class PlanillaError(RuntimeError):
    pass


@dataclass
class Movimiento:
    fecha: date
    tipo: str
    categoria: str
    contraparte: str = ""
    documento: str = ""
    descripcion: str = ""
    neto: int = 0
    afecto_iva: str = "Sí"
    origen: str = "Telegram"
    estado: str = "Confirmado"
    archivo: str = ""
    huella: str = ""
    # Venta de auto
    marca: str = ""
    modelo: str = ""
    anio_auto: int | None = None
    patente: str = ""
    precio_venta: int | None = None
    costo_neto: int | None = None
    comision: int | None = None
    notas: str = ""
    extras: dict = field(default_factory=dict)


def _anio_config(wb) -> int:
    valor = wb["Configuracion"]["B7"].value
    if isinstance(valor, datetime):
        return valor.year
    try:
        return int(valor)
    except (TypeError, ValueError):
        return date.today().year


def preparar_hoja_comprobantes(ws) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.merge_cells("A1:Q1")
    c = ws.cell(1, 1, "  Comprobantes — lo que entra por Telegram")
    c.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    c.fill = fill_navy
    c.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(1, 18):
        ws.cell(1, col).fill = fill_navy
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:Q2")
    ws.cell(
        2,
        1,
        "  Cada foto confirmada es una fila. Los gastos fijos y variables suman desde aquí. No borres las columnas grises.",
    ).font = Font(name="Calibri", size=10, color=WHITE)
    for col in range(1, 18):
        ws.cell(2, col).fill = PatternFill("solid", fgColor="2C5282")
    ws.row_dimensions[2].height = 22

    ws.merge_cells("A3:Q3")
    ws.cell(
        3,
        1,
        "Amarillo = datos del bot. Gris = fórmulas (IVA, total, año y mes).",
    ).font = font_muted

    headers = [
        "N°",
        "Fecha",
        "Tipo",
        "Categoría",
        "Proveedor / cliente",
        "N° documento",
        "Descripción",
        "Neto",
        "Afecto IVA",
        "IVA",
        "Total",
        "Año",
        "Mes",
        "Origen",
        "Estado",
        "Archivo",
        "Huella",
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(5, i, h)
        cell.font = font_h
        cell.fill = fill_navy if i <= 11 or i >= 14 else fill_teal
        cell.alignment = center
        cell.border = thin
    ws.row_dimensions[5].height = 32

    iva = "Configuracion!$B$11"
    for r in range(COMP_FIRST, COMP_LAST + 1):
        for c in range(1, 18):
            ws.cell(r, c).border = thin
            ws.cell(r, c).alignment = center
        for c in (1, 2, 3, 4, 5, 6, 7, 8, 9, 14, 15, 16, 17):
            ws.cell(r, c).fill = fill_input
            ws.cell(r, c).font = font_input
        ws.cell(r, 2).number_format = "DD/MM/YYYY"
        ws.cell(r, 8).number_format = FMT_CLP
        for col, formula in (
            (10, f'=IF(H{r}="","",IF(I{r}="Sí",ROUND(H{r}*{iva},0),0))'),
            (11, f'=IF(H{r}="","",H{r}+J{r})'),
            (12, f'=IF(B{r}="","",YEAR(B{r}))'),
            (13, f'=IF(B{r}="","",MONTH(B{r}))'),
        ):
            cell = ws.cell(r, col, formula)
            cell.fill = fill_formula
            cell.font = font_formula
            cell.border = thin
            cell.alignment = center
            if col in (10, 11):
                cell.number_format = FMT_CLP

    widths = {
        "A": 8, "B": 13, "C": 16, "D": 32, "E": 28, "F": 16, "G": 36,
        "H": 14, "I": 12, "J": 12, "K": 14, "L": 8, "M": 8, "N": 14,
        "O": 14, "P": 28, "Q": 18,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    def lista(formula, celdas):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(celdas)

    lista(
        '"Gasto fijo,Gasto variable,Otro ingreso,Taller,Repuestos,Venta auto"',
        f"C{COMP_FIRST}:C{COMP_LAST}",
    )
    lista('"Sí,No"', f"I{COMP_FIRST}:I{COMP_LAST}")
    lista('"Confirmado,Anulado"', f"O{COMP_FIRST}:O{COMP_LAST}")
    ws.auto_filter.ref = f"A5:Q{COMP_LAST}"
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = "C9A227"
    ws.auto_filter.ref = f"A5:Q{COMP_LAST}"


def _enlazar_gastos(ws, tipo: str) -> None:
    """Convierte montos escritos a filas y deja los meses como suma de comprobantes."""
    # Se llama con el workbook completo; aquí solo reescribe fórmulas de una hoja.
    ultima = 23 if tipo == "Gasto fijo" else 20
    for r in range(6, ultima + 1):
        ws.cell(r, 16).value = suma_iva_concepto(tipo, r)
        ws.cell(r, 16).fill = fill_formula
        ws.cell(r, 16).font = font_formula
        ws.cell(r, 16).number_format = FMT_CLP
        ws.cell(r, 18).value = suma_iva_concepto(tipo, r, "actual")
        ws.cell(r, 18).fill = fill_formula
        ws.cell(r, 18).font = font_formula
        ws.cell(r, 18).number_format = FMT_CLP
        for m in range(1, 13):
            cell = ws.cell(r, 2 + m)
            cell.value = suma_neto(tipo, r, m)
            cell.fill = fill_formula
            cell.font = font_formula
            cell.number_format = FMT_CLP
            cell.border = thin


def _volcar_montos_existentes(wb) -> int:
    """Pasa números ya tipeados en gastos a filas de Comprobantes, una vez."""
    anio = _anio_config(wb)
    ws_c = wb[COMP]
    creados = 0
    for hoja, tipo, ultima in ((GF, "Gasto fijo", 23), (GV, "Gasto variable", 20)):
        ws = wb[hoja]
        for r in range(6, ultima + 1):
            concepto = ws.cell(r, 1).value
            if not concepto or not isinstance(concepto, str):
                continue
            afecto = ws.cell(r, 2).value if ws.cell(r, 2).value in ("Sí", "No") else "Sí"
            for m in range(1, 13):
                cell = ws.cell(r, 2 + m)
                valor = cell.value
                if isinstance(valor, str) and valor.startswith("="):
                    continue
                if not isinstance(valor, (int, float)) or valor == 0:
                    continue
                mov = Movimiento(
                    fecha=date(anio, m, 15),
                    tipo=tipo,
                    categoria=str(concepto),
                    descripcion="Saldo que ya estaba anotado en la planilla",
                    neto=int(round(valor)),
                    afecto_iva=afecto,
                    origen="Planilla",
                    estado="Confirmado",
                )
                _escribir_comprobante(ws_c, mov)
                creados += 1
    return creados


def _enlazar_iva(wb) -> None:
    ws = wb[SH_IVA]
    for m in range(1, 13):
        col = get_column_letter(m + 1)
        ws.cell(15, m + 1).value = suma_iva_tipo("Gasto fijo", m)
        ws.cell(16, m + 1).value = suma_iva_tipo("Gasto variable", m)
        ws.cell(15, m + 1).number_format = FMT_CLP
        ws.cell(16, m + 1).number_format = FMT_CLP
        # col used to keep the month alignment obvious for readers of the sheet
        _ = col


def preparar_planilla(path: Path = PLANILLA) -> str:
    if not path.exists():
        raise PlanillaError(
            f"No encuentro {path.name}. Generala con: python3 scripts/crear_planilla.py"
        )
    wb = load_workbook(path)
    ya = COMP in wb.sheetnames and isinstance(wb[GF]["C6"].value, str) and str(wb[GF]["C6"].value).startswith("=SUMIFS")
    if ya:
        wb.close()
        return "lista"
    if COMP not in wb.sheetnames:
        ws = wb.create_sheet(COMP, 0)
        preparar_hoja_comprobantes(ws)
        _volcar_montos_existentes(wb)
    _enlazar_gastos(wb[GF], "Gasto fijo")
    _enlazar_gastos(wb[GV], "Gasto variable")
    _enlazar_iva(wb)
    nota = wb[GF].cell(3, 1)
    nota.value = (
        "Los montos de cada mes salen de la hoja Comprobantes (Telegram). "
        "Afecto IVA de esta columna es la referencia del concepto."
    )
    wb[GF].cell(3, 1).comment = Comment(
        "No escribas números en los meses: se pisan con la fórmula.",
        "Telegram",
    )
    _guardar(wb, path)
    return "actualizada"


def _guardar(wb, path: Path) -> None:
    try:
        wb.save(path)
    except PermissionError as exc:
        raise PlanillaAbierta(
            "La planilla está abierta. Ciérrala en Excel y confirma de nuevo."
        ) from exc
    finally:
        wb.close()


def _siguiente_fila(ws, columna: int, inicio: int, fin: int) -> int:
    for r in range(inicio, fin + 1):
        if ws.cell(r, columna).value in (None, ""):
            return r
    raise PlanillaError("No quedan filas libres en la planilla.")


def _escribir_comprobante(ws, mov: Movimiento) -> int:
    fila = _siguiente_fila(ws, 2, COMP_FIRST, COMP_LAST)
    n = fila - COMP_FIRST + 1
    ws.cell(fila, 1).value = n
    ws.cell(fila, 2).value = mov.fecha
    ws.cell(fila, 3).value = mov.tipo
    ws.cell(fila, 4).value = mov.categoria
    ws.cell(fila, 5).value = mov.contraparte
    ws.cell(fila, 6).value = mov.documento
    ws.cell(fila, 7).value = mov.descripcion
    ws.cell(fila, 8).value = int(mov.neto)
    ws.cell(fila, 9).value = mov.afecto_iva if mov.afecto_iva in ("Sí", "No") else "Sí"
    ws.cell(fila, 14).value = mov.origen
    ws.cell(fila, 15).value = mov.estado
    ws.cell(fila, 16).value = mov.archivo
    ws.cell(fila, 17).value = mov.huella
    return fila


def huella_ya_cargada(path: Path, huella: str) -> bool:
    if not huella or not path.exists():
        return False
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if COMP not in wb.sheetnames:
            return False
        ws = wb[COMP]
        for row in ws.iter_rows(min_row=COMP_FIRST, max_row=COMP_LAST, min_col=17, max_col=17):
            if row[0].value == huella:
                return True
        return False
    finally:
        wb.close()


def registrar(mov: Movimiento, path: Path = PLANILLA) -> str:
    preparar_planilla(path)
    wb = load_workbook(path)
    ws = wb[COMP]
    fila_c = _escribir_comprobante(ws, mov)
    detalle = f"Comprobantes fila {fila_c}"

    if mov.tipo == "Otro ingreso":
        hoja = wb[OI]
        r = _siguiente_fila(hoja, 1, 6, 105)
        hoja.cell(r, 1).value = mov.fecha
        cat = mov.categoria if mov.categoria in categorias_de("Otro ingreso") else "Otros"
        hoja.cell(r, 2).value = cat
        hoja.cell(r, 3).value = mov.descripcion or mov.contraparte
        hoja.cell(r, 4).value = int(mov.neto)
        hoja.cell(r, 5).value = mov.afecto_iva if mov.afecto_iva in ("Sí", "No") else "Sí"
        detalle += f" · Otros ingresos fila {r}"

    elif mov.tipo in ("Taller", "Repuestos"):
        hoja = wb[TR]
        r = _siguiente_fila(hoja, 1, 6, 205)
        hoja.cell(r, 1).value = mov.fecha
        hoja.cell(r, 2).value = "Repuestos" if mov.tipo == "Repuestos" else "Taller"
        hoja.cell(r, 3).value = mov.documento
        hoja.cell(r, 4).value = mov.contraparte
        hoja.cell(r, 5).value = mov.descripcion
        hoja.cell(r, 6).value = int(mov.neto)
        if mov.costo_neto is not None:
            hoja.cell(r, 9).value = int(mov.costo_neto)
        hoja.cell(r, 10).value = "Sí" if mov.afecto_iva == "Sí" else "No"
        detalle += f" · Taller fila {r}"

    elif mov.tipo == "Venta auto":
        hoja = wb[VA]
        r = _siguiente_fila(hoja, 4, 6, 205)
        tipo_auto = mov.categoria if mov.categoria in ("Nuevo", "Usado", "Consignación") else "Usado"
        hoja.cell(r, 4).value = tipo_auto
        hoja.cell(r, 5).value = mov.marca
        hoja.cell(r, 6).value = mov.modelo
        if mov.anio_auto:
            hoja.cell(r, 7).value = int(mov.anio_auto)
        hoja.cell(r, 8).value = mov.patente
        hoja.cell(r, 9).value = mov.contraparte
        afecto_compra = "Sí" if tipo_auto == "Nuevo" else "No"
        hoja.cell(r, 10).value = afecto_compra
        if mov.costo_neto:
            hoja.cell(r, 2).value = mov.fecha
            hoja.cell(r, 11).value = int(mov.costo_neto)
        precio = mov.precio_venta if mov.precio_venta is not None else int(mov.neto)
        if tipo_auto == "Consignación":
            hoja.cell(r, 15).value = int(precio)
        else:
            hoja.cell(r, 3).value = mov.fecha
            hoja.cell(r, 14).value = int(precio)
        if mov.comision:
            hoja.cell(r, 16).value = int(mov.comision)
        hoja.cell(r, 17).value = "Vendido"
        hoja.cell(r, 18).value = mov.descripcion or mov.documento
        detalle += f" · Ventas fila {r}"

    _guardar(wb, path)
    return detalle


def hash_archivo(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


def afecto_sugerido(categoria: str, leido: str | None) -> str:
    if leido in ("Sí", "No"):
        return leido
    return AFECTO_POR_DEFECTO.get(categoria, "Sí")
