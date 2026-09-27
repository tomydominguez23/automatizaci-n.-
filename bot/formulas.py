"""Fórmulas que conectan Comprobantes con gastos e IVA."""

from __future__ import annotations

ANIO = "Configuracion!$B$7"
MES = "Configuracion!$B$8"
IVA = "Configuracion!$B$11"
COMP = "Comprobantes"
COMP_FIRST = 6
COMP_LAST = 2005


def _rango(col: str) -> str:
    return f"{COMP}!${col}${COMP_FIRST}:${col}${COMP_LAST}"


def suma_neto(tipo: str, fila_concepto: int, mes: int | str) -> str:
    mes_ref = MES if mes == "actual" else str(mes)
    return (
        f'=SUMIFS({_rango("H")},'
        f'{_rango("C")},"{tipo}",'
        f'{_rango("D")},A{fila_concepto},'
        f'{_rango("L")},{ANIO},'
        f'{_rango("M")},{mes_ref},'
        f'{_rango("O")},"Confirmado")'
    )


def suma_iva_concepto(tipo: str, fila_concepto: int, mes: int | str | None = None) -> str:
    partes = [
        f'=SUMIFS({_rango("J")},',
        f'{_rango("C")},"{tipo}",',
        f'{_rango("D")},A{fila_concepto},',
        f'{_rango("L")},{ANIO},',
    ]
    if mes is not None:
        mes_ref = MES if mes == "actual" else str(mes)
        partes.append(f'{_rango("M")},{mes_ref},')
    partes.append(f'{_rango("O")},"Confirmado")')
    return "".join(partes)


def suma_iva_tipo(tipo: str, mes: int) -> str:
    return (
        f'=SUMIFS({_rango("J")},'
        f'{_rango("C")},"{tipo}",'
        f'{_rango("L")},{ANIO},'
        f'{_rango("M")},{mes},'
        f'{_rango("O")},"Confirmado")'
    )
