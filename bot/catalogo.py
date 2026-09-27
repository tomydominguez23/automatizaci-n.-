"""Listas que comparten la planilla y el bot de Telegram."""

from __future__ import annotations

GASTOS_FIJOS = [
    "Arriendo local / patio",
    "Gastos comunes",
    "Sueldos (líquidos)",
    "Leyes sociales empleador",
    "Honorarios contador",
    "Seguros (local, stock, RC)",
    "Electricidad",
    "Agua",
    "Internet y telefonía",
    "Alarma / seguridad",
    "Aseo",
    "Patente municipal",
    "Software / facturación electrónica",
    "Mantención de instalaciones",
    "Publicidad fija (portales, redes)",
    "Contribuciones",
    "Caja chica",
    "Otros gastos fijos",
]

GASTOS_VARIABLES = [
    "Publicidad de campañas",
    "Preparación y detailing",
    "Traslado / flete de vehículos",
    "Combustible",
    "Garantía y postventa",
    "Transferencia / notaría / TAG",
    "Intereses y gastos bancarios",
    "Comisiones Transbank / Webpay",
    "Insumos de taller no cobrados",
    "Viáticos",
    "Multas e imprevistos",
    "Permiso circulación / empadronamiento",
    "Mantención de unidades en stock",
    "Comisiones marketplace / portales extra",
    "Otros gastos variables",
]

OTROS_INGRESOS = [
    "Gestión de financiamiento",
    "Transferencia",
    "Arriendo espacio",
    "Otros",
]

TIPOS = [
    "Gasto fijo",
    "Gasto variable",
    "Otro ingreso",
    "Taller",
    "Repuestos",
    "Venta auto",
]

AFECTO_POR_DEFECTO = {
    "Arriendo local / patio": "No",
    "Gastos comunes": "Sí",
    "Sueldos (líquidos)": "No",
    "Leyes sociales empleador": "No",
    "Honorarios contador": "No",
    "Seguros (local, stock, RC)": "No",
    "Electricidad": "Sí",
    "Agua": "No",
    "Internet y telefonía": "Sí",
    "Alarma / seguridad": "Sí",
    "Aseo": "Sí",
    "Patente municipal": "No",
    "Software / facturación electrónica": "Sí",
    "Mantención de instalaciones": "Sí",
    "Publicidad fija (portales, redes)": "Sí",
    "Contribuciones": "No",
    "Caja chica": "Sí",
    "Otros gastos fijos": "Sí",
    "Publicidad de campañas": "Sí",
    "Preparación y detailing": "Sí",
    "Traslado / flete de vehículos": "Sí",
    "Combustible": "Sí",
    "Garantía y postventa": "Sí",
    "Transferencia / notaría / TAG": "No",
    "Intereses y gastos bancarios": "No",
    "Comisiones Transbank / Webpay": "Sí",
    "Insumos de taller no cobrados": "Sí",
    "Viáticos": "No",
    "Multas e imprevistos": "No",
    "Permiso circulación / empadronamiento": "No",
    "Mantención de unidades en stock": "Sí",
    "Comisiones marketplace / portales extra": "Sí",
    "Otros gastos variables": "Sí",
    "Gestión de financiamiento": "Sí",
    "Transferencia": "Sí",
    "Arriendo espacio": "Sí",
    "Otros": "Sí",
}


ALIASES = {
    "luz": "Electricidad",
    "enel": "Electricidad",
    "cge": "Electricidad",
    "agua": "Agua",
    "bencina": "Combustible",
    "combustible": "Combustible",
    "copec": "Combustible",
    "shell": "Combustible",
    "arriendo": "Arriendo local / patio",
    "sueldo": "Sueldos (líquidos)",
    "sueldos": "Sueldos (líquidos)",
    "contador": "Honorarios contador",
    "honorarios": "Honorarios contador",
    "internet": "Internet y telefonía",
    "telefono": "Internet y telefonía",
    "telefonia": "Internet y telefonía",
    "publicidad": "Publicidad de campañas",
    "flete": "Traslado / flete de vehículos",
    "grua": "Traslado / flete de vehículos",
    "notaria": "Transferencia / notaría / TAG",
    "notaría": "Transferencia / notaría / TAG",
    "tag": "Transferencia / notaría / TAG",
    "detailing": "Preparación y detailing",
    "aseo": "Aseo",
    "alarma": "Alarma / seguridad",
    "seguro": "Seguros (local, stock, RC)",
    "seguros": "Seguros (local, stock, RC)",
    "patente municipal": "Patente municipal",
    "contribuciones": "Contribuciones",
}


def categorias_de(tipo: str) -> list[str]:
    if tipo == "Gasto fijo":
        return GASTOS_FIJOS
    if tipo == "Gasto variable":
        return GASTOS_VARIABLES
    if tipo == "Otro ingreso":
        return OTROS_INGRESOS
    if tipo in ("Taller", "Repuestos"):
        return ["Taller", "Repuestos"]
    if tipo == "Venta auto":
        return ["Nuevo", "Usado", "Consignación"]
    return []


def categoria_respaldo(tipo: str) -> str:
    if tipo == "Gasto fijo":
        return "Otros gastos fijos"
    if tipo == "Gasto variable":
        return "Otros gastos variables"
    if tipo == "Otro ingreso":
        return "Otros"
    if tipo == "Repuestos":
        return "Repuestos"
    if tipo == "Venta auto":
        return "Usado"
    return "Taller"
