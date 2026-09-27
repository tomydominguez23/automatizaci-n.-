# Control financiero — Automotora

Planilla de Excel (también se abre en Google Sheets) para ver **cuánto gana la automotora**, con ingresos, gastos fijos, gastos variables, **IVA** y **PPM**.

Archivo principal: [Control_Financiero_Automotora.xlsx](Control_Financiero_Automotora.xlsx)

## Cómo empezarla

1. Abre el `.xlsx` en Excel o súbelo a Google Drive → *Abrir con Google Sheets* → *Archivo → Guardar como Google Sheets*.
2. En **Configuracion** escribe el nombre, el RUT, el año y el mes que estás cerrando.
3. Confirma la **tasa PPM** con tu contador (viene en 0,25%, típica del primer año comercial).
4. Borra las filas marcadas **EJEMPLO** e ingresa tus datos reales.

Las celdas **amarillas** se escriben. Las **gris azuladas** son fórmulas: no las borres.

## Qué hay en cada hoja

| Hoja | Para qué sirve |
| --- | --- |
| Comprobantes | Cada boleta o factura que entra por Telegram |
| Instrucciones | Cómo usar IVA, PPM y consignación |
| Configuracion | Nombre, mes, IVA 19%, PPM, impuesto de 1ª categoría |
| **Resumen** | Foto del mes: utilidad, F29 (IVA + PPM) y gráfico del año |
| Estado_Resultados | Resultado mes a mes, en neto (sin IVA) |
| Ventas_Autos | Cada unidad: nuevo, usado o consignación |
| Taller_Repuestos | OT, mano de obra y mostrador |
| Otros_Ingresos | Financiamiento, transferencias, etc. |
| Gastos_Fijos | Arriendo, sueldos, seguros… una columna por mes |
| Gastos_Variables | Publicidad, flete, detailing… (sin duplicar comisiones de vendedores) |
| IVA | Débito, crédito, remanente y IVA a pagar |
| PPM | Anticipo mensual y cruce con el impuesto a la renta |

## Fórmulas que usa (Chile)

- **IVA débito** = IVA de ventas y servicios del mes (autos vendidos, taller, comisiones).
- **IVA crédito** = IVA de compras afectas del mes (facturas). Un usado comprado a un particular casi nunca tiene crédito: marca *Compra afecta IVA = No*.
- **IVA a pagar** = `MAX(0; débito − crédito − remanente anterior)`. Si da negativo, queda remanente para el mes siguiente.
- **PPM** = ingresos netos del mes × tasa PPM. Se calcula sobre **ingresos**, no sobre la utilidad. Puede corresponder aunque el mes esté a pérdida.
- **Utilidad operacional** = ingresos − costos de lo vendido − gastos fijos − gastos variables − comisiones.
- **Utilidad neta estimada** = utilidad operacional − provisión de 1ª categoría (27% por defecto, solo si hay utilidad).

El IVA y el PPM **no son un gasto** de la automotora: son impuestos que se pagan en el **F29** (en general el día 12 del mes siguiente).

## Consignación

El auto no es de la automotora. El ingreso es la **comisión con IVA**, no el precio de venta al cliente. El precio de venta queda como referencia.

## Cargar gastos e ingresos por Telegram

Mandas una foto o un PDF de la boleta, la factura o el comprobante. El bot te muestra lo que leyó (fecha, proveedor, neto, IVA y categoría). Cuando confirmas, queda en **Comprobantes**. Los gastos fijos y variables de ese mes se suman solos. Una venta, el taller u otro ingreso además se anota en su hoja.

1. En Telegram habla con [@BotFather](https://t.me/BotFather), crea un bot y copia el token.
2. Entra a [Google AI Studio](https://aistudio.google.com/apikey) y crea una clave de Gemini (tiene cupo gratis).
3. Copia `.env.example` a `.env` y pega el token y la clave.
4. Instala y arranca:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 -m bot
```

5. Abre el bot en Telegram, manda `/start` y después la foto. La primera persona que escribe queda como dueña: hazlo tú.
6. Cierra el Excel antes de confirmar. Si el archivo está abierto, el bot no puede guardarlo.

La foto queda en `data/comprobantes/`. No subas el `.env` ni esa carpeta a ningún lado.

## Regenerar el archivo

Si quieres volver a la versión original (con ejemplos de agosto 2026):

```bash
python3 scripts/crear_planilla.py
```

Esta planilla es un **control de gestión**. No reemplaza el libro de compras/ventas ni el F29 que declara tu contador.
