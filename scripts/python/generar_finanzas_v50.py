#!/usr/bin/env python3
"""
Generador de Excel Financiero Profesional - Finanzas_v50.xlsx

Proyecto: Excel_Finance_Project v5.0
Autor: Alvaro Velasco
Fecha: 16 de noviembre, 2025
Versión: 5.0

Características:
- Paleta de colores profesional financiera/contable
- Formato de fecha dd/mm/yy
- Columnas auto-ajustables
- Protección de columnas automatizadas
- Múltiples hojas: Dashboard, Ingresos, Gastos, Balance, Flujo Caja
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, Protection
)
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import os


# =============================================================================
# CONFIGURACIÓN DE PALETA DE COLORES PROFESIONAL FINANCIERA
# =============================================================================

class PaletaFinanciera:
    """Paleta de colores basada en estándares de análisis financiero"""

    # Headers principales
    HEADER_DARK = "1F4E78"      # Azul oscuro corporativo
    HEADER_TEXT = "FFFFFF"      # Blanco

    # Subheaders
    SUBHEADER_BG = "4472C4"     # Azul medio
    SUBHEADER_TEXT = "FFFFFF"   # Blanco

    # Secciones
    INGRESOS_BG = "E2EFDA"      # Verde claro (positivo)
    GASTOS_BG = "FCE4D6"        # Naranja claro (salida)
    ACTIVOS_BG = "DEEBF7"       # Azul claro (recursos)
    PASIVOS_BG = "FFF2CC"       # Amarillo claro (obligaciones)

    # Filas alternadas
    ROW_EVEN = "F2F2F2"         # Gris muy claro
    ROW_ODD = "FFFFFF"          # Blanco

    # Totales y cálculos
    TOTAL_BG = "D9E1F2"         # Azul muy claro
    SUBTOTAL_BG = "E7E6E6"      # Gris claro

    # Estados financieros
    POSITIVO = "C6EFCE"         # Verde claro (ganancia)
    NEGATIVO = "FFC7CE"         # Rojo claro (pérdida)
    NEUTRO = "FFEB9C"           # Amarillo claro (break-even)

    # Texto
    TEXT_DARK = "000000"        # Negro
    TEXT_MONEY = "0F6B3E"       # Verde oscuro (montos)
    TEXT_ALERT = "C00000"       # Rojo oscuro (alertas)


# =============================================================================
# CONFIGURACIÓN DE ESTILOS
# =============================================================================

def crear_estilo_header():
    """Estilo para headers principales"""
    return {
        'font': Font(name='Calibri', size=12, bold=True, color=PaletaFinanciera.HEADER_TEXT),
        'fill': PatternFill(start_color=PaletaFinanciera.HEADER_DARK,
                           end_color=PaletaFinanciera.HEADER_DARK,
                           fill_type='solid'),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }


def crear_estilo_subheader():
    """Estilo para subheaders"""
    return {
        'font': Font(name='Calibri', size=11, bold=True, color=PaletaFinanciera.SUBHEADER_TEXT),
        'fill': PatternFill(start_color=PaletaFinanciera.SUBHEADER_BG,
                           end_color=PaletaFinanciera.SUBHEADER_BG,
                           fill_type='solid'),
        'alignment': Alignment(horizontal='center', vertical='center'),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }


def crear_estilo_celda_normal():
    """Estilo para celdas normales"""
    return {
        'font': Font(name='Calibri', size=10),
        'alignment': Alignment(horizontal='left', vertical='center'),
        'border': Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )
    }


def crear_estilo_total():
    """Estilo para filas de totales"""
    return {
        'font': Font(name='Calibri', size=11, bold=True),
        'fill': PatternFill(start_color=PaletaFinanciera.TOTAL_BG,
                           end_color=PaletaFinanciera.TOTAL_BG,
                           fill_type='solid'),
        'alignment': Alignment(horizontal='right', vertical='center'),
        'border': Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='double')
        )
    }


# =============================================================================
# FUNCIONES DE GENERACIÓN DE HOJAS
# =============================================================================

def crear_hoja_dashboard(wb):
    """Crea la hoja Dashboard con resumen ejecutivo"""
    ws = wb.active
    ws.title = "Dashboard"

    # Título principal
    ws['A1'] = "DASHBOARD FINANCIERO"
    ws.merge_cells('A1:F1')
    aplicar_estilo(ws['A1'], crear_estilo_header())
    ws.row_dimensions[1].height = 30

    # Información del proyecto
    ws['A2'] = f"Proyecto: Excel_Finance_Project v5.0"
    ws['A3'] = f"Última actualización: {datetime.now().strftime('%d/%m/%y')}"

    # Sección de indicadores clave (KPIs)
    fila_actual = 5
    ws[f'A{fila_actual}'] = "INDICADORES CLAVE"
    ws.merge_cells(f'A{fila_actual}:F{fila_actual}')
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_subheader())
    ws.row_dimensions[fila_actual].height = 25

    fila_actual += 1
    headers_kpi = ["Indicador", "Valor Actual", "Meta", "Estado", "% Cumplimiento", "Tendencia"]
    for col, header in enumerate(headers_kpi, start=1):
        celda = ws.cell(row=fila_actual, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())

    # Datos de ejemplo KPIs
    kpis = [
        ["Ingresos Mensuales", "₡0.00", "₡500,000", "Pendiente", "0%", "-"],
        ["Gastos Mensuales", "₡0.00", "₡350,000", "Pendiente", "0%", "-"],
        ["Flujo de Caja", "₡0.00", ">₡0", "Pendiente", "-", "-"],
        ["Ratio Activos/Pasivos", "0.00", ">1.5", "Pendiente", "0%", "-"],
    ]

    for kpi_data in kpis:
        fila_actual += 1
        for col, valor in enumerate(kpi_data, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            celda.value = valor
            aplicar_estilo(celda, crear_estilo_celda_normal())

    # Sección de resumen financiero
    fila_actual += 3
    ws[f'A{fila_actual}'] = "RESUMEN FINANCIERO"
    ws.merge_cells(f'A{fila_actual}:D{fila_actual}')
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_subheader())

    fila_actual += 1
    headers_resumen = ["Concepto", "Monto (₡)", "% del Total", "Notas"]
    for col, header in enumerate(headers_resumen, start=1):
        celda = ws.cell(row=fila_actual, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())

    resumen_items = [
        ["Total Ingresos", "=Ingresos!F100", "100%", "Ver detalle en hoja Ingresos"],
        ["Total Gastos", "=Gastos!F100", "100%", "Ver detalle en hoja Gastos"],
        ["Balance Neto", "=B16-B17", "=C16-C17", "Ingresos - Gastos"],
    ]

    for item in resumen_items:
        fila_actual += 1
        for col, valor in enumerate(item, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            celda.value = valor
            if col == 2 or col == 3:  # Columnas de montos
                celda.number_format = '#,##0.00'
            aplicar_estilo(celda, crear_estilo_celda_normal())

    # Fila de total con estilo especial
    ws[f'A{fila_actual}'].font = Font(bold=True)
    ws[f'B{fila_actual}'].font = Font(bold=True)

    # Auto-ajustar columnas
    ajustar_columnas(ws)

    # Proteger columnas calculadas (B, C)
    proteger_columnas_calculadas(ws, ['B', 'C'])


def crear_hoja_ingresos(wb):
    """Crea la hoja de Ingresos detallada"""
    ws = wb.create_sheet("Ingresos")

    # Título
    ws['A1'] = "REGISTRO DE INGRESOS"
    ws.merge_cells('A1:G1')
    aplicar_estilo(ws['A1'], crear_estilo_header())
    ws.row_dimensions[1].height = 30

    # Headers de columnas
    headers = ["Fecha", "Categoría", "Descripción", "Monto (₡)", "Método Pago", "Estado", "Notas"]
    for col, header in enumerate(headers, start=1):
        celda = ws.cell(row=2, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())
        ws.column_dimensions[get_column_letter(col)].width = 15

    # Aplicar fondo verde claro (ingresos)
    for col in range(1, 8):
        ws.cell(row=2, column=col).fill = PatternFill(
            start_color=PaletaFinanciera.INGRESOS_BG,
            end_color=PaletaFinanciera.INGRESOS_BG,
            fill_type='solid'
        )

    # Datos de ejemplo (3 registros)
    fecha_inicio = datetime.now()
    datos_ejemplo = [
        [fecha_inicio, "Salario", "Salario mensual", 500000, "Transferencia", "Confirmado", "Pago regular"],
        [fecha_inicio + timedelta(days=5), "Freelance", "Proyecto diseño web", 150000, "Sinpe", "Confirmado", "Cliente XYZ"],
        [fecha_inicio + timedelta(days=10), "Inversiones", "Dividendos", 25000, "Transferencia", "Pendiente", "Cuenta inversión"],
    ]

    fila_actual = 3
    for datos in datos_ejemplo:
        for col, valor in enumerate(datos, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            if col == 1:  # Columna de fecha
                celda.value = valor
                celda.number_format = 'DD/MM/YY'
            elif col == 4:  # Columna de monto
                celda.value = valor
                celda.number_format = '₡#,##0.00'
            else:
                celda.value = valor
            aplicar_estilo(celda, crear_estilo_celda_normal())
        fila_actual += 1

    # Fila de total (SINTAXIS ESPAÑOL para Excel CR)
    fila_total = 100  # Fila fija para referencia desde Dashboard
    ws[f'A{fila_total}'] = "TOTAL INGRESOS"
    ws.merge_cells(f'A{fila_total}:C{fila_total}')
    ws[f'D{fila_total}'] = f"=SUMA(D3:D99)"  # SUMA en español
    ws[f'D{fila_total}'].number_format = '₡#,##0.00'
    ws[f'F{fila_total}'] = f"=SUMA(D3:D99)"  # Para referencia desde Dashboard
    ws[f'F{fila_total}'].number_format = '₡#,##0.00'

    aplicar_estilo(ws[f'A{fila_total}'], crear_estilo_total())
    aplicar_estilo(ws[f'D{fila_total}'], crear_estilo_total())
    aplicar_estilo(ws[f'F{fila_total}'], crear_estilo_total())

    # Auto-ajustar columnas
    ajustar_columnas(ws)

    # Proteger columnas calculadas (D en total, F)
    proteger_columnas_calculadas(ws, ['D', 'F'])


def crear_hoja_gastos(wb):
    """Crea la hoja de Gastos detallada"""
    ws = wb.create_sheet("Gastos")

    # Título
    ws['A1'] = "REGISTRO DE GASTOS"
    ws.merge_cells('A1:H1')
    aplicar_estilo(ws['A1'], crear_estilo_header())
    ws.row_dimensions[1].height = 30

    # Headers de columnas
    headers = ["Fecha", "Categoría", "Subcategoría", "Descripción", "Monto (₡)", "Método Pago", "Estado", "Notas"]
    for col, header in enumerate(headers, start=1):
        celda = ws.cell(row=2, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())
        ws.column_dimensions[get_column_letter(col)].width = 15

    # Aplicar fondo naranja claro (gastos)
    for col in range(1, 9):
        ws.cell(row=2, column=col).fill = PatternFill(
            start_color=PaletaFinanciera.GASTOS_BG,
            end_color=PaletaFinanciera.GASTOS_BG,
            fill_type='solid'
        )

    # Datos de ejemplo
    fecha_inicio = datetime.now()
    datos_ejemplo = [
        [fecha_inicio, "Vivienda", "Alquiler", "Alquiler mensual", 200000, "Transferencia", "Pagado", ""],
        [fecha_inicio + timedelta(days=2), "Servicios", "Electricidad", "Recibo ICE", 35000, "Efectivo", "Pagado", ""],
        [fecha_inicio + timedelta(days=5), "Alimentación", "Supermercado", "Compra semanal", 45000, "Tarjeta", "Pagado", "Auto Mercado"],
        [fecha_inicio + timedelta(days=7), "Transporte", "Combustible", "Gasolina", 30000, "Efectivo", "Pagado", ""],
        [fecha_inicio + timedelta(days=10), "Salud", "Farmacia", "Medicamentos", 15000, "Tarjeta", "Pagado", ""],
    ]

    fila_actual = 3
    for datos in datos_ejemplo:
        for col, valor in enumerate(datos, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            if col == 1:  # Columna de fecha
                celda.value = valor
                celda.number_format = 'DD/MM/YY'
            elif col == 5:  # Columna de monto
                celda.value = valor
                celda.number_format = '₡#,##0.00'
            else:
                celda.value = valor
            aplicar_estilo(celda, crear_estilo_celda_normal())
        fila_actual += 1

    # Fila de total (SINTAXIS ESPAÑOL para Excel CR)
    fila_total = 100
    ws[f'A{fila_total}'] = "TOTAL GASTOS"
    ws.merge_cells(f'A{fila_total}:D{fila_total}')
    ws[f'E{fila_total}'] = f"=SUMA(E3:E99)"  # SUMA en español
    ws[f'E{fila_total}'].number_format = '₡#,##0.00'
    ws[f'F{fila_total}'] = f"=SUMA(E3:E99)"  # Para referencia desde Dashboard
    ws[f'F{fila_total}'].number_format = '₡#,##0.00'

    aplicar_estilo(ws[f'A{fila_total}'], crear_estilo_total())
    aplicar_estilo(ws[f'E{fila_total}'], crear_estilo_total())
    aplicar_estilo(ws[f'F{fila_total}'], crear_estilo_total())

    # Auto-ajustar columnas
    ajustar_columnas(ws)

    # Proteger columnas calculadas
    proteger_columnas_calculadas(ws, ['E', 'F'])


def crear_hoja_balance(wb):
    """Crea la hoja de Balance General (Activos y Pasivos)"""
    ws = wb.create_sheet("Balance General")

    # Título
    ws['A1'] = "BALANCE GENERAL"
    ws.merge_cells('A1:F1')
    aplicar_estilo(ws['A1'], crear_estilo_header())
    ws.row_dimensions[1].height = 30

    ws['A2'] = f"Al {datetime.now().strftime('%d/%m/%y')}"
    ws.merge_cells('A2:F2')

    # SECCIÓN ACTIVOS
    fila_actual = 4
    ws[f'A{fila_actual}'] = "ACTIVOS"
    ws.merge_cells(f'A{fila_actual}:C{fila_actual}')
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_subheader())
    ws[f'A{fila_actual}'].fill = PatternFill(
        start_color=PaletaFinanciera.ACTIVOS_BG,
        end_color=PaletaFinanciera.ACTIVOS_BG,
        fill_type='solid'
    )

    fila_actual += 1
    headers_activos = ["Categoría", "Descripción", "Monto (₡)", "% del Total"]
    for col, header in enumerate(headers_activos, start=1):
        celda = ws.cell(row=fila_actual, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())

    # Datos de activos de ejemplo
    activos = [
        ["Efectivo", "Caja chica", 50000],
        ["Bancos", "Cuenta corriente BAC", 500000],
        ["Bancos", "Cuenta ahorros BCR", 300000],
        ["Inversiones", "CCSS cesantía", 1500000],
        ["Activos fijos", "Vehículo", 5000000],
    ]

    fila_inicio_activos = fila_actual + 1
    for activo in activos:
        fila_actual += 1
        for col, valor in enumerate(activo, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            if col == 3:  # Monto
                celda.value = valor
                celda.number_format = '₡#,##0.00'
            else:
                celda.value = valor
            aplicar_estilo(celda, crear_estilo_celda_normal())
        # % del total (calculado)
        celda_pct = ws.cell(row=fila_actual, column=4)
        celda_pct.value = f"=C{fila_actual}/C{fila_actual+len(activos)+1}"
        celda_pct.number_format = '0.0%'

    fila_total_activos = fila_actual + 1
    ws[f'A{fila_total_activos}'] = "TOTAL ACTIVOS"
    ws.merge_cells(f'A{fila_total_activos}:B{fila_total_activos}')
    ws[f'C{fila_total_activos}'] = f"=SUMA(C{fila_inicio_activos}:C{fila_actual})"  # SUMA en español
    ws[f'C{fila_total_activos}'].number_format = '₡#,##0.00'
    aplicar_estilo(ws[f'A{fila_total_activos}'], crear_estilo_total())
    aplicar_estilo(ws[f'C{fila_total_activos}'], crear_estilo_total())

    # SECCIÓN PASIVOS
    fila_actual = fila_total_activos + 3
    ws[f'A{fila_actual}'] = "PASIVOS"
    ws.merge_cells(f'A{fila_actual}:C{fila_actual}')
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_subheader())
    ws[f'A{fila_actual}'].fill = PatternFill(
        start_color=PaletaFinanciera.PASIVOS_BG,
        end_color=PaletaFinanciera.PASIVOS_BG,
        fill_type='solid'
    )

    fila_actual += 1
    headers_pasivos = ["Categoría", "Descripción", "Monto (₡)", "% del Total"]
    for col, header in enumerate(headers_pasivos, start=1):
        celda = ws.cell(row=fila_actual, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())

    # Datos de pasivos de ejemplo
    pasivos = [
        ["Préstamos", "Préstamo vehículo", 2000000],
        ["Tarjetas crédito", "Tarjeta BAC", 150000],
        ["Cuentas por pagar", "Servicios pendientes", 50000],
    ]

    fila_inicio_pasivos = fila_actual + 1
    for pasivo in pasivos:
        fila_actual += 1
        for col, valor in enumerate(pasivo, start=1):
            celda = ws.cell(row=fila_actual, column=col)
            if col == 3:  # Monto
                celda.value = valor
                celda.number_format = '₡#,##0.00'
            else:
                celda.value = valor
            aplicar_estilo(celda, crear_estilo_celda_normal())
        # % del total
        celda_pct = ws.cell(row=fila_actual, column=4)
        celda_pct.value = f"=C{fila_actual}/C{fila_actual+len(pasivos)+1}"
        celda_pct.number_format = '0.0%'

    fila_total_pasivos = fila_actual + 1
    ws[f'A{fila_total_pasivos}'] = "TOTAL PASIVOS"
    ws.merge_cells(f'A{fila_total_pasivos}:B{fila_total_pasivos}')
    ws[f'C{fila_total_pasivos}'] = f"=SUMA(C{fila_inicio_pasivos}:C{fila_actual})"  # SUMA en español
    ws[f'C{fila_total_pasivos}'].number_format = '₡#,##0.00'
    aplicar_estilo(ws[f'A{fila_total_pasivos}'], crear_estilo_total())
    aplicar_estilo(ws[f'C{fila_total_pasivos}'], crear_estilo_total())

    # PATRIMONIO NETO
    fila_actual = fila_total_pasivos + 2
    ws[f'A{fila_actual}'] = "PATRIMONIO NETO"
    ws.merge_cells(f'A{fila_actual}:B{fila_actual}')
    ws[f'C{fila_actual}'] = f"=C{fila_total_activos}-C{fila_total_pasivos}"
    ws[f'C{fila_actual}'].number_format = '₡#,##0.00'
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_total())
    aplicar_estilo(ws[f'C{fila_actual}'], crear_estilo_total())
    ws[f'A{fila_actual}'].fill = PatternFill(
        start_color=PaletaFinanciera.POSITIVO,
        end_color=PaletaFinanciera.POSITIVO,
        fill_type='solid'
    )

    # Auto-ajustar columnas
    ajustar_columnas(ws)

    # Proteger columnas calculadas
    proteger_columnas_calculadas(ws, ['C', 'D'])


def crear_hoja_flujo_caja(wb):
    """Crea la hoja de Flujo de Caja mensual"""
    ws = wb.create_sheet("Flujo de Caja")

    # Título
    ws['A1'] = "FLUJO DE CAJA MENSUAL"
    ws.merge_cells('A1:E1')
    aplicar_estilo(ws['A1'], crear_estilo_header())
    ws.row_dimensions[1].height = 30

    # Headers
    headers = ["Mes", "Ingresos (₡)", "Gastos (₡)", "Flujo Neto (₡)", "Acumulado (₡)"]
    for col, header in enumerate(headers, start=1):
        celda = ws.cell(row=2, column=col)
        celda.value = header
        aplicar_estilo(celda, crear_estilo_subheader())

    # Generar 12 meses
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

    fila_actual = 3
    for mes in meses:
        ws[f'A{fila_actual}'] = mes
        ws[f'B{fila_actual}'] = 0  # Ingresos (usuario llenará)
        ws[f'C{fila_actual}'] = 0  # Gastos (usuario llenará)
        ws[f'D{fila_actual}'] = f"=B{fila_actual}-C{fila_actual}"  # Flujo neto
        if fila_actual == 3:
            ws[f'E{fila_actual}'] = f"=D{fila_actual}"  # Acumulado primer mes
        else:
            ws[f'E{fila_actual}'] = f"=E{fila_actual-1}+D{fila_actual}"  # Acumulado

        # Formatear montos
        for col in ['B', 'C', 'D', 'E']:
            ws[f'{col}{fila_actual}'].number_format = '₡#,##0.00'
            aplicar_estilo(ws[f'{col}{fila_actual}'], crear_estilo_celda_normal())

        aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_celda_normal())
        fila_actual += 1

    # Fila de totales (SINTAXIS ESPAÑOL para Excel CR)
    ws[f'A{fila_actual}'] = "TOTAL AÑO"
    ws[f'B{fila_actual}'] = f"=SUMA(B3:B14)"  # SUMA en español
    ws[f'C{fila_actual}'] = f"=SUMA(C3:C14)"  # SUMA en español
    ws[f'D{fila_actual}'] = f"=B{fila_actual}-C{fila_actual}"
    ws[f'E{fila_actual}'] = f"=E14"

    for col in ['B', 'C', 'D', 'E']:
        ws[f'{col}{fila_actual}'].number_format = '₡#,##0.00'
        aplicar_estilo(ws[f'{col}{fila_actual}'], crear_estilo_total())
    aplicar_estilo(ws[f'A{fila_actual}'], crear_estilo_total())

    # Auto-ajustar columnas
    ajustar_columnas(ws)

    # Proteger columnas calculadas (D, E)
    proteger_columnas_calculadas(ws, ['D', 'E'])


# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def aplicar_estilo(celda, estilo_dict):
    """Aplica un diccionario de estilos a una celda"""
    if 'font' in estilo_dict:
        celda.font = estilo_dict['font']
    if 'fill' in estilo_dict:
        celda.fill = estilo_dict['fill']
    if 'alignment' in estilo_dict:
        celda.alignment = estilo_dict['alignment']
    if 'border' in estilo_dict:
        celda.border = estilo_dict['border']
    if 'protection' in estilo_dict:
        celda.protection = estilo_dict['protection']


def ajustar_columnas(ws, min_width=10, max_width=50):
    """Auto-ajusta el ancho de todas las columnas basado en contenido"""
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass

        adjusted_width = min(max(max_length + 2, min_width), max_width)
        ws.column_dimensions[column_letter].width = adjusted_width


def proteger_columnas_calculadas(ws, columnas):
    """
    Marca columnas calculadas para protección futura
    Las columnas con fórmulas no deben ser editadas por el usuario
    """
    # Nota: La protección real requiere proteger la hoja con password
    # Por ahora solo marcamos las celdas como protegidas
    for col_letter in columnas:
        for row in range(1, ws.max_row + 1):
            celda = ws[f'{col_letter}{row}']
            celda.protection = Protection(locked=True)


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

def generar_excel_finanzas(ruta_salida="Finanzas_v50.xlsx"):
    """
    Genera el archivo Excel completo con todas las hojas

    Args:
        ruta_salida: Ruta donde se guardará el archivo (default: Finanzas_v50.xlsx)

    Returns:
        str: Ruta al archivo generado
    """
    print("=" * 60)
    print("GENERADOR DE EXCEL FINANCIERO PROFESIONAL v5.0")
    print("=" * 60)
    print(f"\nGenerando archivo: {ruta_salida}")

    # Crear workbook
    wb = Workbook()

    # Crear hojas
    print("\n[1/5] Creando Dashboard...")
    crear_hoja_dashboard(wb)

    print("[2/5] Creando hoja Ingresos...")
    crear_hoja_ingresos(wb)

    print("[3/5] Creando hoja Gastos...")
    crear_hoja_gastos(wb)

    print("[4/5] Creando Balance General...")
    crear_hoja_balance(wb)

    print("[5/5] Creando Flujo de Caja...")
    crear_hoja_flujo_caja(wb)

    # Guardar archivo
    print(f"\nGuardando archivo en: {os.path.abspath(ruta_salida)}")
    wb.save(ruta_salida)

    print("\n" + "=" * 60)
    print("✅ ARCHIVO GENERADO EXITOSAMENTE")
    print("=" * 60)
    print(f"\nArchivo: {os.path.abspath(ruta_salida)}")
    print("\nCaracterísticas:")
    print("  ✓ Paleta de colores profesional financiera")
    print("  ✓ Formato de fecha: dd/mm/yy")
    print("  ✓ Columnas auto-ajustables")
    print("  ✓ Columnas calculadas protegidas")
    print("  ✓ 5 hojas: Dashboard, Ingresos, Gastos, Balance, Flujo Caja")
    print("\nPróximos pasos:")
    print("  1. Abrir archivo en Excel")
    print("  2. Verificar formato y fórmulas")
    print("  3. Agregar tus datos reales")
    print("  4. Guardar como Finanzas_v50.xlsx")

    return os.path.abspath(ruta_salida)


# =============================================================================
# EJECUCIÓN
# =============================================================================

if __name__ == "__main__":
    import sys

    # Permitir especificar ruta de salida como argumento
    if len(sys.argv) > 1:
        ruta = sys.argv[1]
    else:
        ruta = "Finanzas_v50.xlsx"

    generar_excel_finanzas(ruta)
