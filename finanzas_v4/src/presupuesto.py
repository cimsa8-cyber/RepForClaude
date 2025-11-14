#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo PRESUPUESTO - Control presupuestario
Álvaro Velasco - Finanzas v4.0

Control de presupuesto vs real por categoría
"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from datetime import datetime


def crear_hoja_presupuesto(wb):
    """
    Crea la hoja de control presupuestario

    Secciones:
    1. Presupuesto de Ingresos por categoría
    2. Presupuesto de Gastos por categoría
    3. Comparación Presupuesto vs Real
    4. Variaciones y alertas
    """

    ws = wb.create_sheet('PRESUPUESTO')

    # Configurar ancho de columnas
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 15

    # Estilos
    titulo_principal = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
    titulo_seccion = Font(name='Segoe UI', size=12, bold=True, color='FFFFFF')
    encabezado = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    texto_normal = Font(name='Segoe UI', size=10)
    texto_editable = Font(name='Segoe UI', size=10, bold=True, color='0066CC')
    texto_total = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')

    fill_titulo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_seccion_ingreso = PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid')  # Verde
    fill_seccion_egreso = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')  # Rojo
    fill_encabezado = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
    fill_editable = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_total = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    fill_alerta = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')
    fill_ok = PatternFill(start_color='C6E0B4', end_color='C6E0B4', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # ==========================================
    # TÍTULO PRINCIPAL
    # ==========================================
    ws.merge_cells('A1:F1')
    ws['A1'] = 'CONTROL PRESUPUESTARIO - PRESUPUESTO VS REAL'
    ws['A1'].font = titulo_principal
    ws['A1'].fill = fill_titulo
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # ==========================================
    # SECCIÓN 1: PRESUPUESTO DE INGRESOS
    # ==========================================
    fila = 3
    ws.merge_cells(f'A{fila}:F{fila}')
    ws[f'A{fila}'] = '💰 PRESUPUESTO DE INGRESOS MENSUALES'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion_ingreso
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    headers = ['Categoría', 'Presupuestado', 'Real', 'Variación', '% Ejecución', 'Estado']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(fila, col)
        cell.value = header
        cell.font = encabezado
        cell.fill = fill_encabezado
        cell.border = border_thin
        cell.alignment = Alignment(horizontal='center')

    categorias_ingreso = ['Ingresos', 'Ventas', 'Proyectos', 'Otros']
    fila_inicio_ingresos = fila + 1

    for categoria in categorias_ingreso:
        fila += 1
        # Categoría
        ws[f'A{fila}'] = categoria
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        # Presupuestado (editable)
        ws[f'B{fila}'] = 0
        ws[f'B{fila}'].font = texto_editable
        ws[f'B{fila}'].fill = fill_editable
        ws[f'B{fila}'].border = border_thin
        ws[f'B{fila}'].number_format = '₡#,##0.00'

        if fila == fila_inicio_ingresos:
            comment = Comment(
                "INSTRUCCIONES:\n"
                "Ingrese el presupuesto mensual para esta categoría.\n\n"
                "EJEMPLO:\n"
                "Si presupuesta ₡500,000 en ventas, ingrese: 500000\n\n"
                "El sistema comparará automáticamente con los ingresos reales de TRANSACCIONES.",
                "Sistema v4.0"
            )
            ws[f'B{fila}'].comment = comment

        # Real (auto-calculado desde TRANSACCIONES)
        ws[f'C{fila}'] = f'=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO",TRANSACCIONES!D:D,A{fila})'
        ws[f'C{fila}'].font = texto_normal
        ws[f'C{fila}'].border = border_thin
        ws[f'C{fila}'].number_format = '₡#,##0.00'

        # Variación
        ws[f'D{fila}'] = f'=C{fila}-B{fila}'
        ws[f'D{fila}'].font = texto_normal
        ws[f'D{fila}'].border = border_thin
        ws[f'D{fila}'].number_format = '₡#,##0.00'

        # % Ejecución
        ws[f'E{fila}'] = f'=IF(B{fila}=0,0,C{fila}/B{fila})'
        ws[f'E{fila}'].font = texto_normal
        ws[f'E{fila}'].border = border_thin
        ws[f'E{fila}'].number_format = '0.0%'

        # Estado
        ws[f'F{fila}'] = f'=IF(B{fila}=0,"Sin presupuesto",IF(E{fila}>=1,"✅ Cumplido",IF(E{fila}>=0.8,"⚠️ Por debajo","🔴 Crítico")))'
        ws[f'F{fila}'].font = texto_normal
        ws[f'F{fila}'].border = border_thin
        ws[f'F{fila}'].alignment = Alignment(horizontal='center')

    # Total Ingresos
    fila += 1
    fila_total_ingresos = fila
    ws[f'A{fila}'] = 'TOTAL INGRESOS'
    ws[f'A{fila}'].font = texto_total
    ws[f'A{fila}'].fill = fill_total
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = f'=SUM(B{fila_inicio_ingresos}:B{fila-1})'
    ws[f'B{fila}'].font = texto_total
    ws[f'B{fila}'].fill = fill_total
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = f'=SUM(C{fila_inicio_ingresos}:C{fila-1})'
    ws[f'C{fila}'].font = texto_total
    ws[f'C{fila}'].fill = fill_total
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = f'=C{fila}-B{fila}'
    ws[f'D{fila}'].font = texto_total
    ws[f'D{fila}'].fill = fill_total
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    ws[f'E{fila}'] = f'=IF(B{fila}=0,0,C{fila}/B{fila})'
    ws[f'E{fila}'].font = texto_total
    ws[f'E{fila}'].fill = fill_total
    ws[f'E{fila}'].border = border_thin
    ws[f'E{fila}'].number_format = '0.0%'

    ws[f'F{fila}'] = ''
    ws[f'F{fila}'].fill = fill_total
    ws[f'F{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 2: PRESUPUESTO DE GASTOS
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:F{fila}')
    ws[f'A{fila}'] = '💸 PRESUPUESTO DE GASTOS MENSUALES'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion_egreso
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(fila, col)
        cell.value = header
        cell.font = encabezado
        cell.fill = fill_encabezado
        cell.border = border_thin
        cell.alignment = Alignment(horizontal='center')

    categorias_egreso = ['Operaciones', 'Proyectos', 'Administrativo', 'Otros']
    fila_inicio_egresos = fila + 1

    for categoria in categorias_egreso:
        fila += 1
        # Categoría
        ws[f'A{fila}'] = categoria
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        # Presupuestado (editable)
        ws[f'B{fila}'] = 0
        ws[f'B{fila}'].font = texto_editable
        ws[f'B{fila}'].fill = fill_editable
        ws[f'B{fila}'].border = border_thin
        ws[f'B{fila}'].number_format = '₡#,##0.00'

        # Real (auto-calculado desde TRANSACCIONES)
        ws[f'C{fila}'] = f'=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!D:D,A{fila}))'
        ws[f'C{fila}'].font = texto_normal
        ws[f'C{fila}'].border = border_thin
        ws[f'C{fila}'].number_format = '₡#,##0.00'

        # Variación
        ws[f'D{fila}'] = f'=B{fila}-C{fila}'
        ws[f'D{fila}'].font = texto_normal
        ws[f'D{fila}'].border = border_thin
        ws[f'D{fila}'].number_format = '₡#,##0.00'

        # % Ejecución
        ws[f'E{fila}'] = f'=IF(B{fila}=0,0,C{fila}/B{fila})'
        ws[f'E{fila}'].font = texto_normal
        ws[f'E{fila}'].border = border_thin
        ws[f'E{fila}'].number_format = '0.0%'

        # Estado (para gastos, menor es mejor)
        ws[f'F{fila}'] = f'=IF(B{fila}=0,"Sin presupuesto",IF(E{fila}<=1,"✅ Dentro",IF(E{fila}<=1.1,"⚠️ Excedido","🔴 Crítico")))'
        ws[f'F{fila}'].font = texto_normal
        ws[f'F{fila}'].border = border_thin
        ws[f'F{fila}'].alignment = Alignment(horizontal='center')

    # Total Gastos
    fila += 1
    fila_total_egresos = fila
    ws[f'A{fila}'] = 'TOTAL GASTOS'
    ws[f'A{fila}'].font = texto_total
    ws[f'A{fila}'].fill = fill_total
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = f'=SUM(B{fila_inicio_egresos}:B{fila-1})'
    ws[f'B{fila}'].font = texto_total
    ws[f'B{fila}'].fill = fill_total
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = f'=SUM(C{fila_inicio_egresos}:C{fila-1})'
    ws[f'C{fila}'].font = texto_total
    ws[f'C{fila}'].fill = fill_total
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = f'=B{fila}-C{fila}'
    ws[f'D{fila}'].font = texto_total
    ws[f'D{fila}'].fill = fill_total
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    ws[f'E{fila}'] = f'=IF(B{fila}=0,0,C{fila}/B{fila})'
    ws[f'E{fila}'].font = texto_total
    ws[f'E{fila}'].fill = fill_total
    ws[f'E{fila}'].border = border_thin
    ws[f'E{fila}'].number_format = '0.0%'

    ws[f'F{fila}'] = ''
    ws[f'F{fila}'].fill = fill_total
    ws[f'F{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 3: RESUMEN GENERAL
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:F{fila}')
    ws[f'A{fila}'] = '📊 RESUMEN GENERAL'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_titulo
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Concepto'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Presupuestado'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Real'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Variación'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    # Utilidad presupuestada vs real
    fila += 1
    ws[f'A{fila}'] = 'UTILIDAD NETA'
    ws[f'A{fila}'].font = texto_total
    ws[f'A{fila}'].fill = fill_total
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = f'=B{fila_total_ingresos}-B{fila_total_egresos}'
    ws[f'B{fila}'].font = texto_total
    ws[f'B{fila}'].fill = fill_total
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = f'=C{fila_total_ingresos}-C{fila_total_egresos}'
    ws[f'C{fila}'].font = texto_total
    ws[f'C{fila}'].fill = fill_total
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = f'=C{fila}-B{fila}'
    ws[f'D{fila}'].font = texto_total
    ws[f'D{fila}'].fill = fill_total
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    print("   ✅ Hoja PRESUPUESTO creada (presupuestos editables, real auto-calculado)")
    return ws
