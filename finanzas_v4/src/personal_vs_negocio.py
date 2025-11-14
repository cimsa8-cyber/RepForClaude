#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo PERSONAL_VS_NEGOCIO - Análisis de gastos personales vs negocio
Álvaro Velasco - Finanzas v4.0

OBJETIVO: Separar claramente gastos personales de Álvaro vs gastos del negocio
para sanear las finanzas y evitar mezclar fondos.
"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from datetime import datetime


def crear_hoja_personal_vs_negocio(wb):
    """
    Crea la hoja de análisis Personal vs Negocio

    Secciones:
    1. Resumen comparativo
    2. Desglose por categoría
    3. Gráficas de comparación
    4. Alertas de gastos mal clasificados
    """

    ws = wb.create_sheet('PERSONAL_VS_NEGOCIO')

    # Configurar ancho de columnas
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 25

    # Estilos
    titulo_principal = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
    titulo_seccion = Font(name='Segoe UI', size=12, bold=True, color='FFFFFF')
    encabezado = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    texto_normal = Font(name='Segoe UI', size=10)
    texto_bold = Font(name='Segoe UI', size=11, bold=True)
    texto_total = Font(name='Segoe UI', size=12, bold=True, color='FFFFFF')

    fill_titulo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_seccion = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    fill_encabezado = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
    fill_personal = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')  # Amarillo
    fill_negocio = PatternFill(start_color='A9D08E', end_color='A9D08E', fill_type='solid')  # Verde claro
    fill_total = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    fill_alerta = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # ==========================================
    # TÍTULO PRINCIPAL
    # ==========================================
    ws.merge_cells('A1:E1')
    ws['A1'] = '🏢 ANÁLISIS: GASTOS PERSONALES VS NEGOCIO'
    ws['A1'].font = titulo_principal
    ws['A1'].fill = fill_titulo
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws.merge_cells('A2:E2')
    ws['A2'] = 'Objetivo: Sanear finanzas separando claramente lo personal de lo del negocio'
    ws['A2'].font = Font(name='Segoe UI', size=9, italic=True)
    ws['A2'].alignment = Alignment(horizontal='center')

    # ==========================================
    # SECCIÓN 1: RESUMEN COMPARATIVO
    # ==========================================
    fila = 4
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📊 RESUMEN COMPARATIVO'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Tipo'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Total Ingresos'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Total Gastos'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Balance'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    ws[f'E{fila}'] = '% del Total'
    ws[f'E{fila}'].font = encabezado
    ws[f'E{fila}'].fill = fill_encabezado
    ws[f'E{fila}'].border = border_thin

    # NEGOCIO
    fila += 1
    fila_negocio = fila
    ws[f'A{fila}'] = '🏢 NEGOCIO'
    ws[f'A{fila}'].font = texto_bold
    ws[f'A{fila}'].fill = fill_negocio
    ws[f'A{fila}'].border = border_thin

    # Ingresos negocio
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!P:P,"Negocio",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].fill = fill_negocio
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    # Gastos negocio
    ws[f'C{fila}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!P:P,"Negocio",TRANSACCIONES!L:L,"PAGADO"))'
    ws[f'C{fila}'].font = texto_normal
    ws[f'C{fila}'].fill = fill_negocio
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    # Balance negocio
    ws[f'D{fila}'] = f'=B{fila}-C{fila}'
    ws[f'D{fila}'].font = texto_bold
    ws[f'D{fila}'].fill = fill_negocio
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    # % del total
    ws[f'E{fila}'] = f'=C{fila}/(C{fila}+C{fila+1})'
    ws[f'E{fila}'].font = texto_normal
    ws[f'E{fila}'].fill = fill_negocio
    ws[f'E{fila}'].border = border_thin
    ws[f'E{fila}'].number_format = '0.0%'

    # PERSONAL
    fila += 1
    fila_personal = fila
    ws[f'A{fila}'] = '👤 PERSONAL (Álvaro)'
    ws[f'A{fila}'].font = texto_bold
    ws[f'A{fila}'].fill = fill_personal
    ws[f'A{fila}'].border = border_thin

    # Ingresos personales
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!P:P,"Personal",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].fill = fill_personal
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    # Gastos personales
    ws[f'C{fila}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!P:P,"Personal",TRANSACCIONES!L:L,"PAGADO"))'
    ws[f'C{fila}'].font = texto_normal
    ws[f'C{fila}'].fill = fill_personal
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    # Balance personal
    ws[f'D{fila}'] = f'=B{fila}-C{fila}'
    ws[f'D{fila}'].font = texto_bold
    ws[f'D{fila}'].fill = fill_personal
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    # % del total
    ws[f'E{fila}'] = f'=C{fila}/(C{fila_negocio}+C{fila})'
    ws[f'E{fila}'].font = texto_normal
    ws[f'E{fila}'].fill = fill_personal
    ws[f'E{fila}'].border = border_thin
    ws[f'E{fila}'].number_format = '0.0%'

    # TOTAL
    fila += 1
    fila_total = fila
    ws[f'A{fila}'] = 'TOTAL'
    ws[f'A{fila}'].font = texto_total
    ws[f'A{fila}'].fill = fill_total
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = f'=B{fila_negocio}+B{fila_personal}'
    ws[f'B{fila}'].font = texto_total
    ws[f'B{fila}'].fill = fill_total
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = f'=C{fila_negocio}+C{fila_personal}'
    ws[f'C{fila}'].font = texto_total
    ws[f'C{fila}'].fill = fill_total
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = f'=D{fila_negocio}+D{fila_personal}'
    ws[f'D{fila}'].font = texto_total
    ws[f'D{fila}'].fill = fill_total
    ws[f'D{fila}'].border = border_thin
    ws[f'D{fila}'].number_format = '₡#,##0.00'

    ws[f'E{fila}'] = '100%'
    ws[f'E{fila}'].font = texto_total
    ws[f'E{fila}'].fill = fill_total
    ws[f'E{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 2: DESGLOSE POR CATEGORÍA
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📋 DESGLOSE DE GASTOS POR CATEGORÍA'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Categoría'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Negocio'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Personal'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Total'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    ws[f'E{fila}'] = '% Personal'
    ws[f'E{fila}'].font = encabezado
    ws[f'E{fila}'].fill = fill_encabezado
    ws[f'E{fila}'].border = border_thin

    categorias = ['Operaciones', 'Proyectos', 'Administrativo', 'Ventas', 'Otros']

    for categoria in categorias:
        fila += 1
        # Categoría
        ws[f'A{fila}'] = categoria
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        # Negocio
        ws[f'B{fila}'] = f'=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!D:D,A{fila},TRANSACCIONES!P:P,"Negocio",TRANSACCIONES!L:L,"PAGADO"))'
        ws[f'B{fila}'].font = texto_normal
        ws[f'B{fila}'].fill = fill_negocio
        ws[f'B{fila}'].border = border_thin
        ws[f'B{fila}'].number_format = '₡#,##0.00'

        # Personal
        ws[f'C{fila}'] = f'=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!D:D,A{fila},TRANSACCIONES!P:P,"Personal",TRANSACCIONES!L:L,"PAGADO"))'
        ws[f'C{fila}'].font = texto_normal
        ws[f'C{fila}'].fill = fill_personal
        ws[f'C{fila}'].border = border_thin
        ws[f'C{fila}'].number_format = '₡#,##0.00'

        # Total
        ws[f'D{fila}'] = f'=B{fila}+C{fila}'
        ws[f'D{fila}'].font = texto_normal
        ws[f'D{fila}'].border = border_thin
        ws[f'D{fila}'].number_format = '₡#,##0.00'

        # % Personal
        ws[f'E{fila}'] = f'=IF(D{fila}=0,0,C{fila}/D{fila})'
        ws[f'E{fila}'].font = texto_normal
        ws[f'E{fila}'].border = border_thin
        ws[f'E{fila}'].number_format = '0.0%'

    # ==========================================
    # SECCIÓN 3: ALERTAS Y RECOMENDACIONES
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '⚠️ ALERTAS Y RECOMENDACIONES'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Alerta'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin
    ws.merge_cells(f'A{fila}:B{fila}')

    ws[f'C{fila}'] = 'Valor'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Estado'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin
    ws.merge_cells(f'D{fila}:E{fila}')

    # Alerta 1: % gastos personales
    fila += 1
    ws[f'A{fila}'] = 'Porcentaje de gastos personales'
    ws[f'A{fila}'].font = texto_normal
    ws[f'A{fila}'].border = border_thin
    ws.merge_cells(f'A{fila}:B{fila}')

    ws[f'C{fila}'] = f'=E{fila_personal}'
    ws[f'C{fila}'].font = texto_normal
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '0.0%'

    ws[f'D{fila}'] = f'=IF(C{fila}>0.5,"🔴 MUY ALTO - Reducir gastos personales",IF(C{fila}>0.3,"⚠️ MODERADO - Revisar clasificación","✅ OK - Bien separado"))'
    ws[f'D{fila}'].font = texto_normal
    ws[f'D{fila}'].border = border_thin
    ws.merge_cells(f'D{fila}:E{fila}')

    # Alerta 2: Balance personal negativo
    fila += 1
    ws[f'A{fila}'] = 'Balance personal (debe ser positivo)'
    ws[f'A{fila}'].font = texto_normal
    ws[f'A{fila}'].border = border_thin
    ws.merge_cells(f'A{fila}:B{fila}')

    ws[f'C{fila}'] = f'=D{fila_personal}'
    ws[f'C{fila}'].font = texto_normal
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = f'=IF(C{fila}<0,"🔴 NEGATIVO - Gastos personales exceden ingresos","✅ OK - Balance positivo")'
    ws[f'D{fila}'].font = texto_normal
    ws[f'D{fila}'].border = border_thin
    ws.merge_cells(f'D{fila}:E{fila}')

    # ==========================================
    # SECCIÓN 4: INSTRUCCIONES
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📋 INSTRUCCIONES PARA SANEAMIENTO'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    instrucciones = [
        '✅ Al ingresar transacciones en TRANSACCIONES, marca en columna P si es "Personal" o "Negocio"',
        '💡 Gastos de Álvaro (comida, transporte personal, etc.) = Personal',
        '💡 Gastos del negocio (nómina, proveedores, servicios) = Negocio',
        '⚠️ NUNCA mezcles fondos - separa claramente las cuentas',
        '📊 Revisa este reporte mensualmente para verificar que la separación sea correcta',
        '🎯 META: Lograr que gastos personales sean < 30% del total'
    ]

    fill_instrucciones = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
    for instruccion in instrucciones:
        ws[f'A{fila}'] = instruccion
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].fill = fill_instrucciones
        ws.merge_cells(f'A{fila}:E{fila}')
        fila += 1

    print("   ✅ Hoja PERSONAL_VS_NEGOCIO creada (auto-calculada desde TRANSACCIONES columna P)")
    return ws
