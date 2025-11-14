#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo IVA_CONTROL - Control de IVA para declaraciones fiscales
Álvaro Velasco - Finanzas v4.0

IMPORTANTE:
- Empresas de zona franca NO pagan IVA: VWR Internacional Ltda, RS Hughes
- Todos los cálculos son automáticos desde TRANSACCIONES
- Usa la tasa de IVA desde CONFIG
"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime


def crear_hoja_iva_control(wb):
    """
    Crea la hoja de control de IVA

    Secciones:
    1. Resumen IVA del mes
    2. IVA Compras (Crédito Fiscal)
    3. IVA Ventas (Débito Fiscal)
    4. Balance IVA
    5. Detalle por entidad (excluyendo zona franca)
    """

    ws = wb.create_sheet('IVA_CONTROL')

    # Configurar ancho de columnas
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 30

    # Estilos
    titulo_principal = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
    titulo_seccion = Font(name='Segoe UI', size=12, bold=True, color='FFFFFF')
    encabezado = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    texto_normal = Font(name='Segoe UI', size=10)
    texto_bold = Font(name='Segoe UI', size=10, bold=True)
    texto_total = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')

    fill_titulo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_seccion = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    fill_encabezado = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
    fill_credito = PatternFill(start_color='C6E0B4', end_color='C6E0B4', fill_type='solid')  # Verde claro
    fill_debito = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')  # Amarillo
    fill_balance = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')  # Azul
    fill_exento = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')  # Naranja claro

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
    ws['A1'] = 'CONTROL DE IVA - DECLARACIÓN FISCAL'
    ws['A1'].font = titulo_principal
    ws['A1'].fill = fill_titulo
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # Fecha de reporte
    ws.merge_cells('A2:E2')
    ws['A2'] = f'Reporte generado: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A2'].font = Font(name='Segoe UI', size=9, italic=True)
    ws['A2'].alignment = Alignment(horizontal='center')

    # ==========================================
    # SECCIÓN 1: RESUMEN IVA
    # ==========================================
    fila = 4
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📊 RESUMEN IVA'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Concepto'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Monto'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'IVA (13%)'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Estado'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    # IVA COMPRAS (Crédito Fiscal)
    fila += 1
    fila_iva_compras = fila
    ws[f'A{fila}'] = '✅ IVA Compras (Crédito Fiscal)'
    ws[f'A{fila}'].font = texto_bold
    ws[f'A{fila}'].fill = fill_credito
    ws[f'A{fila}'].border = border_thin

    # Suma de todos los montos de EGRESOS (excluyendo zona franca)
    # Excluir VWR y RS Hughes
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO")-SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!F:F,"VWR*")-SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!F:F,"RS Hughes*")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    # IVA de compras (columna M en TRANSACCIONES)
    ws[f'C{fila}'] = '=SUMIFS(TRANSACCIONES!M:M,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO")-SUMIFS(TRANSACCIONES!M:M,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!F:F,"VWR*")-SUMIFS(TRANSACCIONES!M:M,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!F:F,"RS Hughes*")'
    ws[f'C{fila}'].font = texto_bold
    ws[f'C{fila}'].fill = fill_credito
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = 'A favor del contribuyente'
    ws[f'D{fila}'].font = Font(name='Segoe UI', size=9, italic=True)
    ws[f'D{fila}'].border = border_thin

    # IVA VENTAS (Débito Fiscal)
    fila += 1
    fila_iva_ventas = fila
    ws[f'A{fila}'] = '💰 IVA Ventas (Débito Fiscal)'
    ws[f'A{fila}'].font = texto_bold
    ws[f'A{fila}'].fill = fill_debito
    ws[f'A{fila}'].border = border_thin

    # Suma de todos los montos de INGRESOS
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    # IVA de ventas
    ws[f'C{fila}'] = '=SUMIFS(TRANSACCIONES!M:M,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'C{fila}'].font = texto_bold
    ws[f'C{fila}'].fill = fill_debito
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = 'A pagar a Hacienda'
    ws[f'D{fila}'].font = Font(name='Segoe UI', size=9, italic=True)
    ws[f'D{fila}'].border = border_thin

    # BALANCE IVA
    fila += 1
    fila_balance = fila
    ws[f'A{fila}'] = '📈 BALANCE IVA (Débito - Crédito)'
    ws[f'A{fila}'].font = texto_total
    ws[f'A{fila}'].fill = fill_balance
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = ''
    ws[f'B{fila}'].border = border_thin

    # Balance = IVA Ventas - IVA Compras
    ws[f'C{fila}'] = f'=C{fila_iva_ventas}-ABS(C{fila_iva_compras})'
    ws[f'C{fila}'].font = texto_total
    ws[f'C{fila}'].fill = fill_balance
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    # Estado condicional
    ws[f'D{fila}'] = f'=IF(C{fila}>0,"A PAGAR",IF(C{fila}<0,"A FAVOR","EN CERO"))'
    ws[f'D{fila}'].font = texto_total
    ws[f'D{fila}'].fill = fill_balance
    ws[f'D{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 2: EMPRESAS EXENTAS DE IVA
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '🏭 EMPRESAS ZONA FRANCA (EXENTAS DE IVA)'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Empresa'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Total Compras'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'IVA (₡0)'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Estado'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    # VWR Internacional Ltda
    fila += 1
    ws[f'A{fila}'] = 'VWR Internacional Ltda'
    ws[f'A{fila}'].font = texto_normal
    ws[f'A{fila}'].fill = fill_exento
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!F:F,"VWR*",TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].fill = fill_exento
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = 0
    ws[f'C{fila}'].font = texto_bold
    ws[f'C{fila}'].fill = fill_exento
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = 'EXENTO - Zona Franca'
    ws[f'D{fila}'].font = Font(name='Segoe UI', size=9, italic=True, color='CC0000')
    ws[f'D{fila}'].fill = fill_exento
    ws[f'D{fila}'].border = border_thin

    # RS Hughes
    fila += 1
    ws[f'A{fila}'] = 'RS Hughes Co.'
    ws[f'A{fila}'].font = texto_normal
    ws[f'A{fila}'].fill = fill_exento
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!F:F,"RS Hughes*",TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO")'
    ws[f'B{fila}'].font = texto_normal
    ws[f'B{fila}'].fill = fill_exento
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '₡#,##0.00'

    ws[f'C{fila}'] = 0
    ws[f'C{fila}'].font = texto_bold
    ws[f'C{fila}'].fill = fill_exento
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = '₡#,##0.00'

    ws[f'D{fila}'] = 'EXENTO - Zona Franca'
    ws[f'D{fila}'].font = Font(name='Segoe UI', size=9, italic=True, color='CC0000')
    ws[f'D{fila}'].fill = fill_exento
    ws[f'D{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 3: INSTRUCCIONES
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📋 INSTRUCCIONES PARA DECLARACIÓN'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    instrucciones = [
        '✅ Revise el BALANCE IVA antes de hacer la declaración mensual',
        '💡 Si el balance es positivo, debe PAGAR a Hacienda',
        '💡 Si el balance es negativo, tiene CRÉDITO FISCAL para el siguiente mes',
        '🏭 VWR y RS Hughes NO generan IVA (zona franca)',
        '⚠️ Asegúrese de que todas las facturas en TRANSACCIONES tengan el IVA correctamente registrado en columna M',
        '📅 La declaración de IVA se hace mensualmente (15 del mes siguiente)'
    ]

    fill_instrucciones = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
    for instruccion in instrucciones:
        ws[f'A{fila}'] = instruccion
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].fill = fill_instrucciones
        ws.merge_cells(f'A{fila}:E{fila}')
        fila += 1

    print("   ✅ Hoja IVA_CONTROL creada (todo auto-calculado)")
    return ws
