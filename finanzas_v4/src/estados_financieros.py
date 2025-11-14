#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                  📊 ESTADOS FINANCIEROS - Sistema v4.0                       ║
╚══════════════════════════════════════════════════════════════════════════════╝

Genera hojas de estados financieros:
- Estado de Resultados (P&L)
- Balance General
- Flujo de Caja

RESCATADO DE v3.0:
- agregar_hojas_financieras_fase6.py ⭐⭐⭐⭐

Autor: Sistema v4.0
Fecha: 2025-11-14
"""

from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from config import HOJAS_CONFIG


def crear_hoja_estado_resultados(wb):
    """
    Crea la hoja de Estado de Resultados (P&L - Profit & Loss)

    Muestra:
    - Ingresos totales
    - Gastos por categoría
    - Utilidad/Pérdida neta
    """
    print("📊 Creando hoja ESTADO_RESULTADOS (P&L)...")

    if 'ESTADO_RESULTADOS' in wb.sheetnames:
        ws = wb['ESTADO_RESULTADOS']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('ESTADO_RESULTADOS')

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_font = Font(bold=True, color="FFFFFF", size=11)
    subtotal_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    subtotal_font = Font(bold=True, size=10)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=12)

    # Título
    ws['A1'] = 'ESTADO DE RESULTADOS (P&L)'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f'AlvaroVelascoNet SRL - Período: {datetime.now().strftime("%B %Y")}'
    ws['A2'].font = Font(size=10, italic=True)

    # Headers
    fila = 4
    ws[f'A{fila}'] = 'CONCEPTO'
    ws[f'B{fila}'] = 'MONTO (₡)'
    ws[f'C{fila}'] = 'MONTO ($)'

    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = header_fill
        ws[f'{col}{fila}'].font = header_font
        ws[f'{col}{fila}'].alignment = Alignment(horizontal='center', vertical='center')

    fila += 1

    # INGRESOS
    ws[f'A{fila}'] = '📈 INGRESOS'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    # Fórmula: Sumar todos los INGRESOS en estado COBRADO
    ws[f'A{fila}'] = 'Ventas de Servicios'
    ws[f'B{fila}'] = f'=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO",TRANSACCIONES!D:D,"Ingresos")'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    ws[f'A{fila}'] = 'Otros Ingresos'
    ws[f'B{fila}'] = f'=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO",TRANSACCIONES!D:D,"<>Ingresos")'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    # Subtotal Ingresos
    fila_subtotal_ingresos = fila
    ws[f'A{fila}'] = 'TOTAL INGRESOS'
    ws[f'B{fila}'] = f'=SUM(B{fila-2}:B{fila-1})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # EGRESOS
    ws[f'A{fila}'] = '📉 GASTOS OPERATIVOS'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    # Gastos por categoría
    categorias_gasto = ['Operaciones', 'Administrativos', 'Ventas', 'Financieros']
    fila_inicio_gastos = fila

    for categoria in categorias_gasto:
        ws[f'A{fila}'] = f'  {categoria}'
        ws[f'B{fila}'] = f'=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!D:D,"{categoria}")'
        ws[f'B{fila}'].number_format = '₡#,##0.00'
        ws[f'C{fila}'] = f'=B{fila}/540'
        ws[f'C{fila}'].number_format = '$#,##0.00'
        fila += 1

    # Subtotal Gastos
    fila_subtotal_gastos = fila
    ws[f'A{fila}'] = 'TOTAL GASTOS'
    ws[f'B{fila}'] = f'=SUM(B{fila_inicio_gastos}:B{fila-1})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # UTILIDAD/PÉRDIDA NETA
    ws[f'A{fila}'] = '💰 UTILIDAD/PÉRDIDA NETA'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    ws[f'A{fila}'] = 'RESULTADO DEL PERÍODO'
    ws[f'B{fila}'] = f'=B{fila_subtotal_ingresos}-ABS(B{fila_subtotal_gastos})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = total_fill
        ws[f'{col}{fila}'].font = total_font

    # Ajustar anchos
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18

    # Congelar
    ws.freeze_panes = 'A5'

    print("  ✅ Hoja ESTADO_RESULTADOS creada")
    return ws


def crear_hoja_balance_general(wb):
    """
    Crea la hoja de Balance General

    Muestra:
    - Activos (efectivo, CxC)
    - Pasivos (CxP, deudas TC)
    - Patrimonio
    """
    print("📊 Creando hoja BALANCE_GENERAL...")

    if 'BALANCE_GENERAL' in wb.sheetnames:
        ws = wb['BALANCE_GENERAL']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('BALANCE_GENERAL')

    # Estilos (mismos que P&L)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_font = Font(bold=True, color="FFFFFF", size=11)
    subtotal_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    subtotal_font = Font(bold=True, size=10)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=12)

    # Título
    ws['A1'] = 'BALANCE GENERAL'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f'AlvaroVelascoNet SRL - Al {datetime.now().strftime("%d/%m/%Y")}'
    ws['A2'].font = Font(size=10, italic=True)

    # Headers
    fila = 4
    ws[f'A{fila}'] = 'CONCEPTO'
    ws[f'B{fila}'] = 'MONTO (₡)'
    ws[f'C{fila}'] = 'MONTO ($)'

    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = header_fill
        ws[f'{col}{fila}'].font = header_font
        ws[f'{col}{fila}'].alignment = Alignment(horizontal='center', vertical='center')

    fila += 1

    # ACTIVOS
    ws[f'A{fila}'] = '📊 ACTIVOS'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    # Efectivo en Bancos
    ws[f'A{fila}'] = 'Efectivo en Bancos'
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"INGRESO")-SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"EGRESO")'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    # Cuentas por Cobrar
    ws[f'A{fila}'] = 'Cuentas por Cobrar'
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"POR COBRAR")'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    # Total Activos
    fila_total_activos = fila
    ws[f'A{fila}'] = 'TOTAL ACTIVOS'
    ws[f'B{fila}'] = f'=SUM(B{fila-2}:B{fila-1})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # PASIVOS
    ws[f'A{fila}'] = '📉 PASIVOS'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    # Cuentas por Pagar
    ws[f'A{fila}'] = 'Cuentas por Pagar'
    ws[f'B{fila}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PENDIENTE"))'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    # Deuda Tarjetas de Crédito (parte de Cuentas por Pagar)
    ws[f'A{fila}'] = '  (incluye Tarjetas de Crédito)'
    ws[f'A{fila}'].font = Font(italic=True, size=9)
    fila += 1

    # Total Pasivos
    fila_total_pasivos = fila
    ws[f'A{fila}'] = 'TOTAL PASIVOS'
    ws[f'B{fila}'] = f'=B{fila-2}'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # PATRIMONIO
    ws[f'A{fila}'] = '💼 PATRIMONIO'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    ws[f'A{fila}'] = 'Patrimonio Neto'
    ws[f'B{fila}'] = f'=B{fila_total_activos}-B{fila_total_pasivos}'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = total_fill
        ws[f'{col}{fila}'].font = total_font

    # Ajustar anchos
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18

    # Congelar
    ws.freeze_panes = 'A5'

    print("  ✅ Hoja BALANCE_GENERAL creada")
    return ws


def crear_hoja_flujo_caja(wb):
    """
    Crea la hoja de Flujo de Caja

    Muestra:
    - Entradas de efectivo
    - Salidas de efectivo
    - Flujo neto
    """
    print("📊 Creando hoja FLUJO_CAJA...")

    if 'FLUJO_CAJA' in wb.sheetnames:
        ws = wb['FLUJO_CAJA']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('FLUJO_CAJA')

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_font = Font(bold=True, color="FFFFFF", size=11)
    subtotal_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    subtotal_font = Font(bold=True, size=10)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=12)

    # Título
    ws['A1'] = 'FLUJO DE CAJA'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f'AlvaroVelascoNet SRL - Período: {datetime.now().strftime("%B %Y")}'
    ws['A2'].font = Font(size=10, italic=True)

    # Headers
    fila = 4
    ws[f'A{fila}'] = 'CONCEPTO'
    ws[f'B{fila}'] = 'MONTO (₡)'
    ws[f'C{fila}'] = 'MONTO ($)'

    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = header_fill
        ws[f'{col}{fila}'].font = header_font
        ws[f'{col}{fila}'].alignment = Alignment(horizontal='center', vertical='center')

    fila += 1

    # ENTRADAS
    ws[f'A{fila}'] = '📈 ENTRADAS DE EFECTIVO'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    ws[f'A{fila}'] = 'Cobros de Ventas'
    ws[f'B{fila}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    ws[f'A{fila}'] = 'Otros Ingresos'
    ws[f'B{fila}'] = '=0'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    fila_total_entradas = fila
    ws[f'A{fila}'] = 'TOTAL ENTRADAS'
    ws[f'B{fila}'] = f'=SUM(B{fila-2}:B{fila-1})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # SALIDAS
    ws[f'A{fila}'] = '📉 SALIDAS DE EFECTIVO'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    ws[f'A{fila}'] = 'Pagos a Proveedores'
    ws[f'B{fila}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO"))'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    ws[f'A{fila}'] = 'Gastos Operativos'
    ws[f'B{fila}'] = '=0'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    fila += 1

    fila_total_salidas = fila
    ws[f'A{fila}'] = 'TOTAL SALIDAS'
    ws[f'B{fila}'] = f'=SUM(B{fila-2}:B{fila-1})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = subtotal_fill
        ws[f'{col}{fila}'].font = subtotal_font
    fila += 2

    # FLUJO NETO
    ws[f'A{fila}'] = '💰 FLUJO NETO DE CAJA'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:C{fila}')
    fila += 1

    ws[f'A{fila}'] = 'FLUJO DEL PERÍODO'
    ws[f'B{fila}'] = f'=B{fila_total_entradas}-B{fila_total_salidas}'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'C{fila}'] = f'=B{fila}/540'
    ws[f'C{fila}'].number_format = '$#,##0.00'
    for col in ['A', 'B', 'C']:
        ws[f'{col}{fila}'].fill = total_fill
        ws[f'{col}{fila}'].font = total_font

    # Ajustar anchos
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18

    # Congelar
    ws.freeze_panes = 'A5'

    print("  ✅ Hoja FLUJO_CAJA creada")
    return ws


def agregar_estados_financieros(wb):
    """
    Función principal que agrega todas las hojas de estados financieros
    """
    print("\n📊 Agregando Estados Financieros al Excel...")

    crear_hoja_estado_resultados(wb)
    crear_hoja_balance_general(wb)
    crear_hoja_flujo_caja(wb)

    print("✅ Estados Financieros agregados exitosamente\n")
    return wb


if __name__ == "__main__":
    print("Este módulo debe ser importado desde generar_v4.py")
