#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                  🔍 CONCILIACIÓN BANCARIA - Sistema v4.0                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

Herramientas para conciliación bancaria:
- Hoja de conciliación
- Comparación extracto vs sistema
- Detección de diferencias

RESCATADO DE v3.0:
- comparar_extracto_promerica.py ⭐⭐⭐
- auditoria_total_facturas_vs_bancos.py ⭐⭐⭐

Autor: Sistema v4.0
Fecha: 2025-11-14
"""

from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def crear_hoja_conciliacion(wb):
    """
    Crea la hoja de Conciliación Bancaria

    Permite comparar:
    - Saldo según sistema
    - Saldo según extracto bancario
    - Diferencias
    - Transacciones pendientes
    """
    print("🔍 Creando hoja CONCILIACION_BANCARIA...")

    if 'CONCILIACION' in wb.sheetnames:
        ws = wb['CONCILIACION']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('CONCILIACION')

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_font = Font(bold=True, color="FFFFFF", size=11)
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ok_font = Font(color="006100")
    diff_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    diff_font = Font(color="9C0006")

    # Título
    ws['A1'] = 'CONCILIACIÓN BANCARIA'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f'Fecha: {datetime.now().strftime("%d/%m/%Y")}'
    ws['A2'].font = Font(size=10, italic=True)

    # ══════════════════════════════════════════════════════════════════════════
    # INSTRUCCIONES
    # ══════════════════════════════════════════════════════════════════════════

    fila = 4
    ws[f'A{fila}'] = '📋 INSTRUCCIONES:'
    ws[f'A{fila}'].font = Font(bold=True, size=11)
    fila += 1

    instrucciones = [
        '1. Ingresa el saldo según tu extracto bancario en la columna "Saldo Extracto"',
        '2. El sistema calculará automáticamente el saldo según transacciones',
        '3. La columna "Diferencia" mostrará cualquier discrepancia',
        '4. Investiga las diferencias (pueden ser transacciones en tránsito o errores)',
        '',
        '💡 TIPS:',
        '  - Verde = Saldos coinciden',
        '  - Rojo = Hay diferencia (revisar)',
        '  - Considera cheques no cobrados y depósitos en tránsito'
    ]

    for inst in instrucciones:
        ws[f'A{fila}'] = inst
        ws[f'A{fila}'].font = Font(size=9, italic=True)
        fila += 1

    fila += 1

    # ══════════════════════════════════════════════════════════════════════════
    # TABLA DE CONCILIACIÓN
    # ══════════════════════════════════════════════════════════════════════════

    # Headers
    ws[f'A{fila}'] = 'CUENTA'
    ws[f'B{fila}'] = 'SALDO SISTEMA (₡)'
    ws[f'C{fila}'] = 'SALDO EXTRACTO (₡)'
    ws[f'D{fila}'] = 'DIFERENCIA (₡)'
    ws[f'E{fila}'] = 'STATUS'

    for col in ['A', 'B', 'C', 'D', 'E']:
        ws[f'{col}{fila}'].fill = header_fill
        ws[f'{col}{fila}'].font = header_font
        ws[f'{col}{fila}'].alignment = Alignment(horizontal='center', vertical='center')

    fila += 1

    # Cuentas a conciliar
    cuentas = [
        'BNCR Ahorros Colones (***8618)',
        'BNCR Ahorros Dólares (***1066)',
        'BNCR CC Colones (***2186)',
        'BNCR CC Dólares (***9589)',
        'BNCR CC Dólares (***1112)',
        'Promerica SINPE Colones (***1708)',
        'Promerica CC Corporativa Dólares (***1774)'
    ]

    fila_inicio_cuentas = fila

    for cuenta in cuentas:
        # Nombre de cuenta
        ws[f'A{fila}'] = cuenta

        # Saldo según sistema (calculado automáticamente)
        ws[f'B{fila}'] = f'=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!G:G,A{fila},TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"INGRESO")-SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!G:G,A{fila},TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"EGRESO")'
        ws[f'B{fila}'].number_format = '₡#,##0.00'

        # Saldo según extracto (usuario debe ingresar manualmente)
        ws[f'C{fila}'].number_format = '₡#,##0.00'
        ws[f'C{fila}'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        # Diferencia
        ws[f'D{fila}'] = f'=C{fila}-B{fila}'
        ws[f'D{fila}'].number_format = '₡#,##0.00'

        # Status (formula condicional)
        ws[f'E{fila}'] = f'=IF(C{fila}="","Pendiente",IF(ABS(D{fila})<1,"✅ OK","⚠️ Revisar"))'

        # Formato condicional para diferencias
        if fila > fila_inicio_cuentas:  # Aplicar después de la primera fila
            ws[f'D{fila}'].font = diff_font
            ws[f'D{fila}'].fill = diff_fill

        fila += 1

    # ══════════════════════════════════════════════════════════════════════════
    # RESUMEN DE CONCILIACIÓN
    # ══════════════════════════════════════════════════════════════════════════

    fila += 2
    ws[f'A{fila}'] = '📊 RESUMEN'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:E{fila}')
    fila += 1

    ws[f'A{fila}'] = 'Total Saldo Sistema:'
    ws[f'B{fila}'] = f'=SUM(B{fila_inicio_cuentas}:B{fila-3})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'B{fila}'].font = Font(bold=True)
    fila += 1

    ws[f'A{fila}'] = 'Total Saldo Extractos:'
    ws[f'B{fila}'] = f'=SUM(C{fila_inicio_cuentas}:C{fila-3})'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'B{fila}'].font = Font(bold=True)
    fila += 1

    ws[f'A{fila}'] = 'DIFERENCIA TOTAL:'
    ws[f'B{fila}'] = f'=B{fila-1}-B{fila-2}'
    ws[f'B{fila}'].number_format = '₡#,##0.00'
    ws[f'B{fila}'].font = Font(bold=True, size=12)
    ws[f'B{fila}'].fill = diff_fill if True else ok_fill  # Formato condicional
    fila += 2

    ws[f'A{fila}'] = 'Cuentas conciliadas:'
    ws[f'B{fila}'] = f'=COUNTIF(E{fila_inicio_cuentas}:E{fila-4},"✅ OK")'
    ws[f'B{fila}'].font = Font(bold=True, color="006100")
    fila += 1

    ws[f'A{fila}'] = 'Cuentas con diferencias:'
    ws[f'B{fila}'] = f'=COUNTIF(E{fila_inicio_cuentas}:E{fila-4},"⚠️ Revisar")'
    ws[f'B{fila}'].font = Font(bold=True, color="9C0006")

    # ══════════════════════════════════════════════════════════════════════════
    # SECCIÓN: TRANSACCIONES EN TRÁNSITO
    # ══════════════════════════════════════════════════════════════════════════

    fila += 3
    ws[f'A{fila}'] = '🔄 TRANSACCIONES EN TRÁNSITO'
    ws[f'A{fila}'].fill = section_fill
    ws[f'A{fila}'].font = section_font
    ws.merge_cells(f'A{fila}:E{fila}')
    fila += 1

    ws[f'A{fila}'] = 'Posibles causas de diferencias:'
    ws[f'A{fila}'].font = Font(bold=True)
    fila += 1

    causas = [
        '• Cheques emitidos no cobrados',
        '• Depósitos en tránsito',
        '• Cargos bancarios no registrados',
        '• Intereses no contabilizados',
        '• Errores de captura',
        '• Transacciones del extracto no ingresadas al sistema'
    ]

    for causa in causas:
        ws[f'A{fila}'] = causa
        ws[f'A{fila}'].font = Font(size=9)
        fila += 1

    # Ajustar anchos
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 15

    # Congelar
    ws.freeze_panes = f'A{fila_inicio_cuentas}'

    print("  ✅ Hoja CONCILIACION creada")
    return ws


def agregar_conciliacion(wb):
    """
    Función principal que agrega la hoja de conciliación
    """
    print("\n🔍 Agregando Conciliación Bancaria al Excel...")

    crear_hoja_conciliacion(wb)

    print("✅ Conciliación Bancaria agregada exitosamente\n")
    return wb


if __name__ == "__main__":
    print("Este módulo debe ser importado desde generar_v4.py")
