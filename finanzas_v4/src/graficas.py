#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     📈 GRÁFICAS - Sistema v4.0                               ║
╚══════════════════════════════════════════════════════════════════════════════╝

Genera gráficas y dashboards visuales:
- Dashboard de KPIs
- Gráficas de tendencias
- Análisis visual

RESCATADO DE v3.0:
- crear_graficas_profesionales.py ⭐⭐⭐⭐

Autor: Sistema v4.0
Fecha: 2025-11-14
"""

from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.utils import get_column_letter


def crear_hoja_dashboard_visual(wb):
    """
    Crea una hoja de Dashboard Visual con KPIs y gráficas

    Incluye:
    - KPIs principales (números grandes y claros)
    - Gráfica de ingresos vs gastos
    - Gráfica de CxP por vencimiento
    - Gráfica de deuda por tarjeta
    """
    print("📈 Creando hoja DASHBOARD_VISUAL...")

    if 'DASHBOARD_VISUAL' in wb.sheetnames:
        ws = wb['DASHBOARD_VISUAL']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('DASHBOARD_VISUAL')

    # Estilos para KPIs
    kpi_title_font = Font(bold=True, size=10, color="666666")
    kpi_value_font = Font(bold=True, size=24, color="1F4E78")
    kpi_positive_font = Font(bold=True, size=24, color="00B050")
    kpi_negative_font = Font(bold=True, size=24, color="FF0000")
    kpi_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    kpi_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # Título
    ws['A1'] = '📊 DASHBOARD FINANCIERO'
    ws['A1'].font = Font(bold=True, size=16)
    ws['A2'] = f'Actualizado: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A2'].font = Font(size=9, italic=True, color="666666")

    # ══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 1: KPIs PRINCIPALES
    # ══════════════════════════════════════════════════════════════════════════

    fila = 4

    # KPI 1: Efectivo Disponible
    ws[f'A{fila}'] = 'EFECTIVO DISPONIBLE'
    ws[f'A{fila}'].font = kpi_title_font
    ws[f'A{fila+1}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"INGRESO")-SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!L:L,"PAGADO",TRANSACCIONES!C:C,"EGRESO")'
    ws[f'A{fila+1}'].font = kpi_positive_font
    ws[f'A{fila+1}'].number_format = '₡#,##0.00'
    ws[f'A{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'A{fila}:C{fila}')
    ws.merge_cells(f'A{fila+1}:C{fila+1}')
    for col in ['A', 'B', 'C']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    # KPI 2: Cuentas por Cobrar
    ws[f'E{fila}'] = 'CUENTAS POR COBRAR'
    ws[f'E{fila}'].font = kpi_title_font
    ws[f'E{fila+1}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"POR COBRAR")'
    ws[f'E{fila+1}'].font = kpi_value_font
    ws[f'E{fila+1}'].number_format = '₡#,##0.00'
    ws[f'E{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'E{fila}:G{fila}')
    ws.merge_cells(f'E{fila+1}:G{fila+1}')
    for col in ['E', 'F', 'G']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    # KPI 3: Cuentas por Pagar
    ws[f'I{fila}'] = 'CUENTAS POR PAGAR'
    ws[f'I{fila}'].font = kpi_title_font
    ws[f'I{fila+1}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PENDIENTE"))'
    ws[f'I{fila+1}'].font = kpi_negative_font
    ws[f'I{fila+1}'].number_format = '₡#,##0.00'
    ws[f'I{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'I{fila}:K{fila}')
    ws.merge_cells(f'I{fila+1}:K{fila+1}')
    for col in ['I', 'J', 'K']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    fila += 3

    # KPI 4: Ingresos del Mes
    ws[f'A{fila}'] = 'INGRESOS DEL MES'
    ws[f'A{fila}'].font = kpi_title_font
    ws[f'A{fila+1}'] = '=SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"INGRESO",TRANSACCIONES!L:L,"COBRADO")'
    ws[f'A{fila+1}'].font = kpi_positive_font
    ws[f'A{fila+1}'].number_format = '₡#,##0.00'
    ws[f'A{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'A{fila}:C{fila}')
    ws.merge_cells(f'A{fila+1}:C{fila+1}')
    for col in ['A', 'B', 'C']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    # KPI 5: Gastos del Mes
    ws[f'E{fila}'] = 'GASTOS DEL MES'
    ws[f'E{fila}'].font = kpi_title_font
    ws[f'E{fila+1}'] = '=ABS(SUMIFS(TRANSACCIONES!I:I,TRANSACCIONES!C:C,"EGRESO",TRANSACCIONES!L:L,"PAGADO"))'
    ws[f'E{fila+1}'].font = kpi_negative_font
    ws[f'E{fila+1}'].number_format = '₡#,##0.00'
    ws[f'E{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'E{fila}:G{fila}')
    ws.merge_cells(f'E{fila+1}:G{fila+1}')
    for col in ['E', 'F', 'G']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    # KPI 6: Utilidad/Pérdida
    ws[f'I{fila}'] = 'UTILIDAD/PÉRDIDA'
    ws[f'I{fila}'].font = kpi_title_font
    ws[f'I{fila+1}'] = f'=A{fila+1}-E{fila+1}'
    ws[f'I{fila+1}'].font = kpi_value_font
    ws[f'I{fila+1}'].number_format = '₡#,##0.00'
    ws[f'I{fila+1}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'I{fila}:K{fila}')
    ws.merge_cells(f'I{fila+1}:K{fila+1}')
    for col in ['I', 'J', 'K']:
        for r in [fila, fila+1]:
            ws[f'{col}{r}'].fill = kpi_fill
            ws[f'{col}{r}'].border = kpi_border

    # ══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 2: DATOS PARA GRÁFICAS
    # ══════════════════════════════════════════════════════════════════════════

    fila += 4

    # Tabla de datos: Ingresos vs Gastos
    ws[f'A{fila}'] = 'CONCEPTO'
    ws[f'B{fila}'] = 'MONTO'
    ws[f'A{fila}'].font = Font(bold=True)
    ws[f'B{fila}'].font = Font(bold=True)

    ws[f'A{fila+1}'] = 'Ingresos'
    ws[f'B{fila+1}'] = f'=A{fila-4}'  # Referencia a KPI Ingresos

    ws[f'A{fila+2}'] = 'Gastos'
    ws[f'B{fila+2}'] = f'=E{fila-4}'  # Referencia a KPI Gastos

    # Gráfica de Barras: Ingresos vs Gastos
    chart = BarChart()
    chart.title = "Ingresos vs Gastos del Mes"
    chart.style = 10
    chart.y_axis.title = 'Monto (₡)'
    chart.x_axis.title = 'Concepto'

    data = Reference(ws, min_col=2, min_row=fila, max_row=fila+2)
    cats = Reference(ws, min_col=1, min_row=fila+1, max_row=fila+2)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 10
    chart.width = 15

    ws.add_chart(chart, f'A{fila+5}')

    # ══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 3: ALERTAS Y NOTIFICACIONES
    # ══════════════════════════════════════════════════════════════════════════

    fila += 20

    ws[f'A{fila}'] = '⚠️ ALERTAS Y NOTIFICACIONES'
    ws[f'A{fila}'].font = Font(bold=True, size=12, color="FF0000")
    fila += 1

    # Alerta 1: CxP próximas a vencer (siguiente 7 días)
    ws[f'A{fila}'] = 'CxP vencen en 7 días:'
    ws[f'B{fila}'] = '=COUNTIFS(CxP!F:F,"<"&TODAY()+7,CxP!F:F,">="&TODAY())'
    ws[f'B{fila}'].font = Font(bold=True, color="FF6600")
    fila += 1

    # Alerta 2: CxP vencidas
    ws[f'A{fila}'] = 'CxP VENCIDAS:'
    ws[f'B{fila}'] = '=COUNTIF(CxP!F:F,"<"&TODAY())'
    ws[f'B{fila}'].font = Font(bold=True, color="FF0000")
    fila += 1

    # Alerta 3: CxC atrasadas (>30 días)
    ws[f'A{fila}'] = 'CxC atrasadas (+30 días):'
    ws[f'B{fila}'] = '=COUNTIF(CxC!F:F,"<"&TODAY()-30)'
    ws[f'B{fila}'].font = Font(bold=True, color="FF0000")

    # Ajustar anchos
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws.column_dimensions[col].width = 15

    print("  ✅ Hoja DASHBOARD_VISUAL creada")
    return ws


def agregar_graficas(wb):
    """
    Función principal que agrega el dashboard visual
    """
    print("\n📈 Agregando Dashboard Visual al Excel...")

    crear_hoja_dashboard_visual(wb)

    print("✅ Dashboard Visual agregado exitosamente\n")
    return wb


if __name__ == "__main__":
    print("Este módulo debe ser importado desde generar_v4.py")
