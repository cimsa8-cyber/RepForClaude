#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo HOJA CONFIG - Configuraciones globales editables
Álvaro Velasco - Finanzas v4.0

IMPORTANTE: Esta hoja NO está protegida para permitir edición de:
- Tipo de cambio actual
- Fechas de corte de tarjetas
- Configuraciones generales
"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from openpyxl.comments import Comment


def crear_hoja_config(wb):
    """
    Crea la hoja de configuraciones globales del sistema

    Secciones:
    1. Tipo de Cambio (editable)
    2. Fechas de Corte Tarjetas (editable)
    3. Configuraciones Generales
    4. Empresas Zona Franca (exentas de IVA)
    """

    ws = wb.create_sheet('CONFIG')

    # Configurar ancho de columnas
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 20

    # Estilos
    titulo_principal = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
    titulo_seccion = Font(name='Segoe UI', size=12, bold=True, color='FFFFFF')
    encabezado = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    texto_normal = Font(name='Segoe UI', size=10)
    texto_editable = Font(name='Segoe UI', size=10, bold=True, color='0066CC')

    fill_titulo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_seccion = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    fill_encabezado = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
    fill_editable = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_info = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')

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
    ws['A1'] = 'CONFIGURACIONES DEL SISTEMA v4.0'
    ws['A1'].font = titulo_principal
    ws['A1'].fill = fill_titulo
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # ==========================================
    # SECCIÓN 1: TIPO DE CAMBIO ACTUAL
    # ==========================================
    fila = 3
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '💱 TIPO DE CAMBIO'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Tipo de Cambio Actual (USD → CRC)'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Valor'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Fecha'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    fila += 1
    ws[f'A{fila}'] = 'TC USD a CRC'
    ws[f'A{fila}'].font = texto_normal
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 540  # Valor inicial
    ws[f'B{fila}'].font = texto_editable
    ws[f'B{fila}'].fill = fill_editable
    ws[f'B{fila}'].border = border_thin
    ws[f'B{fila}'].number_format = '#,##0.00'
    ws[f'B{fila}'].alignment = Alignment(horizontal='right')

    # Comentario con instrucciones
    comment = Comment(
        "INSTRUCCIONES:\n"
        "Ingrese el tipo de cambio actual del dólar a colones.\n\n"
        "EJEMPLO:\n"
        "Si 1 USD = ₡540.50, ingrese: 540.50\n\n"
        "Este valor se usa en todas las fórmulas del sistema para convertir USD a CRC.\n"
        "Actualice este valor cuando el tipo de cambio cambie.",
        "Sistema v4.0"
    )
    ws[f'B{fila}'].comment = comment

    ws[f'C{fila}'] = datetime.now()
    ws[f'C{fila}'].font = texto_editable
    ws[f'C{fila}'].fill = fill_editable
    ws[f'C{fila}'].border = border_thin
    ws[f'C{fila}'].number_format = 'DD/MM/YYYY'

    comment_fecha = Comment(
        "INSTRUCCIONES:\n"
        "Ingrese la fecha en que actualizó el tipo de cambio.\n\n"
        "EJEMPLO:\n"
        "14/11/2025",
        "Sistema v4.0"
    )
    ws[f'C{fila}'].comment = comment_fecha

    # Guardar referencia a celda del TC
    fila_tc = fila

    # ==========================================
    # SECCIÓN 2: FECHAS DE CORTE TARJETAS
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '💳 FECHAS DE PAGO TARJETAS DE CRÉDITO'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Tarjeta'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Día de Pago'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Límite'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin

    ws[f'D{fila}'] = 'Entidad'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin

    # Tarjetas
    tarjetas = [
        ('Visa Clásica BNCR (***3519)', 17, '$1,200.00', 'BNCR'),
        ('Visa Platino BNCR (***8055)', 5, '$1,600.00', 'BNCR'),
        ('MasterCard Oro BNCR (***9208)', 13, '$5,000.00', 'BNCR'),
        ('Credomatic (***XXXX)', 15, 'Por definir', 'Credomatic')
    ]

    for tarjeta, dia, limite, entidad in tarjetas:
        fila += 1
        ws[f'A{fila}'] = tarjeta
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        ws[f'B{fila}'] = dia
        ws[f'B{fila}'].font = texto_editable
        ws[f'B{fila}'].fill = fill_editable
        ws[f'B{fila}'].border = border_thin
        ws[f'B{fila}'].alignment = Alignment(horizontal='center')

        if fila == 9:  # Primera tarjeta
            comment_dia = Comment(
                "INSTRUCCIONES:\n"
                "Ingrese el día del mes en que vence el pago de esta tarjeta.\n\n"
                "EJEMPLO:\n"
                "Si el pago vence el 17 de cada mes, ingrese: 17\n\n"
                "Este valor se usa para calcular las fechas de vencimiento en CxP.",
                "Sistema v4.0"
            )
            ws[f'B{fila}'].comment = comment_dia

        ws[f'C{fila}'] = limite
        ws[f'C{fila}'].font = texto_normal
        ws[f'C{fila}'].border = border_thin
        ws[f'C{fila}'].alignment = Alignment(horizontal='right')

        ws[f'D{fila}'] = entidad
        ws[f'D{fila}'].font = texto_normal
        ws[f'D{fila}'].border = border_thin

    # ==========================================
    # SECCIÓN 3: CONFIGURACIONES GENERALES
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '⚙️ CONFIGURACIONES GENERALES'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    ws[f'A{fila}'] = 'Configuración'
    ws[f'A{fila}'].font = encabezado
    ws[f'A{fila}'].fill = fill_encabezado
    ws[f'A{fila}'].border = border_thin

    ws[f'B{fila}'] = 'Valor'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin

    ws[f'C{fila}'] = 'Descripción'
    ws[f'C{fila}'].font = encabezado
    ws[f'C{fila}'].fill = fill_encabezado
    ws[f'C{fila}'].border = border_thin
    ws.merge_cells(f'C{fila}:E{fila}')

    configs = [
        ('Empresa', 'AV Business Solutions', 'Nombre de la empresa principal'),
        ('Propietario', 'Álvaro Velasco', 'Propietario del negocio'),
        ('Moneda Base', 'CRC', 'Moneda base para reportes'),
        ('Tasa IVA (%)', 13, 'Tasa de IVA en Costa Rica'),
        ('Días alerta CxP', 7, 'Días antes de vencimiento para alertas'),
        ('Días alerta CxC', 15, 'Días después de vencimiento para alertas')
    ]

    fila_tasa_iva = None
    for config, valor, desc in configs:
        fila += 1
        ws[f'A{fila}'] = config
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        ws[f'B{fila}'] = valor
        ws[f'B{fila}'].font = texto_editable
        ws[f'B{fila}'].fill = fill_editable
        ws[f'B{fila}'].border = border_thin

        if config == 'Tasa IVA (%)':
            fila_tasa_iva = fila

        ws[f'C{fila}'] = desc
        ws[f'C{fila}'].font = texto_normal
        ws[f'C{fila}'].fill = fill_info
        ws[f'C{fila}'].border = border_thin
        ws.merge_cells(f'C{fila}:E{fila}')

    # ==========================================
    # SECCIÓN 4: EMPRESAS ZONA FRANCA (EXENTAS DE IVA)
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

    ws[f'B{fila}'] = 'Nombre Completo'
    ws[f'B{fila}'].font = encabezado
    ws[f'B{fila}'].fill = fill_encabezado
    ws[f'B{fila}'].border = border_thin
    ws.merge_cells(f'B{fila}:C{fila}')

    ws[f'D{fila}'] = 'Notas'
    ws[f'D{fila}'].font = encabezado
    ws[f'D{fila}'].fill = fill_encabezado
    ws[f'D{fila}'].border = border_thin
    ws.merge_cells(f'D{fila}:E{fila}')

    empresas_zf = [
        ('VWR', 'VWR Internacional Ltda', 'Zona franca - NO cobra IVA'),
        ('RS Hughes', 'RS Hughes Co.', 'Zona franca - NO cobra IVA')
    ]

    for siglas, nombre, notas in empresas_zf:
        fila += 1
        ws[f'A{fila}'] = siglas
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].border = border_thin

        ws[f'B{fila}'] = nombre
        ws[f'B{fila}'].font = texto_normal
        ws[f'B{fila}'].border = border_thin
        ws.merge_cells(f'B{fila}:C{fila}')

        ws[f'D{fila}'] = notas
        ws[f'D{fila}'].font = Font(name='Segoe UI', size=10, italic=True, color='CC0000')
        ws[f'D{fila}'].fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
        ws[f'D{fila}'].border = border_thin
        ws.merge_cells(f'D{fila}:E{fila}')

    # ==========================================
    # SECCIÓN 5: INSTRUCCIONES
    # ==========================================
    fila += 2
    ws.merge_cells(f'A{fila}:E{fila}')
    ws[f'A{fila}'] = '📋 INSTRUCCIONES DE USO'
    ws[f'A{fila}'].font = titulo_seccion
    ws[f'A{fila}'].fill = fill_seccion
    ws[f'A{fila}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[fila].height = 25

    fila += 1
    instrucciones = [
        '✅ Esta hoja NO está protegida - puedes editar las celdas amarillas',
        '💡 Actualiza el tipo de cambio periódicamente para cálculos precisos',
        '📅 Verifica las fechas de pago de tarjetas mensualmente',
        '🏭 Las empresas de zona franca (VWR, RS Hughes) NO pagan IVA',
        '⚠️ NO modifiques las fórmulas en otras hojas que referencian esta CONFIG'
    ]

    for instruccion in instrucciones:
        ws[f'A{fila}'] = instruccion
        ws[f'A{fila}'].font = texto_normal
        ws[f'A{fila}'].fill = fill_info
        ws.merge_cells(f'A{fila}:E{fila}')
        fila += 1

    # ==========================================
    # NOTA SOBRE REFERENCIAS
    # ==========================================
    # Otras hojas pueden referenciar el TC usando: CONFIG!$B$5
    # No usamos nombres definidos para simplificar

    print("   ✅ Hoja CONFIG creada (DESPROTEGIDA para edición)")
    print(f"      💱 Tipo de cambio en celda: CONFIG!$B${fila_tc}")
    if fila_tasa_iva:
        print(f"      💰 Tasa IVA en celda: CONFIG!$B${fila_tasa_iva}")

    return ws
