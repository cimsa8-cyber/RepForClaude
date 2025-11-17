#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Herramienta de diagnóstico específica para hojas CxP y CxC
Verifica qué hay EXACTAMENTE en cada celda
"""

import openpyxl

def diagnosticar_hoja(ws, nombre_hoja):
    """Diagnostica una hoja específica (CxP o CxC)."""
    print(f"\n{'='*70}")
    print(f"DIAGNÓSTICO: {nombre_hoja}")
    print(f"{'='*70}\n")

    # Verificar fila 3 (donde deberían estar las fórmulas FILTER)
    print("📋 FILA 3 (Primeras fórmulas FILTER):")
    print(f"{'Celda':<10} {'Tipo':<15} {'Valor/Fórmula':<50}")
    print("-" * 75)

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
        celda_coord = f"{col}3"
        celda = ws[celda_coord]

        tipo = type(celda.value).__name__
        valor = celda.value

        if valor is None:
            valor_str = "[VACÍO]"
        elif isinstance(valor, str):
            if valor.startswith('='):
                # Es una fórmula
                valor_str = f"FÓRMULA: {valor[:60]}..."
            else:
                valor_str = f"TEXTO: {valor[:50]}"
        else:
            valor_str = f"{valor}"

        print(f"{celda_coord:<10} {tipo:<15} {valor_str}")

    # Verificar si hay datos en filas 4-7 (donde deberían aparecer las tarjetas)
    print(f"\n📊 DATOS EN FILAS 4-7 (Resultados FILTER):")
    print(f"{'Fila':<6} {'A (Entidad)':<25} {'C (Monto)':<15} {'H (Equiv USD)':<15}")
    print("-" * 75)

    for fila in range(4, 8):
        a_val = ws[f'A{fila}'].value or "[VACÍO]"
        c_val = ws[f'C{fila}'].value or "[VACÍO]"
        h_val = ws[f'H{fila}'].value or "[VACÍO]"

        print(f"{fila:<6} {str(a_val)[:24]:<25} {str(c_val):<15} {str(h_val):<15}")

    # Verificar headers
    print(f"\n📑 HEADERS (Fila 2):")
    headers = []
    for col in range(1, 13):  # 12 columnas
        header_val = ws.cell(row=2, column=col).value
        headers.append(header_val or "[VACÍO]")

    for i, header in enumerate(headers, 1):
        print(f"  Col {i}: {header}")

def diagnosticar_transacciones(ws):
    """Diagnostica la hoja TRANSACCIONES para ver los datos fuente."""
    print(f"\n{'='*70}")
    print(f"DIAGNÓSTICO: TRANSACCIONES (Datos fuente)")
    print(f"{'='*70}\n")

    print("📊 DATOS PRECARGADOS (Filas 3-6 - Las 4 tarjetas):")
    print(f"{'Fila':<6} {'B (Entidad)':<25} {'C (Cat)':<10} {'F (Monto)':<12} {'M (Estado)':<12}")
    print("-" * 75)

    for fila in range(3, 7):
        b_val = ws[f'B{fila}'].value or "[VACÍO]"
        c_val = ws[f'C{fila}'].value or "[VACÍO]"
        f_val = ws[f'F{fila}'].value or "[VACÍO]"
        m_val = ws[f'M{fila}'].value or "[VACÍO]"

        print(f"{fila:<6} {str(b_val)[:24]:<25} {str(c_val):<10} {str(f_val):<12} {str(m_val):<12}")

    # Verificar columnas calculadas S, T, U, V, W
    print(f"\n🔢 COLUMNAS CALCULADAS (Fila 3 - Primera tarjeta):")
    print(f"{'Columna':<10} {'Nombre':<20} {'Tipo':<15} {'Valor/Fórmula':<40}")
    print("-" * 85)

    cols_calc = {
        'S': 'Días Trans',
        'T': 'Equiv USD',
        'U': 'Fecha Venc',
        'V': 'Prior CxP',
        'W': 'Prior CxC'
    }

    for col, nombre in cols_calc.items():
        celda = ws[f'{col}3']
        tipo = type(celda.value).__name__
        valor = celda.value

        if isinstance(valor, str) and valor.startswith('='):
            valor_str = valor[:50] + '...' if len(valor) > 50 else valor
        else:
            valor_str = str(valor) if valor is not None else "[VACÍO]"

        print(f"{col:<10} {nombre:<20} {tipo:<15} {valor_str}")

if __name__ == "__main__":
    archivo = "AlvaroVelasco_Finanzas_v5.0_COMPLETO.xlsx"

    print(f"\n{'='*70}")
    print(f"DIAGNÓSTICO DETALLADO: CxP y CxC")
    print(f"Archivo: {archivo}")
    print(f"{'='*70}")

    wb = openpyxl.load_workbook(archivo, data_only=False)

    # Diagnosticar TRANSACCIONES primero (fuente de datos)
    if 'TRANSACCIONES' in wb.sheetnames:
        diagnosticar_transacciones(wb['TRANSACCIONES'])

    # Diagnosticar CxP
    if 'CxP' in wb.sheetnames:
        diagnosticar_hoja(wb['CxP'], 'CxP (Cuentas por Pagar)')

    # Diagnosticar CxC
    if 'CxC' in wb.sheetnames:
        diagnosticar_hoja(wb['CxC'], 'CxC (Cuentas por Cobrar)')

    print(f"\n{'='*70}")
    print("DIAGNÓSTICO COMPLETADO")
    print(f"{'='*70}\n")

    print("📌 INTERPRETACIÓN:")
    print("  - Si A3 muestra '[VACÍO]': La fórmula FILTER NO se generó o Excel la removió")
    print("  - Si A3 muestra 'FÓRMULA': La fórmula existe pero puede no funcionar")
    print("  - Si filas 4-7 tienen datos: Las tarjetas aparecen (FILTER funciona)")
    print("  - Si filas 4-7 están vacías: Las tarjetas NO aparecen (FILTER no funciona)")
    print()
