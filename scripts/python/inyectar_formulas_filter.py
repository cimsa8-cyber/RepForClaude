#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Herramienta para inyectar fórmulas FILTER en Excel existente
Usa técnica diferente que puede evitar problemas de compatibilidad
"""

import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy

def inyectar_formulas_cxp(wb):
    """Inyecta fórmulas FILTER en hoja CxP."""
    print("\n📝 Inyectando fórmulas en CxP...")

    ws = wb['CxP']

    # Definir fórmulas FILTER con sintaxis para Excel 365
    # CRÍTICO: Usar punto y coma (;) como separador
    formulas = {
        'A3': '=FILTER(TRANSACCIONES!B3:B1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"Sin CxP pendientes")',
        'B3': '=FILTER(TRANSACCIONES!A3:A1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'C3': '=FILTER(TRANSACCIONES!F3:F1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'D3': '=FILTER(TRANSACCIONES!E3:E1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'E3': '=FILTER(TRANSACCIONES!G3:G1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'F3': '=FILTER(TRANSACCIONES!M3:M1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'G3': '=FILTER(TRANSACCIONES!S3:S1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'H3': '=FILTER(TRANSACCIONES!T3:T1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'I3': '=FILTER(TRANSACCIONES!N3:N1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'J3': '=FILTER(TRANSACCIONES!U3:U1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'K3': '=FILTER(TRANSACCIONES!V3:V1000;(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
    }

    for celda, formula in formulas.items():
        ws[celda] = formula
        print(f"  ✓ {celda}: {formula[:60]}...")

    # Aplicar formatos
    ws['B3'].number_format = 'DD/MM/YY'
    ws['C3'].number_format = '#,##0.00'
    ws['H3'].number_format = '$#,##0.00'
    ws['J3'].number_format = 'DD/MM/YY'

    print("  ✓ Formatos aplicados")

def inyectar_formulas_cxc(wb):
    """Inyecta fórmulas FILTER en hoja CxC."""
    print("\n📝 Inyectando fórmulas en CxC...")

    ws = wb['CxC']

    # Definir fórmulas FILTER
    formulas = {
        'A3': '=FILTER(TRANSACCIONES!B3:B1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"Sin CxC pendientes")',
        'B3': '=FILTER(TRANSACCIONES!A3:A1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'C3': '=FILTER(TRANSACCIONES!F3:F1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'D3': '=FILTER(TRANSACCIONES!E3:E1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'E3': '=FILTER(TRANSACCIONES!G3:G1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'F3': '=FILTER(TRANSACCIONES!M3:M1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'G3': '=FILTER(TRANSACCIONES!S3:S1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'H3': '=FILTER(TRANSACCIONES!T3:T1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'I3': '=FILTER(TRANSACCIONES!N3:N1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'J3': '=FILTER(TRANSACCIONES!U3:U1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
        'K3': '=FILTER(TRANSACCIONES!W3:W1000;(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente");"")',
    }

    for celda, formula in formulas.items():
        ws[celda] = formula
        print(f"  ✓ {celda}: {formula[:60]}...")

    # Aplicar formatos
    ws['B3'].number_format = 'DD/MM/YY'
    ws['C3'].number_format = '#,##0.00'
    ws['H3'].number_format = '$#,##0.00'
    ws['J3'].number_format = 'DD/MM/YY'

    print("  ✓ Formatos aplicados")

def restaurar_dropdowns(wb):
    """Restaura los dropdowns en TRANSACCIONES que se pierden al editar manualmente."""
    print("\n📋 Restaurando dropdowns en TRANSACCIONES...")

    ws = wb['TRANSACCIONES']

    from openpyxl.worksheet.datavalidation import DataValidation

    dropdowns = [
        ('C3:C1000', 'Ingreso,Gasto,Inventario,Servicios,CxP,CxC,Marketing,Personal', 'Categoría'),
        ('E3:E1000', 'USD,CRC', 'Moneda'),
        ('H3:H1000', 'Efectivo,Transferencia,Tarjeta Crédito,Cheque,SINPE,PayPal', 'Método Pago'),
        ('I3:I1000', 'Sí,No', 'Aplica IVA'),
        ('K3:K1000', 'Sí,No', 'Zona Franca'),
        ('M3:M1000', 'Pagado,Pendiente,Cobrado,Cancelado', 'Estado'),
        ('P3:P1000', 'Personal,Negocio', 'Tipo Gasto'),
    ]

    for rango, valores, nombre in dropdowns:
        dv = DataValidation(type="list", formula1=f'"{valores}"', allow_blank=False)
        dv.error = 'Valor inválido'
        dv.errorTitle = f'Error en {nombre}'
        dv.prompt = f'Seleccione {nombre}'
        dv.promptTitle = nombre
        ws.add_data_validation(dv)
        dv.add(rango)
        print(f"  ✓ Dropdown {nombre}: {rango}")

    print("  ✓ Todos los dropdowns restaurados")

if __name__ == "__main__":
    archivo = "AlvaroVelasco_Finanzas_v5.0_COMPLETO.xlsx"
    archivo_salida = "AlvaroVelasco_Finanzas_v5.0_COMPLETO_INYECTADO.xlsx"

    print(f"\n{'='*70}")
    print(f"INYECTOR DE FÓRMULAS FILTER")
    print(f"{'='*70}")
    print(f"\nArchivo entrada: {archivo}")
    print(f"Archivo salida: {archivo_salida}")

    try:
        # Cargar archivo existente
        print("\n📂 Cargando archivo Excel...")
        wb = openpyxl.load_workbook(archivo)
        print(f"  ✓ Archivo cargado correctamente")
        print(f"  ✓ Hojas encontradas: {', '.join(wb.sheetnames)}")

        # Inyectar fórmulas
        if 'CxP' in wb.sheetnames:
            inyectar_formulas_cxp(wb)
        else:
            print("  ⚠️ Hoja CxP no encontrada")

        if 'CxC' in wb.sheetnames:
            inyectar_formulas_cxc(wb)
        else:
            print("  ⚠️ Hoja CxC no encontrada")

        if 'TRANSACCIONES' in wb.sheetnames:
            restaurar_dropdowns(wb)
        else:
            print("  ⚠️ Hoja TRANSACCIONES no encontrada")

        # Guardar archivo
        print(f"\n💾 Guardando archivo: {archivo_salida}...")
        wb.save(archivo_salida)
        print(f"  ✓ Archivo guardado correctamente")

        print(f"\n{'='*70}")
        print("✅ PROCESO COMPLETADO")
        print(f"{'='*70}")
        print(f"\n📌 PRÓXIMOS PASOS:")
        print(f"1. Abrir el archivo: {archivo_salida}")
        print(f"2. Excel pedirá 'Enable Editing' - hacer clic")
        print(f"3. Si Excel muestra warning de fórmulas, aceptar 'Sí' para calcular")
        print(f"4. Verificar que CxP y CxC muestran las 4 tarjetas")
        print(f"5. Verificar que los dropdowns funcionan en TRANSACCIONES")
        print()

    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
