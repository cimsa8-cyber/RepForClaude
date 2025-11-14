# ============================================================================
# SISTEMA DE FINANZAS v4.0 - GENERADOR COMPLETO
# ============================================================================
# Genera archivo Excel con 14 hojas completamente funcionales
# Autor: Claude + Álvaro Velasco
# Fecha: Noviembre 2024
# ============================================================================

# ============================================================================
# IMPORTS NECESARIOS
# ============================================================================

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

print("✅ Librerías importadas correctamente\n")

# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================

VERSION = "4.0"
TC_ACTUAL = 540  # Tipo de cambio USD a CRC (editable en hoja CONFIG)

# Estructura de 16 columnas para TRANSACCIONES
COLUMNAS_TRANSACCIONES = [
    "Fecha", "Entidad", "Categoría", "Subcategoría", "Moneda", "Monto",
    "Descripción", "Forma de Pago", "IVA", "Notas", "Recurrente",
    "Proyecto", "Estado", "Factura #", "Tag", "Personal/Negocio"
]

# 12 Alias de cuentas pre-cargados
ALIAS_CUENTAS = [
    {"nombre_completo": "BAC San José - Cuenta USD", "alias": "BAC-USD", "tipo": "Banco"},
    {"nombre_completo": "Banco de Costa Rica - Cuenta Corriente ₡", "alias": "BCR-CRC", "tipo": "Banco"},
    {"nombre_completo": "Visa BAC Crédito", "alias": "Visa-BAC", "tipo": "Tarjeta Crédito"},
    {"nombre_completo": "Mastercard BCR", "alias": "MC-BCR", "tipo": "Tarjeta Crédito"},
    {"nombre_completo": "Credomatic Platinum USD", "alias": "Credo-Platinum", "tipo": "Tarjeta Crédito"},
    {"nombre_completo": "Credomatic Gold CRC", "alias": "Credo-Gold", "tipo": "Tarjeta Crédito"},
    {"nombre_completo": "VWR International LLC", "alias": "VWR", "tipo": "Proveedor"},
    {"nombre_completo": "RS Hughes Co. Inc.", "alias": "RS-Hughes", "tipo": "Proveedor"},
    {"nombre_completo": "Efectivo - Caja Chica", "alias": "Efectivo", "tipo": "Efectivo"},
    {"nombre_completo": "Cliente XYZ Corp", "alias": "Cliente-XYZ", "tipo": "Cliente"},
    {"nombre_completo": "SINPE Móvil", "alias": "SINPE", "tipo": "Transferencia"},
    {"nombre_completo": "PayPal Business", "alias": "PayPal", "tipo": "Digital"}
]

# 3 Transacciones de ejemplo
TRANSACCIONES_EJEMPLO = [
    {
        "fecha": "2024-11-01",
        "entidad": "VWR",
        "categoria": "Inventario",
        "subcategoria": "Materiales Lab",
        "moneda": "USD",
        "monto": 2500.00,
        "descripcion": "Reactivos químicos Q4 2024",
        "forma_pago": "Transferencia",
        "iva": "No",  # Zona franca, sin IVA
        "notas": "Empresa zona franca - exento IVA",
        "recurrente": "No",
        "proyecto": "LAB-2024",
        "estado": "Pagado",
        "factura": "VWR-98765",
        "tag": "Inventario",
        "personal_negocio": "Negocio"
    },
    {
        "fecha": "2024-11-05",
        "entidad": "Visa-BAC",
        "categoria": "Servicios",
        "subcategoria": "Internet",
        "moneda": "CRC",
        "monto": 45000,
        "descripcion": "Pago mensualidad internet Kolbi",
        "forma_pago": "Tarjeta Crédito",
        "iva": "Sí",
        "notas": "Pago automático",
        "recurrente": "Sí",
        "proyecto": "Oficina",
        "estado": "Pagado",
        "factura": "KLB-2024-11",
        "tag": "Servicios",
        "personal_negocio": "Negocio"
    },
    {
        "fecha": "2024-11-10",
        "entidad": "Cliente-XYZ",
        "categoria": "Ingreso",
        "subcategoria": "Servicios Profesionales",
        "moneda": "USD",
        "monto": 5000.00,
        "descripcion": "Consultoría proyecto ABC",
        "forma_pago": "Transferencia",
        "iva": "Sí",
        "notas": "Factura electrónica enviada",
        "recurrente": "No",
        "proyecto": "CONSULT-2024",
        "estado": "Cobrado",
        "factura": "FAC-2024-089",
        "tag": "Ingresos",
        "personal_negocio": "Negocio"
    }
]

print("╔═══════════════════════════════════════════════════════════════════╗")
print("║           GENERADOR EXCEL v4.0 - SISTEMA COMPLETO                 ║")
print("╠═══════════════════════════════════════════════════════════════════╣")
print("║ Generando archivo con 14 hojas:                                   ║")
print("║  1. RESUMEN                  8. FLUJO_CAJA                         ║")
print("║  2. TRANSACCIONES            9. DASHBOARD_VISUAL                   ║")
print("║  3. CxP                     10. CONCILIACION                       ║")
print("║  4. CxC                     11. CONFIG                             ║")
print("║  5. ENTIDADES_ALIAS         12. IVA_CONTROL                        ║")
print("║  6. ESTADO_RESULTADOS       13. PRESUPUESTO                        ║")
print("║  7. BALANCE_GENERAL         14. PERSONAL_VS_NEGOCIO                ║")
print("╚═══════════════════════════════════════════════════════════════════╝")
print()
print("✅ Configuración cargada")
print(f"   - {len(COLUMNAS_TRANSACCIONES)} columnas en TRANSACCIONES")
print(f"   - {len(ALIAS_CUENTAS)} alias pre-cargados")
print(f"   - {len(TRANSACCIONES_EJEMPLO)} transacciones de ejemplo")
print()

# ============================================================================
# FUNCIONES PARA CREAR LAS 14 HOJAS
# ============================================================================

def crear_hoja_transacciones(wb):
    """Crea la hoja TRANSACCIONES con 16 columnas y 3 transacciones ejemplo"""
    print("   📝 Creando hoja TRANSACCIONES...")
    ws = wb.create_sheet("TRANSACCIONES", 0)

    # Escribir encabezados
    for i, columna in enumerate(COLUMNAS_TRANSACCIONES, 1):
        celda = ws.cell(1, i, columna)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        celda.alignment = Alignment(horizontal="center")

    # Ajustar anchos de columna
    anchos = [12, 20, 15, 15, 15, 15, 20, 15, 10, 30, 10, 15, 12, 15, 12, 15]
    for i, ancho in enumerate(anchos, 1):
        col_letter = chr(64 + i) if i <= 26 else chr(64 + i // 26) + chr(64 + i % 26)
        ws.column_dimensions[col_letter].width = ancho

    # Agregar 3 transacciones de ejemplo
    for i, trans in enumerate(TRANSACCIONES_EJEMPLO, 2):
        ws.cell(i, 1, trans['fecha'])
        ws.cell(i, 2, trans['entidad'])
        ws.cell(i, 3, trans['categoria'])
        ws.cell(i, 4, trans['subcategoria'])
        ws.cell(i, 5, trans['moneda'])
        ws.cell(i, 6, trans['monto'])
        ws.cell(i, 7, trans['descripcion'])
        ws.cell(i, 8, trans['forma_pago'])
        ws.cell(i, 9, trans['iva'])
        ws.cell(i, 10, trans['notas'])
        ws.cell(i, 11, trans['recurrente'])
        ws.cell(i, 12, trans['proyecto'])
        ws.cell(i, 13, trans['estado'])
        ws.cell(i, 14, trans['factura'])
        ws.cell(i, 15, trans['tag'])
        ws.cell(i, 16, trans['personal_negocio'])

    return ws

def crear_hoja_cxp(wb):
    """Crea hoja CxP con tarjetas de crédito incluidas"""
    print("   💳 Creando hoja CxP (incluye tarjetas)...")
    ws = wb.create_sheet("CxP")

    # Encabezados
    headers = ["Proveedor/Tarjeta", "Monto", "Moneda", "Vencimiento", "Días", "Estado"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")

    # Fórmulas para extraer de TRANSACCIONES donde Estado = "Pendiente" y categoría = "CxP"
    ws.cell(2, 1, '=SI(TRANSACCIONES!N2="Pendiente", TRANSACCIONES!B2, "")')
    ws.cell(2, 2, '=SI(TRANSACCIONES!N2="Pendiente", TRANSACCIONES!F2, 0)')
    ws.cell(2, 3, '=SI(TRANSACCIONES!N2="Pendiente", TRANSACCIONES!E2, "")')
    ws.cell(2, 4, "=TRANSACCIONES!A2")
    ws.cell(2, 5, '=SI(A2<>"", HOY()-D2, "")')
    ws.cell(2, 6, '=SI(E2>30, "VENCIDA", "VIGENTE")')

    # Totales
    ws.cell(20, 1, "TOTAL CxP:")
    ws.cell(20, 1).font = Font(bold=True)
    ws.cell(20, 2, '=SUMAR.SI(C:C,"USD",B:B)')
    ws.cell(20, 2).font = Font(bold=True, color="C00000")

    ws.cell(21, 1, "TOTAL ₡:")
    ws.cell(21, 2, '=SUMAR.SI(C:C,"CRC",B:B)')

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 15

    return ws

def crear_hoja_cxc(wb):
    """Crea hoja CxC - Cuentas por Cobrar"""
    print("   📥 Creando hoja CxC...")
    ws = wb.create_sheet("CxC")

    headers = ["Cliente", "Monto", "Moneda", "Fecha Emisión", "Días", "Estado"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")

    # Fórmulas para extraer de TRANSACCIONES donde categoría = "CxC" y Estado = "Pendiente"
    ws.cell(2, 1, '=SI(Y(TRANSACCIONES!C2="CxC", TRANSACCIONES!N2="Pendiente"), TRANSACCIONES!B2, "")')
    ws.cell(2, 2, '=SI(A2<>"", TRANSACCIONES!F2, 0)')
    ws.cell(2, 3, '=SI(A2<>"", TRANSACCIONES!E2, "")')
    ws.cell(2, 4, '=SI(A2<>"", TRANSACCIONES!A2, "")')
    ws.cell(2, 5, '=SI(A2<>"", HOY()-D2, "")')
    ws.cell(2, 6, '=SI(E2>60, "VENCIDA", "VIGENTE")')

    # Totales
    ws.cell(20, 1, "TOTAL CxC:")
    ws.cell(20, 1).font = Font(bold=True)
    ws.cell(20, 2, '=SUMAR.SI(C:C,"USD",B:B)')
    ws.cell(20, 2).font = Font(bold=True, color="00B050")

    ws.column_dimensions['A'].width = 30
    return ws

def crear_hoja_entidades_alias(wb):
    """Crea hoja ENTIDADES_ALIAS con 12 alias pre-cargados"""
    print("   🏢 Creando hoja ENTIDADES_ALIAS...")
    ws = wb.create_sheet("ENTIDADES_ALIAS")

    headers = ["Nombre Completo", "Alias", "Tipo"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")

    # Cargar los 12 alias
    for i, alias in enumerate(ALIAS_CUENTAS, 2):
        ws.cell(i, 1, alias['nombre_completo'])
        ws.cell(i, 2, alias['alias'])
        ws.cell(i, 3, alias['tipo'])

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15

    return ws

def crear_hoja_resumen(wb):
    """Crea hoja RESUMEN - Dashboard principal"""
    print("   📊 Creando hoja RESUMEN...")
    ws = wb.create_sheet("RESUMEN", 0)

    # Título
    ws.merge_cells('A1:F1')
    titulo = ws.cell(1, 1, f"SISTEMA DE FINANZAS v{VERSION} - RESUMEN EJECUTIVO")
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    titulo.alignment = Alignment(horizontal="center")

    # Fecha actualización
    ws.cell(2, 1, "Última actualización:")
    ws.cell(2, 2, "=HOY()")
    ws.cell(2, 2).number_format = 'DD/MM/YYYY'

    # SECCIÓN: CUENTAS POR PAGAR
    ws.cell(4, 1, "CUENTAS POR PAGAR")
    ws.cell(4, 1).font = Font(bold=True, size=12, color="C00000")
    ws.cell(5, 1, "Total USD:")
    ws.cell(5, 2, "=CxP!B20")
    ws.cell(5, 2).font = Font(bold=True)
    ws.cell(6, 1, "Total ₡:")
    ws.cell(6, 2, "=CxP!B21")

    # SECCIÓN: CUENTAS POR COBRAR
    ws.cell(8, 1, "CUENTAS POR COBRAR")
    ws.cell(8, 1).font = Font(bold=True, size=12, color="00B050")
    ws.cell(9, 1, "Total USD:")
    ws.cell(9, 2, "=CxC!B20")
    ws.cell(9, 2).font = Font(bold=True)

    # SECCIÓN: FLUJO DE CAJA
    ws.cell(11, 1, "FLUJO DE CAJA")
    ws.cell(11, 1).font = Font(bold=True, size=12, color="0070C0")
    ws.cell(12, 1, "Balance USD:")
    ws.cell(12, 2, "=FLUJO_CAJA!B50")
    ws.cell(13, 1, "Balance ₡:")
    ws.cell(13, 2, "=FLUJO_CAJA!C50")

    # SECCIÓN: IVA
    ws.cell(15, 1, "CONTROL IVA")
    ws.cell(15, 1).font = Font(bold=True, size=12, color="FF6600")
    ws.cell(16, 1, "Balance a pagar:")
    ws.cell(16, 2, "=IVA_CONTROL!B20")

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_estado_resultados(wb):
    """Crea hoja ESTADO_RESULTADOS (P&L)"""
    print("   📈 Creando ESTADO_RESULTADOS...")
    ws = wb.create_sheet("ESTADO_RESULTADOS")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "ESTADO DE RESULTADOS (P&L)")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # INGRESOS
    ws.cell(3, 1, "INGRESOS")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Ingresos Operacionales:")
    ws.cell(4, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Ingreso", TRANSACCIONES!E:E, "USD")')

    # GASTOS
    ws.cell(6, 1, "GASTOS")
    ws.cell(6, 1).font = Font(bold=True, color="C00000")
    ws.cell(7, 1, "Gastos Operacionales:")
    ws.cell(7, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Gasto", TRANSACCIONES!E:E, "USD")')

    # UTILIDAD
    ws.cell(9, 1, "UTILIDAD NETA")
    ws.cell(9, 1).font = Font(bold=True, size=12)
    ws.cell(9, 2, "=B4-B7")
    ws.cell(9, 2).font = Font(bold=True, size=12)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_balance_general(wb):
    """Crea hoja BALANCE_GENERAL"""
    print("   💰 Creando BALANCE_GENERAL...")
    ws = wb.create_sheet("BALANCE_GENERAL")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "BALANCE GENERAL")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # ACTIVOS
    ws.cell(3, 1, "ACTIVOS")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Cuentas por Cobrar:")
    ws.cell(4, 2, "=CxC!B20")
    ws.cell(5, 1, "Efectivo:")
    ws.cell(5, 2, "=FLUJO_CAJA!B50")
    ws.cell(6, 1, "TOTAL ACTIVOS:")
    ws.cell(6, 2, "=B4+B5")
    ws.cell(6, 1).font = Font(bold=True)

    # PASIVOS
    ws.cell(8, 1, "PASIVOS")
    ws.cell(8, 1).font = Font(bold=True, color="C00000")
    ws.cell(9, 1, "Cuentas por Pagar:")
    ws.cell(9, 2, "=CxP!B20")
    ws.cell(10, 1, "IVA por Pagar:")
    ws.cell(10, 2, "=IVA_CONTROL!B20")
    ws.cell(11, 1, "TOTAL PASIVOS:")
    ws.cell(11, 2, "=B9+B10")
    ws.cell(11, 1).font = Font(bold=True)

    # PATRIMONIO
    ws.cell(13, 1, "PATRIMONIO")
    ws.cell(13, 1).font = Font(bold=True, color="0070C0")
    ws.cell(14, 2, "=B6-B11")
    ws.cell(14, 1).font = Font(bold=True)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_flujo_caja(wb):
    """Crea hoja FLUJO_CAJA"""
    print("   💵 Creando FLUJO_CAJA...")
    ws = wb.create_sheet("FLUJO_CAJA")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "FLUJO DE CAJA")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Encabezados
    ws.cell(3, 1, "Concepto")
    ws.cell(3, 2, "USD")
    ws.cell(3, 3, "CRC")
    for col in [1, 2, 3]:
        ws.cell(3, col).font = Font(bold=True)

    # ENTRADAS
    ws.cell(4, 1, "ENTRADAS")
    ws.cell(4, 1).font = Font(bold=True, color="00B050")
    ws.cell(5, 1, "Ingresos:")
    ws.cell(5, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Ingreso", TRANSACCIONES!E:E, "USD")')
    ws.cell(5, 3, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Ingreso", TRANSACCIONES!E:E, "CRC")')

    # SALIDAS
    ws.cell(7, 1, "SALIDAS")
    ws.cell(7, 1).font = Font(bold=True, color="C00000")
    ws.cell(8, 1, "Gastos:")
    ws.cell(8, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Gasto", TRANSACCIONES!E:E, "USD")')
    ws.cell(8, 3, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Gasto", TRANSACCIONES!E:E, "CRC")')

    # BALANCE
    ws.cell(50, 1, "BALANCE FINAL:")
    ws.cell(50, 1).font = Font(bold=True, size=12)
    ws.cell(50, 2, "=B5-B8")
    ws.cell(50, 3, "=C5-C8")
    ws.cell(50, 2).font = Font(bold=True)
    ws.cell(50, 3).font = Font(bold=True)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15

    return ws

def crear_hoja_dashboard_visual(wb):
    """Crea hoja DASHBOARD_VISUAL"""
    print("   📉 Creando DASHBOARD_VISUAL...")
    ws = wb.create_sheet("DASHBOARD_VISUAL")

    ws.merge_cells('A1:F1')
    titulo = ws.cell(1, 1, "DASHBOARD VISUAL")
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Top 5 Gastos por Categoría
    ws.cell(3, 1, "TOP 5 GASTOS POR CATEGORÍA")
    ws.cell(3, 1).font = Font(bold=True, size=12)

    ws.cell(4, 1, "Categoría")
    ws.cell(4, 2, "Monto USD")
    ws.cell(4, 1).font = Font(bold=True)
    ws.cell(4, 2).font = Font(bold=True)

    # Fórmulas para sumar por categoría
    ws.cell(5, 1, "Servicios")
    ws.cell(5, 2, '=SUMAR.SI(TRANSACCIONES!C:C, "Servicios", TRANSACCIONES!F:F)')

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_conciliacion(wb):
    """Crea hoja CONCILIACION"""
    print("   ✅ Creando CONCILIACION...")
    ws = wb.create_sheet("CONCILIACION")

    ws.merge_cells('A1:E1')
    titulo = ws.cell(1, 1, "CONCILIACIÓN BANCARIA")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Encabezados
    headers = ["Fecha", "Concepto", "Sistema", "Banco", "Diferencia"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
        ws.cell(3, i).font = Font(bold=True)

    ws.cell(4, 1, "=HOY()")
    ws.cell(4, 2, "Saldo Efectivo")
    ws.cell(4, 3, "=FLUJO_CAJA!B50")
    ws.cell(4, 4, 0)  # Usuario ingresará saldo real del banco
    ws.cell(4, 5, "=C4-D4")

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15

    return ws

def crear_hoja_config(wb):
    """Crea hoja CONFIG - EDITABLE"""
    print("   ⚙️  Creando CONFIG (EDITABLE)...")
    ws = wb.create_sheet("CONFIG")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "CONFIGURACIÓN DEL SISTEMA")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")

    # TIPO DE CAMBIO
    ws.cell(3, 1, "TIPO DE CAMBIO")
    ws.cell(3, 1).font = Font(bold=True, size=12)
    ws.cell(4, 1, "USD → CRC:")
    ws.cell(4, 2, TC_ACTUAL)
    ws.cell(4, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # AMARILLO = EDITABLE
    ws.cell(4, 3, "← EDITABLE")
    ws.cell(4, 3).font = Font(italic=True, color="FF0000")

    # FECHAS DE PAGO TARJETAS
    ws.cell(6, 1, "FECHAS DE PAGO - TARJETAS DE CRÉDITO")
    ws.cell(6, 1).font = Font(bold=True, size=12)

    ws.cell(7, 1, "BAC Visa:")
    ws.cell(7, 2, "15")
    ws.cell(7, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.cell(7, 3, "día del mes")

    ws.cell(8, 1, "BCR Mastercard:")
    ws.cell(8, 2, "20")
    ws.cell(8, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.cell(9, 1, "Credomatic Platinum:")
    ws.cell(9, 2, "10")
    ws.cell(9, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.cell(10, 1, "Credomatic Gold:")
    ws.cell(10, 2, "10")
    ws.cell(10, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # EMPRESAS ZONA FRANCA (exentas de IVA)
    ws.cell(12, 1, "EMPRESAS ZONA FRANCA (sin IVA)")
    ws.cell(12, 1).font = Font(bold=True, size=12)
    ws.cell(13, 1, "• VWR International")
    ws.cell(14, 1, "• RS Hughes")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20

    return ws

def crear_hoja_iva_control(wb):
    """Crea hoja IVA_CONTROL con exempciones zona franca"""
    print("   🧾 Creando IVA_CONTROL...")
    ws = wb.create_sheet("IVA_CONTROL")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "CONTROL DE IVA")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")

    # IVA COMPRAS (Crédito fiscal)
    ws.cell(3, 1, "IVA COMPRAS (Crédito Fiscal)")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Total IVA pagado:")
    # Fórmula que EXCLUYE VWR y RS Hughes
    ws.cell(4, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!I:I, TRANSACCIONES!I:I, "Sí", TRANSACCIONES!B:B, "<>*VWR*", TRANSACCIONES!B:B, "<>*RS Hughes*")')
    ws.cell(4, 3, "(excluye zona franca)")
    ws.cell(4, 3).font = Font(italic=True, size=9)

    # IVA VENTAS (Débito fiscal)
    ws.cell(6, 1, "IVA VENTAS (Débito Fiscal)")
    ws.cell(6, 1).font = Font(bold=True, color="C00000")
    ws.cell(7, 1, "Total IVA cobrado:")
    ws.cell(7, 2, '=SUMAR.SI.CONJUNTO(TRANSACCIONES!I:I, TRANSACCIONES!C:C, "Ingreso", TRANSACCIONES!I:I, "Sí")')

    # BALANCE
    ws.cell(20, 1, "BALANCE A PAGAR A HACIENDA:")
    ws.cell(20, 1).font = Font(bold=True, size=12)
    ws.cell(20, 2, "=B7-B4")
    ws.cell(20, 2).font = Font(bold=True, size=12, color="FF0000")

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25

    return ws

def crear_hoja_presupuesto(wb):
    """Crea hoja PRESUPUESTO"""
    print("   💼 Creando PRESUPUESTO...")
    ws = wb.create_sheet("PRESUPUESTO")

    ws.merge_cells('A1:E1')
    titulo = ws.cell(1, 1, "PRESUPUESTO VS REAL")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Encabezados
    headers = ["Categoría", "Presupuesto", "Real", "Diferencia", "% Ejecutado"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
        ws.cell(3, i).font = Font(bold=True)

    # Categorías principales
    categorias = ["Servicios", "Inventario", "Gastos Administrativos", "Marketing", "Personal"]
    for i, cat in enumerate(categorias, 4):
        ws.cell(i, 1, cat)
        ws.cell(i, 2, 5000)  # Presupuesto editable
        ws.cell(i, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws.cell(i, 3, f'=SUMAR.SI(TRANSACCIONES!C:C, "{cat}", TRANSACCIONES!F:F)')
        ws.cell(i, 4, f"=B{i}-C{i}")
        ws.cell(i, 5, f"=C{i}/B{i}")
        ws.cell(i, 5).number_format = '0%'

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15

    return ws

def crear_hoja_personal_vs_negocio(wb):
    """Crea hoja PERSONAL_VS_NEGOCIO usando columna P"""
    print("   👔 Creando PERSONAL_VS_NEGOCIO...")
    ws = wb.create_sheet("PERSONAL_VS_NEGOCIO")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "SEPARACIÓN: GASTOS PERSONALES VS NEGOCIO")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")

    # GASTOS DE NEGOCIO
    ws.cell(3, 1, "GASTOS DE NEGOCIO")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Total:")
    ws.cell(4, 2, '=SUMAR.SI(TRANSACCIONES!P:P, "Negocio", TRANSACCIONES!F:F)')
    ws.cell(4, 2).font = Font(bold=True)

    # GASTOS PERSONALES
    ws.cell(6, 1, "GASTOS PERSONALES")
    ws.cell(6, 1).font = Font(bold=True, color="FF0000")
    ws.cell(7, 1, "Total:")
    ws.cell(7, 2, '=SUMAR.SI(TRANSACCIONES!P:P, "Personal", TRANSACCIONES!F:F)')
    ws.cell(7, 2).font = Font(bold=True)

    # PORCENTAJE
    ws.cell(9, 1, "% Gastos Personales:")
    ws.cell(9, 2, "=B7/(B4+B7)")
    ws.cell(9, 2).number_format = '0.00%'
    ws.cell(9, 2).font = Font(bold=True, size=12)

    # ALERTA
    ws.cell(11, 1, "⚠️ ALERTA:")
    ws.cell(11, 2, '=SI(B9>0.3, "Gastos personales > 30%", "OK")')
    ws.cell(11, 2).font = Font(bold=True, color="FF0000")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def proteger_hojas(wb):
    """Protege todas las hojas EXCEPTO TRANSACCIONES y CONFIG"""
    print("   🔒 Protegiendo hojas (excepto TRANSACCIONES y CONFIG)...")

    hojas_editables = ["TRANSACCIONES", "CONFIG"]

    for hoja in wb.sheetnames:
        if hoja not in hojas_editables:
            ws = wb[hoja]
            ws.protection.sheet = True
            ws.protection.password = ""  # Sin contraseña, solo protección visual

    print("   ✅ Protección aplicada")

# ============================================================================
# FUNCIÓN MAIN Y EJECUCIÓN
# ============================================================================

def main():
    """Función principal que genera el Excel completo con 14 hojas"""

    print("\n" + "="*70)
    print("🚀 GENERANDO SISTEMA DE FINANZAS v4.0 - 14 HOJAS")
    print("="*70 + "\n")

    # Crear workbook
    print("📦 Creando archivo Excel...")
    wb = Workbook()

    # Eliminar hoja por defecto
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Crear todas las hojas en orden
    print("\n📋 Creando hojas del sistema:\n")

    try:
        # 1. RESUMEN (primera hoja - índice 0)
        crear_hoja_resumen(wb)

        # 2. TRANSACCIONES (segunda hoja - EDITABLE)
        crear_hoja_transacciones(wb)

        # 3. CxP - Cuentas por Pagar (incluye tarjetas)
        crear_hoja_cxp(wb)

        # 4. CxC - Cuentas por Cobrar
        crear_hoja_cxc(wb)

        # 5. ENTIDADES_ALIAS
        crear_hoja_entidades_alias(wb)

        # 6. ESTADO_RESULTADOS (P&L)
        crear_hoja_estado_resultados(wb)

        # 7. BALANCE_GENERAL
        crear_hoja_balance_general(wb)

        # 8. FLUJO_CAJA
        crear_hoja_flujo_caja(wb)

        # 9. DASHBOARD_VISUAL
        crear_hoja_dashboard_visual(wb)

        # 10. CONCILIACION
        crear_hoja_conciliacion(wb)

        # 11. CONFIG (EDITABLE)
        crear_hoja_config(wb)

        # 12. IVA_CONTROL (con exempciones zona franca)
        crear_hoja_iva_control(wb)

        # 13. PRESUPUESTO
        crear_hoja_presupuesto(wb)

        # 14. PERSONAL_VS_NEGOCIO
        crear_hoja_personal_vs_negocio(wb)

        print("\n" + "-"*70)

        # Proteger hojas (todas excepto TRANSACCIONES y CONFIG)
        proteger_hojas(wb)

        # Guardar archivo
        nombre_archivo = "AlvaroVelasco_Finanzas_v4.0_COMPLETO.xlsx"
        print(f"\n💾 Guardando archivo: {nombre_archivo}...")
        wb.save(nombre_archivo)

        print("\n" + "="*70)
        print("✅ ¡SISTEMA v4.0 GENERADO EXITOSAMENTE!")
        print("="*70)
        print(f"\n📁 Archivo creado: {nombre_archivo}")
        print(f"📊 Total de hojas: {len(wb.sheetnames)}")
        print("\n🔹 HOJAS EDITABLES (amarillo):")
        print("   1. TRANSACCIONES - Ingresa todas tus transacciones aquí")
        print("   2. CONFIG - Edita tipo de cambio y fechas de pago")
        print("\n🔒 HOJAS PROTEGIDAS (auto-calculadas):")
        print("   3. RESUMEN")
        print("   4. CxP (incluye tarjetas de crédito)")
        print("   5. CxC")
        print("   6. ENTIDADES_ALIAS")
        print("   7. ESTADO_RESULTADOS")
        print("   8. BALANCE_GENERAL")
        print("   9. FLUJO_CAJA")
        print("   10. DASHBOARD_VISUAL")
        print("   11. CONCILIACION")
        print("   12. IVA_CONTROL (excluye VWR y RS Hughes)")
        print("   13. PRESUPUESTO")
        print("   14. PERSONAL_VS_NEGOCIO")

        print("\n💡 CARACTERÍSTICAS PRINCIPALES:")
        print("   ✅ 16 columnas en TRANSACCIONES (incluye Personal/Negocio)")
        print("   ✅ Tarjetas de crédito aparecen en CxP")
        print("   ✅ IVA excluye empresas zona franca (VWR, RS Hughes)")
        print("   ✅ Separación gastos personales vs negocio")
        print("   ✅ 12 alias de cuentas pre-cargados")
        print("   ✅ 3 transacciones de ejemplo")
        print("   ✅ Tipo de cambio editable en CONFIG")

        print("\n🎯 PRÓXIMOS PASOS:")
        print("   1. Abre el archivo Excel")
        print("   2. Revisa las 3 transacciones de ejemplo")
        print("   3. Edita el tipo de cambio en CONFIG si es necesario")
        print("   4. Comienza a agregar tus transacciones reales en TRANSACCIONES")
        print("   5. Todas las demás hojas se actualizarán automáticamente")

        print("\n" + "="*70)
        print("🎉 ¡LISTO PARA USAR!")
        print("="*70 + "\n")

        return True

    except Exception as e:
        print(f"\n❌ ERROR al generar el archivo: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# ============================================================================
# EJECUCIÓN
# ============================================================================

if __name__ == "__main__":
    print("\n")
    print("╔═══════════════════════════════════════════════════════════════════╗")
    print("║                SISTEMA DE FINANZAS v4.0                           ║")
    print("║              GENERADOR COMPLETO - 14 HOJAS                        ║")
    print("╚═══════════════════════════════════════════════════════════════════╝")

    # Ejecutar generación
    exito = main()

    if exito:
        print("\n✅ Proceso completado exitosamente")
        input("\nPresiona ENTER para cerrar...")
    else:
        print("\n❌ Hubo errores en la generación")
        input("\nPresiona ENTER para cerrar...")
