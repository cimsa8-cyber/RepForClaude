# ============================================================================
# SISTEMA DE FINANZAS v4.0 - GENERADOR CORRECTO FINAL
# ============================================================================
# CORRECCIONES:
# - Fórmulas en INGLÉS (IF, SUM, SUMIF, etc.)
# - Dropdowns en columnas clave
# - Comentarios de ayuda en cada columna
# - Saldos iniciales REALES de tarjetas
# ============================================================================

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, date

print("✅ Libraries imported correctly\n")

# ============================================================================
# GLOBAL CONFIGURATION
# ============================================================================

VERSION = "4.0"
TC_ACTUAL = 540  # Exchange rate USD to CRC

# 16-column structure for TRANSACCIONES
COLUMNAS_TRANSACCIONES = [
    "Date", "Entity", "Category", "Subcategory", "Currency", "Amount",
    "Description", "Payment Method", "VAT", "Notes", "Recurring",
    "Project", "Status", "Invoice #", "Tag", "Personal/Business"
]

# Help comments for each column
AYUDA_COLUMNAS = {
    "Date": "Transaction date (YYYY-MM-DD)\nExample: 2024-11-14",
    "Entity": "Entity name or alias\nExample: BAC-USD, VWR, Client-XYZ",
    "Category": "Main category\nOptions: Income, Expense, Inventory, Services, CxP, CxC, Marketing, Personal",
    "Subcategory": "Detail of category\nExample: Lab Materials, Internet, Professional Services",
    "Currency": "Transaction currency\nOptions: USD or CRC",
    "Amount": "Transaction amount\nExample: 2500.00",
    "Description": "Brief description of transaction\nExample: Q4 2024 chemical reagents",
    "Payment Method": "How was it paid?\nOptions: Cash, Transfer, Credit Card, Check, SINPE",
    "VAT": "Does it include VAT?\nOptions: Yes or No",
    "Notes": "Additional notes\nExample: Free zone company - VAT exempt",
    "Recurring": "Is it recurring?\nOptions: Yes or No",
    "Project": "Associated project\nExample: LAB-2024, CONSULT-2024",
    "Status": "Transaction status\nOptions: Paid, Pending, Collected, Canceled",
    "Invoice #": "Invoice or reference number\nExample: VWR-98765",
    "Tag": "Tag for classification\nExample: Inventory, Services, Income",
    "Personal/Business": "Classification\nOptions: Personal or Business"
}

# 12 pre-loaded account aliases
ALIAS_CUENTAS = [
    {"nombre_completo": "BAC San José - USD Account", "alias": "BAC-USD", "tipo": "Bank"},
    {"nombre_completo": "Banco de Costa Rica - CRC Checking Account", "alias": "BCR-CRC", "tipo": "Bank"},
    {"nombre_completo": "BAC Visa Credit", "alias": "Visa-BAC", "tipo": "Credit Card"},
    {"nombre_completo": "BCR Mastercard", "alias": "MC-BCR", "tipo": "Credit Card"},
    {"nombre_completo": "Credomatic Platinum USD", "alias": "Credo-Platinum", "tipo": "Credit Card"},
    {"nombre_completo": "Credomatic Gold CRC", "alias": "Credo-Gold", "tipo": "Credit Card"},
    {"nombre_completo": "VWR International LLC", "alias": "VWR", "tipo": "Supplier"},
    {"nombre_completo": "RS Hughes Co. Inc.", "alias": "RS-Hughes", "tipo": "Supplier"},
    {"nombre_completo": "Cash - Petty Cash", "alias": "Cash", "tipo": "Cash"},
    {"nombre_completo": "Client XYZ Corp", "alias": "Client-XYZ", "tipo": "Client"},
    {"nombre_completo": "SINPE Mobile", "alias": "SINPE", "tipo": "Transfer"},
    {"nombre_completo": "PayPal Business", "alias": "PayPal", "tipo": "Digital"}
]

# REAL initial balances of credit cards (DEBT)
SALDOS_TARJETAS = [
    {
        "fecha": "2024-11-01",
        "entidad": "Visa-BAC",
        "categoria": "CxP",
        "subcategoria": "Credit Card Debt",
        "moneda": "CRC",
        "monto": 1831000,
        "descripcion": "BAC Visa initial balance - Payment day 15",
        "forma_pago": "Credit Card",
        "iva": "No",
        "notas": "Pending monthly payment",
        "recurrente": "Yes",
        "proyecto": "DEBT-2024",
        "estado": "Pending",
        "factura": "BAC-VISA-NOV2024",
        "tag": "Credit Card",
        "personal_negocio": "Business"
    },
    {
        "fecha": "2024-11-01",
        "entidad": "MC-BCR",
        "categoria": "CxP",
        "subcategoria": "Credit Card Debt",
        "moneda": "CRC",
        "monto": 1750000,
        "descripcion": "BCR Mastercard initial balance - Payment day 20",
        "forma_pago": "Credit Card",
        "iva": "No",
        "notas": "Pending monthly payment",
        "recurrente": "Yes",
        "proyecto": "DEBT-2024",
        "estado": "Pending",
        "factura": "BCR-MC-NOV2024",
        "tag": "Credit Card",
        "personal_negocio": "Business"
    },
    {
        "fecha": "2024-11-01",
        "entidad": "Credo-Platinum",
        "categoria": "CxP",
        "subcategoria": "Credit Card Debt",
        "moneda": "USD",
        "monto": 2500.00,
        "descripcion": "Credomatic Platinum initial balance - Payment day 10",
        "forma_pago": "Credit Card",
        "iva": "No",
        "notas": "Pending monthly payment",
        "recurrente": "Yes",
        "proyecto": "DEBT-2024",
        "estado": "Pending",
        "factura": "CREDO-PLAT-NOV2024",
        "tag": "Credit Card",
        "personal_negocio": "Business"
    },
    {
        "fecha": "2024-11-01",
        "entidad": "Credo-Gold",
        "categoria": "CxP",
        "subcategoria": "Credit Card Debt",
        "moneda": "CRC",
        "monto": 2800000,
        "descripcion": "Credomatic Gold initial balance - Payment day 10",
        "forma_pago": "Credit Card",
        "iva": "No",
        "notas": "Pending monthly payment",
        "recurrente": "Yes",
        "proyecto": "DEBT-2024",
        "estado": "Pending",
        "factura": "CREDO-GOLD-NOV2024",
        "tag": "Credit Card",
        "personal_negocio": "Business"
    }
]

print("╔═══════════════════════════════════════════════════════════════════╗")
print("║         EXCEL GENERATOR v4.0 - COMPLETE SYSTEM                    ║")
print("╠═══════════════════════════════════════════════════════════════════╣")
print("║ Generating file with 14 sheets:                                   ║")
print("║  1. SUMMARY                  8. CASH_FLOW                          ║")
print("║  2. TRANSACTIONS             9. VISUAL_DASHBOARD                   ║")
print("║  3. CxP                     10. RECONCILIATION                     ║")
print("║  4. CxC                     11. CONFIG                             ║")
print("║  5. ENTITIES_ALIAS          12. VAT_CONTROL                        ║")
print("║  6. INCOME_STATEMENT        13. BUDGET                             ║")
print("║  7. BALANCE_SHEET           14. PERSONAL_VS_BUSINESS               ║")
print("╚═══════════════════════════════════════════════════════════════════╝")
print()
print("✅ Configuration loaded")
print(f"   - {len(COLUMNAS_TRANSACCIONES)} columns in TRANSACTIONS")
print(f"   - {len(ALIAS_CUENTAS)} pre-loaded aliases")
print(f"   - {len(SALDOS_TARJETAS)} credit cards with REAL balances")
print()

# ============================================================================
# FUNCTIONS TO CREATE THE 14 SHEETS
# ============================================================================

def crear_hoja_transacciones(wb):
    """Creates TRANSACTIONS sheet with 16 columns, dropdowns, help comments, and 4 credit card balances"""
    print("   📝 Creating TRANSACTIONS sheet...")
    ws = wb.create_sheet("TRANSACTIONS", 0)

    # Write headers
    for i, columna in enumerate(COLUMNAS_TRANSACCIONES, 1):
        celda = ws.cell(1, i, columna)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        celda.alignment = Alignment(horizontal="center")

        # Add help comment
        if columna in AYUDA_COLUMNAS:
            celda.comment = Comment(AYUDA_COLUMNAS[columna], "System")

    # Adjust column widths
    anchos = [12, 20, 15, 15, 10, 15, 30, 15, 8, 30, 10, 15, 12, 20, 12, 18]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    # DATA VALIDATIONS (Dropdowns)

    # Category dropdown (Column C)
    dv_category = DataValidation(type="list", formula1='"Income,Expense,Inventory,Services,CxP,CxC,Marketing,Personal"', allow_blank=False)
    dv_category.error = 'Please select a valid category'
    dv_category.errorTitle = 'Invalid Category'
    ws.add_data_validation(dv_category)
    dv_category.add(f'C2:C1000')

    # Currency dropdown (Column E)
    dv_currency = DataValidation(type="list", formula1='"USD,CRC"', allow_blank=False)
    dv_currency.error = 'Please select USD or CRC'
    dv_currency.errorTitle = 'Invalid Currency'
    ws.add_data_validation(dv_currency)
    dv_currency.add(f'E2:E1000')

    # Payment Method dropdown (Column H)
    dv_payment = DataValidation(type="list", formula1='"Cash,Transfer,Credit Card,Check,SINPE,PayPal"', allow_blank=True)
    ws.add_data_validation(dv_payment)
    dv_payment.add(f'H2:H1000')

    # VAT dropdown (Column I)
    dv_vat = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    dv_vat.error = 'Please select Yes or No'
    dv_vat.errorTitle = 'Invalid VAT'
    ws.add_data_validation(dv_vat)
    dv_vat.add(f'I2:I1000')

    # Recurring dropdown (Column K)
    dv_recurring = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv_recurring)
    dv_recurring.add(f'K2:K1000')

    # Status dropdown (Column M)
    dv_status = DataValidation(type="list", formula1='"Paid,Pending,Collected,Canceled"', allow_blank=False)
    dv_status.error = 'Please select a valid status'
    dv_status.errorTitle = 'Invalid Status'
    ws.add_data_validation(dv_status)
    dv_status.add(f'M2:M1000')

    # Personal/Business dropdown (Column P)
    dv_personal = DataValidation(type="list", formula1='"Personal,Business"', allow_blank=False)
    dv_personal.error = 'Please select Personal or Business'
    dv_personal.errorTitle = 'Invalid Classification'
    ws.add_data_validation(dv_personal)
    dv_personal.add(f'P2:P1000')

    # Add 4 credit card REAL balances
    for i, trans in enumerate(SALDOS_TARJETAS, 2):
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

        # Highlight credit card rows in light yellow
        for col in range(1, 17):
            ws.cell(i, col).fill = PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")

    return ws

def crear_hoja_cxp(wb):
    """Creates CxP sheet with credit cards included - FORMULAS IN ENGLISH"""
    print("   💳 Creating CxP sheet (includes credit cards)...")
    ws = wb.create_sheet("CxP")

    # Headers
    headers = ["Supplier/Card", "Amount", "Currency", "Due Date", "Days", "Status"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")

    # ENGLISH formulas to extract from TRANSACTIONS where Status = "Pending"
    ws.cell(2, 1, '=IF(TRANSACTIONS!M2="Pending", TRANSACTIONS!B2, "")')
    ws.cell(2, 2, '=IF(TRANSACTIONS!M2="Pending", TRANSACTIONS!F2, 0)')
    ws.cell(2, 3, '=IF(TRANSACTIONS!M2="Pending", TRANSACTIONS!E2, "")')
    ws.cell(2, 4, "=TRANSACTIONS!A2")
    ws.cell(2, 5, '=IF(A2<>"", TODAY()-D2, "")')
    ws.cell(2, 6, '=IF(E2>30, "OVERDUE", "CURRENT")')

    # Copy formulas down to row 100
    for row in range(3, 101):
        ws.cell(row, 1, f'=IF(TRANSACTIONS!M{row}="Pending", TRANSACTIONS!B{row}, "")')
        ws.cell(row, 2, f'=IF(TRANSACTIONS!M{row}="Pending", TRANSACTIONS!F{row}, 0)')
        ws.cell(row, 3, f'=IF(TRANSACTIONS!M{row}="Pending", TRANSACTIONS!E{row}, "")')
        ws.cell(row, 4, f"=TRANSACTIONS!A{row}")
        ws.cell(row, 5, f'=IF(A{row}<>"", TODAY()-D{row}, "")')
        ws.cell(row, 6, f'=IF(E{row}>30, "OVERDUE", "CURRENT")')

    # Totals
    ws.cell(102, 1, "TOTAL CxP:")
    ws.cell(102, 1).font = Font(bold=True)
    ws.cell(102, 2, '=SUMIF(C:C,"USD",B:B)')
    ws.cell(102, 2).font = Font(bold=True, color="C00000")

    ws.cell(103, 1, "TOTAL ₡:")
    ws.cell(103, 2, '=SUMIF(C:C,"CRC",B:B)')
    ws.cell(103, 2).font = Font(bold=True, color="C00000")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 15

    return ws

def crear_hoja_cxc(wb):
    """Creates CxC - Accounts Receivable - FORMULAS IN ENGLISH"""
    print("   📥 Creating CxC sheet...")
    ws = wb.create_sheet("CxC")

    headers = ["Client", "Amount", "Currency", "Issue Date", "Days", "Status"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")

    # ENGLISH formulas to extract from TRANSACTIONS where category = "CxC" and Status = "Pending"
    ws.cell(2, 1, '=IF(AND(TRANSACTIONS!C2="CxC", TRANSACTIONS!M2="Pending"), TRANSACTIONS!B2, "")')
    ws.cell(2, 2, '=IF(A2<>"", TRANSACTIONS!F2, 0)')
    ws.cell(2, 3, '=IF(A2<>"", TRANSACTIONS!E2, "")')
    ws.cell(2, 4, '=IF(A2<>"", TRANSACTIONS!A2, "")')
    ws.cell(2, 5, '=IF(A2<>"", TODAY()-D2, "")')
    ws.cell(2, 6, '=IF(E2>60, "OVERDUE", "CURRENT")')

    # Copy formulas down
    for row in range(3, 101):
        ws.cell(row, 1, f'=IF(AND(TRANSACTIONS!C{row}="CxC", TRANSACTIONS!M{row}="Pending"), TRANSACTIONS!B{row}, "")')
        ws.cell(row, 2, f'=IF(A{row}<>"", TRANSACTIONS!F{row}, 0)')
        ws.cell(row, 3, f'=IF(A{row}<>"", TRANSACTIONS!E{row}, "")')
        ws.cell(row, 4, f'=IF(A{row}<>"", TRANSACTIONS!A{row}, "")')
        ws.cell(row, 5, f'=IF(A{row}<>"", TODAY()-D{row}, "")')
        ws.cell(row, 6, f'=IF(E{row}>60, "OVERDUE", "CURRENT")')

    # Totals
    ws.cell(102, 1, "TOTAL CxC:")
    ws.cell(102, 1).font = Font(bold=True)
    ws.cell(102, 2, '=SUMIF(C:C,"USD",B:B)')
    ws.cell(102, 2).font = Font(bold=True, color="00B050")

    ws.column_dimensions['A'].width = 30
    return ws

def crear_hoja_entidades_alias(wb):
    """Creates ENTITIES_ALIAS sheet with 12 pre-loaded aliases"""
    print("   🏢 Creating ENTITIES_ALIAS sheet...")
    ws = wb.create_sheet("ENTITIES_ALIAS")

    headers = ["Full Name", "Alias", "Type"]
    for i, h in enumerate(headers, 1):
        celda = ws.cell(1, i, h)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")

    # Load 12 aliases
    for i, alias in enumerate(ALIAS_CUENTAS, 2):
        ws.cell(i, 1, alias['nombre_completo'])
        ws.cell(i, 2, alias['alias'])
        ws.cell(i, 3, alias['tipo'])

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15

    return ws

def crear_hoja_resumen(wb):
    """Creates SUMMARY - Main Dashboard - FORMULAS IN ENGLISH"""
    print("   📊 Creating SUMMARY sheet...")
    ws = wb.create_sheet("SUMMARY", 0)

    # Title
    ws.merge_cells('A1:F1')
    titulo = ws.cell(1, 1, f"FINANCE SYSTEM v{VERSION} - EXECUTIVE SUMMARY")
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    titulo.alignment = Alignment(horizontal="center")

    # Update date
    ws.cell(2, 1, "Last update:")
    ws.cell(2, 2, "=TODAY()")
    ws.cell(2, 2).number_format = 'DD/MM/YYYY'

    # SECTION: ACCOUNTS PAYABLE
    ws.cell(4, 1, "ACCOUNTS PAYABLE")
    ws.cell(4, 1).font = Font(bold=True, size=12, color="C00000")
    ws.cell(5, 1, "Total USD:")
    ws.cell(5, 2, "=CxP!B102")
    ws.cell(5, 2).font = Font(bold=True)
    ws.cell(6, 1, "Total ₡:")
    ws.cell(6, 2, "=CxP!B103")
    ws.cell(6, 2).font = Font(bold=True)

    # SECTION: ACCOUNTS RECEIVABLE
    ws.cell(8, 1, "ACCOUNTS RECEIVABLE")
    ws.cell(8, 1).font = Font(bold=True, size=12, color="00B050")
    ws.cell(9, 1, "Total USD:")
    ws.cell(9, 2, "=CxC!B102")
    ws.cell(9, 2).font = Font(bold=True)

    # SECTION: CASH FLOW
    ws.cell(11, 1, "CASH FLOW")
    ws.cell(11, 1).font = Font(bold=True, size=12, color="0070C0")
    ws.cell(12, 1, "Balance USD:")
    ws.cell(12, 2, "=CASH_FLOW!B50")
    ws.cell(13, 1, "Balance ₡:")
    ws.cell(13, 2, "=CASH_FLOW!C50")

    # SECTION: VAT
    ws.cell(15, 1, "VAT CONTROL")
    ws.cell(15, 1).font = Font(bold=True, size=12, color="FF6600")
    ws.cell(16, 1, "Balance to pay:")
    ws.cell(16, 2, "=VAT_CONTROL!B20")

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_estado_resultados(wb):
    """Creates INCOME_STATEMENT (P&L) - FORMULAS IN ENGLISH"""
    print("   📈 Creating INCOME_STATEMENT...")
    ws = wb.create_sheet("INCOME_STATEMENT")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "INCOME STATEMENT (P&L)")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # INCOME
    ws.cell(3, 1, "INCOME")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Operating Income:")
    ws.cell(4, 2, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Income", TRANSACTIONS!E:E, "USD")')

    # EXPENSES
    ws.cell(6, 1, "EXPENSES")
    ws.cell(6, 1).font = Font(bold=True, color="C00000")
    ws.cell(7, 1, "Operating Expenses:")
    ws.cell(7, 2, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Expense", TRANSACTIONS!E:E, "USD")')

    # NET PROFIT
    ws.cell(9, 1, "NET PROFIT")
    ws.cell(9, 1).font = Font(bold=True, size=12)
    ws.cell(9, 2, "=B4-B7")
    ws.cell(9, 2).font = Font(bold=True, size=12)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_balance_general(wb):
    """Creates BALANCE_SHEET - FORMULAS IN ENGLISH"""
    print("   💰 Creating BALANCE_SHEET...")
    ws = wb.create_sheet("BALANCE_SHEET")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "BALANCE SHEET")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # ASSETS
    ws.cell(3, 1, "ASSETS")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Accounts Receivable:")
    ws.cell(4, 2, "=CxC!B102")
    ws.cell(5, 1, "Cash:")
    ws.cell(5, 2, "=CASH_FLOW!B50")
    ws.cell(6, 1, "TOTAL ASSETS:")
    ws.cell(6, 2, "=B4+B5")
    ws.cell(6, 1).font = Font(bold=True)

    # LIABILITIES
    ws.cell(8, 1, "LIABILITIES")
    ws.cell(8, 1).font = Font(bold=True, color="C00000")
    ws.cell(9, 1, "Accounts Payable:")
    ws.cell(9, 2, "=CxP!B102")
    ws.cell(10, 1, "VAT Payable:")
    ws.cell(10, 2, "=VAT_CONTROL!B20")
    ws.cell(11, 1, "TOTAL LIABILITIES:")
    ws.cell(11, 2, "=B9+B10")
    ws.cell(11, 1).font = Font(bold=True)

    # EQUITY
    ws.cell(13, 1, "EQUITY")
    ws.cell(13, 1).font = Font(bold=True, color="0070C0")
    ws.cell(14, 2, "=B6-B11")
    ws.cell(14, 1).font = Font(bold=True)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_flujo_caja(wb):
    """Creates CASH_FLOW - FORMULAS IN ENGLISH"""
    print("   💵 Creating CASH_FLOW...")
    ws = wb.create_sheet("CASH_FLOW")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "CASH FLOW")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Headers
    ws.cell(3, 1, "Concept")
    ws.cell(3, 2, "USD")
    ws.cell(3, 3, "CRC")
    for col in [1, 2, 3]:
        ws.cell(3, col).font = Font(bold=True)

    # INFLOWS
    ws.cell(4, 1, "INFLOWS")
    ws.cell(4, 1).font = Font(bold=True, color="00B050")
    ws.cell(5, 1, "Income:")
    ws.cell(5, 2, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Income", TRANSACTIONS!E:E, "USD")')
    ws.cell(5, 3, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Income", TRANSACTIONS!E:E, "CRC")')

    # OUTFLOWS
    ws.cell(7, 1, "OUTFLOWS")
    ws.cell(7, 1).font = Font(bold=True, color="C00000")
    ws.cell(8, 1, "Expenses:")
    ws.cell(8, 2, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Expense", TRANSACTIONS!E:E, "USD")')
    ws.cell(8, 3, '=SUMIFS(TRANSACTIONS!F:F, TRANSACTIONS!C:C, "Expense", TRANSACTIONS!E:E, "CRC")')

    # FINAL BALANCE
    ws.cell(50, 1, "FINAL BALANCE:")
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
    """Creates VISUAL_DASHBOARD"""
    print("   📉 Creating VISUAL_DASHBOARD...")
    ws = wb.create_sheet("VISUAL_DASHBOARD")

    ws.merge_cells('A1:F1')
    titulo = ws.cell(1, 1, "VISUAL DASHBOARD")
    titulo.font = Font(bold=True, size=16, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Top 5 Expenses by Category
    ws.cell(3, 1, "TOP 5 EXPENSES BY CATEGORY")
    ws.cell(3, 1).font = Font(bold=True, size=12)

    ws.cell(4, 1, "Category")
    ws.cell(4, 2, "Amount USD")
    ws.cell(4, 1).font = Font(bold=True)
    ws.cell(4, 2).font = Font(bold=True)

    # Formulas to sum by category
    ws.cell(5, 1, "Services")
    ws.cell(5, 2, '=SUMIF(TRANSACTIONS!C:C, "Services", TRANSACTIONS!F:F)')

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def crear_hoja_conciliacion(wb):
    """Creates RECONCILIATION - FORMULAS IN ENGLISH"""
    print("   ✅ Creating RECONCILIATION...")
    ws = wb.create_sheet("RECONCILIATION")

    ws.merge_cells('A1:E1')
    titulo = ws.cell(1, 1, "BANK RECONCILIATION")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Headers
    headers = ["Date", "Concept", "System", "Bank", "Difference"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
        ws.cell(3, i).font = Font(bold=True)

    ws.cell(4, 1, "=TODAY()")
    ws.cell(4, 2, "Cash Balance")
    ws.cell(4, 3, "=CASH_FLOW!B50")
    ws.cell(4, 4, 0)  # User will enter actual bank balance
    ws.cell(4, 5, "=C4-D4")

    # Mark cell D4 as editable (yellow)
    ws.cell(4, 4).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.cell(4, 4).comment = Comment("Enter actual bank balance here", "System")

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15

    return ws

def crear_hoja_config(wb):
    """Creates CONFIG - EDITABLE"""
    print("   ⚙️  Creating CONFIG (EDITABLE)...")
    ws = wb.create_sheet("CONFIG")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "SYSTEM CONFIGURATION")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")

    # EXCHANGE RATE
    ws.cell(3, 1, "EXCHANGE RATE")
    ws.cell(3, 1).font = Font(bold=True, size=12)
    ws.cell(4, 1, "USD → CRC:")
    ws.cell(4, 2, TC_ACTUAL)
    ws.cell(4, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # YELLOW = EDITABLE
    ws.cell(4, 3, "← EDITABLE")
    ws.cell(4, 3).font = Font(italic=True, color="FF0000")

    # CREDIT CARD PAYMENT DATES
    ws.cell(6, 1, "CREDIT CARD PAYMENT DATES")
    ws.cell(6, 1).font = Font(bold=True, size=12)

    ws.cell(7, 1, "BAC Visa:")
    ws.cell(7, 2, "15")
    ws.cell(7, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.cell(7, 3, "day of month")

    ws.cell(8, 1, "BCR Mastercard:")
    ws.cell(8, 2, "20")
    ws.cell(8, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.cell(9, 1, "Credomatic Platinum:")
    ws.cell(9, 2, "10")
    ws.cell(9, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.cell(10, 1, "Credomatic Gold:")
    ws.cell(10, 2, "10")
    ws.cell(10, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # FREE ZONE COMPANIES (VAT exempt)
    ws.cell(12, 1, "FREE ZONE COMPANIES (no VAT)")
    ws.cell(12, 1).font = Font(bold=True, size=12)
    ws.cell(13, 1, "• VWR International")
    ws.cell(14, 1, "• RS Hughes")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20

    return ws

def crear_hoja_iva_control(wb):
    """Creates VAT_CONTROL with free zone exemptions - FORMULAS IN ENGLISH"""
    print("   🧾 Creating VAT_CONTROL...")
    ws = wb.create_sheet("VAT_CONTROL")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "VAT CONTROL")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")

    # VAT PURCHASES (Tax credit)
    ws.cell(3, 1, "VAT PURCHASES (Tax Credit)")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Total VAT paid:")
    # Formula that EXCLUDES VWR and RS Hughes
    ws.cell(4, 2, '=SUMIFS(TRANSACTIONS!I:I, TRANSACTIONS!I:I, "Yes", TRANSACTIONS!B:B, "<>*VWR*", TRANSACTIONS!B:B, "<>*RS Hughes*")')
    ws.cell(4, 3, "(excludes free zone)")
    ws.cell(4, 3).font = Font(italic=True, size=9)

    # VAT SALES (Tax debit)
    ws.cell(6, 1, "VAT SALES (Tax Debit)")
    ws.cell(6, 1).font = Font(bold=True, color="C00000")
    ws.cell(7, 1, "Total VAT collected:")
    ws.cell(7, 2, '=SUMIFS(TRANSACTIONS!I:I, TRANSACTIONS!C:C, "Income", TRANSACTIONS!I:I, "Yes")')

    # BALANCE
    ws.cell(20, 1, "BALANCE TO PAY TO TAX AUTHORITY:")
    ws.cell(20, 1).font = Font(bold=True, size=12)
    ws.cell(20, 2, "=B7-B4")
    ws.cell(20, 2).font = Font(bold=True, size=12, color="FF0000")

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25

    return ws

def crear_hoja_presupuesto(wb):
    """Creates BUDGET - FORMULAS IN ENGLISH"""
    print("   💼 Creating BUDGET...")
    ws = wb.create_sheet("BUDGET")

    ws.merge_cells('A1:E1')
    titulo = ws.cell(1, 1, "BUDGET VS ACTUAL")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    # Headers
    headers = ["Category", "Budget", "Actual", "Difference", "% Executed"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
        ws.cell(3, i).font = Font(bold=True)

    # Main categories
    categorias = ["Services", "Inventory", "Administrative Expenses", "Marketing", "Personal"]
    for i, cat in enumerate(categorias, 4):
        ws.cell(i, 1, cat)
        ws.cell(i, 2, 5000)  # Editable budget
        ws.cell(i, 2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws.cell(i, 3, f'=SUMIF(TRANSACTIONS!C:C, "{cat}", TRANSACTIONS!F:F)')
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
    """Creates PERSONAL_VS_BUSINESS using column P - FORMULAS IN ENGLISH"""
    print("   👔 Creating PERSONAL_VS_BUSINESS...")
    ws = wb.create_sheet("PERSONAL_VS_BUSINESS")

    ws.merge_cells('A1:D1')
    titulo = ws.cell(1, 1, "SEPARATION: PERSONAL VS BUSINESS EXPENSES")
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")

    # BUSINESS EXPENSES
    ws.cell(3, 1, "BUSINESS EXPENSES")
    ws.cell(3, 1).font = Font(bold=True, color="00B050")
    ws.cell(4, 1, "Total:")
    ws.cell(4, 2, '=SUMIF(TRANSACTIONS!P:P, "Business", TRANSACTIONS!F:F)')
    ws.cell(4, 2).font = Font(bold=True)

    # PERSONAL EXPENSES
    ws.cell(6, 1, "PERSONAL EXPENSES")
    ws.cell(6, 1).font = Font(bold=True, color="FF0000")
    ws.cell(7, 1, "Total:")
    ws.cell(7, 2, '=SUMIF(TRANSACTIONS!P:P, "Personal", TRANSACTIONS!F:F)')
    ws.cell(7, 2).font = Font(bold=True)

    # PERCENTAGE
    ws.cell(9, 1, "% Personal Expenses:")
    ws.cell(9, 2, "=B7/(B4+B7)")
    ws.cell(9, 2).number_format = '0.00%'
    ws.cell(9, 2).font = Font(bold=True, size=12)

    # ALERT
    ws.cell(11, 1, "⚠️ ALERT:")
    ws.cell(11, 2, '=IF(B9>0.3, "Personal expenses > 30%", "OK")')
    ws.cell(11, 2).font = Font(bold=True, color="FF0000")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    return ws

def proteger_hojas(wb):
    """Protects all sheets EXCEPT TRANSACTIONS and CONFIG"""
    print("   🔒 Protecting sheets (except TRANSACTIONS and CONFIG)...")

    hojas_editables = ["TRANSACTIONS", "CONFIG"]

    for hoja in wb.sheetnames:
        if hoja not in hojas_editables:
            ws = wb[hoja]
            ws.protection.sheet = True
            ws.protection.password = ""  # No password, just visual protection

    print("   ✅ Protection applied")

# ============================================================================
# MAIN FUNCTION AND EXECUTION
# ============================================================================

def main():
    """Main function that generates complete Excel with 14 sheets"""

    print("\n" + "="*70)
    print("🚀 GENERATING FINANCE SYSTEM v4.0 - 14 SHEETS")
    print("="*70 + "\n")

    # Create workbook
    print("📦 Creating Excel file...")
    wb = Workbook()

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Create all sheets in order
    print("\n📋 Creating system sheets:\n")

    try:
        # 1. SUMMARY (first sheet - index 0)
        crear_hoja_resumen(wb)

        # 2. TRANSACTIONS (second sheet - EDITABLE)
        crear_hoja_transacciones(wb)

        # 3. CxP - Accounts Payable (includes credit cards)
        crear_hoja_cxp(wb)

        # 4. CxC - Accounts Receivable
        crear_hoja_cxc(wb)

        # 5. ENTITIES_ALIAS
        crear_hoja_entidades_alias(wb)

        # 6. INCOME_STATEMENT (P&L)
        crear_hoja_estado_resultados(wb)

        # 7. BALANCE_SHEET
        crear_hoja_balance_general(wb)

        # 8. CASH_FLOW
        crear_hoja_flujo_caja(wb)

        # 9. VISUAL_DASHBOARD
        crear_hoja_dashboard_visual(wb)

        # 10. RECONCILIATION
        crear_hoja_conciliacion(wb)

        # 11. CONFIG (EDITABLE)
        crear_hoja_config(wb)

        # 12. VAT_CONTROL (with free zone exemptions)
        crear_hoja_iva_control(wb)

        # 13. BUDGET
        crear_hoja_presupuesto(wb)

        # 14. PERSONAL_VS_BUSINESS
        crear_hoja_personal_vs_negocio(wb)

        print("\n" + "-"*70)

        # Protect sheets (all except TRANSACTIONS and CONFIG)
        proteger_hojas(wb)

        # Save file
        nombre_archivo = "AlvaroVelasco_Finanzas_v4.0_FINAL_CORRECTED.xlsx"
        print(f"\n💾 Saving file: {nombre_archivo}...")
        wb.save(nombre_archivo)

        print("\n" + "="*70)
        print("✅ SYSTEM v4.0 GENERATED SUCCESSFULLY!")
        print("="*70)
        print(f"\n📁 File created: {nombre_archivo}")
        print(f"📊 Total sheets: {len(wb.sheetnames)}")
        print("\n🔹 EDITABLE SHEETS (yellow cells):")
        print("   1. TRANSACTIONS - Enter all your transactions here")
        print("   2. CONFIG - Edit exchange rate and payment dates")
        print("\n🔒 PROTECTED SHEETS (auto-calculated):")
        print("   3. SUMMARY")
        print("   4. CxP (includes credit cards)")
        print("   5. CxC")
        print("   6. ENTITIES_ALIAS")
        print("   7. INCOME_STATEMENT")
        print("   8. BALANCE_SHEET")
        print("   9. CASH_FLOW")
        print("   10. VISUAL_DASHBOARD")
        print("   11. RECONCILIATION")
        print("   12. VAT_CONTROL (excludes VWR and RS Hughes)")
        print("   13. BUDGET")
        print("   14. PERSONAL_VS_BUSINESS")

        print("\n💡 MAIN FEATURES:")
        print("   ✅ 16 columns in TRANSACTIONS (includes Personal/Business)")
        print("   ✅ DROPDOWNS in key columns (Category, Currency, Status, etc.)")
        print("   ✅ HELP COMMENTS in each column header")
        print("   ✅ Credit cards appear in CxP with REAL balances")
        print("   ✅ VAT excludes free zone companies (VWR, RS Hughes)")
        print("   ✅ Personal vs business expense separation")
        print("   ✅ 12 pre-loaded account aliases")
        print("   ✅ 4 credit cards with REAL initial balances:")
        print("       • BAC Visa: ₡1,831,000")
        print("       • BCR Mastercard: ₡1,750,000")
        print("       • Credomatic Platinum: $2,500 USD")
        print("       • Credomatic Gold: ₡2,800,000")
        print("   ✅ Editable exchange rate in CONFIG")
        print("   ✅ ALL FORMULAS IN ENGLISH (IF, SUM, SUMIF, SUMIFS, TODAY)")

        print("\n🎯 NEXT STEPS:")
        print("   1. Open the Excel file")
        print("   2. Review the 4 credit card balances in TRANSACTIONS")
        print("   3. Edit exchange rate in CONFIG if needed")
        print("   4. Start adding your real transactions in TRANSACTIONS")
        print("   5. All other sheets will update automatically")
        print("   6. Test dropdowns in each column")
        print("   7. Hover over column headers to see help comments")

        print("\n" + "="*70)
        print("🎉 READY TO USE!")
        print("="*70 + "\n")

        return True

    except Exception as e:
        print(f"\n❌ ERROR generating file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# ============================================================================
# EXECUTION
# ============================================================================

if __name__ == "__main__":
    print("\n")
    print("╔═══════════════════════════════════════════════════════════════════╗")
    print("║                FINANCE SYSTEM v4.0 CORRECTED                      ║")
    print("║              COMPLETE GENERATOR - 14 SHEETS                       ║")
    print("║                                                                   ║")
    print("║  ✅ English formulas (IF, SUM, SUMIF, SUMIFS, TODAY)            ║")
    print("║  ✅ Dropdowns in key columns                                     ║")
    print("║  ✅ Help comments in headers                                     ║")
    print("║  ✅ REAL credit card balances loaded                            ║")
    print("╚═══════════════════════════════════════════════════════════════════╝")

    # Execute generation
    exito = main()

    if exito:
        print("\n✅ Process completed successfully")
        input("\nPress ENTER to close...")
    else:
        print("\n❌ Errors occurred during generation")
        input("\nPress ENTER to close...")
