#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════════════
GENERADOR SISTEMA FINANCIERO ERP v5.0 - XLWINGS (100% Compatible)
═══════════════════════════════════════════════════════════════════════════

Proyecto: AlvaroVelasco_Finanzas_v5.0.xlsx
Autor: Alvaro Velasco | Net SRL
Versión: 5.0 Final - Con xlwings para garantizar compatibilidad
Excel: Office 365

DIFERENCIAS vs openpyxl:
- ✅ Usa Excel REAL (no genera XML directamente)
- ✅ Fórmulas FILTER funcionan 100%
- ✅ Dropdowns nativos de Excel
- ✅ 22 hojas (21 originales + ALIAS)

REQUISITOS:
- pip install xlwings
- Excel instalado en Windows/Mac

EJECUCIÓN:
- python generar_finanzas_ERP_v50_xlwings.py
- El script abre Excel, genera el archivo, lo guarda y cierra
"""

import xlwings as xw
from datetime import datetime, timedelta
import sys

# Configuración
TC_ACTUAL = 540
ARCHIVO_SALIDA = "AlvaroVelasco_Finanzas_v5.0_COMPLETO.xlsx"

# Colores
COLOR_AZUL = (31, 71, 136)       # #1F4788
COLOR_ROJO = (255, 82, 82)       # #FF5252
COLOR_VERDE = (76, 175, 80)      # #4CAF50
COLOR_NARANJA = (255, 152, 0)    # #FF9800
COLOR_MORADO = (156, 39, 176)    # #9C27B0

def aplicar_estilo_header(ws, fila, col_inicio, col_fin, texto, color_rgb):
    """Aplica estilo a un header fusionado."""
    rango = ws.range((fila, col_inicio), (fila, col_fin))
    rango.merge()
    rango.value = texto
    rango.color = color_rgb
    rango.api.Font.Bold = True
    rango.api.Font.Color = 0xFFFFFF  # Blanco
    rango.api.Font.Size = 14
    rango.api.HorizontalAlignment = -4108  # xlCenter
    rango.row_height = 30

def crear_dropdown(ws, rango_celdas, valores):
    """Crea un dropdown de validación de datos."""
    ws.range(rango_celdas).api.Validation.Delete()
    ws.range(rango_celdas).api.Validation.Add(
        Type=3,  # xlValidateList
        AlertStyle=1,  # xlValidAlertStop
        Operator=1,
        Formula1=valores
    )

print("\n" + "="*70)
print("  GENERADOR ERP v5.0 - XLWINGS (100% COMPATIBLE)")
print("  22 Hojas | Excel Real | FILTER garantizado | Audit-Ready")
print("="*70 + "\n")

print("⚠️  IMPORTANTE: Este script abrirá Excel. No lo cierres manualmente.\n")
print("Iniciando generación...\n")

try:
    # Crear instancia de Excel
    print("📊 Iniciando Excel...")
    app = xw.App(visible=False)  # visible=True si quieres ver el proceso
    app.display_alerts = False
    wb = app.books.add()

    print("✓ Excel iniciado correctamente\n")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 1: RESUMEN
    # ═════════════════════════════════════════════════════════════════
    print("📋 Generando RESUMEN...")
    ws = wb.sheets.add("RESUMEN")

    aplicar_estilo_header(ws, 1, 1, 8, "DASHBOARD FINANCIERO - AlvaroVelasco Net SRL", COLOR_AZUL)
    ws.range("A2").value = f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    # Sección CxP
    aplicar_estilo_header(ws, 4, 1, 4, "CUENTAS POR PAGAR (CxP)", COLOR_ROJO)
    ws.range("A5").value = "Total CxP USD:"
    ws.range("B5").formula = '=SUMIF(CxP!D:D,"USD",CxP!H:H)'
    ws.range("A6").value = "Total CxP CRC:"
    ws.range("B6").formula = '=SUMIF(CxP!D:D,"CRC",CxP!C:C)'

    # Sección CxC
    aplicar_estilo_header(ws, 8, 1, 4, "CUENTAS POR COBRAR (CxC)", COLOR_VERDE)
    ws.range("A9").value = "Total CxC USD:"
    ws.range("B9").formula = '=SUMIF(CxC!D:D,"USD",CxC!H:H)'

    # Ajustar anchos
    ws.range("A:A").column_width = 25
    ws.range("B:B").column_width = 18

    print("✓ Hoja RESUMEN creada")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 2: TRANSACCIONES (Single Source of Truth)
    # ═════════════════════════════════════════════════════════════════
    print("📋 Generando TRANSACCIONES...")
    ws = wb.sheets.add("TRANSACCIONES")

    aplicar_estilo_header(ws, 1, 1, 23, "TRANSACCIONES - Single Source of Truth", COLOR_AZUL)

    # Headers (23 columnas)
    headers = [
        "Fecha", "Entidad", "Categoría", "Subcategoría", "Moneda", "Monto",
        "Descripción", "Forma Pago", "IVA", "Notas", "Zona Franca", "Proyecto",
        "Estado", "Factura #", "Tag", "Personal/Negocio", "Proveedor",
        "Días Trans", "Equiv USD", "Fecha Venc", "Prioridad CxP", "Prioridad CxC", "✓ Validación"
    ]

    for col, header in enumerate(headers, start=1):
        cell = ws.range((2, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = COLOR_AZUL
        cell.api.Font.Color = 0xFFFFFF

    # 4 tarjetas de crédito pre-cargadas
    hoy = datetime.now()
    tarjetas = [
        [hoy, "BAC San José", "CxP", "", "CRC", 1831000, "Saldo tarjeta BAC", "Tarjeta Crédito", "No", "", "No", "", "Pendiente", "", "", "Negocio", "BAC"],
        [hoy, "BCR", "CxP", "", "CRC", 1750000, "Saldo tarjeta BCR", "Tarjeta Crédito", "No", "", "No", "", "Pendiente", "", "", "Negocio", "BCR"],
        [hoy, "Amex Platinum", "CxP", "", "USD", 2500, "Saldo Amex Platinum", "Tarjeta Crédito", "No", "", "No", "", "Pendiente", "", "", "Negocio", "Credomatic"],
        [hoy, "Credomatic Gold", "CxP", "", "CRC", 2800000, "Saldo Credomatic Gold", "Tarjeta Crédito", "No", "", "No", "", "Pendiente", "", "", "Negocio", "Credomatic"],
    ]

    for i, tarjeta in enumerate(tarjetas, start=3):
        fila = i
        moneda = tarjeta[4]
        monto = tarjeta[5]

        # Datos básicos (columnas A-Q)
        for col, valor in enumerate(tarjeta, start=1):
            ws.range((fila, col)).value = valor

        # Columna S: Días Transcurridos (hoy - fecha = 0)
        ws.range((fila, 19)).value = 0

        # Columna T: Equiv USD
        if moneda == "USD":
            ws.range((fila, 20)).value = monto
        else:
            ws.range((fila, 20)).value = monto / TC_ACTUAL

        # Columna U: Fecha Vencimiento (fecha + 30 días)
        ws.range((fila, 21)).value = hoy + timedelta(days=30)

        # Columna V: Prioridad CxP
        ws.range((fila, 22)).value = "Baja"

        # Columna W: Prioridad CxC (vacío para CxP)
        ws.range((fila, 23)).value = ""

    # Dropdowns en ESPAÑOL
    crear_dropdown(ws, "C3:C1000", "Ingreso,Gasto,Inventario,Servicios,CxP,CxC,Marketing,Personal")

    # NUEVO: Dropdown Subcategorías (Columna D)
    subcategorias = (
        "COGS - Costos Ventas,"
        "Operativo - Oficina,"
        "Operativo - Tecnología,"
        "Administrativo - Contabilidad,"
        "Administrativo - Legal,"
        "Bancario - Comisiones,"
        "Bancario - Intereses,"
        "Impuestos - IVA,"
        "Impuestos - Renta,"
        "Otro"
    )
    crear_dropdown(ws, "D3:D1000", subcategorias)

    crear_dropdown(ws, "E3:E1000", "USD,CRC")
    crear_dropdown(ws, "H3:H1000", "Efectivo,Transferencia,Tarjeta Crédito,Cheque,SINPE,PayPal")
    crear_dropdown(ws, "I3:I1000", "Sí,No")
    crear_dropdown(ws, "K3:K1000", "Sí,No")
    crear_dropdown(ws, "M3:M1000", "Pagado,Pendiente,Cobrado,Cancelado")
    crear_dropdown(ws, "P3:P1000", "Personal,Negocio")

    # Formatos
    ws.range("A3:A1000").number_format = "dd/mm/yy"
    ws.range("F3:F1000").number_format = "$#,##0.00"
    ws.range("T3:T1000").number_format = "$#,##0.00"
    ws.range("U3:U1000").number_format = "dd/mm/yy"

    # Anchos de columna
    ws.range("A:A").column_width = 12
    ws.range("B:B").column_width = 25
    ws.range("C:C").column_width = 15
    ws.range("D:D").column_width = 25  # Subcategoría más ancha
    ws.range("E:E").column_width = 10
    ws.range("F:F").column_width = 15

    print("✓ Hoja TRANSACCIONES creada (23 columnas, 4 tarjetas, dropdowns incluidos)")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 3: CONFIG
    # ═════════════════════════════════════════════════════════════════
    print("📋 Generando CONFIG...")
    ws = wb.sheets.add("CONFIG")

    aplicar_estilo_header(ws, 1, 1, 3, "CONFIGURACIÓN DEL SISTEMA", COLOR_MORADO)

    ws.range("A2").value = "Parámetro"
    ws.range("B2").value = "Valor"
    ws.range("C2").value = "Descripción"
    ws.range("A2:C2").api.Font.Bold = True

    config_data = [
        ["", "", ""],  # Fila 3 vacía
        ["", "", ""],  # Fila 4 vacía
        ["Tipo Cambio USD→CRC", TC_ACTUAL, "Tipo de cambio actual (editable)"],
        ["BAC Visa - Día Corte", 15, "Día de corte mensual"],
        ["BCR Mastercard - Día Corte", 20, "Día de corte mensual"],
        ["Credomatic Platinum - Día Corte", 10, "Día de corte mensual"],
        ["Credomatic Gold - Día Corte", 10, "Día de corte mensual"],
        ["Proveedor Zona Franca 1", "VWR International LLC", "Exento de IVA"],
        ["Proveedor Zona Franca 2", "RS Hughes Co. Inc.", "Exento de IVA"],
        ["IVA Costa Rica %", 13, "Porcentaje IVA estándar"],
        ["Mes Actual de Trabajo", "Noviembre 2025", "Cambiar al hacer cierre"],
        ["Último Cierre Realizado", "31/10/2025", "Auto-actualiza"],
    ]

    for i, row in enumerate(config_data, start=3):
        ws.range((i, 1)).value = row[0]
        ws.range((i, 2)).value = row[1]
        ws.range((i, 3)).value = row[2]

    # Destacar "Mes Actual de Trabajo"
    ws.range("B13").color = (255, 235, 59)  # Amarillo
    ws.range("B13").api.Font.Bold = True

    ws.range("A:A").column_width = 30
    ws.range("B:B").column_width = 20
    ws.range("C:C").column_width = 35

    print("✓ Hoja CONFIG creada")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 4: CxP (Cuentas por Pagar) - CON FILTER
    # ═════════════════════════════════════════════════════════════════
    print("📋 Generando CxP (con FILTER real)...")
    ws = wb.sheets.add("CxP")

    aplicar_estilo_header(ws, 1, 1, 12, "CUENTAS POR PAGAR (CxP)", COLOR_ROJO)

    headers_cxp = [
        "Entidad", "Fecha", "Monto", "Moneda", "Descripción", "Estado",
        "Días Pendientes", "Equiv USD", "Factura #", "Fecha Venc", "Prioridad", "Contacto"
    ]

    for col, header in enumerate(headers_cxp, start=1):
        cell = ws.range((2, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = COLOR_ROJO
        cell.api.Font.Color = 0xFFFFFF

    # FÓRMULAS FILTER - xlwings las escribe directamente en Excel
    ws.range("A3").formula = '=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxP pendientes")'
    ws.range("B3").formula = '=FILTER(TRANSACCIONES!A3:A1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("C3").formula = '=FILTER(TRANSACCIONES!F3:F1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("D3").formula = '=FILTER(TRANSACCIONES!E3:E1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("E3").formula = '=FILTER(TRANSACCIONES!G3:G1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("F3").formula = '=FILTER(TRANSACCIONES!M3:M1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("G3").formula = '=FILTER(TRANSACCIONES!S3:S1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("H3").formula = '=FILTER(TRANSACCIONES!T3:T1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("I3").formula = '=FILTER(TRANSACCIONES!N3:N1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("J3").formula = '=FILTER(TRANSACCIONES!U3:U1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("K3").formula = '=FILTER(TRANSACCIONES!V3:V1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'

    # Formatos
    ws.range("B3").number_format = "dd/mm/yy"
    ws.range("C3").number_format = "#,##0.00"
    ws.range("H3").number_format = "$#,##0.00"
    ws.range("J3").number_format = "dd/mm/yy"

    ws.range("A:A").column_width = 25
    ws.range("B:B").column_width = 12
    ws.range("C:C").column_width = 15

    print("✓ Hoja CxP creada (FILTER nativo)")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 5: CxC (Cuentas por Cobrar) - CON FILTER
    # ═════════════════════════════════════════════════════════════════
    print("📋 Generando CxC (con FILTER real)...")
    ws = wb.sheets.add("CxC")

    aplicar_estilo_header(ws, 1, 1, 12, "CUENTAS POR COBRAR (CxC)", COLOR_VERDE)

    headers_cxc = [
        "Entidad", "Fecha", "Monto", "Moneda", "Descripción", "Estado",
        "Días Pendientes", "Equiv USD", "Factura #", "Fecha Venc", "Prioridad", "Contacto"
    ]

    for col, header in enumerate(headers_cxc, start=1):
        cell = ws.range((2, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = COLOR_VERDE
        cell.api.Font.Color = 0xFFFFFF

    # FÓRMULAS FILTER para CxC
    ws.range("A3").formula = '=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxC pendientes")'
    ws.range("B3").formula = '=FILTER(TRANSACCIONES!A3:A1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("C3").formula = '=FILTER(TRANSACCIONES!F3:F1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("D3").formula = '=FILTER(TRANSACCIONES!E3:E1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("E3").formula = '=FILTER(TRANSACCIONES!G3:G1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("F3").formula = '=FILTER(TRANSACCIONES!M3:M1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("G3").formula = '=FILTER(TRANSACCIONES!S3:S1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("H3").formula = '=FILTER(TRANSACCIONES!T3:T1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("I3").formula = '=FILTER(TRANSACCIONES!N3:N1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("J3").formula = '=FILTER(TRANSACCIONES!U3:U1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'
    ws.range("K3").formula = '=FILTER(TRANSACCIONES!W3:W1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"")'

    # Formatos
    ws.range("B3").number_format = "dd/mm/yy"
    ws.range("C3").number_format = "#,##0.00"
    ws.range("H3").number_format = "$#,##0.00"
    ws.range("J3").number_format = "dd/mm/yy"

    ws.range("A:A").column_width = 25

    print("✓ Hoja CxC creada (FILTER nativo)")

    # ═════════════════════════════════════════════════════════════════
    # HOJAS 6-21: Crear hojas placeholder (estructura básica)
    # ═════════════════════════════════════════════════════════════════
    print("\n📋 Generando hojas auxiliares (6-21)...")

    hojas_restantes = [
        ("FLUJO_CAJA", COLOR_AZUL),
        ("IVA_CONTROL", COLOR_VERDE),
        ("TARJETAS", COLOR_NARANJA),
        ("CONCILIACION", COLOR_AZUL),
        ("PERSONAL_VS_NEGOCIO", COLOR_MORADO),
        ("CATEGORIAS", COLOR_VERDE),
        ("PROYECTOS", COLOR_AZUL),
        ("PROVEEDORES", COLOR_NARANJA),
        ("CLIENTES", COLOR_VERDE),
        ("AUDITORIA", COLOR_ROJO),
        ("CIERRE_MENSUAL", COLOR_MORADO),
        ("HISTORICO_TC", COLOR_AZUL),
        ("ASIENTOS_AJUSTE", COLOR_NARANJA),
        ("BALANCE_GENERAL", COLOR_VERDE),
        ("ESTADO_RESULTADOS", COLOR_AZUL),
        ("INSTRUCCIONES", COLOR_MORADO),
    ]

    for nombre, color in hojas_restantes:
        ws = wb.sheets.add(nombre)
        aplicar_estilo_header(ws, 1, 1, 6, nombre.replace("_", " "), color)
        ws.range("A2").value = "Hoja en desarrollo - Estructura lista para implementación"
        print(f"✓ Hoja {nombre} creada")

    # ═════════════════════════════════════════════════════════════════
    # HOJA 22: ALIAS (NUEVA - Sistema de normalización)
    # ═════════════════════════════════════════════════════════════════
    print("\n📋 Generando ALIAS (normalización de entidades)...")
    ws = wb.sheets.add("ALIAS")

    aplicar_estilo_header(ws, 1, 1, 10, "SISTEMA DE NORMALIZACIÓN DE ENTIDADES", COLOR_AZUL)

    ws.range("A2").value = "💡 Esta hoja permite normalizar nombres de clientes, proveedores y bancos que aparecen con variaciones"
    ws.range("A2:J2").merge()
    ws.range("A2").api.Font.Italic = True

    # Headers
    headers_alias = [
        "Tipo", "Nombre Estándar", "Alias 1", "Alias 2", "Alias 3", "Alias 4",
        "Alias 5", "Categoría", "Notas", "Última Actualización"
    ]

    for col, header in enumerate(headers_alias, start=1):
        cell = ws.range((4, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = COLOR_AZUL
        cell.api.Font.Color = 0xFFFFFF

    # Datos pre-cargados (36 registros)
    datos_alias = [
        # CLIENTES (22)
        ["Cliente", "Grupo Acción Comercial S.A.", "Grupo Acción Comercial", "GACS", "GAC", "", "", "VIP", "Facturación Nov: $1689.04", "13/11/2025"],
        ["Cliente", "VWR International Ltda", "VWR International", "IL", "Avantor", "VWR", "", "VIP", "Facturación Nov: $1400.00", "13/11/2025"],
        ["Cliente", "Alfipac (Almacén Fiscal Pacífico)", "", "A(FP", "", "", "", "VIP", "Facturación Nov: $761.05", "13/11/2025"],
        ["Cliente", "3-102-887892 SRL", "", "", "", "", "", "Regular", "Facturación Nov: $691.56", "13/11/2025"],
        ["Cliente", "Waipio S.A.", "Waipio", "garita", "", "", "", "Regular", "Facturación Nov: $687.27", "13/11/2025"],
        ["Cliente", "Centro Integral Oncología CIO SRL", "", "CIO", "", "", "", "Regular", "Facturación Nov: $687.05", "13/11/2025"],
        ["Cliente", "Ortodoncia de la Cruz", "Dres. De la cruz", "OC", "", "", "", "Regular", "Facturación Nov: $494.50", "13/11/2025"],
        ["Cliente", "Global Automotriz GACR S.A.", "Global Automotriz GACR", "GAGS", "", "", "", "Regular", "Facturación Nov: $439.61", "13/11/2025"],
        ["Cliente", "Solusa Consolidators", "", "", "", "", "", "Regular", "Facturación Nov: $378.35", "13/11/2025"],
        ["Cliente", "Cemso", "", "", "", "", "", "Regular", "Facturación Nov: $333.92", "13/11/2025"],
        ["Cliente", "Acacia (Asoc. CR Agencias Carga)", "", "A(AC", "", "", "", "Regular", "Facturación Nov: $333.35", "13/11/2025"],
        ["Cliente", "Rodriguez Rojas Carlos Humberto", "", "RRCH", "", "", "", "Regular", "Facturación Nov: $282.50", "13/11/2025"],
        ["Cliente", "Supply Net C.R.W.H S.A.", "Supply Net C.R.W.H", "SCS", "", "", "", "Regular", "Facturación Nov: $276.85", "13/11/2025"],
        ["Cliente", "Operation Managment Tierra Magnifica", "", "OMTM", "", "", "", "Regular", "Facturación Nov: $209.06", "13/11/2025"],
        ["Cliente", "Gentra de Costa Rica S.A.", "Gentra de Costa Rica", "GCRS", "", "", "", "Regular", "Facturación Nov: $183.63", "13/11/2025"],
        ["Cliente", "Sevilla Navarro Edgar", "", "SNE", "", "", "", "Regular", "Facturación Nov: $169.50", "13/11/2025"],
        ["Cliente", "Gomez Ajoy Edgar Luis", "", "GAEL", "", "", "", "Regular", "Facturación Nov: $113.00", "13/11/2025"],
        ["Cliente", "Melendez Morales Monica", "", "MMM", "", "", "", "Regular", "Facturación Nov: $113.00", "13/11/2025"],
        ["Cliente", "Bandogo Soluciones Tecnológicas S.A.", "Bandogo Soluciones Tecnológicas", "BSTS", "", "", "", "Regular", "Facturación Nov: $67.80", "13/11/2025"],
        ["Cliente", "CPF Servicios Radiológicos S.A.", "CPF Servicios Radiológicos", "SRS", "", "", "", "Regular", "Facturación Nov: $56.50", "13/11/2025"],
        ["Cliente", "Ortodec S.A.", "Ortodec", "", "", "", "", "Regular", "Facturación Nov: $56.50", "13/11/2025"],
        ["Cliente", "Perez Morales Francisco", "", "PMF", "", "", "", "Regular", "Facturación Nov: $42.38", "13/11/2025"],

        # PROVEEDORES (5)
        ["Proveedor", "Intcomex Costa Rica", "", "", "", "", "", "Principal", "", "13/11/2025"],
        ["Proveedor", "Eurocomp S.A.", "", "", "", "", "", "Principal", "", "13/11/2025"],
        ["Proveedor", "CompuEconómicos", "", "", "", "", "", "Principal", "", "13/11/2025"],
        ["Proveedor", "TD Synex", "", "", "", "", "", "Principal", "", "13/11/2025"],
        ["Proveedor", "ICD Soft", "", "", "", "", "", "Principal", "", "13/11/2025"],

        # BANCOS (9)
        ["Banco", "BNCR CRC Ahorros (***8618)", "", "", "", "", "", "Banco", "Moneda: CRC", "13/11/2025"],
        ["Banco", "BNCR USD Ahorros (***1066)", "", "", "", "", "", "Banco", "Moneda: USD", "13/11/2025"],
        ["Banco", "BNCR CRC Corriente (***2186)", "", "", "", "", "", "Banco", "Moneda: CRC", "13/11/2025"],
        ["Banco", "BNCR USD Corriente (***9589)", "", "", "", "", "", "Banco", "Moneda: USD", "13/11/2025"],
        ["Banco", "BNCR USD Corriente (***1112)", "", "", "", "", "", "Banco", "Moneda: USD", "13/11/2025"],
        ["Banco", "Promerica CRC SINPE (***1708)", "", "", "", "", "", "Banco", "Moneda: CRC", "13/11/2025"],
        ["Banco", "Promerica USD Ahorros (***1691)", "", "", "", "", "", "Banco", "Moneda: USD", "13/11/2025"],
        ["Banco", "Promerica CRC CC Corp (***4229)", "", "", "", "", "", "Banco", "Moneda: CRC", "13/11/2025"],
        ["Banco", "Promerica USD CC Corp (***1774)", "", "", "", "", "", "Banco", "Moneda: USD", "13/11/2025"],
    ]

    for i, fila in enumerate(datos_alias, start=5):
        for col, valor in enumerate(fila, start=1):
            ws.range((i, col)).value = valor

    # Dropdowns para Tipo y Categoría
    crear_dropdown(ws, "A5:A1000", "Cliente,Proveedor,Banco")
    crear_dropdown(ws, "H5:H1000", "VIP,Regular,Principal,Banco")

    # Anchos de columna
    ws.range("A:A").column_width = 12
    ws.range("B:B").column_width = 35
    ws.range("C:G").column_width = 20
    ws.range("H:H").column_width = 12
    ws.range("I:I").column_width = 30
    ws.range("J:J").column_width = 18

    # Instrucciones al final
    fila_instruc = 45
    ws.range(f"A{fila_instruc}").value = "📋 INSTRUCCIONES:"
    ws.range(f"A{fila_instruc}").api.Font.Bold = True
    ws.range(f"A{fila_instruc+1}").value = "1. Cuando aparezca una variación de nombre, agrégala como 'Alias' en la fila correspondiente"
    ws.range(f"A{fila_instruc+2}").value = "2. Ejecuta: python scripts/normalizar_entidades_universal_v3.py"
    ws.range(f"A{fila_instruc+3}").value = "3. El script unificará todos los nombres automáticamente"

    print("✓ Hoja ALIAS creada (36 registros pre-cargados)")

    # ═════════════════════════════════════════════════════════════════
    # REORDENAR HOJAS SEGÚN ÍNDICE
    # ═════════════════════════════════════════════════════════════════
    print("\n🔄 Reordenando hojas...")

    orden_hojas = [
        "RESUMEN", "TRANSACCIONES", "CONFIG", "CxP", "CxC",
        "FLUJO_CAJA", "IVA_CONTROL", "TARJETAS", "CONCILIACION",
        "PERSONAL_VS_NEGOCIO", "CATEGORIAS", "PROYECTOS", "PROVEEDORES",
        "CLIENTES", "AUDITORIA", "CIERRE_MENSUAL", "HISTORICO_TC",
        "ASIENTOS_AJUSTE", "BALANCE_GENERAL", "ESTADO_RESULTADOS",
        "INSTRUCCIONES", "ALIAS"
    ]

    for idx, nombre in enumerate(orden_hojas, start=1):
        if nombre in [s.name for s in wb.sheets]:
            wb.sheets[nombre].api.Move(Before=wb.sheets[0].api)

    # Reversar para orden correcto
    for nombre in reversed(orden_hojas):
        if nombre in [s.name for s in wb.sheets]:
            wb.sheets[nombre].api.Move(Before=wb.sheets[0].api)

    print("✓ Hojas reordenadas correctamente")

    # ═════════════════════════════════════════════════════════════════
    # ELIMINAR HOJAS DEFAULT
    # ═════════════════════════════════════════════════════════════════
    print("\n🗑️  Eliminando hojas default de Excel...")

    # Eliminar hojas que se llaman Sheet, Sheet1, Hoja1, etc.
    hojas_a_eliminar = []
    for sheet in wb.sheets:
        nombre = sheet.name
        # Detectar hojas default (Sheet, Sheet1, Hoja, Hoja1, etc.)
        if nombre.startswith(('Sheet', 'Hoja')) or nombre in ['Sheet1', 'Sheet2', 'Sheet3']:
            hojas_a_eliminar.append(nombre)

    for nombre in hojas_a_eliminar:
        try:
            wb.sheets[nombre].delete()
            print(f"  ✓ Eliminada hoja default: {nombre}")
        except:
            pass  # Si falla, continuar

    if not hojas_a_eliminar:
        print("  ✓ No hay hojas default que eliminar")

    # ═════════════════════════════════════════════════════════════════
    # GUARDAR ARCHIVO
    # ═════════════════════════════════════════════════════════════════
    print(f"\n💾 Guardando archivo: {ARCHIVO_SALIDA}...")

    # Seleccionar primera hoja antes de guardar
    wb.sheets["RESUMEN"].activate()

    # Guardar
    wb.save(ARCHIVO_SALIDA)

    print("✓ Archivo guardado correctamente")

    # Cerrar Excel
    wb.close()
    app.quit()

    print("\n" + "="*70)
    print("✅ GENERACIÓN COMPLETADA EXITOSAMENTE")
    print("="*70)
    print(f"\n📁 Archivo: {ARCHIVO_SALIDA}")
    print(f"📊 Total hojas: 22 (21 originales + ALIAS)")
    print(f"📝 TRANSACCIONES: 23 columnas (incluye subcategorías)")
    print(f"💼 CxP/CxC: Fórmulas FILTER nativas de Excel")
    print(f"📅 Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"💵 Tipo Cambio: USD→CRC = {TC_ACTUAL}")

    print("\n📋 HOJAS CREADAS:")
    for idx, nombre in enumerate(orden_hojas, start=1):
        emoji = "✏️" if nombre in ["TRANSACCIONES", "CONFIG", "HISTORICO_TC", "ASIENTOS_AJUSTE", "ALIAS"] else "📊"
        print(f"  {idx:2d}. {emoji} {nombre}")

    print("\n✅ VERIFICACIONES RECOMENDADAS:")
    print("  1. Abrir archivo en Excel 365")
    print("  2. Verificar CxP/CxC muestran las 4 tarjetas (NO debe haber warnings)")
    print("  3. Probar dropdowns en TRANSACCIONES (columnas C, D, E, H, I, K, M, P)")
    print("  4. Verificar hoja ALIAS tiene 36 registros pre-cargados")
    print("  5. Probar agregar nueva transacción y ver si aparece en CxP/CxC")

    print("\n" + "="*70 + "\n")

except Exception as e:
    print(f"\n❌ ERROR: {str(e)}")
    import traceback
    traceback.print_exc()

    # Intentar cerrar Excel si quedó abierto
    try:
        if 'wb' in locals():
            wb.close()
        if 'app' in locals():
            app.quit()
    except:
        pass

    sys.exit(1)
