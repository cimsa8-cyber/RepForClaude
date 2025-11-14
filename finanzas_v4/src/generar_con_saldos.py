# -*- coding: utf-8 -*-
"""
Generador de Excel v4.0 con Saldos Iniciales Reales
====================================================

Genera el archivo de producción con los saldos bancarios
proporcionados por el usuario al 14/11/2025.
"""

from generar_v4 import *
from datetime import datetime

# Saldos iniciales proporcionados por el usuario (14/11/2025)
SALDOS_INICIALES = [
    # Cuentas de Ahorro BNCR
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial cuenta ahorros',
        'Entidad': 'BNCR',
        'Cuenta': 'BNCR Ahorros Colones (***8618)',
        'Método Pago': '',
        'Monto': 35563.24,
        'Referencia': 'SALDO-INICIAL-001',
        'Notas': 'Saldo confirmado al 14/11/2025 09:59',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial cuenta ahorros dólares',
        'Entidad': 'BNCR',
        'Cuenta': 'BNCR Ahorros Dólares (***1066)',
        'Método Pago': '',
        'Monto': 1014.39,
        'Referencia': 'SALDO-INICIAL-002',
        'Notas': 'Saldo confirmado al 14/11/2025 09:59 - USD',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    # Cuentas Corrientes BNCR
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial cuenta corriente',
        'Entidad': 'BNCR',
        'Cuenta': 'BNCR CC Colones (***2186)',
        'Método Pago': '',
        'Monto': 7950.50,
        'Referencia': 'SALDO-INICIAL-003',
        'Notas': 'Saldo confirmado al 14/11/2025 09:59',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    # Cuentas Promerica
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial SINPE Promerica',
        'Entidad': 'Promerica',
        'Cuenta': 'Promerica SINPE Colones (***1708)',
        'Método Pago': '',
        'Monto': 1090.00,
        'Referencia': 'SALDO-INICIAL-012',
        'Notas': 'AlvaroVelascoNet SRL - Saldo confirmado 14/11/2025',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial cuenta corriente dólares',
        'Entidad': 'BNCR',
        'Cuenta': 'BNCR CC Dólares (***9589)',
        'Método Pago': '',
        'Monto': 0.43,
        'Referencia': 'SALDO-INICIAL-004',
        'Notas': 'Saldo confirmado al 14/11/2025 09:59 - USD',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial cuenta corriente dólares',
        'Entidad': 'BNCR',
        'Cuenta': 'BNCR CC Dólares (***1112)',
        'Método Pago': '',
        'Monto': 21.84,
        'Referencia': 'SALDO-INICIAL-005',
        'Notas': 'Saldo confirmado al 14/11/2025 09:59 - USD',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'INGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Saldo Inicial',
        'Descripción': 'Saldo inicial CC corporativa dólares',
        'Entidad': 'Promerica',
        'Cuenta': 'Promerica CC Corporativa Dólares (***1774)',
        'Método Pago': '',
        'Monto': 3676.44,
        'Referencia': 'SALDO-INICIAL-013',
        'Notas': 'AlvaroVelascoNet SRL - Saldo confirmado 14/11/2025 - USD',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    # Tarjetas de Crédito BNCR (DEUDA) - ⚠️ Estado PENDIENTE para aparecer en CxP
    # Tarjeta 1: Visa Clásica (***3519) - Fecha pago: 17 de cada mes
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta Visa Clásica ₡ - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta Visa Clásica BNCR (***3519)',
        'Método Pago': 'Tarjeta',
        'Monto': -590158.64,
        'Referencia': 'TC-VISA-CL-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - Límite $1,200.00 - Pago mensual: 17',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 17),
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta Visa Clásica USD - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta Visa Clásica BNCR (***3519)',
        'Método Pago': 'Tarjeta',
        'Monto': -65.04,
        'Referencia': 'TC-VISA-CL-USD-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - USD - Pago mensual: 17',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 17),
        'Número Factura': ''
    },
    # Tarjeta 2: Visa Platino (***9837) - Fecha pago: 5 de cada mes
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta Visa Platino ₡ - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta Visa Platino BNCR (***9837)',
        'Método Pago': 'Tarjeta',
        'Monto': -2086984.01,
        'Referencia': 'TC-VISA-PT-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - Límite $9,000.00 - Pago mensual: 5',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 5),
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta Visa Platino USD - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta Visa Platino BNCR (***9837)',
        'Método Pago': 'Tarjeta',
        'Monto': -1775.45,
        'Referencia': 'TC-VISA-PT-USD-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - USD - Pago mensual: 5',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 5),
        'Número Factura': ''
    },
    # Tarjeta 3: MasterCard Oro (***8759) - Fecha pago: 13 de cada mes
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta MasterCard Oro ₡ - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta MasterCard Oro BNCR (***8759)',
        'Método Pago': 'Tarjeta',
        'Monto': -2847410.17,
        'Referencia': 'TC-MC-ORO-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - Límite $6,000.00 - Pago mensual: 13',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 13),
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': 'Pago tarjeta MasterCard Oro USD - Corte Nov 2025',
        'Entidad': 'BNCR',
        'Cuenta': 'Tarjeta MasterCard Oro BNCR (***8759)',
        'Método Pago': 'Tarjeta',
        'Monto': -256.09,
        'Referencia': 'TC-MC-ORO-USD-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - USD - Pago mensual: 13',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 12, 13),
        'Número Factura': ''
    },
    # Tarjeta 4: Credomatic - Fecha pago: 15 de cada mes ⚠️ URGENTE MAÑANA
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': '🔴 URGENTE: Pago tarjeta Credomatic ₡ - Vence MAÑANA',
        'Entidad': 'Credomatic',
        'Cuenta': 'Tarjeta Credomatic',
        'Método Pago': 'Tarjeta',
        'Monto': -1419305.54,
        'Referencia': 'TC-CREDO-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - Pago mensual: 15 - ⚠️ VENCE 15/11/2025',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 11, 15),
        'Número Factura': ''
    },
    {
        'Fecha': datetime(2025, 11, 14),
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Deuda Tarjeta Crédito',
        'Descripción': '🔴 URGENTE: Pago tarjeta Credomatic USD - Vence MAÑANA',
        'Entidad': 'Credomatic',
        'Cuenta': 'Tarjeta Credomatic',
        'Método Pago': 'Tarjeta',
        'Monto': -209.92,
        'Referencia': 'TC-CREDO-USD-NOV2025',
        'Notas': 'Saldo al 14/11/2025 - USD - Pago mensual: 15 - ⚠️ VENCE 15/11/2025',
        'Estado': 'PENDIENTE',
        'IVA': 0,
        'Fecha Vencimiento': datetime(2025, 11, 15),
        'Número Factura': ''
    }
]

def crear_hoja_transacciones_con_saldos(wb):
    """Crea la hoja de TRANSACCIONES con saldos iniciales reales"""
    print("📝 Creando hoja TRANSACCIONES con saldos iniciales...")

    # Crear o obtener hoja
    if 'TRANSACCIONES' in wb.sheetnames:
        ws = wb['TRANSACCIONES']
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet('TRANSACCIONES')

    # Escribir headers (fila 1)
    for nombre_campo, config in ESTRUCTURA_TRANSACCIONES.items():
        col = config['col']
        ws.cell(1, col, nombre_campo)

    # Aplicar estilo a headers
    aplicar_estilo_header(ws, 1, HOJAS_CONFIG['TRANSACCIONES']['color_header'])

    # Ajustar anchos de columna
    ajustar_anchos_columnas(ws, ESTRUCTURA_TRANSACCIONES)

    # Congelar primera fila
    ws.freeze_panes = 'A2'

    # Insertar saldos iniciales
    print("  💰 Insertando saldos iniciales reales...")
    for saldo in SALDOS_INICIALES:
        fila = ws.max_row + 1

        for nombre_campo, config in ESTRUCTURA_TRANSACCIONES.items():
            col = config['col']
            valor = saldo.get(nombre_campo)

            # Escribir valor
            celda = ws.cell(fila, col, valor)

            # Aplicar formato
            formato = config.get('formato_openpyxl')
            if formato:
                celda.number_format = formato

    cantidad_saldos = len(SALDOS_INICIALES)
    print(f"  ✅ Hoja TRANSACCIONES creada con {cantidad_saldos} saldos iniciales")
    return ws

def generar_excel_produccion(nombre_archivo='AlvaroVelasco_Finanzas_v4.0.xlsx'):
    """
    Genera el archivo Excel v4.0 de PRODUCCIÓN con saldos iniciales reales
    """
    print(f"""
╔═══════════════════════════════════════════════════════════════════╗
║     GENERADOR DE ARCHIVO DE PRODUCCIÓN v4.0 - INICIO              ║
╠═══════════════════════════════════════════════════════════════════╣
║ Archivo: {nombre_archivo:<55} ║
║ Saldos iniciales: 15 transacciones (7 cuentas + 8 tarjetas)     ║
║ Fecha saldos: 14/11/2025 09:59                                   ║
╚═══════════════════════════════════════════════════════════════════╝
    """)

    # Crear workbook
    wb = openpyxl.Workbook()

    # Eliminar hoja por defecto
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Crear hojas en orden
    crear_hoja_transacciones_con_saldos(wb)  # Con saldos reales
    crear_hoja_cxp(wb)
    crear_hoja_cxc(wb)
    crear_hoja_entidades_alias(wb)
    crear_hoja_resumen(wb)

    # ══════════════════════════════════════════════════════════════════════════
    # NUEVAS HOJAS v4.0 COMPLETO (100%)
    # ══════════════════════════════════════════════════════════════════════════
    agregar_estados_financieros(wb)  # P&L, Balance, Flujo Caja
    agregar_graficas(wb)              # Dashboard Visual con KPIs
    agregar_conciliacion(wb)          # Conciliación Bancaria

    # Proteger hojas (todas excepto TRANSACCIONES)
    proteger_hojas(wb)

    # Activar hoja RESUMEN por defecto
    wb.active = wb['RESUMEN']

    # Guardar
    print(f"\n💾 Guardando archivo: {nombre_archivo}")
    wb.save(nombre_archivo)
    wb.close()

    # Calcular totales
    cuentas_colones = 35563.24 + 7950.50 + 1090.00
    cuentas_dolares = 1014.39 + 0.43 + 21.84 + 3676.44
    tarjetas_colones = 590158.64 + 2086984.01 + 2847410.17 + 1419305.54
    tarjetas_dolares = 65.04 + 1775.45 + 256.09 + 209.92
    total_colones = cuentas_colones - tarjetas_colones
    total_dolares = cuentas_dolares - tarjetas_dolares

    print(f"""
╔═══════════════════════════════════════════════════════════════════╗
║      ✅ ARCHIVO DE PRODUCCIÓN v4.0 GENERADO EXITOSAMENTE           ║
╠═══════════════════════════════════════════════════════════════════╣
║ Archivo creado: {nombre_archivo:<49} ║
║                                                                   ║
║ 💰 SALDOS BANCARIOS (7 cuentas):                                  ║
║  🏦 BNCR:                                                          ║
║   ✅ Ahorros Colones (***8618): ₡35,563.24                        ║
║   ✅ Ahorros Dólares (***1066): $1,014.39                         ║
║   ✅ CC Colones (***2186): ₡7,950.50                              ║
║   ✅ CC Dólares (***9589): $0.43                                  ║
║   ✅ CC Dólares (***1112): $21.84                                 ║
║  🏦 PROMERICA (AlvaroVelascoNet SRL):                             ║
║   ✅ SINPE Colones (***1708): ₡1,090.00                           ║
║   ✅ CC Corporativa Dólares (***1774): $3,676.44                  ║
║     Subtotal: ₡44,603.74 / $4,713.10                             ║
║                                                                   ║
║ 💳 TARJETAS DE CRÉDITO - DEUDAS (4 tarjetas):                     ║
║  ❌ Visa Clásica (***3519): -₡590,158.64 / -$65.04               ║
║     Límite: $1,200.00                                             ║
║  ❌ Visa Platino (***9837): -₡2,086,984.01 / -$1,775.45          ║
║     Límite: $9,000.00                                             ║
║  ❌ MasterCard Oro (***8759): -₡2,847,410.17 / -$256.09          ║
║     Límite: $6,000.00                                             ║
║  ❌ Credomatic: -₡1,419,305.54 / -$209.92                        ║
║     Pago: 15/11/2025                                              ║
║     Subtotal deuda: -₡6,943,858.36 / -$2,306.50                  ║
║                                                                   ║
║ 📊 POSICIÓN FINANCIERA NETA:                                      ║
║  Colones: ₡{total_colones:>20,.2f} 🔴                            ║
║  Dólares: ${total_dolares:>20,.2f} 🟢                            ║
║                                                                   ║
║ 📋 Total transacciones: 15 saldos iniciales                       ║
║                                                                   ║
║ Hojas creadas:                                                    ║
║  ✅ RESUMEN (Dashboard) - 🔒 PROTEGIDA                             ║
║  ✅ TRANSACCIONES (15 saldos) - 🔓 EDITABLE                        ║
║  ✅ CxP (Cuentas por Pagar) - 🔒 PROTEGIDA                         ║
║  ✅ CxC (Cuentas por Cobrar) - 🔒 PROTEGIDA                        ║
║  ✅ ENTIDADES_ALIAS (12 alias) - 🔒 PROTEGIDA                      ║
║                                                                   ║
║ 🔐 PROTECCIÓN APLICADA:                                            ║
║  - Solo TRANSACCIONES es editable por el usuario                  ║
║  - Todas las demás hojas están protegidas (solo fórmulas)         ║
║                                                                   ║
║ ⚠️  ANÁLISIS FINANCIERO:                                           ║
║  💰 Efectivo total: ₡44,603.74 + $4,713.10                        ║
║  💳 Deuda total: ₡6,943,858.36 + $2,306.50                        ║
║  📊 Posición neta: -₡6,899,254.62 + $2,406.60                    ║
║                                                                   ║
║  🔴 Deuda en colones CRÍTICA: -₡6.90M (+₡1.42M Credomatic)       ║
║  🟢 Posición en dólares POSITIVA: +$2,406.60                     ║
║  ⚠️  Límite total disponible BNCR: $16,200.00 (3 tarjetas)       ║
║  ⚠️  Pago urgente Credomatic: 15/11/2025 (MAÑANA)                ║
║                                                                   ║
║  🚨 RECOMENDACIÓN: Estrategia URGENTE de pago de tarjetas        ║
║     Prioridad #1: Credomatic (vence mañana)                      ║
╚═══════════════════════════════════════════════════════════════════╝
    """)

    return nombre_archivo

if __name__ == '__main__':
    import sys
    import os

    # Cambiar al directorio raíz del proyecto
    os.chdir('..')

    # Nombre del archivo
    nombre_archivo = 'AlvaroVelasco_Finanzas_v4.0.xlsx'

    # Generar
    try:
        archivo_creado = generar_excel_produccion(nombre_archivo)

        print("\n🎉 ¡LISTO! El archivo de PRODUCCIÓN está listo para usar.")
        print(f"📁 Ubicación: {os.path.abspath(archivo_creado)}")
        print("\n📝 Próximos pasos:")
        print("  1. Abrir archivo en Excel")
        print("  2. Verificar saldos en TRANSACCIONES")
        print("  3. Ver el dashboard en RESUMEN")
        print("  4. Comenzar a registrar nuevas transacciones")

        sys.exit(0)

    except Exception as e:
        print(f"\n❌ ERROR al generar archivo: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
