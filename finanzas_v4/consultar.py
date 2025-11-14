#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                  🔍 CONSULTAR ESTADO - Sistema v4.0                          ║
╚══════════════════════════════════════════════════════════════════════════════╝

Script para consultar estado financiero rápidamente sin abrir Excel.

USO:
    python consultar.py               # Ver resumen completo
    python consultar.py --cxp         # Solo CxP
    python consultar.py --cxc         # Solo CxC
    python consultar.py --saldos      # Solo saldos bancarios
"""

import sys
from pathlib import Path
from datetime import datetime
import openpyxl

# Agregar src al path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from config import ESTRUCTURA_TRANSACCIONES

# Archivo de producción (en raíz del proyecto)
ARCHIVO_EXCEL = Path(__file__).parent.parent / "AlvaroVelasco_Finanzas_v4.0.xlsx"

def leer_transacciones(wb):
    """Lee todas las transacciones"""
    ws = wb['TRANSACCIONES']
    transacciones = []

    # Leer desde fila 2 (fila 1 son headers)
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Si la fila está vacía, terminar
        if not row[0]:  # Si no hay fecha, terminar
            break

        # Mapear a dict usando la estructura
        trans = {}
        for nombre_campo, config in ESTRUCTURA_TRANSACCIONES.items():
            col_idx = config['col'] - 1  # -1 porque values_only usa índice 0
            trans[nombre_campo] = row[col_idx] if col_idx < len(row) else None

        transacciones.append(trans)

    return transacciones

def calcular_saldos(transacciones):
    """Calcula saldos por cuenta"""
    saldos = {}

    for trans in transacciones:
        cuenta = trans.get('Cuenta', '')
        monto = trans.get('Monto', 0)
        estado = trans.get('Estado', '')

        # Solo contar transacciones completadas
        if estado in ['PAGADO', 'COBRADO']:
            if cuenta:
                saldos[cuenta] = saldos.get(cuenta, 0) + (monto or 0)

    return saldos

def obtener_cxp(transacciones):
    """Obtiene cuentas por pagar"""
    cxp = []

    for trans in transacciones:
        tipo = trans.get('Tipo', '')
        estado = trans.get('Estado', '')

        if tipo == 'EGRESO' and estado in ['PENDIENTE', 'PARCIAL']:
            cxp.append(trans)

    # Ordenar por fecha de vencimiento
    cxp.sort(key=lambda x: x.get('Fecha Vencimiento') or datetime(2099, 12, 31))

    return cxp

def obtener_cxc(transacciones):
    """Obtiene cuentas por cobrar"""
    cxc = []

    for trans in transacciones:
        tipo = trans.get('Tipo', '')
        estado = trans.get('Estado', '')

        if tipo == 'INGRESO' and estado in ['POR COBRAR', 'PARCIAL']:
            cxc.append(trans)

    # Ordenar por fecha de vencimiento
    cxc.sort(key=lambda x: x.get('Fecha Vencimiento') or datetime(2099, 12, 31))

    return cxc

def mostrar_saldos(saldos):
    """Muestra saldos bancarios"""
    print("\n" + "═"*80)
    print("💰 SALDOS BANCARIOS")
    print("═"*80)

    # Separar cuentas bancarias de tarjetas
    cuentas_banco = {}
    tarjetas = {}

    for cuenta, saldo in saldos.items():
        if 'Tarjeta' in cuenta:
            tarjetas[cuenta] = saldo
        else:
            cuentas_banco[cuenta] = saldo

    # Mostrar cuentas bancarias
    if cuentas_banco:
        print("\n🏦 Cuentas Bancarias:")
        total_banco = 0
        for cuenta, saldo in sorted(cuentas_banco.items()):
            signo = "🟢" if saldo >= 0 else "🔴"
            print(f"  {signo} {cuenta:50} ₡{saldo:>15,.2f}")
            total_banco += saldo

        print(f"  {'─'*70}")
        print(f"  {'TOTAL BANCOS':50} ₡{total_banco:>15,.2f}\n")

    # Mostrar tarjetas
    if tarjetas:
        print("💳 Tarjetas de Crédito:")
        total_tarjetas = 0
        for cuenta, saldo in sorted(tarjetas.items()):
            signo = "🟢" if saldo >= 0 else "🔴"
            print(f"  {signo} {cuenta:50} ₡{saldo:>15,.2f}")
            total_tarjetas += saldo

        print(f"  {'─'*70}")
        print(f"  {'TOTAL TARJETAS':50} ₡{total_tarjetas:>15,.2f}\n")

    # Total general
    total_general = sum(saldos.values())
    print("═"*80)
    print(f"{'POSICIÓN FINANCIERA NETA':50} ₡{total_general:>15,.2f}")
    print("═"*80)

def mostrar_cxp(cxp):
    """Muestra cuentas por pagar"""
    print("\n" + "═"*80)
    print("📤 CUENTAS POR PAGAR (CxP)")
    print("═"*80)

    if not cxp:
        print("\n✅ No hay cuentas pendientes de pago\n")
        return

    total = 0
    hoy = datetime.now()

    for i, trans in enumerate(cxp, 1):
        fecha_venc = trans.get('Fecha Vencimiento')
        descripcion = trans.get('Descripción', '')
        monto = abs(trans.get('Monto', 0))
        entidad = trans.get('Entidad', '')
        estado = trans.get('Estado', '')

        # Determinar si está vencida
        if fecha_venc and fecha_venc < hoy:
            status = "🔴 VENCIDA"
        elif fecha_venc and (fecha_venc - hoy).days <= 7:
            status = "⚠️  PRÓXIMA"
        else:
            status = "🟡 PENDIENTE"

        venc_str = fecha_venc.strftime('%d/%m/%Y') if fecha_venc else 'Sin fecha'

        print(f"\n{i}. {status}")
        print(f"   Entidad: {entidad}")
        print(f"   Descripción: {descripcion}")
        print(f"   Monto: ₡{monto:,.2f}")
        print(f"   Vencimiento: {venc_str}")
        print(f"   Estado: {estado}")

        total += monto

    print("\n" + "─"*80)
    print(f"TOTAL POR PAGAR: ₡{total:,.2f}")
    print("═"*80)

def mostrar_cxc(cxc):
    """Muestra cuentas por cobrar"""
    print("\n" + "═"*80)
    print("📥 CUENTAS POR COBRAR (CxC)")
    print("═"*80)

    if not cxc:
        print("\n✅ No hay cuentas pendientes de cobro\n")
        return

    total = 0
    hoy = datetime.now()

    for i, trans in enumerate(cxc, 1):
        fecha_venc = trans.get('Fecha Vencimiento')
        descripcion = trans.get('Descripción', '')
        monto = abs(trans.get('Monto', 0))
        entidad = trans.get('Entidad', '')
        estado = trans.get('Estado', '')

        # Determinar si está vencida
        if fecha_venc and fecha_venc < hoy:
            status = "🔴 VENCIDA"
        elif fecha_venc and (fecha_venc - hoy).days <= 7:
            status = "⚠️  PRÓXIMA"
        else:
            status = "🟡 PENDIENTE"

        venc_str = fecha_venc.strftime('%d/%m/%Y') if fecha_venc else 'Sin fecha'

        print(f"\n{i}. {status}")
        print(f"   Cliente: {entidad}")
        print(f"   Descripción: {descripcion}")
        print(f"   Monto: ₡{monto:,.2f}")
        print(f"   Vencimiento: {venc_str}")
        print(f"   Estado: {estado}")

        total += monto

    print("\n" + "─"*80)
    print(f"TOTAL POR COBRAR: ₡{total:,.2f}")
    print("═"*80)

def mostrar_resumen_completo(transacciones):
    """Muestra resumen completo"""
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                 📊 ESTADO FINANCIERO - Sistema v4.0                          ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")

    print(f"📅 Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"📝 Total transacciones: {len(transacciones)}")

    # Calcular y mostrar saldos
    saldos = calcular_saldos(transacciones)
    mostrar_saldos(saldos)

    # Obtener y mostrar CxP
    cxp = obtener_cxp(transacciones)
    mostrar_cxp(cxp)

    # Obtener y mostrar CxC
    cxc = obtener_cxc(transacciones)
    mostrar_cxc(cxc)

    print("\n")

def main():
    """Función principal"""

    # Verificar que existe el archivo
    if not ARCHIVO_EXCEL.exists():
        print(f"\n❌ Error: No se encuentra el archivo {ARCHIVO_EXCEL}")
        print("   Ejecuta primero: python finanzas_v4/src/generar_con_saldos.py")
        sys.exit(1)

    # Abrir archivo
    try:
        wb = openpyxl.load_workbook(ARCHIVO_EXCEL, data_only=True)
    except Exception as e:
        print(f"\n❌ Error al abrir archivo: {e}")
        sys.exit(1)

    # Leer transacciones
    transacciones = leer_transacciones(wb)

    # Procesar según argumentos
    if '--saldos' in sys.argv:
        print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     💰 SALDOS BANCARIOS - Sistema v4.0                       ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")
        saldos = calcular_saldos(transacciones)
        mostrar_saldos(saldos)

    elif '--cxp' in sys.argv:
        print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                   📤 CUENTAS POR PAGAR - Sistema v4.0                        ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")
        cxp = obtener_cxp(transacciones)
        mostrar_cxp(cxp)

    elif '--cxc' in sys.argv:
        print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                   📥 CUENTAS POR COBRAR - Sistema v4.0                       ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")
        cxc = obtener_cxc(transacciones)
        mostrar_cxc(cxc)

    else:
        # Mostrar resumen completo
        mostrar_resumen_completo(transacciones)

    print("\n")

if __name__ == "__main__":
    main()
