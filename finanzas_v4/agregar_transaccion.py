#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║              ➕ AGREGAR TRANSACCIÓN RÁPIDA - Sistema v4.0                    ║
╚══════════════════════════════════════════════════════════════════════════════╝

Script para agregar transacciones de forma rápida e interactiva.
Ideal para uso diario.

USO:
    python agregar_transaccion.py

    O para agregar múltiples transacciones:
    python agregar_transaccion.py --multiple
"""

import sys
from pathlib import Path
from datetime import datetime

# Agregar src al path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from operaciones import insertar_transaccion

# Archivo de producción
ARCHIVO_EXCEL = Path(__file__).parent / "AlvaroVelasco_Finanzas_v4.0.xlsx"

# Plantillas rápidas para operaciones comunes
PLANTILLAS = {
    '1': {
        'nombre': 'Ingreso por Venta',
        'defaults': {
            'Tipo': 'INGRESO',
            'Categoría': 'Ingresos',
            'Subcategoría': 'Venta de Servicios',
            'Estado': 'COBRADO'
        }
    },
    '2': {
        'nombre': 'Pago Tarjeta de Crédito',
        'defaults': {
            'Tipo': 'EGRESO',
            'Categoría': 'Operaciones',
            'Subcategoría': 'Pago Tarjeta',
            'Método Pago': 'Tarjeta',
            'Estado': 'PAGADO'
        }
    },
    '3': {
        'nombre': 'Gasto Operativo',
        'defaults': {
            'Tipo': 'EGRESO',
            'Categoría': 'Operaciones',
            'Estado': 'PAGADO'
        }
    },
    '4': {
        'nombre': 'Cuenta por Pagar',
        'defaults': {
            'Tipo': 'EGRESO',
            'Estado': 'PENDIENTE'
        }
    },
    '5': {
        'nombre': 'Cuenta por Cobrar',
        'defaults': {
            'Tipo': 'INGRESO',
            'Estado': 'POR COBRAR'
        }
    }
}

def mostrar_menu():
    """Muestra el menú de plantillas"""
    print("\n" + "="*80)
    print("📋 PLANTILLAS RÁPIDAS")
    print("="*80)
    for key, plantilla in PLANTILLAS.items():
        print(f"  {key}. {plantilla['nombre']}")
    print("  0. Transacción personalizada (ingresar todo)")
    print("="*80)

def solicitar_fecha(mensaje="Fecha (DD/MM/YYYY o Enter para hoy)"):
    """Solicita una fecha al usuario"""
    while True:
        fecha_str = input(f"\n{mensaje}: ").strip()

        if not fecha_str:
            return datetime.now()

        try:
            return datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            print("❌ Formato inválido. Usa DD/MM/YYYY (ejemplo: 14/11/2025)")

def solicitar_monto(mensaje="Monto"):
    """Solicita un monto al usuario"""
    while True:
        monto_str = input(f"\n{mensaje}: ").strip().replace(",", "")

        try:
            monto = float(monto_str)
            return monto
        except ValueError:
            print("❌ Monto inválido. Usa números (ejemplo: 150000 o 150000.50)")

def solicitar_campo(nombre_campo, valor_default=None, requerido=True):
    """Solicita un campo genérico"""
    if valor_default:
        prompt = f"{nombre_campo} [{valor_default}]"
    else:
        prompt = nombre_campo

    if requerido:
        prompt += " *"

    valor = input(f"\n{prompt}: ").strip()

    if not valor and valor_default:
        return valor_default

    if not valor and requerido:
        print(f"❌ {nombre_campo} es requerido")
        return solicitar_campo(nombre_campo, valor_default, requerido)

    return valor if valor else ""

def agregar_transaccion_interactiva(plantilla_id=None):
    """Agrega una transacción de forma interactiva"""

    print("\n" + "╔" + "="*78 + "╗")
    print("║" + " "*20 + "➕ AGREGAR TRANSACCIÓN" + " "*37 + "║")
    print("╚" + "="*78 + "╝")

    # Seleccionar plantilla si no se proporcionó
    if plantilla_id is None:
        mostrar_menu()
        plantilla_id = input("\nSelecciona plantilla (0-5): ").strip()

    # Obtener defaults de la plantilla
    if plantilla_id in PLANTILLAS:
        defaults = PLANTILLAS[plantilla_id]['defaults']
        print(f"\n✅ Usando plantilla: {PLANTILLAS[plantilla_id]['nombre']}")
    else:
        defaults = {}
        print("\n📝 Transacción personalizada")

    # Recopilar datos
    transaccion = {}

    # Fecha
    transaccion['Fecha'] = solicitar_fecha()

    # Tipo
    tipo_default = defaults.get('Tipo')
    if tipo_default:
        print(f"\nTipo: {tipo_default}")
        transaccion['Tipo'] = tipo_default
    else:
        tipo = input("\nTipo (INGRESO/EGRESO) *: ").strip().upper()
        while tipo not in ['INGRESO', 'EGRESO']:
            print("❌ Debe ser INGRESO o EGRESO")
            tipo = input("Tipo (INGRESO/EGRESO) *: ").strip().upper()
        transaccion['Tipo'] = tipo

    # Descripción
    transaccion['Descripción'] = solicitar_campo("Descripción", requerido=True)

    # Monto
    monto = solicitar_monto("Monto")
    # Ajustar signo según tipo
    if transaccion['Tipo'] == 'EGRESO' and monto > 0:
        monto = -monto
    transaccion['Monto'] = monto

    # Categoría
    transaccion['Categoría'] = solicitar_campo(
        "Categoría",
        defaults.get('Categoría'),
        requerido=True
    )

    # Subcategoría
    transaccion['Subcategoría'] = solicitar_campo(
        "Subcategoría",
        defaults.get('Subcategoría'),
        requerido=False
    )

    # Entidad
    transaccion['Entidad'] = solicitar_campo("Entidad", requerido=False)

    # Cuenta
    transaccion['Cuenta'] = solicitar_campo("Cuenta/Tarjeta", requerido=False)

    # Método de Pago
    transaccion['Método Pago'] = solicitar_campo(
        "Método Pago",
        defaults.get('Método Pago'),
        requerido=False
    )

    # Estado
    estado_default = defaults.get('Estado', 'PAGADO')
    print(f"\nEstado [{estado_default}]:")
    print("  1. PAGADO")
    print("  2. PENDIENTE")
    print("  3. POR COBRAR")
    print("  4. COBRADO")
    print("  5. CANCELADO")
    print("  6. PARCIAL")
    estado_opcion = input("Selecciona (Enter para default): ").strip()

    estados = {
        '1': 'PAGADO',
        '2': 'PENDIENTE',
        '3': 'POR COBRAR',
        '4': 'COBRADO',
        '5': 'CANCELADO',
        '6': 'PARCIAL'
    }

    transaccion['Estado'] = estados.get(estado_opcion, estado_default)

    # Referencia
    transaccion['Referencia'] = solicitar_campo("Referencia", requerido=False)

    # Notas
    transaccion['Notas'] = solicitar_campo("Notas", requerido=False)

    # IVA (opcional)
    iva_str = input("\nIVA (dejar vacío si no aplica): ").strip().replace(",", "")
    transaccion['IVA'] = float(iva_str) if iva_str else 0

    # Fecha vencimiento (opcional)
    if transaccion['Estado'] in ['PENDIENTE', 'POR COBRAR']:
        usar_vencimiento = input("\n¿Tiene fecha de vencimiento? (s/N): ").strip().lower()
        if usar_vencimiento == 's':
            transaccion['Fecha Vencimiento'] = solicitar_fecha("Fecha vencimiento")
        else:
            transaccion['Fecha Vencimiento'] = ''
    else:
        transaccion['Fecha Vencimiento'] = ''

    # Número de factura (opcional)
    transaccion['Número Factura'] = solicitar_campo("Número Factura", requerido=False)

    # Mostrar resumen
    print("\n" + "="*80)
    print("📋 RESUMEN DE LA TRANSACCIÓN")
    print("="*80)
    print(f"Fecha: {transaccion['Fecha'].strftime('%d/%m/%Y')}")
    print(f"Tipo: {transaccion['Tipo']}")
    print(f"Descripción: {transaccion['Descripción']}")
    print(f"Monto: ₡{transaccion['Monto']:,.2f}" if transaccion['Monto'] >= 0 else f"Monto: -₡{abs(transaccion['Monto']):,.2f}")
    print(f"Categoría: {transaccion['Categoría']}")
    if transaccion['Subcategoría']:
        print(f"Subcategoría: {transaccion['Subcategoría']}")
    if transaccion['Entidad']:
        print(f"Entidad: {transaccion['Entidad']}")
    if transaccion['Cuenta']:
        print(f"Cuenta: {transaccion['Cuenta']}")
    print(f"Estado: {transaccion['Estado']}")
    if transaccion['Referencia']:
        print(f"Referencia: {transaccion['Referencia']}")
    if transaccion['Notas']:
        print(f"Notas: {transaccion['Notas']}")
    print("="*80)

    # Confirmar
    confirmar = input("\n¿Guardar esta transacción? (S/n): ").strip().lower()

    if confirmar == 'n':
        print("\n❌ Transacción cancelada")
        return False

    # Agregar transacción (la función ya maneja todo internamente)
    try:
        insertar_transaccion(ARCHIVO_EXCEL, transaccion, validar=True)
        print("\n✅ ¡Transacción agregada exitosamente!")
        return True

    except Exception as e:
        print(f"\n❌ Error al agregar transacción: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal"""

    # Verificar que existe el archivo
    if not ARCHIVO_EXCEL.exists():
        print(f"\n❌ Error: No se encuentra el archivo {ARCHIVO_EXCEL}")
        print("   Ejecuta primero: python finanzas_v4/src/generar_con_saldos.py")
        sys.exit(1)

    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    Sistema de Finanzas v4.0 - Álvaro Velasco                 ║
║                         Agregar Transacciones Rápido                          ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")

    # Modo múltiple
    if '--multiple' in sys.argv:
        while True:
            success = agregar_transaccion_interactiva()

            if success:
                otra = input("\n¿Agregar otra transacción? (S/n): ").strip().lower()
                if otra == 'n':
                    break
            else:
                break
    else:
        agregar_transaccion_interactiva()

    print("\n👋 ¡Hasta luego!\n")

if __name__ == "__main__":
    main()
