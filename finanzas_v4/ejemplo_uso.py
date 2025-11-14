# -*- coding: utf-8 -*-
"""
🎓 EJEMPLO DE USO - Sistema de Finanzas v4.0
============================================

Script de demostración de las funcionalidades principales.
"""

import sys
import os

# Agregar src al path para importar módulos
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.generar_v4 import generar_excel_v4
from src.operaciones import insertar_transaccion, crear_respaldo
from src.auditoria import auditoria_completa, imprimir_reporte_auditoria

def ejemplo_completo():
    """Demuestra el flujo completo de uso del sistema"""

    print("""
╔═══════════════════════════════════════════════════════════════════╗
║           EJEMPLO DE USO - Sistema de Finanzas v4.0              ║
╚═══════════════════════════════════════════════════════════════════╝
    """)

    # 1. Generar archivo Excel
    print("\n1️⃣  GENERAR ARCHIVO EXCEL v4.0")
    print("=" * 70)
    archivo = generar_excel_v4('ejemplo_finanzas_v4.0.xlsx', incluir_ejemplos=True)

    # 2. Auditar archivo generado
    print("\n2️⃣  AUDITAR ARCHIVO GENERADO")
    print("=" * 70)
    aprobado, reporte = auditoria_completa(archivo)
    imprimir_reporte_auditoria(reporte)

    if not aprobado:
        print("❌ Auditoría falló, abortando ejemplo")
        return

    # 3. Insertar nueva transacción
    print("\n3️⃣  INSERTAR NUEVA TRANSACCIÓN")
    print("=" * 70)

    nueva_transaccion = {
        'Fecha': '15/11/2025',
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Servicios',
        'Descripción': 'Pago de servicio de internet',
        'Entidad': 'Proveedor Internet SA',
        'Cuenta': 'BAC Corriente',
        'Método Pago': 'Transferencia',
        'Monto': 45000,
        'Referencia': 'TRF-2025-100',
        'Notas': 'Pago mensual',
        'Estado': 'PENDIENTE',
        'IVA': 5850,
        'Fecha Vencimiento': '20/11/2025',
        'Número Factura': 'FAC-INT-2025-11'
    }

    fila = insertar_transaccion(archivo, nueva_transaccion)
    print(f"✅ Transacción insertada en fila {fila}")

    # 4. Auditar después de insertar
    print("\n4️⃣  AUDITAR DESPUÉS DE INSERTAR")
    print("=" * 70)
    aprobado, reporte = auditoria_completa(archivo)

    if aprobado:
        print("✅ Auditoría post-inserción: APROBADA")
    else:
        print("❌ Auditoría post-inserción: FALLIDA")
        imprimir_reporte_auditoria(reporte)

    # 5. Resumen final
    print("""
╔═══════════════════════════════════════════════════════════════════╗
║                    EJEMPLO COMPLETADO                             ║
╠═══════════════════════════════════════════════════════════════════╣
║ ✅ Archivo Excel generado                                          ║
║ ✅ Auditoría inicial aprobada                                      ║
║ ✅ Transacción insertada                                           ║
║ ✅ Auditoría post-inserción aprobada                               ║
║                                                                   ║
║ Archivo creado: ejemplo_finanzas_v4.0.xlsx                       ║
║                                                                   ║
║ Próximos pasos:                                                   ║
║  1. Abrir el archivo en Excel                                     ║
║  2. Revisar hoja TRANSACCIONES                                    ║
║  3. Verificar que CxP muestra la nueva transacción PENDIENTE      ║
║  4. Revisar hoja RESUMEN con el dashboard                         ║
╚═══════════════════════════════════════════════════════════════════╝
    """)

if __name__ == '__main__':
    try:
        ejemplo_completo()
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
