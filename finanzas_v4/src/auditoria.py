# -*- coding: utf-8 -*-
"""
🔍 SCRIPT DE AUDITORÍA - Sistema de Finanzas v4.0
==================================================

Ejecuta auditoría completa del archivo Excel y genera reporte.

FUNCIONALIDADES:
- Valida estructura de todas las hojas
- Verifica formatos de fecha y moneda
- Valida integridad de datos
- Verifica que CxP/CxC funcionen
- Genera reporte detallado
- Guarda log de auditoría

USO:
    python auditoria.py --archivo AlvaroVelasco_Finanzas_v4.0.xlsx
    python auditoria.py --archivo v4.0.xlsx --verbose
    python auditoria.py --archivo v4.0.xlsx --guardar-log
"""

import sys
import argparse
from datetime import datetime
from pathlib import Path

from validaciones import (
    auditoria_completa,
    imprimir_reporte_auditoria
)
from config import LOGS_CONFIG

# ==============================================================================
# GUARDAR LOG
# ==============================================================================

def guardar_log_auditoria(reporte):
    """Guarda el reporte de auditoría en un archivo de log"""
    # Crear directorio de logs si no existe
    logs_dir = Path(LOGS_CONFIG['directorio'])
    logs_dir.mkdir(exist_ok=True)

    # Nombre del log
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_log = f"auditoria_{timestamp}.log"
    path_log = logs_dir / nombre_log

    # Escribir log
    with open(path_log, 'w', encoding='utf-8') as f:
        f.write("=" * 70 + "\n")
        f.write("REPORTE DE AUDITORÍA - Sistema de Finanzas v4.0\n")
        f.write("=" * 70 + "\n\n")

        f.write(f"Archivo: {reporte['archivo']}\n")
        f.write(f"Fecha: {reporte['fecha_auditoria']}\n")
        f.write(f"Errores críticos: {reporte['errores_criticos']}\n")
        f.write(f"Advertencias: {reporte['advertencias']}\n")
        f.write(f"Estado: {'APROBADO' if reporte['aprobado'] else 'FALLIDO'}\n")
        f.write("\n" + "=" * 70 + "\n\n")

        # Detalle de validaciones
        for nombre, resultado in reporte['validaciones'].items():
            f.write(f"VALIDACIÓN: {nombre}\n")
            f.write(f"Resultado: {resultado['resultado']}\n")

            if resultado['errores']:
                f.write("Errores:\n")
                for error in resultado['errores']:
                    f.write(f"  - {error}\n")

            f.write("\n")

    print(f"📄 Log guardado en: {path_log}")
    return str(path_log)

# ==============================================================================
# AUDITORÍA INTERACTIVA
# ==============================================================================

def auditoria_interactiva(archivo_excel):
    """Ejecuta auditoría con preguntas interactivas"""
    print("\n🔍 AUDITORÍA INTERACTIVA")
    print("=" * 70)

    # Preguntar qué validar
    print("\n¿Qué deseas validar?")
    print("1. Todo (recomendado)")
    print("2. Solo estructura")
    print("3. Solo CxP/CxC")
    print("4. Solo formatos")

    opcion = input("\nOpción (1-4): ").strip()

    if opcion == '1':
        aprobado, reporte = auditoria_completa(archivo_excel)
        imprimir_reporte_auditoria(reporte)
        return aprobado, reporte

    # Implementar otras opciones según necesidad
    print("⚠️  Solo opción 1 implementada por ahora")
    return False, {}

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Auditar archivo Excel del sistema de finanzas v4.0',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python auditoria.py --archivo v4.0.xlsx
  python auditoria.py --archivo v4.0.xlsx --verbose
  python auditoria.py --archivo v4.0.xlsx --guardar-log
  python auditoria.py --archivo v4.0.xlsx --interactivo
        """
    )

    parser.add_argument('--archivo', required=True,
                        help='Archivo Excel a auditar')
    parser.add_argument('--verbose', action='store_true',
                        help='Mostrar información detallada')
    parser.add_argument('--guardar-log', action='store_true',
                        help='Guardar reporte en archivo de log')
    parser.add_argument('--interactivo', action='store_true',
                        help='Modo interactivo con opciones')

    args = parser.parse_args()

    # Verificar que archivo existe
    if not Path(args.archivo).exists():
        print(f"❌ ERROR: Archivo no encontrado: {args.archivo}")
        sys.exit(1)

    # Ejecutar auditoría
    try:
        if args.interactivo:
            aprobado, reporte = auditoria_interactiva(args.archivo)
        else:
            aprobado, reporte = auditoria_completa(args.archivo)
            imprimir_reporte_auditoria(reporte)

        # Guardar log si se solicitó
        if args.guardar_log and reporte:
            guardar_log_auditoria(reporte)

        # Verbose
        if args.verbose and reporte:
            print("\n📊 DETALLE COMPLETO:")
            print("=" * 70)
            for nombre, resultado in reporte['validaciones'].items():
                print(f"\n{nombre}:")
                print(f"  Resultado: {resultado['resultado']}")
                if resultado['errores']:
                    print("  Errores:")
                    for error in resultado['errores']:
                        print(f"    - {error}")

        # Exit code
        if aprobado:
            print("\n✅ AUDITORÍA APROBADA")
            sys.exit(0)
        else:
            print("\n❌ AUDITORÍA FALLIDA - Revisar errores arriba")
            sys.exit(1)

    except Exception as e:
        print(f"\n❌ ERROR DURANTE AUDITORÍA: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
