# -*- coding: utf-8 -*-
"""
🔄 SCRIPT DE MIGRACIÓN v3.0 → v4.0 - Sistema de Finanzas
==========================================================

Migra datos de un archivo v3.0 (potencialmente corrupto) a v4.0 limpio.

FUNCIONALIDADES:
- Lee v3.0 con columnas potencialmente desalineadas
- Mapea campos correctamente
- Limpia datos inválidos
- Elimina duplicados
- Valida integridad después de migrar
- Genera reporte detallado

LECCIÓN v3.0: No asumir estructura, mapear dinámicamente
"""

import openpyxl
from datetime import datetime
from pathlib import Path
import sys

from config import (
    ESTRUCTURA_TRANSACCIONES,
    get_columna_por_nombre,
    get_formato_campo
)
from validaciones import (
    auditoria_completa,
    imprimir_reporte_auditoria
)
from operaciones import crear_respaldo

# ==============================================================================
# MAPEO DE COLUMNAS v3.0
# ==============================================================================

def mapear_columnas_v3(ws_v3):
    """
    Lee los headers de v3.0 y crea un mapa de columnas

    CRÍTICO: Esto evita el error de columnas hardcodeadas
    """
    headers_v3 = {}

    for col in range(1, 30):  # Buscar hasta columna 30
        header = ws_v3.cell(1, col).value

        if header:
            # Normalizar nombre (eliminar espacios extra, capitalizar)
            header_normalizado = str(header).strip()
            headers_v3[header_normalizado] = col

    print(f"  📋 Headers detectados en v3.0: {len(headers_v3)}")
    for nombre, col in sorted(headers_v3.items(), key=lambda x: x[1]):
        print(f"     Col {col:2d}: {nombre}")

    return headers_v3

def mapear_campo_v3_a_v4(nombre_campo_v4, headers_v3):
    """
    Encuentra el índice de columna en v3.0 para un campo de v4.0

    Maneja variaciones de nombres (ej: "Fecha Vencimiento" vs "Fecha de Vencimiento")
    """
    # Mapeo exacto
    if nombre_campo_v4 in headers_v3:
        return headers_v3[nombre_campo_v4]

    # Mapeo con variaciones
    variaciones = {
        'Método Pago': ['Metodo Pago', 'Método de Pago', 'Metodo de Pago', 'Método'],
        'Fecha Vencimiento': ['Fecha de Vencimiento', 'Vencimiento', 'Fecha Venc'],
        'Número Factura': ['Numero Factura', 'Número de Factura', 'No. Factura', 'Factura #', 'Factura'],
        'Categoría': ['Categoria'],
        'Subcategoría': ['Subcategoria'],
        'Descripción': ['Descripcion', 'Detalle']
    }

    if nombre_campo_v4 in variaciones:
        for variacion in variaciones[nombre_campo_v4]:
            if variacion in headers_v3:
                print(f"  🔄 Mapeando '{variacion}' → '{nombre_campo_v4}'")
                return headers_v3[variacion]

    return None

# ==============================================================================
# VALIDACIÓN Y LIMPIEZA DE DATOS
# ==============================================================================

def validar_fila_transaccion(fila_data):
    """
    Valida que una fila tenga datos mínimos para ser migrada

    Returns:
        (bool, str): (es_valida, razon_si_invalida)
    """
    # Campos críticos
    if not fila_data.get('Fecha'):
        return False, "Falta fecha"

    if not fila_data.get('Tipo'):
        return False, "Falta tipo"

    if not fila_data.get('Descripción'):
        return False, "Falta descripción"

    if not fila_data.get('Monto'):
        return False, "Falta monto"

    if not fila_data.get('Estado'):
        return False, "Falta estado"

    # Validar tipo
    if fila_data['Tipo'] not in ['INGRESO', 'EGRESO', 'TRANSFERENCIA']:
        return False, f"Tipo inválido: {fila_data['Tipo']}"

    # Validar estado
    estados_validos = ['PAGADO', 'PENDIENTE', 'POR COBRAR', 'CANCELADO', 'PARCIAL']
    if fila_data['Estado'] not in estados_validos:
        return False, f"Estado inválido: {fila_data['Estado']}"

    return True, ""

def limpiar_valor(valor, tipo_campo):
    """Limpia y convierte un valor según su tipo"""
    if valor is None:
        return None

    # String
    if tipo_campo == 'string':
        return str(valor).strip() if valor else None

    # Float
    if tipo_campo == 'float':
        try:
            return float(valor)
        except (ValueError, TypeError):
            return None

    # Datetime
    if tipo_campo == 'datetime':
        if isinstance(valor, datetime):
            return valor
        if isinstance(valor, str):
            # Intentar parsear varios formatos
            formatos = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d']
            for formato in formatos:
                try:
                    return datetime.strptime(valor, formato)
                except ValueError:
                    continue
        return None

    return valor

# ==============================================================================
# MIGRACIÓN PRINCIPAL
# ==============================================================================

def migrar_v3_a_v4(archivo_v3, archivo_v4, modo='test'):
    """
    Migra transacciones de v3.0 a v4.0

    Args:
        archivo_v3: Path al archivo v3.0 (origen)
        archivo_v4: Path al archivo v4.0 (destino, debe existir)
        modo: 'test' = no guardar, solo reportar | 'produccion' = guardar cambios

    Returns:
        dict: Reporte de migración
    """
    print(f"""
╔═══════════════════════════════════════════════════════════════════╗
║              MIGRACIÓN v3.0 → v4.0 - INICIO                       ║
╠═══════════════════════════════════════════════════════════════════╣
║ Origen: {archivo_v3:<57} ║
║ Destino: {archivo_v4:<56} ║
║ Modo: {modo.upper():<59} ║
╚═══════════════════════════════════════════════════════════════════╝
    """)

    reporte = {
        'origen': archivo_v3,
        'destino': archivo_v4,
        'modo': modo,
        'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'filas_leidas': 0,
        'filas_validas': 0,
        'filas_invalidas': 0,
        'filas_duplicadas': 0,
        'filas_migradas': 0,
        'errores': [],
        'advertencias': []
    }

    # Verificar que archivos existen
    if not Path(archivo_v3).exists():
        reporte['errores'].append(f"❌ Archivo origen no existe: {archivo_v3}")
        return reporte

    if not Path(archivo_v4).exists():
        reporte['errores'].append(f"❌ Archivo destino no existe: {archivo_v4}")
        reporte['advertencias'].append("💡 Ejecuta 'python generar_v4.py' primero")
        return reporte

    # Crear respaldo de v4.0 antes de modificar
    if modo == 'produccion':
        print("\n💾 Creando respaldo de v4.0...")
        crear_respaldo(archivo_v4, 'antes_migracion')

    # Abrir archivos
    print("\n📖 Leyendo archivos...")
    wb_v3 = openpyxl.load_workbook(archivo_v3, data_only=True)
    wb_v4 = openpyxl.load_workbook(archivo_v4)

    ws_v3 = wb_v3['TRANSACCIONES']
    ws_v4 = wb_v4['TRANSACCIONES']

    # Mapear columnas v3
    print("\n🗺️  Mapeando estructura de v3.0...")
    headers_v3 = mapear_columnas_v3(ws_v3)

    # Leer transacciones de v3
    print("\n📊 Leyendo transacciones de v3.0...")
    transacciones_v3 = []
    transacciones_vistas = set()  # Para detectar duplicados

    for fila_v3 in range(2, ws_v3.max_row + 1):
        reporte['filas_leidas'] += 1

        # Leer datos de la fila
        fila_data = {}

        for nombre_campo_v4, config_v4 in ESTRUCTURA_TRANSACCIONES.items():
            # Encontrar columna en v3
            col_v3 = mapear_campo_v3_a_v4(nombre_campo_v4, headers_v3)

            if col_v3:
                valor_crudo = ws_v3.cell(fila_v3, col_v3).value
                valor_limpio = limpiar_valor(valor_crudo, config_v4.get('tipo', 'string'))
                fila_data[nombre_campo_v4] = valor_limpio
            else:
                fila_data[nombre_campo_v4] = None

        # Validar fila
        es_valida, razon = validar_fila_transaccion(fila_data)

        if not es_valida:
            reporte['filas_invalidas'] += 1
            reporte['advertencias'].append(f"⚠️  Fila {fila_v3}: {razon} - SALTADA")
            continue

        # Detectar duplicados (misma fecha + descripción + monto)
        clave = (
            str(fila_data.get('Fecha')),
            str(fila_data.get('Descripción')),
            str(fila_data.get('Monto'))
        )

        if clave in transacciones_vistas:
            reporte['filas_duplicadas'] += 1
            reporte['advertencias'].append(
                f"⚠️  Fila {fila_v3}: Duplicado detectado - SALTADA"
            )
            continue

        transacciones_vistas.add(clave)
        reporte['filas_validas'] += 1
        transacciones_v3.append(fila_data)

    print(f"  ✅ {reporte['filas_validas']} transacciones válidas encontradas")
    print(f"  ⚠️  {reporte['filas_invalidas']} filas inválidas")
    print(f"  ⚠️  {reporte['filas_duplicadas']} duplicados")

    # Escribir a v4
    print("\n✍️  Escribiendo transacciones en v4.0...")

    for fila_data in transacciones_v3:
        fila_v4 = ws_v4.max_row + 1

        for nombre_campo, config in ESTRUCTURA_TRANSACCIONES.items():
            col_v4 = config['col']
            valor = fila_data.get(nombre_campo)

            # Escribir valor
            celda = ws_v4.cell(fila_v4, col_v4, valor)

            # Aplicar formato
            formato = get_formato_campo(nombre_campo)
            if formato and valor:
                celda.number_format = formato

        reporte['filas_migradas'] += 1

    print(f"  ✅ {reporte['filas_migradas']} transacciones migradas a v4.0")

    # Guardar o cerrar según modo
    if modo == 'produccion':
        print("\n💾 Guardando archivo v4.0...")
        wb_v4.save(archivo_v4)
        print("  ✅ Archivo guardado")
    else:
        print("\n⚠️  MODO TEST: No se guardaron cambios")

    wb_v3.close()
    wb_v4.close()

    # Validar integridad después de migrar
    if modo == 'produccion':
        print("\n🔍 Validando integridad de v4.0 después de migración...")
        aprobado, reporte_auditoria = auditoria_completa(archivo_v4)

        if aprobado:
            print("  ✅ Validación exitosa")
        else:
            print("  ❌ Validación falló")
            reporte['errores'].append("Validación post-migración falló")

        reporte['auditoria'] = reporte_auditoria

    return reporte

def imprimir_reporte_migracion(reporte):
    """Imprime el reporte de migración de forma legible"""
    print("\n╔═══════════════════════════════════════════════════════════════════╗")
    print("║                   REPORTE DE MIGRACIÓN v3.0 → v4.0                ║")
    print("╠═══════════════════════════════════════════════════════════════════╣")
    print(f"║ Fecha: {reporte['fecha']:<57} ║")
    print(f"║ Modo: {reporte['modo'].upper():<59} ║")
    print("╠═══════════════════════════════════════════════════════════════════╣")
    print(f"║ Filas leídas: {reporte['filas_leidas']:<51} ║")
    print(f"║ Filas válidas: {reporte['filas_validas']:<50} ║")
    print(f"║ Filas inválidas: {reporte['filas_invalidas']:<48} ║")
    print(f"║ Duplicados encontrados: {reporte['filas_duplicadas']:<41} ║")
    print(f"║ Filas migradas: {reporte['filas_migradas']:<47} ║")
    print("╠═══════════════════════════════════════════════════════════════════╣")

    if reporte['errores']:
        print("║ ERRORES:                                                          ║")
        for error in reporte['errores'][:5]:
            print(f"║  {error:<64} ║")

    if reporte['advertencias']:
        print("║ ADVERTENCIAS (primeras 5):                                        ║")
        for adv in reporte['advertencias'][:5]:
            print(f"║  {adv:<64} ║")

    print("╠═══════════════════════════════════════════════════════════════════╣")

    if reporte.get('auditoria'):
        aud = reporte['auditoria']
        if aud['aprobado']:
            print("║ Estado: ✅ MIGRACIÓN EXITOSA Y VALIDADA                           ║")
        else:
            print("║ Estado: ⚠️  MIGRACIÓN COMPLETA PERO CON ADVERTENCIAS              ║")
    else:
        print("║ Estado: ⚠️  MODO TEST - NO SE GUARDARON CAMBIOS                   ║")

    print("╚═══════════════════════════════════════════════════════════════════╝\n")

# ==============================================================================
# EJECUCIÓN PRINCIPAL
# ==============================================================================

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Migrar datos de v3.0 a v4.0')
    parser.add_argument('--origen', required=True, help='Archivo v3.0 (origen)')
    parser.add_argument('--destino', required=True, help='Archivo v4.0 (destino)')
    parser.add_argument('--modo', default='test', choices=['test', 'produccion'],
                        help='test = no guardar, produccion = guardar cambios')

    args = parser.parse_args()

    try:
        # Ejecutar migración
        reporte = migrar_v3_a_v4(args.origen, args.destino, args.modo)

        # Imprimir reporte
        imprimir_reporte_migracion(reporte)

        # Imprimir auditoría si existe
        if reporte.get('auditoria'):
            imprimir_reporte_auditoria(reporte['auditoria'])

        # Exit code
        if reporte['errores']:
            sys.exit(1)
        else:
            sys.exit(0)

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
