# -*- coding: utf-8 -*-
"""
🔐 MÓDULO DE GESTIÓN DE ALIAS - v4.0
=====================================

Gestiona el sistema de alias para cuentas bancarias enmascaradas.

FUNCIONALIDAD RESCATADA DE v3.0:
- 48 alias registrados en v3.0
- 11 alias de cuentas bancarias
- 7 alias BNCR (2 enmascarados)
- Mapeo automático de cuentas enmascaradas (XXXXXXXXXX1066X → BNCR CC Colones)

CASOS DE USO:
1. Importación de CSV bancario con cuentas enmascaradas
2. Conciliación bancaria automática
3. Reportes ejecutivos legibles

LECCIÓN APRENDIDA v3.0:
- ✅ Sistema de alias FUNCIONA correctamente
- ✅ Crítico para reportes legibles
- ✅ Evita duplicados de cuentas
- ⚠️ DEBE ser gestionado de forma centralizada (no edición manual)
"""

import openpyxl
from datetime import datetime
from config import ESTRUCTURA_ENTIDADES_ALIAS, HOJAS_CONFIG
import os

# ==============================================================================
# DATOS DE ALIAS PRE-CARGADOS DE v3.0
# ==============================================================================

ALIAS_PRECARGADOS_V3 = [
    # Cuentas BNCR enmascaradas (críticas)
    {
        'Alias': 'XXXXXXXXXX1066X',
        'Entidad Real': 'BNCR Cuenta Corriente Colones',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta corriente BNCR en colones (enmascarada)'
    },
    {
        'Alias': 'XXXXXXXXXX8618X',
        'Entidad Real': 'BNCR Cuenta Corriente Dólares',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta corriente BNCR en dólares (enmascarada)'
    },
    # Cuentas BNCR parcialmente visibles
    {
        'Alias': 'BNCR CRC Ahorros (***8618)',
        'Entidad Real': 'BNCR Ahorros Colones',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta de ahorros BNCR en colones'
    },
    {
        'Alias': 'BNCR USD Ahorros (***1066)',
        'Entidad Real': 'BNCR Ahorros Dólares',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta de ahorros BNCR en dólares'
    },
    {
        'Alias': 'BNCR CRC Corriente (***2186)',
        'Entidad Real': 'BNCR Cuenta Corriente Colones 2186',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta corriente BNCR en colones'
    },
    {
        'Alias': 'BNCR USD Corriente (***9589)',
        'Entidad Real': 'BNCR Cuenta Corriente Dólares 9589',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta corriente BNCR en dólares'
    },
    {
        'Alias': 'BNCR USD Corriente (***1112)',
        'Entidad Real': 'BNCR Cuenta Corriente Dólares 1112',
        'Tipo': 'Cuenta',
        'Notas': 'Cuenta corriente BNCR en dólares (alternativa)'
    },
    # Tarjetas de crédito comunes
    {
        'Alias': 'XXXX-XXXX-XXXX-1234',
        'Entidad Real': 'Tarjeta Visa BAC 1234',
        'Tipo': 'Tarjeta',
        'Notas': 'Tarjeta de crédito BAC Visa'
    },
    {
        'Alias': 'XXXX-XXXX-XXXX-5678',
        'Entidad Real': 'Tarjeta MasterCard BAC 5678',
        'Tipo': 'Tarjeta',
        'Notas': 'Tarjeta de crédito BAC MasterCard'
    },
    # Proveedores comunes (ejemplos)
    {
        'Alias': 'ICE',
        'Entidad Real': 'Instituto Costarricense de Electricidad',
        'Tipo': 'Proveedor',
        'Notas': 'Servicios eléctricos y telecomunicaciones'
    },
    {
        'Alias': 'KOLBI',
        'Entidad Real': 'Instituto Costarricense de Electricidad - Kolbi',
        'Tipo': 'Proveedor',
        'Notas': 'Servicios de telecomunicaciones móviles'
    }
]

# ==============================================================================
# FUNCIONES PRINCIPALES
# ==============================================================================

def resolver_alias(nombre_entrada, archivo_excel=None):
    """
    Resuelve un alias enmascarado a su nombre real.

    Prioridad de búsqueda:
    1. Coincidencia exacta de alias
    2. Coincidencia parcial (últimos 4 dígitos)
    3. Nombre original si no se encuentra

    Args:
        nombre_entrada (str): Nombre/alias a resolver
        archivo_excel (str, optional): Ruta al archivo Excel. Si no se proporciona,
                                       usa solo alias pre-cargados.

    Returns:
        str: Nombre real de la entidad

    Ejemplo:
        >>> resolver_alias('XXXXXXXXXX1066X')
        'BNCR Cuenta Corriente Colones'

        >>> resolver_alias('BNCR CRC Ahorros (***8618)')
        'BNCR Ahorros Colones'
    """
    if not nombre_entrada or nombre_entrada.strip() == '':
        return nombre_entrada

    nombre_entrada = nombre_entrada.strip()

    # 1. Buscar en alias pre-cargados (coincidencia exacta)
    for alias_info in ALIAS_PRECARGADOS_V3:
        if alias_info['Alias'].upper() == nombre_entrada.upper():
            return alias_info['Entidad Real']

    # 2. Si hay archivo Excel, buscar allí
    if archivo_excel and os.path.exists(archivo_excel):
        try:
            wb = openpyxl.load_workbook(archivo_excel, data_only=True)
            if 'ENTIDADES_ALIAS' in wb.sheetnames:
                ws = wb['ENTIDADES_ALIAS']

                # Buscar coincidencia exacta
                for fila in range(2, ws.max_row + 1):
                    alias = ws.cell(fila, 1).value
                    entidad_real = ws.cell(fila, 2).value

                    if alias and alias.upper() == nombre_entrada.upper():
                        return entidad_real if entidad_real else nombre_entrada

                # Buscar coincidencia parcial (últimos 4 dígitos)
                ultimos_4 = nombre_entrada[-4:] if len(nombre_entrada) >= 4 else None
                if ultimos_4:
                    for fila in range(2, ws.max_row + 1):
                        alias = ws.cell(fila, 1).value
                        entidad_real = ws.cell(fila, 2).value

                        if alias and ultimos_4 in alias:
                            return entidad_real if entidad_real else nombre_entrada

            wb.close()
        except Exception as e:
            print(f"⚠️ Error al buscar alias en Excel: {e}")

    # 3. No se encontró, retornar nombre original
    return nombre_entrada


def agregar_alias(archivo_excel, alias, entidad_real, tipo='Cuenta', notas=''):
    """
    Agrega un nuevo alias a la hoja ENTIDADES_ALIAS.

    Args:
        archivo_excel (str): Ruta al archivo Excel
        alias (str): Alias enmascarado
        entidad_real (str): Nombre real de la entidad
        tipo (str): Tipo de entidad (Cuenta, Tarjeta, Proveedor, Cliente, Otro)
        notas (str): Notas adicionales

    Returns:
        (bool, str): (éxito, mensaje)

    VALIDACIONES:
    - Archivo debe existir
    - Hoja ENTIDADES_ALIAS debe existir
    - Alias no debe estar duplicado
    - Tipo debe ser válido

    Ejemplo:
        >>> agregar_alias('finanzas_v4.xlsx', 'XXXX1234', 'BAC CC Dólares', 'Cuenta', 'Nueva cuenta')
        (True, "Alias agregado exitosamente en fila 50")
    """
    # Validaciones
    if not os.path.exists(archivo_excel):
        return False, f"Archivo {archivo_excel} no existe"

    tipos_validos = ESTRUCTURA_ENTIDADES_ALIAS['Tipo']['valores_validos']
    if tipo not in tipos_validos:
        return False, f"Tipo '{tipo}' no válido. Válidos: {tipos_validos}"

    if not alias or not entidad_real:
        return False, "Alias y Entidad Real son obligatorios"

    try:
        wb = openpyxl.load_workbook(archivo_excel)

        if 'ENTIDADES_ALIAS' not in wb.sheetnames:
            return False, "Hoja ENTIDADES_ALIAS no existe en el archivo"

        ws = wb['ENTIDADES_ALIAS']

        # Verificar si alias ya existe
        for fila in range(2, ws.max_row + 1):
            alias_existente = ws.cell(fila, 1).value
            if alias_existente and alias_existente.upper() == alias.upper():
                return False, f"Alias '{alias}' ya existe en fila {fila}"

        # Encontrar primera fila vacía
        fila_nueva = ws.max_row + 1

        # Insertar datos
        ws.cell(fila_nueva, 1, alias)
        ws.cell(fila_nueva, 2, entidad_real)
        ws.cell(fila_nueva, 3, tipo)
        ws.cell(fila_nueva, 4, notas)

        # Guardar
        wb.save(archivo_excel)
        wb.close()

        return True, f"Alias agregado exitosamente en fila {fila_nueva}"

    except Exception as e:
        return False, f"Error al agregar alias: {str(e)}"


def listar_alias(archivo_excel=None, solo_tipo=None):
    """
    Lista todos los alias registrados.

    Args:
        archivo_excel (str, optional): Ruta al archivo Excel
        solo_tipo (str, optional): Filtrar por tipo (Cuenta, Tarjeta, etc.)

    Returns:
        list: Lista de diccionarios con información de alias

    Ejemplo:
        >>> listar_alias(solo_tipo='Cuenta')
        [
            {
                'Alias': 'XXXXXXXXXX1066X',
                'Entidad Real': 'BNCR Cuenta Corriente Colones',
                'Tipo': 'Cuenta',
                'Notas': 'Cuenta corriente BNCR en colones (enmascarada)'
            },
            ...
        ]
    """
    alias_list = []

    # 1. Agregar alias pre-cargados
    for alias_info in ALIAS_PRECARGADOS_V3:
        if solo_tipo is None or alias_info['Tipo'] == solo_tipo:
            alias_list.append(alias_info.copy())

    # 2. Si hay archivo Excel, agregar de allí
    if archivo_excel and os.path.exists(archivo_excel):
        try:
            wb = openpyxl.load_workbook(archivo_excel, data_only=True)
            if 'ENTIDADES_ALIAS' in wb.sheetnames:
                ws = wb['ENTIDADES_ALIAS']

                for fila in range(2, ws.max_row + 1):
                    alias = ws.cell(fila, 1).value
                    entidad_real = ws.cell(fila, 2).value
                    tipo = ws.cell(fila, 3).value
                    notas = ws.cell(fila, 4).value

                    if alias:  # Solo si hay alias
                        # Evitar duplicados
                        ya_existe = any(a['Alias'].upper() == alias.upper() for a in alias_list)
                        if not ya_existe:
                            if solo_tipo is None or tipo == solo_tipo:
                                alias_list.append({
                                    'Alias': alias,
                                    'Entidad Real': entidad_real,
                                    'Tipo': tipo,
                                    'Notas': notas
                                })

            wb.close()
        except Exception as e:
            print(f"⚠️ Error al listar alias de Excel: {e}")

    return alias_list


def validar_alias(archivo_excel):
    """
    Valida la integridad de la hoja ENTIDADES_ALIAS.

    Validaciones:
    - Hoja existe
    - Estructura correcta (4 columnas)
    - No hay duplicados
    - Tipos válidos
    - Campos obligatorios no vacíos

    Args:
        archivo_excel (str): Ruta al archivo Excel

    Returns:
        (bool, list): (es_valido, lista_errores)

    Ejemplo:
        >>> validar_alias('finanzas_v4.xlsx')
        (True, [])

        >>> validar_alias('finanzas_corrupto.xlsx')
        (False, ['Fila 10: Alias vacío', 'Fila 15: Tipo inválido'])
    """
    errores = []

    if not os.path.exists(archivo_excel):
        return False, [f"Archivo {archivo_excel} no existe"]

    try:
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)

        if 'ENTIDADES_ALIAS' not in wb.sheetnames:
            return False, ["Hoja ENTIDADES_ALIAS no existe"]

        ws = wb['ENTIDADES_ALIAS']

        # Validar estructura (headers)
        headers_esperados = ['Alias', 'Entidad Real', 'Tipo', 'Notas']
        for col, header_esperado in enumerate(headers_esperados, start=1):
            header_real = ws.cell(1, col).value
            if header_real != header_esperado:
                errores.append(f"Columna {col}: Se esperaba '{header_esperado}', encontrado '{header_real}'")

        # Validar datos
        alias_vistos = set()
        tipos_validos = ESTRUCTURA_ENTIDADES_ALIAS['Tipo']['valores_validos']

        for fila in range(2, ws.max_row + 1):
            alias = ws.cell(fila, 1).value
            entidad_real = ws.cell(fila, 2).value
            tipo = ws.cell(fila, 3).value

            # Saltar filas completamente vacías
            if not any([alias, entidad_real, tipo]):
                continue

            # Validar campos obligatorios
            if not alias:
                errores.append(f"Fila {fila}: Alias vacío")
            if not entidad_real:
                errores.append(f"Fila {fila}: Entidad Real vacía")
            if not tipo:
                errores.append(f"Fila {fila}: Tipo vacío")

            # Validar duplicados
            if alias:
                alias_upper = alias.upper()
                if alias_upper in alias_vistos:
                    errores.append(f"Fila {fila}: Alias '{alias}' duplicado")
                alias_vistos.add(alias_upper)

            # Validar tipo
            if tipo and tipo not in tipos_validos:
                errores.append(f"Fila {fila}: Tipo '{tipo}' no válido. Válidos: {tipos_validos}")

        wb.close()

        if errores:
            return False, errores
        else:
            return True, []

    except Exception as e:
        return False, [f"Error al validar: {str(e)}"]


def cargar_alias_precargados(archivo_excel):
    """
    Carga los alias pre-cargados de v3.0 en la hoja ENTIDADES_ALIAS.

    Esta función debe ejecutarse solo una vez al crear el archivo inicial.
    NO debe ejecutarse si ya hay alias cargados.

    Args:
        archivo_excel (str): Ruta al archivo Excel

    Returns:
        (bool, str): (éxito, mensaje)

    Ejemplo:
        >>> cargar_alias_precargados('finanzas_v4.xlsx')
        (True, "11 alias pre-cargados exitosamente")
    """
    if not os.path.exists(archivo_excel):
        return False, f"Archivo {archivo_excel} no existe"

    try:
        wb = openpyxl.load_workbook(archivo_excel)

        if 'ENTIDADES_ALIAS' not in wb.sheetnames:
            return False, "Hoja ENTIDADES_ALIAS no existe"

        ws = wb['ENTIDADES_ALIAS']

        # Verificar si ya hay datos (evitar duplicados)
        if ws.max_row > 1:
            # Ya hay datos
            alias_existentes = []
            for fila in range(2, ws.max_row + 1):
                alias = ws.cell(fila, 1).value
                if alias:
                    alias_existentes.append(alias)

            if alias_existentes:
                return False, f"Ya hay {len(alias_existentes)} alias cargados. No se cargarán duplicados."

        # Cargar alias pre-cargados
        fila_actual = 2
        for alias_info in ALIAS_PRECARGADOS_V3:
            ws.cell(fila_actual, 1, alias_info['Alias'])
            ws.cell(fila_actual, 2, alias_info['Entidad Real'])
            ws.cell(fila_actual, 3, alias_info['Tipo'])
            ws.cell(fila_actual, 4, alias_info['Notas'])
            fila_actual += 1

        wb.save(archivo_excel)
        wb.close()

        cantidad = len(ALIAS_PRECARGADOS_V3)
        return True, f"{cantidad} alias pre-cargados exitosamente"

    except Exception as e:
        return False, f"Error al cargar alias: {str(e)}"


def buscar_alias_por_ultimos_digitos(ultimos_digitos, archivo_excel=None):
    """
    Busca alias que contengan los últimos dígitos especificados.

    Útil para encontrar cuentas bancarias por sus últimos 4 dígitos.

    Args:
        ultimos_digitos (str): Últimos dígitos a buscar (ej: '1066')
        archivo_excel (str, optional): Ruta al archivo Excel

    Returns:
        list: Lista de alias encontrados

    Ejemplo:
        >>> buscar_alias_por_ultimos_digitos('1066')
        [
            {
                'Alias': 'XXXXXXXXXX1066X',
                'Entidad Real': 'BNCR Cuenta Corriente Colones',
                'Tipo': 'Cuenta',
                'Notas': '...'
            },
            {
                'Alias': 'BNCR USD Ahorros (***1066)',
                'Entidad Real': 'BNCR Ahorros Dólares',
                'Tipo': 'Cuenta',
                'Notas': '...'
            }
        ]
    """
    todos_alias = listar_alias(archivo_excel)
    resultados = []

    for alias_info in todos_alias:
        if ultimos_digitos in alias_info['Alias']:
            resultados.append(alias_info)

    return resultados


# ==============================================================================
# FUNCIÓN DE DEMOSTRACIÓN
# ==============================================================================

def demo_alias():
    """Demostración del sistema de alias"""
    print("=" * 70)
    print("🔐 DEMOSTRACIÓN DEL SISTEMA DE ALIAS v4.0")
    print("=" * 70)

    print("\n1️⃣ Resolución de alias enmascarados:\n")

    tests = [
        'XXXXXXXXXX1066X',
        'XXXXXXXXXX8618X',
        'BNCR CRC Ahorros (***8618)',
        'Cuenta no existente'
    ]

    for test in tests:
        resultado = resolver_alias(test)
        print(f"   Entrada: {test:<30} → Resultado: {resultado}")

    print("\n2️⃣ Listado de alias por tipo:\n")

    for tipo in ['Cuenta', 'Tarjeta', 'Proveedor']:
        alias_tipo = listar_alias(solo_tipo=tipo)
        print(f"   {tipo}: {len(alias_tipo)} alias registrados")

    print("\n3️⃣ Búsqueda por últimos dígitos:\n")

    resultados_1066 = buscar_alias_por_ultimos_digitos('1066')
    print(f"   Buscando '1066': {len(resultados_1066)} coincidencias")
    for r in resultados_1066:
        print(f"      - {r['Alias']} → {r['Entidad Real']}")

    print("\n" + "=" * 70)
    print("✅ Demostración completada")
    print("=" * 70)


if __name__ == '__main__':
    # Si se ejecuta directamente, mostrar demo
    demo_alias()
