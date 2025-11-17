#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Herramienta de diagnóstico de referencias circulares en Excel
Lee el archivo Excel generado y encuentra EXACTAMENTE dónde están las referencias circulares
"""

import openpyxl
from openpyxl.formula import Tokenizer
import re
from collections import defaultdict, deque

def extraer_referencias(formula, hoja_actual):
    """Extrae todas las referencias de celda de una fórmula."""
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return []

    referencias = []

    # Buscar referencias con hoja (ej: TRANSACCIONES!A3:A1000)
    patron_con_hoja = re.findall(r'([A-Z_]+)!([A-Z]+\d+(?::[A-Z]+\d+)?)', formula)
    for hoja, rango in patron_con_hoja:
        referencias.append(f"{hoja}!{rango}")

    # Buscar referencias sin hoja (ej: A3, B5:C10)
    patron_sin_hoja = re.findall(r'\b([A-Z]+\d+(?::[A-Z]+\d+)?)\b', formula)
    for rango in patron_sin_hoja:
        # Verificar que no sea parte de una referencia con hoja
        if not any(rango in ref for ref in referencias):
            referencias.append(f"{hoja_actual}!{rango}")

    return referencias

def expandir_rango(rango_str):
    """Expande un rango A3:A10 a lista de celdas individuales (limitado a primeras 10)."""
    if ':' not in rango_str:
        return [rango_str]

    # Extraer hoja si existe
    if '!' in rango_str:
        hoja, rango = rango_str.split('!')
        hoja += '!'
    else:
        hoja = ''
        rango = rango_str

    inicio, fin = rango.split(':')

    # Extraer columna y fila
    col_inicio = re.match(r'([A-Z]+)', inicio).group(1)
    fila_inicio = int(re.match(r'[A-Z]+(\d+)', inicio).group(1))
    col_fin = re.match(r'([A-Z]+)', fin).group(1)
    fila_fin = int(re.match(r'[A-Z]+(\d+)', fin).group(1))

    # Limitar a primeras 10 celdas para evitar explosión de memoria
    celdas = []
    if col_inicio == col_fin:
        # Rango vertical (A3:A1000)
        for fila in range(fila_inicio, min(fila_inicio + 10, fila_fin + 1)):
            celdas.append(f"{hoja}{col_inicio}{fila}")
    elif fila_inicio == fila_fin:
        # Rango horizontal (A3:Z3)
        col_actual = col_inicio
        while col_actual <= col_fin and len(celdas) < 10:
            celdas.append(f"{hoja}{col_actual}{fila_inicio}")
            col_actual = chr(ord(col_actual) + 1)

    return celdas if celdas else [f"{hoja}{inicio}"]

def analizar_excel(archivo):
    """Analiza el Excel y encuentra referencias circulares."""
    print(f"\n{'='*70}")
    print(f"ANÁLISIS DE REFERENCIAS CIRCULARES: {archivo}")
    print(f"{'='*70}\n")

    wb = openpyxl.load_workbook(archivo, data_only=False)

    # Diccionario: celda -> lista de celdas que referencia
    dependencias = defaultdict(list)
    formulas = {}

    print("📊 Extrayendo fórmulas de todas las hojas...\n")

    for hoja_nombre in wb.sheetnames:
        ws = wb[hoja_nombre]
        print(f"  Analizando hoja: {hoja_nombre}")

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    celda_id = f"{hoja_nombre}!{cell.coordinate}"
                    formula = cell.value
                    formulas[celda_id] = formula

                    # Extraer referencias
                    refs = extraer_referencias(formula, hoja_nombre)

                    # Expandir rangos (limitado)
                    refs_expandidas = []
                    for ref in refs:
                        refs_expandidas.extend(expandir_rango(ref))

                    dependencias[celda_id] = refs_expandidas

    print(f"\n✓ Total de celdas con fórmulas: {len(formulas)}")
    print(f"✓ Total de dependencias: {sum(len(deps) for deps in dependencias.values())}")

    # Buscar ciclos usando DFS
    print(f"\n{'='*70}")
    print("🔍 BUSCANDO REFERENCIAS CIRCULARES...")
    print(f"{'='*70}\n")

    visitados = set()
    en_proceso = set()
    ciclos_encontrados = []

    def dfs(celda, camino):
        """DFS para detectar ciclos."""
        if celda in en_proceso:
            # Encontramos un ciclo
            idx = camino.index(celda)
            ciclo = camino[idx:] + [celda]
            ciclos_encontrados.append(ciclo)
            return True

        if celda in visitados:
            return False

        en_proceso.add(celda)
        camino.append(celda)

        for dependencia in dependencias.get(celda, []):
            dfs(dependencia, camino[:])

        en_proceso.remove(celda)
        visitados.add(celda)
        return False

    # Ejecutar DFS desde cada celda con fórmula
    for celda in formulas.keys():
        if celda not in visitados:
            dfs(celda, [])

    # Mostrar resultados
    if ciclos_encontrados:
        print(f"❌ ENCONTRADAS {len(ciclos_encontrados)} REFERENCIAS CIRCULARES:\n")

        for i, ciclo in enumerate(ciclos_encontrados[:10], 1):  # Mostrar primeras 10
            print(f"CICLO #{i}:")
            print(f"  Longitud: {len(ciclo) - 1} pasos")
            print(f"  Camino: {' → '.join(ciclo)}")

            # Mostrar fórmulas involucradas
            print(f"  Fórmulas:")
            for celda in ciclo[:-1]:  # Excluir última (es repetida)
                if celda in formulas:
                    formula_corta = formulas[celda][:80] + '...' if len(formulas[celda]) > 80 else formulas[celda]
                    print(f"    {celda}: {formula_corta}")
            print()

        if len(ciclos_encontrados) > 10:
            print(f"  ... y {len(ciclos_encontrados) - 10} ciclos más\n")

        return False
    else:
        print("✅ NO SE ENCONTRARON REFERENCIAS CIRCULARES\n")
        print("El archivo Excel NO tiene referencias circulares detectables.")
        print("El problema puede ser:")
        print("  1. Fórmulas que Excel interpreta como circulares pero el script no")
        print("  2. Fórmulas volátiles (TODAY, NOW) que causan recálculos")
        print("  3. Configuración de cálculo de Excel")
        print()
        return True

def mostrar_resumen_formulas(archivo):
    """Muestra un resumen de fórmulas por hoja."""
    print(f"\n{'='*70}")
    print("📋 RESUMEN DE FÓRMULAS POR HOJA")
    print(f"{'='*70}\n")

    wb = openpyxl.load_workbook(archivo, data_only=False)

    for hoja_nombre in wb.sheetnames:
        ws = wb[hoja_nombre]
        formulas_hoja = []

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    formulas_hoja.append((cell.coordinate, cell.value))

        if formulas_hoja:
            print(f"Hoja: {hoja_nombre}")
            print(f"  Total fórmulas: {len(formulas_hoja)}")
            print(f"  Primeras 5 fórmulas:")
            for coord, formula in formulas_hoja[:5]:
                formula_corta = formula[:60] + '...' if len(formula) > 60 else formula
                print(f"    {coord}: {formula_corta}")
            print()

if __name__ == "__main__":
    archivo = "AlvaroVelasco_Finanzas_v5.0_COMPLETO.xlsx"

    # Análisis de referencias circulares
    resultado = analizar_excel(archivo)

    # Resumen de fórmulas
    mostrar_resumen_formulas(archivo)

    print(f"{'='*70}")
    print("ANÁLISIS COMPLETADO")
    print(f"{'='*70}\n")
