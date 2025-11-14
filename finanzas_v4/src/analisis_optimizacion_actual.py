#!/usr/bin/env python3
"""
Análisis de Optimización con Datos Reales
Basado en Cuestionario Fundacional 12/11/2025
"""

from optimizacion import OptimizadorPagos, TarjetaCredito, crear_tarjetas_manual

print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║              🎯 ANÁLISIS DE OPTIMIZACIÓN - ÁLVARO VELASCO                    ║
║                   Basado en Cuestionario Fundacional                          ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")

# ══════════════════════════════════════════════════════════════════════════════
# DATOS DEL CUESTIONARIO FUNDACIONAL (12/11/2025)
# ══════════════════════════════════════════════════════════════════════════════

print("📋 SITUACIÓN FINANCIERA ACTUAL (14/11/2025):\n")

# Situación Operativa del Cuestionario
facturacion_mensual = 9466.42  # USD - Noviembre 2025 real
break_even = 8000  # USD
margen_operativo = facturacion_mensual - break_even  # ~$1,466/mes
efectivo_actual = 3444.54  # USD
dias_cobertura = 12.9  # días

print(f"💰 Facturación mensual (Nov): ${facturacion_mensual:,.2f}")
print(f"📊 Break-even: ${break_even:,.2f}")
print(f"✅ Margen operativo: ${margen_operativo:,.2f}/mes")
print(f"🏦 Efectivo disponible: ${efectivo_actual:,.2f}")
print(f"⏰ Cobertura operativa: {dias_cobertura} días (🔴 CRÍTICO: <15 días)\n")

# Deuda Total del Cuestionario
deuda_tc = 16536  # USD equivalente (datos cuestionario - actualizados hoy ~$15,166)
deuda_nissan = 18680.75
deuda_hacienda = 10215.83
deuda_total = deuda_tc + deuda_nissan + deuda_hacienda

print(f"💳 Deuda tarjetas: ${deuda_tc:,.2f}")
print(f"🚗 Deuda Nissan: ${deuda_nissan:,.2f}")
print(f"🏛️ Deuda Hacienda: ${deuda_hacienda:,.2f}")
print(f"📊 DEUDA TOTAL: ${deuda_total:,.2f}\n")

# Pagos Mensuales Actuales del Cuestionario
pago_tc_actual = 556  # USD - 1.5x mínimo en 4 tarjetas empresa
pago_nissan = 800
pago_hacienda_intereses = 204.32  # Solo intereses (2% mensual)
pagos_mensuales_minimos = pago_tc_actual + pago_nissan + pago_hacienda_intereses

print(f"💳 Pago TC actual (1.5x mínimo): ${pago_tc_actual:,.2f}/mes")
print(f"🚗 Pago Nissan: ${pago_nissan:,.2f}/mes")
print(f"🏛️ Hacienda (solo intereses): ${pago_hacienda_intereses:,.2f}/mes")
print(f"📊 Total pagos actuales: ${pagos_mensuales_minimos:,.2f}/mes\n")

# ALERTA CRÍTICA del Cuestionario
intereses_tc_mensuales = 625  # USD - Del cuestionario
print("🚨 ALERTA CRÍTICA DEL CUESTIONARIO:")
print(f"   Pagos TC actuales: ${pago_tc_actual:,.2f}/mes")
print(f"   Intereses TC generados: ${intereses_tc_mensuales:,.2f}/mes")
print(f"   DIFERENCIA: ${pago_tc_actual - intereses_tc_mensuales:,.2f}/mes")
print("   ❌ ¡ESTÁS PAGANDO MENOS QUE LOS INTERESES!")
print("   ❌ ¡LA DEUDA ESTÁ CRECIENDO $69/MES!\n")

# ══════════════════════════════════════════════════════════════════════════════
# PRESUPUESTO DISPONIBLE - ANÁLISIS CONSERVADOR
# ══════════════════════════════════════════════════════════════════════════════

print("═" * 80)
print("💰 ANÁLISIS DE PRESUPUESTO DISPONIBLE")
print("═" * 80 + "\n")

# Escenario 1: Conservador (usar solo margen operativo)
presupuesto_conservador = margen_operativo  # $1,466/mes
presupuesto_conservador_colones = presupuesto_conservador * 540

print(f"📊 ESCENARIO CONSERVADOR:")
print(f"   Basado en margen operativo actual: ${presupuesto_conservador:,.2f}/mes")
print(f"   En colones: ₡{presupuesto_conservador_colones:,.2f}/mes")
print(f"   Disponible después de Nissan (${pago_nissan}): ${presupuesto_conservador - pago_nissan:,.2f}\n")

# Escenario 2: Moderado (margen + reducir costos 10%)
ahorro_reduccion_costos = break_even * 0.10  # 10% de $8k = $800
presupuesto_moderado = margen_operativo + ahorro_reduccion_costos
presupuesto_moderado_colones = presupuesto_moderado * 540

print(f"📊 ESCENARIO MODERADO:")
print(f"   Margen operativo: ${margen_operativo:,.2f}")
print(f"   + Reducción costos 10%: ${ahorro_reduccion_costos:,.2f}")
print(f"   = TOTAL: ${presupuesto_moderado:,.2f}/mes")
print(f"   En colones: ₡{presupuesto_moderado_colones:,.2f}/mes")
print(f"   Disponible después de Nissan: ${presupuesto_moderado - pago_nissan:,.2f}\n")

# Escenario 3: Agresivo (aumentar facturación 20%)
aumento_facturacion = facturacion_mensual * 0.20
presupuesto_agresivo = margen_operativo + aumento_facturacion
presupuesto_agresivo_colones = presupuesto_agresivo * 540

print(f"📊 ESCENARIO AGRESIVO:")
print(f"   Margen actual: ${margen_operativo:,.2f}")
print(f"   + Facturación +20%: ${aumento_facturacion:,.2f}")
print(f"   = TOTAL: ${presupuesto_agresivo:,.2f}/mes")
print(f"   En colones: ₡{presupuesto_agresivo_colones:,.2f}/mes")
print(f"   Disponible después de Nissan: ${presupuesto_agresivo - pago_nissan:,.2f}\n")

# ══════════════════════════════════════════════════════════════════════════════
# EJECUTAR ANÁLISIS DE OPTIMIZACIÓN
# ══════════════════════════════════════════════════════════════════════════════

print("═" * 80)
print("🔄 ANÁLISIS DE ESTRATEGIAS DE PAGO - TARJETAS DE CRÉDITO")
print("═" * 80 + "\n")

# Cargar tarjetas (4 tarjetas actuales)
tarjetas = crear_tarjetas_manual()
optimizador = OptimizadorPagos(tarjetas)

# Ejecutar análisis con los 3 escenarios
escenarios_presupuesto = [
    ("CONSERVADOR", presupuesto_conservador_colones),
    ("MODERADO", presupuesto_moderado_colones),
    ("AGRESIVO", presupuesto_agresivo_colones)
]

resultados_por_escenario = {}

for nombre_escenario, presupuesto_colones in escenarios_presupuesto:
    print(f"\n{'═' * 80}")
    print(f"💰 ESCENARIO {nombre_escenario}: ₡{presupuesto_colones:,.2f}/mes (${presupuesto_colones/540:,.2f})")
    print(f"{'═' * 80}\n")

    # Generar análisis
    situacion = optimizador.analizar_situacion_actual()

    print(f"📊 Deuda total TC: ₡{situacion['deuda_total']:,.2f}")
    print(f"💰 Pago mínimo total: ₡{situacion['pago_minimo_total']:,.2f}")
    print(f"📈 Interés mensual: ₡{situacion['interes_mensual_total']:,.2f}\n")

    # Comparar estrategias
    escenarios = optimizador.comparar_estrategias(presupuesto_colones)
    mejor_estrategia, mejor_escenario = optimizador.recomendar_estrategia(presupuesto_colones)

    # Mostrar comparativa
    print("┌─────────────┬──────────────┬─────────────────┬────────────────────┐")
    print("│ Estrategia  │ Meses Libre  │ Total Intereses │ Ahorro vs Mínimo   │")
    print("├─────────────┼──────────────┼─────────────────┼────────────────────┤")

    for nombre, esc in escenarios.items():
        if nombre != 'minimum':
            print(f"│ {nombre:11} │ {esc.meses_hasta_libre:12} │ ₡{esc.total_intereses:13,.2f} │ ₡{esc.ahorro_vs_minimo:16,.2f} │")

    # Mostrar escenario mínimo como baseline
    esc_min = escenarios['minimum']
    print(f"│ {'minimum (baseline)':11} │ {esc_min.meses_hasta_libre:12} │ ₡{esc_min.total_intereses:13,.2f} │ {'N/A':>18} │")
    print("└─────────────┴──────────────┴─────────────────┴────────────────────┘\n")

    print(f"🏆 MEJOR ESTRATEGIA: {mejor_estrategia.upper()}")
    print(f"   Tiempo hasta libre de deuda: {mejor_escenario.meses_hasta_libre} meses ({mejor_escenario.meses_hasta_libre/12:.1f} años)")
    print(f"   Total en intereses: ₡{mejor_escenario.total_intereses:,.2f}")
    print(f"   Ahorro vs pago mínimo: ₡{mejor_escenario.ahorro_vs_minimo:,.2f}\n")

    resultados_por_escenario[nombre_escenario] = {
        'presupuesto': presupuesto_colones,
        'mejor_estrategia': mejor_estrategia,
        'meses_libre': mejor_escenario.meses_hasta_libre,
        'total_intereses': mejor_escenario.total_intereses,
        'ahorro': mejor_escenario.ahorro_vs_minimo
    }

# ══════════════════════════════════════════════════════════════════════════════
# RESUMEN EJECUTIVO Y RECOMENDACIONES
# ══════════════════════════════════════════════════════════════════════════════

print("\n" + "═" * 80)
print("📊 RESUMEN EJECUTIVO - COMPARATIVA DE ESCENARIOS")
print("═" * 80 + "\n")

print("┌──────────────┬─────────────────┬────────────────┬─────────────────┬───────────────────┐")
print("│ Escenario    │ Presupuesto/mes │ Meses Libre    │ Total Intereses │ Ahorro vs Mínimo  │")
print("├──────────────┼─────────────────┼────────────────┼─────────────────┼───────────────────┤")

for nombre_esc, datos in resultados_por_escenario.items():
    print(f"│ {nombre_esc:12} │ ₡{datos['presupuesto']:13,.2f} │ {datos['meses_libre']:14} │ ₡{datos['total_intereses']:13,.2f} │ ₡{datos['ahorro']:15,.2f} │")

print("└──────────────┴─────────────────┴────────────────┴─────────────────┴───────────────────┘\n")

# ══════════════════════════════════════════════════════════════════════════════
# RECOMENDACIONES CRÍTICAS
# ══════════════════════════════════════════════════════════════════════════════

print("═" * 80)
print("🎯 RECOMENDACIONES CRÍTICAS BASADAS EN TU SITUACIÓN")
print("═" * 80 + "\n")

print("🚨 URGENTE - PAGO CREDOMATIC MAÑANA (15/11/2025):")
print("   Deuda: ₡1,419,305.54 + $209.92")
print("   Pago mínimo: ₡64,728 + $24 (~₡77,688 total)")
print("   Efectivo disponible: ₡1,860,048 (~$3,444)")
print("   ✅ SÍ tienes liquidez para el pago mínimo\n")

print("💰 ESTRATEGIA RECOMENDADA:")

if presupuesto_conservador > pago_nissan + pago_hacienda_intereses:
    disponible_tc = presupuesto_conservador - pago_nissan
    print(f"   1. Continuar pagando Nissan: ${pago_nissan}/mes (obligatorio)")
    print(f"   2. Destinar ${disponible_tc:.2f}/mes a tarjetas de crédito")
    print(f"   3. ⚠️  CRÍTICO: ${disponible_tc:.2f} < ${intereses_tc_mensuales:.2f} (intereses)")
    print(f"   4. Necesitas aumentar presupuesto a mínimo ${pago_nissan + intereses_tc_mensuales:.2f}/mes")
else:
    print(f"   ⛔ CRISIS: No hay suficiente margen para pagar deudas")
    print(f"   Necesitas: ${pago_nissan + intereses_tc_mensuales:.2f}/mes mínimo")
    print(f"   Tienes: ${presupuesto_conservador:.2f}/mes")
    print(f"   DÉFICIT: ${(pago_nissan + intereses_tc_mensuales) - presupuesto_conservador:.2f}/mes")

print("\n📈 OPCIONES PARA AUMENTAR PRESUPUESTO:")
print("   A. Cobrar CxC más rápido: Tienes $9k-$11k por cobrar")
print("   B. Negociar con Hacienda: Arreglo de pago formal")
print("   C. Aumentar facturación: +20% = +$1,893/mes extra")
print("   D. Reducir costos operativos: -10% = +$800/mes extra")
print("   E. Reestructurar Nissan: Extender plazo para bajar cuota\n")

print("⚠️  ALERTA HACIENDA:")
print(f"   Deuda actual: ${deuda_hacienda:,.2f}")
print(f"   Creciendo: ${pago_hacienda_intereses:.2f}/mes (2% mensual)")
print(f"   En 12 meses: ${deuda_hacienda * 1.268:,.2f} (+26.8%)")
print("   🚨 URGENTE: Establecer arreglo de pago formal\n")

print("═" * 80)
print("✅ ANÁLISIS COMPLETADO")
print("═" * 80)
