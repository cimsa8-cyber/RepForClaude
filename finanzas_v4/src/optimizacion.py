#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     🎯 MÓDULO DE OPTIMIZACIÓN FINANCIERA                     ║
║                              Sistema v4.0                                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

FUNCIONALIDAD:
- Optimización de estrategias de pago de deudas
- Proyecciones de flujo de caja
- Simulación de escenarios financieros
- Priorización inteligente de pagos
- Análisis de ahorro en intereses

ESTRATEGIAS IMPLEMENTADAS:
1. Avalanche (mayor tasa de interés primero)
2. Snowball (menor deuda primero)
3. Cashflow (optimizar liquidez)
4. Hybrid (combinación inteligente)

RESCATADO DE v3.0:
- optimize_payments.py ⭐⭐⭐⭐⭐
- plan_pago_cxp.py ⭐⭐⭐⭐⭐
- proyecciones_flujo.py ⭐⭐⭐⭐
- simulate_cycles.py ⭐⭐⭐⭐

Autor: Sistema v4.0
Fecha: 2025-11-14
"""

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional
from enum import Enum
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class EstrategiaPago(Enum):
    """Estrategias de optimización de pagos"""
    AVALANCHE = "avalanche"  # Mayor tasa de interés primero
    SNOWBALL = "snowball"    # Menor deuda primero
    CASHFLOW = "cashflow"    # Optimizar flujo de caja
    HYBRID = "hybrid"        # Combinación inteligente
    MINIMUM = "minimum"      # Solo pagos mínimos


@dataclass
class TarjetaCredito:
    """Representa una tarjeta de crédito con su deuda"""
    nombre: str
    deuda_colones: float
    deuda_dolares: float
    limite_colones: float
    limite_dolares: float
    tasa_interes_mensual: float
    pago_minimo_colones: float
    pago_minimo_dolares: float
    fecha_corte: int  # Día del mes
    fecha_pago: int   # Día del mes

    @property
    def deuda_total_colones(self) -> float:
        """Deuda total en colones (usando tipo de cambio fijo 540)"""
        return self.deuda_colones + (self.deuda_dolares * 540)

    @property
    def pago_minimo_total_colones(self) -> float:
        """Pago mínimo total en colones"""
        return self.pago_minimo_colones + (self.pago_minimo_dolares * 540)

    @property
    def porcentaje_utilizado(self) -> float:
        """Porcentaje del límite utilizado"""
        limite_total = self.limite_colones + (self.limite_dolares * 540)
        if limite_total == 0:
            return 0
        return (self.deuda_total_colones / limite_total) * 100

    @property
    def interes_mensual_colones(self) -> float:
        """Interés mensual estimado en colones"""
        return self.deuda_total_colones * (self.tasa_interes_mensual / 100)


@dataclass
class EscenarioPago:
    """Escenario de pago proyectado"""
    estrategia: EstrategiaPago
    meses_hasta_libre: int
    total_intereses: float
    pago_mensual_promedio: float
    ahorro_vs_minimo: float
    cronograma: List[Dict]

    def __str__(self) -> str:
        return f"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  Estrategia: {self.estrategia.value.upper()}
║  Tiempo hasta estar libre de deuda: {self.meses_hasta_libre} meses ({self.meses_hasta_libre/12:.1f} años)
║  Total en intereses: ₡{self.total_intereses:,.2f}
║  Pago mensual promedio: ₡{self.pago_mensual_promedio:,.2f}
║  Ahorro vs pago mínimo: ₡{self.ahorro_vs_minimo:,.2f}
╚══════════════════════════════════════════════════════════════════════════════╝
"""


class OptimizadorPagos:
    """
    Optimizador de pagos de tarjetas de crédito

    Analiza múltiples estrategias y recomienda la mejor opción
    basándose en las finanzas del usuario.
    """

    TIPO_CAMBIO = 540  # CRC por USD

    def __init__(self, tarjetas: List[TarjetaCredito]):
        self.tarjetas = tarjetas
        self.deuda_total = sum(t.deuda_total_colones for t in tarjetas)
        self.pago_minimo_total = sum(t.pago_minimo_total_colones for t in tarjetas)

    def analizar_situacion_actual(self) -> Dict:
        """Analiza la situación financiera actual"""
        return {
            'num_tarjetas': len(self.tarjetas),
            'deuda_total': self.deuda_total,
            'pago_minimo_total': self.pago_minimo_total,
            'interes_mensual_total': sum(t.interes_mensual_colones for t in self.tarjetas),
            'tarjeta_mayor_deuda': max(self.tarjetas, key=lambda t: t.deuda_total_colones),
            'tarjeta_mayor_interes': max(self.tarjetas, key=lambda t: t.tasa_interes_mensual),
            'tarjeta_mas_utilizada': max(self.tarjetas, key=lambda t: t.porcentaje_utilizado),
            'promedio_utilizacion': sum(t.porcentaje_utilizado for t in self.tarjetas) / len(self.tarjetas)
        }

    def calcular_escenario_avalanche(self, presupuesto_mensual: float) -> EscenarioPago:
        """
        Estrategia Avalanche: Pagar primero las tarjetas con mayor tasa de interés

        Esta estrategia minimiza el total pagado en intereses.
        """
        tarjetas_ordenadas = sorted(self.tarjetas,
                                   key=lambda t: t.tasa_interes_mensual,
                                   reverse=True)
        return self._simular_pagos(tarjetas_ordenadas, presupuesto_mensual, EstrategiaPago.AVALANCHE)

    def calcular_escenario_snowball(self, presupuesto_mensual: float) -> EscenarioPago:
        """
        Estrategia Snowball: Pagar primero las tarjetas con menor deuda

        Esta estrategia proporciona victorias rápidas y motivación psicológica.
        """
        tarjetas_ordenadas = sorted(self.tarjetas,
                                   key=lambda t: t.deuda_total_colones)
        return self._simular_pagos(tarjetas_ordenadas, presupuesto_mensual, EstrategiaPago.SNOWBALL)

    def calcular_escenario_cashflow(self, presupuesto_mensual: float) -> EscenarioPago:
        """
        Estrategia Cashflow: Optimizar flujo de caja mensual

        Balancea entre liberación de límites y reducción de intereses.
        """
        # Priorizar tarjetas con alto porcentaje de utilización
        tarjetas_ordenadas = sorted(self.tarjetas,
                                   key=lambda t: t.porcentaje_utilizado,
                                   reverse=True)
        return self._simular_pagos(tarjetas_ordenadas, presupuesto_mensual, EstrategiaPago.CASHFLOW)

    def calcular_escenario_hybrid(self, presupuesto_mensual: float) -> EscenarioPago:
        """
        Estrategia Híbrida: Combinación inteligente

        Balancea entre tasa de interés, tamaño de deuda y utilización.
        """
        # Calcular score compuesto para cada tarjeta
        def calcular_score(tarjeta: TarjetaCredito) -> float:
            # Normalizar valores
            max_interes = max(t.tasa_interes_mensual for t in self.tarjetas)
            max_utilizacion = max(t.porcentaje_utilizado for t in self.tarjetas)

            score_interes = (tarjeta.tasa_interes_mensual / max_interes) * 40
            score_utilizacion = (tarjeta.porcentaje_utilizado / max_utilizacion) * 40
            score_deuda_pequena = 20 if tarjeta.deuda_total_colones < self.deuda_total * 0.2 else 0

            return score_interes + score_utilizacion + score_deuda_pequena

        tarjetas_ordenadas = sorted(self.tarjetas,
                                   key=calcular_score,
                                   reverse=True)
        return self._simular_pagos(tarjetas_ordenadas, presupuesto_mensual, EstrategiaPago.HYBRID)

    def calcular_escenario_minimo(self) -> EscenarioPago:
        """
        Escenario de solo pagos mínimos (usado como baseline)

        ⚠️ ADVERTENCIA: Esta estrategia resulta en el mayor costo total.
        """
        return self._simular_pagos(self.tarjetas, self.pago_minimo_total, EstrategiaPago.MINIMUM)

    def _simular_pagos(self, tarjetas_ordenadas: List[TarjetaCredito],
                      presupuesto_mensual: float,
                      estrategia: EstrategiaPago) -> EscenarioPago:
        """
        Simula pagos mes a mes hasta liquidar todas las deudas

        Args:
            tarjetas_ordenadas: Lista de tarjetas en orden de prioridad
            presupuesto_mensual: Presupuesto disponible para pagos
            estrategia: Estrategia utilizada

        Returns:
            EscenarioPago con la proyección completa
        """
        # Copiar deudas para no modificar originales
        deudas_actuales = {t.nombre: t.deuda_total_colones for t in tarjetas_ordenadas}
        cronograma = []
        mes = 0
        total_intereses = 0

        while any(deuda > 1 for deuda in deudas_actuales.values()) and mes < 600:  # Max 50 años
            mes += 1
            fecha = datetime.now() + timedelta(days=30 * mes)
            presupuesto_restante = presupuesto_mensual
            pagos_mes = {}

            # Aplicar intereses a todas las tarjetas
            for tarjeta in tarjetas_ordenadas:
                if deudas_actuales[tarjeta.nombre] > 0:
                    interes = deudas_actuales[tarjeta.nombre] * (tarjeta.tasa_interes_mensual / 100)
                    deudas_actuales[tarjeta.nombre] += interes
                    total_intereses += interes

            # Distribuir presupuesto según prioridad
            for tarjeta in tarjetas_ordenadas:
                if presupuesto_restante <= 0:
                    break

                if deudas_actuales[tarjeta.nombre] <= 0:
                    continue

                # Pago mínimo obligatorio
                pago_minimo = min(tarjeta.pago_minimo_total_colones, deudas_actuales[tarjeta.nombre])

                # Si es la tarjeta prioritaria, pagar todo el presupuesto restante
                if tarjeta == tarjetas_ordenadas[0]:
                    pago = min(presupuesto_restante, deudas_actuales[tarjeta.nombre])
                else:
                    pago = pago_minimo

                if pago > presupuesto_restante:
                    pago = presupuesto_restante

                deudas_actuales[tarjeta.nombre] -= pago
                presupuesto_restante -= pago
                pagos_mes[tarjeta.nombre] = pago

                # Si se liquidó esta tarjeta, moverla al final
                if deudas_actuales[tarjeta.nombre] <= 0:
                    deudas_actuales[tarjeta.nombre] = 0
                    tarjetas_ordenadas.remove(tarjeta)
                    tarjetas_ordenadas.append(tarjeta)

            cronograma.append({
                'mes': mes,
                'fecha': fecha,
                'pagos': pagos_mes.copy(),
                'deudas_restantes': deudas_actuales.copy(),
                'total_pagado': sum(pagos_mes.values()),
                'deuda_total_restante': sum(deudas_actuales.values())
            })

        # Calcular escenario mínimo para comparación (si no es ya el escenario mínimo)
        if estrategia != EstrategiaPago.MINIMUM:
            escenario_minimo = self.calcular_escenario_minimo()
            ahorro = escenario_minimo.total_intereses - total_intereses
        else:
            ahorro = 0

        return EscenarioPago(
            estrategia=estrategia,
            meses_hasta_libre=mes,
            total_intereses=total_intereses,
            pago_mensual_promedio=presupuesto_mensual if estrategia != EstrategiaPago.MINIMUM else self.pago_minimo_total,
            ahorro_vs_minimo=ahorro,
            cronograma=cronograma
        )

    def comparar_estrategias(self, presupuesto_mensual: float) -> Dict[str, EscenarioPago]:
        """
        Compara todas las estrategias disponibles

        Args:
            presupuesto_mensual: Presupuesto disponible para pagos

        Returns:
            Diccionario con todos los escenarios calculados
        """
        print("\n🔄 Calculando escenarios...")

        escenarios = {
            'avalanche': self.calcular_escenario_avalanche(presupuesto_mensual),
            'snowball': self.calcular_escenario_snowball(presupuesto_mensual),
            'cashflow': self.calcular_escenario_cashflow(presupuesto_mensual),
            'hybrid': self.calcular_escenario_hybrid(presupuesto_mensual),
            'minimum': self.calcular_escenario_minimo()
        }

        return escenarios

    def recomendar_estrategia(self, presupuesto_mensual: float) -> Tuple[str, EscenarioPago]:
        """
        Recomienda la mejor estrategia basándose en el análisis

        Args:
            presupuesto_mensual: Presupuesto disponible para pagos

        Returns:
            Tupla (nombre_estrategia, escenario_recomendado)
        """
        escenarios = self.comparar_estrategias(presupuesto_mensual)

        # Excluir escenario mínimo de la recomendación
        escenarios_validos = {k: v for k, v in escenarios.items()
                             if v.estrategia != EstrategiaPago.MINIMUM}

        # Seleccionar mejor opción (menor interés total)
        mejor = min(escenarios_validos.items(),
                   key=lambda x: x[1].total_intereses)

        return mejor

    def generar_reporte(self, presupuesto_mensual: float, output_path: str = "analisis_optimizacion.txt"):
        """
        Genera un reporte completo de análisis y recomendaciones

        Args:
            presupuesto_mensual: Presupuesto disponible para pagos
            output_path: Ruta del archivo de salida
        """
        situacion = self.analizar_situacion_actual()
        escenarios = self.comparar_estrategias(presupuesto_mensual)
        mejor_estrategia, mejor_escenario = self.recomendar_estrategia(presupuesto_mensual)

        reporte = f"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                   📊 ANÁLISIS DE OPTIMIZACIÓN FINANCIERA                     ║
║                              Sistema v4.0                                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

📅 Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}

═══════════════════════════════════════════════════════════════════════════════
🎯 SITUACIÓN ACTUAL
═══════════════════════════════════════════════════════════════════════════════

💳 Número de tarjetas: {situacion['num_tarjetas']}
💰 Deuda total: ₡{situacion['deuda_total']:,.2f}
📉 Pago mínimo total: ₡{situacion['pago_minimo_total']:,.2f}
📈 Interés mensual estimado: ₡{situacion['interes_mensual_total']:,.2f}
📊 Promedio utilización: {situacion['promedio_utilizacion']:.1f}%

🔥 TARJETA CON MAYOR DEUDA:
   {situacion['tarjeta_mayor_deuda'].nombre}: ₡{situacion['tarjeta_mayor_deuda'].deuda_total_colones:,.2f}

⚠️  TARJETA CON MAYOR INTERÉS:
   {situacion['tarjeta_mayor_interes'].nombre}: {situacion['tarjeta_mayor_interes'].tasa_interes_mensual:.2f}% mensual

🚨 TARJETA MÁS UTILIZADA:
   {situacion['tarjeta_mas_utilizada'].nombre}: {situacion['tarjeta_mas_utilizada'].porcentaje_utilizado:.1f}% del límite

═══════════════════════════════════════════════════════════════════════════════
💡 PRESUPUESTO DE PAGO
═══════════════════════════════════════════════════════════════════════════════

💵 Presupuesto mensual disponible: ₡{presupuesto_mensual:,.2f}
📊 Pago mínimo requerido: ₡{situacion['pago_minimo_total']:,.2f}
✅ Presupuesto extra para pagar deudas: ₡{presupuesto_mensual - situacion['pago_minimo_total']:,.2f}

═══════════════════════════════════════════════════════════════════════════════
📈 COMPARACIÓN DE ESTRATEGIAS
═══════════════════════════════════════════════════════════════════════════════
"""

        # Agregar cada escenario
        for nombre, escenario in escenarios.items():
            reporte += f"\n{escenario}"

        # Agregar recomendación
        reporte += f"""
═══════════════════════════════════════════════════════════════════════════════
🏆 RECOMENDACIÓN
═══════════════════════════════════════════════════════════════════════════════

✅ ESTRATEGIA RECOMENDADA: {mejor_estrategia.upper()}

{mejor_escenario}

🎯 RAZONES DE LA RECOMENDACIÓN:
"""

        if mejor_estrategia == 'avalanche':
            reporte += """
   • Minimiza el costo total en intereses
   • Ataca primero las deudas más caras
   • Recomendada si eres disciplinado y orientado a números
"""
        elif mejor_estrategia == 'snowball':
            reporte += """
   • Proporciona victorias rápidas
   • Motivación psicológica al eliminar deudas pequeñas
   • Recomendada si necesitas ver progreso rápido
"""
        elif mejor_estrategia == 'cashflow':
            reporte += """
   • Optimiza tu flujo de caja mensual
   • Libera límites de crédito rápidamente
   • Recomendada si necesitas mantener flexibilidad financiera
"""
        elif mejor_estrategia == 'hybrid':
            reporte += """
   • Balance óptimo entre costo y motivación
   • Considera múltiples factores simultáneamente
   • Recomendada para la mayoría de situaciones
"""

        reporte += f"""
💰 AHORRO ESTIMADO VS PAGO MÍNIMO: ₡{mejor_escenario.ahorro_vs_minimo:,.2f}
⏱️  TIEMPO DE LIBERTAD: {mejor_escenario.meses_hasta_libre} meses ({mejor_escenario.meses_hasta_libre/12:.1f} años)

═══════════════════════════════════════════════════════════════════════════════
⚠️  ADVERTENCIAS Y CONSIDERACIONES
═══════════════════════════════════════════════════════════════════════════════

1. 🛑 NO usar las tarjetas mientras se pagan las deudas
2. 📅 Configurar pagos automáticos para evitar atrasos
3. 💰 Si recibes ingresos extra, destinarlos a deuda prioritaria
4. 📊 Revisar el plan mensualmente y ajustar si es necesario
5. 🎯 Considerar consolidación si tasas son muy altas (>3% mensual)

═══════════════════════════════════════════════════════════════════════════════
📝 PRÓXIMOS PASOS
═══════════════════════════════════════════════════════════════════════════════

1. ✅ Revisar y aprobar la estrategia recomendada
2. ✅ Configurar pagos automáticos en fechas clave
3. ✅ Detener uso de tarjetas de crédito
4. ✅ Crear fondo de emergencia pequeño (₡200,000 mínimo)
5. ✅ Ejecutar el plan con disciplina
6. ✅ Celebrar cada tarjeta liquidada 🎉

═══════════════════════════════════════════════════════════════════════════════

🚀 ¡Puedes liberarte de estas deudas! La clave es disciplina y consistencia.

╔══════════════════════════════════════════════════════════════════════════════╗
║              Generado por Sistema de Finanzas v4.0 - Módulo de               ║
║                          Optimización de Pagos                               ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

        # Guardar reporte
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(reporte)

        print(f"\n✅ Reporte generado: {output_path}")
        return reporte


# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE UTILIDAD
# ══════════════════════════════════════════════════════════════════════════════

def cargar_tarjetas_desde_excel(ruta_excel: str) -> List[TarjetaCredito]:
    """
    Carga información de tarjetas desde el archivo Excel v4.0

    Args:
        ruta_excel: Ruta al archivo Excel de finanzas

    Returns:
        Lista de objetos TarjetaCredito
    """
    # TODO: Implementar lectura desde TRANSACCIONES y ENTIDADES_ALIAS
    # Por ahora retornar datos hardcoded para las 4 tarjetas conocidas
    pass


def crear_tarjetas_manual() -> List[TarjetaCredito]:
    """
    Crea manualmente las 4 tarjetas conocidas del usuario

    Returns:
        Lista con las 4 tarjetas de crédito
    """
    return [
        TarjetaCredito(
            nombre="Visa Clásica ***3519",
            deuda_colones=590158.64,
            deuda_dolares=65.04,
            limite_colones=0,
            limite_dolares=1200,
            tasa_interes_mensual=3.5,  # Estimado para Visa clásica
            pago_minimo_colones=18000,
            pago_minimo_dolares=14,
            fecha_corte=7,
            fecha_pago=17
        ),
        TarjetaCredito(
            nombre="Visa Platino ***9837",
            deuda_colones=2086984.01,
            deuda_dolares=1775.45,
            limite_colones=0,
            limite_dolares=9000,
            tasa_interes_mensual=2.5,  # Estimado para Visa Platino
            pago_minimo_colones=65100,
            pago_minimo_dolares=174,
            fecha_corte=26,
            fecha_pago=5
        ),
        TarjetaCredito(
            nombre="MasterCard Oro ***8759",
            deuda_colones=2847410.17,
            deuda_dolares=256.09,
            limite_colones=0,
            limite_dolares=6000,
            tasa_interes_mensual=2.8,  # Estimado para MasterCard Oro
            pago_minimo_colones=93500,
            pago_minimo_dolares=23,
            fecha_corte=3,
            fecha_pago=13
        ),
        TarjetaCredito(
            nombre="Credomatic",
            deuda_colones=1419305.54,
            deuda_dolares=209.92,
            limite_colones=0,
            limite_dolares=3000,  # Estimado
            tasa_interes_mensual=3.2,  # Estimado
            pago_minimo_colones=64728,
            pago_minimo_dolares=24,
            fecha_corte=5,  # Estimado
            fecha_pago=15
        )
    ]


# ══════════════════════════════════════════════════════════════════════════════
# SCRIPT PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║              🎯 OPTIMIZADOR DE PAGOS - SISTEMA FINANZAS V4.0                ║
╚══════════════════════════════════════════════════════════════════════════════╝
""")

    # Cargar tarjetas
    print("📋 Cargando información de tarjetas...")
    tarjetas = crear_tarjetas_manual()
    print(f"✅ {len(tarjetas)} tarjetas cargadas\n")

    # Crear optimizador
    optimizador = OptimizadorPagos(tarjetas)

    # Solicitar presupuesto al usuario
    print("💰 Ingresa tu presupuesto mensual disponible para pagar deudas:")
    print(f"   (Pago mínimo requerido: ₡{optimizador.pago_minimo_total:,.2f})")

    try:
        presupuesto = float(input("\n   Presupuesto mensual (₡): ").replace(",", ""))

        if presupuesto < optimizador.pago_minimo_total:
            print(f"\n⚠️  ADVERTENCIA: El presupuesto es menor que el pago mínimo requerido.")
            print(f"   Ajustando a pago mínimo: ₡{optimizador.pago_minimo_total:,.2f}")
            presupuesto = optimizador.pago_minimo_total

        # Generar análisis completo
        print("\n🔄 Analizando todas las estrategias posibles...")
        print("   (Esto puede tomar unos segundos...)\n")

        reporte = optimizador.generar_reporte(presupuesto)
        print(reporte)

        print("\n✅ ¡Análisis completado!")
        print("📄 Archivo generado: analisis_optimizacion.txt")

    except KeyboardInterrupt:
        print("\n\n❌ Análisis cancelado por el usuario.")
    except ValueError:
        print("\n\n❌ Error: Debes ingresar un número válido.")
