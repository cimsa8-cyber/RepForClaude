# 💰 Sistema de Finanzas v4.0 - Documentación Completa

**Autor:** Álvaro Velasco + Claude
**Fecha:** Noviembre 2024
**Versión:** 4.0 (Completa y Auditada)

---

## 📋 Índice

1. [Resumen Ejecutivo](#resumen-ejecutivo)
2. [Características Principales](#características-principales)
3. [Estructura del Sistema](#estructura-del-sistema)
4. [Instalación y Uso](#instalación-y-uso)
5. [Arquitectura de Datos](#arquitectura-de-datos)
6. [Hojas de Excel (14 Total)](#hojas-de-excel)
7. [Reglas de Negocio](#reglas-de-negocio)
8. [Datos de Ejemplo](#datos-de-ejemplo)
9. [Próximos Pasos](#próximos-pasos)
10. [Soporte y Contacto](#soporte-y-contacto)

---

## 🎯 Resumen Ejecutivo

El **Sistema de Finanzas v4.0** es una solución completa en Excel para gestión financiera personal y empresarial, generada mediante Python con openpyxl. Este sistema permite:

- ✅ Control total de **Cuentas por Pagar (CxP)** y **Cuentas por Cobrar (CxC)**
- ✅ Gestión de **4 tarjetas de crédito** integradas en CxP
- ✅ Estados financieros automáticos (P&L, Balance General, Flujo de Caja)
- ✅ Control de **IVA** con exempciones para empresas de zona franca
- ✅ Separación de **gastos personales vs negocio**
- ✅ Presupuesto vs Real por categoría
- ✅ Dashboard visual con análisis de gastos

**Principio fundamental:** Una sola hoja editable (TRANSACCIONES), todas las demás se calculan automáticamente.

---

## ⚡ Características Principales

### 1. **16 Columnas en TRANSACCIONES**
- Fecha, Entidad, Categoría, Subcategoría, Moneda, Monto
- Descripción, Forma de Pago, IVA, Notas, Recurrente
- Proyecto, Estado, Factura #, Tag, **Personal/Negocio**

### 2. **Tarjetas de Crédito Integradas**
Las 4 tarjetas aparecen automáticamente en CxP:
- **BAC Visa** (USD) - Pago día 15
- **BCR Mastercard** (CRC) - Pago día 20
- **Credomatic Platinum** (USD) - Pago día 10
- **Credomatic Gold** (CRC) - Pago día 10

**Deuda actual:** ~₡6.9M (incluida en CxP)

### 3. **Control de IVA Inteligente**
- Excluye automáticamente empresas de zona franca: **VWR** y **RS Hughes**
- Calcula IVA Compras (crédito fiscal) vs IVA Ventas (débito fiscal)
- Balance a pagar a Hacienda

### 4. **Separación Personal vs Negocio**
- Nueva columna P: "Personal/Negocio"
- Alerta automática si gastos personales > 30%
- Análisis de % de gastos personales

### 5. **12 Alias de Cuentas Pre-cargados**
| Alias | Nombre Completo | Tipo |
|-------|-----------------|------|
| BAC-USD | BAC San José - Cuenta USD | Banco |
| BCR-CRC | Banco de Costa Rica - Cuenta ₡ | Banco |
| Visa-BAC | Visa BAC Crédito | Tarjeta Crédito |
| MC-BCR | Mastercard BCR | Tarjeta Crédito |
| Credo-Platinum | Credomatic Platinum USD | Tarjeta Crédito |
| Credo-Gold | Credomatic Gold CRC | Tarjeta Crédito |
| VWR | VWR International LLC | Proveedor |
| RS-Hughes | RS Hughes Co. Inc. | Proveedor |
| Efectivo | Efectivo - Caja Chica | Efectivo |
| Cliente-XYZ | Cliente XYZ Corp | Cliente |
| SINPE | SINPE Móvil | Transferencia |
| PayPal | PayPal Business | Digital |

---

## 📁 Estructura del Sistema

```
finanzas_v4/
├── README.md                          # Este archivo
├── LESSONS_LEARNED.md                 # Errores y soluciones
├── requirements.txt                   # Dependencias Python
├── src/
│   └── generar_v4_completo.py        # Generador principal (810 líneas)
├── docs/
│   ├── architecture.md                # Arquitectura técnica
│   └── user_guide.md                  # Guía de usuario
├── examples/
│   └── transacciones_ejemplo.xlsx     # Ejemplos de transacciones
├── data/
│   └── cuentas_reales.json           # Datos de cuentas reales (CxP, CxC)
└── tests/
    └── test_generador.py              # Tests unitarios
```

---

## 🚀 Instalación y Uso

### Requisitos Previos
- Python 3.11+
- pip (gestor de paquetes)

### Paso 1: Instalar Dependencias

```bash
pip install openpyxl
```

O usar el archivo requirements.txt:

```bash
pip install -r requirements.txt
```

### Paso 2: Generar el Excel

```bash
cd finanzas_v4/src
python generar_v4_completo.py
```

### Paso 3: Abrir el Excel

Se creará el archivo:
```
AlvaroVelasco_Finanzas_v4.0_COMPLETO.xlsx
```

**¡Listo para usar!** 🎉

---

## 📊 Hojas de Excel (14 Total)

### 🔹 Hojas EDITABLES (2)

#### 1. **TRANSACCIONES** (Hoja principal)
- **Color:** Azul
- **Editable:** ✅ SÍ
- **Función:** Registrar TODAS las transacciones aquí
- **Columnas:** 16 (incluye Personal/Negocio)
- **Transacciones de ejemplo:** 3 pre-cargadas

#### 2. **CONFIG** (Configuración)
- **Color:** Naranja
- **Editable:** ✅ SÍ
- **Función:** Configurar tipo de cambio y fechas de pago
- **Celdas amarillas:** Editables
- **Incluye:**
  - Tipo de cambio USD → CRC (540 por defecto)
  - Fechas de pago de tarjetas (15, 20, 10, 10)
  - Lista de empresas zona franca (VWR, RS Hughes)

---

### 🔒 Hojas PROTEGIDAS (12 - Auto-calculadas)

#### 3. **RESUMEN** (Dashboard Ejecutivo)
- Vista general de CxP, CxC, Flujo de Caja, IVA
- Actualización en tiempo real
- **Primera hoja** que se ve al abrir el Excel

#### 4. **CxP** (Cuentas por Pagar)
- Incluye **tarjetas de crédito** automáticamente
- Extrae de TRANSACCIONES donde Estado = "Pendiente"
- Total USD y Total ₡
- Días vencidos automáticos

#### 5. **CxC** (Cuentas por Cobrar)
- Facturas pendientes de cobro
- Extrae de TRANSACCIONES donde Categoría = "CxC" y Estado = "Pendiente"
- Total USD
- Alerta de facturas > 60 días vencidas

#### 6. **ENTIDADES_ALIAS**
- Lista de 12 alias pre-cargados
- Tipo: Banco, Tarjeta, Proveedor, Cliente, etc.

#### 7. **ESTADO_RESULTADOS** (P&L)
- Ingresos Operacionales
- Gastos Operacionales
- **Utilidad Neta** = Ingresos - Gastos
- Solo en USD

#### 8. **BALANCE_GENERAL**
- **Activos:** CxC + Efectivo
- **Pasivos:** CxP + IVA por Pagar
- **Patrimonio** = Activos - Pasivos

#### 9. **FLUJO_CAJA**
- Entradas (Ingresos) en USD y CRC
- Salidas (Gastos) en USD y CRC
- **Balance Final** por moneda

#### 10. **DASHBOARD_VISUAL**
- Top 5 Gastos por Categoría
- Análisis visual de gastos
- (Preparado para gráficos futuros)

#### 11. **CONCILIACION**
- Conciliación bancaria
- Compara saldo del sistema vs saldo bancario real
- Campo editable para ingresar saldo bancario

#### 12. **IVA_CONTROL**
- IVA Compras (Crédito Fiscal)
- IVA Ventas (Débito Fiscal)
- **Excluye automáticamente VWR y RS Hughes** (zona franca)
- Balance a pagar a Hacienda

#### 13. **PRESUPUESTO**
- Presupuesto vs Real por categoría
- Celdas amarillas editables para presupuesto
- % de ejecución automático
- Categorías: Servicios, Inventario, G. Administrativos, Marketing, Personal

#### 14. **PERSONAL_VS_NEGOCIO**
- Total gastos de Negocio
- Total gastos Personales
- **% de gastos personales**
- ⚠️ **Alerta** si gastos personales > 30%

---

## 📐 Arquitectura de Datos

### Principio Fundamental: MÁXIMA SIMPLICIDAD

```
┌─────────────────────────────────────┐
│     TRANSACCIONES (Editable)        │  ← Única fuente de verdad
│  16 columnas, datos crudos          │
└──────────────┬──────────────────────┘
               │
               │ (Fórmulas de Excel)
               │
               ▼
┌──────────────────────────────────────────────────────────┐
│  12 HOJAS AUTO-CALCULADAS (Protegidas)                   │
│  • CxP, CxC, Estados Financieros, IVA, etc.              │
│  • Sin macros, solo fórmulas nativas de Excel            │
└──────────────────────────────────────────────────────────┘
```

### Flujo de Información

1. **Input:** Usuario ingresa transacciones en **TRANSACCIONES**
2. **Procesamiento:** Fórmulas de Excel (SUMAR.SI, SUMAR.SI.CONJUNTO, etc.)
3. **Output:** Actualización automática de las 12 hojas protegidas
4. **Config:** Usuario ajusta tipo de cambio en **CONFIG** si es necesario

### Tecnologías Utilizadas
- **Python 3.11+**
- **openpyxl 3.1.2** - Biblioteca para crear archivos Excel
- **Excel formulas** - Sin VBA, sin macros

---

## ⚖️ Reglas de Negocio

### 1. **Regla de Oro: Una Sola Fuente de Verdad**
✅ **HACER:** Ingresar TODO en TRANSACCIONES
❌ **NO HACER:** Editar celdas en hojas protegidas

### 2. **Tarjetas de Crédito = CxP**
- Las tarjetas de crédito **DEBEN** aparecer en CxP
- Se registran como transacciones con Estado = "Pendiente"
- Categoría puede ser "CxP" o la categoría del gasto

### 3. **IVA y Zona Franca**
- **VWR International** y **RS Hughes** están exentos de IVA
- El sistema los excluye automáticamente del cálculo de IVA
- Para agregar más empresas zona franca: editar fórmula en IVA_CONTROL

### 4. **Gastos Personales vs Negocio**
- **Columna P (16):** "Personal" o "Negocio"
- Target: Mantener gastos personales < 30%
- Sistema alerta automáticamente si se excede

### 5. **Tipo de Cambio**
- Editable en CONFIG (celda B5)
- Valor por defecto: 540 CRC/USD
- Actualizar manualmente según necesidad

### 6. **Estados de Transacciones**
- **Pagado:** Ya ejecutado
- **Pendiente:** Por pagar (aparece en CxP o CxC)
- **Cobrado:** Ya cobrado (para ingresos)
- **Cancelado:** Transacción anulada

---

## 📝 Datos de Ejemplo

### Transacción 1: Compra a Zona Franca (Sin IVA)
```
Fecha: 2024-11-01
Entidad: VWR
Categoría: Inventario
Subcategoría: Materiales Lab
Moneda: USD
Monto: 2500.00
IVA: No  ← Zona franca
Estado: Pagado
Personal/Negocio: Negocio
```

### Transacción 2: Pago de Servicio con Tarjeta
```
Fecha: 2024-11-05
Entidad: Visa-BAC
Categoría: Servicios
Subcategoría: Internet
Moneda: CRC
Monto: 45000
IVA: Sí
Estado: Pagado
Personal/Negocio: Negocio
```

### Transacción 3: Ingreso por Consultoría
```
Fecha: 2024-11-10
Entidad: Cliente-XYZ
Categoría: Ingreso
Subcategoría: Servicios Profesionales
Moneda: USD
Monto: 5000.00
IVA: Sí
Estado: Cobrado
Personal/Negocio: Negocio
```

---

## 🎯 Próximos Pasos

### Fase 1: Configuración Inicial (AHORA)
- [x] Generar archivo Excel v4.0
- [ ] Revisar las 3 transacciones de ejemplo
- [ ] Ajustar tipo de cambio en CONFIG si es necesario
- [ ] Familiarizarse con las 14 hojas

### Fase 2: Migración de Datos (PRÓXIMO)
- [ ] Cargar datos reales de **CxP** (~$6,500 en 10 facturas)
- [ ] Cargar datos reales de **CxC** ($9,923 pendiente)
- [ ] Registrar deudas de tarjetas (₡6.9M total)
- [ ] Validar que todas las fórmulas funcionen correctamente

### Fase 3: Uso Diario (FUTURO)
- [ ] Registrar transacciones diarias en TRANSACCIONES
- [ ] Revisar RESUMEN semanalmente
- [ ] Conciliar bancos mensualmente (hoja CONCILIACION)
- [ ] Declarar IVA usando hoja IVA_CONTROL
- [ ] Ajustar presupuestos trimestralmente

### Fase 4: Optimización (OPCIONAL)
- [ ] Agregar gráficos en DASHBOARD_VISUAL
- [ ] Crear reporte mensual automatizado
- [ ] Integrar con API bancaria (futuro)
- [ ] Exportar a software contable (futuro)

---

## 📚 Documentación Adicional

- **LESSONS_LEARNED.md** - Errores que cometimos y cómo los resolvimos
- **docs/architecture.md** - Arquitectura técnica detallada
- **docs/user_guide.md** - Guía paso a paso para usuarios

---

## 🛠️ Soporte y Contacto

### Reportar Problemas
Si encuentras algún error:
1. Verifica que tienes Python 3.11+ y openpyxl instalado
2. Revisa LESSONS_LEARNED.md para errores comunes
3. Documenta el error con captura de pantalla

### Autor
**Álvaro Velasco**
Proyecto desarrollado en colaboración con Claude AI
Noviembre 2024

---

## 📄 Licencia

Este proyecto es de uso personal y privado.
© 2024 Álvaro Velasco. Todos los derechos reservados.

---

## ✅ Checklist de Validación

Antes de usar el sistema, verifica:

- [ ] Python 3.11+ instalado
- [ ] openpyxl instalado (`pip install openpyxl`)
- [ ] Archivo Excel generado exitosamente
- [ ] 14 hojas visibles en el Excel
- [ ] 3 transacciones de ejemplo presentes
- [ ] Hojas TRANSACCIONES y CONFIG editables (sin protección)
- [ ] 12 hojas restantes protegidas
- [ ] Tipo de cambio en CONFIG = 540 (o ajustado)
- [ ] Fórmulas funcionando (RESUMEN muestra datos)

---

**¡Sistema de Finanzas v4.0 - Listo para Producción!** 🚀
