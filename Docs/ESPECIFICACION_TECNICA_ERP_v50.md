# 📐 ESPECIFICACIÓN TÉCNICA DETALLADA - ERP v5.0

**Proyecto:** AlvaroVelasco_Finanzas_v5.0.xlsx
**Autor:** Alvaro Velasco | Net SRL
**Fecha:** 16 de noviembre, 2025
**Versión:** 5.0 Final
**Excel:** Office 365 en INGLÉS

---

## ⚠️ CHECKLIST PRE-GENERACIÓN COMPLETADO

- [x] **Idioma Excel:** INGLÉS (Office 365)
- [x] **Fórmulas:** SUM, IF, SUMIF, FILTER (NO SUMA)
- [x] **Separador:** `,` (coma)
- [x] **Hojas:** 15 confirmadas
- [x] **Datos:** Actuales y verificados
- [x] **Consistencia:** Español (nombres) + Inglés (fórmulas)
- [x] **Imports:** Completos al inicio
- [x] **Validaciones:** Estrictas

---

## 📋 ÍNDICE DE HOJAS (15 TOTAL)

| # | Hoja | Tipo | Propósito |
|---|------|------|-----------|
| 1 | RESUMEN | Auto | Dashboard ejecutivo KPIs |
| 2 | TRANSACCIONES | ✏️ Editable | Single Source of Truth (16+1 cols) |
| 3 | CONFIG | ✏️ Editable | Parámetros configurables |
| 4 | CxP | Auto | Cuentas por Pagar (FILTER dinámico) |
| 5 | CxC | Auto | Cuentas por Cobrar (FILTER dinámico) |
| 6 | FLUJO_CAJA | Auto | Cash flow mensual |
| 7 | IVA_CONTROL | Auto | Gestión IVA 13% (excluye zona franca) |
| 8 | TARJETAS | Auto | Control 4 tarjetas + días corte |
| 9 | CONCILIACION | Auto | Conciliación bancaria |
| 10 | PERSONAL_VS_NEGOCIO | Auto | Separación gastos |
| 11 | CATEGORIAS | Auto | Análisis por categoría |
| 12 | PROYECTOS | Auto | Centros de costo |
| 13 | PROVEEDORES | Auto | Análisis proveedores |
| 14 | CLIENTES | Auto | Análisis clientes |
| 15 | AUDITORIA | Auto | Detección anomalías |

---

## 🏗️ HOJA 1: RESUMEN (Dashboard)

### Propósito
Vista consolidada en tiempo real de KPIs operacionales y financieros.

### Layout

```
┌─────────────────────────────────────────────────────────────┐
│ A1:H1 (merged)                                              │
│ DASHBOARD FINANCIERO - AlvaroVelasco Net SRL               │
│ [Estilo: Header dark, altura 35px]                         │
├─────────────────────────────────────────────────────────────┤
│ A2:H2 (merged)                                              │
│ Última actualización: 16/11/2025 14:30                     │
│ [Estilo: Italic, size 9]                                   │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ A4:D4 (merged) - CUENTAS POR PAGAR (CxP)                   │
│ [Fondo: Rojo #FF5252]                                      │
│                                                             │
│ A5: Total CxP USD:    B5: =SUMIF(CxP!D:D,"USD",CxP!C:C)   │
│ A6: Total CxP CRC:    B6: =SUMIF(CxP!D:D,"CRC",CxP!C:C)   │
│ A7: Facturas vencidas: B7: =COUNTIF(CxP!F:F,">0")         │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│ A9:D9 (merged) - CUENTAS POR COBRAR (CxC)                  │
│ [Fondo: Verde #4CAF50]                                     │
│                                                             │
│ A10: Total CxC USD:   B10: =SUMIF(CxC!D:D,"USD",CxC!C:C)  │
│ A11: Total CxC CRC:   B11: =SUMIF(CxC!D:D,"CRC",CxC!C:C)  │
│ A12: Facturas >60 días: B12: =COUNTIF(CxC!F:F,">60")      │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│ A14:D14 (merged) - FLUJO DE CAJA (MES ACTUAL)              │
│                                                             │
│ A15: Ingresos mes USD: B15: =SUMIFS(...)                  │
│ A16: Gastos mes USD:   B16: =SUMIFS(...)                  │
│ A17: Balance mes USD:  B17: =B15-B16                       │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│ A19:D19 (merged) - IVA CONTROL                             │
│                                                             │
│ A20: IVA Ventas:      B20: =SUM(IVA_CONTROL!...)          │
│ A21: IVA Compras:     B21: =SUM(IVA_CONTROL!...)          │
│ A22: A pagar Hacienda: B22: =B20-B21                       │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Columnas

| Col | Ancho | Contenido |
|-----|-------|-----------|
| A | 30 | Labels |
| B | 20 | Valores calculados |
| C-H | Auto | Espaciado |

### Formatos Número

- Montos USD: `$#,##0.00`
- Montos CRC: `₡#,##0.00`
- Contadores: `0`

---

## 🏗️ HOJA 2: TRANSACCIONES (Core - Single Source of Truth)

### Propósito
**ÚNICA HOJA EDITABLE PARA DATOS.** Todas las demás hojas se calculan desde aquí.

### Layout

```
Fila 1: Título (A1:Q1 merged)
"TRANSACCIONES - ÚNICA FUENTE DE VERDAD"
[Fondo: Azul #366092, altura 30px]

Fila 2: Headers (17 columnas)
[Fondo: Azul medio #4472C4, texto blanco]
```

### Columnas (17 TOTAL: 16 inputs + 1 validación)

| Col | Campo | Tipo | Ancho | Validación | Formato | Comentario |
|-----|-------|------|-------|------------|---------|------------|
| **A** | Fecha | Date | 12 | Required, Date > 2020-01-01 | DD/MM/YY | "Formato: DD/MM/YYYY" |
| **B** | Entidad | Text | 25 | Required, Not empty | Text | "Banco/Proveedor/Cliente" |
| **C** | Categoría | List | 15 | Dropdown | Text | "Income, Expense, CxP, CxC..." |
| **D** | Subcategoría | Text | 18 | Optional | Text | "Detalle de categoría" |
| **E** | Moneda | List | 10 | Dropdown (USD/CRC) | Text | "USD o CRC" |
| **F** | Monto | Number | 15 | > 0 (decimal) | $#,##0.00 o ₡#,##0.00 | "Valor > 0" |
| **G** | Descripción | Text | 30 | Optional | Text | "Detalle operación" |
| **H** | Forma Pago | List | 15 | Dropdown | Text | "Cash, Transfer, CC..." |
| **I** | IVA | List | 8 | Dropdown (Yes/No) | Text | "Yes/No para IVA 13%" |
| **J** | Notas | Text | 25 | Optional | Text | "Observaciones" |
| **K** | Recurrente | List | 12 | Dropdown (Yes/No) | Text | "Yes/No" |
| **L** | Proyecto | Text | 15 | Optional | Text | "Centro costo" |
| **M** | Estado | List | 12 | Dropdown | Text | "Paid, Pending, Collected" |
| **N** | Factura # | Text | 15 | Optional | Text | "# Documento" |
| **O** | Tag | Text | 15 | Optional | Text | "Etiquetas" |
| **P** | Personal/Negocio | List | 18 | Dropdown | Text | "Personal o Business" |
| **Q** | ✓ Validación | Formula | 20 | AUTO (no editable) | Text | "✓ OK o ✗ ERROR" |

### Fórmula Columna Q (Validación Automática)

```excel
=IF(AND(
  A3<>"",
  ISNUMBER(A3),
  B3<>"",
  C3<>"",
  OR(E3="USD",E3="CRC"),
  F3>0
), "✓ OK", "✗ ERROR: Revisa datos")
```

**Formato condicional Q:**
- Si contiene "✓ OK" → Fondo verde #C6EFCE
- Si contiene "✗ ERROR" → Fondo rojo #FFC7CE

### Dropdowns (Data Validation)

#### Categoría (C3:C1000)
```
Income, Expense, Inventory, Services, CxP, CxC, Marketing, Personal
```

#### Moneda (E3:E1000)
```
USD, CRC
```

#### Forma Pago (H3:H1000)
```
Cash, Transfer, Credit Card, Check, SINPE, PayPal
```

#### IVA (I3:I1000)
```
Yes, No
```

#### Recurrente (K3:K1000)
```
Yes, No
```

#### Estado (M3:M1000)
```
Paid, Pending, Collected, Canceled
```

#### Personal/Negocio (P3:P1000)
```
Personal, Business
```

### Datos Pre-cargados (Filas 3-6)

| Fila | Fecha | Entidad | Cat | SubCat | Mon | Monto | Desc | Pago | IVA | Notas | Rec | Proy | Estado | Fact | Tag | P/N |
|------|-------|---------|-----|--------|-----|-------|------|------|-----|-------|-----|------|--------|------|-----|-----|
| 3 | HOY | BAC San José | CxP | Tarjeta Crédito | CRC | 1831000 | Saldo tarjeta Visa Clásica | Credit Card | No | Pre-cargado - Corte día 15 | No | | Pending | | Deuda Inicial | Business |
| 4 | HOY | BCR | CxP | Tarjeta Crédito | CRC | 1750000 | Saldo tarjeta Mastercard | Credit Card | No | Pre-cargado - Corte día 20 | No | | Pending | | Deuda Inicial | Business |
| 5 | HOY | Credomatic | CxP | Tarjeta Crédito | USD | 2500 | Saldo tarjeta Platinum | Credit Card | No | Pre-cargado - Corte día 10 | No | | Pending | | Deuda Inicial | Business |
| 6 | HOY | Credomatic | CxP | Tarjeta Crédito | CRC | 2800000 | Saldo tarjeta Gold | Credit Card | No | Pre-cargado - Corte día 10 | No | | Pending | | Deuda Inicial | Business |

---

## 🏗️ HOJA 3: CONFIG

### Propósito
Parámetros configurables del sistema. Única hoja editable además de TRANSACCIONES.

### Layout

```
A1:C1 (merged): "CONFIGURACIÓN DEL SISTEMA"
[Fondo: Naranja #FF6600, altura 30px]

A2:C2 (merged): "⚠️ Celdas amarillas son EDITABLES - No modificar fórmulas"
[Texto rojo, bold]

Fila 4: Headers
A4: Parámetro | B4: Valor | C4: Descripción
```

### Parámetros (Filas 5-12)

| Fila | A (Parámetro) | B (Valor) | C (Descripción) |
|------|---------------|-----------|-----------------|
| 5 | Tipo Cambio USD→CRC | **540** | Tipo de cambio actual |
| 6 | BAC Visa - Día Corte | **15** | Día de corte mensual |
| 7 | BCR Mastercard - Día Corte | **20** | Día de corte mensual |
| 8 | Credomatic Platinum - Día Corte | **10** | Día de corte mensual |
| 9 | Credomatic Gold - Día Corte | **10** | Día de corte mensual |
| 10 | Proveedor Zona Franca 1 | **VWR International LLC** | Exento de IVA |
| 11 | Proveedor Zona Franca 2 | **RS Hughes Co. Inc.** | Exento de IVA |
| 12 | IVA Costa Rica % | **13** | Porcentaje IVA estándar |

**Formato columna B:**
- Fondo: Amarillo #FFFF00 (indica editable)
- Font: Bold
- Números: General
- Texto: Text

**Anchos:**
- A: 35px
- B: 25px
- C: 40px

---

## 🏗️ HOJA 4: CxP (Cuentas por Pagar)

### Propósito
Auto-calculada desde TRANSACCIONES usando FILTER(). Muestra todas las obligaciones pendientes.

### Layout

```
A1:H1 (merged): "CUENTAS POR PAGAR (CxP)"
[Fondo: Rojo #FF5252, altura 30px]

Fila 2: Headers
```

### Columnas

| Col | Campo | Fórmula | Formato | Ancho |
|-----|-------|---------|---------|-------|
| A | Entidad | FILTER (ver abajo) | Text | 25 |
| B | Fecha | FILTER | DD/MM/YY | 12 |
| C | Monto | FILTER | $#,##0.00 | 15 |
| D | Moneda | FILTER | Text | 10 |
| E | Descripción | FILTER | Text | 30 |
| F | Estado | FILTER | Text | 12 |
| G | Días Vencidos | Calculado | 0 | 15 |
| H | Equiv USD | Calculado | $#,##0.00 | 15 |

### Fórmula FILTER Principal (A3)

```excel
=FILTER(
  TRANSACCIONES!B:B,
  (TRANSACCIONES!C:C="CxP") * (TRANSACCIONES!M:M="Pending"),
  "Sin CxP pendientes"
)
```

**Nota:** Cada columna B, C, D, E, F usa FILTER similar pero referencia la columna correspondiente.

### Columna G (Días Vencidos)

```excel
=IF(A3<>"", TODAY() - B3, "")
```

**Formato condicional:**
- Si > 30 → Fondo rojo
- Si > 0 AND <= 30 → Fondo amarillo
- Si <= 0 → Sin formato

### Columna H (Equiv USD)

```excel
=IF(D3="USD", C3, C3/CONFIG!$B$5)
```

### Totales (Fila final + 2)

```excel
A: "TOTAL CxP"
C: =SUM(C3:C100)  // Ajustar rango según necesidad
H: =SUM(H3:H100)
```

**Protección:** Hoja protegida (sin password, solo lectura)

---

## 🏗️ HOJA 5: CxC (Cuentas por Cobrar)

### Igual que CxP pero filtra:
```excel
=FILTER(
  TRANSACCIONES!B:B,
  (TRANSACCIONES!C:C="CxC") * (TRANSACCIONES!M:M="Pending"),
  "Sin CxC pendientes"
)
```

**Alerta días pendientes:**
- > 60 días → Fondo rojo
- 30-60 días → Fondo amarillo
- < 30 días → Sin formato

---

## 🏗️ HOJA 6: FLUJO_CAJA

### Propósito
Proyección mensual de ingresos vs gastos.

### Columnas

| Col | Campo | Fórmula | Ancho |
|-----|-------|---------|-------|
| A | Mes | Texto ("Enero", "Febrero"...) | 15 |
| B | Ingresos USD | =SUMIFS(TRANSACCIONES!F:F, TRANSACCIONES!C:C, "Income", TRANSACCIONES!E:E, "USD", TRANSACCIONES!A:A, ">=FECHA_INICIO_MES", TRANSACCIONES!A:A, "<=FECHA_FIN_MES") | 18 |
| C | Gastos USD | Similar a B pero "Expense" | 18 |
| D | Flujo Neto | =B3-C3 | 18 |
| E | Acumulado | =IF(A3="Enero", D3, E2+D3) | 18 |

**Formato:**
- Números: $#,##0.00
- Fila total: Bold, fondo azul

---

## 🏗️ HOJA 7: IVA_CONTROL

### Propósito
Calcular IVA 13% EXCLUYENDO proveedores zona franca.

### Columnas

| Col | Campo | Fórmula | Ancho |
|-----|-------|---------|-------|
| A | Fecha | Referencia TRANSACCIONES | 12 |
| B | Entidad | Referencia | 25 |
| C | Monto Base | Referencia | 15 |
| D | IVA Aplicable | =IF(AND(TRANSACCIONES!I="Yes", NOT(OR(B3=CONFIG!$B$10, B3=CONFIG!$B$11))), C3*0.13, 0) | 15 |
| E | Tipo | "Venta" o "Compra" | 12 |

**Fórmula explicada columna D:**
- SI IVA="Yes" en TRANSACCIONES
- Y Entidad NO es zona franca (CONFIG B10 o B11)
- ENTONCES Monto * 13%
- SINO 0

### Totales

```
IVA Ventas: =SUMIF(E:E, "Venta", D:D)
IVA Compras: =SUMIF(E:E, "Compra", D:D)
A pagar Hacienda: =IVA_Ventas - IVA_Compras
```

---

## 🏗️ HOJA 8: TARJETAS

### Propósito
Control individual de 4 tarjetas con días hasta próximo corte.

### Layout

```
Fila 1: Header

Filas 3-6: Una fila por tarjeta

Columnas:
A: Tarjeta (nombre)
B: Saldo Actual (SUMIF desde TRANSACCIONES)
C: Día Corte (desde CONFIG)
D: Días hasta corte (fórmula)
E: Próxima fecha corte (calculada)
F: Pagos mes actual (SUMIFS)
```

### Fórmula Días hasta Corte (D3)

```excel
=IF(DAY(TODAY()) <= CONFIG!B6,
    CONFIG!B6 - DAY(TODAY()),
    (DAY(EOMONTH(TODAY(),0)) - DAY(TODAY())) + CONFIG!B6
)
```

**Formato condicional:**
- Si < 5 días → Fondo rojo (urgente)
- Si 5-10 días → Fondo amarillo
- Si > 10 días → Verde

---

## 🏗️ HOJA 9: CONCILIACION

### Propósito
Conciliación bancaria mensual.

### Layout

```
SECCIÓN 1: SALDO BANCO
A5: Saldo según banco (input manual)
A6: Menos: Cheques en tránsito
A7: Más: Depósitos en tránsito
A8: = Saldo ajustado banco

SECCIÓN 2: SALDO LIBROS
A12: Saldo según TRANSACCIONES (auto)
A13: Menos: Cargos bancarios no registrados
A14: Más: Intereses no registrados
A15: = Saldo ajustado libros

SECCIÓN 3: DIFERENCIA
A19: Diferencia (=A8-A15)
```

**Alerta:**
- Si diferencia = 0 → Verde "✓ Conciliado"
- Si diferencia <> 0 → Rojo "✗ Revisar diferencia"

---

## 🏗️ HOJA 10: PERSONAL_VS_NEGOCIO

### Propósito
Separar gastos personales vs negocio.

### Columnas

| Tipo | Ingresos | Gastos | Balance |
|------|----------|--------|---------|
| PERSONAL | =SUMIFS(TRANSACCIONES!F:F, P:P, "Personal", C:C, "Income") | =SUMIFS(..., "Expense") | =Ing-Gas |
| BUSINESS | =SUMIFS(..., "Business", ...) | | |
| TOTAL | =SUM | =SUM | =SUM |

**Formato:**
- Ingresos: Verde
- Gastos: Naranja
- Balance positivo: Verde bold
- Balance negativo: Rojo bold

---

## 🏗️ HOJA 11: CATEGORIAS

### Propósito
Análisis de gastos por categoría.

### Layout

```
Top 10 categorías con más gasto

Columnas:
A: Categoría
B: Total Gastado
C: % del Total
D: # Transacciones
E: Promedio por transacción
```

**Fórmulas:**
- B: SUMIF por categoría
- C: =B/SUM($B:$B)
- D: COUNTIF
- E: =B/D

**Ordenar:** Por columna B descendente

---

## 🏗️ HOJA 12: PROYECTOS

### Similar a CATEGORIAS pero agrupado por columna L (Proyecto)

---

## 🏗️ HOJA 13: PROVEEDORES

### Layout

```
Top 20 proveedores

Columnas:
A: Proveedor (Entidad)
B: Total Comprado
C: # Facturas
D: Promedio Factura
E: Última Compra (MAX fecha)
F: Días desde última compra
```

---

## 🏗️ HOJA 14: CLIENTES

### Similar a PROVEEDORES pero filtrando Income/CxC

---

## 🏗️ HOJA 15: AUDITORIA (NUEVA)

### Propósito
Detectar automáticamente anomalías en los datos.

### Secciones

#### 1. Errores Críticos
- Montos negativos
- Fechas futuras (> HOY + 30)
- Moneda inválida
- Entidad vacía

#### 2. Alertas
- Montos outliers (> $50,000)
- Zona franca con IVA=Yes
- Transacciones incompletas (columna Q con ERROR)

#### 3. Inconsistencias
- CxP/CxC con Estado=Paid (deberían estar Pending)
- Montos redondos sospechosos (múltiplos exactos de 1000)

### Columnas

| Col | Campo | Descripción |
|-----|-------|-------------|
| A | Tipo | "ERROR", "ALERTA", "INFO" |
| B | Fila | # fila en TRANSACCIONES |
| C | Campo | Columna con problema |
| D | Problema | Descripción |
| E | Valor Actual | Valor problemático |
| F | Acción Sugerida | Qué hacer |

### Fórmulas Ejemplo

```excel
// Detectar montos negativos
=IF(TRANSACCIONES!F3<0, "ERROR - Monto negativo", "")

// Detectar zona franca con IVA
=IF(AND(
  OR(TRANSACCIONES!B3=CONFIG!B10, TRANSACCIONES!B3=CONFIG!B11),
  TRANSACCIONES!I3="Yes"
), "ALERTA - Zona franca no debe tener IVA", "")
```

---

## 📊 RESUMEN TÉCNICO

### Total Columnas
- TRANSACCIONES: **17** (16 input + 1 validación)
- Otras hojas: Variable según necesidad

### Total Fórmulas Únicas
- **~50-60 fórmulas** distintas en total
- Uso intensivo de: FILTER(), SUMIF(), SUMIFS(), IF(), AND(), OR()

### Formato Números Estándar
- USD: `$#,##0.00`
- CRC: `₡#,##0.00`
- Fechas: `DD/MM/YY`
- Porcentajes: `0.0%`

### Protección
- Hojas 1, 4-15: **Protegidas** (sin password, solo lectura)
- Hojas 2, 3: **Sin protección** (editables)

---

## ✅ VALIDACIÓN PRE-CODIFICACIÓN

**Antes de generar código Python, confirmar:**

1. ¿Esta especificación cubre TODO lo que necesitas?
2. ¿Alguna fórmula debe cambiarse?
3. ¿Alguna columna falta o sobra?
4. ¿Los formatos de número son correctos?
5. ¿Los anchos de columna son adecuados?

**Una vez aprobada esta especificación, procedo a generar el código Python completo (1500-2000 líneas) que implementa EXACTAMENTE esto.**

---

**Especificación creada por:** Claude (Sonnet 4.5)
**Fecha:** 16 de noviembre, 2025
**Status:** PENDIENTE APROBACIÓN DEL USUARIO
