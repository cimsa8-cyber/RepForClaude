# 📐 INVENTARIO DE FÓRMULAS CRÍTICAS

**Proyecto:** Excel Finance ERP v5.0
**Última actualización:** 18 de noviembre, 2025
**Total fórmulas documentadas:** 15+ críticas

---

## ⚠️ REGLAS PARA MODIFICAR FÓRMULAS

1. **NUNCA modificar sin entender** qué hace la fórmula
2. **BACKUP antes de cambiar** cualquier fórmula
3. **Probar en celda de prueba** antes de aplicar masivamente
4. **Documentar cambios** en commit message
5. **Verificar separadores** (`;` vs `,`) según idioma Excel

---

## 🔴 FÓRMULAS FILTER (Excel 365 Dynamic Arrays)

### CxP - Filtrar Cuentas por Pagar Pendientes

**Ubicación:** Hoja `CxP`, Celda `A3`

**Fórmula (openpyxl - NO FUNCIONA):**
```excel
=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxP pendientes")
```

**Fórmula (xlwings - FUNCIONA):**
```excel
=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxP pendientes")
```

**Qué hace:**
- Filtra la columna `B` (Entidad) de TRANSACCIONES
- **Condición 1:** Columna `C` (Categoría) = "CxP"
- **Condición 2:** Columna `M` (Estado) = "Pendiente"
- **Operador:** `*` (AND lógico)
- **Si vacío:** Muestra "Sin CxP pendientes"

**Columnas referenciadas:**
- `B3:B1000` - Entidad (resultado)
- `C3:C1000` - Categoría (filtro)
- `M3:M1000` - Estado (filtro)

**Rango:** 997 filas (B3 a B1000)

**Resultado:** Array dinámico que se expande verticalmente

**Separadores:**
- Excel español: **punto y coma (;)**
- Excel inglés: **coma (,)**

**Problema conocido:**
- openpyxl genera con comas → Excel español lo rechaza
- xlwings usa Excel real → funciona con cualquier separador

---

### CxP - Columnas B-K (11 fórmulas FILTER adicionales)

Misma estructura que A3, pero filtrando diferentes columnas:

| Celda | Columna Resultado | Contenido |
|-------|-------------------|-----------|
| `B3` | Fecha | `TRANSACCIONES!A3:A1000` |
| `C3` | Monto | `TRANSACCIONES!F3:F1000` |
| `D3` | Moneda | `TRANSACCIONES!E3:E1000` |
| `E3` | Descripción | `TRANSACCIONES!G3:G1000` |
| `F3` | Estado | `TRANSACCIONES!M3:M1000` |
| `G3` | Días Pendientes | `TRANSACCIONES!S3:S1000` |
| `H3` | Equiv USD | `TRANSACCIONES!T3:T1000` |
| `I3` | Factura # | `TRANSACCIONES!N3:N1000` |
| `J3` | Fecha Venc | `TRANSACCIONES!U3:U1000` |
| `K3` | Prioridad | `TRANSACCIONES!V3:V1000` |

**Todas usan:** `(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente")`

---

### CxC - Filtrar Cuentas por Cobrar Pendientes

**Ubicación:** Hoja `CxC`, Celda `A3`

**Fórmula:**
```excel
=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxC pendientes")
```

**Diferencia vs CxP:** Cambia filtro de "CxP" a "CxC"

**CxC tiene 11 columnas adicionales** (B3-K3) similar a CxP

**Nota:** Columna `K3` usa `TRANSACCIONES!W3:W1000` (Prioridad CxC, no V)

---

## 🟢 FÓRMULAS SUMIF/SUMIFS

### RESUMEN - Total CxP USD

**Ubicación:** Hoja `RESUMEN`, Celda `B5`

**Fórmula:**
```excel
=SUMIF(CxP!D:D,"USD",CxP!H:H)
```

**Qué hace:**
- Suma la columna `H` (Equiv USD) de la hoja CxP
- **Condición:** Columna `D` (Moneda) = "USD"
- **Resultado:** Total USD en cuentas por pagar

**Columnas referenciadas:**
- `CxP!D:D` - Moneda (criterio)
- `CxP!H:H` - Equiv USD (suma)

**Nota:** Usa columnas completas (D:D) en vez de rango específico

---

### RESUMEN - Total CxP CRC

**Ubicación:** Hoja `RESUMEN`, Celda `B6`

**Fórmula:**
```excel
=SUMIF(CxP!D:D,"CRC",CxP!C:C)
```

**Qué hace:**
- Suma la columna `C` (Monto) de la hoja CxP
- **Condición:** Columna `D` (Moneda) = "CRC"
- **Resultado:** Total CRC en cuentas por pagar

---

### RESUMEN - Total CxC USD

**Ubicación:** Hoja `RESUMEN`, Celda `B9`

**Fórmula:**
```excel
=SUMIF(CxC!D:D,"USD",CxC!H:H)
```

**Similar a CxP pero para cuentas por cobrar**

---

### PERSONAL_VS_NEGOCIO - Ingresos Personales

**Ubicación:** Hoja `PERSONAL_VS_NEGOCIO`, Celda `B3`

**Fórmula:**
```excel
=SUMIFS(TRANSACCIONES!F:F,TRANSACCIONES!P:P,"Personal",TRANSACCIONES!C:C,"Ingreso")
```

**Qué hace:**
- Suma la columna `F` (Monto) de TRANSACCIONES
- **Condición 1:** Columna `P` (Personal/Negocio) = "Personal"
- **Condición 2:** Columna `C` (Categoría) = "Ingreso"
- **Resultado:** Total ingresos personales

**Diferencia SUMIF vs SUMIFS:**
- `SUMIF`: 1 condición
- `SUMIFS`: Múltiples condiciones (AND lógico)

---

### PERSONAL_VS_NEGOCIO - Gastos Personales

**Ubicación:** Hoja `PERSONAL_VS_NEGOCIO`, Celda `C3`

**Fórmula:**
```excel
=SUMIFS(TRANSACCIONES!F:F,TRANSACCIONES!P:P,"Personal",TRANSACCIONES!C:C,"Gasto")
```

**Diferencia vs B3:** Cambia "Ingreso" a "Gasto"

---

## 🔵 FÓRMULAS IF CONDICIONALES

### IVA_CONTROL - IVA Aplicable

**Ubicación:** Hoja `IVA_CONTROL`, Celda `D3`

**Fórmula:**
```excel
=IF(AND(TRANSACCIONES!I3="Sí",NOT(OR(B3=CONFIG!$B$10,B3=CONFIG!$B$11))),C3*0.13,0)
```

**Qué hace:**
1. **Verifica dos condiciones (AND):**
   - `TRANSACCIONES!I3="Sí"` → IVA está marcado como "Sí"
   - `NOT(OR(B3=CONFIG!$B$10,B3=CONFIG!$B$11))` → NO es proveedor zona franca

2. **Si AMBAS son verdaderas:**
   - Calcula IVA: `C3 * 0.13` (13% del monto base)

3. **Si alguna es falsa:**
   - Devuelve `0` (sin IVA)

**Zona Franca (exentos IVA):**
- `CONFIG!$B$10` = VWR International LLC
- `CONFIG!$B$11` = RS Hughes Co. Inc.

**Lógica:**
```
SI (IVA="Sí" Y NO es zona franca)
   ENTONCES: Monto * 13%
   SINO: 0
```

**Columnas referenciadas:**
- `TRANSACCIONES!I3` - IVA (Sí/No)
- `B3` - Entidad (comparar con zona franca)
- `C3` - Monto Base
- `CONFIG!$B$10, $B$11` - Proveedores zona franca

**Nota:** `$` hace referencia absoluta (no cambia al copiar fórmula)

---

### TRANSACCIONES - Equivalente USD

**Ubicación:** Hoja `TRANSACCIONES`, Columna `T` (calculada)

**Fórmula (fila 3):**
```excel
=IF(E3="USD",F3,F3/CONFIG!$B$5)
```

**Qué hace:**
1. **Si moneda es USD:** Devuelve monto tal cual
2. **Si moneda NO es USD (CRC):** Divide monto entre tipo de cambio

**Ejemplo:**
- `E3 = "USD", F3 = 1000` → Resultado: `1000`
- `E3 = "CRC", F3 = 540000, TC = 540` → Resultado: `1000` (540000/540)

**Columnas referenciadas:**
- `E3` - Moneda (USD/CRC)
- `F3` - Monto original
- `CONFIG!$B$5` - Tipo Cambio USD→CRC (540)

---

## 🟡 FÓRMULAS CALCULADAS (TRANSACCIONES)

### Columna S - Días Transcurridos

**Ubicación:** Hoja `TRANSACCIONES`, Columna `S`

**Fórmula (fila 3):**
```excel
=TODAY()-A3
```

**Qué hace:**
- Calcula días desde la fecha de transacción hasta hoy
- `TODAY()` devuelve fecha actual
- Resultado: Número de días

**Ejemplo:**
- Hoy: 18/11/2025
- A3: 10/11/2025
- Resultado: 8 días

---

### Columna U - Fecha Vencimiento

**Ubicación:** Hoja `TRANSACCIONES`, Columna `U`

**Fórmula (fila 3):**
```excel
=A3+30
```

**Qué hace:**
- Suma 30 días a la fecha de transacción
- Asume plazo estándar de 30 días

**Ejemplo:**
- A3: 10/11/2025
- Resultado: 10/12/2025

**Mejora futura:** Podría leer plazo de CONFIG en vez de hardcoded 30

---

### Columna V - Prioridad CxP

**Ubicación:** Hoja `TRANSACCIONES`, Columna `V`

**Fórmula (fila 3):**
```excel
=IF(C3="CxP",IF(S3>60,"Alta",IF(S3>30,"Media","Baja")),"")
```

**Qué hace:**
1. **Si es CxP:**
   - Más de 60 días → "Alta"
   - 30-60 días → "Media"
   - Menos de 30 días → "Baja"

2. **Si NO es CxP:** Vacío

**Lógica anidada (3 niveles):**
```
SI C3="CxP"
   SI S3>60 → "Alta"
   SINO
      SI S3>30 → "Media"
      SINO → "Baja"
SINO → ""
```

**Columnas referenciadas:**
- `C3` - Categoría
- `S3` - Días Transcurridos (calculado)

---

### Columna W - Prioridad CxC

**Ubicación:** Hoja `TRANSACCIONES`, Columna `W`

**Fórmula (fila 3):**
```excel
=IF(C3="CxC",IF(S3>90,"Alta",IF(S3>60,"Media","Baja")),"")
```

**Diferencia vs Prioridad CxP:**
- CxC usa umbrales diferentes (90/60 en vez de 60/30)
- CxC permite más días antes de ser "Alta" prioridad

---

## 🔴 FÓRMULAS DE VALIDACIÓN

### TRANSACCIONES - Validación de Datos

**Ubicación:** Hoja `TRANSACCIONES`, Columna `R`

**Fórmula (fila 3):**
```excel
=IF(AND(A3<>"",ISNUMBER(A3),B3<>"",C3<>"",OR(E3="USD",E3="CRC"),F3>0),"✓ OK","✗ ERROR: Revisa datos")
```

**Qué hace:**
Valida que todos los campos requeridos estén correctos:

1. `A3<>""` - Fecha no vacía
2. `ISNUMBER(A3)` - Fecha es número (fecha válida)
3. `B3<>""` - Entidad no vacía
4. `C3<>""` - Categoría no vacía
5. `OR(E3="USD",E3="CRC")` - Moneda es USD o CRC
6. `F3>0` - Monto es positivo

**Si TODAS son verdaderas:** "✓ OK"
**Si alguna es falsa:** "✗ ERROR: Revisa datos"

**Uso:** Permite identificar rápidamente filas con datos incompletos

---

## 🟣 FÓRMULAS DE CONCILIACIÓN

### CONCILIACION - Diferencia

**Ubicación:** Hoja `CONCILIACION`, Celda `B19`

**Fórmula:**
```excel
=B8-B15
```

**Qué hace:**
- `B8` = Saldo ajustado según banco
- `B15` = Saldo ajustado según libros
- **Resultado:** Diferencia (debe ser 0 si todo cuadra)

**Concepto contable:**
Si B19 ≠ 0 → Hay error de conciliación que investigar

---

## 📊 FÓRMULAS DE REFERENCIAS CIRCULARES (RESUELTAS)

### RESUMEN - Ratio Liquidez (ANTES - ERROR)

**Fórmula errónea:**
```excel
=IFERROR(B19/B20,0)
```

**Problema:** B20 referenciaba a sí misma → circular reference

**Fórmula corregida:**
```excel
=IFERROR(B18/B19,0)
```

**Qué hace:**
- Divide `B18` (Activo Corriente) entre `B19` (Pasivo Corriente)
- Ratio > 1.5 es saludable
- Si error (división por 0), devuelve 0

**Commit fix:** `ff9bb74`

---

## 🎯 CONVENCIONES DE FÓRMULAS

### Uso de $ (Referencias Absolutas)

```excel
$B$5    → Fila Y columna fijas (no cambia al copiar)
B$5     → Fila fija, columna relativa
$B5     → Columna fija, fila relativa
B5      → Ambas relativas (cambian al copiar)
```

**Ejemplo:**
- `CONFIG!$B$5` → Siempre apunta a B5 de CONFIG
- Útil para: Tipo cambio, parámetros globales

---

### Separadores según Idioma

| Idioma | Separador args | Separador decimal | Ejemplo |
|--------|----------------|-------------------|---------|
| Inglés | `,` (coma) | `.` (punto) | `=IF(A1>0,B1,0)` |
| Español | `;` (punto y coma) | `,` (coma) | `=SI(A1>0;B1;0)` |

**Crítico:** Excel rechaza fórmulas con separadores incorrectos

---

## 📝 CHECKLIST ANTES DE MODIFICAR FÓRMULA

```markdown
[ ] Entiendo qué hace la fórmula
[ ] Leí la documentación en este archivo
[ ] Hice backup del archivo
[ ] Probé en celda de prueba primero
[ ] Verifiqué separadores (';' vs ',')
[ ] Documenté el cambio en commit
[ ] Probé que sigue funcionando
```

---

## 🔧 HERRAMIENTAS DE DEBUGGING

### Ver todas las fórmulas en Excel

1. Abrir Excel
2. Ir a pestaña "Fórmulas"
3. Click en "Mostrar fórmulas" (Ctrl + `)
4. Ver todas las fórmulas en vez de resultados

### Evaluar fórmula paso a paso

1. Seleccionar celda con fórmula
2. Fórmulas → Evaluar fórmula
3. Ver cálculo paso a paso

### Buscar errores

- `#NAME?` → Función mal escrita o separadores incorrectos
- `#VALUE!` → Tipo de dato incorrecto
- `#REF!` → Referencia a celda eliminada
- `#DIV/0!` → División por cero
- `#N/A` → Valor no disponible (común en FILTER vacío)

---

**Última actualización:** 18 de noviembre, 2025
**Próxima revisión:** Cuando se agreguen nuevas fórmulas
**Mantenedor:** Actualizar este documento ANTES de modificar fórmulas
