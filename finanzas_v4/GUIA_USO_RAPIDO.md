# 🚀 GUÍA DE USO RÁPIDO - Sistema de Finanzas v4.0

**Álvaro Velasco - AlvaroVelascoNet SRL**
**Versión:** 4.0
**Última actualización:** 14 Noviembre 2025

---

## 📋 ÍNDICE RÁPIDO

1. [Operaciones Diarias](#operaciones-diarias)
2. [Agregar Transacciones](#agregar-transacciones)
3. [Consultar Estado Financiero](#consultar-estado-financiero)
4. [Hojas del Excel](#hojas-del-excel)
5. [Solución de Problemas](#solución-de-problemas)

---

## ⚡ INICIO RÁPIDO (5 minutos)

### Archivo Principal
```
📁 AlvaroVelasco_Finanzas_v4.0.xlsx
```

### Hojas Disponibles
- **TRANSACCIONES** (📝 EDITABLE) - Agrega aquí tus movimientos
- **CxP** (🔒 Protegida) - Cuentas por Pagar (se actualiza sola)
- **CxC** (🔒 Protegida) - Cuentas por Cobrar (se actualiza sola)
- **ENTIDADES_ALIAS** (🔒 Protegida) - Catálogo de bancos/tarjetas
- **DASHBOARDS** (🔒 Protegida) - Resúmenes automáticos

### ⚠️ IMPORTANTE
- ✅ Solo puedes editar la hoja **TRANSACCIONES**
- ✅ Las demás hojas se actualizan automáticamente
- ✅ NO elimines filas/columnas
- ✅ Siempre usa el script para agregar transacciones (más seguro)

---

## 📝 OPERACIONES DIARIAS

### 1️⃣ Agregar una Transacción (Método Recomendado)

**Desde la terminal:**
```bash
cd ~/desktop/debt-sanitization-strategy
python agregar_transaccion.py
```

**Plantillas disponibles:**
1. Ingreso por Venta
2. Pago Tarjeta de Crédito
3. Gasto Operativo
4. Cuenta por Pagar
5. Cuenta por Cobrar
0. Transacción personalizada

**Ejemplo de uso:**
```
Selecciona plantilla (0-5): 2
Descripción: Pago mensual Visa Clásica
Monto: 18000
Entidad: BNCR
Cuenta: Tarjeta Visa Clásica BNCR (***3519)
```

### 2️⃣ Agregar Múltiples Transacciones

```bash
python agregar_transaccion.py --multiple
```

Esto te permitirá agregar varias transacciones seguidas sin salir del programa.

---

## 🏦 TIPOS DE TRANSACCIONES COMUNES

### ✅ Ingreso por Venta de Servicio
```
Tipo: INGRESO
Categoría: Ingresos
Subcategoría: Venta de Servicios
Estado: COBRADO (si ya lo cobraste) o POR COBRAR (si es factura pendiente)
```

### 💳 Pago de Tarjeta de Crédito
```
Tipo: EGRESO
Categoría: Operaciones
Subcategoría: Pago Tarjeta
Cuenta: [Nombre de la tarjeta]
Método Pago: Tarjeta
Estado: PAGADO
```

### 💰 Gasto Operativo
```
Tipo: EGRESO
Categoría: Operaciones
Subcategoría: [Tipo de gasto: Servicios, Suministros, etc.]
Estado: PAGADO
```

### 📄 Factura por Pagar (CxP)
```
Tipo: EGRESO
Estado: PENDIENTE
Fecha Vencimiento: [Fecha de vencimiento]
Número Factura: [Número]
```

### 📄 Factura por Cobrar (CxC)
```
Tipo: INGRESO
Estado: POR COBRAR
Fecha Vencimiento: [Fecha esperada de cobro]
Número Factura: [Número]
```

---

## 📊 CONSULTAR ESTADO FINANCIERO

### Ver Resumen Rápido

Abre el Excel y ve a la hoja **DASHBOARDS**:
- Saldo actual en bancos
- Total CxP pendientes
- Total CxC pendientes
- Posición financiera neta

### Ver Cuentas por Pagar

Abre la hoja **CxP**:
- Todas las facturas/pagos pendientes
- Ordenadas por fecha de vencimiento
- Totales automáticos

### Ver Cuentas por Cobrar

Abre la hoja **CxC**:
- Todas las facturas por cobrar
- Ordenadas por fecha de vencimiento
- Totales automáticos

---

## 🎯 ESTADOS DE TRANSACCIONES

| Estado | Cuándo Usarlo | Aparece en |
|--------|---------------|------------|
| **PAGADO** | Gasto ya pagado | Ninguna (completado) |
| **PENDIENTE** | Tengo que pagar esto | CxP |
| **POR COBRAR** | Me tienen que pagar | CxC |
| **COBRADO** | Ya me pagaron | Ninguna (completado) |
| **CANCELADO** | Se canceló la transacción | Ninguna |
| **PARCIAL** | Pago/cobro parcial | CxP o CxC |

---

## 🔧 OPERACIONES ESPECIALES

### Registrar Pago de Tarjeta de Crédito

**Opción 1: Pago Total**
```
Tipo: EGRESO
Descripción: Pago total Visa Clásica
Monto: [Monto pagado]
Categoría: Operaciones
Subcategoría: Pago Tarjeta
Cuenta: Tarjeta Visa Clásica BNCR (***3519)
Método Pago: Tarjeta
Estado: PAGADO
```

**Opción 2: Pago Mínimo**
```
Tipo: EGRESO
Descripción: Pago mínimo Visa Clásica
Monto: [Pago mínimo]
Categoría: Operaciones
Subcategoría: Pago Mínimo Tarjeta
Cuenta: Tarjeta Visa Clásica BNCR (***3519)
Método Pago: Tarjeta
Estado: PAGADO
```

### Registrar Compra con Tarjeta
```
Tipo: EGRESO
Descripción: [Qué compraste]
Monto: [Monto]
Categoría: [Categoría del gasto]
Cuenta: Tarjeta Visa Clásica BNCR (***3519)
Método Pago: Tarjeta
Estado: PAGADO
```

### Registrar Transferencia entre Cuentas

**Salida de cuenta origen:**
```
Tipo: EGRESO
Descripción: Transferencia a CC Dólares
Monto: [Monto]
Categoría: Transferencias
Cuenta: BNCR Ahorros Colones (***8618)
Estado: PAGADO
Referencia: TRANS-001
```

**Entrada a cuenta destino:**
```
Tipo: INGRESO
Descripción: Transferencia desde Ahorros
Monto: [Monto]
Categoría: Transferencias
Cuenta: BNCR CC Dólares (***9589)
Estado: COBRADO
Referencia: TRANS-001
```

---

## 💡 TIPS Y BUENAS PRÁCTICAS

### ✅ DO's (Hacer)

1. **Usa el script** `agregar_transaccion.py` siempre que puedas
2. **Usa referencias únicas** para transferencias (TRANS-001, TRANS-002...)
3. **Agrega notas** cuando sea relevante (ayuda después)
4. **Revisa CxP y CxC** semanalmente
5. **Mantén las categorías consistentes** (usa siempre las mismas)

### ❌ DON'Ts (No Hacer)

1. **NO edites las hojas protegidas** (CxP, CxC, DASHBOARDS)
2. **NO elimines columnas** de TRANSACCIONES
3. **NO cambies el orden** de las columnas
4. **NO uses acentos raros** o emojis (pueden causar errores)
5. **NO borres filas** sin antes crear backup

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### Problema: "No aparece mi transacción en CxP/CxC"

**Solución:**
- Verifica que el **Estado** sea correcto:
  - CxP: PENDIENTE o PARCIAL
  - CxC: POR COBRAR o PARCIAL
- Verifica que el **Tipo** sea correcto:
  - CxP: EGRESO
  - CxC: INGRESO

### Problema: "Los totales no cuadran"

**Solución:**
1. Cierra y vuelve a abrir el Excel (fuerza recálculo)
2. Verifica que no hayas editado las fórmulas
3. Si persiste, contacta soporte

### Problema: "Agregué algo mal, ¿cómo lo borro?"

**Solución:**
1. Abre el Excel manualmente
2. Ve a hoja TRANSACCIONES
3. Busca la fila con el error
4. Elimina TODA la fila (clic derecho → Eliminar fila)
5. Guarda el archivo

**IMPORTANTE:** Siempre hay backups automáticos en `backups/`

---

## 🎓 CATEGORÍAS RECOMENDADAS

### Ingresos
- `Ingresos > Venta de Servicios`
- `Ingresos > Venta de Productos`
- `Ingresos > Intereses Bancarios`
- `Transferencias` (entre cuentas propias)

### Egresos
- `Operaciones > Servicios` (luz, agua, internet)
- `Operaciones > Suministros`
- `Operaciones > Pago Tarjeta`
- `Operaciones > Pago Mínimo Tarjeta`
- `Operaciones > Sueldo`
- `Operaciones > Honorarios`
- `Transferencias` (entre cuentas propias)

---

## 📞 SOPORTE

### Backups Automáticos
Cada vez que usas el script, se crea un backup automático en:
```
backups/AlvaroVelasco_Finanzas_v4.0_backup_[fecha].xlsx
```

Se mantienen los últimos 10 backups.

### Restaurar un Backup
```bash
cd ~/desktop/debt-sanitization-strategy
cp backups/[nombre_backup].xlsx AlvaroVelasco_Finanzas_v4.0.xlsx
```

---

## 🔮 FUNCIONALIDADES FUTURAS (Ya Disponibles, No Activadas)

El sistema ya tiene listos estos módulos para cuando los necesites:

### 📊 Análisis de Optimización de Pagos
```bash
python finanzas_v4/src/analisis_optimizacion_actual.py
```

**Qué hace:**
- Analiza tus 4 tarjetas de crédito
- Calcula estrategias óptimas de pago
- Proyecta cuándo estarás libre de deudas
- Estima ahorro en intereses

**Cuándo usarlo:**
Cuando definas un presupuesto mensual para pagar deudas.

---

## ✅ CHECKLIST DIARIO

- [ ] Registrar todas las transacciones del día
- [ ] Verificar saldo en DASHBOARDS
- [ ] Revisar CxP próximos a vencer
- [ ] Revisar CxC atrasadas

## ✅ CHECKLIST SEMANAL

- [ ] Revisar todas las CxP
- [ ] Revisar todas las CxC
- [ ] Planear pagos de la semana
- [ ] Crear backup manual (opcional)

## ✅ CHECKLIST MENSUAL

- [ ] Conciliar con extractos bancarios
- [ ] Revisar categorización de gastos
- [ ] Analizar tendencias de ingresos/egresos
- [ ] Actualizar proyecciones (si usas el módulo de optimización)

---

## 🚀 ATAJOS RÁPIDOS

```bash
# Agregar una transacción
python agregar_transaccion.py

# Agregar múltiples transacciones
python agregar_transaccion.py --multiple

# Regenerar archivo desde cero (¡CUIDADO! Pierdes datos nuevos)
python finanzas_v4/src/generar_con_saldos.py

# Ver análisis de optimización
python finanzas_v4/src/analisis_optimizacion_actual.py
```

---

**🎯 ¿Listo para empezar? Ejecuta:**
```bash
cd ~/desktop/debt-sanitization-strategy
python agregar_transaccion.py
```

**¡Buena suerte con tu proyecto de sanitización financiera! 💪**
