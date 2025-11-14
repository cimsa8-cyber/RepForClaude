# 🔐 ALIAS CUENTAS BNCR ENMASCARADAS - v3.0

**Fecha de implementación:** 14 de Noviembre, 2025
**Estado:** ✅ COMPLETADO EN v3.0
**Acción:** 🔄 PENDIENTE INTEGRAR EN v4.0

---

## 📋 RESUMEN

Sistema existente en v3.0 que permite mapear cuentas bancarias enmascaradas a nombres reales.

**Propósito:** Reconocer automáticamente cuentas bancarias cuando aparecen enmascaradas en CSVs o transacciones sensibles.

---

## ✅ ALIAS IMPLEMENTADOS EN v3.0

### **Cuentas Corrientes Enmascaradas:**

1. **Alias:** `XXXXXXXXXX1066X`
   - **Entidad Real:** `BNCR Cuenta Corriente Colones`
   - **Tipo:** Cuenta
   - **Moneda:** ₡ (CRC)
   - **Ubicación v3.0:** ENTIDADES_ALIAS fila 47

2. **Alias:** `XXXXXXXXXX8618X`
   - **Entidad Real:** `BNCR Cuenta Corriente Dólares`
   - **Tipo:** Cuenta
   - **Moneda:** $ (USD)
   - **Ubicación v3.0:** ENTIDADES_ALIAS fila 48

---

## 📊 SISTEMA DE ALIAS COMPLETO v3.0

**Estadísticas según documentación:**

| Métrica | Valor |
|---------|-------|
| **Total alias en sistema** | 48 |
| **Alias de cuentas bancarias** | 11 |
| **Alias BNCR** | 7 |
| **Alias BNCR enmascarados** | 2 |
| **Tipos de alias** | 13 categorías |

### **Lista de Alias BNCR (11 total):**

**Cuentas de Ahorro:**
1. `BNCR CRC Ahorros (***8618)` → Cuenta ahorros colones
2. `BNCR USD Ahorros (***1066)` → Cuenta ahorros dólares

**Cuentas Corrientes (enmascaradas):**
3. `XXXXXXXXXX1066X` → BNCR Cuenta Corriente Colones
4. `XXXXXXXXXX8618X` → BNCR Cuenta Corriente Dólares

**Cuentas Corrientes (parcialmente visibles):**
5. `BNCR CRC Corriente (***2186)` → Cuenta corriente colones
6. `BNCR USD Corriente (***9589)` → Cuenta corriente dólares
7. `BNCR USD Corriente (***1112)` → Cuenta corriente dólares alternativa
8-11. (Otros alias BNCR)

---

## 🎯 FUNCIONAMIENTO

### **Estructura ENTIDADES_ALIAS:**
```
Columna A (Alias):        XXXXXXXXXX1066X
Columna B (Entidad Real): BNCR Cuenta Corriente Colones
Columna C (Tipo):         Cuenta
Columna D (Notas):        Cuenta corriente BNCR en colones (enmascarada)
```

### **Flujo de reconocimiento:**
1. Sistema encuentra `XXXXXXXXXX1066X` en transacción/CSV
2. Busca en ENTIDADES_ALIAS columna A
3. Encuentra coincidencia
4. Reemplaza con columna B: "BNCR Cuenta Corriente Colones"
5. Aplica moneda: CRC (₡)

---

## 💡 CASOS DE USO DOCUMENTADOS

### **1. Importación de CSV enmascarado**
```
CSV bancario: Cuenta = XXXXXXXXXX1066X
Sistema: Reconoce como "BNCR Cuenta Corriente Colones"
Moneda: Automática CRC
```

### **2. Conciliación bancaria**
```
Extracto BNCR: XXXXXXXXXX8618X - Transferencia $500
Excel: BNCR Cuenta Corriente Dólares - Transferencia $500
Sistema: ✓ CONCILIADO (misma cuenta)
```

### **3. Reportes ejecutivos**
```
Dashboard muestra:
- Efectivo en BNCR CC Colones: ₡500,000
- Efectivo en BNCR CC Dólares: $2,000
(En lugar de XXXXXXXXXX...)
```

---

## 🔍 IDENTIFICACIÓN INTELIGENTE

El sistema puede identificar por:
1. Nombre completo exacto
2. Enmascaramiento completo (XXXXXXXXXX1066X)
3. Últimos 4 dígitos (***1066)
4. Variaciones parciales

**Prioridad de búsqueda:**
1. Nombre completo exacto
2. Alias enmascarado
3. Últimos dígitos
4. Nombre parcial

---

## 🔐 SEGURIDAD

**Razones de enmascaramiento bancario:**
- Protección de datos sensibles
- Cumplimiento GDPR, PCI-DSS
- Prevención de fraude
- Exportaciones seguras

**Seguridad del archivo:**
- ✅ Excel protegido con contraseña
- ✅ Acceso solo personal autorizado
- ✅ Respaldos en ubicación segura
- ✅ No compartir públicamente

---

## 🚀 INTEGRACIÓN EN v4.0

### **PENDIENTE IMPLEMENTAR:**

1. **Nueva hoja:** `ENTIDADES_ALIAS`
   - Columna A: Alias
   - Columna B: Entidad Real
   - Columna C: Tipo
   - Columna D: Notas

2. **Módulo Python:** `alias.py`
   - Función: `resolver_alias(nombre_entrada)`
   - Función: `agregar_alias(alias, entidad_real, tipo, notas)`
   - Función: `listar_alias()`
   - Función: `validar_alias()`

3. **Integración con operaciones:**
   - Modificar `insertar_transaccion()` para resolver alias
   - Modificar `migración.py` para mapear alias v3.0
   - Agregar validación de alias en `validaciones.py`

4. **Importación CSV:**
   - Script para importar extractos bancarios
   - Auto-resolución de cuentas enmascaradas
   - Mapeo automático a entidades reales

---

## 📝 DATOS A MIGRAR DE v3.0

**CRÍTICO: Al migrar v3.0 → v4.0, capturar:**

1. **Hoja ENTIDADES_ALIAS completa (48 registros)**
2. **Estructura de 4 columnas**
3. **11 alias de cuentas bancarias**
4. **13 categorías de tipos**

**Script de migración debe:**
- Leer ENTIDADES_ALIAS de v3.0
- Crear hoja ENTIDADES_ALIAS en v4.0
- Copiar todos los alias
- Validar integridad
- Verificar funcionamiento

---

## ⚠️ IMPACTO EN v4.0

**SIN este sistema:**
- ❌ Reportes con cuentas enmascaradas (ilegibles)
- ❌ Imposible conciliar con extractos bancarios
- ❌ Duplicados (misma cuenta con diferentes formatos)
- ❌ Monedas incorrectas (no se identifica USD vs CRC)

**CON este sistema:**
- ✅ Reportes claros y profesionales
- ✅ Conciliación automática con bancos
- ✅ Consolidación correcta de cuentas
- ✅ Detección automática de moneda

---

## 📌 ACCIÓN REQUERIDA

**PRÓXIMOS PASOS:**

1. ✅ Documentación guardada (este archivo)
2. ⏳ Esperar más archivos del usuario
3. ⏳ Diseñar módulo `alias.py` para v4.0
4. ⏳ Modificar `migracion.py` para incluir ENTIDADES_ALIAS
5. ⏳ Actualizar `generar_v4.py` para crear hoja ENTIDADES_ALIAS
6. ⏳ Integrar resolución de alias en todas las operaciones

---

**Documento preservado:** 14/11/2025
**Origen:** Sistema v3.0 existente
**Estado:** Analizado y pendiente integración en v4.0
**Prioridad:** 🔴 ALTA (funcionalidad crítica)
