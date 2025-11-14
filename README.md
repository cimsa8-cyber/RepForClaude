# 💼 Sistema de Finanzas v4.0 - CRM/SAP en Excel

**Versión:** 4.0
**Fecha:** 14 de Noviembre, 2025
**Estado:** ✅ PRODUCCIÓN

---

## 🎯 OBJETIVO

Sistema de gestión financiera en Excel que funciona como CRM/SAP, con:
- Control de ingresos y egresos
- Gestión de Cuentas por Pagar (CxP)
- Gestión de Cuentas por Cobrar (CxC)
- Dashboard de resumen
- Automatización con Python

## 🔧 MEJORAS DE V4.0

Esta versión corrige **TODOS** los errores críticos de v3.0:

### ✅ Problemas Resueltos

1. **Columnas Alineadas:** Sistema de mapeo dinámico de columnas
2. **CxP/CxC Funcionales:** Validación automática de fórmulas
3. **Formato Correcto:** DD/MM/YYYY para fechas, ₡#,##0.00 para moneda
4. **Validación Pre/Post:** Verificación antes y después de cada operación
5. **Scripts Seguros:** No más hardcoded indices, todo mapeado

### 🚀 Nuevas Características

- 🔍 Auditoría automática después de cada operación
- 💾 Respaldos automáticos antes de cambios destructivos
- 📊 Validador de integridad de datos
- 🧪 Modo prueba (archivo separado) antes de producción
- 📈 Reportes de migración con estadísticas

---

## 📁 ESTRUCTURA DEL PROYECTO

```
finanzas_v4/
├── src/
│   ├── config.py              # Definición de estructura Excel
│   ├── validaciones.py        # Funciones de validación
│   ├── operaciones.py         # CRUD seguro
│   ├── generar_v4.py          # Genera Excel limpio v4.0
│   ├── migracion.py           # Migra datos de v3 a v4
│   └── auditoria.py           # Auditoría post-operación
├── docs/
│   ├── DIAGNOSTICO_V3.md      # Análisis de errores v3.0
│   └── GUIA_USO.md            # Guía de usuario
├── backups/                   # Respaldos automáticos
├── tests/                     # Archivos de prueba
├── requirements.txt           # Dependencias Python
└── README.md                  # Este archivo
```

---

## 🛠️ INSTALACIÓN

### Requisitos
- Python 3.8+
- pip

### Instalación de Dependencias

```bash
pip install -r requirements.txt
```

---

## 📖 USO RÁPIDO

### 1. Generar Archivo Excel Limpio v4.0

```bash
cd finanzas_v4/src
python generar_v4.py
```

**Salida:** `AlvaroVelasco_Finanzas_v4.0.xlsx` con estructura correcta

### 2. Migrar Datos de v3.0 a v4.0

```bash
python migracion.py --origen ../AlvaroVelasco_Finanzas_v3.0.xlsx --destino ../AlvaroVelasco_Finanzas_v4.0.xlsx
```

**Funcionalidad:**
- Lee v3.0 (potencialmente corrupta)
- Valida cada fila
- Migra solo datos válidos
- Limpia duplicados
- Valida CxP/CxC funcionan
- Genera reporte de migración

### 3. Insertar Transacción (Seguro)

```python
from operaciones import insertar_transaccion

data = {
    'Fecha': '01/11/2025',
    'Tipo': 'EGRESO',
    'Categoría': 'Operaciones',
    'Subcategoría': 'Nómina',
    'Descripción': 'Pago de salario',
    'Monto': 500000,
    'Estado': 'PAGADO'
}

insertar_transaccion('v4.0.xlsx', data)
```

### 4. Auditar Archivo

```bash
python auditoria.py --archivo ../AlvaroVelasco_Finanzas_v4.0.xlsx
```

**Validaciones:**
- ✅ Estructura de columnas correcta
- ✅ Formatos de fecha/moneda aplicados
- ✅ CxP tiene datos si existen PENDIENTES
- ✅ CxC tiene datos si existen POR COBRAR
- ✅ Fórmulas funcionando
- ✅ No hay duplicados

---

## 📊 ESTRUCTURA DEL EXCEL v4.0

### Hoja: TRANSACCIONES

| Col | Campo              | Formato       | Ejemplo                |
|-----|--------------------|---------------|------------------------|
| A   | Fecha              | DD/MM/YYYY    | 01/11/2025             |
| B   | Tipo               | Texto         | EGRESO                 |
| C   | Categoría          | Texto         | Operaciones            |
| D   | Subcategoría       | Texto         | Nómina                 |
| E   | Descripción        | Texto         | Pago de salario        |
| F   | Entidad            | Texto         | Juan Pérez             |
| G   | Cuenta             | Texto         | BAC Corriente          |
| H   | Método Pago        | Texto         | Transferencia          |
| I   | Monto              | ₡#,##0.00     | ₡500,000.00            |
| J   | Referencia         | Texto         | TRF-001                |
| K   | Notas              | Texto         | Pago quincenal         |
| L   | Estado             | Texto         | PAGADO / PENDIENTE     |
| M   | IVA                | ₡#,##0.00     | ₡65,000.00             |
| N   | Fecha Vencimiento  | DD/MM/YYYY    | 15/11/2025             |
| O   | Número Factura     | Texto         | FAC-2025-001           |

### Hoja: CxP (Cuentas por Pagar)

Muestra automáticamente transacciones con **Estado = PENDIENTE**

### Hoja: CxC (Cuentas por Cobrar)

Muestra automáticamente transacciones con **Estado = POR COBRAR**

### Hoja: RESUMEN

Dashboard con totales, gráficos y métricas clave.

---

## 🔒 VALIDACIONES AUTOMÁTICAS

### Pre-Operación
- ✅ Headers coinciden con estructura definida
- ✅ Datos tienen formato correcto
- ✅ No hay filas duplicadas antes de insertar

### Post-Operación
- ✅ CxP muestra PENDIENTES correctamente
- ✅ CxC muestra POR COBRAR correctamente
- ✅ Formatos aplicados (fecha/moneda)
- ✅ Fórmulas funcionando
- ✅ Integridad de referencias

---

## 🚨 RESPALDOS

Antes de cualquier operación destructiva:

```python
from operaciones import crear_respaldo

crear_respaldo('v4.0.xlsx')
# Crea: backups/v4.0_20251114_083000.xlsx
```

---

## 📚 DOCUMENTACIÓN ADICIONAL

- [DIAGNOSTICO_V3.md](docs/DIAGNOSTICO_V3.md) - Análisis completo de errores v3.0
- [GUIA_USO.md](docs/GUIA_USO.md) - Guía detallada de uso

---

## 🐛 SOLUCIÓN DE PROBLEMAS

### CxP/CxC vacías

```bash
python auditoria.py --archivo v4.0.xlsx --verbose
```

Verifica:
- ¿Columna L (Estado) tiene valores?
- ¿Existen transacciones PENDIENTES/POR COBRAR?
- ¿Fórmulas apuntan a columna correcta?

### Formato de fecha incorrecto

```python
from operaciones import corregir_formatos

corregir_formatos('v4.0.xlsx')
```

### Migración falló

Revisa: `logs/migracion_YYYYMMDD_HHMMSS.log`

---

## ✅ CHECKLIST PRE-PRODUCCIÓN

Antes de usar v4.0:

- [ ] ✅ Ejecutar `generar_v4.py`
- [ ] ✅ Validar estructura con `auditoria.py`
- [ ] ✅ Migrar datos con `migracion.py` (si aplica)
- [ ] ✅ Probar inserción en archivo TEST
- [ ] ✅ Verificar CxP tiene datos
- [ ] ✅ Verificar CxC tiene datos
- [ ] ✅ Confirmar formatos DD/MM/YYYY visibles
- [ ] ✅ Crear respaldo antes de producción

---

## 🤝 CONTRIBUCIONES

Sistema desarrollado aplicando lecciones de:
- v1.0: Primera versión
- v2.0: Mejoras de estructura
- v3.0: ❌ Errores críticos (documentados)
- **v4.0:** ✅ Sistema robusto y validado

---

## 📞 SOPORTE

Para reportar problemas o sugerencias:
- Revisar `docs/DIAGNOSTICO_V3.md` para entender errores previos
- Ejecutar `auditoria.py` para diagnóstico automático

---

**Desarrollado con:** Python 3.x + openpyxl
**Licencia:** MIT
**Versión Estable:** v4.0 (14/11/2025)
