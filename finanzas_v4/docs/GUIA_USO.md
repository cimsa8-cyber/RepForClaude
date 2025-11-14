# 📖 GUÍA DE USO - Sistema de Finanzas v4.0

## 🚀 INICIO RÁPIDO

### 1. Instalación

```bash
cd finanzas_v4
pip install -r requirements.txt
```

### 2. Generar archivo Excel nuevo

```bash
cd src
python generar_v4.py
```

Esto crea: `AlvaroVelasco_Finanzas_v4.0.xlsx` en el directorio raíz.

### 3. Verificar que todo funciona

```bash
python auditoria.py --archivo ../AlvaroVelasco_Finanzas_v4.0.xlsx
```

---

## 📝 OPERACIONES COMUNES

### Insertar una transacción

```python
from operaciones import insertar_transaccion

datos = {
    'Fecha': '14/11/2025',
    'Tipo': 'EGRESO',
    'Categoría': 'Operaciones',
    'Descripción': 'Pago de proveedor',
    'Entidad': 'Proveedor XYZ',
    'Cuenta': 'BAC Corriente',
    'Método Pago': 'Transferencia',
    'Monto': 250000,
    'Estado': 'PENDIENTE',
    'Fecha Vencimiento': '30/11/2025',
    'Número Factura': 'FAC-001'
}

insertar_transaccion('../AlvaroVelasco_Finanzas_v4.0.xlsx', datos)
```

### Actualizar una transacción

```python
from operaciones import actualizar_transaccion

# Actualizar fila 5: cambiar estado a PAGADO
actualizar_transaccion(
    '../AlvaroVelasco_Finanzas_v4.0.xlsx',
    fila=5,
    datos={'Estado': 'PAGADO'}
)
```

### Migrar datos desde v3.0

```bash
# Modo TEST (no guarda cambios, solo reporta)
python migracion.py --origen ../v3.0.xlsx --destino ../v4.0.xlsx --modo test

# Modo PRODUCCIÓN (guarda cambios)
python migracion.py --origen ../v3.0.xlsx --destino ../v4.0.xlsx --modo produccion
```

### Auditar archivo

```bash
# Auditoría básica
python auditoria.py --archivo ../v4.0.xlsx

# Auditoría con detalles
python auditoria.py --archivo ../v4.0.xlsx --verbose

# Guardar log de auditoría
python auditoria.py --archivo ../v4.0.xlsx --guardar-log
```

### Corregir formatos

```python
from operaciones import corregir_formatos

corregir_formatos('../AlvaroVelasco_Finanzas_v4.0.xlsx')
```

---

## 🔧 SOLUCIÓN DE PROBLEMAS

### CxP está vacía pero hay transacciones PENDIENTES

1. Verificar que el Estado sea exactamente "PENDIENTE" (mayúsculas)
2. Ejecutar auditoría:
   ```bash
   python auditoria.py --archivo ../v4.0.xlsx --verbose
   ```
3. Revisar que las fórmulas apunten a la columna L (Estado)

### CxC está vacía pero hay transacciones POR COBRAR

1. Verificar que el Estado sea exactamente "POR COBRAR" (mayúsculas)
2. Ejecutar auditoría como arriba

### Fechas se ven en formato incorrecto

```python
from operaciones import corregir_formatos
corregir_formatos('../v4.0.xlsx')
```

### Migración falló

1. Revisar el reporte de migración
2. Verificar que v3.0 tenga datos válidos
3. Ejecutar en modo TEST primero para ver errores

---

## 📚 REFERENCIA DE CAMPOS

### Campos de TRANSACCIONES

| Campo              | Requerido | Tipo     | Ejemplo                |
|--------------------|-----------|----------|------------------------|
| Fecha              | ✅        | Fecha    | 14/11/2025             |
| Tipo               | ✅        | Texto    | EGRESO / INGRESO       |
| Categoría          | ✅        | Texto    | Operaciones            |
| Subcategoría       | ❌        | Texto    | Nómina                 |
| Descripción        | ✅        | Texto    | Pago de salario        |
| Entidad            | ❌        | Texto    | Juan Pérez             |
| Cuenta             | ✅        | Texto    | BAC Corriente          |
| Método Pago        | ❌        | Texto    | Transferencia          |
| Monto              | ✅        | Número   | 500000                 |
| Referencia         | ❌        | Texto    | TRF-001                |
| Notas              | ❌        | Texto    | Pago quincenal         |
| Estado             | ✅        | Texto    | PAGADO / PENDIENTE     |
| IVA                | ❌        | Número   | 65000                  |
| Fecha Vencimiento  | ❌        | Fecha    | 30/11/2025             |
| Número Factura     | ❌        | Texto    | FAC-2025-001           |

### Valores válidos para Estado

- **PAGADO**: Transacción completada
- **PENDIENTE**: Aparece en CxP (Cuentas por Pagar)
- **POR COBRAR**: Aparece en CxC (Cuentas por Cobrar)
- **CANCELADO**: Transacción cancelada
- **PARCIAL**: Pago parcial

---

## 🎯 MEJORES PRÁCTICAS

1. **Siempre hacer backup antes de cambios importantes**
   ```python
   from operaciones import crear_respaldo
   crear_respaldo('v4.0.xlsx', 'antes_cambios_grandes')
   ```

2. **Validar después de modificaciones importantes**
   ```bash
   python auditoria.py --archivo v4.0.xlsx
   ```

3. **Usar modo TEST antes de PRODUCCIÓN**
   ```bash
   python migracion.py --origen v3.0.xlsx --destino v4.0.xlsx --modo test
   ```

4. **Mantener Estado en mayúsculas**
   - ✅ PENDIENTE
   - ❌ pendiente
   - ❌ Pendiente

5. **Usar formato DD/MM/YYYY para fechas**
   - ✅ 14/11/2025
   - ❌ 2025-11-14
   - ❌ 11/14/2025

---

## 📞 SOPORTE

Para problemas o preguntas:
1. Revisar este documento
2. Ejecutar `python auditoria.py --archivo tu_archivo.xlsx --verbose`
3. Revisar logs en `logs/`
4. Consultar `docs/DIAGNOSTICO_V3.md` para errores conocidos

---

**Última actualización:** 14/11/2025
