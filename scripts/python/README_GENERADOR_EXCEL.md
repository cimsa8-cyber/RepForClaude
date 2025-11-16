# 📊 Generador de Excel Financiero Profesional v5.0

## 📋 Descripción

Script Python que genera automáticamente el archivo `Finanzas_v50.xlsx` con estructura profesional para contabilidad y finanzas personales.

**Características:**
- ✅ Paleta de colores profesional (estándares financieros/contables)
- ✅ Formato de fecha: `dd/mm/yy` (Costa Rica)
- ✅ Fórmulas en ESPAÑOL (`SUMA`, `SI`, separador `;`)
- ✅ Columnas auto-ajustables
- ✅ Protección de columnas calculadas
- ✅ 5 hojas: Dashboard, Ingresos, Gastos, Balance General, Flujo de Caja

---

## 🔧 Requisitos

### Instalación de Python
```powershell
# Verificar si Python está instalado
python --version

# Si no está instalado, descargar desde:
# https://www.python.org/downloads/
```

### Instalación de openpyxl
```powershell
# Instalar librería necesaria
pip install openpyxl
```

---

## 🚀 Uso

### Opción 1: Generación básica (ubicación actual)
```powershell
# Navegar a la carpeta del script
cd "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad\scripts\python"

# Ejecutar script
python generar_finanzas_v50.py

# Se creará: Finanzas_v50.xlsx en la misma carpeta
```

### Opción 2: Especificar ruta de salida
```powershell
# Generar en carpeta data/
python generar_finanzas_v50.py "../../data/Finanzas_v50.xlsx"

# Generar en escritorio
python generar_finanzas_v50.py "$env:USERPROFILE\Desktop\Finanzas_v50.xlsx"
```

---

## 📂 Estructura del Excel Generado

### 1. Dashboard
- Resumen ejecutivo con KPIs
- Indicadores clave (Ingresos, Gastos, Flujo, Ratios)
- Enlaces a hojas detalladas

### 2. Ingresos
- Registro detallado de todos los ingresos
- Categorías: Salario, Freelance, Inversiones, etc.
- Total calculado automáticamente

### 3. Gastos
- Registro detallado de todos los gastos
- Categorías: Vivienda, Servicios, Alimentación, Transporte, Salud, etc.
- Subcategorías para mayor detalle
- Total calculado automáticamente

### 4. Balance General
- Activos (Efectivo, Bancos, Inversiones, Activos fijos)
- Pasivos (Préstamos, Tarjetas, Cuentas por pagar)
- Patrimonio Neto (Activos - Pasivos)
- Porcentajes calculados automáticamente

### 5. Flujo de Caja
- Proyección mensual (12 meses)
- Flujo neto mensual
- Acumulado anual
- Totales anuales

---

## 🎨 Paleta de Colores Profesional

| Elemento | Color | Uso |
|----------|-------|-----|
| Header principal | Azul oscuro (#1F4E78) | Títulos de hojas |
| Subheader | Azul medio (#4472C4) | Encabezados de columnas |
| Ingresos | Verde claro (#E2EFDA) | Sección de ingresos |
| Gastos | Naranja claro (#FCE4D6) | Sección de gastos |
| Activos | Azul claro (#DEEBF7) | Sección de activos |
| Pasivos | Amarillo claro (#FFF2CC) | Sección de pasivos |
| Totales | Azul muy claro (#D9E1F2) | Filas de totales |
| Positivo | Verde claro (#C6EFCE) | Ganancias/superávit |
| Negativo | Rojo claro (#FFC7CE) | Pérdidas/déficit |

---

## 🔒 Columnas Protegidas (Calculadas Automáticamente)

Las siguientes columnas contienen fórmulas y NO deben editarse manualmente:

**Dashboard:**
- Columna B: Valores actuales (referencia a otras hojas)
- Columna C: Porcentajes

**Ingresos:**
- Columna D (fila 100): Total ingresos
- Columna F (fila 100): Total para referencia

**Gastos:**
- Columna E (fila 100): Total gastos
- Columna F (fila 100): Total para referencia

**Balance General:**
- Columna C: Totales calculados
- Columna D: Porcentajes

**Flujo de Caja:**
- Columna D: Flujo neto (Ingresos - Gastos)
- Columna E: Acumulado

---

## ✏️ Cómo Personalizar tus Datos

### 1. Datos de Ejemplo Incluidos
El archivo generado incluye datos de ejemplo para que veas la estructura. Estos datos deben ser reemplazados por tus datos reales.

### 2. Agregar Ingresos
1. Ir a hoja "Ingresos"
2. Agregar filas entre la fila 3 y 99
3. Completar: Fecha, Categoría, Descripción, Monto, Método Pago, Estado, Notas
4. El total se calcula automáticamente en fila 100

### 3. Agregar Gastos
1. Ir a hoja "Gastos"
2. Agregar filas entre la fila 3 y 99
3. Completar: Fecha, Categoría, Subcategoría, Descripción, Monto, Método Pago, Estado, Notas
4. El total se calcula automáticamente en fila 100

### 4. Actualizar Balance General
1. Ir a hoja "Balance General"
2. Modificar montos de Activos y Pasivos según tu situación real
3. Los totales y patrimonio se calculan automáticamente

### 5. Proyectar Flujo de Caja
1. Ir a hoja "Flujo de Caja"
2. Ingresar proyecciones mensuales en columnas B (Ingresos) y C (Gastos)
3. El flujo neto y acumulado se calculan automáticamente

---

## 🔍 Verificación Post-Generación

Después de generar el archivo, verificar:

- [ ] Archivo se abre correctamente en Excel
- [ ] Todas las hojas están presentes (5 hojas)
- [ ] Formato de fechas: dd/mm/yy
- [ ] Fórmulas en español (SUMA, no SUM)
- [ ] Colores aplicados correctamente
- [ ] Columnas con ancho adecuado
- [ ] Sin errores #NAME?, #REF!, #VALUE!

---

## 🐛 Solución de Problemas

### Error: "No module named 'openpyxl'"
**Solución:**
```powershell
pip install openpyxl
```

### Error: "Python no reconocido"
**Solución:**
1. Instalar Python desde https://www.python.org/downloads/
2. Durante instalación, marcar "Add Python to PATH"
3. Reiniciar PowerShell

### Fórmulas no funcionan (muestran #NAME?)
**Causa:** Excel instalado en idioma diferente
**Solución:** Verificar que tu Excel esté en español. Si está en inglés, contactar para versión bilingüe del script.

### Archivo no se genera
**Solución:**
1. Verificar permisos de escritura en la carpeta
2. Cerrar cualquier archivo Finanzas_v50.xlsx que esté abierto
3. Ejecutar PowerShell como Administrador

---

## 📝 Notas Importantes

1. **Backup:** Siempre hacer backup antes de modificar datos importantes
2. **Formato de moneda:** ₡ (Colón costarricense) - Cambiar si usas otra moneda
3. **Datos sensibles:** No hacer commit del .xlsx a Git (está en .gitignore)
4. **Actualizaciones:** Regenerar archivo solo si necesitas estructura nueva
5. **Datos reales:** Guardar copia antes de regenerar para no perder datos

---

## 🔄 Integración con Gobernanza

Este generador se integra con el sistema de gobernanza:

1. **Validación de integridad:**
   ```powershell
   python audit_integrity_check.py --data-dir "../../data"
   ```

2. **Commit tras generación:**
   ```powershell
   git add Finanzas_v50.xlsx
   git commit -m "feat: Generar Excel financiero v5.0 profesional"
   ```

3. **Snapshot de integridad:**
   ```powershell
   python audit_integrity_check.py
   git add audit_checksums.json
   git commit -m "audit: Snapshot post-generación Excel v5.0"
   ```

---

## 📞 Soporte

Si encuentras problemas:
1. Revisar sección "Solución de Problemas" arriba
2. Verificar logs de error en PowerShell
3. Consultar documentación en `Docs/GOBERNANZA/README.md`

---

**Autor:** Alvaro Velasco
**Proyecto:** Excel_Finance_Project v5.0
**Fecha:** 16 de noviembre, 2025
**Versión script:** 1.0
