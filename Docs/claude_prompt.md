# 🔐 PLAN DE SANITIZACIÓN - Eliminar Datos Sensibles

**Propósito:** Limpiar datos confidenciales antes de hacer repo público  
**Fecha creación:** 16 de noviembre, 2025  
**Estado:** Plantilla para ejecución futura

---

## 🎯 OBJETIVO

Eliminar **TODOS** los datos sensibles del repositorio para poder:
- Hacer el repo público en GitHub
- Compartir código como portfolio
- Usar como template para otros proyectos
- Publicar en tu sitio web profesional

---

## 📊 INVENTARIO DE DATOS SENSIBLES

### Ubicaciones Confirmadas

| Archivo | Datos Sensibles | Acción |
|---------|-----------------|--------|
| `Docs/claude_prompt.md` | Saldos tarjetas, proveedores | Sanitizar |
| `Docs/especificacion_tecnica_v5.md` | Saldos, CxP, CxC | Sanitizar |
| `scripts/python/generar_v5_FINAL.py` | Saldos en código | Sanitizar |
| `data/confidencial/*.xlsx` | TODO | Gitignored ya |
| `changelog.md` | Posibles referencias | Revisar |
| `README.md` | Empresa, contacto | Sanitizar |
| Commits Git | Mensajes con datos | Reescribir historia |

### Tipos de Datos a Eliminar

✅ **Financieros:**
- Saldos de tarjetas (₡1,831,000, etc)
- Montos CxP ($6,500)
- Montos CxC ($9,923)
- Tipo de cambio específico si cambia

✅ **Identificadores:**
- Nombre completo: Alvaro Velasco
- Empresa: Net SRL
- Email corporativo
- Números de cuenta/tarjeta si aparecen

✅ **Proveedores:**
- Nombres clientes reales
- Proveedores zona franca (opcional mantener genérico)

---

## 🛠️ PROCESO DE SANITIZACIÓN

### Fase 1: Preparación (antes de empezar)
````bash
# 1. Crear branch sanitización
git checkout -b sanitize-for-public

# 2. Backup completo del repo privado
cd ..
tar -czf Excel_Finance_Project_PRIVATE_BACKUP_$(date +%Y%m%d).tar.gz Proyecto\ FinanzasContabilidad/
# Guardar en disco externo seguro

# 3. Clonar repo en nueva ubicación
cd ..
git clone Proyecto\ FinanzasContabilidad/ Excel_Finance_Project_PUBLIC
cd Excel_Finance_Project_PUBLIC
````

### Fase 2: Sanitización de Archivos

#### A. Docs/claude_prompt.md
````bash
# Abrir y buscar/reemplazar
notepad Docs\claude_prompt.md

# Buscar y reemplazar:
₡1,831,000 → ₡XXX,XXX (ejemplo)
₡1,750,000 → ₡XXX,XXX (ejemplo)
$2,500 → $XXX (ejemplo)
₡2,800,000 → ₡XXX,XXX (ejemplo)
$6,500 → $X,XXX (ejemplo)
$9,923 → $X,XXX (ejemplo)
Alvaro Velasco → [Tu Nombre]
Net SRL → [Tu Empresa]
[email] → [contacto]
VWR International LLC → Proveedor ZF 1
RS Hughes Co. Inc. → Proveedor ZF 2
````

#### B. Docs/especificacion_tecnica_v5.md
````bash
notepad Docs\especificacion_tecnica_v5.md

# Mismo proceso de reemplazo
# Usar valores genéricos de ejemplo
````

#### C. scripts/python/generar_v5_FINAL.py
````python
# Cambiar TARJETAS de:
TARJETAS = [
    {"nombre": "BAC Visa", "saldo": 1831000, "moneda": "CRC", "dia": 15},
    # ...
]

# A:
TARJETAS = [
    {"nombre": "Banco A - Visa", "saldo": 150000, "moneda": "CRC", "dia": 15},
    {"nombre": "Banco B - MC", "saldo": 120000, "moneda": "CRC", "dia": 20},
    {"nombre": "Banco C - Platinum", "saldo": 2000, "moneda": "USD", "dia": 10},
    {"nombre": "Banco D - Gold", "saldo": 180000, "moneda": "CRC", "dia": 10}
]

TC_USD_CRC = 500  # Valor genérico redondeado

ZONA_FRANCA = ["Proveedor ZF 1", "Proveedor ZF 2"]
````

#### D. README.md
````markdown
# Cambiar de:
**Autor:** Alvaro Velasco
**Empresa:** Net SRL

# A:
**Autor:** [Tu Nombre] - Consultor Financiero
**Uso:** Template para sistemas financieros empresariales
````

#### E. changelog.md
````bash
# Revisar commit por commit
# Eliminar cualquier referencia a montos reales
# Mantener solo cambios técnicos
````

### Fase 3: Sanitización de Historial Git

**⚠️ CRÍTICO:** Git guarda TODO el historial, incluyendo datos borrados.
````bash
# Opción A: Squash todo el historial (recomendado)
git checkout --orphan clean-history
git add .
git commit -m "INIT: Sistema Financiero CRM v5.0

Sistema de gestión financiera empresarial comparable a SAP/QuickBooks

Features:
- 14 módulos integrados (TRANSACCIONES, CxP, CxC, etc)
- Multi-moneda USD/CRC
- Validaciones exhaustivas
- Conciliación contable automática
- Reportería ejecutiva en tiempo real
- Scripts Python para automatización
- Calidad 99.5% production-ready

Tech Stack:
- Python 3.11+
- openpyxl
- Excel con fórmulas avanzadas
- Git para versionado

NOTA: Todos los datos son ejemplos ficticios
"

git branch -D main
git branch -m main
git push -f origin main

# Opción B: BFG Repo-Cleaner (más complejo pero preserva más historia)
# https://rtyley.github.io/bfg-repo-cleaner/
````

### Fase 4: Limpieza de Archivos Temporales
````bash
# Buscar archivos Excel olvidados
find . -name "*.xlsx" -type f
# Eliminar cualquiera encontrado

# Buscar backups
find . -name "*_BACKUP*" -type f
find . -name "*.bak" -type f

# Limpiar caché Python
find . -name "__pycache__" -type d -exec rm -rf {} +
find . -name "*.pyc" -delete
````

### Fase 5: Validación Final
````bash
# Script de validación automatizada
cat > scripts/validate_sanitization.sh << 'EOF'
#!/bin/bash

echo "🔍 VALIDANDO SANITIZACIÓN..."

# Buscar patrones sensibles
echo ""
echo "Buscando montos en colones..."
grep -r "₡[0-9,]" --include="*.md" --include="*.py" .

echo ""
echo "Buscando montos en dólares altos..."
grep -r "\$[0-9][0-9][0-9][0-9]" --include="*.md" --include="*.py" .

echo ""
echo "Buscando nombre completo..."
grep -r "Alvaro Velasco" --include="*.md" --include="*.py" .

echo ""
echo "Buscando empresa..."
grep -r "Net SRL" --include="*.md" --include="*.py" .

echo ""
echo "Buscando emails..."
grep -r "@" --include="*.md" --include="*.py" . | grep -v "example.com"

echo ""
echo "✅ Validación completada"
EOF

chmod +x scripts/validate_sanitization.sh
./scripts/validate_sanitization.sh
````

### Fase 6: Crear README Público
````markdown
# 📊 Sistema Financiero Empresarial - Template

Sistema de gestión financiera comparable a SAP Business One / QuickBooks Enterprise, 
implementado en Excel con automatización Python.

## 🎯 Características

- ✅ 14 módulos integrados
- ✅ Multi-moneda (USD/cualquier moneda local)
- ✅ Validaciones exhaustivas en todos los inputs
- ✅ Conciliación contable automática
- ✅ Reportería ejecutiva en tiempo real
- ✅ Calidad 99.5% production-ready
- ✅ Scripts Python para generación automatizada

## 🚀 Quick Start
```bash
# 1. Clonar repositorio
git clone https://github.com/[tu-usuario]/Excel_Finance_Project

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Generar sistema
python scripts/python/generar_v5_FINAL.py

# 4. Abrir Excel generado
start data/confidencial/Sistema_Financiero_v5.0.xlsx
```

## 📖 Documentación

- [Especificación Técnica](Docs/especificacion_tecnica_v5.md)
- [Guía de Usuario](Docs/guia_usuario_v5.md)
- [Workflow Git](Docs/flujo_trabajo_git.md)

## 🛡️ Datos de Ejemplo

**IMPORTANTE:** Todos los datos en este repositorio son ficticios y de ejemplo.
Para uso en producción, configurar datos reales en `data/confidencial/datos_reales.json`

## 📄 Licencia

MIT License - Uso libre para proyectos comerciales y personales

## 👤 Autor

**[Tu Nombre]**
- Portfolio: [tu-sitio-web]
- LinkedIn: [tu-linkedin]
- Email: [tu-email]

## 🙏 Créditos

Sistema desarrollado con asistencia de Claude AI (Anthropic)
Inspirado en mejores prácticas de SAP, QuickBooks, NetSuite
````

---

## 📋 CHECKLIST FINAL SANITIZACIÓN
````markdown
### Archivos
- [ ] Docs/claude_prompt.md sanitizado
- [ ] Docs/especificacion_tecnica_v5.md sanitizado
- [ ] scripts/python/generar_v5_FINAL.py sanitizado
- [ ] README.md reescrito para público
- [ ] changelog.md revisado
- [ ] .gitignore actualizado
- [ ] Todos los .xlsx eliminados de repo

### Git History
- [ ] Historial limpiado (squash o BFG)
- [ ] Commits revisados
- [ ] Tags revisados
- [ ] Branches limpiados

### Validación
- [ ] Script validación ejecutado
- [ ] 0 resultados de búsqueda datos sensibles
- [ ] README público creado
- [ ] Licencia agregada
- [ ] Backup privado guardado seguro

### Testing
- [ ] Repo clonado en máquina limpia
- [ ] Scripts ejecutan sin datos reales
- [ ] Documentación tiene sentido sin contexto privado
- [ ] Enlaces externos funcionan

### Publicación
- [ ] Repo marcado como público en GitHub
- [ ] About/Description actualizado
- [ ] Topics/Tags agregados
- [ ] GitHub Pages configurado (opcional)
````

---

## 🚨 IMPORTANTES RECORDATORIOS

### ⚠️ Antes de Hacer Público

1. **Triple verificación** de datos sensibles eliminados
2. **Test en máquina limpia** (clonar y probar)
3. **Backup seguro** del repo privado original
4. **No se puede deshacer** una vez público

### ✅ Después de Hacer Público

1. **Monitorear issues** por si alguien encuentra datos
2. **Responder rápido** si hay problema de seguridad
3. **Mantener actualizado** con mejoras sin datos privados

---

## 📞 CONTACTO PARA SANITIZACIÓN

Si necesitas ayuda cuando llegue el momento:

1. Abrir nueva conversación con Claude
2. Adjuntar este documento
3. Decir: "Necesito ejecutar plan de sanitización"
4. Claude te guiará paso a paso

---

**Guardado en:** `Docs/plan_sanitizacion.md`  
**Ejecutar cuando:** Proyecto terminado y validado  
**Tiempo estimado:** 4-6 horas de trabajo cuidadoso

---

> "La privacidad no es el acto de ocultar algo, es el acto de proteger algo" - Unknown