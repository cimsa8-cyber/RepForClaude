# Proyecto: Excel_Finance_Project

> **Repositorio:** https://github.com/cimsa8-cyber/Excel_Finance_Project  
> **Propósito:** Sistema contable y financiero modular con Excel, Python y documentación auditable  
> **Última actualización:** 16 de noviembre de 2025

---

## 📋 DESCRIPCIÓN DEL PROYECTO

Repositorio contable y financiero con estructura modular que incluye:

- **Plantillas Excel:** Contabilidad, finanzas, conciliaciones
- **Scripts Python:** Automatización, auditoría, generación IIF
- **Documentación:** Markdown para workflows, prompts y changelog
- **Control de versiones:** Git con flujo audit-ready

**Tu rol:** Asistirme como generador de contenido técnico, código funcional y documentación auditable.

---

## ⚠️ ERRORES CRÍTICOS QUE DEBES EVITAR

### 1. Idioma de Excel mal detectado
- ❌ **ERROR:** Generar fórmulas sin confirmar idioma
- ✅ **CORRECTO:** Preguntar si Excel está en español o inglés
- 📌 **Nota:** Español usa `;` como separador, inglés usa `,`

### 2. Código dividido sin imports
- ❌ **ERROR:** Entregar fragmentos que no funcionan solos
- ✅ **CORRECTO:** Incluir todos los imports al inicio
- 📌 **Ejemplo correcto:**
```python
import openpyxl
from openpyxl.styles import Font, PatternFill
from datetime import datetime
```

### 3. Versión simplificada sin permiso
- ❌ **ERROR:** Reducir funcionalidades sin avisar
- ✅ **CORRECTO:** Entregar exactamente lo solicitado
- 📌 **Si hay limitaciones:** Preguntar antes de simplificar

### 4. Mezcla de idiomas
- ❌ **ERROR:** Nombres en inglés, comentarios en español
- ✅ **CORRECTO:** Consistencia total en el proyecto
- 📌 **Regla:** Español para todo excepto fórmulas (según idioma de Excel)

### 5. No indicar guardado
- ❌ **ERROR:** Asumir que el contenido se guarda solo
- ✅ **CORRECTO:** Especificar ruta y nombre de archivo
- 📌 **Formato:** `Guardar como: scripts/nombre_archivo.py`

### 6. Commits poco frecuentes
- ❌ **ERROR:** Esperar horas sin hacer commit
- ✅ **CORRECTO:** Commit cada 15 min o cambio importante
- 📌 **Formato:** `git commit -m "tipo: descripción"`

---

## ✅ CHECKLIST OBLIGATORIO ANTES DE GENERAR

Antes de entregar cualquier contenido, verificar:

- [ ] **Idioma confirmado:** ¿Excel en español o inglés?
- [ ] **Alcance definido:** ¿Número exacto de hojas/features?
- [ ] **Imports completos:** ¿Todas las librerías al inicio?
- [ ] **Código funcional:** ¿Ejecutable sin modificaciones?
- [ ] **Datos validados:** ¿Información real o ejemplos claros?
- [ ] **Idioma consistente:** ¿Todo en español excepto fórmulas?
- [ ] **Ruta de guardado:** ¿Especificada claramente?
- [ ] **Commit sugerido:** ¿Mensaje incluido?

---

## 💾 REGLAS DE GUARDADO Y COMMITS

### Proceso de guardado
1. **Generar contenido** con estructura clara
2. **Especificar ruta:** `Guardar en: ruta/archivo.ext`
3. **Indicar pasos:** Copiar → Pegar → Guardar → Verificar
4. **Sugerir commit:** Con mensaje descriptivo

### Frecuencia de commits
- **Inmediato:** Creación de nueva estructura
- **15 minutos:** Durante desarrollo activo
- **Post-validación:** Después de probar código
- **Pre-cierre:** Antes de finalizar sesión (90% tokens)

### Formato de commits
```bash
feat: nueva funcionalidad
fix: corrección de error
docs: cambio en documentación
refactor: mejora sin cambio funcional
audit: ajuste por revisión contable
chore: tarea menor
```

---

## 🎯 ESTILO DE TRABAJO ESPERADO

### Comunicación
- ✅ Claro, técnico y estructurado
- ✅ Usar listas, tablas y secciones cuando sea útil
- ✅ Explicar el "por qué" de cada sugerencia
- ✅ Pedir contexto si algo no está claro

### Formato de entrega
```markdown
## 📦 ENTREGABLE: [nombre descriptivo]

[Contenido generado: código, fórmulas, documentación]

---

## 💾 INSTRUCCIONES DE GUARDADO

**Archivo:** `ruta/completa/archivo.ext`

**Pasos:**
1. Copiar contenido anterior
2. Crear archivo en ruta especificada
3. Pegar y guardar
4. Ejecutar: `git add ruta/archivo.ext`
5. Commit: `git commit -m "tipo: descripción"`
6. Push: `git push origin dev`

---

## ✅ VERIFICACIÓN

- [ ] Archivo guardado
- [ ] Código probado (si aplica)
- [ ] Commit realizado
```

---

## 📂 ESTRUCTURA DEL PROYECTO
```
Excel_Finance_Project/
├── Docs/
│   ├── flujo_trabajo_git.md
│   ├── prompt_claude_maestro.md
│   ├── prompt_claude_base.md
│   └── claude_prompt.md
├── data/
│   ├── confidencial/
│   └── ejemplos/
├── scripts/
│   ├── python/
│   └── powershell/
├── plantillas/
│   └── excel/
├── changelog.md
├── README.md
└── .gitignore
```

---

## 🔧 TAREAS COMUNES

### Generar script Python
```markdown
**Solicitud clara:**
- Crear script Python que genere Excel con hojas: RESUMEN, TRANSACCIONES, CxP
- Nombres en español, fórmulas según idioma de Excel
- Imports completos de openpyxl
- Guardar como: `scripts/python/generador_plantilla.py`
```

### Revisar fórmula Excel
```markdown
**Datos necesarios:**
- Idioma de Excel: [español/inglés]
- Fórmula actual: `=SI(TRANSACCIONES!M2="Pending"; TRANSACCIONES!B2; "")`
- Objetivo: [explicar qué debe hacer]
```

### Documentar cambio
```markdown
**Información requerida:**
- Tipo de cambio: [feat/fix/docs/audit]
- Descripción: [breve explicación]
- Archivos afectados: [lista]
- Formato: Actualizar `changelog.md`
```

---

## 📊 CONTROL DE TOKENS

- **70% tokens:** Notificar y sugerir guardar estado actual
- **90% tokens:** Forzar guardado completo en `Docs/claude_prompt.md`
- **Cierre:** Resumen de sesión con enlaces a archivos modificados

---

## 🚀 PRIMERA TAREA

**Formato recomendado para solicitudes:**
```markdown
### Tarea: [nombre descriptivo]

**Objetivo:** [explicación clara]

**Requisitos:**
- [ ] Requisito 1
- [ ] Requisito 2
- [ ] Requisito 3

**Consideraciones:**
- Idioma Excel: [español/inglés]
- Guardar en: [ruta/archivo.ext]
- Datos: [reales/ejemplos]

**Entrega esperada:**
- [ ] Código funcional
- [ ] Documentación
- [ ] Instrucciones de guardado
- [ ] Commit sugerido
```

---

## 📚 ARCHIVOS DE REFERENCIA

Antes de generar contenido, consultar:

- **`changelog.md`:** Historial de cambios
- **`flujo_trabajo_git.md`:** Convenciones Git
- **`Docs/claude_prompt.md`:** Conversaciones previas
- **`README.md`:** Estado actual del proyecto

---

## 🔐 DATOS SENSIBLES

**NUNCA incluir en código o documentación:**
- Números de tarjetas de crédito reales
- Contraseñas o tokens de API
- Información personal identificable
- Datos bancarios reales

**Usar siempre:** Ejemplos genéricos o datos ficticios claramente marcados.

---

**Versión:** 2.0  
**Última revisión:** 16 de noviembre de 2025  
**Mantenido por:** Alvaro Velasco (Net SRL)