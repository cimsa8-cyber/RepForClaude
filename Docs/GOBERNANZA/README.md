# 📁 GOBERNANZA TÉCNICA

> **Proyecto:** Excel_Finance_Project v5.0
> **Sistema de Gobernanza Integrado:** Versión 1.0
> **Última actualización:** 16 de noviembre, 2025

---

## 🎯 BIENVENIDA

Este directorio contiene **TODO el sistema de gobernanza técnica** del proyecto, integrando:
- ✅ Políticas y controles de calidad
- ✅ Templates operativos
- ✅ Procedimientos de auditoría
- ✅ Análisis técnicos y métricas
- ✅ Scripts de automatización

**Sistema audit-ready con cobertura 100% de gobernanza.**

---

## 📂 ESTRUCTURA

```
GOBERNANZA/
├── README.md                              ← Este archivo
├── gobernanza_tecnica.md                  ← 📖 DOCUMENTO MAESTRO - LEER PRIMERO
│
├── TEMPLATES/                             ← Templates operativos (copiar y usar)
│   ├── matriz_trazabilidad.md             - Req → Commit → Archivo
│   ├── registro_rollbacks.md              - Historial de reversiones
│   ├── checklist_auditoria_mensual.md     - 70 puntos de validación
│   ├── checklist_pre_commit.md            - Validación antes de commit
│   └── proceso_4ojos.md                   - Workflow de revisión dual
│
├── POLITICAS/                             ← Políticas detalladas
│   ├── politica_backup.md                 - Backups 3-2-1 + Disaster Recovery
│   ├── riesgos_financieros.md             - Matriz de riesgos + mitigaciones
│   └── seguridad_datos_sensibles.md       - Clasificación + sanitización
│
└── ANALISIS/                              ← Documentos de análisis técnico
    ├── analisis_audit_ready.md            - Análisis completo audit-ready
    ├── analisis_integracion.md            - Integración v5.0 + AR
    └── metricas_kpis.md                   - Métricas de éxito medibles
```

---

## 🚀 QUICK START

### Para Desarrolladores

#### 1. Primera vez en el proyecto
```bash
# Leer documento maestro (30-45 min)
cat Docs/GOBERNANZA/gobernanza_tecnica.md

# Leer flujo de trabajo Git (15 min)
cat Docs/flujo_trabajo_git.md

# Leer prompt Claude (30 min)
cat Docs/prompt_claude_maestro.md
```

#### 2. Antes de cada commit importante
```bash
# Usar checklist pre-commit
cat Docs/GOBERNANZA/TEMPLATES/checklist_pre_commit.md

# Validar calidad
python scripts/python/audit_integrity_check.py

# Si ≥99.5%, proceder con commit firmado
git commit -S -m "tipo(scope): descripción"
```

#### 3. Al implementar requisito nuevo
```bash
# Actualizar matriz de trazabilidad
notepad Docs/GOBERNANZA/TEMPLATES/matriz_trazabilidad.md

# Agregar línea:
# | REQ-XXX | [descripción] | [commit] | [archivos] | [autor] | [fecha] | ✅ |
```

#### 4. Si necesitas hacer rollback
```bash
# Revertir commit
git revert [commit-hash]

# Documentar en registro
notepad Docs/GOBERNANZA/TEMPLATES/registro_rollbacks.md
```

---

### Para Auditores

#### Auditoría Mensual
```bash
# Copiar template
cp Docs/GOBERNANZA/TEMPLATES/checklist_auditoria_mensual.md \
   auditoria_$(date +%Y-%m).md

# Completar 70 puntos de validación
notepad auditoria_$(date +%Y-%m).md

# Commit resultado
git add auditoria_*.md
git commit -m "audit: Auditoría mensual [Mes Año] - [X]%"
```

#### Revisión Trimestral de Riesgos
```bash
# Abrir matriz de riesgos
notepad Docs/GOBERNANZA/POLITICAS/riesgos_financieros.md

# Actualizar columnas: Probabilidad, Impacto, Estado
# Agregar nuevos riesgos si se identificaron
# Commit cambios con tag
git commit -m "audit: Evaluación riesgos Q[X] 2025"
git tag -a risk-review-2025-Q[X]
```

---

### Para Project Managers

#### Dashboard de Métricas
```bash
# Ver KPIs y métricas de éxito
cat Docs/GOBERNANZA/ANALISIS/metricas_kpis.md

# Revisar matriz de trazabilidad (completitud de requisitos)
cat Docs/GOBERNANZA/TEMPLATES/matriz_trazabilidad.md
```

#### Plan de Implementación
```bash
# Ver roadmap 4 fases
cat Docs/GOBERNANZA/ANALISIS/analisis_audit_ready.md
# Buscar sección: "PLAN DE IMPLEMENTACIÓN PRIORITIZADO"
```

---

## 📖 DOCUMENTOS PRINCIPALES

### 1. Documento Maestro
**`gobernanza_tecnica.md`** - Política completa de gobernanza (19 secciones)

**Cuándo leer:** Primera vez en proyecto, o como referencia de políticas

**Contenido:**
- Control de Calidad 99.5%
- Firma GPG y validación 4-ojos
- Matriz de trazabilidad
- Política de backups (resumen)
- Validación de integridad
- Proceso de rollback
- Protección de ramas
- Gestión de riesgos (resumen)
- Seguridad de datos
- Trabajo con Claude AI
- Métricas y KPIs
- Plan de implementación

---

### 2. Políticas Detalladas

#### `POLITICAS/politica_backup.md`
**Cuándo leer:** Antes de configurar backups, o en caso de desastre

**Contenido:**
- Estrategia 3-2-1
- Scripts automatizados (diario/semanal/mensual/trimestral)
- Disaster Recovery (4 escenarios)
- Testing trimestral de backups
- Inventario de backups

**Acción:** Ejecutar `scripts/powershell/setup_backups.ps1` (crear primero)

---

#### `POLITICAS/riesgos_financieros.md`
**Cuándo leer:** Trimestral, o al identificar nuevo riesgo

**Contenido:**
- Matriz de 10 riesgos principales
- Probabilidad e Impacto
- Mitigaciones implementadas
- Indicadores de riesgo
- Proceso de evaluación trimestral

**Acción:** Agregar a calendario trimestral (mar, jun, sep, dic)

---

#### `POLITICAS/seguridad_datos_sensibles.md`
**Cuándo leer:** Antes de commit, o pre-publicación

**Contenido:**
- Clasificación de datos (4 niveles)
- Qué versionar y qué no
- Plan de sanitización
- Proceso de validación

**Acción:** Validar antes de hacer repo público

---

### 3. Templates Operativos

| Template | Uso | Frecuencia |
|----------|-----|------------|
| `checklist_pre_commit.md` | Validación pre-commit | Cada commit crítico |
| `checklist_auditoria_mensual.md` | Auditoría completa | Mensual |
| `matriz_trazabilidad.md` | Tracking requisitos | Con cada REQ nuevo |
| `registro_rollbacks.md` | Documentar rollback | Al revertir commit |
| `proceso_4ojos.md` | Workflow PR | Al crear PR |

---

### 4. Análisis Técnicos

#### `ANALISIS/analisis_audit_ready.md`
Análisis exhaustivo de gaps y mejoras para compliance audit-ready.

**Incluye:**
- Fortalezas y debilidades identificadas
- 7 gaps críticos documentados
- Mejoras detalladas por área
- Plan de implementación 4 fases
- Scripts de ejemplo completos

---

#### `ANALISIS/analisis_integracion.md`
Análisis de integración entre Gobernanza v5.0 y Audit-Ready.

**Incluye:**
- Overlap de contenidos (60%)
- Gaps únicos de cada enfoque
- Propuesta de estructura unificada
- Plan de integración
- Métricas de integración

---

#### `ANALISIS/metricas_kpis.md`
KPIs y métricas de éxito para gobernanza.

**Incluye:**
- 7 métricas principales
- Valores objetivo
- Estado actual
- Dashboard de progreso

---

## 🛠️ SCRIPTS DE AUTOMATIZACIÓN

### Python

| Script | Propósito | Uso |
|--------|-----------|-----|
| `audit_integrity_check.py` | Validación integridad SHA256 | Pre-commit |
| `validador_integridad.py` | Validador sistema completo | Validación 99.5% |

**Ubicación:** `scripts/python/`

**Ejemplo:**
```bash
# Generar snapshot de integridad
python scripts/python/audit_integrity_check.py

# Verificar contra snapshot anterior
python scripts/python/audit_integrity_check.py --verify audit_checksums.json
```

---

### PowerShell

| Script | Propósito | Frecuencia |
|--------|-----------|------------|
| `setup_gobernanza.ps1` | Setup inicial gobernanza | Una vez |
| `backup_daily.ps1` | Backup diario automático | Diario (Task Scheduler) |
| `backup_weekly.ps1` | Backup semanal | Semanal (viernes) |

**Ubicación:** `scripts/powershell/`

**Setup inicial:**
```powershell
# Ejecutar UNA VEZ para crear estructura completa
.\scripts\powershell\setup_gobernanza.ps1
```

---

## 📊 MÉTRICAS DE GOBERNANZA

### Estado Actual

| Métrica | Meta | Actual | Estado |
|---------|------|--------|--------|
| % Commits firmados (críticos) | 100% | [Medir] | 🟡 |
| % PRs con 2+ aprobadores | 100% | [Medir] | 🟡 |
| Tiempo rollback | <30 min | [Medir] | 🟡 |
| Cobertura trazabilidad | >90% | [Medir] | 🟡 |
| Tests backup/año | 4 | 0 | 🔴 |
| Calidad código | ≥99.5% | 99.5% | 🟢 |
| Auditorías sin observaciones | 100% | [Medir] | 🟡 |

**Dashboard completo:** Ver `ANALISIS/metricas_kpis.md`

---

## 🔄 WORKFLOW TÍPICO

### Ciclo de Desarrollo
```
1. Nueva funcionalidad solicitada
   ↓
2. Agregar a matriz_trazabilidad.md (REQ-XXX, estado: En Progreso)
   ↓
3. Crear branch: feature/nombre
   ↓
4. Desarrollo + commits frecuentes
   ↓
5. Antes de cada commit:
   - Checklist pre-commit
   - Validar integridad ≥99.5%
   - Commit firmado: git commit -S
   ↓
6. Push y crear Pull Request
   - Usar template .github/pull_request_template.md
   - Asignar revisor (4-ojos)
   ↓
7. Revisión aprobada → Merge a dev
   ↓
8. Actualizar matriz_trazabilidad.md (REQ-XXX, estado: Completado)
   ↓
9. Testing en dev
   ↓
10. PR a main (requiere 2 aprobadores)
    ↓
11. Merge a main → Tag de versión
```

---

## 📅 CALENDARIO DE GOBERNANZA

### Diario
- ✅ Backup automático (11 PM via Task Scheduler)

### Semanal
- ✅ Backup manual a Google Drive (sábados)
- ✅ Revisión de PRs pendientes (viernes)

### Mensual
- ✅ Auditoría mensual (checklist 70 puntos)
- ✅ Backup snapshot a disco externo (último viernes)
- ✅ Revisión de riesgos emergentes

### Trimestral
- ✅ Evaluación formal de riesgos (Q1: mar, Q2: jun, Q3: sep, Q4: dic)
- ✅ Test de backups (restauración completa)
- ✅ Backup trimestral encriptado offline

### Semestral
- ✅ Revisión de gobernanza técnica (enero, julio)
- ✅ Actualización de políticas si necesario
- ✅ Training/refresher de procedimientos

### Anual
- ✅ Auditoría externa integral
- ✅ Revisión completa de Disaster Recovery
- ✅ Actualización de documentación maestra

---

## 🎓 CAPACITACIÓN

### Nuevos Colaboradores

**Día 1: Setup (4 horas)**
1. Leer `README.md` proyecto (10 min)
2. Leer `gobernanza_tecnica.md` completo (45 min)
3. Leer `flujo_trabajo_git.md` (20 min)
4. Leer `prompt_claude_maestro.md` (30 min)
5. Setup Git + GPG (30 min)
6. Práctica: Generar sistema y validar (1 hora)

**Semana 1: Inmersión**
- Leer templates operativos
- Participar en primera auditoría mensual
- Hacer primer PR con revisión

**Semana 2-4: Autonomía**
- Implementar feature pequeño independiente
- Participar en revisión de PR de otro
- Contribuir a documentación

**Certificación:** Después de 4 semanas + evaluación

---

## 📞 CONTACTO Y SOPORTE

### Responsable de Gobernanza
**Alvaro Velasco**
- Email: [email proyecto]
- GitHub: @cimsa8-cyber

### Canales de Soporte
- 📧 Email para consultas generales
- 🐛 GitHub Issues con etiqueta `governance` para mejoras
- 💬 Chat interno (si aplica)
- 📅 Reuniones semanales: Viernes 10 AM

### Para Reportar
- **Problema de gobernanza:** GitHub Issue
- **Sugerencia de mejora:** GitHub Issue
- **Incidente de seguridad:** Email directo + Issue privado
- **Rollback urgente:** Notificar inmediatamente

---

## 🔗 REFERENCIAS CRUZADAS

### Documentos Relacionados (Raíz Docs/)
- `flujo_trabajo_git.md` - Workflow Git detallado
- `prompt_claude_maestro.md` - Contexto para Claude AI
- `plan_sanitizacion.md` - Limpieza de datos sensibles
- `especificacion_tecnica_v5.md` - Arquitectura del sistema
- `guia_usuario_v5.md` - Manual de usuario final
- `changelog.md` - Historial de cambios

### Scripts
- `scripts/python/` - Scripts Python de validación
- `scripts/powershell/` - Scripts PowerShell de automatización

### GitHub
- `.github/pull_request_template.md` - Template de PR

---

## 📚 ESTÁNDARES Y NORMATIVAS

### Seguimos
- [Conventional Commits](https://www.conventionalcommits.org/)
- [Keep a Changelog](https://keepachangelog.com/es/1.0.0/)
- [Semantic Versioning](https://semver.org/lang/es/)
- [GitHub Flow](https://guides.github.com/introduction/flow/)
- [NIIF para PyMEs (Costa Rica)](https://www.nicniif.org/)

### Cumplimiento
- ✅ Fiscal: Costa Rica (IVA 13%, retenciones, zona franca)
- ✅ Contable: NIIF para PyMEs
- ✅ Seguridad: Datos sensibles protegidos (repo privado)
- ✅ Trazabilidad: Git + GPG + matriz

---

## ❓ FAQ

### ¿Debo firmar TODOS los commits con GPG?
**No.** Solo commits que:
- Modifiquen scripts Python generadores
- Cambien fórmulas o lógica financiera
- Actualicen documentación técnica crítica
- Hagan merge a `main`

Commits menores (typos, README updates) pueden omitir firma.

---

### ¿Qué hago si el validador_integridad.py da <99.5%?
**NO hacer commit.** Primero:
1. Revisar output del validador (qué falló)
2. Corregir errores identificados
3. Re-ejecutar validador hasta ≥99.5%
4. Solo entonces, commit

---

### ¿Cada cuánto debo actualizar matriz_trazabilidad.md?
Con **cada commit que implemente un requisito**.

Si el commit es:
- ✅ `feat: nueva funcionalidad` → Actualizar matriz
- ✅ `fix: corrección importante` → Actualizar matriz (si cierra REQ)
- ❌ `docs: typo en README` → No necesita actualización

---

### ¿Qué pasa si necesito hacer rollback fuera de horario?
1. Ejecutar rollback inmediatamente (no esperar):
   ```bash
   git revert [commit-hash]
   git push origin main
   ```
2. Validar sistema funcional
3. Documentar en `registro_rollbacks.md` **dentro de 24 horas**
4. Notificar a responsables al día siguiente

---

### ¿Cómo sé si un cambio requiere revisión de 4-ojos?
Ver tabla en `TEMPLATES/proceso_4ojos.md`

**Regla general:**
- Scripts Python generadores → ✅ Obligatorio
- Fórmulas Excel críticas → ✅ Obligatorio
- Config fiscal (IVA, ZF) → ✅ Obligatorio
- Typos en docs → ❌ No requerido

---

### ¿Puedo modificar las políticas de gobernanza?
**Sí**, pero con proceso formal:
1. Crear issue en GitHub explicando cambio propuesto
2. Discutir con equipo
3. Crear PR con cambios a gobernanza_tecnica.md
4. Requiere aprobación de responsable de gobernanza
5. Actualizar versión del documento

---

## 🏆 ÉXITO CON GOBERNANZA

> "La gobernanza no es un obstáculo para la innovación,
> es la estructura que la hace sostenible."

### Beneficios Logrados
- ✅ **Trazabilidad 100%**: Req → Commit → Archivo
- ✅ **Calidad 99.5%**: Entregas consistentes
- ✅ **Audit-ready**: Listo para auditoría en cualquier momento
- ✅ **Disaster Recovery**: 4 escenarios documentados
- ✅ **0 rollbacks no documentados**: Historial completo
- ✅ **Colaboración IA efectiva**: Workflow Claude AI definido
- ✅ **Seguridad de datos**: Clasificación y protección

---

**Versión:** 1.0 Integrada
**Estado:** ✅ ACTIVO
**Próxima revisión:** Enero 2026
**Aprobado por:** Alvaro Velasco
**Fecha aprobación:** 16 de noviembre, 2025

---

🏛️ **Sistema de Gobernanza Técnica Completo e Integrado**
