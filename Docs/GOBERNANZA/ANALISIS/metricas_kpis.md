# 📊 MÉTRICAS Y KPIs DE GOBERNANZA

> **Proyecto:** Excel_Finance_Project v5.0
> **Propósito:** Medir efectividad del sistema de gobernanza
> **Última actualización:** 16 de noviembre, 2025

---

## 🎯 DASHBOARD DE MÉTRICAS

### Estado General

| Área | Métrica | Meta | Actual | Estado | Tendencia |
|------|---------|------|--------|--------|-----------|
| **Calidad** | % Calidad código | ≥99.5% | 99.5% | 🟢 OK | ➡️ Estable |
| **Seguridad** | % Commits firmados GPG (críticos) | 100% | [Medir] | 🟡 Pendiente | - |
| **Revisión** | % PRs con 2+ aprobadores (alto impacto) | 100% | [Medir] | 🟡 Pendiente | - |
| **Operaciones** | Tiempo promedio rollback | <30 min | [Medir] | 🟡 Pendiente | - |
| **Trazabilidad** | % Requisitos documentados en matriz | >90% | [Medir] | 🟡 Pendiente | - |
| **Backups** | Tests de backup exitosos/año | 4 | 0 | 🔴 Urgente | - |
| **Auditoría** | % Auditorías sin observaciones críticas | 100% | [Medir] | 🟡 Pendiente | - |

**Leyenda:**
- 🟢 OK: Cumple meta
- 🟡 Pendiente: Aún no medido o en progreso
- 🔴 Urgente: No cumple meta, acción requerida
- ➡️ Tendencia: Mejorando ↗️ | Estable ➡️ | Empeorando ↘️

---

## 📈 MÉTRICAS DETALLADAS

### 1. Calidad de Código

#### 1.1 Calidad General
**Métrica:** % de calidad en validador_integridad.py
**Meta:** ≥99.5%
**Actual:** 99.5% ✅
**Frecuencia medición:** Pre-commit (cada commit crítico)

**Cómo medir:**
```bash
python scripts/python/validador_integridad.py
# Output: "Calidad general: XX.X%"
```

**Acciones si <99.5%:**
1. Revisar errores reportados
2. Corregir antes de commit
3. Re-ejecutar validador
4. Solo commit cuando ≥99.5%

---

#### 1.2 Errores en Fórmulas Excel
**Métrica:** Cantidad de errores #NAME?, #REF!, #VALUE!
**Meta:** 0 errores
**Actual:** 0 ✅
**Frecuencia:** Pre-commit + Auditoría mensual

**Cómo medir:**
1. Abrir Excel generado
2. Buscar errores visualmente o con:
   ```excel
   =COUNTIF(Range, "#NAME?") + COUNTIF(Range, "#REF!") + COUNTIF(Range, "#VALUE!")
   ```

**Acciones si >0:**
1. Identificar celda con error
2. Corregir fórmula en script generador
3. Regenerar Excel
4. Validar 0 errores

---

### 2. Seguridad

#### 2.1 Commits Firmados
**Métrica:** % de commits críticos firmados con GPG
**Meta:** 100%
**Actual:** [Por medir] 🟡
**Frecuencia:** Semanal

**Commits críticos:**
- Cambios en scripts Python generadores
- Modificaciones a fórmulas financieras
- Merge a `main`

**Cómo medir:**
```bash
# Commits totales en main en último mes
git log --oneline --since="1 month ago" main | wc -l

# Commits firmados en main en último mes
git log --show-signature --since="1 month ago" main | grep "Good signature" | wc -l

# Calcular %
```

**Meta por mes:** 100% de commits críticos firmados

**Acciones si <100%:**
1. Identificar commits sin firma
2. Configurar `commit.gpgsign true` globalmente
3. Re-educar en checklist pre-commit

---

#### 2.2 Datos Sensibles Expuestos
**Métrica:** Cantidad de commits con datos sensibles detectados
**Meta:** 0
**Actual:** 0 ✅
**Frecuencia:** Continua (pre-commit)

**Cómo medir:**
```bash
# Buscar patrones sensibles en commits recientes
git log -p --since="1 month ago" | grep -E "(saldo|tarjeta.*[0-9]{4}|password|secret)"
```

**Acciones si >0:**
1. Revertir commit inmediatamente
2. Limpiar historial con BFG o filter-branch
3. Force push (requiere aprobación)
4. Actualizar .gitignore

---

### 3. Revisión y Control

#### 3.1 Pull Requests con Revisión Adecuada
**Métrica:** % de PRs críticos con 2+ aprobadores
**Meta:** 100%
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**PRs críticos:**
- Merge a `main`
- Cambios en lógica financiera core
- Config fiscal (IVA, zona franca)

**Cómo medir:**
```bash
# Listar PRs merged en último mes
gh pr list --state merged --limit 50

# Revisar cada uno si tiene 2+ aprobadores en GitHub UI
```

**Acciones si <100%:**
1. Configurar branch protection en GitHub
2. Requerir 2 approvals para `main`
3. Documentar en proceso_4ojos.md

---

#### 3.2 Tiempo de Revisión de PR
**Métrica:** Tiempo promedio desde PR creado hasta merged
**Meta:** <24 horas
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**Cómo medir:**
Manualmente en GitHub o con API:
```bash
gh pr list --state merged --json createdAt,mergedAt
# Calcular diferencia
```

**Acciones si >24h:**
1. Asignar revisor proactivamente
2. Notificaciones automáticas
3. Revisar si bloqueadores técnicos

---

### 4. Operaciones

#### 4.1 Tiempo de Rollback
**Métrica:** Tiempo desde detección de error hasta sistema restaurado
**Meta:** <30 minutos
**Actual:** [Por medir] 🟡
**Frecuencia:** Post-rollback (ad-hoc)

**Cómo medir:**
Documentar en `registro_rollbacks.md`:
- Timestamp detección error
- Timestamp sistema restaurado
- Calcular diferencia

**Acciones si >30 min:**
1. Revisar por qué demoró
2. Optimizar proceso de rollback
3. Tener backups más frecuentes
4. Practicar disaster recovery

---

#### 4.2 Frecuencia de Rollbacks
**Métrica:** Cantidad de rollbacks en `main` por trimestre
**Meta:** 0
**Actual:** 0 ✅
**Frecuencia:** Trimestral

**Cómo medir:**
```bash
# Contar commits de tipo revert en main en Q actual
git log --oneline main --since="3 months ago" | grep -i "revert"
```

**Acciones si >0:**
1. Analizar causa raíz de cada rollback
2. Implementar prevenciones
3. Mejorar validaciones pre-merge

---

### 5. Trazabilidad

#### 5.1 Cobertura de Matriz de Trazabilidad
**Métrica:** % de requisitos documentados en matriz
**Meta:** >90%
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**Cómo medir:**
1. Contar requisitos identificados (de backlog, issues, etc.)
2. Contar requisitos en `matriz_trazabilidad.md`
3. Calcular %

**Ejemplo:**
- Requisitos totales identificados: 20
- Requisitos en matriz: 18
- Cobertura: 18/20 = 90% ✅

**Acciones si <90%:**
1. Identificar requisitos sin documentar
2. Agregar a matriz con estado correspondiente
3. Entrenar a equipo en uso de matriz

---

#### 5.2 Requisitos Sin Commit Asociado
**Métrica:** Cantidad de REQs completados sin commit hash
**Meta:** 0
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**Cómo medir:**
Revisar `matriz_trazabilidad.md`:
```bash
# Buscar filas con estado "Completado" pero sin commit hash
grep "Completado" Docs/GOBERNANZA/TEMPLATES/matriz_trazabilidad.md | grep -v "[a-f0-9]\{7\}"
```

**Acciones si >0:**
1. Identificar commit que implementó req
2. Actualizar matriz con hash

---

### 6. Backups

#### 6.1 Ejecución de Backups Automáticos
**Métrica:** % de backups programados ejecutados exitosamente
**Meta:** 100%
**Actual:** [Por medir] 🟡
**Frecuencia:** Semanal

**Backups programados:**
- Diario (OneDrive): 7/semana
- Semanal (Google Drive): 1/semana
- Mensual (Disco externo): 1/mes

**Cómo medir:**
```powershell
# Verificar logs de Task Scheduler
Get-ScheduledTask -TaskName "Excel Finance Backup*" | Get-ScheduledTaskInfo

# Verificar timestamps de backups recientes
ls "$env:OneDrive\Backups\FinanzasContabilidad\Daily" | Sort-Object -Descending | Select-Object -First 7
```

**Acciones si <100%:**
1. Revisar logs de Task Scheduler
2. Corregir scripts con errores
3. Validar permisos y rutas

---

#### 6.2 Tests de Backup Exitosos
**Métrica:** Cantidad de tests de backup realizados por año
**Meta:** 4 (trimestral)
**Actual:** 0 🔴
**Frecuencia:** Trimestral

**Test de backup:**
1. Restaurar backup completo en entorno de prueba
2. Validar integridad de archivos
3. Verificar que Excel abre sin errores
4. Documentar resultado

**Cómo medir:**
Buscar en `changelog.md` o crear archivo `tests_backup.md`:
```markdown
| Fecha | Backup Probado | Resultado | Tiempo Restauración |
|-------|----------------|-----------|---------------------|
| 2025-11-16 | 2025-11-15 | ✅ Exitoso | 12 min |
```

**Acciones si <4/año:**
1. Agendar tests trimestrales (mar, jun, sep, dic)
2. Ejecutar procedimiento de testing
3. Documentar resultado

---

### 7. Auditoría

#### 7.1 Auditorías Sin Observaciones Críticas
**Métrica:** % de auditorías mensuales sin observaciones críticas
**Meta:** 100%
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**Observación crítica:** Cualquier item con puntuación 0 en checklist mensual

**Cómo medir:**
Revisar archivos `auditoria_YYYY-MM.md`:
```bash
# Contar auditorías con 0 observaciones críticas
grep "❌" auditoria_*.md | wc -l
```

**Acciones si <100%:**
1. Analizar observaciones recurrentes
2. Implementar acciones correctivas
3. Validar efectividad en siguiente auditoría

---

#### 7.2 Calificación Promedio de Auditorías
**Métrica:** Calificación promedio en auditorías mensuales
**Meta:** ≥95%
**Actual:** [Por medir] 🟡
**Frecuencia:** Mensual

**Cómo medir:**
Extraer calificación total de cada `auditoria_YYYY-MM.md`:
```
Mes 1: 68/70 = 97.1%
Mes 2: 67/70 = 95.7%
Mes 3: 69/70 = 98.6%
Promedio: (97.1 + 95.7 + 98.6) / 3 = 97.1% ✅
```

**Acciones si <95%:**
1. Identificar áreas con puntuación baja
2. Plan de mejora específico
3. Re-evaluar en siguiente mes

---

## 📅 CALENDARIO DE MEDICIÓN

### Diaria
- ✅ Calidad código (pre-commit)
- ✅ Ejecución backup automático

### Semanal
- ✅ Commits firmados
- ✅ Backups ejecutados

### Mensual
- ✅ PRs con revisión adecuada
- ✅ Tiempo revisión PR
- ✅ Cobertura matriz trazabilidad
- ✅ Auditoría mensual completa
- ✅ Calificación auditoría

### Trimestral
- ✅ Frecuencia rollbacks
- ✅ Test de backup completo
- ✅ Evaluación riesgos

### Anual
- ✅ Tests de backup (total 4)
- ✅ Auditoría externa integral
- ✅ Revisión de todas las métricas

---

## 📊 REPORTES

### Reporte Mensual de Gobernanza

**Template:** `reporte_gobernanza_YYYY-MM.md`

```markdown
# Reporte de Gobernanza - [Mes Año]

## Métricas Principales

| Métrica | Meta | Actual | Estado | Acción Requerida |
|---------|------|--------|--------|------------------|
| Calidad código | ≥99.5% | XX.X% | [🟢/🟡/🔴] | [Descripción] |
| Commits firmados | 100% | XX% | [🟢/🟡/🔴] | [Descripción] |
| PRs con 2+ approvals | 100% | XX% | [🟢/🟡/🔴] | [Descripción] |
| Backups exitosos | 100% | XX% | [🟢/🟡/🔴] | [Descripción] |
| Calificación auditoría | ≥95% | XX% | [🟢/🟡/🔴] | [Descripción] |

## Logros del Mes
- [Logro 1]
- [Logro 2]

## Desafíos Enfrentados
- [Desafío 1 + solución]
- [Desafío 2 + solución]

## Plan Próximo Mes
- [Acción 1]
- [Acción 2]

---

**Responsable:** Alvaro Velasco
**Fecha:** [YYYY-MM-DD]
```

---

## 🎯 METAS A 3 MESES

### Q1 2026 (Ene-Mar)

| Métrica | Estado Actual | Meta Q1 | Plan |
|---------|---------------|---------|------|
| Commits firmados | [Medir] | 100% | Configurar GPG + educar |
| Tests backup | 0 | 1 | Ejecutar en marzo |
| Auditorías ≥95% | [Medir] | 3/3 ✅ | Mantener estándares |
| Cobertura trazabilidad | [Medir] | >90% | Documentar todos REQs |

---

## 🏆 RECONOCIMIENTOS

### Cuando alcanzar 100% en todas las métricas durante 3 meses consecutivos:

**Certificación: "Sistema de Gobernanza Production-Ready"**

Beneficios:
- ✅ Confianza total en trazabilidad
- ✅ Listo para auditoría externa en cualquier momento
- ✅ Sistema resiliente ante desastres
- ✅ Cultura de calidad establecida

---

## 📞 RESPONSABLE DE MÉTRICAS

**Alvaro Velasco**
- Medición mensual de métricas
- Generación de reportes
- Identificación de acciones correctivas
- Comunicación a stakeholders

---

**Última actualización:** 16 de noviembre, 2025
**Próxima revisión:** Diciembre 2025
**Versión:** 1.0
