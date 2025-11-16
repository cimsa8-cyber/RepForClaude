## 🎯 Propósito del cambio
<!-- Describe qué problema resuelve este PR y por qué es necesario -->



## 📋 Tipo de cambio
- [ ] `feat`: Nueva funcionalidad
- [ ] `fix`: Corrección de error
- [ ] `audit`: Cambio por revisión contable
- [ ] `docs`: Solo documentación
- [ ] `refactor`: Mejora sin cambio funcional
- [ ] `chore`: Tareas menores (limpieza, actualizaciones)

## 💼 Nivel de Impacto Financiero
<!-- Selecciona el nivel de impacto más alto que aplique -->

- [ ] **🔴 ALTO:** Afecta balance, estado de resultados, flujo de efectivo, o reportes regulatorios
- [ ] **🟡 MEDIO:** Afecta auxiliares, reportes internos, o cálculos no críticos
- [ ] **🟢 BAJO:** Solo documentación, scripts de soporte, o mejoras de proceso

## ✅ Checklist de Validación

### Validación Técnica
- [ ] Código/fórmulas probadas manualmente sin errores
- [ ] Datos de prueba ejecutados exitosamente
- [ ] Sin warnings o errores en logs/consola
- [ ] Archivos temporales eliminados (~$*.xlsx, .tmp, etc.)

### Documentación
- [ ] `README.md` actualizado (si aplica cambio de estructura)
- [ ] `changelog.md` actualizado con entrada de este cambio
- [ ] `TRAZABILIDAD.md` actualizado con referencia (si aplica cambio financiero)
- [ ] Comentarios en código/fórmulas para contexto futuro

### Control de Versiones
- [ ] Commits con mensajes claros siguiendo convención `tipo(scope): descripción`
- [ ] Commits firmados con GPG (si es cambio de impacto ALTO)
- [ ] Branch actualizado con últimos cambios de `dev` (sin conflictos)

### Integridad de Datos (si aplica archivos Excel/CSV)
- [ ] Checksum de integridad generado: `python scripts/audit_integrity_check.py`
- [ ] `audit_checksums.json` incluido en el commit
- [ ] Backup manual creado antes de modificaciones

### Colaboración con Claude (si aplica)
- [ ] Prompts documentados en `Docs/claude_prompt.md` o `Docs/sessions/`
- [ ] Decisiones técnicas de IA explicadas en comentarios
- [ ] Validación humana de código/fórmulas generadas por IA

## 👥 Revisores Requeridos

<!--
Política de aprobación:
- Impacto ALTO: Requiere 2 aprobadores (incluyendo auditor/contador si disponible)
- Impacto MEDIO: Requiere 1 aprobador
- Impacto BAJO: Auto-merge permitido después de 24 horas sin objeciones
-->

**Aprobadores asignados:**
- [ ] @reviewer1
- [ ] @reviewer2 (solo si impacto ALTO)

**Auditor/Contador** (mencionar si aplica):
- [ ] @auditor-name

## 📎 Referencias

**Issues relacionados:**
- Closes #
- Related to #

**Documentación externa:**
- Norma contable aplicable: [NIC/NIIF X]
- Requerimiento regulatorio: [enlace]
- Ticket/solicitud interna: [enlace]

**Commits clave:**
- abc123: descripción del commit principal
- def456: otro commit relevante

## 🧪 Plan de Pruebas
<!-- Describe paso a paso cómo validaste este cambio -->

**Casos de prueba ejecutados:**
1.
2.
3.

**Datos de prueba utilizados:**
- [ ] Datos reales (producción)
- [ ] Datos de ejemplo/mock
- [ ] Datos históricos (año: _____)

**Resultados esperados vs obtenidos:**
| Caso | Esperado | Obtenido | Status |
|------|----------|----------|--------|
| 1    |          |          | ✅/❌   |

## 📸 Capturas de Pantalla / Evidencia
<!-- Si aplica, adjunta screenshots de:
- Resultados de fórmulas Excel
- Salida de scripts
- Reportes generados
- Comparativas antes/después
-->



## 🔄 Plan de Rollback
<!-- Solo para cambios de impacto MEDIO o ALTO -->

**Si este cambio causa problemas, el rollback es:**
- [ ] Fácil: revertir commit es suficiente
- [ ] Medio: requiere restaurar backup de archivos
- [ ] Complejo: requiere intervención manual

**Pasos de rollback:**
1.
2.
3.

## 📝 Notas Adicionales para Revisores
<!-- Contexto adicional, decisiones de diseño, trade-offs considerados, etc. -->



---

## ✍️ Checklist del Autor (antes de solicitar review)
- [ ] Leí completamente este template
- [ ] Completé todas las secciones aplicables
- [ ] Auto-revisé el código/cambios antes de solicitar review
- [ ] Este PR está listo para revisión (no es WIP/draft)
- [ ] He considerado el impacto de este cambio en auditoría futura

## 👀 Checklist del Revisor (antes de aprobar)
- [ ] Revisé todos los archivos modificados
- [ ] Validé que el checklist de validación esté completo
- [ ] Probé el cambio localmente (si aplica)
- [ ] El nivel de impacto declarado es correcto
- [ ] La documentación es suficiente para auditoría futura
- [ ] Estoy de acuerdo en aprobar este cambio

---

**Fecha creación PR:** <!-- Auto-generado por GitHub -->
**Última actualización:** <!-- Auto-actualizado por GitHub -->
