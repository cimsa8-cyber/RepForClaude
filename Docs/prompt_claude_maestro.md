# 🎯 PROMPT MAESTRO - Sistema Anti-Loop v3.0

**Proyecto:** Excel Finance ERP v5.0
**Usuario:** Alvaro Velasco | Net SRL
**Última actualización:** 18 de noviembre, 2025
**Costo acumulado en errores:** $195 USD + 7 horas

---

## ⚠️ REGLA DE ORO - LEE ESTO PRIMERO

**ANTES DE HACER CUALQUIER CAMBIO:**

```
┌─────────────────────────────────────────────────────┐
│  🛑 STOP - CHECKLIST OBLIGATORIO                   │
│                                                     │
│  [ ] ¿Leí ESTADO_ACTUAL_PROYECTO.md?              │
│  [ ] ¿Hice backup del archivo que voy a cambiar?   │
│  [ ] ¿Comparé ANTES vs DESPUÉS del cambio?         │
│  [ ] ¿Verifiqué que NO pierdo trabajo?             │
│  [ ] ¿Documenté qué hace cada fórmula crítica?     │
│  [ ] ¿Pedí permiso si voy a simplificar algo?      │
│                                                     │
│  SI ALGUNA RESPUESTA ES "NO" → NO PROCEDER         │
└─────────────────────────────────────────────────────┘
```

---

## 🔴 ERRORES CRÍTICOS DOCUMENTADOS (NO REPETIR)

### Error #1: No confirmar idioma de Excel
**Costo:** $80 USD + 3-4 horas
**Lo que pasó:** Generé fórmulas en español, Excel estaba en inglés
**Prevención:** SIEMPRE preguntar idioma ANTES de generar

### Error #2: Código dividido sin imports
**Costo:** $30 USD + 1 hora
**Lo que pasó:** Imports estaban en parte 2, usuario ejecutó parte 1
**Prevención:** Imports COMPLETOS al inicio SIEMPRE

### Error #3: Versión simplificada sin permiso 🔥
**Costo:** $70 USD + 2 horas + pérdida de confianza
**Lo que pasó:** Simplifiqué hojas sin avisar (placeholder en vez de completo)
**Prevención:** NUNCA simplificar sin permiso EXPLÍCITO del usuario
**Regla:** Completo > Parcial. SIEMPRE.

### Error #4: Mezclar idiomas
**Costo:** $15 USD + 30 min
**Lo que pasó:** Nombres en inglés, contenido en español
**Prevención:** TODO en el mismo idioma (excepto fórmulas técnicas)

### Error #10: No verificar idioma Excel
**Repetición del Error #1** - Aprender la lección

### Error #11: Sin imports completos
**Repetición del Error #2** - Aprender la lección

### Error #12: Entregar versión "simplificada" 🔥🔥🔥
**Repetición del Error #3** - **ESTO ES INACEPTABLE**
**NUNCA MÁS SIMPLIFICAR SIN PERMISO**

### Error #13: Mezclar idiomas
**Repetición del Error #4** - Aprender la lección

### ⚡ ERROR NUEVO: Perder trabajo crítico al migrar
**Costo:** ???
**Lo que pasó:** Al migrar de openpyxl a xlwings, dejé hojas 6-21 como placeholders
**Prevención:** BACKUP + COMPARAR líneas ANTES vs DESPUÉS

---

## 📋 CHECKLIST OBLIGATORIO PRE-CAMBIO

### ✅ FASE 1: PREPARACIÓN (5 min)

```markdown
[ ] 1. Leer ESTADO_ACTUAL_PROYECTO.md
      → Entender QUÉ funciona, QUÉ no
      → Ver lista de archivos críticos

[ ] 2. Identificar archivos que voy a modificar
      → Listar rutas completas
      → Verificar que existen

[ ] 3. Crear BACKUP de cada archivo
      → cp archivo.py archivo_BACKUP_$(date +%Y%m%d_%H%M%S).py
      → Commitear backup
      → Confirmar commit exitoso

[ ] 4. Documentar ESTADO ANTES del cambio
      → Cuántas líneas tiene
      → Qué funcionalidades tiene
      → Qué fórmulas críticas contiene
```

### ✅ FASE 2: VALIDACIÓN (3 min)

```markdown
[ ] 5. ¿Voy a simplificar algo?
      SI → Pedir permiso EXPLÍCITO al usuario
      NO → Continuar

[ ] 6. ¿Voy a eliminar/cambiar funcionalidad existente?
      SI → Pedir permiso EXPLÍCITO al usuario
      NO → Continuar

[ ] 7. ¿Tengo TODOS los datos necesarios?
      → Idioma Excel confirmado
      → Número exacto de hojas
      → Fórmulas exactas
      → Datos reales

[ ] 8. ¿Leí las guías de error en Docs/?
      → Guia claude aprendizaje v2.md
      → GUIA_COMPLETA_TRABAJO_LOCAL.md
```

### ✅ FASE 3: EJECUCIÓN (variable)

```markdown
[ ] 9. Implementar cambio COMPLETO
      → NO versiones parciales
      → NO placeholders
      → TODO funcional

[ ] 10. Comparar ANTES vs DESPUÉS
       ANTES: X líneas, Y funcionalidades
       DESPUÉS: ? líneas, ? funcionalidades

       SI DESPUÉS < ANTES → REVISAR

[ ] 11. Verificar que NO se perdió nada
       → Todas las hojas completas
       → Todas las fórmulas funcionan
       → Todos los dropdowns existen
       → Todas las instrucciones presentes
```

### ✅ FASE 4: COMMIT (2 min)

```markdown
[ ] 12. Commitear con mensaje detallado
       → Qué cambió
       → Por qué cambió
       → Qué se probó
       → Próximos pasos

[ ] 13. Push exitoso
       → Verificar que llegó al remote

[ ] 14. Actualizar ESTADO_ACTUAL_PROYECTO.md
       → Nueva versión
       → Cambios realizados
       → Estado actual
```

---

## 📊 ESTADO ACTUAL DEL PROYECTO

**SIEMPRE consultar:** `Docs/ESTADO_ACTUAL_PROYECTO.md`

Este archivo debe contener:

```markdown
# Estado Actual - ERP v5.0

Última actualización: [FECHA HORA]

## Archivos Funcionales

### generar_finanzas_ERP_v50_COMPLETO.py (openpyxl)
- ✅ Líneas: 1601
- ✅ Hojas: 22 (21 + ALIAS)
- ✅ Hojas 1-5: COMPLETAS con fórmulas
- ✅ Hojas 6-21: COMPLETAS con fórmulas
- ❌ CxP/CxC: FILTER con separadores incorrectos

### generar_finanzas_ERP_v50_xlwings.py (xlwings)
- ✅ Líneas: 569
- ✅ Hojas 1-5: COMPLETAS con fórmulas FILTER funcionando
- ❌ Hojas 6-21: PLACEHOLDERS (perdidas)
- ❌ Hoja INSTRUCCIONES: PLACEHOLDER (perdida)

## Decisión Pendiente
[Qué hacer para recuperar hojas 6-21]
```

---

## 🔄 WORKFLOW ANTI-LOOP

```
┌────────────────────────────────────────────┐
│ Usuario pide cambio                        │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ 🛑 LEER ESTADO_ACTUAL_PROYECTO.md         │
│    ¿Qué funciona? ¿Qué no?                │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ 🛑 CREAR BACKUP de archivos a modificar   │
│    cp archivo archivo_BACKUP_fecha         │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ ¿Voy a simplificar o eliminar algo?        │
│    SI → Pedir permiso EXPLÍCITO            │
│    NO → Continuar                          │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ IMPLEMENTAR cambio COMPLETO                │
│    - NO versiones parciales                │
│    - NO placeholders                       │
│    - TODO funcional                        │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ 🛑 COMPARAR líneas ANTES vs DESPUÉS        │
│    SI DESPUÉS < ANTES → INVESTIGAR         │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ COMMIT + PUSH                              │
└────────────────┬───────────────────────────┘
                 │
                 ▼
┌────────────────────────────────────────────┐
│ ACTUALIZAR ESTADO_ACTUAL_PROYECTO.md       │
└────────────────────────────────────────────┘
```

---

## 📐 INVENTARIO DE FÓRMULAS CRÍTICAS

**SIEMPRE consultar:** `Docs/FORMULAS_CRITICAS.md`

### CxP - Fórmula FILTER (Celda A3)
```excel
=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxP")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxP pendientes")
```
**Qué hace:** Filtra entidades de TRANSACCIONES donde Categoría="CxP" Y Estado="Pendiente"
**Columnas referenciadas:** B (Entidad), C (Categoría), M (Estado)
**Rango:** B3:B1000 (997 filas)
**Mensaje si vacío:** "Sin CxP pendientes"

### CxC - Fórmula FILTER (Celda A3)
```excel
=FILTER(TRANSACCIONES!B3:B1000,(TRANSACCIONES!C3:C1000="CxC")*(TRANSACCIONES!M3:M1000="Pendiente"),"Sin CxC pendientes")
```
**Qué hace:** Filtra entidades de TRANSACCIONES donde Categoría="CxC" Y Estado="Pendiente"
**Similar a CxP pero para cuentas por cobrar**

### RESUMEN - Total CxP USD (Celda B5)
```excel
=SUMIF(CxP!D:D,"USD",CxP!H:H)
```
**Qué hace:** Suma columna H de CxP donde columna D="USD"
**Columnas:** D (Moneda), H (Equiv USD)

### IVA_CONTROL - IVA Aplicable (Celda D3)
```excel
=IF(AND(TRANSACCIONES!I3="Sí",NOT(OR(B3=CONFIG!$B$10,B3=CONFIG!$B$11))),C3*0.13,0)
```
**Qué hace:** Calcula IVA 13% SI: IVA="Sí" Y NO es zona franca
**Zona franca:** CONFIG B10 y B11 (VWR International, RS Hughes)

[Continuar documentando TODAS las fórmulas críticas]

---

## 🚨 SEÑALES DE ALERTA - DETENER Y PREGUNTAR

Si ves alguna de estas señales, **DETENTE Y PREGUNTA AL USUARIO:**

```
🚩 Voy a crear "versión simplificada"
🚩 Voy a dejar algo "para después"
🚩 Voy a usar "placeholders"
🚩 Archivo DESPUÉS tiene menos líneas que ANTES
🚩 Estoy eliminando funcionalidad existente
🚩 No estoy seguro de qué hace una fórmula
🚩 Voy a asumir algo sin confirmar
🚩 No hice backup antes del cambio
🚩 No documenté el estado actual
🚩 Voy a mezclar idiomas
```

---

## 📝 POLÍTICA DE BACKUP OBLIGATORIA

### ANTES de modificar CUALQUIER archivo:

```bash
# 1. Crear backup con timestamp
cp archivo.py archivo_BACKUP_$(date +%Y%m%d_%H%M%S).py

# 2. Commitear backup
git add archivo_BACKUP_*.py
git commit -m "backup: Guardar versión antes de [CAMBIO]"
git push

# 3. SOLO DESPUÉS hacer el cambio
```

### Verificar que backup fue exitoso:
```bash
git log -1
# Debe mostrar commit de backup

git status
# Debe estar clean
```

---

## 📚 ARCHIVOS QUE DEBO CONSULTAR

### Antes de CUALQUIER cambio:
1. `Docs/ESTADO_ACTUAL_PROYECTO.md` - Qué funciona ahora
2. `Docs/Guia claude aprendizaje v2.md` - Errores documentados
3. `Docs/FORMULAS_CRITICAS.md` - Qué hace cada fórmula
4. `Docs/ESPECIFICACION_TECNICA_ERP_v50_COMPLETA.md` - Arquitectura

### Si voy a modificar fórmulas:
5. `Docs/FORMULAS_CRITICAS.md` - Entender qué hace
6. Verificar separadores (`;` vs `,`)
7. Verificar funciones (FILTER vs FILTRAR)

### Si voy a migrar código:
8. Comparar líneas: `wc -l archivo_viejo.py archivo_nuevo.py`
9. Comparar funcionalidades: listar qué hace cada uno
10. Validar que NADA se pierde

---

## ✅ CHECKLIST ENTREGA FINAL

Antes de decir "terminado", verificar:

```markdown
[ ] TODAS las hojas están completas (no placeholders)
[ ] TODAS las fórmulas tienen comentarios explicando qué hacen
[ ] TODOS los dropdowns funcionan
[ ] TODAS las instrucciones están presentes
[ ] Backup creado y commiteado
[ ] Comparación ANTES vs DESPUÉS documentada
[ ] ESTADO_ACTUAL_PROYECTO.md actualizado
[ ] Usuario confirmó que TODO funciona
```

---

## 🎯 REGLAS DE ORO (NUNCA OLVIDAR)

1. **BACKUP ANTES DE CAMBIAR** - Sin excepción
2. **COMPLETO > PARCIAL** - Nunca simplificar sin permiso
3. **COMPARAR ANTES vs DESPUÉS** - Validar que no se pierde nada
4. **DOCUMENTAR ESTADO ACTUAL** - Siempre saber dónde estamos
5. **PREGUNTAR SI HAY DUDA** - Mejor preguntar que asumir
6. **LEER GUÍAS PRIMERO** - Errores ya están documentados
7. **NO REPETIR ERRORES** - Aprender la lección

---

## 📞 CUANDO PEDIR AYUDA AL USUARIO

- ✋ Si no entiendo qué hace una fórmula
- ✋ Si voy a simplificar algo
- ✋ Si voy a eliminar funcionalidad
- ✋ Si archivo DESPUÉS < archivo ANTES
- ✋ Si no sé cómo probar que funciona
- ✋ Si estoy asumiendo algo no confirmado
- ✋ Si detecto que voy a repetir un error documentado

---

**Última revisión:** 18 de noviembre, 2025
**Versión:** 3.0 - Sistema Anti-Loop con validaciones obligatorias
**Próxima revisión:** Después de cada error nuevo

---

## 📊 MÉTRICAS DE ÉXITO

- ❌ Errores repetidos: 0 (objetivo)
- ✅ Backups antes de cambios: 100%
- ✅ Cambios completos (no parciales): 100%
- ✅ Documentación actualizada: 100%
- ✅ Usuario satisfecho: SÍ
