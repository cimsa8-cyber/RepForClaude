# Flujo de trabajo Git para Proyecto FinanzasContabilidad

> **Propósito:** Flujo Git audit-ready para proyectos financieros con Claude  
> **Última revisión:** 16/nov/2025

Este documento describe paso a paso cómo se estructura, inicializa y mantiene el repositorio.

## 1. Estructura del repositorio
```
Proyecto FinanzasContabilidad/
├── Docs/                  # Documentación técnica y prompts para Claude
├── data/                  # Archivos fuente (Excel, CSV, etc.)
├── scripts/               # Scripts auxiliares (PowerShell, Python, etc.)
├── changelog.md           # Registro de cambios con fechas y motivos
├── README.md              # Descripción general del proyecto
├── .gitignore             # Exclusiones para Git
```

## 2. Inicialización y configuración

Los siguientes comandos se ejecutan en PowerShell para inicializar el repositorio local y conectarlo con GitHub:
```powershell
cd "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad"
git init
git branch -M main
git add .
git commit -m "Inicialización del proyecto con estructura base y documentación"
git remote add origin https://github.com/cimsa8-cyber/Excel_Finance_Project
git push -u origin main
git checkout -b dev
git push -u origin dev
```

---

## 3. Ramas y flujo de trabajo

El proyecto utiliza una estrategia de ramas que permite mantener orden, trazabilidad y colaboración audit-ready. Cada rama tiene un propósito específico:

### Ramas principales

- `main`: rama estable. Solo se actualiza desde `dev` mediante revisión y validación.
- `dev`: rama activa de desarrollo. Aquí se integran los cambios antes de pasar a producción.

### Ramas temáticas

- `feature/nombre`: para nuevas funcionalidades. Se crean desde `dev` y se integran mediante pull request.
- `audit/nombre`: para revisiones contables, validaciones o ajustes con trazabilidad.
- `claude/nombre`: para prompts, documentación generada por IA o colaboraciones con Claude.

### Ejemplo de creación de rama
```powershell
git checkout -b feature/plantilla_gastos
```

---

## 4. Convenciones de commit

Los mensajes de commit deben ser claros, consistentes y auditables. Se recomienda seguir el formato:
```
tipo: descripción breve
```

### Tipos comunes

- `feat`: nueva funcionalidad
- `fix`: corrección de errores
- `docs`: cambios en documentación
- `refactor`: reestructuración de código sin cambiar funcionalidad
- `style`: cambios de formato (espacios, indentación, etc.)
- `test`: adición o mejora de pruebas
- `chore`: tareas menores (actualizaciones, limpieza, etc.)

### Ejemplos
```bash
git commit -m "feat: agregado template de gastos mensuales"
git commit -m "fix: corregido error en fórmula de balance"
git commit -m "docs: actualizado flujo de trabajo Git"
```

---

## 5. Integración con Claude

Claude se utiliza como asistente técnico y documental. Su colaboración se organiza en los siguientes archivos:

- `Docs/prompt_claude_base.md`: prompts base para tareas repetitivas.
- `Docs/prompt_claude_maestro.md`: prompts estructurados para tareas complejas.
- `Docs/claude_prompt.md`: respuestas clave generadas por Claude.

Claude puede ayudarte a:

- Validar fórmulas y cálculos financieros.
- Generar tablas comparativas y resúmenes.
- Sugerir mejoras en `changelog.md`.
- Crear documentación técnica y explicaciones.

Toda interacción relevante con Claude debe quedar registrada en `claude_prompt.md` para trazabilidad y revisión futura.

---

## 6. Auditoría y trazabilidad

Para mantener un historial claro y auditable, se recomienda utilizar los siguientes comandos de Git:

### Ver historial de commits
```bash
git log
```

### Ver diferencias entre versiones
```bash
git diff
```

### Ver quién modificó cada línea
```bash
git blame archivo
```

### Ver cambios en archivos específicos
```bash
git log -- archivo
```

Toda modificación relevante debe reflejarse en `changelog.md` con fecha, motivo y referencia al commit correspondiente. Esto garantiza trazabilidad técnica y contable.

---

## 7. Colaboración y pull requests

Este proyecto está diseñado para facilitar la colaboración entre humanos y asistentes IA como Claude. Para mantener orden y trazabilidad, se recomienda seguir estas prácticas al trabajar en equipo:

### Flujo de colaboración

1. Crear una rama temática desde `dev` (ej. `feature/validacion_ingresos`)
2. Realizar cambios y commits con mensajes claros y auditables
3. Subir la rama al repositorio remoto
4. Crear un pull request (PR) hacia `dev` desde GitHub
5. Incluir en el PR:
   - Descripción clara del propósito
   - Referencia a commits relevantes
   - Enlace a documentación o prompts si aplica
6. Solicitar revisión por parte de otro colaborador o Claude
7. Aprobar y fusionar el PR tras validación

### Ejemplo de comandos
```bash
git push -u origin feature/validacion_ingresos
```

---

## 7.5. Sincronización con Claude Memory

Para mantener continuidad entre sesiones:
```bash
# Antes de cerrar sesión con Claude al 90% tokens
git add Docs/claude_prompt.md
git commit -m "docs: sync Claude session $(Get-Date -Format 'yyyy-MM-dd')"
git push origin dev
```

**Archivos clave para Claude:**
- `Docs/claude_prompt.md`: historial de prompts y respuestas
- `changelog.md`: contexto de cambios recientes
- `README.md`: estado actual del proyecto

---

## 8. Sugerencias para .gitignore
```gitignore
# Datos sensibles
*.xlsx
data/confidencial/
*.csv
~$*

# Temporales de Excel
~$*.xlsx
~$*.xls

# Claude temporales
.claude_temp/

# Logs
*.log
```

---

## 9. Mantenimiento y evolución del flujo

Este flujo de trabajo está vivo y puede evolucionar con el tiempo. Toda mejora, ajuste o aprendizaje nuevo debe documentarse en este archivo mediante un pull request.

### Recomendaciones:

- Revisar este documento al menos una vez al mes
- Registrar cambios en `changelog.md`
- Usar Claude para sugerencias de mejora continua
- Compartir este flujo con nuevos colaboradores como punto de partida

---

**Última actualización:** 16 de noviembre de 2025