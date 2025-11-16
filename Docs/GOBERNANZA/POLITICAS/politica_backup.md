# 💾 Política de Backup y Retención

> **Propósito:** Garantizar disponibilidad, integridad y recuperabilidad de datos financieros
> **Alcance:** Repositorio Git + archivos Excel/CSV en carpeta `data/`
> **Responsable:** Alvaro Velasco
> **Última revisión:** 16 de noviembre de 2025

---

## 📋 Resumen Ejecutivo

Esta política define cómo se protegen los activos digitales del proyecto Excel_Finance_Project contra pérdida, corrupción o eliminación accidental. Cumple con requerimientos de auditoría financiera que exigen retención mínima de 7 años para documentos contables.

### Principio rector: **3-2-1**
- **3** copias de cada archivo (original + 2 backups)
- **2** medios de almacenamiento diferentes (cloud + local)
- **1** copia fuera del sitio (offsite/cloud)

---

## 🗂️ Clasificación de Activos

### 1. Repositorio Git (código, docs, scripts)
- **Criticidad:** Alta
- **Tipo:** Texto versionado
- **Tamaño:** ~100 MB
- **Frecuencia de cambio:** Diaria
- **Retención:** Ilimitada (historial completo)

### 2. Archivos Excel/CSV (data/)
- **Criticidad:** Crítica
- **Tipo:** Binario
- **Tamaño:** ~500 MB - 2 GB
- **Frecuencia de cambio:** Diaria a semanal
- **Retención:** 7 años (requerimiento fiscal Costa Rica)

### 3. Documentos generados por Claude (Docs/)
- **Criticidad:** Media
- **Tipo:** Texto (Markdown)
- **Tamaño:** ~50 MB
- **Frecuencia de cambio:** Semanal
- **Retención:** Ilimitada

---

## 📅 Estrategia de Backup

### A. Repositorio Git

#### Ubicaciones de backup
```
PRIMARY:   GitHub (https://github.com/cimsa8-cyber/Excel_Finance_Project)
           ├── Automático en cada `git push`
           └── Redundancia geográfica (servidores GitHub)

MIRROR 1:  GitLab (backup espejo)
           ├── Sincronización: Semanal (sábados 2:00 AM)
           └── URL: https://gitlab.com/cimsa8-cyber/Excel_Finance_Project

MIRROR 2:  Disco local (laptop)
           ├── Ubicación: C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad\.git
           └── Actualización: Continua (git pull)

MIRROR 3:  Disco externo (offline backup)
           ├── Ubicación: E:\Backups_Git\Excel_Finance_Project\
           └── Frecuencia: Mensual (último viernes del mes)
```

#### Comandos de backup Git
```powershell
# Backup a GitLab (ejecutar semanalmente)
git remote add gitlab https://gitlab.com/cimsa8-cyber/Excel_Finance_Project.git
git push gitlab --all
git push gitlab --tags

# Backup a disco externo (ejecutar mensualmente)
# 1. Conectar disco externo E:\
# 2. Ejecutar:
$date = Get-Date -Format "yyyy-MM"
git bundle create "E:\Backups_Git\Excel_Finance_Project_$date.bundle" --all
# 3. Verificar:
git bundle verify "E:\Backups_Git\Excel_Finance_Project_$date.bundle"
```

---

### B. Archivos Excel/CSV (data/)

#### Estrategia de backup incremental + snapshot

```
DAILY INCREMENTAL:  OneDrive
                    ├── Ruta: OneDrive\Backups\FinanzasContabilidad\Daily\
                    ├── Retención: Últimos 30 días
                    └── Sincronización: Automática (OneDrive sync)

WEEKLY COMPLETE:    Google Drive
                    ├── Ruta: Google Drive\Backups\FinanzasContabilidad\Weekly\
                    ├── Retención: Últimas 12 semanas
                    └── Sincronización: Manual (sábados)

MONTHLY SNAPSHOT:   Disco externo
                    ├── Ruta: E:\Backups\FinanzasContabilidad\Monthly\
                    ├── Retención: 7 años
                    └── Formato: ZIP con password

QUARTERLY ARCHIVE:  Disco externo offline (guardado en caja fuerte)
                    ├── Ruta: F:\Archives\FinanzasContabilidad\
                    ├── Retención: Permanente
                    └── Formato: ZIP encriptado AES-256
```

#### Script de backup automático (Windows Task Scheduler)

**Archivo:** `scripts/backup_daily.ps1`
```powershell
# ============================================
# Script de Backup Diario - Excel Finance Project
# Ejecutar: Diariamente a las 11:00 PM (Task Scheduler)
# ============================================

param(
    [string]$SourceDir = "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad\data",
    [string]$BackupRoot = "$env:OneDrive\Backups\FinanzasContabilidad\Daily",
    [switch]$GenerateChecksum = $true,
    [switch]$SendEmail = $false
)

# Configuración
$date = Get-Date -Format "yyyy-MM-dd"
$destDir = Join-Path $BackupRoot $date
$logFile = Join-Path $destDir "backup.log"

# Crear directorio de destino
New-Item -ItemType Directory -Path $destDir -Force | Out-Null

# Iniciar log
Start-Transcript -Path $logFile

Write-Host "================================================"
Write-Host "BACKUP DIARIO - Excel Finance Project"
Write-Host "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "================================================`n"

# Verificar que directorio fuente existe
if (-not (Test-Path $SourceDir)) {
    Write-Error "❌ Error: Directorio fuente no existe: $SourceDir"
    exit 1
}

# Ejecutar backup incremental con robocopy
Write-Host "📁 Copiando archivos..."
Write-Host "   Origen: $SourceDir"
Write-Host "   Destino: $destDir`n"

$robocopyLog = Join-Path $destDir "robocopy.log"
robocopy $SourceDir $destDir /MIR /R:3 /W:5 /LOG:$robocopyLog /NP /NDL

# Verificar resultado
if ($LASTEXITCODE -le 7) {
    Write-Host "✅ Backup completado exitosamente"
} else {
    Write-Error "❌ Error en backup (código: $LASTEXITCODE)"
    exit $LASTEXITCODE
}

# Generar checksums para integridad
if ($GenerateChecksum) {
    Write-Host "`n🔐 Generando checksums de integridad..."
    $checksumFile = Join-Path $destDir "checksums.csv"

    Get-ChildItem $destDir -Recurse -File |
        Where-Object { $_.Extension -match '\.(xlsx|xls|csv|xlsm)$' } |
        ForEach-Object {
            [PSCustomObject]@{
                File = $_.FullName.Replace($destDir, ".")
                SHA256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
                Size = $_.Length
                Modified = $_.LastWriteTime
            }
        } | Export-Csv $checksumFile -NoTypeInformation

    Write-Host "✅ Checksums guardados en: checksums.csv"
}

# Limpieza de backups antiguos (>30 días)
Write-Host "`n🧹 Limpiando backups antiguos..."
$cutoffDate = (Get-Date).AddDays(-30)
Get-ChildItem $BackupRoot -Directory |
    Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' -and $_.CreationTime -lt $cutoffDate } |
    ForEach-Object {
        Write-Host "   Eliminando: $($_.Name)"
        Remove-Item $_.FullName -Recurse -Force
    }

# Resumen
Write-Host "`n================================================"
Write-Host "RESUMEN DE BACKUP"
Write-Host "================================================"
Write-Host "Directorio: $destDir"
Write-Host "Archivos copiados: $(( Get-ChildItem $destDir -Recurse -File | Measure-Object ).Count)"
Write-Host "Tamaño total: $([math]::Round((Get-ChildItem $destDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB, 2)) MB"
Write-Host "Duración: $((Get-Date) - $date)"
Write-Host "================================================`n"

Stop-Transcript

# Opcional: Enviar email de confirmación
if ($SendEmail) {
    # Configurar con tus credenciales SMTP
    # Send-MailMessage -To "alvaro@example.com" -Subject "Backup exitoso $date" ...
}

exit 0
```

**Configurar en Task Scheduler (Windows):**
```powershell
# Ejecutar como Administrador en PowerShell

$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -File `"C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad\scripts\backup_daily.ps1`""

$trigger = New-ScheduledTaskTrigger -Daily -At "11:00PM"

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable

Register-ScheduledTask -TaskName "Excel Finance Backup Daily" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Backup diario automático del proyecto Excel Finance"
```

---

### C. Backup Semanal (Google Drive)

**Archivo:** `scripts/backup_weekly.ps1`
```powershell
# Backup semanal completo a Google Drive
# Ejecutar: Sábados a las 3:00 AM

$date = Get-Date -Format "yyyy-MM-dd"
$week = Get-Date -UFormat "%Y-W%V"
$source = "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad\data"
$dest = "G:\Mi unidad\Backups\FinanzasContabilidad\Weekly\$week"  # G: = Google Drive

# Crear backup completo
New-Item -ItemType Directory -Path $dest -Force | Out-Null
robocopy $source $dest /MIR /LOG:"$dest\backup_$date.log"

# Comprimir para ahorrar espacio
Compress-Archive -Path $dest -DestinationPath "$dest.zip" -CompressionLevel Optimal
Remove-Item $dest -Recurse -Force

Write-Host "✅ Backup semanal completado: $dest.zip"

# Retener solo últimas 12 semanas
$parent = Split-Path $dest -Parent
Get-ChildItem $parent -Filter "*.zip" |
    Sort-Object CreationTime -Descending |
    Select-Object -Skip 12 |
    Remove-Item -Force
```

---

### D. Snapshot Mensual (Disco Externo)

**Proceso manual (último viernes del mes):**
```powershell
# 1. Conectar disco externo E:\
# 2. Ejecutar script de snapshot

$month = Get-Date -Format "yyyy-MM"
$source = "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad"
$dest = "E:\Backups\FinanzasContabilidad\Monthly\$month"

# Backup completo del proyecto
robocopy $source $dest /MIR /XD .git node_modules /LOG:"$dest\backup.log"

# Comprimir con contraseña
7z a -tzip -p -mhe=on "$dest.zip" $dest

# Generar checksum del ZIP
certutil -hashfile "$dest.zip" SHA256 > "$dest.zip.sha256"

Write-Host "✅ Snapshot mensual completado"
Write-Host "📦 Archivo: $dest.zip"
Write-Host "🔐 Checksum: $dest.zip.sha256"
Write-Host ""
Write-Host "⚠️  IMPORTANTE: Guardar contraseña en gestor de contraseñas"
```

---

### E. Archivo Trimestral (Offline, Encriptado)

**Proceso manual (fin de trimestre fiscal: mar, jun, sep, dic):**
```powershell
# Backup trimestral encriptado para compliance a largo plazo

$quarter = "Q$(([math]::Ceiling((Get-Date).Month / 3)))-$(Get-Date -Format 'yyyy')"
$source = "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad"
$dest = "F:\Archives\FinanzasContabilidad\$quarter"  # F: = Disco externo offline

# 1. Crear snapshot completo
robocopy $source $dest /MIR /XD .git /LOG:"$dest\backup.log"

# 2. Generar checksums ANTES de comprimir
python scripts/audit_integrity_check.py --data-dir $dest\data --output $dest\checksums.json

# 3. Comprimir con encriptación AES-256 (usar 7-Zip)
7z a -t7z -p -mhe=on -mx=9 -mcu=on "$dest.7z" $dest

# 4. Checksum del archivo comprimido
certutil -hashfile "$dest.7z" SHA256 > "$dest.7z.sha256"

# 5. Crear README con metadata
@"
# Archivo Trimestral - $quarter
- Fecha creación: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
- Período fiscal: $quarter
- Tamaño original: $((Get-ChildItem $dest -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB) MB
- Tamaño comprimido: $((Get-Item "$dest.7z").Length / 1MB) MB
- SHA256: $(Get-Content "$dest.7z.sha256" -Tail 1)
- Retención: 7 años (hasta $(((Get-Date).AddYears(7)).ToString('yyyy-MM-dd')))
- Encriptación: AES-256
- Contraseña: [Ver gestor de contraseñas - Entry: Finance_Archive_$quarter]
"@ | Out-File "$dest.README.txt"

Write-Host "✅ Archivo trimestral creado"
Write-Host "📦 Archivo: $dest.7z"
Write-Host "📄 Metadata: $dest.README.txt"
Write-Host ""
Write-Host "⚠️  ACCIÓN REQUERIDA:"
Write-Host "   1. Verificar archivo: 7z t $dest.7z"
Write-Host "   2. Guardar contraseña en KeePass/1Password"
Write-Host "   3. Etiquetar disco externo con: FINANCE_ARCHIVE_$quarter"
Write-Host "   4. Guardar en caja fuerte/ubicación segura"
Write-Host "   5. Registrar en inventario de backups"
```

---

## 🔄 Plan de Disaster Recovery

### Escenario 1: Pérdida de repositorio local

**Síntomas:** Laptop robada, disco dañado, eliminación accidental de carpeta

**RTO (Recovery Time Objective):** 2 horas
**RPO (Recovery Point Objective):** Último push a GitHub

**Procedimiento:**
```powershell
# 1. Instalar Git en nueva máquina
# 2. Clonar desde GitHub
git clone https://github.com/cimsa8-cyber/Excel_Finance_Project
cd Excel_Finance_Project

# 3. Configurar Git
git config user.name "Alvaro Velasco"
git config user.email "tu@email.com"
git config commit.gpgsign true

# 4. Restaurar archivos data/ desde último backup
# Opción A: OneDrive (más reciente)
Copy-Item "$env:OneDrive\Backups\FinanzasContabilidad\Daily\2025-11-16\*" -Destination "data\" -Recurse

# Opción B: Google Drive (semanal)
# Descomprimir último backup semanal
Expand-Archive "G:\Mi unidad\Backups\FinanzasContabilidad\Weekly\2025-W46.zip" -DestinationPath "temp\"
Copy-Item "temp\*" -Destination "data\" -Recurse

# 5. Verificar integridad
python scripts/audit_integrity_check.py --verify audit_checksums.json

# 6. Validar que todo funciona
# - Abrir archivos Excel
# - Ejecutar scripts Python
# - Verificar fórmulas

Write-Host "✅ Recuperación completada"
```

---

### Escenario 2: Corrupción de archivo Excel crítico

**Síntomas:** Archivo no abre, errores en fórmulas, datos incoherentes

**RTO:** 30 minutos
**RPO:** Último backup (máximo 24 horas)

**Procedimiento:**
```powershell
# 1. Identificar archivo corrupto
$corruptedFile = "data\balance_2024.xlsx"

# 2. Renombrar archivo corrupto (no eliminar aún)
Move-Item $corruptedFile "$corruptedFile.corrupted.$(Get-Date -Format 'yyyyMMdd-HHmmss')"

# 3. Restaurar desde backup más reciente
# Verificar OneDrive primero
$backupDate = (Get-ChildItem "$env:OneDrive\Backups\FinanzasContabilidad\Daily" | Sort-Object Name -Descending | Select-Object -First 1).Name
Copy-Item "$env:OneDrive\Backups\FinanzasContabilidad\Daily\$backupDate\balance_2024.xlsx" -Destination "data\"

# 4. Validar integridad del archivo restaurado
python scripts/audit_integrity_check.py

# 5. Abrir Excel y verificar datos manualmente
# - Revisar fórmulas clave
# - Comparar totales con reportes conocidos
# - Validar fechas de última modificación

# 6. Si archivo restaurado está OK, documentar incidente
git add data/balance_2024.xlsx
git commit -m "fix: restaurar balance_2024.xlsx desde backup $backupDate por corrupción detectada

Detalles:
- Archivo original corrupto guardado como: balance_2024.xlsx.corrupted.*
- Restaurado desde: OneDrive\Daily\$backupDate
- Validación de integridad: OK
- Pérdida de datos: [Última modificación en backup: verificar timestamp]
"

# 7. Analizar causa de corrupción
# - ¿Fue cierre inesperado de Excel?
# - ¿Problemas de disco?
# - ¿Virus/malware?
# Tomar medidas preventivas
```

---

### Escenario 3: Commit erróneo en branch main

**Síntomas:** Merge accidental, commit con errores en producción

**RTO:** 1 hora
**RPO:** N/A (revertir en Git no pierde datos)

**Procedimiento:**
```bash
# NUNCA hacer git reset --hard en main
# SIEMPRE usar git revert para mantener historial auditable

# 1. Identificar commit problemático
git log --oneline main
# Ejemplo: commit abc123 tiene el error

# 2. Revertir commit (crea nuevo commit que deshace cambios)
git revert abc123 -m "Revert: [razón del rollback]

Detalles:
- Commit revertido: abc123
- Razón: [descripción del problema detectado]
- Detectado por: Alvaro Velasco
- Validado por: [nombre auditor si aplica]
- Issue: #[número]
"

# 3. Si fueron múltiples commits, revertir en orden inverso
git revert abc125 abc124 abc123

# 4. Push del revert
git push origin main

# 5. Documentar en changelog.md
# Ver sección "Plantilla de rollback" más abajo

# 6. Notificar a stakeholders
# - Enviar email a auditores/contadores
# - Actualizar tickets relacionados
# - Documentar lecciones aprendidas
```

---

### Escenario 4: Pérdida total (desastre)

**Síntomas:** Incendio, robo de todos los equipos, falla de múltiples servicios cloud

**RTO:** 1 día
**RPO:** Último backup mensual (máximo 30 días)

**Procedimiento:**
```powershell
# FASE 1: Recuperación del repositorio (1-2 horas)

# 1. Nueva máquina, instalar herramientas
# - Git, Python, Excel, VSCode
# - Ver: Docs/GUIA_COMPLETA_TRABAJO_LOCAL.md

# 2. Recuperar desde GitHub (si disponible)
git clone https://github.com/cimsa8-cyber/Excel_Finance_Project

# 3. Si GitHub no disponible, usar bundle del disco externo
git clone E:\Backups_Git\Excel_Finance_Project_2025-11.bundle Excel_Finance_Project

# FASE 2: Recuperación de datos (2-4 horas)

# 4. Recuperar archivos desde disco externo mensual
$lastMonthly = "E:\Backups\FinanzasContabilidad\Monthly\2025-11"
7z x "$lastMonthly.zip" -p  # Solicita contraseña

# 5. Copiar data/ al repositorio
Copy-Item "$lastMonthly\data\*" -Destination "Excel_Finance_Project\data\" -Recurse

# 6. Validar checksums
cd Excel_Finance_Project
python scripts/audit_integrity_check.py --verify "$lastMonthly\checksums.json"

# FASE 3: Reconstrucción de cambios recientes (4-8 horas)

# 7. Identificar trabajo perdido (desde último backup hasta desastre)
# - Revisar emails con reportes enviados
# - Consultar con contador/auditor qué reportes recibieron
# - Revisar impresiones físicas si existen

# 8. Recrear cambios manualmente
# - Priorizar datos críticos (transacciones, balance)
# - Documentar estimaciones si es necesario
# - Marcar claramente datos reconstruidos vs originales

# FASE 4: Validación y documentación (2-4 horas)

# 9. Validación cruzada
# - Comparar totales con reportes externos (bancos, proveedores)
# - Validar con auditor/contador
# - Ejecutar reconciliaciones

# 10. Documentar recuperación
git add -A
git commit -m "disaster: recuperación completa desde backup $lastMonthly

Contexto del desastre:
- Fecha del incidente: [fecha]
- Último backup utilizado: $lastMonthly
- Datos perdidos: [período desde último backup hasta desastre]
- Datos reconstruidos: [lista de archivos/transacciones]
- Validado por: [nombre auditor]
- Fecha de validación: [fecha]

Notas:
- [Detalles adicionales sobre el proceso de recuperación]
- [Lecciones aprendidas]
- [Mejoras a implementar en política de backup]
"

# 11. Actualizar TRAZABILIDAD.md con nota de recuperación

# 12. Restablecer backups automáticos
# - Reconfigurar Task Scheduler
# - Validar sincronización OneDrive/Google Drive
# - Crear nuevo snapshot mensual
```

**Checklist post-desastre:**
- [ ] Repositorio Git completamente funcional
- [ ] Todos los archivos Excel abren sin errores
- [ ] Checksums validados vs último backup conocido
- [ ] Fórmulas validadas manualmente
- [ ] Totales reconciliados con fuentes externas
- [ ] Auditor/contador ha validado datos reconstruidos
- [ ] Backups automáticos restablecidos y probados
- [ ] Documentación de desastre y recuperación completada
- [ ] Post-mortem realizado (causas, mejoras, prevención)

---

## 🧪 Testing de Backups

### Frecuencia: **Trimestral** (junto con archivo trimestral)

**Objetivo:** Validar que los backups son recuperables y útiles en caso de desastre real.

**Proceso de testing:**
```powershell
# TEST 1: Recuperación de archivo individual (15 min)
# ======================================================

# 1. Seleccionar archivo aleatorio de data/
$testFile = Get-ChildItem data\*.xlsx | Get-Random

# 2. Crear copia temporal
Copy-Item $testFile "$testFile.original"

# 3. Eliminar archivo
Remove-Item $testFile

# 4. Restaurar desde backup OneDrive
$latestBackup = Get-ChildItem "$env:OneDrive\Backups\FinanzasContabilidad\Daily" | Sort-Object Name -Descending | Select-Object -First 1
Copy-Item "$env:OneDrive\Backups\FinanzasContabilidad\Daily\$($latestBackup.Name)\$($testFile.Name)" -Destination (Split-Path $testFile)

# 5. Validar integridad
python scripts/audit_integrity_check.py --verify audit_checksums.json

# 6. Comparar archivos
fc $testFile "$testFile.original" > $null
if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ TEST 1 PASSED: Archivo restaurado correctamente"
} else {
    Write-Host "❌ TEST 1 FAILED: Archivo restaurado difiere del original"
}

# 7. Limpiar
Remove-Item "$testFile.original"
```

```powershell
# TEST 2: Recuperación de repositorio completo (30 min)
# ======================================================

# 1. Crear directorio temporal
$testDir = "C:\Temp\BackupTest_$(Get-Date -Format 'yyyyMMddHHmmss')"
New-Item -ItemType Directory -Path $testDir | Out-Null

# 2. Clonar desde GitHub
git clone https://github.com/cimsa8-cyber/Excel_Finance_Project $testDir\repo

# 3. Restaurar data/ desde backup semanal
$latestWeekly = Get-ChildItem "G:\Mi unidad\Backups\FinanzasContabilidad\Weekly" | Sort-Object Name -Descending | Select-Object -First 1
Expand-Archive $latestWeekly.FullName -DestinationPath $testDir\weekly
Copy-Item "$testDir\weekly\*" -Destination "$testDir\repo\data\" -Recurse

# 4. Validar integridad
Push-Location "$testDir\repo"
python scripts/audit_integrity_check.py
$testResult = $LASTEXITCODE
Pop-Location

# 5. Verificar resultado
if ($testResult -eq 0) {
    Write-Host "✅ TEST 2 PASSED: Repositorio recuperado completamente"
} else {
    Write-Host "❌ TEST 2 FAILED: Problemas en recuperación de repositorio"
}

# 6. Limpiar
Remove-Item $testDir -Recurse -Force
```

```powershell
# TEST 3: Verificación de backup mensual encriptado (15 min)
# ===========================================================

# 1. Seleccionar backup mensual más reciente
$latestMonthly = Get-ChildItem "E:\Backups\FinanzasContabilidad\Monthly\*.zip" | Sort-Object Name -Descending | Select-Object -First 1

# 2. Verificar integridad del ZIP
$actualHash = (Get-FileHash $latestMonthly.FullName -Algorithm SHA256).Hash
$expectedHash = (Get-Content "$($latestMonthly.FullName).sha256" | Select-Object -Last 1).Trim()

if ($actualHash -eq $expectedHash) {
    Write-Host "✅ Checksum del ZIP coincide"
} else {
    Write-Host "❌ ALERTA: Checksum del ZIP NO coincide - posible corrupción"
    exit 1
}

# 3. Extraer archivo (requiere contraseña)
$testExtract = "C:\Temp\MonthlyTest_$(Get-Date -Format 'yyyyMMdd')"
7z x $latestMonthly.FullName -o$testExtract -p  # Solicita contraseña

# 4. Validar que se extrajo correctamente
if (Test-Path "$testExtract\data") {
    $fileCount = (Get-ChildItem "$testExtract\data" -Recurse -File | Measure-Object).Count
    Write-Host "✅ TEST 3 PASSED: Backup mensual extraído ($fileCount archivos)"
} else {
    Write-Host "❌ TEST 3 FAILED: No se pudo extraer backup mensual"
}

# 5. Limpiar
Remove-Item $testExtract -Recurse -Force
```

**Documentar resultado en changelog.md:**
```markdown
## Test de Backups - Q4 2025

**Fecha:** 2025-11-16
**Ejecutado por:** Alvaro Velasco

### Resultados
- ✅ Test 1: Recuperación archivo individual - PASSED
- ✅ Test 2: Recuperación repositorio completo - PASSED
- ✅ Test 3: Backup mensual encriptado - PASSED

### Métricas
- Archivos validados: 47/47
- Tiempo de restauración: 12 minutos
- Tamaño total backups: 2.3 GB
- Backups disponibles:
  - Diarios: 30 (último 30 días)
  - Semanales: 12 (últimas 12 semanas)
  - Mensuales: 12 (último año)
  - Trimestrales: 4 (último año)

### Acciones correctivas
- Ninguna requerida

### Próximo test programado
- Fecha: 2026-02-16
- Responsable: Alvaro Velasco
```

---

## 📊 Inventario de Backups

**Actualizar mensualmente en:** `Docs/INVENTARIO_BACKUPS.md`

| Tipo | Ubicación | Última actualización | Tamaño | Retención | Estado |
|------|-----------|----------------------|--------|-----------|--------|
| Git - GitHub | github.com/cimsa8-cyber | 2025-11-16 | 120 MB | Ilimitado | 🟢 OK |
| Git - GitLab | gitlab.com/cimsa8-cyber | 2025-11-13 | 120 MB | Ilimitado | 🟢 OK |
| Git - Bundle E:\ | E:\Backups_Git\ | 2025-11-01 | 125 MB | Ilimitado | 🟢 OK |
| Daily - OneDrive | OneDrive\Backups\Daily | 2025-11-16 | 850 MB | 30 días | 🟢 OK |
| Weekly - GDrive | GDrive\Backups\Weekly | 2025-11-13 | 2.1 GB | 12 semanas | 🟢 OK |
| Monthly - E:\ | E:\Backups\Monthly | 2025-11-01 | 980 MB | 7 años | 🟢 OK |
| Quarterly - F:\ | F:\Archives\Q3-2025 | 2025-09-30 | 1.2 GB | Permanente | 🟢 OK |

**Leyenda:**
- 🟢 OK: Backup actualizado y validado
- 🟡 WARNING: Backup desactualizado (>7 días del esperado)
- 🔴 CRITICAL: Backup faltante o corrupto

---

## 📝 Plantilla de Documentación de Rollback

**Usar en `changelog.md` cuando se ejecute un rollback:**

```markdown
### YYYY-MM-DD - 🔄 ROLLBACK CRÍTICO

#### Contexto
- **Commit revertido:** [hash]
- **Rama afectada:** [main/dev/otra]
- **Razón del rollback:** [descripción del problema]
- **Detectado por:** [nombre]
- **Fecha de detección:** [fecha/hora]
- **Issue relacionado:** #[número]

#### Impacto
- **Archivos afectados:**
  - archivo1.xlsx
  - archivo2.py
- **Reportes impactados:** [listado]
- **Período afectado:** [desde - hasta]
- **Magnitud:** [ALTA/MEDIA/BAJA]

#### Acciones Tomadas
1. [Descripción paso 1]
2. [Descripción paso 2]
3. ...

#### Recuperación
- **Método utilizado:** [git revert / restauración desde backup / manual]
- **Backup restaurado:** [ruta y fecha del backup]
- **Validación:**
  - [ ] Checksums verificados
  - [ ] Fórmulas validadas
  - [ ] Datos reconciliados
  - [ ] Aprobado por auditor/contador
- **Aprobado por:** [nombre del aprobador]
- **Fecha de aprobación:** [fecha]

#### Lecciones Aprendidas
- [Qué falló]
- [Cómo se puede prevenir en el futuro]
- [Mejoras a implementar]

#### Seguimiento
- [ ] Actualizar documentación afectada
- [ ] Crear issue para mejoras preventivas: #[número]
- [ ] Comunicar a stakeholders
- [ ] Actualizar TRAZABILIDAD.md
```

---

## 🔐 Seguridad y Acceso

### Contraseñas de backups encriptados
- **Almacenamiento:** KeePass / 1Password / Bitwarden
- **Entrada:** `Finance_Archive_YYYY-QX`
- **Compartir:** Solo con auditor autorizado, mediante canal seguro

### Acceso a servicios cloud
- **GitHub:** 2FA habilitado (obligatorio)
- **OneDrive:** 2FA habilitado
- **Google Drive:** 2FA habilitado

### Cifrado
- **Backups mensuales:** ZIP con contraseña (AES-256 via 7-Zip)
- **Backups trimestrales:** 7z con AES-256
- **Git:** Commits firmados con GPG (para commits críticos)

---

## ✅ Checklist de Cumplimiento

### Diario
- [ ] Backup automático OneDrive ejecutado (verificar Task Scheduler)
- [ ] Sincronización OneDrive completada

### Semanal (Sábados)
- [ ] Backup manual a Google Drive ejecutado
- [ ] Verificar espacio disponible en cloud
- [ ] Sincronizar repositorio a GitLab

### Mensual (Último viernes)
- [ ] Snapshot mensual a disco externo E:\
- [ ] Generar checksum del snapshot
- [ ] Comprimir con contraseña
- [ ] Guardar contraseña en gestor
- [ ] Actualizar inventario de backups

### Trimestral (Fin de Q)
- [ ] Archivo trimestral encriptado a disco F:\
- [ ] Test de backups (3 tests completos)
- [ ] Documentar resultados en changelog.md
- [ ] Validar cumplimiento de retención (eliminar backups >7 años si aplica)
- [ ] Revisar y actualizar esta política

### Anual
- [ ] Auditoría externa de backups (si aplica)
- [ ] Revisión completa de disaster recovery plan
- [ ] Actualización de esta política
- [ ] Training/refresher de procedimientos

---

## 📞 Contactos de Emergencia

En caso de desastre crítico:

1. **Responsable principal:** Alvaro Velasco
   - Email: [tu@email.com]
   - Teléfono: [número]

2. **Auditor/Contador:** [Nombre]
   - Email: [email]
   - Teléfono: [número]

3. **Soporte técnico GitHub:** https://support.github.com
4. **Soporte OneDrive:** https://support.microsoft.com
5. **Soporte Google Drive:** https://support.google.com

---

**Aprobado por:** Alvaro Velasco
**Fecha de aprobación:** 16 de noviembre de 2025
**Próxima revisión:** 16 de noviembre de 2026
**Versión:** 1.0
