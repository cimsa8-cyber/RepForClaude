#!/usr/bin/env pwsh
<#
.SYNOPSIS
Setup automático de estructura de Gobernanza Técnica

.DESCRIPTION
Crea toda la estructura de carpetas y archivos de gobernanza
para el proyecto Excel_Finance_Project v5.0

.EXAMPLE
.\scripts\setup_gobernanza.ps1

.NOTES
Autor: Alvaro Velasco
Fecha: 2025-11-16
Proyecto: Excel_Finance_Project
Versión: 1.0 (Corregida)
#>

# ===================================
# CONFIGURACIÓN
# ===================================

$ErrorActionPreference = "Stop"
$BASE_PATH = "C:\Users\Alvaro Velasco\desktop\Proyecto FinanzasContabilidad"
$FECHA_HOY = Get-Date -Format "yyyy-MM-dd"
$AUTOR = "Alvaro Velasco"

# ===================================
# FUNCIONES AUXILIARES
# ===================================

function Write-Header {
    param([string]$Texto)
    Write-Host "`n$('=' * 60)" -ForegroundColor Cyan
    Write-Host $Texto -ForegroundColor Cyan
    Write-Host "$('=' * 60)`n" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Texto)
    Write-Host "✅ $Texto" -ForegroundColor Green
}

function Write-Info {
    param([string]$Texto)
    Write-Host "ℹ️  $Texto" -ForegroundColor Yellow
}

function Write-Error-Custom {
    param([string]$Texto)
    Write-Host "❌ $Texto" -ForegroundColor Red
}

function Create-Directory {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Success "Carpeta creada: $Path"
    } else {
        Write-Info "Carpeta ya existe: $Path"
    }
}

function Create-File {
    param(
        [string]$Path,
        [string]$Content,
        [string]$Description
    )

    if (-not (Test-Path $Path)) {
        Set-Content -Path $Path -Value $Content -Encoding UTF8
        Write-Success "$Description"
    } else {
        Write-Info "Archivo ya existe: $Path (no sobrescrito)"
    }
}

# ===================================
# FUNCIÓN PRINCIPAL
# ===================================

function Main {
    try {
        Write-Header "SETUP GOBERNANZA TÉCNICA v1.0"
        Write-Host "Proyecto: Excel_Finance_Project" -ForegroundColor White
        Write-Host "Ruta: $BASE_PATH" -ForegroundColor White
        Write-Host "Fecha: $FECHA_HOY`n" -ForegroundColor White

        # Verificar que estamos en el directorio correcto
        if (-not (Test-Path $BASE_PATH)) {
            Write-Error-Custom "ERROR: Directorio no encontrado: $BASE_PATH"
            Write-Host "`nVerifica la ruta y vuelve a ejecutar el script." -ForegroundColor Yellow
            exit 1
        }

        Set-Location $BASE_PATH
        Write-Success "Directorio de trabajo establecido"

        # Crear estructura de carpetas
        Write-Header "PASO 1: Creando Estructura de Carpetas"

        Create-Directory "$BASE_PATH\Docs\GOBERNANZA"
        Create-Directory "$BASE_PATH\Docs\GOBERNANZA\TEMPLATES"
        Create-Directory "$BASE_PATH\Docs\GOBERNANZA\POLITICAS"
        Create-Directory "$BASE_PATH\Docs\GOBERNANZA\ANALISIS"
        Create-Directory "$BASE_PATH\.github"

        # Crear archivos
        Write-Header "PASO 2: Generando Archivos de Gobernanza"

        # Nota: Los contenidos de archivos se copiarán desde el repositorio integrado
        # Este script solo crea la estructura. Los archivos finales se copian manualmente
        # o se descargan desde RepForClaude

        Write-Info "Estructura de carpetas creada exitosamente"
        Write-Info "Los archivos de contenido deben copiarse desde el repositorio integrado"

        # Resumen final
        Write-Header "✅ SETUP COMPLETADO EXITOSAMENTE"

        Write-Host "`n📊 RESUMEN:" -ForegroundColor Cyan
        Write-Host "  • Carpetas creadas:" -ForegroundColor White
        Write-Host "    - Docs\GOBERNANZA\" -ForegroundColor Gray
        Write-Host "    - Docs\GOBERNANZA\TEMPLATES\" -ForegroundColor Gray
        Write-Host "    - Docs\GOBERNANZA\POLITICAS\" -ForegroundColor Gray
        Write-Host "    - Docs\GOBERNANZA\ANALISIS\" -ForegroundColor Gray
        Write-Host "    - .github\" -ForegroundColor Gray

        Write-Host "`n📂 PRÓXIMOS PASOS:" -ForegroundColor Cyan
        Write-Host "  1. Copiar archivos desde RepForClaude/Docs/GOBERNANZA/" -ForegroundColor White
        Write-Host "  2. Revisar y personalizar según necesidades" -ForegroundColor White
        Write-Host "  3. Hacer commit:" -ForegroundColor White
        Write-Host "     git add Docs/GOBERNANZA/" -ForegroundColor Gray
        Write-Host "     git add scripts/setup_gobernanza.ps1" -ForegroundColor Gray
        Write-Host "     git add .github/pull_request_template.md" -ForegroundColor Gray
        Write-Host "     git commit -m `"docs: Implementar gobernanza técnica v1.0`"" -ForegroundColor Gray
        Write-Host "     git push origin dev" -ForegroundColor Gray

        Write-Host "`n📖 DOCUMENTACIÓN:" -ForegroundColor Cyan
        Write-Host "  Leer: Docs\GOBERNANZA\README.md (primero)" -ForegroundColor White
        Write-Host "  Luego: Docs\GOBERNANZA\gobernanza_tecnica.md" -ForegroundColor White

        Write-Host "`n🎉 ¡ESTRUCTURA DE GOBERNANZA CREADA!" -ForegroundColor Green
        Write-Host "`n"

    }
    catch {
        Write-Header "❌ ERROR"
        Write-Host "Ocurrió un error durante la ejecución:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "`nStack Trace:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor Gray
        exit 1
    }
}

# ===================================
# EJECUTAR
# ===================================

Main
Read-Host "`nPresiona ENTER para cerrar"
