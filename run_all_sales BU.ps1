param(
    [switch]$SkipRedshift,
    [switch]$ContinueOnRedshiftFail
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " PROPIETARIOS " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Use-ProjectEnv {
    if (Test-Path ".\env\Scripts\Activate.ps1") {
        Write-Host "Activando env existente: env" -ForegroundColor Green
        . ".\env\Scripts\Activate.ps1"
    }
    elseif (Test-Path ".\.venv\Scripts\Activate.ps1") {
        Write-Host "Activando env existente: .venv" -ForegroundColor Green
        . ".\.venv\Scripts\Activate.ps1"
    }
    else {
        Write-Host "No se encontró env/.venv. Creando .venv..." -ForegroundColor Yellow
        python -m venv .venv
        . ".\.venv\Scripts\Activate.ps1"
    }
}

function Invoke-Step {
    param(
        [string]$StepName,
        [scriptblock]$Command
    )

    Write-Host ""
    Write-Host $StepName -ForegroundColor Cyan
    & $Command

    if ($LASTEXITCODE -ne 0) {
        throw "Falló el paso: $StepName"
    }
}

Use-ProjectEnv

Write-Host "Instalando / validando dependencias..." -ForegroundColor Cyan
python -m pip install --upgrade pip
if (Test-Path ".\requirements.txt") {
    pip install -r requirements.txt
}
else {
    pip install pandas numpy pyarrow openpyxl python-dotenv sqlalchemy psycopg2-binary rapidfuzz scikit-learn pyyaml
}

if (-not $SkipRedshift) {
    try {
        Invoke-Step "0/11 Extrayendo Redshift con cache diario..." {
            python .\tools\extract_redshift_daily.py
        }
    }
    catch {
        if ($ContinueOnRedshiftFail) {
            Write-Host "Redshift falló, pero continúo con parquets locales existentes." -ForegroundColor Yellow
        }
        else {
            throw
        }
    }
}
else {
    Write-Host "0/11 Saltando Redshift. Usando parquets locales existentes..." -ForegroundColor Yellow
}

Invoke-Step "1/ Ejecutando Ventas..." {
    python .\ventas_por_cobrar\main_pipeline.py
}

Invoke-Step "2.0/11 Ejecutando Ingresos Cobrados..." {
    
    $DashboardPath = Join-Path $PSScriptRoot "ventas_dashboard\app_ventas_dashboard.py"
    streamlit run $DashboardPath
}


    
Write-Host "Validando patch EFAC priority en cobranza merge..." -ForegroundColor Cyan
$MergeScript = ".\3_merge_venta_recibido\cobranza_pipeline.py"
$MergeContent = Get-Content $MergeScript -Raw
if ($MergeContent -notmatch "add_efac_priority_columns_to_pagos_eventos") {
    Write-Host "ADVERTENCIA: No detecto patch EFAC priority en $MergeScript" -ForegroundColor Yellow
} else {
    Write-Host "OK: patch EFAC priority detectado." -ForegroundColor Green
}


Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host " PIPELINE COMPLETADO" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green
Write-Host "Output matriz:" -ForegroundColor Green
Write-Host "data\gold\power_bi\matriz_venta_cobranza" -ForegroundColor Green
Write-Host ""