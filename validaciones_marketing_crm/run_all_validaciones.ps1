$ErrorActionPreference = "Stop"

Write-Host "=========================================="
Write-Host "Ejecutando validaciones Marketing / CRM"
Write-Host "=========================================="

if (!(Test-Path "data\leads_crm.xlsx")) {
    Write-Host "ERROR: No se encontro data\leads_crm.xlsx"
    Write-Host "Coloca tu export del CRM/Formularios en la carpeta data y renombralo como leads_crm.xlsx"
    exit 1
}

Write-Host "1/2 Auditoria leads formularios..."
python audit_leads_formularios.py

Write-Host "2/2 BBDD negativa exclusion..."
python generar_bbdd_negativa.py

Write-Host "=========================================="
Write-Host "Listo. Revisa la carpeta outputs"
Write-Host "=========================================="
