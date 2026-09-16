$ErrorActionPreference = "Stop"

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$exporter = Join-Path $PSScriptRoot "scripts\export-catalogs.py"

if (-not (Test-Path $python)) {
    Write-Error "Python environment not found. Create .venv and install openpyxl first."
}

& $python $exporter
Write-Host "Catalogs synchronized from the Excel workbook. Review the catalogs folder before publishing."
