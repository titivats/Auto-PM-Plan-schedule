$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$python = Join-Path $root "venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    $python = "python"
}

& $python ".\tools\create_icon.py"

if (Test-Path ".\build") {
    Remove-Item -LiteralPath ".\build" -Recurse -Force
}

if (Test-Path ".\dist") {
    Remove-Item -LiteralPath ".\dist" -Recurse -Force
}

& $python -m PyInstaller --noconfirm ".\PMPlanAutoSchedule.spec"

if (Test-Path ".\build") {
    Remove-Item -LiteralPath ".\build" -Recurse -Force
}

Write-Output ""
Write-Output "Built: $root\dist\PMPlanAutoSchedule.exe"
