# Build Windows .exe for Kristina.
# Run on Windows: .\build-windows.ps1
# Prereqs: Python 3.10+, Node 18+, PowerShell 5+

$ErrorActionPreference = "Stop"

$ROOT = Resolve-Path (Join-Path $PSScriptRoot "..")
Write-Host "==> 1. Build React frontend"
Set-Location (Join-Path $ROOT "frontend")
if (-not (Test-Path "node_modules")) { npm install }
npm run build
if ($LASTEXITCODE -ne 0) { throw "npm build failed" }

Write-Host "==> 2. Install Python build deps"
Set-Location (Join-Path $ROOT "backend")
python -m pip install --quiet -r requirements.txt
python -m pip install --quiet pyinstaller pywin32

Write-Host "==> 3. Run PyInstaller"
Set-Location (Join-Path $ROOT "build")
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
python -m PyInstaller accrual.spec --noconfirm
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$EXE = Join-Path $ROOT "build\dist\7t-Accrual-Updater.exe"
if (-not (Test-Path $EXE)) { throw "Exe not produced: $EXE" }

Write-Host ""
Write-Host "==> Done. EXE: $EXE"
Write-Host "   Run: & '$EXE'"
Write-Host "   Browser opens automatically at http://127.0.0.1:8100"
