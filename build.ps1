$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$pythonRoot = (& $python -c "import sys; print(sys.base_prefix)").Trim()

if (-not (Test-Path $python)) {
    throw "No existe el entorno virtual .venv. Créalo antes de compilar."
}

& $python -m pip install -r (Join-Path $projectRoot "requirements-build.txt")
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name PDFToolKit `
    --collect-all fitz `
    --hidden-import win32com.client `
    --hidden-import tkinter `
    --add-data "$(Join-Path $pythonRoot 'tcl');tcl" `
    --add-data "$(Join-Path $projectRoot 'img');img" `
    (Join-Path $projectRoot "PdfToolKit.py")

Write-Host "Ejecutable generado en dist\PDFToolKit\PDFToolKit.exe"
