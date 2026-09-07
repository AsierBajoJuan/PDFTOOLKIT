$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$pythonRoot = (& $python -c "import sys; print(sys.base_prefix)").Trim()
$version = (& $python -c "from version import __version__; print(__version__)").Trim()
$applicationName = "PDFToolKit-$version"

if (-not (Test-Path $python)) {
    throw "No existe el entorno virtual .venv. Créalo antes de compilar."
}

& $python -m pip install -r (Join-Path $projectRoot "requirements-build.txt")
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name $applicationName `
    --collect-all fitz `
    --hidden-import win32com.client `
    --hidden-import tkinter `
    --hidden-import _tkinter `
    --add-data "$(Join-Path $pythonRoot 'Lib\tkinter');tkinter" `
    --add-data "$(Join-Path $pythonRoot 'tcl');tcl" `
    --add-binary "$(Join-Path $pythonRoot 'DLLs\_tkinter.pyd');." `
    --add-binary "$(Join-Path $pythonRoot 'DLLs\tcl86t.dll');." `
    --add-binary "$(Join-Path $pythonRoot 'DLLs\tk86t.dll');." `
    --add-data "$(Join-Path $projectRoot 'img');img" `
    (Join-Path $projectRoot "PdfToolKit.py")

Write-Host "Ejecutable generado en dist\$applicationName\$applicationName.exe"
