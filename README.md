# PDFToolKit

![Version](https://img.shields.io/badge/version-1.0.0-2563eb.svg)
![Python](https://img.shields.io/badge/python-3.13-3776ab.svg)
![Platform](https://img.shields.io/badge/platform-Windows-0078d4.svg)
![License](https://img.shields.io/badge/license-MIT-22c55e.svg)

Aplicación de escritorio para trabajar con documentos PDF de forma local. PDFToolKit ofrece una interfaz gráfica sencilla para convertir, unir, dividir, comprimir, proteger y transformar documentos sin subir los archivos a servicios externos.

## Versión 1.0.0

Esta versión marca la primera versión estable del proyecto e incluye:

- Interfaz gráfica renovada con modo claro y oscuro.
- Procesamiento en segundo plano para no bloquear la ventana.
- Entorno compatible con Python 3.13.
- Tests automatizados para las operaciones principales.
- Generación de ejecutable Windows mediante PyInstaller.
- Detección automática de Microsoft Office y LibreOffice para conversiones.

## Funcionalidades

### PDF y documentos

- PDF a Word.
- Word a PDF.
- PDF a PowerPoint.
- PowerPoint a PDF.
- PDF a Excel, exportando el texto por páginas.
- Excel a PDF.
- HTML a PDF.
- Unir PDFs.
- Dividir PDFs.
- Comprimir PDFs.
- Reparar PDFs.
- Comparar el texto de dos PDFs.

### Páginas e imágenes

- Rotar páginas.
- Ordenar páginas.
- Enumerar páginas.
- PDF a JPG.
- JPG/PNG a PDF.
- Marca de agua.
- Edición de metadatos.

### Seguridad y OCR

- Proteger PDF con contraseña.
- Desbloquear PDF utilizando la contraseña válida.
- OCR mediante Tesseract.

## Requisitos

- Windows 10/11 de 64 bits.
- Python 3.13.x.
- Dependencias del proyecto de [requirements.txt](requirements.txt).
- Tesseract OCR para utilizar OCR.
- Microsoft Office o LibreOffice para convertir Word, Excel, PowerPoint y HTML a PDF.

La guía completa de instalación está en [docs/SETUP.md](docs/SETUP.md).

## Instalación desde código fuente

```powershell
git clone <URL_DEL_REPOSITORIO>
cd PDFTOOLKIT-main
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -r requirements-dev.txt
```

Para habilitar conversiones mediante Microsoft Office:

```powershell
python -m pip install -r requirements-office.txt
```

## Ejecución

```powershell
python PdfToolKit.py
```

## Motores de conversión

PDFToolKit detecta los motores instalados automáticamente:

- Si solo está disponible Microsoft Office, utiliza COM mediante `pywin32`.
- Si solo está disponible LibreOffice, utiliza `soffice` en modo headless.
- Si están disponibles ambos, la aplicación permite elegir.
- Si no está disponible ninguno, muestra una alerta explicativa.

## OCR

El OCR necesita tener instalado Tesseract y sus paquetes de idioma. Al ejecutar la herramienta se puede indicar el idioma, por ejemplo:

```text
eng
spa+eng
```

Si Tesseract no está instalado o no está en el `PATH`, la aplicación informa del problema sin modificar el PDF original.

## Tests

```powershell
python -m pytest -q
```

## Generar el ejecutable Windows

```powershell
.\build.ps1
```

El ejecutable se genera en:

```text
dist\PDFToolKit\PDFToolKit.exe
```

El proceso incluye los iconos, Tcl/Tk y las dependencias necesarias para ejecutar la interfaz fuera del entorno virtual.

## Limitaciones conocidas

Las siguientes funciones todavía requieren trabajo o herramientas especializadas:

- Firma digital PAdES con certificado real.
- Conversión PDF/A validada mediante un estándar externo.
- Escaneo directo desde hardware.
- Edición avanzada del contenido de una página.

Estas funciones no se presentan como disponibles hasta contar con una implementación verificable.

## Estructura principal

```text
PdfToolKit.py          Interfaz gráfica
pdf_operations.py      Operaciones PDF
conversion_engines.py  Microsoft Office y LibreOffice
tests/                 Tests automatizados
docs/SETUP.md          Guía de instalación
build.ps1              Generación del ejecutable Windows
```

## Licencia

PDFToolKit se distribuye bajo la licencia [MIT](LICENSE).

Copyright © 2024–2026 AsierBajo.
