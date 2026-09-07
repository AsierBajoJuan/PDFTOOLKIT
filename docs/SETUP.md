# Preparación del entorno

## Versión de Python

La versión objetivo para este proyecto es **Python 3.13.x de 64 bits en Windows**.
Es una versión estable y compatible con las dependencias fijadas actualmente.

## Instalación

Desde la carpeta raíz del proyecto:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -r requirements-dev.txt
```

## Comprobación básica

```powershell
python --version
python -m pip check
python -m py_compile PdfToolKit.py
```

La aplicación usa OCR mediante `pytesseract`. Para habilitar OCR es necesario
instalar también el programa Tesseract OCR y añadirlo al `PATH` de Windows.
Además, deben estar instalados los idiomas que se indiquen en la ventana OCR,
por ejemplo `eng` o `spa+eng`. Si Tesseract no está disponible, PDFToolKit
mostrará un error controlado y no modificará el PDF original.

## Conversión de documentos a PDF

Las conversiones de Word, Excel, PowerPoint y HTML comprueban automáticamente
los motores disponibles:

- Si solo está instalado Microsoft Office, se usa mediante COM.
- Si solo está instalado LibreOffice, se usa `soffice` en modo headless.
- Si están instalados los dos, la aplicación pregunta cuál utilizar.
- Si no hay ninguno, se muestra una alerta y no se inicia la conversión.

Para habilitar Microsoft Office mediante COM instala la dependencia opcional:

```powershell
python -m pip install -r requirements-office.txt
```

LibreOffice no necesita una dependencia Python adicional. Debe estar instalado
en una ruta estándar o disponible mediante `soffice` en el `PATH`.

## Ejecución

```powershell
python PdfToolKit.py
```

## Tests y ejecutable Windows

Para ejecutar los tests:

```powershell
python -m pytest -q
```

Para generar el ejecutable distribuible:

```powershell
.\build.ps1
```

El resultado queda en `dist\PDFToolKit-1.0.0\PDFToolKit-1.0.0.exe`. El nombre se
obtiene automáticamente desde `version.py`, por lo que cambiará al preparar
una nueva versión. El script incluye los
iconos y los archivos Tcl/Tk necesarios para que la interfaz funcione fuera del
entorno virtual.

## Estado actual del entorno

La instalación detectada está en `Python313` y devuelve Python 3.13.0. En este
equipo los alias `python` y `py` de Windows no apuntan correctamente a ella, por
lo que debe usarse el ejecutable instalado o corregirse el `PATH`.

PDFToolKit detecta automáticamente la carpeta `tcl` de esa instalación para
evitar el error de Tkinter `Can't find a usable init.tcl`.
