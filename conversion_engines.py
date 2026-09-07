"""Detección y uso de motores externos para convertir documentos a PDF."""

from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
from typing import Literal
import winreg


ConversionKind = Literal["word", "excel", "powerpoint", "html"]


def find_libreoffice() -> str | None:
    """Devuelve la ruta de soffice si LibreOffice está disponible."""
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def find_microsoft_office() -> bool:
    """Comprueba Office mediante sus identificadores COM de Windows."""
    for prog_id in ("Word.Application", "Excel.Application", "PowerPoint.Application"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id):
                return True
        except OSError:
            continue
    return False


def available_engines() -> list[str]:
    """Devuelve los motores detectados: ``office`` y/o ``libreoffice``."""
    engines: list[str] = []
    if find_microsoft_office():
        engines.append("office")
    if find_libreoffice():
        engines.append("libreoffice")
    return engines


def convert_to_pdf(source: str | Path, destination: str | Path, kind: ConversionKind, engine: str) -> None:
    """Convierte un documento usando el motor elegido."""
    source_path = Path(source)
    destination_path = Path(destination)
    if not source_path.is_file():
        raise FileNotFoundError(f"No existe el archivo: {source_path}")
    if source_path.resolve() == destination_path.resolve():
        raise ValueError("El archivo de salida debe ser distinto del archivo de entrada.")
    destination_path.parent.mkdir(parents=True, exist_ok=True)

    if engine == "office":
        try:
            _convert_with_office(source_path, destination_path, kind)
        except RuntimeError:
            raise
        except Exception as error:
            raise RuntimeError(
                "Microsoft Office está instalado, pero no se pudo iniciar para convertir el archivo. "
                "Puedes volver a intentarlo o elegir LibreOffice."
            ) from error
    elif engine == "libreoffice":
        _convert_with_libreoffice(source_path, destination_path)
    else:
        raise ValueError(f"Motor de conversión desconocido: {engine}")


def _convert_with_libreoffice(source: Path, destination: Path) -> None:
    executable = find_libreoffice()
    if not executable:
        raise RuntimeError("LibreOffice no está instalado.")
    with TemporaryDirectory(prefix="pdftoolkit-convert-") as temporary_directory:
        profile_directory = Path(temporary_directory) / "profile"
        profile_directory.mkdir()
        subprocess.run(
            [
                executable,
                "--headless",
                f"-env:UserInstallation={profile_directory.as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                temporary_directory,
                str(source),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        generated = Path(temporary_directory) / f"{source.stem}.pdf"
        if not generated.is_file():
            raise RuntimeError("LibreOffice no generó el PDF esperado.")
        shutil.copyfile(generated, destination)


def _convert_with_office(source: Path, destination: Path, kind: ConversionKind) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Para usar Microsoft Office instala la dependencia opcional pywin32."
        ) from error

    pythoncom.CoInitialize()
    application = None
    document = None
    try:
        if kind in ("word", "html"):
            application = win32com.client.DispatchEx("Word.Application")
            application.Visible = False
            document = application.Documents.Open(str(source.resolve()))
            document.ExportAsFixedFormat(str(destination.resolve()), 17)
        elif kind == "excel":
            application = win32com.client.DispatchEx("Excel.Application")
            application.Visible = False
            document = application.Workbooks.Open(str(source.resolve()))
            document.ExportAsFixedFormat(0, str(destination.resolve()))
        elif kind == "powerpoint":
            application = win32com.client.DispatchEx("PowerPoint.Application")
            document = application.Presentations.Open(str(source.resolve()), WithWindow=False)
            document.ExportAsFixedFormat(str(destination.resolve()), 2)
        else:
            raise ValueError(f"Tipo de documento desconocido: {kind}")
    finally:
        if document is not None:
            document.Close(False)
        if application is not None:
            application.Quit()
        pythoncom.CoUninitialize()
