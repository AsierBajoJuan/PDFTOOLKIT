"""Interfaz gráfica de PDFToolKit."""

from __future__ import annotations

import os
import logging
from pathlib import Path
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.simpledialog import askstring
import sys
from typing import Any, Callable

from PIL import Image, ImageOps, ImageTk

from pdf_operations import (
    compress_pdf,
    compare_pdfs,
    convert_pdf_to_powerpoint,
    convert_pdf_to_word,
    add_watermark,
    jpg_to_pdf,
    edit_pdf_metadata,
    merge_pdfs,
    number_pdf,
    ocr_pdf,
    pdf_to_excel,
    pdf_to_jpg,
    protect_pdf,
    repair_pdf,
    reorder_pdf,
    rotate_pdf,
    split_pdf,
    unlock_pdf,
)
from conversion_engines import available_engines, convert_to_pdf
from version import __version__


BASE_DIR = Path(__file__).resolve().parent
IMAGE_DIR = BASE_DIR / "img"
logger = logging.getLogger(__name__)

LIGHT = {
    "background": "#f4f7fb",
    "surface": "#ffffff",
    "text": "#172033",
    "muted": "#667085",
    "accent": "#2563eb",
    "accent_hover": "#1d4ed8",
    "border": "#dce3ee",
}
DARK = {
    "background": "#111827",
    "surface": "#1f2937",
    "text": "#f3f4f6",
    "muted": "#aab4c3",
    "accent": "#60a5fa",
    "accent_hover": "#93c5fd",
    "border": "#374151",
}


def _configure_tcl_paths() -> None:
    """Corrige la ruta Tcl/Tk en instalaciones de Python con layout no estándar."""
    roots = []
    if getattr(sys, "_MEIPASS", None):
        roots.append(Path(sys._MEIPASS))
    roots.append(Path(sys.base_prefix))
    for root in roots:
        tcl_root = root / "tcl"
        tcl_library = tcl_root / "tcl8.6"
        tk_library = tcl_root / "tk8.6"
        if (tcl_library / "init.tcl").exists() and "TCL_LIBRARY" not in os.environ:
            os.environ["TCL_LIBRARY"] = str(tcl_library)
        if (tk_library / "tk.tcl").exists() and "TK_LIBRARY" not in os.environ:
            os.environ["TK_LIBRARY"] = str(tk_library)
        if os.environ.get("TCL_LIBRARY") and os.environ.get("TK_LIBRARY"):
            break


class PdfToolKitApp:
    """Ventana principal y coordinación de acciones de usuario."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"PDFToolKit v{__version__}")
        self.root.geometry("1120x760")
        self.root.minsize(900, 650)
        self.is_day_mode = True
        self.colors = LIGHT
        self.buttons: list[ttk.Button] = []
        self.card_frames: list[ttk.Frame] = []
        self.images: list[ImageTk.PhotoImage] = []

        self._configure_styles()
        self.status_var = tk.StringVar(value="Listo para trabajar con tus documentos")
        self.frame = ttk.Frame(root, style="App.TFrame", padding=(28, 22))
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.images = [self._resize_image(IMAGE_DIR / f"image{i}.png", (52, 52)) for i in range(1, 26)]
        self.day_icon = self._resize_image(IMAGE_DIR / "day_icon.png", (22, 22))
        self.night_icon = self._resize_image(IMAGE_DIR / "night_icon.png", (22, 22))

        self._build_header()
        self._build_menu()
        self._build_footer()

        icon = ImageTk.PhotoImage(Image.open(IMAGE_DIR / "icono.png"))
        self.root.iconphoto(True, icon)
        self.root._icon_reference = icon

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background=self.colors["background"])
        style.configure("Header.TLabel", background=self.colors["background"], foreground=self.colors["text"])
        style.configure("Subtitle.TLabel", background=self.colors["background"], foreground=self.colors["muted"])
        style.configure("Status.TLabel", background=self.colors["surface"], foreground=self.colors["muted"], padding=(12, 8))
        style.configure(
            "Tool.TButton",
            background=self.colors["surface"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            padding=(10, 12),
            font=("Segoe UI", 10),
            relief="flat",
        )
        style.map(
            "Tool.TButton",
            background=[("active", self.colors["accent"]), ("disabled", self.colors["border"])],
            foreground=[("active", "white"), ("disabled", self.colors["muted"])],
        )
        style.configure(
            "Mode.TButton",
            background=self.colors["surface"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            padding=(10, 6),
            font=("Segoe UI", 9),
        )

    def _build_header(self) -> None:
        header = ttk.Frame(self.frame, style="App.TFrame")
        header.pack(fill=tk.X, pady=(0, 22))
        title_block = ttk.Frame(header, style="App.TFrame")
        title_block.pack(side=tk.LEFT)
        ttk.Label(title_block, text="PDFToolKit", style="Header.TLabel", font=("Segoe UI", 25, "bold")).pack(anchor="w")
        ttk.Label(
            title_block,
            text="Tus herramientas PDF, rápidas y en local",
            style="Subtitle.TLabel",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 0))
        self.toggle_button = ttk.Button(
            header,
            image=self.night_icon,
            command=self.toggle_mode,
            text=" Modo noche",
            compound="left",
            style="Mode.TButton",
        )
        self.toggle_button.pack(side=tk.RIGHT, anchor="n")

    def _build_footer(self) -> None:
        footer = ttk.Frame(self.frame, style="App.TFrame")
        footer.pack(fill=tk.X, pady=(20, 0))
        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(footer, text="AsierBajo", style="Subtitle.TLabel", font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=(12, 0))

    @staticmethod
    def _resize_image(image_path: Path, size: tuple[int, int]) -> ImageTk.PhotoImage:
        image = Image.open(image_path).convert("RGBA")
        pixels = []
        for pixel in image.getdata():
            if all(channel > 200 for channel in pixel[:3]):
                pixels.append((*pixel[:3], 0))
            else:
                pixels.append(pixel)
        image.putdata(pixels)
        image = ImageOps.contain(image, size, Image.Resampling.LANCZOS)
        canvas = Image.new("RGBA", size, (255, 255, 255, 0))
        canvas.alpha_composite(image, ((size[0] - image.width) // 2, (size[1] - image.height) // 2))
        return ImageTk.PhotoImage(canvas)

    def _build_menu(self) -> None:
        options = [
            ("Convertir PDF a Word", self.convert_word),
            ("Unir PDF", self.merge),
            ("Dividir PDF", self.split),
            ("Comprimir PDF", self.compress),
            ("PDF a PowerPoint", self.convert_powerpoint),
            ("PDF a Excel", self.pdf_to_excel),
            ("Word a PDF", self.word_to_pdf),
            ("PowerPoint a PDF", self.powerpoint_to_pdf),
            ("Excel a PDF", self.excel_to_pdf),
            ("Editar PDF", self.edit_metadata),
            ("PDF a JPG", self.pdf_to_jpg),
            ("JPG a PDF", self.jpg_to_pdf),
            ("Firmar PDF", self.not_implemented),
            ("Marca de agua", self.watermark),
            ("Rotar PDF", self.rotate),
            ("Html a PDF", self.html_to_pdf),
            ("Desbloquear PDF", self.unlock),
            ("Proteger PDF", self.protect),
            ("Ordenar PDF", self.reorder),
            ("PDF a PDF/a", self.not_implemented),
            ("Reparar PDF", self.repair),
            ("Enumerar páginas", self.number_pages),
            ("Escanear a PDF", self.not_implemented),
            ("OCR PDF", self.ocr),
            ("Comparar PDF", self.compare),
        ]

        self.menu_frame = ttk.Frame(self.frame, style="App.TFrame")
        self.menu_frame.pack(fill=tk.BOTH, expand=True)
        for column in range(5):
            self.menu_frame.columnconfigure(column, weight=1)
        for index, (label, command) in enumerate(options):
            card = ttk.Frame(self.menu_frame, style="App.TFrame", padding=4)
            card.grid(row=index // 5, column=index % 5, padx=7, pady=7, sticky="nsew")
            self.card_frames.append(card)
            button = ttk.Button(
                card,
                text=label,
                image=self.images[index],
                compound="top",
                command=command,
                style="Tool.TButton",
            )
            button.pack(fill=tk.BOTH, expand=True)
            self.buttons.append(button)

    def toggle_mode(self) -> None:
        self.is_day_mode = not self.is_day_mode
        self.colors = LIGHT if self.is_day_mode else DARK
        self._configure_styles()
        self.toggle_button.config(
            image=self.night_icon if self.is_day_mode else self.day_icon,
            text=" Modo noche" if self.is_day_mode else " Modo día",
        )

    def _select_pdf(self, title: str = "Selecciona un PDF") -> str | None:
        path = filedialog.askopenfilename(title=title, filetypes=[("PDF files", "*.pdf")])
        return path or None

    def _save_file(self, extension: str, file_type: str) -> str | None:
        path = filedialog.asksaveasfilename(
            defaultextension=extension,
            filetypes=[(file_type, f"*{extension}")],
        )
        return path or None

    def _show_error(self, error: Exception) -> None:
        self.status_var.set("Se ha producido un error")
        messagebox.showerror("Error", str(error))

    def _run_in_background(
        self,
        operation: Callable[[], Any],
        on_success: Callable[[Any], None],
    ) -> None:
        """Ejecuta una operación pesada sin bloquear el bucle de Tkinter."""
        for button in self.buttons:
            button.config(state=tk.DISABLED)

        def worker() -> None:
            try:
                result = operation()
            except Exception as error:
                logger.exception("Error ejecutando una operación PDF")
                self.root.after(0, lambda error=error: self._finish_background(error=error))
            else:
                self.root.after(0, lambda result=result: self._finish_background(on_success, result))

        Thread(target=worker, daemon=True).start()

    def _finish_background(
        self,
        on_success: Callable[[Any], None] | None = None,
        result: Any = None,
        error: Exception | None = None,
    ) -> None:
        for button in self.buttons:
            button.config(state=tk.NORMAL)
        if error is not None:
            self._show_error(error)
        elif on_success is not None:
            self.status_var.set("Operación completada correctamente")
            on_success(result)

    def convert_word(self) -> None:
        self.status_var.set("Preparando conversión a Word...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".docx", "Word files")
        if not destination:
            return
        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Archivo guardado en:\n{destination}")

        self._run_in_background(
            lambda: convert_pdf_to_word(source, destination),
            completed,
        )

    def merge(self) -> None:
        self.status_var.set("Preparando unión de PDFs...")
        sources = filedialog.askopenfilenames(
            title="Selecciona al menos dos PDFs",
            filetypes=[("PDF files", "*.pdf")],
        )
        if len(sources) < 2:
            if sources:
                messagebox.showerror("Error", "Selecciona al menos dos archivos PDF.")
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return
        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(
            lambda: merge_pdfs(list(sources), destination),
            completed,
        )

    def split(self) -> None:
        self.status_var.set("Preparando división del PDF...")
        source = self._select_pdf()
        if not source:
            return
        output_directory = filedialog.askdirectory(title="Selecciona la carpeta de salida")
        if not output_directory:
            return
        def completed(outputs: list[Path]) -> None:
            messagebox.showinfo("Éxito", f"Se han creado {len(outputs)} archivos.")

        self._run_in_background(
            lambda: split_pdf(source, output_directory),
            completed,
        )

    def convert_powerpoint(self) -> None:
        self.status_var.set("Preparando conversión a PowerPoint...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".pptx", "PowerPoint files")
        if not destination:
            return
        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Presentación guardada en:\n{destination}")

        self._run_in_background(
            lambda: convert_pdf_to_powerpoint(source, destination),
            completed,
        )

    def compress(self) -> None:
        self.status_var.set("Preparando compresión del PDF...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF comprimido y guardado en:\n{destination}")

        self._run_in_background(
            lambda: compress_pdf(source, destination),
            completed,
        )

    def pdf_to_jpg(self) -> None:
        self.status_var.set("Preparando conversión a JPG...")
        source = self._select_pdf()
        if not source:
            return
        output_directory = filedialog.askdirectory(title="Selecciona la carpeta de salida")
        if not output_directory:
            return

        def completed(outputs: list[Path]) -> None:
            messagebox.showinfo("Éxito", f"Se han creado {len(outputs)} imágenes JPG.")

        self._run_in_background(lambda: pdf_to_jpg(source, output_directory), completed)

    def pdf_to_excel(self) -> None:
        self.status_var.set("Preparando extracción a Excel...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".xlsx", "Excel files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Excel guardado en:\n{destination}")

        self._run_in_background(lambda: pdf_to_excel(source, destination), completed)

    def jpg_to_pdf(self) -> None:
        self.status_var.set("Preparando conversión a PDF...")
        sources = filedialog.askopenfilenames(
            title="Selecciona imágenes",
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png")],
        )
        if not sources:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(lambda: jpg_to_pdf(list(sources), destination), completed)

    def ocr(self) -> None:
        self.status_var.set("Preparando OCR...")
        source = self._select_pdf()
        if not source:
            return
        language = askstring(
            "OCR PDF",
            "Idiomas instalados en Tesseract (ejemplo: spa+eng):",
            parent=self.root,
            initialvalue="eng",
        )
        if not language:
            return
        destination = self._save_file(".txt", "Text files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Texto OCR guardado en:\n{destination}")

        self._run_in_background(lambda: ocr_pdf(source, destination, language), completed)

    def watermark(self) -> None:
        self.status_var.set("Preparando marca de agua...")
        source = self._select_pdf()
        if not source:
            return
        text = askstring("Marca de agua", "Texto de la marca de agua:", parent=self.root)
        if not text:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(lambda: add_watermark(source, destination, text), completed)

    def rotate(self) -> None:
        self.status_var.set("Preparando rotación...")
        source = self._select_pdf()
        if not source:
            return
        angle_text = askstring("Rotar PDF", "Ángulo (90, 180 o 270):", parent=self.root, initialvalue="90")
        if not angle_text:
            return
        try:
            angle = int(angle_text)
        except ValueError:
            messagebox.showerror("Error", "El ángulo debe ser 90, 180 o 270.")
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF rotado y guardado en:\n{destination}")

        self._run_in_background(lambda: rotate_pdf(source, destination, angle), completed)

    def _password_operation(self, title: str, operation: Callable[[str, str, str], None]) -> None:
        source = self._select_pdf()
        if not source:
            return
        password = askstring(title, "Contraseña:", show="*", parent=self.root)
        if not password:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(lambda: operation(source, destination, password), completed)

    def protect(self) -> None:
        self.status_var.set("Preparando protección del PDF...")
        self._password_operation("Proteger PDF", protect_pdf)

    def unlock(self) -> None:
        self.status_var.set("Preparando desbloqueo del PDF...")
        self._password_operation("Desbloquear PDF", unlock_pdf)

    def reorder(self) -> None:
        self.status_var.set("Preparando ordenación de páginas...")
        source = self._select_pdf()
        if not source:
            return
        order_text = askstring(
            "Ordenar PDF",
            "Introduce el orden separado por comas (ejemplo: 2,1,3):",
            parent=self.root,
        )
        if not order_text:
            return
        try:
            order = [int(value.strip()) for value in order_text.split(",")]
        except ValueError:
            messagebox.showerror("Error", "El orden debe contener números separados por comas.")
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(lambda: reorder_pdf(source, destination, order), completed)

    def number_pages(self) -> None:
        self.status_var.set("Preparando numeración de páginas...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF numerado y guardado en:\n{destination}")

        self._run_in_background(lambda: number_pdf(source, destination), completed)

    def _choose_conversion_engine(self, kind: str) -> str | None:
        engines = available_engines()
        if not engines:
            messagebox.showerror(
                "Motor no encontrado",
                "Esta función necesita Microsoft Office o LibreOffice.\n"
                "Instala uno de ellos y vuelve a intentarlo.",
            )
            self.status_var.set("Falta Microsoft Office o LibreOffice")
            return None
        if len(engines) == 1:
            return engines[0]
        use_office = messagebox.askyesno(
            "Elegir motor",
            "Se han detectado Microsoft Office y LibreOffice.\n\n"
            "¿Quieres usar Microsoft Office?\n"
            "Pulsa No para usar LibreOffice.",
            parent=self.root,
        )
        return "office" if use_office else "libreoffice"

    def _convert_document(
        self,
        kind: str,
        title: str,
        file_types: list[tuple[str, str]],
    ) -> None:
        self.status_var.set(f"Preparando {title}...")
        engine = self._choose_conversion_engine(kind)
        if not engine:
            return
        source = filedialog.askopenfilename(title=title, filetypes=file_types)
        if not source:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{destination}")

        self._run_in_background(
            lambda: convert_to_pdf(source, destination, kind, engine),
            completed,
        )

    def word_to_pdf(self) -> None:
        self._convert_document("word", "Convertir Word a PDF", [("Word", "*.doc *.docx")])

    def powerpoint_to_pdf(self) -> None:
        self._convert_document("powerpoint", "Convertir PowerPoint a PDF", [("PowerPoint", "*.ppt *.pptx")])

    def excel_to_pdf(self) -> None:
        self._convert_document("excel", "Convertir Excel a PDF", [("Excel", "*.xls *.xlsx")])

    def html_to_pdf(self) -> None:
        self._convert_document("html", "Convertir HTML a PDF", [("HTML", "*.html *.htm")])

    def edit_metadata(self) -> None:
        self.status_var.set("Preparando edición de metadatos...")
        source = self._select_pdf()
        if not source:
            return
        title = askstring("Editar PDF", "Título:", parent=self.root) or ""
        author = askstring("Editar PDF", "Autor:", parent=self.root) or ""
        subject = askstring("Editar PDF", "Asunto:", parent=self.root) or ""
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Metadatos guardados en:\n{destination}")

        self._run_in_background(
            lambda: edit_pdf_metadata(source, destination, title, author, subject),
            completed,
        )

    def repair(self) -> None:
        self.status_var.set("Preparando reparación del PDF...")
        source = self._select_pdf()
        if not source:
            return
        destination = self._save_file(".pdf", "PDF files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"PDF reparado y guardado en:\n{destination}")

        self._run_in_background(lambda: repair_pdf(source, destination), completed)

    def compare(self) -> None:
        self.status_var.set("Preparando comparación de PDFs...")
        sources = filedialog.askopenfilenames(
            title="Selecciona exactamente dos PDFs",
            filetypes=[("PDF files", "*.pdf")],
        )
        if len(sources) != 2:
            if sources:
                messagebox.showerror("Error", "Selecciona exactamente dos archivos PDF.")
            return
        destination = self._save_file(".txt", "Text files")
        if not destination:
            return

        def completed(_: Any) -> None:
            messagebox.showinfo("Éxito", f"Informe guardado en:\n{destination}")

        self._run_in_background(
            lambda: compare_pdfs(sources[0], sources[1], destination),
            completed,
        )

    def not_implemented(self) -> None:
        self.status_var.set("Funcionalidad pendiente de implementar")
        messagebox.showinfo("Pendiente", "Esta funcionalidad todavía no está implementada.")


def main() -> None:
    logging.basicConfig(
        filename=BASE_DIR / "pdftoolkit.log",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    logger.info("Iniciando PDFToolKit")
    _configure_tcl_paths()
    root = tk.Tk()
    PdfToolKitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
