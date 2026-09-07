from pathlib import Path

import fitz
from PyPDF2 import PdfReader, PdfWriter

from pdf_operations import (
    add_watermark,
    compare_pdfs,
    compress_pdf,
    edit_pdf_metadata,
    jpg_to_pdf,
    merge_pdfs,
    number_pdf,
    pdf_to_jpg,
    protect_pdf,
    repair_pdf,
    reorder_pdf,
    rotate_pdf,
    split_pdf,
    unlock_pdf,
)


def create_pdf(path: Path, texts: list[str]) -> None:
    document = fitz.open()
    for text in texts:
        page = document.new_page(width=400, height=400)
        page.insert_text((40, 80), text, fontsize=16)
    document.save(path)
    document.close()


def test_merge_split_rotate_and_reorder(tmp_path: Path) -> None:
    first = tmp_path / "first.pdf"
    second = tmp_path / "second.pdf"
    create_pdf(first, ["uno"])
    create_pdf(second, ["dos"])

    merged = tmp_path / "merged.pdf"
    merge_pdfs([first, second], merged)
    assert len(PdfReader(str(merged)).pages) == 2

    pages = split_pdf(merged, tmp_path / "pages")
    assert len(pages) == 2

    rotated = tmp_path / "rotated.pdf"
    rotate_pdf(merged, rotated, 90)
    assert PdfReader(str(rotated)).pages[0].rotation == 90

    reordered = tmp_path / "reordered.pdf"
    reorder_pdf(merged, reordered, [2, 1])
    assert len(PdfReader(str(reordered)).pages) == 2


def test_images_compression_watermark_and_numbering(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    create_pdf(source, ["contenido"])

    jpgs = pdf_to_jpg(source, tmp_path / "images")
    assert len(jpgs) == 1

    from_images = tmp_path / "from-images.pdf"
    jpg_to_pdf(jpgs, from_images)
    assert len(PdfReader(str(from_images)).pages) == 1

    for function, name in (
        (compress_pdf, "compressed.pdf"),
        (lambda src, dst: add_watermark(src, dst, "BORRADOR"), "watermarked.pdf"),
        (number_pdf, "numbered.pdf"),
        (repair_pdf, "repaired.pdf"),
    ):
        destination = tmp_path / name
        function(source, destination)
        assert destination.exists() and destination.stat().st_size > 0


def test_passwords_metadata_and_comparison(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    other = tmp_path / "other.pdf"
    create_pdf(source, ["texto original"])
    create_pdf(other, ["texto cambiado"])

    edited = tmp_path / "edited.pdf"
    edit_pdf_metadata(source, edited, title="Prueba", author="PDFToolKit")
    with fitz.open(edited) as document:
        assert document.metadata["title"] == "Prueba"

    protected = tmp_path / "protected.pdf"
    protect_pdf(source, protected, "clave")
    assert PdfReader(str(protected)).is_encrypted

    unlocked = tmp_path / "unlocked.pdf"
    unlock_pdf(protected, unlocked, "clave")
    assert not PdfReader(str(unlocked)).is_encrypted

    report = tmp_path / "comparison.txt"
    compare_pdfs(source, other, report)
    report_text = report.read_text(encoding="utf-8")
    assert "texto original" in report_text
    assert "texto cambiado" in report_text
