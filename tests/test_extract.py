from __future__ import annotations

import io
import zipfile

import pytest

from ggpost.extract import ExtractError, extract_text


def make_pdf(text: str) -> bytes:
    pytest.importorskip("reportlab")
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    y = 800
    for line in text.splitlines():
        c.drawString(50, y, line)
        y -= 20
    c.save()
    return buf.getvalue()


def make_docx(paragraphs: list[str]) -> bytes:
    docx = pytest.importorskip("docx")
    d = docx.Document()
    for p in paragraphs:
        d.add_paragraph(p)
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "Menge"
    t.cell(0, 1).text = "40"
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def make_xlsx() -> bytes:
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bestellung"
    ws.append(["Artikel", "Menge", "Preis"])
    ws.append(["Gitterbox", 40, 89.5])
    ws["E5"] = "Notiz"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_pdf():
    text = extract_text("bestellung.pdf", make_pdf("Bestellung 4711\nLiefertermin Freitag"))
    assert "Bestellung 4711" in text and "Freitag" in text


def test_docx():
    text = extract_text("angebot.docx", make_docx(["Sehr geehrte Damen und Herren,", "anbei unser Angebot."]))
    assert "anbei unser Angebot." in text and "Menge" in text


def test_xlsx():
    text = extract_text("liste.xlsx", make_xlsx())
    assert "## Bestellung" in text
    assert "Gitterbox | 40 | 89.5" in text
    assert "Notiz" in text


def test_text_encodings():
    assert extract_text("a.csv", "Größe;Menge\n".encode("cp1252")).startswith("Größe")
    assert extract_text("a.txt", "Grüße".encode("utf-8")) == "Grüße"
    assert extract_text("a.html", b"<p>Hallo <b>Welt</b></p>") == "Hallo Welt"


def test_unsupported_and_broken():
    with pytest.raises(ExtractError, match="old Office formats"):
        extract_text("alt.doc", b"x")
    with pytest.raises(ExtractError, match="cannot read file type"):
        extract_text("bild.png", b"\x89PNG")
    with pytest.raises(ExtractError, match="damaged"):
        extract_text("kaputt.docx", b"kein zip")
    with pytest.raises(ExtractError):
        extract_text("leer.pdf", b"%PDF-1.4 kaputt")


def test_zip_bomb_guard():
    # Write in blocks: the guard reads the size recorded in the zip, so the test
    # does not need the 130 MB in memory at once (which would raise MemoryError).
    buf = io.BytesIO()
    chunk = b"0" * (1024 * 1024)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        with zf.open("word/document.xml", "w") as fh:
            for _ in range(130):
                fh.write(chunk)
    with pytest.raises(ExtractError, match="too large"):
        extract_text("bombe.docx", buf.getvalue())
