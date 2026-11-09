"""Pull text out of attachments: PDF, Word, Excel, text, HTML, mail.

Deliberately without heavy dependencies. Word and Excel are ZIP files holding
XML, which we read directly; PDF goes through pypdf. Everything has size limits
so a doctored file cannot bring anything down.
"""

from __future__ import annotations

import io
import re
import zipfile
from email import policy
from email.parser import BytesParser
from xml.etree import ElementTree as ET

from ggpost.text import html_to_text, normalize_whitespace

MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024
MAX_UNZIPPED_BYTES = 60 * 1024 * 1024
MAX_PDF_PAGES = 60
MAX_XLSX_CELLS = 20000

TEXT_EXT = {".txt", ".csv", ".tsv", ".md", ".json", ".xml", ".log", ".ics", ".vcf"}


class ExtractError(ValueError):
    pass


def _ext(name: str) -> str:
    m = re.search(r"(\.[A-Za-z0-9]{1,5})$", name or "")
    return m.group(1).lower() if m else ""


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _zip_member(zf: zipfile.ZipFile, name: str) -> bytes:
    info = zf.getinfo(name)
    if info.file_size > MAX_UNZIPPED_BYTES:
        raise ExtractError("Unpacked, the file is too large.")
    return zf.read(name)


def _check_zip(zf: zipfile.ZipFile) -> None:
    total = sum(i.file_size for i in zf.infolist())
    if total > MAX_UNZIPPED_BYTES * 2:
        raise ExtractError("Unpacked, the file is too large.")


def _xml_text(data: bytes, para_tag: str, text_tag: str) -> str:
    root = ET.fromstring(data)
    out: list[str] = []
    for para in root.iter(para_tag):
        parts = []
        for node in para.iter():
            if node.tag == text_tag and node.text:
                parts.append(node.text)
            elif node.tag.endswith("}tab"):
                parts.append("\t")
            elif node.tag.endswith("}br"):
                parts.append("\n")
        out.append("".join(parts))
    return "\n".join(out)


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def docx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        _check_zip(zf)
        return _xml_text(_zip_member(zf, "word/document.xml"), W + "p", W + "t")


def _col_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref or "A")
    n = 0
    for ch in letters.group(0) if letters else "A":
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def xlsx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        _check_zip(zf)
        names = zf.namelist()
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(_zip_member(zf, "xl/sharedStrings.xml"))
            for si in root.iter(S + "si"):
                shared.append("".join(t.text or "" for t in si.iter(S + "t")))
        sheet_names: list[str] = []
        if "xl/workbook.xml" in names:
            wb = ET.fromstring(_zip_member(zf, "xl/workbook.xml"))
            sheet_names = [s.get("name", "") for s in wb.iter(S + "sheet")]
        sheets = sorted(
            (n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
            key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1)),
        )
        out: list[str] = []
        cells = 0
        for idx, sheet in enumerate(sheets):
            title = sheet_names[idx] if idx < len(sheet_names) else f"Blatt {idx + 1}"
            out.append(f"## {title}")
            root = ET.fromstring(_zip_member(zf, sheet))
            for row in root.iter(S + "row"):
                values: dict[int, str] = {}
                for c in row.iter(S + "c"):
                    cells += 1
                    if cells > MAX_XLSX_CELLS:
                        out.append("[… weitere Zellen ausgelassen]")
                        return "\n".join(out)
                    t = c.get("t")
                    v = c.find(S + "v")
                    if t == "s" and v is not None and v.text and v.text.isdigit():
                        i = int(v.text)
                        text = shared[i] if i < len(shared) else ""
                    elif t == "inlineStr":
                        text = "".join(x.text or "" for x in c.iter(S + "t"))
                    else:
                        text = v.text if v is not None and v.text else ""
                    if text:
                        values[_col_index(c.get("r", ""))] = text
                if values:
                    width = max(values) + 1
                    out.append(" | ".join(values.get(i, "") for i in range(width)))
        return "\n".join(out)


def pdf_text(data: bytes) -> str:
    try:
        from pypdf import PdfReader  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover
        raise ExtractError("PDF support is missing (package pypdf).") from exc
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception as exc:
                raise ExtractError("The PDF is password protected.") from exc
        pages = []
        for i, page in enumerate(reader.pages):
            if i >= MAX_PDF_PAGES:
                pages.append(f"[… {len(reader.pages) - MAX_PDF_PAGES} further pages omitted]")
                break
            pages.append(page.extract_text() or "")
    except ExtractError:
        raise
    except Exception as exc:
        raise ExtractError(f"The PDF could not be read ({exc}).") from exc
    text = "\n\n".join(pages)
    if not text.strip():
        raise ExtractError("The PDF holds no readable text (probably a scan).")
    return text


def eml_text(data: bytes) -> str:
    from ggpost.backends.imap import extract_body  # noqa: PLC0415

    msg = BytesParser(policy=policy.default).parsebytes(data)
    head = "\n".join(f"{h}: {msg.get(h, '')}" for h in ("From", "To", "Date", "Subject") if msg.get(h))
    return head + "\n\n" + extract_body(msg)


def extract_text(name: str, data: bytes) -> str:
    if len(data) > MAX_ATTACHMENT_BYTES:
        raise ExtractError("The attachment is too large to read (over 20 MB).")
    ext = _ext(name)
    try:
        if ext == ".pdf" or data[:5] == b"%PDF-":
            text = pdf_text(data)
        elif ext in (".docx", ".docm"):
            text = docx_text(data)
        elif ext in (".xlsx", ".xlsm"):
            text = xlsx_text(data)
        elif ext in (".html", ".htm"):
            text = html_to_text(_decode(data))
        elif ext == ".eml":
            text = eml_text(data)
        elif ext in TEXT_EXT:
            text = _decode(data)
        elif ext in (".doc", ".xls", ".ppt"):
            raise ExtractError("post. cannot read the old Office formats (.doc, .xls).")
        else:
            raise ExtractError(f"post. cannot read file type {ext or 'unknown'} as text.")
    except (zipfile.BadZipFile, ET.ParseError, KeyError) as exc:
        raise ExtractError(f"The file is damaged or not a valid {ext} document.") from exc
    return normalize_whitespace(text)
