"""Read-only PDF text extraction for Office Forge.

The bundled Python runtime intentionally has no PDF dependency.  This module
therefore tries pypdf / PyMuPDF / pdfplumber when the selected interpreter has
one of them, and otherwise falls back to a conservative, dependency-free
stream scanner.  The fallback extracts literal strings from uncompressed or
zlib-compressed content streams; it is deliberately labelled as a fallback,
not as a full PDF parser.
"""

from __future__ import annotations

import os
import re
import sys
import zlib
from typing import Any, Mapping

_VENDOR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
if _VENDOR not in sys.path:
    sys.path.insert(0, _VENDOR)

from office_common import EngineError, assert_input_file, scalar_text, truncate

_ALLOWED = (".pdf",)
_TEXT_TOKEN = re.compile(rb"\((?:\\.|[^()\\])*\)|<[0-9A-Fa-f\s]+>")
_STREAM = re.compile(rb"stream\r?\n(.*?)\r?\nendstream", re.DOTALL)
_PAGE = re.compile(rb"/Type\s*/Page\b")


def _literal(raw: bytes) -> str:
    """Decode a PDF literal string body (without the surrounding parentheses)."""
    out: list[str] = []
    index = 0
    while index < len(raw):
        byte = raw[index]
        if byte != 0x5C:  # backslash
            out.append(chr(byte))
            index += 1
            continue
        index += 1
        if index >= len(raw):
            break
        escape = raw[index]
        index += 1
        if escape in b"nrtbf":
            out.append({ord("n"): "\n", ord("r"): "\r", ord("t"): "\t", ord("b"): "\b", ord("f"): "\f"}[escape])
        elif escape in b"()\\":
            out.append(chr(escape))
        elif 48 <= escape <= 55:
            digits = bytes([escape])
            while index < len(raw) and len(digits) < 3 and 48 <= raw[index] <= 55:
                digits += bytes([raw[index]])
                index += 1
            out.append(chr(int(digits, 8)))
        elif escape in b"\r\n":
            if escape == 13 and index < len(raw) and raw[index] == 10:
                index += 1
        else:
            out.append(chr(escape))
    return "".join(out)


def _content_text(content: bytes) -> str:
    parts: list[str] = []
    for match in _TEXT_TOKEN.finditer(content):
        token = match.group(0)
        if token.startswith(b"("):
            value = _literal(token[1:-1])
        else:
            hex_digits = re.sub(rb"\s+", b"", token[1:-1])
            if len(hex_digits) % 2:
                hex_digits += b"0"
            try:
                value = bytes.fromhex(hex_digits.decode("ascii")).decode("latin-1")
            except Exception:
                continue
        value = value.replace("\x00", "").strip()
        if value:
            parts.append(value)
    return " ".join(parts)


def _fallback_pages(raw: bytes) -> list[str]:
    pages: list[str] = []
    for match in _STREAM.finditer(raw):
        content = match.group(1)
        try:
            content = zlib.decompress(content)
        except Exception:
            pass
        text = _content_text(content)
        if text:
            pages.append(text)
    if not pages:
        text = _content_text(raw)
        if text:
            pages.append(text)
    return pages


def _pypdf_pages(path: str, password: Any = None) -> list[str] | None:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        return None
    try:
        reader = PdfReader(path)
        if reader.is_encrypted:
            if not password:
                raise EngineError("PDF_PASSWORD_REQUIRED", "this PDF is encrypted and requires a password", {"path": path})
            try:
                result = reader.decrypt(str(password))
            except Exception as error:
                raise EngineError("PDF_PASSWORD_INVALID", "the PDF password is not valid", {"detail": f"{type(error).__name__}: {error}"}) from error
            if result == 0:
                raise EngineError("PDF_PASSWORD_INVALID", "the PDF password is not valid", {"path": path})
        return [page.extract_text() or "" for page in reader.pages]
    except EngineError:
        raise
    except Exception:
        return None


def _fitz_pages(path: str) -> list[str] | None:
    try:
        import fitz  # type: ignore
    except Exception:
        return None
    try:
        document = fitz.open(path)
        return [page.get_text() or "" for page in document]
    except Exception:
        return None


def _pdfplumber_pages(path: str) -> list[str] | None:
    try:
        import pdfplumber  # type: ignore
    except Exception:
        return None
    try:
        with pdfplumber.open(path) as document:
            return [page.extract_text() or "" for page in document.pages]
    except Exception:
        return None


def read_pdf(path: str, args: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    info = assert_input_file(path, _ALLOWED, int(limits["maxInputBytes"]))
    with open(path, "rb") as handle:
        raw = handle.read()
    if b"%PDF-" not in raw[:1024]:
        raise EngineError(
            "INVALID_PDF",
            "the file is not a valid PDF",
            {"path": path, "bytes": len(raw)},
        )
    pages = _pypdf_pages(path, args.get("password"))
    engine = "pypdf"
    if pages is None:
        pages = _fitz_pages(path)
        engine = "fitz"
    if pages is None:
        pages = _pdfplumber_pages(path)
        engine = "pdfplumber"
    if pages is None:
        pages = _fallback_pages(raw)
        engine = "fallback"
    if not pages:
        pages = []
    page_count = len(pages)
    if page_count == 0:
        page_count = max(1, len(_PAGE.findall(raw)))
    selected_pages = args.get("pages")
    if isinstance(selected_pages, list) and selected_pages:
        indices = []
        for value in selected_pages:
            try:
                number = int(value)
            except Exception:
                continue
            if 1 <= number <= len(pages):
                indices.append(number - 1)
        pages = [pages[index] for index in indices] if indices else pages
    max_pages = int(args.get("maxPages") or 50)
    truncated_pages = len(pages) > max_pages
    pages = pages[:max_pages]
    max_chars = int(args.get("maxChars") or limits.get("maxTextChars", 200_000))
    page_rows = []
    total = 0
    truncated_text = False
    for index, text in enumerate(pages, start=1):
        text = scalar_text(text)
        total += len(text)
        text, cut = truncate(text, max_chars)
        truncated_text = truncated_text or cut
        page_rows.append({"page": index, "text": text})
    text, cut = truncate("\n\n".join(row["text"] for row in page_rows), max_chars)
    truncated_text = truncated_text or cut
    fmt = str(args.get("format") or "json").lower()
    result: dict[str, Any] = {
        "type": "pdf",
        "path": path,
        "bytes": info.st_size,
        "pageCount": page_count,
        "engine": engine,
        "pages": page_rows,
        "text": text,
        "truncated": truncated_pages or truncated_text,
    }
    if fmt == "markdown":
        lines = ["# PDF 文本", ""]
        for row in page_rows:
            lines.extend([f"## 第 {row['page']} 页", "", row["text"], ""])
        result["markdown"] = "\n".join(lines)
    return result
