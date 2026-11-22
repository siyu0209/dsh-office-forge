"""File-level PDF operations for Office Forge.

The bundled Python runtime does not ship a PDF writer, so Office Forge carries
a pure-Python pypdf copy under ``lib/vendor``.  This keeps merge/split/rotate
available out of the box without installing anything into the host profile.
"""

from __future__ import annotations

import io
import os
import sys
from typing import Any, Mapping

_VENDOR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
if _VENDOR not in sys.path:
    sys.path.insert(0, _VENDOR)

try:
    from pypdf import PageObject, PdfReader, PdfWriter
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject
except Exception as error:  # pragma: no cover - import failure is environment-specific
    PageObject = None  # type: ignore[assignment]
    PdfReader = None  # type: ignore[assignment]
    PdfWriter = None  # type: ignore[assignment]
    DecodedStreamObject = None  # type: ignore[assignment]
    DictionaryObject = None  # type: ignore[assignment]
    NameObject = None  # type: ignore[assignment]
    _PYPDF_ERROR = f"{type(error).__name__}: {error}"
else:
    _PYPDF_ERROR = ""

from office_common import EngineError, assert_inside, atomic_save, safe_filename


def _require_pypdf() -> None:
    if PdfReader is None or PdfWriter is None:
        raise EngineError("BACKEND_UNAVAILABLE", "the bundled pypdf module is unavailable", {"detail": _PYPDF_ERROR})


def _checked(path: Any, label: str, workspace: str) -> str:
    if not path:
        raise EngineError("INVALID_ARGUMENT", f"{label} is required")
    raw = str(path)
    absolute = raw if os.path.isabs(raw) else os.path.join(workspace, raw)
    absolute = os.path.abspath(absolute)
    assert_inside(absolute, workspace, label)
    if not os.path.isfile(absolute):
        raise EngineError("INPUT_NOT_FOUND", f"{label} does not exist: {absolute}", {"path": absolute})
    return absolute


def _checked_output(path: Any, workspace: str, label: str = "output") -> str:
    if not path:
        raise EngineError("INVALID_ARGUMENT", f"{label} is required")
    raw = str(path)
    absolute = raw if os.path.isabs(raw) else os.path.join(workspace, raw)
    absolute = os.path.abspath(absolute)
    assert_inside(absolute, workspace, label)
    return absolute


def _reader(path: str, password: Any = None) -> Any:
    _require_pypdf()
    try:
        reader = PdfReader(path)
        if getattr(reader, "is_encrypted", False):
            if not password:
                raise EngineError("PDF_PASSWORD_REQUIRED", "PDF is password protected", {"path": path})
            try:
                result = reader.decrypt(str(password))
            except Exception as error:
                raise EngineError("PDF_PASSWORD_INVALID", "cannot decrypt PDF with the supplied password", {"detail": f"{type(error).__name__}: {error}"}) from error
            if not result:
                raise EngineError("PDF_PASSWORD_INVALID", "cannot decrypt PDF with the supplied password")
        return reader
    except EngineError:
        raise
    except Exception as error:
        raise EngineError("INVALID_PDF", f"cannot read PDF: {path}", {"detail": f"{type(error).__name__}: {error}"}) from error


def _page_numbers(value: Any, page_count: int) -> list[int]:
    if not isinstance(value, list) or not value:
        return list(range(page_count))
    result: list[int] = []
    for item in value:
        try:
            number = int(item)
        except Exception:
            continue
        if 1 <= number <= page_count and number - 1 not in result:
            result.append(number - 1)
    return result or list(range(page_count))


def merge_pdfs(paths: Any, output: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    if not isinstance(paths, list) or not paths:
        raise EngineError("INVALID_ARGUMENT", "inputs must be a non-empty array of PDF paths")
    if len(paths) > 50:
        raise EngineError("LIMIT_EXCEEDED", "cannot merge more than 50 PDF files", {"count": len(paths)})
    sources = [_checked(item, f"inputs[{index}]", workspace) for index, item in enumerate(paths)]
    target = _checked_output(output, workspace)
    writer = PdfWriter()
    total_pages = 0
    for source in sources:
        reader = _reader(source)
        for page in reader.pages:
            writer.add_page(page)
            total_pages += 1

    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)

    size, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "merge", "path": target, "bytes": size, "pages": total_pages, "inputs": len(sources)}


def split_pdf(path: Any, output_dir: Any, pages: Any, every_page: bool, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    directory = _checked_output(output_dir, workspace, "outputDir")
    os.makedirs(directory, exist_ok=True)
    reader = _reader(source)
    page_count = len(reader.pages)
    selected = _page_numbers(pages, page_count)
    grouped: list[list[int]] = [[index] for index in selected] if every_page or not isinstance(pages, list) else [selected]
    stem = os.path.splitext(os.path.basename(source))[0]
    outputs: list[dict[str, Any]] = []
    for group_index, group in enumerate(grouped, start=1):
        writer = PdfWriter()
        for page_index in group:
            writer.add_page(reader.pages[page_index])
        suffix = f"{group_index:03d}" if every_page or len(grouped) > 1 else "part"
        target = os.path.join(directory, f"{safe_filename(stem)}-{suffix}.pdf")

        def write(temp: str, current_writer=writer) -> None:
            with open(temp, "wb") as handle:
                current_writer.write(handle)

        size, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
        outputs.append({"path": target, "bytes": size, "pages": [index + 1 for index in group]})
    return {"type": "pdf", "operation": "split", "path": source, "pages": page_count, "outputs": outputs}


def rotate_pdf(path: Any, output: Any, pages: Any, angle: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    target = _checked_output(output, workspace)
    try:
        degrees = int(angle)
    except Exception as error:
        raise EngineError("INVALID_ARGUMENT", "angle must be an integer number of degrees") from error
    if degrees % 90 != 0:
        raise EngineError("INVALID_ARGUMENT", "angle must be a multiple of 90 degrees")
    reader = _reader(source)
    selected = set(_page_numbers(pages, len(reader.pages)))
    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        if index in selected:
            page.rotate(degrees)
        writer.add_page(page)

    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)

    size, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "rotate", "path": target, "bytes": size, "pages": len(reader.pages), "angle": degrees}


def watermark_pdf(path: Any, output: Any, watermark_path: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    overlay_path = _checked(watermark_path, "watermarkPath", workspace)
    target = _checked_output(output, workspace)
    reader = _reader(source)
    overlay = _reader(overlay_path)
    if not overlay.pages:
        raise EngineError("INVALID_PDF", "watermark PDF has no pages", {"path": overlay_path})
    watermark_page = overlay.pages[0]
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)

    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)

    size, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "watermark", "path": target, "bytes": size, "pages": len(reader.pages), "watermark": overlay_path}


def _pdf_text_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)").replace("\r", " ").replace("\n", " ")


def _page_number_text(template: str, page: int, total: int) -> str:
    return template.replace("{page}", str(page)).replace("{total}", str(total)).replace("{pages}", str(total))


def page_numbers_pdf(path: Any, output: Any, pages: Any, template: Any, position: Any, start: Any, font_size: Any, margin: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool, password: Any = None) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    target = _checked_output(output, workspace)
    reader = _reader(source, password)
    selected = set(_page_numbers(pages, len(reader.pages)))
    try:
        first = int(start)
    except Exception:
        first = 1
    try:
        size = float(font_size)
    except Exception:
        size = 10.0
    try:
        gap = float(margin)
    except Exception:
        gap = 24.0
    size = max(6.0, min(size, 36.0))
    gap = max(4.0, min(gap, 120.0))
    text_template = str(template or "{page} / {total}")
    place = str(position or "bottom-center").lower()
    writer = PdfWriter()
    font = DictionaryObject({
        NameObject("/Type"): NameObject("/Font"),
        NameObject("/Subtype"): NameObject("/Type1"),
        NameObject("/BaseFont"): NameObject("/Helvetica"),
    })
    font_ref = writer._add_object(font)
    numbered = 0
    for index, page in enumerate(reader.pages):
        if index in selected:
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            text = _page_number_text(text_template, first + index, len(reader.pages))
            escaped = _pdf_text_escape(text)
            text_width = len(text) * size * 0.5
            if place in ("bottom-right", "top-right"):
                x = max(gap, width - gap - text_width)
            elif place in ("bottom-left", "top-left"):
                x = gap
            else:
                x = max(gap, (width - text_width) / 2)
            y = gap if place.startswith("bottom") else max(gap, height - gap - size)
            overlay = PageObject.create_blank_page(width=width, height=height)
            overlay[NameObject("/Resources")] = DictionaryObject({
                NameObject("/Font"): DictionaryObject({NameObject("/F1"): font_ref}),
            })
            stream = DecodedStreamObject()
            stream.set_data(f"BT /F1 {size:g} Tf 1 0 0 1 {x:.2f} {y:.2f} Tm ({escaped}) Tj ET".encode("latin-1", "replace"))
            overlay[NameObject("/Contents")] = writer._add_object(stream)
            page.merge_page(overlay)
            numbered += 1
        writer.add_page(page)

    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)

    written, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "pageNumbers", "path": target, "bytes": written, "pages": len(reader.pages), "numbered": numbered, "position": place}


def _field_type(field: Any) -> str:
    ft = str(field.get("/FT") or "")
    if ft == "/Tx":
        return "text"
    if ft == "/Ch":
        return "choice"
    if ft == "/Btn":
        flags = int(field.get("/Ff") or 0)
        return "radio_group" if flags & (1 << 15) else "checkbox"
    if ft == "/Sig":
        return "signature"
    return ft or "unknown"


def list_form_fields(path: Any, password: Any, workspace: str) -> dict[str, Any]:
    source = _checked(path, "path", workspace)
    reader = _reader(source, password)
    try:
        raw_fields = reader.get_fields() or {}
    except Exception as error:
        raise EngineError("INVALID_PDF", "cannot read PDF form fields", {"detail": f"{type(error).__name__}: {error}"}) from error
    fields: list[dict[str, Any]] = []
    for name, field in raw_fields.items():
        entry: dict[str, Any] = {"name": str(name), "type": _field_type(field), "value": str(field.get("/V") or "")}
        options = field.get("/Opt")
        if options:
            entry["options"] = [str(option) for option in options]
        fields.append(entry)
    return {"type": "pdf", "operation": "formFields", "path": source, "fields": fields, "count": len(fields)}


def fill_form_pdf(path: Any, output: Any, values: Any, password: Any, flatten: bool, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    target = _checked_output(output, workspace)
    if not isinstance(values, Mapping):
        raise EngineError("INVALID_ARGUMENT", "fields must be an object of field name to value")
    reader = _reader(source, password)
    try:
        writer = PdfWriter(clone_from=reader)
        values_map = {str(key): value for key, value in values.items()}
        for page in writer.pages:
            writer.update_page_form_field_values(
                page,
                values_map,
                auto_regenerate=True,
                flatten=bool(flatten),
            )
    except Exception as error:
        raise EngineError("INVALID_ARGUMENT", "cannot fill PDF form fields", {"detail": f"{type(error).__name__}: {error}"}) from error

    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)

    written, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "fillForm", "path": target, "bytes": written, "fields": len(values), "flatten": bool(flatten)}


def extract_images_pdf(path: Any, output_dir: Any, pages: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool, password: Any = None) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    directory = _checked_output(output_dir, workspace, "outputDir")
    os.makedirs(directory, exist_ok=True)
    reader = _reader(source, password)
    selected = set(_page_numbers(pages, len(reader.pages)))
    stem = safe_filename(os.path.splitext(os.path.basename(source))[0]) or "pdf"
    outputs: list[dict[str, Any]] = []
    total = 0
    for page_index, page in enumerate(reader.pages):
        if page_index not in selected:
            continue
        try:
            images = list(page.images)
        except Exception:
            images = []
        for image_index, image in enumerate(images, start=1):
            if total >= 500:
                raise EngineError("LIMIT_EXCEEDED", "cannot extract more than 500 images", {"count": total})
            data = getattr(image, "data", b"") or b""
            if not data:
                continue
            try:
                from PIL import Image  # type: ignore
                with Image.open(io.BytesIO(data)) as bitmap:
                    buffer = io.BytesIO()
                    bitmap.save(buffer, format="PNG")
                    payload = buffer.getvalue()
            except Exception:
                payload = data
            target = os.path.join(directory, f"{stem}-p{page_index + 1:03d}-{image_index:03d}.png")

            def write(temp: str, current_payload=payload) -> None:
                with open(temp, "wb") as handle:
                    handle.write(current_payload)

            written, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
            outputs.append({"path": target, "bytes": written, "page": page_index + 1})
            total += 1
    return {"type": "pdf", "operation": "extractImages", "path": source, "images": outputs, "count": len(outputs)}


def ocr_pdf(path: Any, pages: Any, language: Any, dpi: Any, workspace: str, limits: Mapping[str, Any], password: Any = None) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    try:
        import fitz  # type: ignore
        import pytesseract  # type: ignore
        from PIL import Image  # type: ignore
    except Exception as error:
        raise EngineError("BACKEND_UNAVAILABLE", "OCR requires PyMuPDF (fitz) and pytesseract", {"detail": f"{type(error).__name__}: {error}"}) from error
    try:
        pytesseract.get_tesseract_version()
    except Exception as error:
        raise EngineError("BACKEND_UNAVAILABLE", "the tesseract executable is not available", {"detail": f"{type(error).__name__}: {error}"}) from error
    reader = _reader(source, password)
    selected = set(_page_numbers(pages, len(reader.pages)))
    document = fitz.open(source)
    try:
        texts: list[str] = []
        for index in sorted(selected):
            page = document[index]
            pix = page.get_pixmap(dpi=max(72, min(int(dpi or 200), 400)))
            bitmap = Image.open(io.BytesIO(pix.tobytes("png")))
            texts.append(pytesseract.image_to_string(bitmap, lang=str(language or "eng")))
    finally:
        document.close()
    return {"type": "pdf", "operation": "ocr", "path": source, "pages": [index + 1 for index in sorted(selected)], "text": texts, "engine": "pytesseract"}


def encrypt_pdf(path: Any, output: Any, user_password: Any, owner_password: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    target = _checked_output(output, workspace)
    reader = _reader(source)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(str(user_password or ""), str(owner_password or user_password or ""))
    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)
    written, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "encrypt", "path": target, "bytes": written, "pages": len(reader.pages)}


def decrypt_pdf(path: Any, output: Any, password: Any, workspace: str, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    _require_pypdf()
    source = _checked(path, "path", workspace)
    target = _checked_output(output, workspace)
    reader = _reader(source, password)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    def write(temp: str) -> None:
        with open(temp, "wb") as handle:
            writer.write(handle)
    written, _ = atomic_save(target, write, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pdf", "operation": "decrypt", "path": target, "bytes": written, "pages": len(reader.pages)}


def pdf_operation(payload: Mapping[str, Any], workspace: str, limits: Mapping[str, Any]) -> dict[str, Any]:
    operation = str(payload.get("operation") or "").lower()
    overwrite = bool(payload.get("overwrite", False))
    password = payload.get("password")
    if operation == "merge":
        return merge_pdfs(payload.get("inputs"), payload.get("output"), workspace, limits, overwrite)
    if operation == "split":
        return split_pdf(payload.get("path"), payload.get("outputDir"), payload.get("pages"), bool(payload.get("everyPage")), workspace, limits, overwrite)
    if operation == "rotate":
        return rotate_pdf(payload.get("path"), payload.get("output"), payload.get("pages"), payload.get("angle", 90), workspace, limits, overwrite)
    if operation == "watermark":
        return watermark_pdf(payload.get("path"), payload.get("output"), payload.get("watermarkPath"), workspace, limits, overwrite)
    if operation == "pagenumbers":
        return page_numbers_pdf(payload.get("path"), payload.get("output"), payload.get("pages"), payload.get("format"), payload.get("position"), payload.get("start", 1), payload.get("fontSize", 10), payload.get("margin", 24), workspace, limits, overwrite, password)
    if operation == "formfields":
        return list_form_fields(payload.get("path"), password, workspace)
    if operation == "fillform":
        return fill_form_pdf(payload.get("path"), payload.get("output"), payload.get("fields"), password, bool(payload.get("flatten")), workspace, limits, overwrite)
    if operation == "extractimages":
        return extract_images_pdf(payload.get("path"), payload.get("outputDir"), payload.get("pages"), workspace, limits, overwrite, password)
    if operation == "ocr":
        return ocr_pdf(payload.get("path"), payload.get("pages"), payload.get("language"), payload.get("dpi"), workspace, limits, password)
    if operation == "encrypt":
        return encrypt_pdf(payload.get("path"), payload.get("output"), payload.get("userPassword"), payload.get("ownerPassword"), workspace, limits, overwrite)
    if operation == "decrypt":
        return decrypt_pdf(payload.get("path"), payload.get("output"), password, workspace, limits, overwrite)
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported PDF operation: {operation or '(empty)'}", {"allowed": ["merge", "split", "rotate", "watermark", "pageNumbers", "formFields", "fillForm", "extractImages", "ocr", "encrypt", "decrypt"]})
