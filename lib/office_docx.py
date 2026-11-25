"""Word (.docx) reading, creation, editing and template expansion."""

from __future__ import annotations

import copy
import hashlib
import os
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping

from docx import Document
from docx.document import Document as _Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.table import Table, _Row
from docx.text.paragraph import Paragraph
from lxml import etree

from office_common import (
    EngineError,
    apply_core_properties,
    assert_inside,
    atomic_save,
    get_path,
    jsonable,
    md_escape,
    md_table,
    render_placeholders,
    require_dict,
    require_feature,
    require_list,
    scalar_text,
    truncate,
    word_counts,
)

_ALLOWED = (".docx",)
_HEADING_RE = re.compile(r"^(?:Heading|标题)\s*([1-9])$", re.IGNORECASE)
_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------


def _iter_block_items(parent: Any) -> Iterable[Any]:
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def _paragraph_text(paragraph: Paragraph) -> str:
    pieces: list[str] = []
    for node in paragraph._p.iter():
        if node.tag == qn("w:t"):
            pieces.append(node.text or "")
        elif node.tag == qn("w:tab"):
            pieces.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            pieces.append("\n")
    return "".join(pieces)


def _heading_level(paragraph: Paragraph) -> int | None:
    name = paragraph.style.name if paragraph.style is not None else ""
    match = _HEADING_RE.match(name or "")
    if match:
        return int(match.group(1))
    if name and name.lower() in ("title", "标题"):
        return 0
    return None


def _table_rows(table: Table) -> tuple[list[str], list[list[str]]]:
    rows: list[list[str]] = []
    for row in table.rows:
        rows.append([_cell_text(cell) for cell in row.cells])
    if not rows:
        return [], []
    width = max(len(row) for row in rows)
    normalised = [row + [""] * (width - len(row)) for row in rows]
    return normalised[0], normalised[1:]


def _cell_text(cell: Any) -> str:
    parts: list[str] = []
    for paragraph in cell.paragraphs:
        text = _paragraph_text(paragraph)
        if text:
            parts.append(text)
    return "\n".join(parts)


def _build_blocks(doc: Document) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    index = 0
    for item in _iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = _paragraph_text(item)
            level = _heading_level(item)
            style = item.style.name if item.style is not None else ""
            if text.strip() or level is not None:
                blocks.append(
                    {
                        "index": index,
                        "kind": "heading" if level is not None else "paragraph",
                        "level": level,
                        "style": style,
                        "text": text,
                    }
                )
                index += 1
        elif isinstance(item, Table):
            headers, rows = _table_rows(item)
            blocks.append({"index": index, "kind": "table", "headers": headers, "rows": rows})
            index += 1
    return blocks


def _blocks_markdown(blocks: list[dict[str, Any]], include_tables: bool = True) -> str:
    lines: list[str] = []
    for block in blocks:
        kind = block["kind"]
        if kind == "heading":
            level = int(block.get("level") or 0)
            if level == 0:
                lines.append(f"# {block['text']}".rstrip())
            else:
                lines.append(f"{'#' * min(level + 1, 6)} {block['text']}".rstrip())
        elif kind == "paragraph":
            text = block.get("text") or ""
            if text.strip():
                lines.append(text)
        elif kind == "table" and include_tables:
            headers = block.get("headers") or []
            rows = block.get("rows") or []
            if headers:
                lines.append(md_table(headers, rows))
    return "\n\n".join(line for line in lines if line != "").strip()


def _blocks_text(blocks: list[dict[str, Any]]) -> str:
    lines: list[str] = []
    for block in blocks:
        if block["kind"] == "table":
            for row in [block.get("headers") or []] + (block.get("rows") or []):
                lines.append("\t".join(scalar_text(cell) for cell in row))
        else:
            lines.append(block.get("text") or "")
    return "\n".join(lines).strip()


def _search_blocks(blocks: list[dict[str, Any]], needle: str, match_case: bool) -> list[dict[str, Any]]:
    if not needle:
        raise EngineError("INVALID_ARGUMENT", "search requires a non-empty search string")
    needle_cmp = needle if match_case else needle.casefold()
    matches: list[dict[str, Any]] = []
    for block in blocks:
        if block["kind"] == "table":
            headers = block.get("headers") or []
            for row_index, row in enumerate(block.get("rows") or []):
                row_text = " | ".join(scalar_text(cell) for cell in row)
                if (needle if match_case else needle_cmp) in (row_text if match_case else row_text.casefold()):
                    matches.append({"location": f"table[{block['index']}] row {row_index + 2}", "text": row_text, "index": block["index"]})
                    if len(matches) >= 100:
                        return matches
            continue
        text = block.get("text") or ""
        haystack = text if match_case else text.casefold()
        if needle_cmp in haystack:
            matches.append(
                {
                    "location": f"{block['kind']}[{block['index']}]",
                    "text": text,
                    "index": block["index"],
                    "headingLevel": block.get("level"),
                }
            )
            if len(matches) >= 100:
                break
    return matches


def _metadata(doc: Document) -> dict[str, Any]:
    props = doc.core_properties
    return {
        "title": props.title or "",
        "author": props.author or "",
        "subject": props.subject or "",
        "keywords": props.keywords or "",
        "comments": props.comments or "",
        "created": jsonable(props.created),
        "modified": jsonable(props.modified),
        "revision": props.revision,
        "lastModifiedBy": props.last_modified_by or "",
    }


def read_docx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    doc = Document(path)
    blocks = _build_blocks(doc)
    mode = str(args.get("mode") or "content")
    include_tables = bool(args.get("includeTables", True))
    max_chars = int(args.get("maxChars") or limits["maxTextChars"])
    metadata = _metadata(doc) if args.get("includeMetadata", True) else {}

    headings = [
        {"index": block["index"], "level": block.get("level"), "text": block["text"]}
        for block in blocks
        if block["kind"] == "heading"
    ]
    stats = {
        "blocks": len(blocks),
        "paragraphs": sum(1 for block in blocks if block["kind"] == "paragraph"),
        "tables": sum(1 for block in blocks if block["kind"] == "table"),
        "images": len(doc.inline_shapes),
        "sections": len(doc.sections),
        "headings": len(headings),
    }
    stats.update(word_counts("\n".join(block.get("text") or "" for block in blocks if block["kind"] != "table")))

    if mode == "outline":
        value = {"mode": "outline", "headings": headings, "stats": stats, "metadata": metadata}
        return {"type": "docx", **value, "totalChars": 0, "truncated": False}
    if mode == "search":
        matches = _search_blocks(blocks, str(args.get("search") or ""), bool(args.get("matchCase", False)))
        return {"type": "docx", "mode": "search", "matches": matches, "stats": stats, "metadata": metadata, "totalChars": 0, "truncated": False}

    output_format = str(args.get("format") or "markdown")
    if output_format == "json":
        payload = {"blocks": blocks, "stats": stats, "metadata": metadata}
        full_text = __import__("json").dumps(payload, ensure_ascii=False, default=jsonable)
    elif output_format == "text":
        full_text = _blocks_text(blocks)
    else:
        full_text = _blocks_markdown(blocks, include_tables=include_tables)
    text, truncated_flag = truncate(full_text, max_chars)
    return {
        "type": "docx",
        "mode": "content",
        "format": output_format,
        "text": text,
        "totalChars": len(full_text),
        "truncated": truncated_flag,
        "stats": stats,
        "metadata": metadata,
        "blocks": blocks if output_format == "json" else None,
    }


# ---------------------------------------------------------------------------
# Building primitives
# ---------------------------------------------------------------------------


def _apply_style_font(style: Any, name: str, size: float | None = None) -> None:
    try:
        style.font.name = name
        if size:
            style.font.size = Pt(size)
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), name)
    except Exception:
        # Font wiring is a presentation nicety; never fail a document for it.
        pass


def _set_cell_background(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _repeat_table_header(row: Any) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def _add_page_number(paragraph: Paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _add_hyperlink(paragraph: Paragraph, url: str, text: str) -> None:
    relationship_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.append(color)
    properties.append(underline)
    run.append(properties)
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def _enable_update_fields(doc: Document) -> None:
    """Ask Word to refresh fields (TOC, captions, cross-references) on open."""
    try:
        settings = doc.settings.element
        update = settings.find(qn("w:updateFields"))
        if update is None:
            update = OxmlElement("w:updateFields")
            settings.append(update)
        update.set(qn("w:val"), "true")
    except Exception:
        # A missing settings hook must not make a document fail to build.
        pass


def _caption_label(value: Any) -> str:
    key = scalar_text(value).strip().lower()
    return {
        "图": "图",
        "figure": "图",
        "fig": "图",
        "表": "表",
        "table": "表",
        "tab": "表",
    }.get(key, scalar_text(value).strip() or "图")


def _add_caption(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "captions", "Word captions")
    label = _caption_label(operation.get("label") or operation.get("kind"))
    text = scalar_text(operation.get("text") or "")
    position = str(operation.get("position") or "below").lower()
    paragraph = doc.add_paragraph()
    try:
        paragraph.style = doc.styles["Caption"]
    except Exception:
        pass

    anchor = None
    if operation.get("table") is not None:
        index = int(operation.get("table") or 0) - 1
        if index < 0 or index >= len(doc.tables):
            raise EngineError("TABLE_NOT_FOUND", f"table index {index + 1} is out of range", {"tables": len(doc.tables)})
        anchor = doc.tables[index]._tbl
    elif operation.get("paragraph") is not None:
        index = int(operation.get("paragraph") or 0) - 1
        if index < 0 or index >= len(doc.paragraphs):
            raise EngineError("PARAGRAPH_NOT_FOUND", f"paragraph index {index + 1} is out of range", {"paragraphs": len(doc.paragraphs)})
        anchor = doc.paragraphs[index]._p
    if anchor is not None:
        if position == "above":
            anchor.addprevious(paragraph._p)
        else:
            anchor.addnext(paragraph._p)

    label_run = paragraph.add_run(label + " ")
    label_run.bold = True
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), f" SEQ {label} \\* ARABIC ")
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "1"
    field_run.append(field_text)
    field.append(field_run)
    paragraph._p.append(field)
    if text:
        paragraph.add_run(" " + text)
    return {"label": label, "position": position, "text": text}


def _set_style_properties(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "styles", "Word style editing")
    style_name = scalar_text(operation.get("style") or "Normal")
    try:
        style = doc.styles[style_name]
    except KeyError as error:
        raise EngineError("STYLE_NOT_FOUND", f"Word style {style_name!r} was not found") from error

    font_name = operation.get("font") or operation.get("fontName")
    font_size = operation.get("size") or operation.get("fontSize")
    if font_name or font_size:
        _apply_style_font(style, scalar_text(font_name or style.font.name or "Microsoft YaHei"), float(font_size) if font_size else None)
    if "bold" in operation:
        style.font.bold = bool(operation["bold"])
    if "italic" in operation:
        style.font.italic = bool(operation["italic"])
    if operation.get("color"):
        style.font.color.rgb = RGBColor.from_string(scalar_text(operation["color"]).lstrip("#").upper())
    paragraph_format = getattr(style, "paragraph_format", None)
    if paragraph_format is not None:
        alignment = operation.get("alignment")
        if alignment and str(alignment).lower() in _ALIGN:
            paragraph_format.alignment = _ALIGN[str(alignment).lower()]
        for field, attribute in (("spaceBefore", "space_before"), ("spaceAfter", "space_after")):
            if operation.get(field) is not None:
                setattr(paragraph_format, attribute, Pt(float(operation[field])))
        if operation.get("lineSpacing") is not None:
            paragraph_format.line_spacing = float(operation["lineSpacing"])
    return {"style": style_name, "updated": sorted(operation.keys())}



def _select_paragraph(doc: Document, operation: Mapping[str, Any]) -> tuple[Paragraph, int]:
    """Select a paragraph by 1-based index or by a literal text fragment."""
    target = operation.get("paragraph")
    if target is not None:
        index = int(target) - 1
        if index < 0 or index >= len(doc.paragraphs):
            raise EngineError("PARAGRAPH_NOT_FOUND", f"paragraph index {index + 1} is out of range", {"paragraphs": len(doc.paragraphs)})
        return doc.paragraphs[index], index + 1
    find = scalar_text(operation.get("find") or "")
    if find:
        for index, paragraph in enumerate(doc.paragraphs):
            if find in _paragraph_text(paragraph):
                return paragraph, index + 1
        raise EngineError("TEXT_NOT_FOUND", f"text was not found: {find}")
    raise EngineError("INVALID_ARGUMENT", "paragraph index or find text is required")


def _revision_timestamp(value: Any = None) -> str:
    if value:
        return scalar_text(value)
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _add_comment(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "comments", "Word comments")
    paragraph, index = _select_paragraph(doc, operation)
    text = scalar_text(operation.get("text") or operation.get("comment") or "")
    if not text:
        raise EngineError("INVALID_ARGUMENT", "addComment.text must not be empty")
    author = scalar_text(operation.get("author") or "Office Forge")
    initials = scalar_text(operation.get("initials") or author[:2])
    runs = list(paragraph.runs)
    if not runs:
        runs = [paragraph.add_run("")]
    comment = doc.add_comment(runs, text=text, author=author, initials=initials)
    return {"paragraph": index, "text": text, "author": author, "commentId": int(comment.comment_id)}


def _add_tracked_revision(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "revisions", "Word tracked changes")
    paragraph, index = _select_paragraph(doc, operation)
    old_text = _paragraph_text(paragraph)
    find = scalar_text(operation.get("find") or "")
    if find and find not in old_text:
        raise EngineError("TEXT_NOT_FOUND", f"text was not found: {find}")
    new_text = operation.get("replace")
    if new_text is None:
        new_text = old_text + scalar_text(operation.get("text") or "")
    new_text = scalar_text(new_text)
    author = scalar_text(operation.get("author") or "Office Forge")
    timestamp = _revision_timestamp(operation.get("date"))
    revision_id = str(900000 + abs(hash((index, old_text, new_text))) % 100000)

    # Mark the existing paragraph content as deleted, then insert the new text.
    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)
    deletion = OxmlElement("w:del")
    deletion.set(qn("w:id"), revision_id)
    deletion.set(qn("w:author"), author)
    deletion.set(qn("w:date"), timestamp)
    old_run = OxmlElement("w:r")
    old_text_node = OxmlElement("w:delText")
    old_text_node.set(qn("xml:space"), "preserve")
    old_text_node.text = old_text
    old_run.append(old_text_node)
    deletion.append(old_run)
    paragraph._p.append(deletion)

    insertion = OxmlElement("w:ins")
    insertion.set(qn("w:id"), str(int(revision_id) + 1))
    insertion.set(qn("w:author"), author)
    insertion.set(qn("w:date"), timestamp)
    new_run = OxmlElement("w:r")
    new_text_node = OxmlElement("w:t")
    new_text_node.set(qn("xml:space"), "preserve")
    new_text_node.text = new_text
    new_run.append(new_text_node)
    insertion.append(new_run)
    paragraph._p.append(insertion)
    return {"paragraph": index, "oldText": old_text, "newText": new_text, "author": author, "revisionId": revision_id}


def _mail_merge(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "mailMerge", "Word mail merge")
    data = require_dict(operation.get("data") or operation.get("values"), "mailMerge.data")
    missing = str(operation.get("missing") or "keep")
    _replace_document_text(doc, data, missing)
    return {"fields": sorted(str(key) for key in data.keys()), "missing": missing}


# ---------------------------------------------------------------------------
# Footnotes, bookmarks and cross-references (v0.8)
# ---------------------------------------------------------------------------

_FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_FOOTNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_FOOTNOTES_PARTNAME = "/word/footnotes.xml"
_FOOTNOTES_SEPARATORS = (
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:separator/></w:r></w:p></w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
)
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _footnotes_part(doc: Document) -> Any:
    """Return the footnotes part, creating it with Word's two separator entries."""
    package = doc.part.package
    for part in package.iter_parts():
        if str(part.partname) == _FOOTNOTES_PARTNAME:
            return part
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:footnotes xmlns:w="{_W_NS}">' + _FOOTNOTES_SEPARATORS + "</w:footnotes>"
    )
    part = Part(PackURI(_FOOTNOTES_PARTNAME), _FOOTNOTES_CT, xml.encode("utf-8"), package)
    doc.part.relate_to(part, _FOOTNOTES_RT)
    return part


def _footnotes_root(part: Any) -> Any:
    from lxml import etree

    return etree.fromstring(part.blob)


def _store_footnotes_root(part: Any, root: Any) -> None:
    from lxml import etree

    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _next_footnote_id(root: Any) -> int:
    used = []
    for node in root.findall(qn("w:footnote")):
        raw = node.get(qn("w:id"))
        try:
            used.append(int(raw))
        except (TypeError, ValueError):
            continue
    return max([1] + [value + 1 for value in used if value >= 1])


def _add_footnote(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "footnotes", "Word footnotes")
    paragraph, index = _select_paragraph(doc, operation)
    text = scalar_text(operation.get("text") or operation.get("note") or "")
    if not text:
        raise EngineError("INVALID_ARGUMENT", "addFootnote.text must not be empty")

    part = _footnotes_part(doc)
    root = _footnotes_root(part)
    footnote_id = _next_footnote_id(root)

    # Body reference: a run carrying w:footnoteReference.
    run = paragraph.add_run()
    properties = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), "FootnoteReference")
    properties.append(style)
    run._r.append(properties)
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), str(footnote_id))
    run._r.append(reference)

    # Footnote body: <w:footnote><w:p> ref-run + text </w:p></w:footnote>
    note = OxmlElement("w:footnote")
    note.set(qn("w:id"), str(footnote_id))
    body_paragraph = OxmlElement("w:p")
    paragraph_properties = OxmlElement("w:pPr")
    paragraph_style = OxmlElement("w:pStyle")
    paragraph_style.set(qn("w:val"), "FootnoteText")
    paragraph_properties.append(paragraph_style)
    body_paragraph.append(paragraph_properties)

    marker_run = OxmlElement("w:r")
    marker_properties = OxmlElement("w:rPr")
    marker_style = OxmlElement("w:rStyle")
    marker_style.set(qn("w:val"), "FootnoteReference")
    marker_properties.append(marker_style)
    marker_run.append(marker_properties)
    marker_run.append(OxmlElement("w:footnoteRef"))
    body_paragraph.append(marker_run)

    text_run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = " " + text
    text_run.append(text_node)
    body_paragraph.append(text_run)

    note.append(body_paragraph)
    root.append(note)
    _store_footnotes_root(part, root)
    return {"paragraph": index, "footnoteId": footnote_id, "text": text}


def _bookmark_name(value: Any) -> str:
    """Word only accepts bookmark names that start with a letter/underscore and
    contain letters, digits or underscores, so non-ASCII requests are hashed into
    a stable ASCII name instead of collapsing into a row of underscores."""
    raw = scalar_text(value).strip()
    if not raw:
        return ""
    cleaned = re.sub(r"[^0-9A-Za-z_]", "_", raw)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned:
        digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]
        return f"bm_{digest}"
    if not re.match(r"[A-Za-z_]", cleaned[0]):
        cleaned = "bm_" + cleaned
    return cleaned[:40]


def _add_bookmark(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "bookmarks", "Word bookmarks")
    paragraph, index = _select_paragraph(doc, operation)
    name = _bookmark_name(operation.get("name") or operation.get("bookmark") or "")
    if not name:
        raise EngineError("INVALID_ARGUMENT", "addBookmark.name must not be empty")

    existing = {
        node.get(qn("w:name"))
        for node in doc.element.body.iter(qn("w:bookmarkStart"))
        if node.get(qn("w:name"))
    }
    if name in existing:
        raise EngineError("BOOKMARK_EXISTS", f"bookmark already exists: {name}", {"bookmark": name})

    used_ids = []
    for node in doc.element.body.iter(qn("w:bookmarkStart")):
        try:
            used_ids.append(int(node.get(qn("w:id"))))
        except (TypeError, ValueError):
            continue
    bookmark_id = max([0] + used_ids) + 1

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    anchor = paragraph.add_run("")
    anchor._r.addprevious(start)
    anchor._r.addnext(end)
    return {"paragraph": index, "bookmark": name, "bookmarkId": bookmark_id}


def _add_cross_reference(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "crossReferences", "Word cross-references")
    paragraph, index = _select_paragraph(doc, operation)
    requested = scalar_text(operation.get("bookmark") or operation.get("target") or "")
    name = _bookmark_name(requested)
    if not name:
        raise EngineError("INVALID_ARGUMENT", "addCrossReference.bookmark must not be empty")
    known = {
        node.get(qn("w:name"))
        for node in doc.element.body.iter(qn("w:bookmarkStart"))
        if node.get(qn("w:name"))
    }
    if name not in known:
        raise EngineError("BOOKMARK_NOT_FOUND", f"bookmark was not found: {name}", {"bookmark": name, "requested": requested})

    display = scalar_text(operation.get("text") or requested or name)
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = f" REF {name} \\h "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = display
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instruction, separate, placeholder, end):
        run._r.append(node)
    _enable_update_fields(doc)
    return {"paragraph": index, "bookmark": name, "display": display}


def _add_toc(doc: Document, title: Any = "目录") -> None:
    if title:
        doc.add_heading(scalar_text(title), level=1)
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = r'TOC \o "1-3" \h \z \u'
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "右键更新目录"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instruction, separate, placeholder, end):
        run._r.append(node)
    try:
        settings = doc.settings.element
        update = settings.find(qn("w:updateFields"))
        if update is None:
            update = OxmlElement("w:updateFields")
            settings.append(update)
        update.set(qn("w:val"), "true")
    except Exception:
        # A missing settings hook must not make a document fail to build.
        pass


def _add_picture_run(paragraph: Paragraph, value: Any, workspace: str | None) -> None:
    data = require_dict(value, "image")
    if not workspace:
        raise EngineError("INVALID_ARGUMENT", "workspace is required for image insertion")
    image_path = os.path.abspath(str(data.get("path") or ""))
    assert_inside(image_path, workspace, "image.path")
    if not os.path.isfile(image_path):
        raise EngineError("INPUT_NOT_FOUND", f"image not found: {data.get('path')}")
    kwargs: dict[str, Any] = {}
    if data.get("width"):
        kwargs["width"] = Inches(float(data["width"]))
    if data.get("height"):
        kwargs["height"] = Inches(float(data["height"]))
    run = paragraph.add_run()
    run.add_picture(image_path, **kwargs)


def _configure_document(doc: Document, spec: Mapping[str, Any]) -> None:
    font = str(spec.get("font") or "Microsoft YaHei")
    _apply_style_font(doc.styles["Normal"], font, float(spec.get("fontSize") or 10.5))
    for level in range(1, 7):
        style_name = f"Heading {level}"
        if style_name in [style.name for style in doc.styles]:
            _apply_style_font(doc.styles[style_name], font, None)
    for section in doc.sections:
        page_size = str(spec.get("pageSize") or "A4").lower()
        if page_size == "letter":
            section.page_width, section.page_height = Cm(21.59), Cm(27.94)
        else:
            section.page_width, section.page_height = Cm(21.0), Cm(29.7)
        if str(spec.get("orientation") or "portrait").lower() == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            if section.page_width < section.page_height:
                section.page_width, section.page_height = section.page_height, section.page_width
        margins = spec.get("margins") if isinstance(spec.get("margins"), Mapping) else {}
        if margins:
            section.top_margin = Cm(float(margins.get("top", 2.54)))
            section.bottom_margin = Cm(float(margins.get("bottom", 2.54)))
            section.left_margin = Cm(float(margins.get("left", 2.54)))
            section.right_margin = Cm(float(margins.get("right", 2.54)))
    header = spec.get("header")
    if header is not None:
        for section in doc.sections:
            section.header.is_linked_to_previous = False
            paragraph = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
            paragraph.text = scalar_text(header)
    footer = spec.get("footer")
    if footer is not None:
        for section in doc.sections:
            section.footer.is_linked_to_previous = False
            paragraph = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
            paragraph.text = scalar_text(footer)
            if spec.get("pageNumbers", True):
                paragraph.add_run("    ")
                _add_page_number(paragraph)


def _add_paragraph_spec(doc: Document, value: Any, workspace: str | None = None, features: Any = None) -> Paragraph:
    if isinstance(value, str):
        return doc.add_paragraph(value)
    data = require_dict(value, "paragraph")
    paragraph = doc.add_paragraph()
    text = scalar_text(data.get("text") or "")
    if data.get("link"):
        require_feature(features, "word", "hyperlinks", "Word hyperlinks")
        _add_hyperlink(paragraph, str(data["link"]), text)
    else:
        run = paragraph.add_run(text)
        if data.get("bold"):
            run.bold = True
        if data.get("italic"):
            run.italic = True
        if data.get("size"):
            run.font.size = Pt(float(data["size"]))
        if data.get("color"):
            color = str(data["color"]).lstrip("#")
            try:
                run.font.color.rgb = RGBColor.from_string(color)
            except Exception:
                pass
    if data.get("image") is not None:
        require_feature(features, "word", "images", "Word images")
        _add_picture_run(paragraph, data["image"], workspace)
    if data.get("style"):
        try:
            paragraph.style = doc.styles[str(data["style"])]
        except Exception:
            pass
    align = _ALIGN.get(str(data.get("align") or "").lower())
    if align is not None:
        paragraph.alignment = align
    if data.get("spaceBefore") is not None:
        paragraph.paragraph_format.space_before = Pt(float(data["spaceBefore"]))
    if data.get("spaceAfter") is not None:
        paragraph.paragraph_format.space_after = Pt(float(data["spaceAfter"]))
    return paragraph


def _add_list(doc: Document, values: Any, style_name: str) -> None:
    for value in require_list(values, "list"):
        if isinstance(value, Mapping):
            text = scalar_text(value.get("text") or "")
            level = max(0, min(8, int(value.get("level") or 0)))
            paragraph = doc.add_paragraph(text, style=style_name)
            if level:
                paragraph.paragraph_format.left_indent = Cm(0.75 * level)
        else:
            doc.add_paragraph(scalar_text(value), style=style_name)


def _add_table_spec(doc: Document, value: Any) -> Table:
    data = require_dict(value, "table")
    headers = list(require_list(data.get("headers"), "table.headers"))
    rows = list(require_list(data.get("rows"), "table.rows", max_items=20000))
    all_rows = ([headers] if headers else []) + rows
    if not all_rows:
        raise EngineError("INVALID_ARGUMENT", "table must contain headers or rows")
    width = max(len(row) if isinstance(row, (list, tuple)) else 1 for row in all_rows)
    table = doc.add_table(rows=len(all_rows), cols=width)
    table.style = str(data.get("style") or "Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row_index, row in enumerate(all_rows):
        cells = list(row) if isinstance(row, (list, tuple)) else [row]
        for column_index in range(width):
            cell = table.cell(row_index, column_index)
            cell.text = scalar_text(cells[column_index]) if column_index < len(cells) else ""
            if row_index == 0 and headers:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                _set_cell_background(cell, str(data.get("headerFill") or "1F4E78").lstrip("#"))
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    if headers:
        _repeat_table_header(table.rows[0])
    widths = data.get("widths")
    if isinstance(widths, list) and widths:
        for row in table.rows:
            for index, width_cm in enumerate(widths[:width]):
                row.cells[index].width = Cm(float(width_cm))
    return table


def _add_section_spec(doc: Document, value: Any, workspace: str | None = None, features: Any = None) -> None:
    data = require_dict(value, "section")
    if data.get("pageBreakBefore"):
        require_feature(features, "word", "pageBreaks", "Word page breaks")
        doc.add_page_break()
    heading = data.get("heading")
    if heading is not None:
        level = int(data.get("level") or 1)
        level = max(1, min(9, level))
        doc.add_heading(scalar_text(heading), level=level)
    for text in require_list(data.get("paragraphs"), "section.paragraphs"):
        _add_paragraph_spec(doc, text, workspace, features)
    _add_list(doc, data.get("bullets"), "List Bullet")
    _add_list(doc, data.get("numbered"), "List Number")
    for image in require_list(data.get("images"), "section.images") if data.get("images") is not None else []:
        require_feature(features, "word", "images", "Word images")
        _add_picture_run(doc.add_paragraph(), image, workspace)
    table = data.get("table")
    if table is not None:
        _add_table_spec(doc, table)
    if data.get("pageBreakAfter"):
        require_feature(features, "word", "pageBreaks", "Word page breaks")
        doc.add_page_break()


def _apply_sections(doc: Document, sections: Any, workspace: str | None = None, features: Any = None) -> None:
    for section in require_list(sections, "sections", max_items=2000):
        _add_section_spec(doc, section, workspace, features)


def _set_columns(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "columns", "Word section columns")
    count = max(1, min(4, int(operation.get("count") or operation.get("columns") or 2)))
    space = float(operation.get("space") or 36)
    section_index = int(operation.get("section") or 1) - 1
    if section_index < 0 or section_index >= len(doc.sections):
        raise EngineError("SECTION_NOT_FOUND", f"section index {section_index + 1} is out of range", {"sections": len(doc.sections)})
    section = doc.sections[section_index]
    sect_pr = section._sectPr
    columns = sect_pr.find(qn("w:cols"))
    if columns is None:
        columns = OxmlElement("w:cols")
        sect_pr.append(columns)
    columns.set(qn("w:num"), str(count))
    columns.set(qn("w:space"), str(int(round(space * 20))))
    if operation.get("separator"):
        columns.set(qn("w:sep"), "1")
    return {"section": section_index + 1, "count": count, "space": space}


def _add_watermark(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "watermark", "Word watermark")
    text = scalar_text(operation.get("text") or "CONFIDENTIAL")
    color = scalar_text(operation.get("color") or "C0C0C0").lstrip("#").upper()
    font_size = float(operation.get("fontSize") or 48)
    rotation = float(operation.get("rotation") or -45)
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        pict = OxmlElement("w:pict")
        vml_ns = "urn:schemas-microsoft-com:vml"
        shape = etree.SubElement(pict, f"{{{vml_ns}}}shape", nsmap={"v": vml_ns})
        shape.set("id", "PowerPlusWaterMarkObjectOfficeForge")
        shape.set("type", "#_x0000_t136")
        shape.set("style", (
            "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;"
            f"rotation:{rotation:g};z-index:-251658752;mso-position-horizontal:center;"
            "mso-position-vertical:center;mso-wrap-edited:f"
        ))
        shape.set("fillcolor", f"#{color}")
        shape.set("stroked", "f")
        textpath = etree.SubElement(shape, f"{{{vml_ns}}}textpath")
        textpath.set("style", f"font-family:'Microsoft YaHei';font-size:{font_size:g}pt")
        textpath.set("string", text)
        run._r.append(pict)
    return {"text": text, "color": color, "fontSize": font_size, "rotation": rotation}


def _add_content_control(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "contentControls", "Word content controls")
    tag = scalar_text(operation.get("tag") or "officeforge-control")
    alias = scalar_text(operation.get("alias") or tag)
    text = scalar_text(operation.get("text") or "")
    paragraph = doc.add_paragraph()
    if text:
        paragraph.add_run(text)
    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    alias_node = OxmlElement("w:alias")
    alias_node.set(qn("w:val"), alias)
    tag_node = OxmlElement("w:tag")
    tag_node.set(qn("w:val"), tag)
    sdt_pr.append(alias_node)
    sdt_pr.append(tag_node)
    sdt.append(sdt_pr)
    sdt_content = OxmlElement("w:sdtContent")
    body = paragraph._p.getparent()
    body.replace(paragraph._p, sdt)
    sdt_content.append(paragraph._p)
    sdt.append(sdt_content)
    return {"tag": tag, "alias": alias, "text": text}


def _add_citation(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "citations", "Word citations")
    if operation.get("paragraph") is not None or operation.get("find"):
        paragraph, index = _select_paragraph(doc, operation)
    else:
        paragraph = doc.add_paragraph()
        index = len(doc.paragraphs)
    if operation.get("index") is not None:
        marker = f"[{int(operation['index'])}]"
    else:
        marker = scalar_text(operation.get("text") or operation.get("marker") or "")
    if not marker:
        raise EngineError("INVALID_ARGUMENT", "addCitation requires index or text")
    run = paragraph.add_run(marker)
    if bool(operation.get("superscript", True)):
        run.font.superscript = True
    return {"citation": marker, "paragraph": index}


def _add_bibliography(doc: Document, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "word", "citations", "Word citations")
    entries = require_list(operation.get("entries"), "bibliography.entries", max_items=500)
    title = scalar_text(operation.get("title") or "参考文献")
    doc.add_heading(title, level=1)
    for index, entry in enumerate(entries, 1):
        if isinstance(entry, Mapping):
            text = scalar_text(entry.get("text") or entry.get("title") or entry.get("citation") or "")
        else:
            text = scalar_text(entry)
        doc.add_paragraph(f"[{index}] {text}" if text else f"[{index}]")
    return {"title": title, "count": len(entries)}


def _convert_to_dotx(path: str) -> None:
    with zipfile.ZipFile(path, "r") as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                )
            archive.writestr(info, data)


def _save_as_template(doc: Document, output: Any, workspace: str | None, limits: Mapping[str, Any], overwrite: bool) -> dict[str, Any]:
    if not output:
        raise EngineError("INVALID_ARGUMENT", "saveAsTemplate.output is required")
    if not workspace:
        raise EngineError("INVALID_ARGUMENT", "workspace is required for saveAsTemplate")
    target = str(output)
    absolute = target if os.path.isabs(target) else os.path.join(workspace, target)
    absolute = os.path.abspath(absolute)
    assert_inside(absolute, workspace, "saveAsTemplate.output")
    if not absolute.lower().endswith(".dotx"):
        raise EngineError("INVALID_ARGUMENT", "saveAsTemplate.output must end with .dotx")

    def writer(temp: str) -> None:
        doc.save(temp)
        _convert_to_dotx(temp)

    size, backup = atomic_save(absolute, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]), make_backup=False)
    return {"templatePath": absolute, "bytes": size, "backupPath": backup}


def create_docx(path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str | None = None, features: Any = None) -> dict[str, Any]:
    doc = Document()
    _configure_document(doc, spec)
    title = spec.get("title")
    if title is not None:
        doc.add_heading(scalar_text(title), level=0)
    subtitle = spec.get("subtitle")
    if subtitle is not None:
        paragraph = doc.add_paragraph(scalar_text(subtitle))
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if spec.get("toc"):
        require_feature(features, "word", "toc", "Word table of contents")
        _add_toc(doc, spec.get("tocTitle") or (spec["toc"] if isinstance(spec["toc"], str) else "目录"))
    _apply_sections(doc, spec.get("sections"), workspace, features)
    if spec.get("columns") is not None:
        columns_spec = spec["columns"] if isinstance(spec["columns"], Mapping) else {"count": spec["columns"]}
        _set_columns(doc, columns_spec, features)
    if spec.get("watermark") is not None:
        watermark_spec = spec["watermark"] if isinstance(spec["watermark"], Mapping) else {"text": spec["watermark"]}
        _add_watermark(doc, watermark_spec, features)
    for control in require_list(spec.get("contentControls"), "contentControls", max_items=200) if spec.get("contentControls") is not None else []:
        _add_content_control(doc, require_dict(control, "contentControl"), features)
    if spec.get("bibliography") is not None:
        _add_bibliography(doc, {"entries": spec.get("bibliography"), "title": spec.get("bibliographyTitle")}, features)
    for section in doc.sections:
        if section.footer.paragraphs and spec.get("pageNumbers", True) and spec.get("footer") is None:
            _add_page_number(section.footer.paragraphs[0])

    # ``properties`` is the cross-format spelling; the top-level keys are the
    # historical docx spelling and stay supported so old specs keep working.
    metadata = dict(spec.get("properties") or {})
    for legacy in ("title", "subject", "author", "keywords", "comments", "category"):
        if legacy in spec and legacy not in metadata:
            metadata[legacy] = spec[legacy]
    properties = apply_core_properties(doc.core_properties, metadata)

    def writer(temp: str) -> None:
        doc.save(temp)
        if path.lower().endswith(".dotx"):
            _convert_to_dotx(temp)

    size, _ = atomic_save(path, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {
        "type": "docx",
        "path": path,
        "bytes": size,
        "sections": len(require_list(spec.get("sections"), "sections")),
        "properties": properties,
        "template": path.lower().endswith(".dotx"),
    }


# ---------------------------------------------------------------------------
# Editing
# ---------------------------------------------------------------------------


def _replace_in_paragraph(paragraph: Paragraph, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    original = _paragraph_text(paragraph)
    if not original:
        return 0
    flags = 0 if match_case else re.IGNORECASE
    if regex:
        pattern = re.compile(find, flags)
        new_text, count = pattern.subn(replacement, original)
    else:
        if match_case:
            count = original.count(find)
            new_text = original.replace(find, replacement)
        else:
            pattern = re.compile(re.escape(find), re.IGNORECASE)
            new_text, count = pattern.subn(lambda _m: replacement, original)
    if count == 0:
        return 0
    # First try run-local replacement, which preserves all original run formatting.
    run_text = "".join(run.text for run in paragraph.runs)
    if run_text == original:
        local_count = 0
        local_ok = True
        for run in paragraph.runs:
            if regex:
                run.text, number = re.subn(find, replacement, run.text, flags=flags)
            else:
                if match_case:
                    number = run.text.count(find)
                    run.text = run.text.replace(find, replacement)
                else:
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    run.text, number = pattern.subn(lambda _m: replacement, run.text)
            local_count += number
        if local_count == count:
            return count
        # A match crossed a run boundary: rebuild the paragraph in its first run.
        for run in paragraph.runs[1:]:
            run._r.getparent().remove(run._r)
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)
        return count
    # Hyperlinks/complex content: rewrite the first text node and blank the rest.
    text_nodes = [node for node in paragraph._p.iter() if node.tag == qn("w:t")]
    if not text_nodes:
        return 0
    text_nodes[0].text = new_text
    for node in text_nodes[1:]:
        node.text = ""
    return count


def _replace_in_table(table: Table, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    # Iterative walk: nested tables are queued instead of recursively called.
    count = 0
    stack = [table]
    while stack:
        current = stack.pop()
        for row in current.rows:
            seen: set[int] = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for paragraph in cell.paragraphs:
                    count += _replace_in_paragraph(paragraph, find, replacement, regex, match_case)
                stack.extend(cell.tables)
    return count


def _replace_everywhere(doc: Document, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    count = 0
    for paragraph in doc.paragraphs:
        count += _replace_in_paragraph(paragraph, find, replacement, regex, match_case)
    for table in doc.tables:
        count += _replace_in_table(table, find, replacement, regex, match_case)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                count += _replace_in_paragraph(paragraph, find, replacement, regex, match_case)
            for table in part.tables:
                count += _replace_in_table(table, find, replacement, regex, match_case)
    return count


def _set_core(doc: Document, values: Mapping[str, Any]) -> None:
    props = doc.core_properties
    mapping = {
        "title": "title",
        "author": "author",
        "subject": "subject",
        "keywords": "keywords",
        "comments": "comments",
        "category": "category",
        "lastModifiedBy": "last_modified_by",
        "revision": "revision",
    }
    for key, attribute in mapping.items():
        if key in values:
            setattr(props, attribute, values[key] if key == "revision" else scalar_text(values[key]))


def edit_docx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str | None = None, features: Any = None) -> dict[str, Any]:
    doc = Document(path)
    operations = require_list(args.get("operations"), "operations", max_items=200)
    if not operations:
        raise EngineError("INVALID_ARGUMENT", "operations must contain at least one edit")
    applied: list[dict[str, Any]] = []
    template_output: Any = None
    for item in operations:
        operation = require_dict(item, "operation")
        name = str(operation.get("op") or "")
        if name == "replace":
            find = scalar_text(operation.get("find") or "")
            if not find:
                raise EngineError("INVALID_ARGUMENT", "replace.find must not be empty")
            count = _replace_everywhere(doc, find, scalar_text(operation.get("replace") or ""), bool(operation.get("regex", False)), bool(operation.get("matchCase", False)))
            applied.append({"op": "replace", "matches": count})
        elif name == "append":
            _apply_sections(doc, operation.get("sections"), workspace, features)
            applied.append({"op": "append", "sections": len(require_list(operation.get("sections"), "sections"))})
        elif name == "addImage":
            require_feature(features, "word", "images", "Word images")
            image = operation.get("image") if operation.get("image") is not None else operation
            _add_picture_run(doc.add_paragraph(), image, workspace)
            applied.append({"op": "addImage"})
        elif name == "addToc":
            require_feature(features, "word", "toc", "Word table of contents")
            _add_toc(doc, operation.get("title") or "目录")
            applied.append({"op": "addToc"})
        elif name == "addCaption":
            result = _add_caption(doc, operation, features)
            applied.append({"op": "addCaption", **result})
        elif name == "addComment":
            result = _add_comment(doc, operation, features)
            applied.append({"op": "addComment", **result})
        elif name == "trackRevision":
            result = _add_tracked_revision(doc, operation, features)
            applied.append({"op": "trackRevision", **result})
        elif name == "mailMerge":
            result = _mail_merge(doc, operation, features)
            applied.append({"op": "mailMerge", **result})
        elif name == "addFootnote":
            result = _add_footnote(doc, operation, features)
            applied.append({"op": "addFootnote", **result})
        elif name == "addBookmark":
            result = _add_bookmark(doc, operation, features)
            applied.append({"op": "addBookmark", **result})
        elif name == "addCrossReference":
            result = _add_cross_reference(doc, operation, features)
            applied.append({"op": "addCrossReference", **result})
        elif name == "addCitation":
            result = _add_citation(doc, operation, features)
            applied.append({"op": "addCitation", **result})
        elif name == "addBibliography":
            result = _add_bibliography(doc, operation, features)
            applied.append({"op": "addBibliography", **result})
        elif name == "setColumns":
            result = _set_columns(doc, operation, features)
            applied.append({"op": "setColumns", **result})
        elif name == "addWatermark":
            result = _add_watermark(doc, operation, features)
            applied.append({"op": "addWatermark", **result})
        elif name == "addContentControl":
            result = _add_content_control(doc, operation, features)
            applied.append({"op": "addContentControl", **result})
        elif name == "saveAsTemplate":
            template_output = operation.get("output") or operation.get("path")
            if not template_output:
                raise EngineError("INVALID_ARGUMENT", "saveAsTemplate.output is required")
            applied.append({"op": "saveAsTemplate", "output": template_output})
        elif name == "setStyle":
            result = _set_style_properties(doc, operation, features)
            applied.append({"op": "setStyle", **result})
        elif name == "updateFields":
            require_feature(features, "word", "fields", "Word field refresh")
            _enable_update_fields(doc)
            applied.append({"op": "updateFields"})
        elif name == "setCore":
            values = require_dict(operation.get("values"), "setCore.values")
            _set_core(doc, values)
            applied.append({"op": "setCore", "keys": sorted(values.keys())})
        else:
            raise EngineError("UNSUPPORTED_OPERATION", f"unsupported docx edit operation: {name or '(empty)'}", {"allowed": ["replace", "append", "addImage", "addToc", "addCaption", "addComment", "trackRevision", "mailMerge", "addFootnote", "addBookmark", "addCrossReference", "addCitation", "addBibliography", "setColumns", "addWatermark", "addContentControl", "saveAsTemplate", "setStyle", "updateFields", "setCore"]})

    if template_output is not None:
        result = _save_as_template(doc, template_output, workspace, limits, bool(args.get("overwrite", False)))
        return {"type": "docx", "templatePath": result["templatePath"], "bytes": result["bytes"], "operations": applied}
    size, backup = atomic_save(path, lambda temp: doc.save(temp), overwrite=True, max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(args.get("backup", True)))
    return {"type": "docx", "path": path, "bytes": size, "backupPath": backup, "operations": applied}


# ---------------------------------------------------------------------------
# Template expansion
# ---------------------------------------------------------------------------


def _expand_table_loops(table: Table, data: Any, missing: str) -> int:
    """Expand one-level ``{{#each items}}`` table row loops in place."""
    count = 0
    index = 0
    while index < len(table.rows):
        rows = list(table.rows)
        if index >= len(rows):
            break
        start_row = rows[index]
        start_text = "\n".join(_cell_text(cell) for cell in start_row.cells)
        match = re.search(r"\{\{\s*#(?:each\s+)?([\w.]+)\s*\}\}", start_text)
        if not match:
            index += 1
            continue
        key = match.group(1)
        end_index = index + 1
        end_pattern = re.compile(r"\{\{\s*/(?:each(?:\s+([\w.]+))?|([\w.]+))\s*\}\}")
        while end_index < len(rows):
            end_text = "\n".join(_cell_text(cell) for cell in rows[end_index].cells)
            end_match = end_pattern.search(end_text)
            end_key = (end_match.group(1) or end_match.group(2)) if end_match else None
            if end_match and (end_key is None or end_key == key):
                break
            end_index += 1
        if end_index >= len(rows):
            index += 1
            continue
        items = get_path(data, key, [])
        if not isinstance(items, list):
            items = [] if items in (None, "") else [items]
        template_rows = rows[index + 1 : end_index]
        end_tr = rows[end_index]._tr
        for item_index, item in enumerate(items):
            context = dict(data) if isinstance(data, Mapping) else {}
            if isinstance(item, Mapping):
                context.update(item)
            context["this"] = item
            context["index"] = item_index + 1
            for template_row in template_rows:
                clone = copy.deepcopy(template_row._tr)
                end_tr.addprevious(clone)
                new_row = _Row(clone, table)
                for cell in new_row.cells:
                    for paragraph in cell.paragraphs:
                        _replace_paragraph_text(paragraph, context, missing)
                    for nested in cell.tables:
                        _replace_table_text(nested, context, missing)
        for row in [start_row] + template_rows + [rows[end_index]]:
            row._tr.getparent().remove(row._tr)
        count += len(items)
        index += len(items) * len(template_rows)
    return count


def _replace_paragraph_text(paragraph: Paragraph, data: Any, missing: str) -> None:
    original = _paragraph_text(paragraph)
    replaced = render_placeholders(original, data, missing)
    if replaced == original:
        return
    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        paragraph.add_run(replaced)


def _replace_table_text(table: Table, data: Any, missing: str) -> None:
    # Iterative walk: never recurse through nested tables, so a malformed or
    # self-referential document cannot grow the Python stack.
    stack = [table]
    while stack:
        current = stack.pop()
        for row in current.rows:
            seen: set[int] = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for paragraph in cell.paragraphs:
                    _replace_paragraph_text(paragraph, data, missing)
                stack.extend(cell.tables)


def _replace_document_text(doc: Document, data: Any, missing: str) -> None:
    for paragraph in doc.paragraphs:
        _replace_paragraph_text(paragraph, data, missing)
    for table in doc.tables:
        _replace_table_text(table, data, missing)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                _replace_paragraph_text(paragraph, data, missing)
            for table in part.tables:
                _replace_table_text(table, data, missing)


def template_docx(template: str, output: str, data: Mapping[str, Any], options: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    doc = Document(template)
    missing = str(options.get("missing") or "keep")
    loops = 0
    if options.get("expandLoops", True):
        for table in doc.tables:
            loops += _expand_table_loops(table, data, missing)
    _replace_document_text(doc, data, missing)
    size, backup = atomic_save(output, lambda temp: doc.save(temp), overwrite=bool(options.get("overwrite", False)), max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(options.get("backup", False)))
    return {"type": "docx", "path": output, "bytes": size, "tablesExpanded": loops, "backupPath": backup}


# ---------------------------------------------------------------------------
# Read-only document audit (health / profile)
# ---------------------------------------------------------------------------

_PLACEHOLDER_RE = re.compile(r"\{\{\s*[^{}\n]+\s*\}\}")
_MARKER_RE = re.compile(r"\b(?:TODO|TBD|FIXME)\b", re.IGNORECASE)


def _docx_audit_sequence(doc: Document) -> list[dict[str, Any]]:
    """Raw body sequence including empty paragraphs, for structural checks."""
    sequence: list[dict[str, Any]] = []
    for index, item in enumerate(_iter_block_items(doc)):
        if isinstance(item, Paragraph):
            text = _paragraph_text(item)
            level = _heading_level(item)
            style = item.style.name if item.style is not None else ""
            sequence.append(
                {
                    "index": index,
                    "kind": "heading" if level is not None else "paragraph",
                    "level": level,
                    "style": style,
                    "text": text,
                }
            )
        else:
            sequence.append({"index": index, "kind": "table", "text": ""})
    return sequence


def _health_docx(path: str, args: Mapping[str, Any]) -> dict[str, Any]:
    doc = Document(path)
    sequence = _docx_audit_sequence(doc)
    issues: list[dict[str, Any]] = []
    max_issues = max(1, min(int(args.get("maxIssues") or 200), 1000))

    def add(code: str, severity: str, message: str, **extra: Any) -> None:
        issues.append({"code": code, "severity": severity, "message": message, **extra})

    body_text = [block for block in sequence if block["kind"] in ("paragraph", "heading") and block["text"].strip()]
    has_table = any(block["kind"] == "table" for block in sequence)
    if not body_text and not has_table:
        add("empty_document", "high", "document contains no body text or tables")

    previous_heading: dict[str, Any] | None = None
    seen_headings: dict[tuple[int, str], dict[str, Any]] = {}
    for position, block in enumerate(sequence):
        if block["kind"] != "heading":
            continue
        level = int(block["level"] or 0)
        text = block["text"].strip()
        if level > 0 and (previous_heading is None or int(previous_heading["level"] or 0) == 0):
            if level > 1:
                add("heading_jump", "medium", "first numbered heading starts below level 1", index=block["index"], level=level, text=text)
        elif level > 0 and previous_heading is not None:
            previous_level = int(previous_heading["level"] or 0)
            if previous_level > 0 and level > previous_level + 1:
                add(
                    "heading_jump",
                    "medium",
                    f"heading level jumps from {previous_level} to {level}",
                    index=block["index"],
                    level=level,
                    previousLevel=previous_level,
                    text=text,
                )
        previous_heading = block

        key = (level, text.casefold())
        if text:
            first = seen_headings.get(key)
            if first is None:
                seen_headings[key] = block
            else:
                add("duplicate_heading", "medium", "heading text repeats at the same level", index=block["index"], firstIndex=first["index"], level=level, text=text)

        # A heading followed by another heading at the same or higher level, or
        # ending the document, has no body content. A table counts as content:
        # "1.1 数据" directly above a table is a normal report pattern, not an
        # orphan heading. A deeper heading is a normal outline transition
        # (document title -> chapter -> section), not an orphan heading.
        next_block = sequence[position + 1] if position + 1 < len(sequence) else None
        if next_block is None:
            add("orphan_heading", "medium", "heading is not followed by body text", index=block["index"], level=level, text=text)
        elif next_block["kind"] == "heading" and next_block.get("level", 0) <= level:
            add("orphan_heading", "medium", "heading is not followed by body text", index=block["index"], level=level, text=text)

    empty_run: list[dict[str, Any]] = []

    def flush_empty_run() -> None:
        if len(empty_run) >= 2:
            add(
                "empty_paragraph_run",
                "low",
                f"{len(empty_run)} consecutive empty paragraphs",
                index=empty_run[0]["index"],
                count=len(empty_run),
            )
        empty_run.clear()

    for block in sequence:
        if block["kind"] == "paragraph" and not block["text"].strip():
            empty_run.append(block)
        else:
            flush_empty_run()
        if block["kind"] == "paragraph":
            text = block["text"]
            if _PLACEHOLDER_RE.search(text):
                add("placeholder_left", "high", "unresolved template placeholder remains in body text", index=block["index"], text=text[:200])
            elif _MARKER_RE.search(text):
                add("placeholder_left", "low", "draft marker remains in body text", index=block["index"], text=text[:200])
            if len(text) > 1500:
                add("long_paragraph", "low", "paragraph exceeds 1500 characters", index=block["index"], length=len(text))
    flush_empty_run()

    summary: dict[str, int] = {}
    for issue in issues:
        summary[issue["code"]] = summary.get(issue["code"], 0) + 1
    high = sum(1 for issue in issues if issue["severity"] == "high")
    medium = sum(1 for issue in issues if issue["severity"] == "medium")
    low = sum(1 for issue in issues if issue["severity"] == "low")
    score = max(0, 100 - high * 10 - medium * 5 - low * 2)
    return {
        "type": "docx",
        "mode": "health",
        "score": score,
        "issueCount": len(issues),
        "summary": summary,
        "issues": issues[:max_issues],
        "truncated": len(issues) > max_issues,
        "paragraphs": sum(1 for block in sequence if block["kind"] == "paragraph"),
        "headings": sum(1 for block in sequence if block["kind"] == "heading"),
        "tables": sum(1 for block in sequence if block["kind"] == "table"),
        "images": len(doc.inline_shapes),
    }


def _profile_docx(path: str) -> dict[str, Any]:
    doc = Document(path)
    blocks = _build_blocks(doc)
    headings_by_level: dict[str, int] = {}
    style_counts: dict[str, int] = {}
    for block in blocks:
        if block["kind"] == "heading":
            key = str(block.get("level"))
            headings_by_level[key] = headings_by_level.get(key, 0) + 1
        if block["kind"] in ("paragraph", "heading"):
            style = block.get("style") or "(none)"
            style_counts[style] = style_counts.get(style, 0) + 1
    stats = {
        "blocks": len(blocks),
        "paragraphs": sum(1 for block in blocks if block["kind"] == "paragraph"),
        "headings": sum(1 for block in blocks if block["kind"] == "heading"),
        "tables": len(doc.tables),
        "images": len(doc.inline_shapes),
        "sections": len(doc.sections),
    }
    stats.update(word_counts("\n".join(block.get("text") or "" for block in blocks if block["kind"] != "table")))
    sections = []
    for index, section in enumerate(doc.sections):
        width = section.page_width.cm if section.page_width is not None else None
        height = section.page_height.cm if section.page_height is not None else None
        sections.append(
            {
                "index": index,
                "orientation": "landscape" if section.orientation == WD_ORIENT.LANDSCAPE else "portrait",
                "pageWidthCm": round(width, 2) if width is not None else None,
                "pageHeightCm": round(height, 2) if height is not None else None,
            }
        )
    tables = [{"index": index, "rows": len(table.rows), "columns": len(table.columns)} for index, table in enumerate(doc.tables)]
    return {
        "type": "docx",
        "mode": "profile",
        "stats": stats,
        "headingsByLevel": headings_by_level,
        "styles": [{"name": name, "count": count} for name, count in sorted(style_counts.items(), key=lambda item: (-item[1], item[0]))],
        "sections": sections,
        "tables": tables,
        "metadata": _metadata(doc),
    }


def audit_docx(path: str, mode: str, args: Mapping[str, Any]) -> dict[str, Any]:
    requested = str(mode or "health").lower()
    if requested == "health":
        return _health_docx(path, args)
    if requested == "profile":
        return _profile_docx(path)
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported docx audit mode: {requested}", {"supported": ["health", "profile"]})
