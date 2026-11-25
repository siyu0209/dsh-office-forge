"""Best-effort PPTX -> Deck IR importer.

The importer never pretends a complex PowerPoint round-trips perfectly.  It
extracts every object it understands into Deck IR, copies pictures into a
workspace-local media folder, and reports advanced objects it detected but
could not fully reconstruct in ``fidelity`` / ``warnings``.
"""

from __future__ import annotations

import hashlib
import os
import re
from typing import Any, Mapping

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

import office_pptx
from office_common import EngineError, assert_inside


def _inches(value: Any) -> float | None:
    try:
        return round(float(value) / 914400.0, 3)
    except Exception:
        return None


def _bounds(shape: Any) -> dict[str, Any]:
    return {
        "left": _inches(getattr(shape, "left", None)),
        "top": _inches(getattr(shape, "top", None)),
        "width": _inches(getattr(shape, "width", None)),
        "height": _inches(getattr(shape, "height", None)),
    }


def _chart_block(chart: Any) -> dict[str, Any] | None:
    try:
        name = str(chart.chart_type or "").upper()
        kind = "pie" if "PIE" in name else "line" if "LINE" in name else "bar"
        categories: list[str] = []
        try:
            categories = [str(item) for item in chart.plots[0].categories]
        except Exception:
            pass
        series: list[dict[str, Any]] = []
        for item in chart.series:
            try:
                values = [float(value) if value is not None else 0 for value in item.values]
            except Exception:
                values = []
            series.append({"name": str(item.name or "Series"), "values": values})
        if not series:
            return None
        return {"type": "chart", "chart": {"type": kind, "categories": categories, "series": series}}
    except Exception:
        return None


def _shape_name(shape: Any) -> str:
    try:
        value = getattr(shape, "auto_shape_type", None)
        raw = str(value or "RECTANGLE").split(".")[-1].split(" ")[0]
        return raw.upper()
    except Exception:
        return "RECTANGLE"


def _is_smartart(shape: Any) -> bool:
    try:
        return "drawingml/2006/diagram" in shape.element.xml
    except Exception:
        return False


def _media_kind(shape: Any) -> str | None:
    try:
        xml = shape.element.xml
        if "<p:video" in xml or "videoFile" in xml:
            return "video"
        if "<p:audio" in xml or "audioFile" in xml:
            return "audio"
    except Exception:
        pass
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
            return "media"
    except Exception:
        pass
    return None


def _shape_style(shape: Any) -> dict[str, Any]:
    style: dict[str, Any] = {}
    try:
        fill = shape.fill
        if fill.type is not None and str(fill.type).endswith("SOLID (1)"):
            style["fill"] = str(fill.fore_color.rgb)
    except Exception:
        pass
    try:
        line = shape.line
        if line.color and line.color.type is not None:
            style["line"] = str(line.color.rgb)
        if line.width:
            style["lineWidth"] = round(float(line.width) / 12700.0, 2)
    except Exception:
        pass
    try:
        if shape.rotation:
            style["rotation"] = round(float(shape.rotation), 2)
    except Exception:
        pass
    return style


def _media_part(shape: Any) -> Any | None:
    """Return the first embedded audio/video part reachable from a shape."""
    try:
        relations = shape.part.rels.values()
    except Exception:
        return None
    for relation in relations:
        try:
            if relation.is_external:
                continue
            part = relation.target_part
            name = str(part.partname).lower()
            if name.endswith((".mp4", ".m4a", ".mp3", ".wav", ".mov", ".avi", ".wmv", ".mpeg", ".mpg")):
                return part
        except Exception:
            continue
    return None


def _extract_media(shape: Any, media_dir: str, stem: str, slide_no: int, index: int, kind: str) -> dict[str, Any] | None:
    part = _media_part(shape)
    if part is None:
        return None
    extension = os.path.splitext(str(part.partname))[1].lower() or ".bin"
    target = os.path.join(media_dir, f"{stem}-s{slide_no}-{index + 1}-{kind}{extension}")
    try:
        with open(target, "wb") as handle:
            handle.write(part.blob)
    except Exception:
        return None
    return {"type": "media", "kind": kind, "path": target, "name": shape.name, **_bounds(shape)}


def _smartart_block(shape: Any) -> dict[str, Any]:
    texts: list[str] = []
    try:
        for node in shape.element.iter():
            if str(node.tag).endswith("}t") and node.text and node.text.strip():
                text = node.text.strip()
                if text not in texts:
                    texts.append(text)
    except Exception:
        pass
    return {"type": "smartart", "name": shape.name, "items": texts or [shape.name], **_bounds(shape)}


def _transition_block(slide: Any) -> dict[str, Any] | None:
    try:
        node = slide.element.find(qn("p:transition"))
    except Exception:
        return None
    if node is None:
        return None
    effect = None
    for child in node:
        tag = str(child.tag).split("}")[-1]
        if tag not in ("extLst", "sndAc"):
            effect = tag
            break
    raw = getattr(node, "xml", "") or ""
    if not effect and "morph" in raw.lower():
        effect = "morph"
    return {"effect": effect or "fade"}


def _extract_blocks(slide: Any, media_dir: str, stem: str, slide_no: int) -> tuple[list[dict[str, Any]], dict[str, int]]:
    blocks: list[dict[str, Any]] = []
    stats = {"text": 0, "table": 0, "chart": 0, "image": 0, "shape": 0, "smartart": 0, "media": 0}
    title_shape = slide.shapes.title
    title = title_shape.text if title_shape is not None and title_shape.has_text_frame else ""
    if title.strip():
        blocks.append({"type": "title", "text": title.strip(), **_bounds(title_shape)})
        stats["text"] += 1
    for index, shape in enumerate(slide.shapes):
        if title_shape is not None and shape == title_shape:
            continue
        bounds = _bounds(shape)
        if _is_smartart(shape):
            blocks.append(_smartart_block(shape))
            stats["smartart"] += 1
            continue
        media_kind = _media_kind(shape)
        if media_kind:
            media_block = _extract_media(shape, media_dir, stem, slide_no, index, media_kind)
            if media_block is not None:
                blocks.append(media_block)
            else:
                blocks.append({"type": "text", "text": f"[Media:{media_kind}] {shape.name}", **bounds})
            stats["media"] += 1
            continue
        kind = office_pptx._shape_kind(shape)
        if kind == "table":
            rows = office_pptx._table_matrix(shape.table)
            if rows:
                blocks.append({"type": "table", "headers": rows[0], "rows": rows[1:], **bounds})
                stats["table"] += 1
            continue
        if kind == "chart":
            block = _chart_block(shape.chart)
            if block:
                block.update(bounds)
                blocks.append(block)
                stats["chart"] += 1
            continue
        if kind == "picture":
            try:
                image = shape.image
                ext = str(image.ext or "png").lower()
                if ext == "jpeg":
                    ext = "jpg"
                name = f"{stem}-s{slide_no}-{index + 1}.{ext}"
                target = os.path.join(media_dir, name)
                with open(target, "wb") as handle:
                    handle.write(image.blob)
                blocks.append({"type": "image", "path": target, "name": shape.name, **bounds})
                stats["image"] += 1
            except Exception:
                blocks.append({"type": "text", "text": f"[Image] {shape.name}", **bounds})
                stats["image"] += 1
            continue
        if kind == "text" and shape.has_text_frame:
            paragraphs = [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]
            if paragraphs:
                blocks.append({"type": "paragraphs", "items": paragraphs, "name": shape.name, **bounds})
                stats["text"] += 1
            continue
        # Auto shapes / connectors / freeforms without text still carry layout.
        shape_block = {"type": "shape", "shape": _shape_name(shape), "name": shape.name, "text": shape.text_frame.text if shape.has_text_frame else "", **bounds}
        shape_block.update(_shape_style(shape))
        blocks.append(shape_block)
        stats["shape"] += 1
    return blocks, stats


def import_pptx(path: str, workspace: str, limits: Mapping[str, Any], args: Mapping[str, Any] | None = None) -> dict[str, Any]:
    del args  # reserved for future selection options
    source = os.path.abspath(path)
    assert_inside(source, workspace, "PPTX import")
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"file not found: {path}")
    if not source.lower().endswith(".pptx"):
        raise EngineError("UNSUPPORTED_FORMAT", "PPTX import requires a .pptx file")
    raw = open(source, "rb").read()
    digest = hashlib.sha256(raw).hexdigest()
    stem = re.sub(r"[^0-9A-Za-z_-]+", "-", os.path.splitext(os.path.basename(source))[0]).strip("-") or "deck"
    media_dir = os.path.join(workspace, "_imports", stem)
    os.makedirs(media_dir, exist_ok=True)
    presentation = Presentation(source)
    deck: dict[str, Any] = {
        "title": str(presentation.core_properties.title or stem),
        "theme": None,
        "slides": [],
        "importMeta": {
            "source": os.path.basename(source),
            "sourceSha256": digest,
            "slideSize": {k: _inches(v) for k, v in (("width", presentation.slide_width), ("height", presentation.slide_height))},
        },
    }
    totals = {"slides": 0, "text": 0, "table": 0, "chart": 0, "image": 0, "shape": 0, "smartart": 0, "media": 0, "notes": 0, "comments": 0, "transitions": 0, "animations": 0}
    warnings: list[str] = []
    comments = office_pptx._list_comments(presentation)
    for slide_no, slide in enumerate(presentation.slides, start=1):
        blocks, stats = _extract_blocks(slide, media_dir, stem, slide_no)
        slide_spec: dict[str, Any] = {"layout": "content", "blocks": blocks}
        for block in blocks:
            if isinstance(block, Mapping) and str(block.get("type") or "").lower() == "title" and block.get("text"):
                slide_spec["title"] = str(block["text"])
                break
        for key, value in stats.items():
            totals[key] += value
        notes = ""
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text
        except Exception:
            pass
        if notes:
            slide_spec["notes"] = notes
            totals["notes"] += 1
        matching_comments = [item for item in comments if int(item.get("slide") or 0) == slide_no]
        if matching_comments:
            slide_spec["comments"] = [{"text": item.get("text"), "author": item.get("author") or "Office Forge"} for item in matching_comments]
            totals["comments"] += len(matching_comments)
        transition = _transition_block(slide)
        if transition:
            slide_spec["transition"] = transition
            totals["transitions"] += 1
        if slide.element.find(qn("p:timing")) is not None:
            slide_spec["animations"] = True
            totals["animations"] += 1
        if stats["smartart"]:
            slide_spec["smartart"] = True
        if stats["media"]:
            slide_spec["media"] = True
        deck["slides"].append(slide_spec)
        totals["slides"] += 1
    if totals["smartart"]:
        warnings.append("SmartArt detected; text items are imported as a smartart block and rendered as editable shape diagrams, not native diagrams")
    if totals["media"]:
        warnings.append("Audio/video detected; embedded media is copied to the workspace and imported as media blocks where the package exposes it")
    if totals["animations"]:
        warnings.append("Animation timing XML detected; timing can be preserved only when the source shape mapping is supported")
    if totals["transitions"]:
        warnings.append("Slide transitions detected and marked; legacy renderer supports fade/push/wipe/... including morph")
    return {
        "type": "deck",
        "mode": "import",
        "schemaVersion": 1,
        "deck": deck,
        "fidelity": {"counts": totals, "mediaDir": media_dir, "warnings": warnings},
    }
