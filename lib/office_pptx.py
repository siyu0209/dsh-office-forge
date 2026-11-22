"""PowerPoint (.pptx) reading, creation, editing and template expansion."""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime, timezone
from typing import Any, Mapping

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from office_common import (
    EngineError,
    apply_core_properties,
    assert_inside,
    atomic_save,
    jsonable,
    render_placeholders,
    require_dict,
    require_feature,
    require_list,
    scalar_text,
    truncate,
)

_ALLOWED = (".pptx",)
_LAYOUTS = {
    "title": 0,
    "title_slide": 0,
    "content": 1,
    "title_content": 1,
    "section": 2,
    "section_header": 2,
    "two_content": 3,
    "comparison": 4,
    "title_only": 5,
    "blank": 6,
    "caption": 7,
    "picture": 8,
}
_ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_THEMES_FALLBACK = {
    "forge": {"label": "Forge 深蓝", "background": "0B1F33", "titleColor": "FFFFFF", "bodyColor": "D9E2EC", "accent": "2D9CDB", "font": "Microsoft YaHei"},
    "academic": {"label": "Academic 学术", "background": "FFFFFF", "titleColor": "1F3864", "bodyColor": "222222", "accent": "B8860B", "font": "SimSun"},
}


def _load_themes() -> dict[str, dict[str, Any]]:
    """Load themes from the package resource file.

    The resource file is data-only; a malformed or absent file falls back to
    the built-in minimum so a damaged installation cannot stop the plugin.
    """
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resources", "themes.json")
    try:
        import json
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        themes = payload.get("themes") if isinstance(payload, dict) else None
        result = {}
        for item in themes or []:
            if isinstance(item, dict) and item.get("id"):
                result[str(item["id"])] = dict(item)
        return result or dict(_THEMES_FALLBACK)
    except Exception:
        return dict(_THEMES_FALLBACK)


_THEMES = _load_themes()


def _load_json_resource(name: str, key: str) -> dict[str, dict[str, Any]]:
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resources", name)
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        rows = payload.get(key) if isinstance(payload, dict) else None
        result: dict[str, dict[str, Any]] = {}
        for row in rows or []:
            if isinstance(row, dict) and row.get("id"):
                result[str(row["id"])] = dict(row)
        return result
    except Exception:
        return {}


_LAYOUT_CATALOG = _load_json_resource("layouts.json", "layouts")
_TEMPLATE_CATALOG = _load_json_resource("templates.json", "templates")


def _shape_kind(shape: Any) -> str:
    if shape.has_text_frame:
        return "text"
    if shape.has_table:
        return "table"
    if shape.has_chart:
        return "chart"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    return "shape"


def _table_matrix(table: Any) -> list[list[str]]:
    return [[scalar_text(cell.text) for cell in row.cells] for row in table.rows]


def _chart_summary(chart: Any) -> dict[str, Any]:
    try:
        series_names = [series.name for series in chart.series]
    except Exception:
        series_names = []
    categories: list[str] = []
    try:
        categories = [scalar_text(category) for category in chart.plots[0].categories]
    except Exception:
        pass
    return {"chartType": str(chart.chart_type), "series": series_names, "categories": categories}


def _slide_record(slide: Any, index: int, limits: Mapping[str, Any]) -> dict[str, Any]:
    title_shape = slide.shapes.title
    title = title_shape.text if title_shape is not None and title_shape.has_text_frame else ""
    shapes: list[dict[str, Any]] = []
    text_parts: list[str] = []
    for shape in slide.shapes:
        kind = _shape_kind(shape)
        if kind == "text" and shape.has_text_frame:
            text = shape.text_frame.text
            if text.strip():
                shapes.append({"kind": "text", "name": shape.name, "text": text})
                text_parts.append(text)
        elif kind == "table":
            matrix = _table_matrix(shape.table)
            shapes.append({"kind": "table", "name": shape.name, "rows": matrix})
            text_parts.extend(" | ".join(row) for row in matrix)
        elif kind == "chart":
            shapes.append({"kind": "chart", "name": shape.name, **_chart_summary(shape.chart)})
        elif kind == "picture":
            shapes.append({"kind": "picture", "name": shape.name, "width": shape.width, "height": shape.height})
    notes = ""
    try:
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
    except Exception:
        notes = ""
    return {"index": index, "title": title, "shapes": shapes, "text": "\n".join(text_parts), "notes": notes}


def _slide_markdown(slide: Mapping[str, Any]) -> str:
    lines = [f"## Slide {slide['index']}: {slide['title'] or '(untitled)'}"]
    for shape in slide.get("shapes") or []:
        if shape["kind"] == "text":
            lines.append(shape["text"])
        elif shape["kind"] == "table":
            rows = shape.get("rows") or []
            if rows:
                header = rows[0]
                body = rows[1:]
                lines.append("| " + " | ".join(header) + " |")
                lines.append("| " + " | ".join("---" for _ in header) + " |")
                for row in body:
                    lines.append("| " + " | ".join(row) + " |")
    if slide.get("notes"):
        lines.append(f"_Notes: {slide['notes']}_")
    return "\n\n".join(lines)


def read_pptx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    presentation = Presentation(path)
    slides = [_slide_record(slide, index + 1, limits) for index, slide in enumerate(presentation.slides)]
    total = len(slides)
    max_slides = min(int(args.get("maxSlides") or limits["maxSlides"]), int(limits["maxSlides"]))
    selected = slides[:max_slides]
    mode = str(args.get("mode") or "content")
    if mode == "outline":
        return {
            "type": "pptx",
            "mode": "outline",
            "slides": [{"index": slide["index"], "title": slide["title"], "shapeCount": len(slide["shapes"]), "notes": bool(slide["notes"])} for slide in slides],
            "totalSlides": total,
            "slideSize": {"width": presentation.slide_width, "height": presentation.slide_height},
        }
    if mode == "search":
        needle = str(args.get("search") or "")
        if not needle:
            raise EngineError("INVALID_ARGUMENT", "search requires a non-empty search string")
        match_case = bool(args.get("matchCase", False))
        needle_cmp = needle if match_case else needle.casefold()
        matches: list[dict[str, Any]] = []
        for slide in slides:
            haystack = (slide["title"] + "\n" + slide["text"] + "\n" + slide["notes"])
            if needle_cmp in (haystack if match_case else haystack.casefold()):
                matches.append({"slide": slide["index"], "title": slide["title"], "text": slide["text"][:500]})
        return {"type": "pptx", "mode": "search", "matches": matches, "totalSlides": total}
    output_format = str(args.get("format") or "markdown")
    if output_format == "json":
        import json
        payload = {"slides": selected, "totalSlides": total}
        full_text = json.dumps(payload, ensure_ascii=False, default=jsonable)
    elif output_format == "text":
        full_text = "\n\n".join(f"[Slide {slide['index']}] {slide['title']}\n{slide['text']}" for slide in selected)
    else:
        full_text = "\n\n".join(_slide_markdown(slide) for slide in selected)
    text, truncated_flag = truncate(full_text, int(args.get("maxChars") or limits["maxTextChars"]))
    return {
        "type": "pptx",
        "mode": "content",
        "format": output_format,
        "text": text,
        "totalChars": len(full_text),
        "truncated": truncated_flag,
        "slides": selected,
        "totalSlides": total,
    }


# ---------------------------------------------------------------------------
# Creating
# ---------------------------------------------------------------------------


def _set_slide_size(presentation: Any, spec: Mapping[str, Any]) -> None:
    size = str(spec.get("slideSize") or "16:9")
    if size == "4:3":
        presentation.slide_width = Inches(10)
        presentation.slide_height = Inches(7.5)
    else:
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)


def _get_body_placeholder(slide: Any) -> Any:
    for placeholder in slide.placeholders:
        if placeholder.placeholder_format.idx != 0:
            return placeholder
    return None


def _style_table(table: Any, header_fill: str = "1F4E78") -> None:
    fill = RGBColor.from_string(header_fill.lstrip("#"))
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            if row_index == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = fill
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _add_table(slide: Any, value: Any) -> None:
    data = require_dict(value, "table")
    rows = list(require_list(data.get("rows"), "table.rows", max_items=500))
    if not rows:
        return
    width = max(len(row) if isinstance(row, (list, tuple)) else 1 for row in rows)
    left = Inches(float(data.get("left") or 0.8))
    top = Inches(float(data.get("top") or 1.8))
    table_width = Inches(float(data.get("width") or 11.5))
    table_height = Inches(float(data.get("height") or max(0.4, 0.4 * len(rows))))
    table = slide.shapes.add_table(len(rows), width, left, top, table_width, table_height).table
    for row_index, row in enumerate(rows):
        cells = list(row) if isinstance(row, (list, tuple)) else [row]
        for column_index in range(width):
            table.cell(row_index, column_index).text = scalar_text(cells[column_index]) if column_index < len(cells) else ""
    _style_table(table, str(data.get("headerFill") or "1F4E78"))


def _add_chart(slide: Any, value: Any) -> None:
    data = require_dict(value, "chart")
    kinds = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    chart_type = kinds.get(str(data.get("type") or "bar").lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
    chart_data = CategoryChartData()
    chart_data.categories = [scalar_text(item) for item in require_list(data.get("categories"), "chart.categories", max_items=200)]
    for series in require_list(data.get("series"), "chart.series", max_items=20):
        series_data = require_dict(series, "chart.series item")
        chart_data.add_series(scalar_text(series_data.get("name") or "Series"), list(require_list(series_data.get("values"), "chart.series.values", max_items=200)))
    slide.shapes.add_chart(
        chart_type,
        Inches(float(data.get("left") or 0.8)),
        Inches(float(data.get("top") or 1.8)),
        Inches(float(data.get("width") or 11.5)),
        Inches(float(data.get("height") or 4.8)),
        chart_data,
    )


def _add_transition(slide: Any, value: Any, features: Any) -> None:
    require_feature(features, "ppt", "transitions", "PowerPoint transitions")
    spec = require_dict(value, "transition")
    effect = str(spec.get("effect") or "fade").lower()
    allowed = {"fade", "push", "wipe", "split", "cover", "uncover", "zoom", "dissolve", "cut", "morph"}
    if effect not in allowed:
        raise EngineError("INVALID_ARGUMENT", f"unsupported transition effect: {effect}", {"allowed": sorted(allowed)})
    for existing in slide.element.findall(qn("p:transition")):
        slide.element.remove(existing)
    transition = etree.SubElement(slide.element, qn("p:transition"))
    if spec.get("speed"):
        transition.set("spd", str(spec["speed"]))
    if spec.get("advanceAfter") is not None:
        transition.set("advTm", str(int(float(spec["advanceAfter"]) * 1000)))
    node = etree.SubElement(transition, qn("p:" + effect))
    if effect == "morph" and spec.get("option"):
        node.set("option", str(spec["option"]))


def _add_shape(slide: Any, value: Any, features: Any) -> None:
    require_feature(features, "ppt", "shapes", "PowerPoint shapes")
    spec = require_dict(value, "shape")
    shape_name = str(spec.get("shape") or "rectangle").upper()
    shape_type = getattr(MSO_SHAPE, shape_name, MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(
        shape_type,
        Inches(float(spec.get("left") or 0.8)),
        Inches(float(spec.get("top") or 1.8)),
        Inches(float(spec.get("width") or 2.0)),
        Inches(float(spec.get("height") or 1.0)),
    )
    if spec.get("fill"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(str(spec["fill"]).lstrip("#"))
    if spec.get("line"):
        shape.line.color.rgb = RGBColor.from_string(str(spec["line"]).lstrip("#"))
        shape.line.width = Pt(float(spec.get("lineWidth") or 1))
    if spec.get("text") is not None:
        shape.text_frame.text = scalar_text(spec["text"])
    if spec.get("rotation"):
        shape.rotation = float(spec["rotation"])


def _add_textbox(slide: Any, value: Any, features: Any) -> None:
    require_feature(features, "ppt", "shapes", "PowerPoint text boxes")
    spec = require_dict(value, "textbox")
    shape = slide.shapes.add_textbox(
        Inches(float(spec.get("left") or 0.8)),
        Inches(float(spec.get("top") or 1.8)),
        Inches(float(spec.get("width") or 4.0)),
        Inches(float(spec.get("height") or 1.0)),
    )
    frame = shape.text_frame
    frame.word_wrap = bool(spec.get("wordWrap", True))
    frame.text = scalar_text(spec.get("text") or "")
    paragraph = frame.paragraphs[0]
    alignment = str(spec.get("align") or "").lower()
    if alignment in _ALIGN:
        paragraph.alignment = _ALIGN[alignment]
    for run in paragraph.runs:
        if spec.get("font"):
            run.font.name = scalar_text(spec["font"])
        if spec.get("fontSize"):
            run.font.size = Pt(float(spec["fontSize"]))
        if "bold" in spec:
            run.font.bold = bool(spec["bold"])
        if "italic" in spec:
            run.font.italic = bool(spec["italic"])
        if spec.get("color"):
            run.font.color.rgb = RGBColor.from_string(str(spec["color"]).lstrip("#"))
    if spec.get("fill"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(str(spec["fill"]).lstrip("#"))
    if spec.get("line"):
        shape.line.color.rgb = RGBColor.from_string(str(spec["line"]).lstrip("#"))
        shape.line.width = Pt(float(spec.get("lineWidth") or 1))


def _add_connector(slide: Any, value: Any, features: Any) -> None:
    require_feature(features, "ppt", "shapes", "PowerPoint connectors")
    spec = require_dict(value, "connector")
    kind = str(spec.get("connector") or "straight").lower()
    connector_type = {"straight": MSO_CONNECTOR.STRAIGHT, "elbow": MSO_CONNECTOR.ELBOW, "curve": MSO_CONNECTOR.CURVE}.get(kind)
    if connector_type is None:
        raise EngineError("INVALID_ARGUMENT", f"unsupported connector type: {kind}", {"allowed": ["straight", "elbow", "curve"]})
    begin = require_dict(spec.get("begin") or {"x": 1, "y": 1}, "connector.begin")
    end = require_dict(spec.get("end") or {"x": 3, "y": 2}, "connector.end")
    connector = slide.shapes.add_connector(
        connector_type,
        Inches(float(begin.get("x") or 0)),
        Inches(float(begin.get("y") or 0)),
        Inches(float(end.get("x") or 0)),
        Inches(float(end.get("y") or 0)),
    )
    if spec.get("color"):
        connector.line.color.rgb = RGBColor.from_string(str(spec["color"]).lstrip("#"))
    if spec.get("width"):
        connector.line.width = Pt(float(spec["width"]))
    if spec.get("dash"):
        dash = getattr(MSO_LINE_DASH_STYLE, str(spec["dash"]).upper(), None)
        if dash is not None:
            connector.line.dash_style = dash



def _add_animation(slide: Any, value: Any, features: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "animations", "PowerPoint animations")
    spec = require_dict(value, "animation")
    effect = str(spec.get("effect") or "fade").lower()
    allowed = {"fade", "fly", "wipe", "zoom", "appear"}
    if effect not in allowed:
        raise EngineError("INVALID_ARGUMENT", f"unsupported animation effect: {effect}", {"allowed": sorted(allowed)})
    shape_id = spec.get("shapeId") or (slide.shapes[0].shape_id if len(slide.shapes) else 2)
    duration = int(float(spec.get("duration") or 1.0) * 1000)
    for existing in slide.element.findall(qn("p:timing")):
        slide.element.remove(existing)
    timing = etree.SubElement(slide.element, qn("p:timing"))
    tn_lst = etree.SubElement(timing, qn("p:tnLst"))
    par = etree.SubElement(tn_lst, qn("p:par"))
    c_tn = etree.SubElement(par, qn("p:cTn"))
    c_tn.set("id", "1")
    c_tn.set("dur", str(duration))
    c_tn.set("fill", "hold")
    child = etree.SubElement(c_tn, qn("p:childTnLst"))
    par2 = etree.SubElement(child, qn("p:par"))
    c_tn2 = etree.SubElement(par2, qn("p:cTn"))
    c_tn2.set("id", "2")
    c_tn2.set("dur", str(duration))
    c_tn2.set("fill", "hold")
    child2 = etree.SubElement(c_tn2, qn("p:childTnLst"))
    effect_node = etree.SubElement(child2, qn("p:animEffect"))
    effect_node.set("transition", "in")
    effect_node.set("filter", "fade" if effect == "fade" else effect)
    behavior = etree.SubElement(effect_node, qn("p:cBhvr"))
    behavior_tn = etree.SubElement(behavior, qn("p:cTn"))
    behavior_tn.set("id", "3")
    behavior_tn.set("dur", str(duration))
    target = etree.SubElement(behavior, qn("p:tgtEl"))
    sp_target = etree.SubElement(target, qn("p:spTgt"))
    sp_target.set("spid", str(shape_id))
    return {"effect": effect, "shapeId": int(shape_id), "durationMs": duration}


def _add_smartart(slide: Any, value: Any, features: Any) -> dict[str, Any]:
    """Create an editable SmartArt-style diagram from basic shapes.

    Native SmartArt requires a diagram data part; these are ordinary PowerPoint
    shapes so they remain editable without external Office automation.
    """
    require_feature(features, "ppt", "smartart", "PowerPoint SmartArt-style diagrams")
    spec = require_dict(value, "smartArt")
    items = require_list(spec.get("items"), "smartArt.items", max_items=20)
    if not items:
        raise EngineError("INVALID_ARGUMENT", "smartArt.items must not be empty")
    layout = str(spec.get("layout") or "horizontal").lower()
    left = float(spec.get("left") or 1.0)
    top = float(spec.get("top") or 2.0)
    width = float(spec.get("width") or 2.2)
    height = float(spec.get("height") or 1.0)
    gap = float(spec.get("gap") or 0.2)
    fill = spec.get("fill") or "1F4E79"
    line = spec.get("line") or "FFFFFF"
    for index, item in enumerate(items):
        if layout == "vertical":
            x, y = left, top + index * (height + gap)
        else:
            x, y = left + index * (width + gap), top
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(width), Inches(height),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(str(fill).lstrip("#"))
        shape.line.color.rgb = RGBColor.from_string(str(line).lstrip("#"))
        text = item.get("text") if isinstance(item, Mapping) else item
        shape.text_frame.text = scalar_text(text)
        if index:
            if layout == "vertical":
                begin = (left + width, top + (index - 1) * (height + gap) + height)
                end = (x, y + height / 2)
            else:
                begin = (left + (index - 1) * (width + gap) + width, top + height / 2)
                end = (x, y + height / 2)
            connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(begin[0]), Inches(begin[1]), Inches(end[0]), Inches(end[1]))
            connector.line.color.rgb = RGBColor.from_string(str(line).lstrip("#"))
    return {"items": len(items), "layout": layout, "nativeSmartArt": False}


def _add_media(slide: Any, value: Any, features: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "media", "PowerPoint audio/video")
    spec = require_dict(value, "media")
    kind = str(spec.get("kind") or "video").lower()
    if kind not in {"audio", "video"}:
        raise EngineError("INVALID_ARGUMENT", f"unsupported media kind: {kind}", {"allowed": ["audio", "video"]})
    path_value = spec.get("path") or spec.get("source")
    if not path_value:
        raise EngineError("INVALID_ARGUMENT", "media.path is required")
    media_path = os.path.abspath(str(path_value))
    assert_inside(media_path, os.environ.get("OFFICEFORGE_WORKSPACE", os.getcwd()))
    if not os.path.isfile(media_path):
        raise EngineError("INPUT_NOT_FOUND", f"media not found: {path_value}")
    poster = spec.get("poster")
    poster_path = None
    if poster:
        poster_path = os.path.abspath(str(poster))
        if not os.path.isfile(poster_path):
            raise EngineError("INPUT_NOT_FOUND", f"poster not found: {poster}")
    mime = str(spec.get("mimeType") or ("audio/mpeg" if kind == "audio" else "video/mp4"))
    movie = slide.shapes.add_movie(
        media_path,
        Inches(float(spec.get("left") or 0.8)),
        Inches(float(spec.get("top") or 1.8)),
        Inches(float(spec.get("width") or 5.0)),
        Inches(float(spec.get("height") or 2.8)),
        poster_frame_image=poster_path,
        mime_type=mime,
    )
    if spec.get("autoplay") or spec.get("loop"):
        for tag in ("p:video", "p:audio"):
            node = movie._element.find(".//" + qn(tag))
            if node is not None:
                if spec.get("autoplay"):
                    node.set("autoplay", "1")
                if spec.get("loop"):
                    node.set("loop", "1")
                break
    return {"kind": kind, "path": media_path, "mimeType": mime, "autoplay": bool(spec.get("autoplay")), "loop": bool(spec.get("loop"))}


_COMMENTS_CT = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
_COMMENT_AUTHORS_CT = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _presentation_part(presentation: Any) -> Any:
    return presentation.part


def _comments_part(slide: Any) -> Any:
    from pptx.opc.packuri import PackURI
    from pptx.opc.package import Part
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    for relation in slide.part.rels.values():
        if relation.reltype == RT.COMMENTS and not relation.is_external:
            return relation.target_part
    used = {str(relation.target_part.partname) for relation in slide.part.rels.values() if not relation.is_external}
    index = 1
    while True:
        partname = f"/ppt/comments/comment{index}.xml"
        if partname not in used and not any(str(part.partname) == partname for part in slide.part.package.iter_parts()):
            break
        index += 1
    blob = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:cmLst xmlns:p="{_P_NS}"/>'.encode("utf-8")
    part = Part(PackURI(partname), _COMMENTS_CT, slide.part.package, blob)
    slide.part.relate_to(part, RT.COMMENTS)
    return part


def _comment_authors_part(presentation: Any) -> Any:
    from pptx.opc.packuri import PackURI
    from pptx.opc.package import Part
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    for relation in presentation.part.rels.values():
        if relation.reltype == RT.COMMENT_AUTHORS and not relation.is_external:
            return relation.target_part
    blob = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:cmAuthorLst xmlns:p="{_P_NS}"/>'.encode("utf-8")
    part = Part(PackURI("/ppt/commentAuthors.xml"), _COMMENT_AUTHORS_CT, presentation.part.package, blob)
    presentation.part.relate_to(part, RT.COMMENT_AUTHORS)
    return part


def _register_comment_author(presentation: Any, author: str, initials: str) -> int:
    part = _comment_authors_part(presentation)
    root = etree.fromstring(part.blob)
    existing = {node.get("name"): int(node.get("id")) for node in root.findall(qn("p:cmAuthor"))}
    if author in existing:
        return existing[author]
    author_id = max(existing.values()) + 1 if existing else 0
    node = etree.SubElement(root, qn("p:cmAuthor"))
    node.set("id", str(author_id))
    node.set("name", author)
    node.set("initials", initials or author[:1])
    node.set("lastIdx", "0")
    node.set("clrIdx", str(author_id % 6))
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return author_id


def _add_comment(slide: Any, value: Any, features: Any, presentation: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "comments", "PowerPoint comments")
    spec = require_dict(value, "comment")
    text = scalar_text(spec.get("text") or "")
    if not text:
        raise EngineError("INVALID_ARGUMENT", "comment.text must not be empty")
    author = scalar_text(spec.get("author") or "Office Forge")
    author_id = _register_comment_author(presentation, author, scalar_text(spec.get("initials") or ""))
    part = _comments_part(slide)
    root = etree.fromstring(part.blob)
    existing = root.findall(qn("p:cm"))
    index = max((int(node.get("idx") or 0) for node in existing), default=0) + 1
    node = etree.SubElement(root, qn("p:cm"))
    node.set("authorId", str(author_id))
    node.set("dt", scalar_text(spec.get("date") or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000")))
    node.set("idx", str(index))
    position = etree.SubElement(node, qn("p:pos"))
    position.set("x", str(int(float(spec.get("x") or 10))))
    position.set("y", str(int(float(spec.get("y") or 10))))
    text_node = etree.SubElement(node, qn("p:text"))
    text_node.text = text
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return {"commentIndex": index, "author": author, "authorId": author_id, "text": text}


def _list_comments(presentation: Any) -> list[dict[str, Any]]:
    authors: dict[str, str] = {}
    for relation in presentation.part.rels.values():
        if relation.reltype.endswith("/commentAuthors") and not relation.is_external:
            root = etree.fromstring(relation.target_part.blob)
            authors = {node.get("id"): node.get("name") or "" for node in root.findall(qn("p:cmAuthor"))}
    result: list[dict[str, Any]] = []
    for index, slide in enumerate(presentation.slides, start=1):
        for relation in slide.part.rels.values():
            if not relation.reltype.endswith("/comments") or relation.is_external:
                continue
            root = etree.fromstring(relation.target_part.blob)
            for node in root.findall(qn("p:cm")):
                result.append({
                    "slide": index,
                    "index": int(node.get("idx") or 0),
                    "authorId": node.get("authorId"),
                    "author": authors.get(node.get("authorId"), ""),
                    "date": node.get("dt"),
                    "text": (node.findtext(qn("p:text")) or ""),
                })
    return result


def _catalog_items(spec: Mapping[str, Any], key: str = "items") -> list[Any]:
    value = spec.get(key)
    if value is None:
        value = spec.get("bullets")
    if value is None:
        value = spec.get("paragraphs")
    if value is None:
        value = spec.get("body")
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def _add_catalog_text(slide: Any, text: Any, left: float, top: float, width: float, height: float, size: float = 18, bold: bool = False, align: Any = None) -> Any:
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.word_wrap = True
    frame.text = scalar_text(text)
    for paragraph in frame.paragraphs:
        if align is not None:
            paragraph.alignment = align
        for run in paragraph.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
    return shape


def _add_catalog_card(slide: Any, text: Any, left: float, top: float, width: float, height: float, fill: str = "1F4E79") -> Any:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(fill).lstrip("#"))
    shape.line.color.rgb = RGBColor.from_string("FFFFFF")
    shape.text_frame.text = scalar_text(text)
    for paragraph in shape.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(15)
    return shape


def _split_catalog_items(items: list[Any], groups: int) -> list[list[Any]]:
    groups = max(1, groups)
    buckets: list[list[Any]] = [[] for _ in range(groups)]
    for index, item in enumerate(items):
        buckets[index % groups].append(item)
    return buckets


def _apply_catalog_layout(slide: Any, spec: Mapping[str, Any], layout: Mapping[str, Any], workspace: str, features: Any = None) -> None:
    require_feature(features, "ppt", "layouts", "PowerPoint layout library")
    family = str(layout.get("family") or "bullets")
    title = spec.get("title")
    if title is not None and slide.shapes.title is not None:
        slide.shapes.title.text = scalar_text(title)
    subtitle = spec.get("subtitle")
    if subtitle is not None:
        _add_catalog_text(slide, subtitle, 0.9, 1.55, 11.2, 0.8, size=22, align=PP_ALIGN.CENTER if family in {"title", "section", "closing", "quote"} else None)
    items = _catalog_items(spec)
    if family in {"title", "section", "closing"}:
        if not subtitle:
            _add_catalog_text(slide, spec.get("body") or "", 1.0, 2.0, 11.0, 2.0, size=20, align=PP_ALIGN.CENTER)
        if family == "closing":
            _add_catalog_text(slide, spec.get("contact") or "", 1.0, 4.6, 11.0, 0.8, size=16, align=PP_ALIGN.CENTER)
        return
    if family == "bullets":
        body = "\n".join(scalar_text(item.get("text") if isinstance(item, Mapping) else item) for item in items)
        _add_catalog_text(slide, body, 1.0, 2.0, 11.0, 4.2, size=20)
        return
    if family in {"two-col", "three-col"}:
        count = 2 if family == "two-col" else 3
        columns = spec.get("columns")
        if isinstance(columns, list) and columns:
            buckets = [columns[i] if i < len(columns) else [] for i in range(count)]
        else:
            buckets = _split_catalog_items(items, count)
        width = 11.0 / count - 0.3
        for index, bucket in enumerate(buckets):
            text = "\n".join(scalar_text(item.get("text") if isinstance(item, Mapping) else item) for item in bucket)
            _add_catalog_text(slide, text, 0.8 + index * (width + 0.3), 2.0, width, 3.8, size=17)
        return
    if family == "four-quadrant":
        labels = items or ["价值", "成本", "风险", "机会"]
        for index in range(4):
            col, row = index % 2, index // 2
            _add_catalog_card(slide, labels[index] if index < len(labels) else "", 0.8 + col * 5.9, 1.9 + row * 2.25, 5.5, 1.9,
                              fill="1F4E79" if index % 3 else "2D9CDB")
        return
    if family == "kpi":
        kpis = spec.get("kpis") or items or [{"label": "指标", "value": "—"}]
        for index, item in enumerate(kpis[:4]):
            value = item.get("value") if isinstance(item, Mapping) else item
            label = item.get("label") if isinstance(item, Mapping) else ""
            _add_catalog_text(slide, value, 0.8 + index * 2.95, 2.0, 2.6, 1.2, size=36, bold=True, align=PP_ALIGN.CENTER)
            _add_catalog_text(slide, label, 0.8 + index * 2.95, 3.25, 2.6, 0.7, size=15, align=PP_ALIGN.CENTER)
        return
    if family == "timeline":
        labels = items or ["阶段一", "阶段二", "阶段三", "阶段四"]
        y = 2.7
        _add_catalog_text(slide, "→", 0.8, y - 0.45, 11.2, 0.5, size=30)
        for index, label in enumerate(labels[:6]):
            x = 0.9 + index * (10.9 / max(1, min(6, len(labels))))
            _add_catalog_card(slide, label, x, y, 1.6, 1.0, fill="2D9CDB" if index % 2 else "1F4E79")
        return
    if family == "process":
        labels = items or ["输入", "处理", "输出"]
        for index, label in enumerate(labels[:6]):
            _add_catalog_card(slide, label, 0.8 + index * 1.95, 2.3, 1.7, 1.2, fill="1F4E79" if index % 2 == 0 else "2D9CDB")
        return
    if family == "comparison":
        labels = items or ["方案 A", "方案 B"]
        for index in range(2):
            _add_catalog_card(slide, labels[index] if index < len(labels) else "", 0.8 + index * 5.9, 1.9, 5.4, 3.8, fill="1F4E79" if index == 0 else "2D9CDB")
        return
    if family == "matrix":
        labels = items or ["高", "低", "高", "低"]
        for index in range(4):
            col, row = index % 2, index // 2
            _add_catalog_card(slide, labels[index] if index < len(labels) else "", 0.8 + col * 5.9, 1.9 + row * 2.25, 5.5, 1.9, fill="1F4E79" if index % 2 == 0 else "2D9CDB")
        return
    if family == "quote":
        text = spec.get("quote") or spec.get("body") or (items[0] if items else "")
        _add_catalog_text(slide, text, 0.9, 2.1, 11.2, 2.6, size=28, bold=True, align=PP_ALIGN.CENTER)
        return
    if family in {"image-left", "image-right", "full-image", "gallery"}:
        images = spec.get("images") or ([spec["image"]] if spec.get("image") else [])
        image = images[0] if images else None
        if image:
            image_path = os.path.abspath(str(image.get("path") if isinstance(image, Mapping) else image))
            assert_inside(image_path, workspace)
            if not os.path.isfile(image_path):
                raise EngineError("INPUT_NOT_FOUND", f"image not found: {image_path}")
            if family == "full-image":
                slide.shapes.add_picture(image_path, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
            else:
                left = 0.8 if family == "image-left" else 6.8
                slide.shapes.add_picture(image_path, Inches(left), Inches(1.8), width=Inches(5.7), height=Inches(4.3))
        body = "\n".join(scalar_text(item) for item in items)
        if family != "full-image":
            _add_catalog_text(slide, body, 0.8 if family == "image-right" else 6.8, 2.0, 5.7, 3.8, size=18)
        return
    if family == "table":
        table = spec.get("table")
        if table is not None:
            _add_table(slide, table)
        return
    if family == "chart":
        chart = spec.get("chart")
        if chart is not None:
            require_feature(features, "ppt", "charts", "PowerPoint charts")
            _add_chart(slide, chart)
        return
    if family == "chart-split":
        chart = spec.get("chart")
        if chart is not None:
            require_feature(features, "ppt", "charts", "PowerPoint charts")
            _add_chart(slide, chart)
        body = "\n".join(scalar_text(item) for item in items)
        _add_catalog_text(slide, body, 7.0, 2.0, 5.4, 3.8, size=17)
        return
    if family == "cards":
        cards = spec.get("cards") or items or ["卡片一", "卡片二", "卡片三"]
        count = min(3, len(cards))
        for index in range(count):
            item = cards[index]
            text = item.get("text") if isinstance(item, Mapping) else item
            title_text = item.get("title") if isinstance(item, Mapping) else ""
            _add_catalog_card(slide, f"{title_text}\n{text}" if title_text else text, 0.7 + index * 4.15, 2.0, 3.8, 3.2,
                              fill="1F4E79" if index % 2 == 0 else "2D9CDB")
        return
    body = "\n".join(scalar_text(item) for item in items)
    _add_catalog_text(slide, body, 1.0, 2.0, 11.0, 3.8, size=18)


def _populate_slide(slide: Any, spec: Mapping[str, Any], workspace: str, features: Any = None) -> None:
    title = spec.get("title")
    if title is not None and slide.shapes.title is not None:
        slide.shapes.title.text = scalar_text(title)
    subtitle = spec.get("subtitle")
    if subtitle is not None:
        body = _get_body_placeholder(slide)
        if body is not None and body.has_text_frame:
            body.text_frame.text = scalar_text(subtitle)
    bullets = spec.get("bullets")
    paragraphs = spec.get("paragraphs")
    body_shape = _get_body_placeholder(slide)
    body_frame = body_shape.text_frame if body_shape is not None and body_shape.has_text_frame else None
    if bullets is not None or paragraphs is not None:
        if body_frame is None:
            body_shape = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(4.8))
            body_frame = body_shape.text_frame
        body_frame.clear()
        first = True
        for item in require_list(bullets, "bullets", max_items=200) if bullets is not None else []:
            paragraph = body_frame.paragraphs[0] if first else body_frame.add_paragraph()
            first = False
            if isinstance(item, Mapping):
                paragraph.text = scalar_text(item.get("text") or "")
                paragraph.level = max(0, min(8, int(item.get("level") or 0)))
            else:
                paragraph.text = scalar_text(item)
        for item in require_list(paragraphs, "paragraphs", max_items=200) if paragraphs is not None else []:
            paragraph = body_frame.paragraphs[0] if first else body_frame.add_paragraph()
            first = False
            if isinstance(item, Mapping):
                paragraph.text = scalar_text(item.get("text") or "")
                if item.get("size"):
                    for run in paragraph.runs:
                        run.font.size = Pt(float(item["size"]))
                if item.get("bold"):
                    for run in paragraph.runs:
                        run.font.bold = True
                align = _ALIGN.get(str(item.get("align") or "").lower())
                if align is not None:
                    paragraph.alignment = align
            else:
                paragraph.text = scalar_text(item)
    table = spec.get("table")
    if table is not None:
        _add_table(slide, table)
    chart = spec.get("chart")
    if chart is not None:
        require_feature(features, "ppt", "charts", "PowerPoint charts")
        _add_chart(slide, chart)
    for image in require_list(spec.get("images"), "images", max_items=20) if spec.get("images") is not None else []:
        require_feature(features, "ppt", "images", "PowerPoint images")
        image_spec = require_dict(image, "image")
        image_path = os.path.abspath(str(image_spec.get("path") or ""))
        assert_inside(image_path, workspace)
        if not os.path.isfile(image_path):
            raise EngineError("INPUT_NOT_FOUND", f"image not found: {image_spec.get('path')}")
        kwargs: dict[str, Any] = {"left": Inches(float(image_spec.get("left") or 0.8)), "top": Inches(float(image_spec.get("top") or 1.8))}
        if image_spec.get("width"):
            kwargs["width"] = Inches(float(image_spec["width"]))
        if image_spec.get("height"):
            kwargs["height"] = Inches(float(image_spec["height"]))
        slide.shapes.add_picture(image_path, **kwargs)
    for shape_spec in require_list(spec.get("shapes"), "shapes", max_items=100) if spec.get("shapes") is not None else []:
        _add_shape(slide, shape_spec, features)
    if spec.get("transition") is not None:
        _add_transition(slide, spec["transition"], features)
    if spec.get("animation") is not None:
        _add_animation(slide, spec["animation"], features)
    if spec.get("smartArt") is not None:
        _add_smartart(slide, spec["smartArt"], features)
    if spec.get("media") is not None:
        _add_media(slide, spec["media"], features)
    notes = spec.get("notes")
    if notes is not None:
        slide.notes_slide.notes_text_frame.text = scalar_text(notes)
    background = spec.get("background")
    if background:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(str(background).lstrip("#"))


def _add_slide(presentation: Any, spec: Mapping[str, Any], workspace: str, features: Any = None) -> Any:
    layout_name = str(spec.get("layout") or "content").lower()
    if layout_name in _LAYOUT_CATALOG:
        layout = presentation.slide_layouts[min(5, len(presentation.slide_layouts) - 1)]
        slide = presentation.slides.add_slide(layout)
        _apply_catalog_layout(slide, spec, _LAYOUT_CATALOG[layout_name], workspace, features)
        return slide
    layout_index = _LAYOUTS.get(layout_name, 1)
    layout = presentation.slide_layouts[min(layout_index, len(presentation.slide_layouts) - 1)]
    slide = presentation.slides.add_slide(layout)
    _populate_slide(slide, spec, workspace, features)
    return slide


def _resolve_template_spec(spec: Mapping[str, Any]) -> dict[str, Any]:
    template_name = str(spec.get("template") or "").strip().lower()
    if not template_name:
        return dict(spec)
    template = _TEMPLATE_CATALOG.get(template_name)
    if template is None:
        raise EngineError("TEMPLATE_NOT_FOUND", f"unknown PowerPoint template: {template_name}", {"allowed": sorted(_TEMPLATE_CATALOG)})
    resolved = dict(spec)
    if not resolved.get("slides"):
        resolved["slides"] = [dict(item) for item in template.get("slides") or []]
    if not resolved.get("theme"):
        resolved["theme"] = template.get("theme")
    resolved["_templateName"] = template_name
    return resolved


def _master_textbox(master: Any, text: Any, left: float, top: float, width: float, height: float, size: float = 10) -> None:
    """Append a plain text box to a slide master.

    ``MasterShapes`` exposes no ``add_textbox`` in python-pptx, so the ``p:sp``
    element is written directly into the master shape tree.
    """
    ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    emu = lambda value: int(round(float(value) * 914400))
    shape_id = len(master.shapes._spTree.findall(qn("p:sp"))) + 100
    safe_text = (
        scalar_text(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    xml = (
        f'<p:sp {ns}>'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="OfficeForgeMasterText{shape_id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{emu(left)}" y="{emu(top)}"/><a:ext cx="{emu(width)}" cy="{emu(height)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>'
        f'<a:p><a:r><a:rPr lang="en-US" sz="{int(size * 100)}"/><a:t>{safe_text}</a:t></a:r></a:p>'
        f'</p:txBody></p:sp>'
    )
    master.shapes._spTree.append(etree.fromstring(xml))


def _apply_master(presentation: Any, value: Any, features: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "master", "PowerPoint slide master")
    spec = require_dict(value, "master")
    master = presentation.slide_masters[0]
    if spec.get("background"):
        fill = master.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(str(spec["background"]).lstrip("#"))
    if spec.get("footer"):
        _master_textbox(master, spec["footer"], 0.6, 7.05, 12.0, 0.3, size=10)
    if spec.get("header"):
        _master_textbox(master, spec["header"], 0.6, 0.2, 12.0, 0.3, size=10)
    return {"background": spec.get("background"), "footer": bool(spec.get("footer")), "header": bool(spec.get("header"))}


def create_pptx(path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
    spec = _resolve_template_spec(spec)
    presentation = Presentation()
    _set_slide_size(presentation, spec)
    if spec.get("master") is not None:
        _apply_master(presentation, spec["master"], features)
    slides = require_list(spec.get("slides"), "slides", max_items=int(limits["maxSlides"]))
    if not slides:
        raise EngineError("INVALID_ARGUMENT", "slides must contain at least one slide")
    for slide_spec in slides:
        _add_slide(presentation, require_dict(slide_spec, "slide"), workspace, features)
    if spec.get("theme"):
        _apply_theme(presentation, spec["theme"], features)
    if spec.get("properties"):
        props = require_dict(spec["properties"], "properties")
        for key in ("title", "subject", "author", "keywords", "comments", "category"):
            if key in props:
                setattr(presentation.core_properties, key, scalar_text(props[key]))
    applied_properties = apply_core_properties(presentation.core_properties, spec.get("properties"))

    def writer(temp: str) -> None:
        presentation.save(temp)

    size, _ = atomic_save(path, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "pptx", "path": path, "bytes": size, "slides": len(slides), "template": spec.get("_templateName"), "properties": applied_properties}


# ---------------------------------------------------------------------------
# Editing
# ---------------------------------------------------------------------------


def _replace_in_text_frame(text_frame: Any, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    count = 0
    flags = 0 if match_case else re.IGNORECASE
    pattern = re.compile(find if regex else re.escape(find), flags)
    for paragraph in text_frame.paragraphs:
        original = "".join(run.text for run in paragraph.runs)
        if not original:
            continue
        if regex:
            new_text, number = pattern.subn(replacement, original)
        elif match_case:
            number = original.count(find)
            new_text = original.replace(find, replacement)
        else:
            new_text, number = pattern.subn(lambda _m: replacement, original)
        if number:
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                for run in paragraph.runs[1:]:
                    run._r.getparent().remove(run._r)
            else:
                paragraph.text = new_text
            count += number
    return count


def _replace_in_slide(slide: Any, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    count = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            count += _replace_in_text_frame(shape.text_frame, find, replacement, regex, match_case)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    count += _replace_in_text_frame(cell.text_frame, find, replacement, regex, match_case)
    try:
        if slide.has_notes_slide:
            count += _replace_in_text_frame(slide.notes_slide.notes_text_frame, find, replacement, regex, match_case)
    except Exception:
        pass
    return count


def _delete_slide(presentation: Any, index: int) -> None:
    slide_id = presentation.slides._sldIdLst[index]
    presentation.part.drop_rel(slide_id.rId)
    presentation.slides._sldIdLst.remove(slide_id)


def _apply_notes_batch(presentation: Any, values: Any, features: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "notes", "PowerPoint speaker notes")
    items = require_list(values, "notes", max_items=300)
    applied = 0
    cleared = 0
    for item in items:
        spec = require_dict(item, "note")
        index = int(spec.get("slide") or 0) - 1
        if index < 0 or index >= len(presentation.slides):
            raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
        text = scalar_text(spec.get("text") or "")
        frame = presentation.slides[index].notes_slide.notes_text_frame
        frame.text = text
        if text:
            applied += 1
        else:
            cleared += 1
    return {"slides": len(items), "applied": applied, "cleared": cleared}


def _sections_element(presentation: Any, create: bool = False) -> Any:
    """Return the ``p14:sectionLst`` element of the presentation extension list.

    PowerPoint stores slide sections in a ``p:extLst`` extension with the
    ``http://schemas.microsoft.com/office/powerpoint/2010/main`` namespace, so
    the element has to be created by hand; python-pptx has no public API.

    The extension list must hang off the ``p:presentation`` root element: putting
    it inside ``p:sldIdLst`` corrupts the slide id list and makes python-pptx
    raise ``AttributeError: 'lxml.etree._Element' object has no attribute 'rId'``
    the next time the presentation is opened.
    """
    root = presentation.part._element
    if root is None:
        return None
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        if not create:
            return None
        ext_lst = etree.SubElement(root, qn("p:extLst"))
    for ext in ext_lst.findall(qn("p:ext")):
        candidate = ext.find(f"{{{_P14_NS}}}sectionLst")
        if candidate is not None:
            return candidate
    if not create:
        return None
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
    section_lst = etree.SubElement(ext, f"{{{_P14_NS}}}sectionLst")
    return section_lst


def _section_slide_ids(presentation: Any, slide_numbers: Any) -> list[str]:
    ids: list[str] = []
    total = len(presentation.slides)
    for raw in require_list(slide_numbers, "sections.slides", max_items=300):
        number = int(raw)
        if number < 1 or number > total:
            raise EngineError("SLIDE_NOT_FOUND", f"slide index {number} is out of range")
        ids.append(str(presentation.slides[number - 1].slide_id))
    return ids


def _add_section(presentation: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "ppt", "sections", "PowerPoint sections")
    name = scalar_text(operation.get("name") or "").strip()
    if not name:
        raise EngineError("INVALID_ARGUMENT", "section.name must not be empty")
    ids = _section_slide_ids(presentation, operation.get("slides") or [])
    if not ids:
        raise EngineError("INVALID_ARGUMENT", "section.slides must list at least one slide")
    section_lst = _sections_element(presentation, create=True)
    if section_lst is None:
        raise EngineError("INVALID_OPERATION", "presentation has no slide id list to attach sections to")
    assigned = {
        node.get("id")
        for node in section_lst.iter(f"{{{_P14_NS}}}sldId")
        if node.get("id")
    }
    clash = [item for item in ids if item in assigned]
    if clash:
        raise EngineError("SECTION_CONFLICT", "slides already belong to a section", {"slides": clash})
    section = etree.SubElement(section_lst, f"{{{_P14_NS}}}section")
    section.set("name", name)
    section.set("id", _section_guid(name))
    for slide_id in ids:
        node = etree.SubElement(section, f"{{{_P14_NS}}}sldId")
        node.set("id", slide_id)
    return {"name": name, "slides": len(ids)}


def _section_guid(seed: str) -> str:
    digest = hashlib.sha1(seed.encode("utf-8")).hexdigest().upper()
    return f"{{{digest[:8]}-{digest[8:12]}-{digest[12:16]}-{digest[16:20]}-{digest[20:32]}}}"


def _list_sections(presentation: Any) -> list[dict[str, Any]]:
    section_lst = _sections_element(presentation)
    if section_lst is None:
        return []
    order = {str(slide.slide_id): index + 1 for index, slide in enumerate(presentation.slides)}
    result: list[dict[str, Any]] = []
    for section in section_lst.findall(f"{{{_P14_NS}}}section"):
        slides = [order.get(node.get("id"), 0) for node in section.iter(f"{{{_P14_NS}}}sldId")]
        result.append({"name": section.get("name") or "", "slides": [item for item in slides if item]})
    return result


def edit_pptx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None) -> dict[str, Any]:
    presentation = Presentation(path)
    operations = require_list(args.get("operations"), "operations", max_items=200)
    if not operations:
        raise EngineError("INVALID_ARGUMENT", "operations must contain at least one edit")
    applied: list[dict[str, Any]] = []
    for item in operations:
        operation = require_dict(item, "operation")
        name = str(operation.get("op") or "")
        if name == "replace":
            find = scalar_text(operation.get("find") or "")
            if not find:
                raise EngineError("INVALID_ARGUMENT", "replace.find must not be empty")
            scope = str(operation.get("scope") or "all")
            total = 0
            indices = range(len(presentation.slides)) if scope == "all" else [int(operation.get("slide") or 1) - 1]
            for slide_index in indices:
                if 0 <= slide_index < len(presentation.slides):
                    total += _replace_in_slide(presentation.slides[slide_index], find, scalar_text(operation.get("replace") or ""), bool(operation.get("regex", False)), bool(operation.get("matchCase", False)))
            applied.append({"op": "replace", "matches": total})
        elif name == "appendSlides":
            added = 0
            for slide_spec in require_list(operation.get("slides"), "slides", max_items=100):
                _add_slide(presentation, require_dict(slide_spec, "slide"), workspace, features)
                added += 1
            applied.append({"op": "appendSlides", "slides": added})
        elif name == "setTransition":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            _add_transition(presentation.slides[index], operation.get("transition") or operation, features)
            applied.append({"op": "setTransition", "slide": index + 1})
        elif name == "addShape":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            shapes = operation.get("shapes") if operation.get("shapes") is not None else [operation]
            for shape_spec in require_list(shapes, "shapes", max_items=100):
                _add_shape(presentation.slides[index], shape_spec, features)
            applied.append({"op": "addShape", "slide": index + 1, "count": len(shapes)})
        elif name == "addTextBox":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            textboxes = operation.get("textboxes") if operation.get("textboxes") is not None else [operation]
            for textbox in require_list(textboxes, "textboxes", max_items=100):
                _add_textbox(presentation.slides[index], textbox, features)
            applied.append({"op": "addTextBox", "slide": index + 1, "count": len(textboxes)})
        elif name == "addConnector":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            connectors = operation.get("connectors") if operation.get("connectors") is not None else [operation]
            for connector in require_list(connectors, "connectors", max_items=100):
                _add_connector(presentation.slides[index], connector, features)
            applied.append({"op": "addConnector", "slide": index + 1, "count": len(connectors)})
        elif name == "addAnimation":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            result = _add_animation(presentation.slides[index], operation.get("animation") if operation.get("animation") is not None else operation, features)
            applied.append({"op": "addAnimation", "slide": index + 1, **result})
        elif name == "addSmartArt":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            result = _add_smartart(presentation.slides[index], operation.get("smartArt") if operation.get("smartArt") is not None else operation, features)
            applied.append({"op": "addSmartArt", "slide": index + 1, **result})
        elif name == "setNotes":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            presentation.slides[index].notes_slide.notes_text_frame.text = scalar_text(operation.get("text") or "")
            applied.append({"op": "setNotes", "slide": index + 1})
        elif name == "setNotesBatch":
            result = _apply_notes_batch(presentation, operation.get("notes") if operation.get("notes") is not None else [operation], features)
            applied.append({"op": "setNotesBatch", **result})
        elif name == "addComment":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            result = _add_comment(presentation.slides[index], operation.get("comment") if operation.get("comment") is not None else operation, features, presentation)
            applied.append({"op": "addComment", "slide": index + 1, **result})
        elif name == "listComments":
            applied.append({"op": "listComments", "comments": _list_comments(presentation)})
        elif name == "addMedia":
            index = int(operation.get("slide") or 1) - 1
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            result = _add_media(presentation.slides[index], operation.get("media") if operation.get("media") is not None else operation, features)
            applied.append({"op": "addMedia", "slide": index + 1, **result})
        elif name == "applyMaster":
            result = _apply_master(presentation, operation.get("master") if operation.get("master") is not None else operation, features)
            applied.append({"op": "applyMaster", **result})
        elif name == "addSection":
            result = _add_section(presentation, operation, features)
            applied.append({"op": "addSection", **result})
        elif name == "listSections":
            applied.append({"op": "listSections", "sections": _list_sections(presentation)})
        elif name == "deleteSlide":
            index = int(operation.get("slide") or 1) - 1
            if len(presentation.slides) <= 1:
                raise EngineError("INVALID_OPERATION", "cannot delete the only slide")
            if index < 0 or index >= len(presentation.slides):
                raise EngineError("SLIDE_NOT_FOUND", f"slide index {index + 1} is out of range")
            _delete_slide(presentation, index)
            applied.append({"op": "deleteSlide", "slide": index + 1})
        elif name == "moveSlide":
            source = int(operation.get("from") or 1) - 1
            target = int(operation.get("to") or 1) - 1
            slides = presentation.slides._sldIdLst
            if source < 0 or source >= len(slides) or target < 0 or target >= len(slides):
                raise EngineError("SLIDE_NOT_FOUND", "moveSlide indices are out of range")
            element = slides[source]
            slides.remove(element)
            slides.insert(target, element)
            applied.append({"op": "moveSlide", "from": source + 1, "to": target + 1})
        elif name == "applyTheme":
            theme_name = _apply_theme(presentation, operation.get("theme") or operation, features)
            applied.append({"op": "applyTheme", "theme": theme_name})
        elif name == "setCore":
            values = require_dict(operation.get("values"), "setCore.values")
            mapping = {"title": "title", "subject": "subject", "author": "author", "keywords": "keywords", "comments": "comments", "category": "category", "lastModifiedBy": "last_modified_by"}
            for key, attribute in mapping.items():
                if key in values:
                    setattr(presentation.core_properties, attribute, scalar_text(values[key]))
            applied.append({"op": "setCore", "keys": sorted(values.keys())})
        else:
            raise EngineError("UNSUPPORTED_OPERATION", f"unsupported pptx edit operation: {name or '(empty)'}", {"allowed": ["replace", "appendSlides", "setNotes", "setNotesBatch", "addSection", "listSections", "deleteSlide", "moveSlide", "setTransition", "addShape", "addTextBox", "addConnector", "addAnimation", "addSmartArt", "addComment", "listComments", "addMedia", "applyMaster", "applyTheme", "setCore"]})

    size, backup = atomic_save(path, lambda temp: presentation.save(temp), overwrite=True, max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(args.get("backup", True)))
    return {"type": "pptx", "path": path, "bytes": size, "backupPath": backup, "operations": applied}


# ---------------------------------------------------------------------------
# Template expansion
# ---------------------------------------------------------------------------


def _replace_slide_text(slide: Any, data: Any, missing: str) -> int:
    count = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                original = "".join(run.text for run in paragraph.runs)
                if "{{" not in original:
                    continue
                replaced = render_placeholders(original, data, missing)
                if replaced != original:
                    if paragraph.runs:
                        paragraph.runs[0].text = replaced
                        for run in paragraph.runs[1:]:
                            run._r.getparent().remove(run._r)
                    else:
                        paragraph.text = replaced
                    count += 1
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        original = "".join(run.text for run in paragraph.runs)
                        if "{{" in original:
                            paragraph.runs[0].text = render_placeholders(original, data, missing)
                            for run in paragraph.runs[1:]:
                                run._r.getparent().remove(run._r)
                            count += 1
    try:
        if slide.has_notes_slide:
            for paragraph in slide.notes_slide.notes_text_frame.paragraphs:
                original = "".join(run.text for run in paragraph.runs)
                if "{{" in original:
                    paragraph.runs[0].text = render_placeholders(original, data, missing)
                    for run in paragraph.runs[1:]:
                        run._r.getparent().remove(run._r)
                    count += 1
    except Exception:
        pass
    return count


def template_pptx(template: str, output: str, data: Mapping[str, Any], options: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    presentation = Presentation(template)
    missing = str(options.get("missing") or "keep")
    replacements = sum(_replace_slide_text(slide, data, missing) for slide in presentation.slides)
    size, backup = atomic_save(output, lambda temp: presentation.save(temp), overwrite=bool(options.get("overwrite", False)), max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(options.get("backup", False)))
    return {"type": "pptx", "path": output, "bytes": size, "paragraphsReplaced": replacements, "backupPath": backup}


# ---------------------------------------------------------------------------
# Theme presets and deck quality checks (v0.4)
# ---------------------------------------------------------------------------


def _theme_name(value: Any) -> str:
    if isinstance(value, Mapping):
        value = value.get("name") or value.get("theme") or ""
    return str(value or "").strip().lower()


def _apply_theme(presentation: Any, value: Any, features: Any) -> str:
    require_feature(features, "ppt", "themes", "PowerPoint themes")
    name = _theme_name(value)
    if name not in _THEMES:
        raise EngineError("THEME_NOT_FOUND", f"unknown PowerPoint theme: {name or '(empty)'}", {"allowed": sorted(_THEMES)})
    theme = _THEMES[name]
    for slide in presentation.slides:
        try:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor.from_string(theme["background"].lstrip("#"))
        except Exception:
            pass
        title_shape = slide.shapes.title
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            is_title = title_shape is not None and shape == title_shape
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = theme["font"]
                    try:
                        run.font.color.rgb = RGBColor.from_string((theme["titleColor"] if is_title else theme["bodyColor"]).lstrip("#"))
                    except Exception:
                        pass
                    if is_title:
                        run.font.bold = True
    return name


def _deck_check(presentation: Any, args: Mapping[str, Any]) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []
    title_fonts: dict[str, int] = {}
    for index, slide in enumerate(presentation.slides, start=1):
        title_shape = slide.shapes.title
        title = title_shape.text.strip() if title_shape is not None and title_shape.has_text_frame else ""
        text_parts: list[str] = []
        bullets = 0
        small_fonts = 0
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if title_shape is None or shape != title_shape:
                text_parts.append(shape.text_frame.text)
                bullets += sum(1 for paragraph in shape.text_frame.paragraphs if paragraph.text.strip())
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.name:
                        if title_shape is not None and shape == title_shape:
                            title_fonts[run.font.name] = title_fonts.get(run.font.name, 0) + 1
                    if run.font.size is not None and run.font.size < Pt(12):
                        small_fonts += 1
        total_text = len(title) + sum(len(text) for text in text_parts)
        if not title and not any(text.strip() for text in text_parts):
            issues.append({"code": "empty_slide", "severity": "high", "slide": index, "message": "slide has no title or body text"})
        elif not title:
            issues.append({"code": "missing_title", "severity": "medium", "slide": index, "message": "slide has no title"})
        if len(title) > 60:
            issues.append({"code": "long_title", "severity": "low", "slide": index, "message": "title is longer than 60 characters"})
        if total_text > 800:
            issues.append({"code": "dense_text", "severity": "medium", "slide": index, "message": "slide carries more than 800 characters"})
        if bullets > 8:
            issues.append({"code": "too_many_bullets", "severity": "low", "slide": index, "message": "slide has more than 8 body paragraphs"})
        if small_fonts:
            issues.append({"code": "small_font", "severity": "medium", "slide": index, "message": f"slide has {small_fonts} run(s) below 12 pt"})
    if len(title_fonts) > 1:
        issues.append({"code": "inconsistent_title_fonts", "severity": "low", "slide": 0, "message": "title fonts vary across the deck", "fonts": sorted(title_fonts)})
    summary: dict[str, int] = {}
    for issue in issues:
        summary[issue["code"]] = summary.get(issue["code"], 0) + 1
    high = sum(1 for issue in issues if issue["severity"] == "high")
    medium = sum(1 for issue in issues if issue["severity"] == "medium")
    low = sum(1 for issue in issues if issue["severity"] == "low")
    score = max(0, 100 - high * 12 - medium * 5 - low * 2)
    return {
        "type": "pptx",
        "mode": "deck",
        "score": score,
        "issueCount": len(issues),
        "summary": summary,
        "issues": issues[: max(1, min(int(args.get("maxIssues") or 200), 1000))],
        "slideCount": len(presentation.slides),
        "truncated": len(issues) > max(1, min(int(args.get("maxIssues") or 200), 1000)),
    }


def audit_pptx(path: str, mode: str, args: Mapping[str, Any]) -> dict[str, Any]:
    if mode == "themes":
        return {"type": "pptx", "mode": "themes", "themes": [{"name": key, **value} for key, value in _THEMES.items()]}
    if mode == "layouts":
        return {"type": "pptx", "mode": "layouts", "layouts": [dict(item) for item in _LAYOUT_CATALOG.values()]}
    if mode == "templates":
        return {"type": "pptx", "mode": "templates", "templates": [dict(item) for item in _TEMPLATE_CATALOG.values()]}
    if mode == "profile":
        presentation = Presentation(path)
        return {
            "type": "pptx",
            "mode": "profile",
            "slideCount": len(presentation.slides),
            "slideSize": {"width": presentation.slide_width, "height": presentation.slide_height},
            "title": presentation.core_properties.title,
            "author": presentation.core_properties.author,
        }
    if mode == "deck":
        return _deck_check(Presentation(path), args)
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported pptx audit mode: {mode}", {"allowed": ["deck", "profile", "themes", "layouts", "templates"]})
