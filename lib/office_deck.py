"""Deck Runtime: Deck IR + adapter registry + template packs.

This module is the narrow orchestrator used by the ``officeforge_deck`` tool.
It never talks to an LLM and never registers prompt hooks; it only plans,
renders, audits and repairs a deck through registered adapters.
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Mapping

import office_deck_ir
import office_deck_registry
import office_ppt_import
import office_webdeck
from office_common import EngineError


def plan(spec: Any) -> dict[str, Any]:
    deck = office_deck_ir.normalize_deck(spec)
    return {
        "type": "deck",
        "mode": "plan",
        "schemaVersion": office_deck_ir.DECK_SCHEMA_VERSION,
        "deck": deck,
        "summary": office_deck_ir.summarize_deck(deck),
    }


def storyboard(spec: Any, fmt: str = "json") -> dict[str, Any]:
    """Return a presenter-oriented storyboard without touching any file."""
    deck = office_deck_ir.normalize_deck(spec)
    slides: list[dict[str, Any]] = []
    total_seconds = 0
    for index, slide in enumerate(deck.get("slides", [])):
        blocks = slide.get("blocks") or []
        block_types = [str(block.get("type") or "") for block in blocks]
        title = slide.get("title")
        subtitle = slide.get("subtitle")
        for block in blocks:
            block_type = block.get("type")
            if not title and block_type == "title":
                title = block.get("text")
            if not subtitle and block_type == "subtitle":
                subtitle = block.get("text")
        words = 0
        for block in blocks:
            for key in ("text", "items", "rows", "data"):
                value = block.get(key)
                if isinstance(value, list):
                    for item in value:
                        if isinstance(item, (list, tuple)):
                            words += len(" ".join(str(part) for part in item).split())
                        else:
                            words += len(str(item).split())
                elif value:
                    words += len(str(value).split())
        notes = str(slide.get("notes") or "")
        words += len(notes.split())
        seconds = max(20, min(180, 15 + int(words * 0.25)))
        total_seconds += seconds
        slides.append({
            "index": index + 1,
            "layout": slide.get("layout") or "content",
            "title": str(title or ""),
            "subtitle": str(subtitle or ""),
            "blockTypes": block_types,
            "notes": notes,
            "estimatedSeconds": seconds,
        })
    result = {
        "type": "deck",
        "mode": "storyboard",
        "schemaVersion": office_deck_ir.DECK_SCHEMA_VERSION,
        "title": str(deck.get("title") or ""),
        "theme": str(deck.get("theme") or ""),
        "slideCount": len(slides),
        "slides": slides,
        "totalEstimatedSeconds": total_seconds,
    }
    if str(fmt or "json").lower() == "markdown":
        lines = [
            f"# {result['title'] or 'Deck storyboard'}",
            "",
            f"- Theme: {result['theme'] or '-'}",
            f"- Slides: {result['slideCount']}",
            f"- Estimated: {result['totalEstimatedSeconds']}s",
            "",
        ]
        for slide in slides:
            lines.append(f"## {slide['index']}. {slide['title'] or '(untitled)'}")
            lines.append("")
            lines.append(f"- Layout: {slide['layout']}")
            if slide["subtitle"]:
                lines.append(f"- Subtitle: {slide['subtitle']}")
            if slide["blockTypes"]:
                lines.append(f"- Blocks: {', '.join(slide['blockTypes'])}")
            lines.append(f"- Estimated: {slide['estimatedSeconds']}s")
            if slide["notes"]:
                lines.append("")
                lines.append(f"> Notes: {slide['notes']}")
            lines.append("")
        result["format"] = "markdown"
        result["markdown"] = "\n".join(lines).rstrip() + "\n"
    else:
        result["format"] = "json"
    return result


def render(path: str, spec: Any, limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    deck = office_deck_ir.normalize_deck(spec)
    if not deck.get("slides") and not deck.get("template"):
        raise EngineError("DECK_EMPTY", "deck must contain at least one slide")
    if str(path).lower().endswith((".html", ".htm")):
        html_engine = "guizang" if str(engine or "auto").lower() == "auto" else engine
        if str(html_engine).lower() in {"guizang", "guizang-html", "guizang-swiss"}:
            result = render_html(path, spec, limits, overwrite, workspace, features, html_engine)
        else:
            result = office_deck_registry.render(path, deck, limits, overwrite, workspace, features, html_engine)
        result["summary"] = office_deck_ir.summarize_deck(deck)
        return result
    result = office_deck_registry.render(path, deck, limits, overwrite, workspace, features, engine)
    result["summary"] = office_deck_ir.summarize_deck(deck)
    return result


def audit(path: str, mode: str, args: Mapping[str, Any], features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    return office_deck_registry.audit(path, mode, args, features, engine)


def repair(path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    return office_deck_registry.repair(path, args, limits, workspace, features, engine)


def catalog() -> dict[str, Any]:
    result = office_deck_ir.catalog()
    result["adapters"] = office_deck_registry.adapters()
    result["deckSchema"] = "resources/deck_schema.json"
    return result


def capabilities() -> dict[str, Any]:
    """Return the full adapter and dsh-ppt catalog matrix without rendering."""
    resource = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "resources",
        "dsh_ppt_catalog.json",
    )
    try:
        with open(resource, "r", encoding="utf-8") as handle:
            dsh_ppt = json.load(handle)
    except (OSError, ValueError) as error:
        raise EngineError("CAPABILITY_CATALOG_UNAVAILABLE", f"could not read dsh-ppt catalog: {error}")
    return {
        "type": "deck",
        "mode": "capabilities",
        "schemaVersion": office_deck_ir.DECK_SCHEMA_VERSION,
        "adapters": office_deck_registry.adapters(),
        "dshPpt": dsh_ppt,
        "legacy": office_deck_ir.catalog(),
        "design": office_deck_registry.design_catalog(),
    }


def design_catalog() -> dict[str, Any]:
    return office_deck_registry.design_catalog()


def import_pptx(path: str, workspace: str, limits: Mapping[str, Any], args: Mapping[str, Any] | None = None) -> dict[str, Any]:
    return office_ppt_import.import_pptx(path, workspace, limits, args)


def edit_deck(source: str | None, output: str, spec: Any, operations: Any, limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    """Import or accept a deck, apply design-aware Deck IR edits, then render."""
    if spec is None:
        if not source:
            raise EngineError("INVALID_ARGUMENT", "deck edit requires path or spec")
        imported = office_ppt_import.import_pptx(source, workspace, limits, {})
        deck = imported["deck"]
        base_source = source
    else:
        deck = office_deck_ir.normalize_deck(spec)
        base_source = None
    edited = office_deck_ir.apply_deck_operations(deck, operations)
    result = render(output, edited, limits, overwrite, workspace, features, engine)
    result["operationsApplied"] = len(operations) if isinstance(operations, list) else 0
    if base_source:
        result["source"] = base_source
    return result


def loss_audit(spec: Any, engine: Any = "auto") -> dict[str, Any]:
    """Static loss audit; see office_deck_registry.loss_audit."""
    deck = office_deck_ir.normalize_deck(spec)
    result = office_deck_registry.loss_audit(deck, engine)
    result["summary"] = office_deck_ir.summarize_deck(deck)
    return result


def validate_template_pack(pack_id: str | None = None) -> dict[str, Any]:
    return office_deck_ir.validate_template_pack(pack_id)


def render_html(path: str, spec: Any, limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None, engine: Any = "webdeck") -> dict[str, Any]:
    deck = office_deck_ir.normalize_deck(spec)
    result = office_webdeck.create_html_deck(path, office_deck_ir.deck_to_backend_spec(deck), limits, overwrite, workspace, features, engine)
    result["deckSchemaVersion"] = deck.get("schemaVersion", office_deck_ir.DECK_SCHEMA_VERSION)
    result["summary"] = office_deck_ir.summarize_deck(deck)
    return result


def _verify_norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _verify_image(path: str) -> dict[str, Any]:
    from PIL import Image, ImageStat

    image = Image.open(path)
    image.load()
    width, height = image.size
    gray = image.convert("L")
    extrema = gray.getextrema()
    variance = float(ImageStat.Stat(gray).var[0])
    thumb = gray.resize((64, 64))
    pixels = list(thumb.getdata())
    nonwhite = sum(1 for value in pixels if value < 245) / max(1, len(pixels))
    blank = (extrema[1] - extrema[0]) < 6 and nonwhite < 0.01
    return {"width": width, "height": height, "variance": round(variance, 2), "nonWhiteRatio": round(nonwhite, 4), "blank": blank}


def verify(path: str, spec: Any, args: Mapping[str, Any]) -> dict[str, Any]:
    """Verify a rendered deck against its PPTX structure and rendered PNG pages."""
    from pptx import Presentation

    payload = dict(args or {})
    rendered_images = payload.get("images") if isinstance(payload.get("images"), list) else []
    missing_fonts = payload.get("missingFonts") if isinstance(payload.get("missingFonts"), list) else []
    rendered_page_count = payload.get("pageCount")
    deck = office_deck_ir.normalize_deck(spec) if spec else None
    expected_slides = list(deck.get("slides") or []) if deck else []
    presentation = Presentation(path)
    slides = list(presentation.slides)
    slide_width = int(presentation.slide_width or 0)
    slide_height = int(presentation.slide_height or 0)

    issues: list[dict[str, Any]] = []
    checks: list[dict[str, Any]] = []
    slide_rows: list[dict[str, Any]] = []

    def add(code: str, severity: str, message: str, **extra: Any) -> None:
        issue = {"code": code, "severity": severity, "message": message}
        issue.update(extra)
        issues.append(issue)

    if deck is not None and len(expected_slides) != len(slides):
        add("SLIDE_COUNT_MISMATCH", "high", "rendered deck slide count does not match the Deck IR", expected=len(expected_slides), actual=len(slides))
    checks.append({"name": "slideCount", "ok": deck is None or len(expected_slides) == len(slides), "detail": f"{len(slides)} slides"})

    if rendered_page_count is not None and int(rendered_page_count) != len(slides):
        add("PAGE_COUNT_MISMATCH", "high", "renderer reported a different page count", expected=len(slides), reported=int(rendered_page_count))

    image_by_page: dict[int, dict[str, Any]] = {}
    for index, item in enumerate(rendered_images):
        if not isinstance(item, Mapping):
            continue
        page = item.get("page")
        page_number = int(page) if isinstance(page, (int, float)) and not isinstance(page, bool) else index + 1
        if page_number not in image_by_page:
            image_by_page[page_number] = dict(item)

    for index, slide in enumerate(slides):
        page_number = index + 1
        texts: list[str] = []
        kinds: list[str] = []
        shape_count = 0
        for shape in slide.shapes:
            shape_count += 1
            if getattr(shape, "has_text_frame", False):
                value = str(shape.text or "").strip()
                if value:
                    texts.append(value)
            if getattr(shape, "has_table", False):
                kinds.append("table")
            if getattr(shape, "has_chart", False):
                kinds.append("chart")
            try:
                if int(shape.shape_type) == 13:
                    kinds.append("picture")
            except Exception:
                pass
            try:
                left = int(shape.left or 0)
                top = int(shape.top or 0)
                width = int(shape.width or 0)
                height = int(shape.height or 0)
                if slide_width and slide_height and (left < 0 or top < 0 or left + width > slide_width or top + height > slide_height):
                    add("SHAPE_OUT_OF_BOUNDS", "medium", "shape extends beyond the slide bounds", slide=page_number, shape=str(getattr(shape, "name", "")))
            except Exception:
                pass
        if not texts and not kinds:
            add("EMPTY_SLIDE", "high", "slide has no text, table, chart or picture", slide=page_number)
        expected_title = ""
        if index < len(expected_slides):
            expected_title = str(expected_slides[index].get("title") or "").strip()
        joined = _verify_norm(" ".join(texts))
        if expected_title and _verify_norm(expected_title) not in joined:
            add("TITLE_MISSING", "medium", "expected slide title was not found in the rendered text", slide=page_number, expected=expected_title)
        image = image_by_page.get(page_number)
        image_info = None
        if image is None:
            add("RENDER_PAGE_MISSING", "high", "no rendered PNG was produced for this slide", slide=page_number)
        else:
            try:
                image_info = _verify_image(str(image.get("path") or ""))
                if image_info["blank"]:
                    add("BLANK_RENDER", "high", "rendered slide appears blank", slide=page_number, path=str(image.get("path") or ""))
                if image_info["width"] < 120 or image_info["height"] < 80:
                    add("RENDER_TOO_SMALL", "medium", "rendered slide is too small to inspect", slide=page_number, width=image_info["width"], height=image_info["height"])
            except Exception as error:
                add("RENDER_UNREADABLE", "high", "rendered PNG could not be decoded", slide=page_number, detail=str(error))
        slide_rows.append({"index": page_number, "shapeCount": shape_count, "textLength": len(joined), "kinds": sorted(set(kinds)), "image": image_info})

    checks.append({"name": "renderPages", "ok": len(image_by_page) >= len(slides), "detail": f"{len(image_by_page)}/{len(slides)} pages rendered"})
    checks.append({"name": "textCoverage", "ok": not any(issue["code"] == "TITLE_MISSING" for issue in issues), "detail": f"{len(expected_slides)} expected titles" if deck else "no Deck IR supplied"})
    checks.append({"name": "geometry", "ok": not any(issue["code"] == "SHAPE_OUT_OF_BOUNDS" for issue in issues), "detail": "all shapes inside slide bounds" if not any(issue["code"] == "SHAPE_OUT_OF_BOUNDS" for issue in issues) else "out-of-bounds shapes found"})
    if missing_fonts:
        add("MISSING_FONTS", "low", "some fonts were substituted by the renderer", fonts=missing_fonts)

    weights = {"high": 15, "medium": 8, "low": 3}
    score = max(0, min(100, 100 - sum(weights.get(issue["severity"], 3) for issue in issues)))
    summary: dict[str, int] = {}
    for issue in issues:
        summary[issue["code"]] = summary.get(issue["code"], 0) + 1
    return {
        "type": "deck",
        "mode": "verify",
        "path": path,
        "pageCount": len(slides),
        "renderedPages": len(image_by_page),
        "missingFonts": missing_fonts,
        "score": score,
        "issueCount": len(issues),
        "summary": summary,
        "checks": checks,
        "slides": slide_rows,
        "issues": issues,
        "valid": score >= 70 and not any(issue["severity"] == "high" for issue in issues),
    }
