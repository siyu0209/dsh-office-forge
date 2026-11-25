"""Render a Deck IR spec through the bundled Guizang HTML skill template.

The Guizang skill is HTML-first: its visual system lives in a self-contained
template (CSS + motion runtime), not in a native PowerPoint theme.  This
module uses the real template shipped inside ``skills/guizang-ppt-skill`` and
replaces only the slide sections and speaker notes.  Output is a portable,
single-file HTML deck.
"""

from __future__ import annotations

import html
import json
import os
import re
from typing import Any, Mapping

from office_common import EngineError, atomic_save, require_list, scalar_text


def _plugin_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _template_path() -> str:
    return os.path.join(_plugin_root(), "skills", "guizang-ppt-skill", "assets", "template-swiss.html")


def _esc(value: Any) -> str:
    return html.escape(scalar_text(value or ""))


def _bullet_list(values: Any) -> str:
    if not isinstance(values, list) or not values:
        return ""
    items = "".join(f"<li>{_esc(item)}</li>" for item in values)
    return f'<ul class="guizang-bullets" style="font-size:clamp(14px,1.5vw,20px);line-height:1.6;padding-left:1.2em;margin:2vh 0 0">{items}</ul>'


def _paragraphs(values: Any) -> str:
    if not isinstance(values, list) or not values:
        return ""
    return "".join(
        f'<p style="font-size:clamp(13px,1.35vw,18px);line-height:1.7;color:var(--text-helper);margin:1.2vh 0">{_esc(item)}</p>'
        for item in values
    )


def _table(table: Any) -> str:
    if not isinstance(table, Mapping):
        return ""
    headers = table.get("headers") or table.get("columns") or []
    rows = table.get("rows") or []
    if not isinstance(headers, list) or not isinstance(rows, list):
        return ""
    head = "".join(f"<th>{_esc(item)}</th>" for item in headers)
    body = "".join("<tr>" + "".join(f"<td>{_esc(cell)}</td>" for cell in row) + "</tr>" for row in rows if isinstance(row, list))
    return f'<table class="guizang-table" style="width:100%;border-collapse:collapse;font-size:clamp(11px,1.15vw,15px);margin-top:2vh"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def _slide_section(slide: Mapping[str, Any], index: int, total: int) -> str:
    title = _esc(slide.get("title") or slide.get("heading") or f"Slide {index}")
    subtitle = _esc(slide.get("subtitle") or "")
    bullets = _bullet_list(slide.get("bullets"))
    paragraphs = _paragraphs(slide.get("paragraphs"))
    table = _table(slide.get("table"))
    notes = _esc(slide.get("notes") or "")
    accent = index == 1
    cls = "slide accent" if accent else "slide"
    kicker = "OPENING" if index == 1 else f"SECTION {index:02d}"
    body = bullets or paragraphs or table
    if not body:
        body = "<p style=\"font-size:clamp(14px,1.5vw,20px);color:var(--text-helper)\">—</p>"
    return (
        f'<section class="{cls}" data-animate="grid-reveal" data-slide-id="slide-{index:02d}" data-title="{title}">'
        '<div class="canvas-card">'
        '<div class="chrome-min"><div class="l">' + _esc("Office Forge · Guizang") + '</div>'
        f'<div class="r">{index:02d} / {total:02d}</div></div>'
        f'<div class="t-meta" style="letter-spacing:.22em;margin-top:4vh">{kicker}</div>'
        f'<h2 class="h-xl" style="font-size:min(7vw,12vh);line-height:1.05;font-weight:300;margin:1.6vh 0 0">{title}</h2>'
        + (f'<div class="lead" style="max-width:60ch;margin-top:1.4vh">{subtitle}</div>' if subtitle else "")
        + body
        + (f'<aside class="notes" style="display:none">{notes}</aside>' if notes else "")
        + '</div></section>'
    )


def render_html(path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
    """Render ``spec`` with the real Guizang skill HTML template."""
    slides = require_list(spec.get("slides"), "slides", max_items=int(limits["maxSlides"]))
    if not slides:
        raise EngineError("INVALID_ARGUMENT", "slides must contain at least one slide")
    template_path = _template_path()
    try:
        with open(template_path, "r", encoding="utf-8") as handle:
            template = handle.read()
    except OSError as exc:
        raise EngineError("BACKEND_UNAVAILABLE", "Guizang template is missing", {"path": template_path}) from exc

    sections = "\n".join(_slide_section(slide, i, len(slides)) for i, slide in enumerate(slides, 1))
    deck_pattern = re.compile(r'(<div id="deck">)(.*?)(</div>\s*<div id="nav">)', re.S)
    if not deck_pattern.search(template):
        raise EngineError("BACKEND_UNAVAILABLE", "Guizang template deck marker not found", {"path": template_path})
    document = deck_pattern.sub(lambda match: match.group(1) + "\n" + sections + "\n" + match.group(3), template, count=1)

    notes = []
    for i, slide in enumerate(slides, 1):
        notes.append({
            "id": f"slide-{i:02d}",
            "title": str(slide.get("title") or f"Slide {i}"),
            "section": "开场" if i == 1 else "内容",
            "minutes": 1.0,
            "purpose": str(slide.get("subtitle") or ""),
            "talk": [str(item) for item in (slide.get("bullets") or [])][:6],
            "transition": "进入下一页",
            "cue": "",
        })
    note_pattern = re.compile(r"const SPEAKER_NOTES = \[.*?\];", re.S)
    document = note_pattern.sub("const SPEAKER_NOTES = " + json.dumps(notes, ensure_ascii=False) + ";", document, count=1)

    def writer(temp: str) -> None:
        with open(temp, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(document)

    size, _ = atomic_save(path, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {
        "type": "html",
        "path": path,
        "bytes": size,
        "slides": len(slides),
        "theme": "guizang-swiss",
        "engine": "guizang",
        "template": os.path.basename(template_path),
    }
