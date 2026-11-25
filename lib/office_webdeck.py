"""Self-contained HTML deck renderer.

The web deck is a sibling output of the native PPTX renderer: it consumes the
same high-level ``slides`` spec but emits one portable HTML file with keyboard
and click navigation.  It deliberately has no external runtime or CDN
dependency, so a generated deck opens on any machine with a browser.
"""

from __future__ import annotations

import base64
import html
import json
import os
from typing import Any, Mapping

from office_common import EngineError, atomic_save, require_dict, require_feature, require_list, scalar_text


def _theme_path() -> str:
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resources", "themes.json")


def _load_themes() -> dict[str, dict[str, Any]]:
    try:
        with open(_theme_path(), "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        themes = payload.get("themes") if isinstance(payload, dict) else None
        return {str(item["id"]): dict(item) for item in (themes or []) if isinstance(item, dict) and item.get("id")}
    except Exception:
        return {}


def _theme(name: Any) -> dict[str, Any]:
    themes = _load_themes()
    key = str(name or "forge").lower()
    theme = themes.get(key) or themes.get("forge") or {}
    return {
        "id": key,
        "background": theme.get("background", "0B1F33"),
        "panel": theme.get("panel", "132F4C"),
        "titleColor": theme.get("titleColor", "FFFFFF"),
        "bodyColor": theme.get("bodyColor", "D9E2EC"),
        "accent": theme.get("accent", "2D9CDB"),
        "accent2": theme.get("accent2", "56CCF2"),
        "font": theme.get("font", "Microsoft YaHei, Segoe UI, sans-serif"),
    }


def _esc(value: Any) -> str:
    return html.escape(scalar_text(value or ""))


def _image_src(value: Any, workspace: str) -> str:
    raw = str(value or "")
    if not raw:
        return ""
    if raw.startswith(("data:", "http://", "https://")):
        return raw
    path = raw if os.path.isabs(raw) else os.path.join(workspace, raw)
    path = os.path.abspath(path)
    if not path.startswith(os.path.abspath(workspace) + os.sep) and path != os.path.abspath(workspace):
        raise EngineError("PATH_OUTSIDE_WORKSPACE", "image path is outside the workspace", {"path": raw})
    try:
        with open(path, "rb") as handle:
            data = handle.read()
        mime = "image/png"
        if path.lower().endswith((".jpg", ".jpeg")):
            mime = "image/jpeg"
        elif path.lower().endswith(".gif"):
            mime = "image/gif"
        elif path.lower().endswith(".webp"):
            mime = "image/webp"
        return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
    except OSError:
        return ""


def _bullets_html(values: Any) -> str:
    if not isinstance(values, list):
        return ""
    return "<ul>" + "".join(f"<li>{_esc(item)}</li>" for item in values) + "</ul>"


def _paragraphs_html(values: Any) -> str:
    if not isinstance(values, list):
        return ""
    return "".join(f"<p>{_esc(item)}</p>" for item in values)


def _table_html(table: Mapping[str, Any]) -> str:
    columns = table.get("columns") or table.get("headers") or []
    rows = table.get("rows") or []
    head = "".join(f"<th>{_esc(item)}</th>" for item in columns)
    body = ""
    for row in rows if isinstance(rows, list) else []:
        if isinstance(row, list):
            body += "<tr>" + "".join(f"<td>{_esc(cell)}</td>" for cell in row) + "</tr>"
        elif isinstance(row, Mapping):
            body += "<tr>" + "".join(f"<td>{_esc(row.get(str(key), ''))}</td>" for key in columns) + "</tr>"
    return f"<table>{'<thead><tr>'+head+'</tr></thead>' if head else ''}<tbody>{body}</tbody></table>"


def _kpi_html(values: Any) -> str:
    if not isinstance(values, list):
        return ""
    cards = []
    for item in values:
        if isinstance(item, Mapping):
            cards.append(
                '<div class="kpi"><div class="kpi-value">'
                + _esc(item.get("value") or item.get("number") or "")
                + '</div><div class="kpi-label">'
                + _esc(item.get("label") or item.get("name") or "")
                + "</div></div>"
            )
    return '<div class="kpi-row">' + "".join(cards) + "</div>"


def _image_html(images: Any, workspace: str) -> str:
    if isinstance(images, str):
        images = [images]
    if not isinstance(images, list):
        return ""
    parts = []
    for item in images:
        if isinstance(item, Mapping):
            src = _image_src(item.get("path") or item.get("src"), workspace)
            alt = _esc(item.get("alt") or item.get("caption") or "")
        else:
            src = _image_src(item, workspace)
            alt = ""
        if src:
            parts.append(f'<img src="{src}" alt="{alt}">')
    return '<div class="image-row">' + "".join(parts) + "</div>" if parts else ""


def _chart_html(chart: Any) -> str:
    if not isinstance(chart, Mapping):
        return ""
    kind = str(chart.get("type") or "bar").lower()
    categories = chart.get("categories") or []
    series = chart.get("series") or []
    if not isinstance(series, list):
        return ""
    if series and isinstance(series[0], Mapping) and isinstance(series[0].get("values"), list):
        values = series[0].get("values") or []
        labels = categories or [str(i + 1) for i in range(len(values))]
        if kind == "pie":
            rows = [f"<tr><td>{_esc(label)}</td><td>{_esc(value)}</td></tr>" for label, value in zip(labels, values)]
            return '<table class="chart-table"><tbody>' + "".join(rows) + "</tbody></table>"
        max_value = max([abs(float(v or 0)) for v in values] or [1]) or 1
        bars = []
        for label, value in zip(labels, values):
            width = max(2, int(abs(float(value or 0)) / max_value * 100))
            bars.append(f'<div class="bar-row"><span>{_esc(label)}</span><div class="bar" style="width:{width}%"></div><b>{_esc(value)}</b></div>')
        return '<div class="chart">' + "".join(bars) + "</div>"
    return ""


def _shapes_html(shapes: Any) -> str:
    if not isinstance(shapes, list):
        return ""
    parts = []
    for shape in shapes:
        if not isinstance(shape, Mapping):
            continue
        fill = str(shape.get("fill") or "1F4E79").lstrip("#")
        style = f"background:#{_esc(fill)};border-radius:6px;padding:12px 16px;color:#fff;min-width:100px"
        text = _esc(shape.get("text") or shape.get("name") or "")
        parts.append(f'<div class="shape" style="{style}">{text}</div>')
    return '<div class="shape-row">' + "".join(parts) + "</div>" if parts else ""


def _media_src(path: Any, workspace: str) -> tuple[str, str]:
    raw = str(path or "")
    if not raw:
        return "", ""
    if raw.startswith(("data:", "http://", "https://")):
        return raw, "video/mp4"
    full = raw if os.path.isabs(raw) else os.path.join(workspace, raw)
    full = os.path.abspath(full)
    if not full.startswith(os.path.abspath(workspace) + os.sep) and full != os.path.abspath(workspace):
        return "", ""
    mime = "video/mp4" if full.lower().endswith((".mp4", ".webm", ".mov")) else "audio/mpeg"
    try:
        with open(full, "rb") as handle:
            data = handle.read()
        return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}", mime
    except OSError:
        return "", ""


def _media_html(media: Any, workspace: str) -> str:
    items = media if isinstance(media, list) else [media]
    parts = []
    for item in items:
        if not isinstance(item, Mapping):
            continue
        src, mime = _media_src(item.get("path") or item.get("source"), workspace)
        if not src:
            continue
        if str(item.get("kind") or "video").lower() == "audio" or mime.startswith("audio"):
            parts.append(f'<audio controls src="{src}"></audio>')
        else:
            parts.append(f'<video controls src="{src}"></video>')
    return '<div class="media-row">' + "".join(parts) + "</div>" if parts else ""


def _comments_html(comments: Any) -> str:
    items = comments if isinstance(comments, list) else [comments]
    lines = [f"<li>{_esc(item.get('text') or item)}</li>" for item in items if isinstance(item, Mapping)]
    return '<aside class="comments"><b>Comments</b><ul>' + "".join(lines) + "</ul></aside>" if lines else ""


def _slide_html(slide: Mapping[str, Any], index: int, total: int, workspace: str) -> str:
    layout = str(slide.get("layout") or "content").lower()
    title = slide.get("title") or slide.get("heading") or ""
    subtitle = slide.get("subtitle") or ""
    parts = [f'<section class="slide slide-{_esc(layout)}" data-slide="{index}">']
    if layout in ("title", "cover", "title_slide"):
        parts.append(f'<div class="title-wrap"><h1>{_esc(title)}</h1>')
        if subtitle:
            parts.append(f'<p class="subtitle">{_esc(subtitle)}</p>')
        parts.append("</div>")
    elif layout in ("section", "section_header"):
        parts.append(f'<div class="section-wrap"><div class="section-number">{index:02d}</div><h2>{_esc(title)}</h2></div>')
    else:
        if title:
            parts.append(f"<h2>{_esc(title)}</h2>")
        if subtitle:
            parts.append(f'<p class="subtitle">{_esc(subtitle)}</p>')
        body = []
        if slide.get("kpis") is not None:
            body.append(_kpi_html(slide.get("kpis")))
        if slide.get("bullets") is not None:
            body.append(_bullets_html(slide.get("bullets")))
        if slide.get("paragraphs") is not None:
            body.append(_paragraphs_html(slide.get("paragraphs")))
        if slide.get("table") is not None:
            body.append(_table_html(require_dict(slide["table"], "table")))
        if slide.get("images") is not None:
            body.append(_image_html(slide.get("images"), workspace))
        if slide.get("image") is not None:
            body.append(_image_html(slide.get("image"), workspace))
        if slide.get("chart") is not None:
            body.append(_chart_html(slide.get("chart")))
        if slide.get("shapes") is not None:
            body.append(_shapes_html(slide.get("shapes")))
        if slide.get("media") is not None:
            body.append(_media_html(slide.get("media"), workspace))
        if slide.get("comments") is not None:
            body.append(_comments_html(slide.get("comments")))
        parts.append('<div class="body">' + "".join(body) + "</div>")
    notes = slide.get("notes")
    if notes:
        parts.append(f'<aside class="notes">{_esc(notes)}</aside>')
    parts.append("</section>")
    return "".join(parts)


def create_html_deck(path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None, engine: str = "webdeck") -> dict[str, Any]:
    require_feature(features, "ppt", "web", "HTML deck output")
    if str(engine or "webdeck").lower() in {"guizang", "guizang-html", "guizang-swiss"}:
        import office_guizang
        return office_guizang.render_html(path, spec, limits, overwrite, workspace, features)
    slides = require_list(spec.get("slides"), "slides", max_items=int(limits["maxSlides"]))
    if not slides:
        raise EngineError("INVALID_ARGUMENT", "slides must contain at least one slide")
    theme = _theme(spec.get("theme"))
    slide_html = []
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, Mapping):
            raise EngineError("INVALID_ARGUMENT", "each slide must be an object")
        slide_html.append(_slide_html(slide, index, len(slides), workspace))
    title = _esc(spec.get("title") or "Office Forge deck")
    doc = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<style>
:root{{--bg:#{theme['background']};--panel:#{theme['panel']};--title:#{theme['titleColor']};--body:#{theme['bodyColor']};--accent:#{theme['accent']};--accent2:#{theme['accent2']};--font:{theme['font']};}}
*{{box-sizing:border-box}}
html,body{{margin:0;height:100%;background:#0a0a0a;color:var(--body);font-family:var(--font)}}
body{{display:grid;place-items:center;overflow:hidden}}
#deck{{position:relative;width:min(96vw,1280px);aspect-ratio:16/9;background:var(--bg);box-shadow:0 20px 80px #0008;overflow:hidden}}
.slide{{position:absolute;inset:0;padding:7% 8%;display:none;background:var(--bg);flex-direction:column}}
.slide.active{{display:flex;animation:fade .25s ease}}
@keyframes fade{{from{{opacity:.3;transform:translateX(12px)}}to{{opacity:1;transform:none}}}}
h1,h2{{color:var(--title);margin:0 0 .6em;line-height:1.15}}
h1{{font-size:clamp(32px,5.5vw,72px);letter-spacing:-.03em}}
h2{{font-size:clamp(24px,3.4vw,44px)}}
.subtitle{{font-size:clamp(16px,1.8vw,24px);opacity:.8;margin:.4em 0 1em}}
.body{{flex:1;display:flex;flex-direction:column;justify-content:center;gap:1.2em}}
ul{{margin:0;padding-left:1.2em}} li{{font-size:clamp(15px,1.7vw,22px);margin:.35em 0;line-height:1.45}}
p{{font-size:clamp(15px,1.7vw,22px);line-height:1.55;margin:.3em 0}}
table{{width:100%;border-collapse:collapse;font-size:clamp(13px,1.3vw,18px)}} th,td{{border:1px solid #ffffff33;padding:.55em .7em;text-align:left}} th{{background:#ffffff14;color:var(--title)}}
.title-wrap,.section-wrap{{margin:auto;max-width:85%}}
.title-wrap{{border-left:6px solid var(--accent);padding-left:1em}}
.section-number{{font-size:clamp(60px,10vw,160px);font-weight:800;color:var(--accent);opacity:.35;line-height:1}}
.kpi-row{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:1em}}
.kpi{{background:var(--panel);border-top:3px solid var(--accent);padding:1em}}
.kpi-value{{font-size:clamp(28px,4vw,56px);font-weight:800;color:var(--title)}} .kpi-label{{font-size:clamp(12px,1.2vw,16px);opacity:.7;margin-top:.3em}}
.image-row{{display:flex;gap:1em;justify-content:center;align-items:center;flex:1}} .image-row img{{max-width:100%;max-height:100%;object-fit:contain;border-radius:8px}}
.chart{{display:flex;flex-direction:column;gap:8px}} .bar-row{{display:grid;grid-template-columns:8em 1fr 4em;gap:8px;align-items:center}} .bar{{height:12px;background:var(--accent);border-radius:6px}} .bar-row b{{text-align:right;color:var(--title)}}
.shape-row{{display:flex;gap:12px;flex-wrap:wrap;justify-content:center}} .shape{{min-height:54px;display:flex;align-items:center;justify-content:center;text-align:center}} .media-row{{display:flex;justify-content:center}} .media-row video,.media-row audio{{max-width:100%;max-height:60%}} .comments{{font-size:12px;opacity:.7;border-left:2px solid var(--accent);padding-left:10px}}
.notes{{display:none}}
#bar{{position:absolute;left:0;bottom:0;height:4px;background:var(--accent);width:0;transition:width .2s}}
#ui{{position:absolute;right:18px;bottom:16px;display:flex;gap:8px;align-items:center;font-size:13px;color:#fff9}}
button{{border:1px solid #ffffff44;background:#0008;color:#fff;border-radius:8px;padding:7px 12px;cursor:pointer;font:inherit}} button:hover{{background:var(--accent);border-color:var(--accent)}}
@media print{{#ui,#bar{{display:none}} body{{display:block;overflow:visible;background:#fff}} #deck{{width:100%;box-shadow:none}} .slide{{display:flex;position:relative;page-break-after:always;aspect-ratio:16/9}}}}
</style>
</head>
<body><main id="deck">{''.join(slide_html)}<div id="bar"></div><div id="ui"><button id="prev">←</button><span id="count">1 / {len(slides)}</span><button id="next">→</button></div></main>
<script>
const slides=[...document.querySelectorAll('.slide')];let i=0;
function show(n){{i=Math.max(0,Math.min(slides.length-1,n));slides.forEach((s,k)=>s.classList.toggle('active',k===i));document.getElementById('count').textContent=(i+1)+' / '+slides.length;document.getElementById('bar').style.width=((i+1)/slides.length*100)+'%';}}
document.getElementById('prev').onclick=()=>show(i-1);document.getElementById('next').onclick=()=>show(i+1);
addEventListener('keydown',e=>{{if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown')show(i+1);if(e.key==='ArrowLeft'||e.key==='PageUp')show(i-1);}});
show(0);
</script></body></html>"""
    def writer(temp: str) -> None:
        with open(temp, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(doc)

    size, _ = atomic_save(path, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {"type": "html", "path": path, "bytes": size, "slides": len(slides), "theme": theme["id"], "engine": "webdeck"}
