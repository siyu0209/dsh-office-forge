"""Office Forge Deck IR.

Deck IR is the stable, backend-neutral representation of a presentation.
Renderers, auditors and template packs exchange this object instead of each
other's private structures.  The schema is versioned so future adapters can
migrate old decks without changing the public ``officeforge_*`` tools.
"""

from __future__ import annotations

import copy
import json
import os
from typing import Any, Mapping

import office_pptx
from office_common import EngineError

DECK_SCHEMA_VERSION = 1
DEFAULT_THEME = "forge"
TEMPLATE_PACK_ENV = "OFFICEFORGE_TEMPLATE_PACKS"
DECK_BLOCK_TYPES = (
    "title",
    "subtitle",
    "text",
    "bullets",
    "paragraphs",
    "paragraph",
    "kpi",
    "kpis",
    "metrics",
    "stats",
    "table",
    "chart",
    "image",
    "images",
    "quote",
    "note",
    "notes",
    "timeline",
    "matrix",
    "shape",
    "shapes",
    "line",
    "lines",
    "media",
    "comment",
    "comments",
    "smartart",
    "transition",
    "animation",
)


def resource_path(name: str) -> str:
    """Return an absolute path inside this plugin's ``resources`` directory."""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resources", name)


def _read_json(name: str) -> dict[str, Any]:
    path = resource_path(name)
    try:
        with open(path, "r", encoding="utf-8") as handle:
            value = json.load(handle)
    except FileNotFoundError:
        raise EngineError("RESOURCE_NOT_FOUND", f"deck resource is missing: {name}", {"path": path})
    except json.JSONDecodeError as error:
        raise EngineError("RESOURCE_INVALID", f"deck resource is not valid JSON: {name}", {"error": str(error)})
    if not isinstance(value, dict):
        raise EngineError("RESOURCE_INVALID", f"deck resource must be an object: {name}")
    return value


def _plugin_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read_json_path(path: str) -> Any:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except FileNotFoundError:
        raise EngineError("RESOURCE_NOT_FOUND", f"template pack asset is missing: {path}", {"path": path})
    except json.JSONDecodeError as error:
        raise EngineError("RESOURCE_INVALID", f"template pack asset is not valid JSON: {path}", {"path": path, "error": str(error)})


def _pack_root(pack: Mapping[str, Any]) -> str:
    root = str(pack.get("root") or "").strip()
    if root:
        return root if os.path.isabs(root) else os.path.join(_plugin_root(), root)
    return _plugin_root()


def _pack_asset_path(pack: Mapping[str, Any], key: str) -> str | None:
    value = pack.get(key)
    if not value:
        return None
    path = str(value)
    return path if os.path.isabs(path) else os.path.join(_pack_root(pack), path)


def _asset_count(value: Any, key: str) -> int:
    if isinstance(value, list):
        return len(value)
    if isinstance(value, Mapping):
        inner = value.get(key)
        if isinstance(inner, list):
            return len(inner)
        if isinstance(inner, Mapping):
            return len(inner)
    return 0


def _pack_assets(pack: Mapping[str, Any]) -> dict[str, Any]:
    assets: dict[str, Any] = {}
    for key in ("themes", "layouts", "templates"):
        path = _pack_asset_path(pack, key)
        if path and os.path.isfile(path):
            assets[key] = _read_json_path(path)
        else:
            assets[key] = {} if key == "themes" else []
    return assets


def _external_pack_entries() -> list[dict[str, Any]]:
    """Read optional pack manifests from ``OFFICEFORGE_TEMPLATE_PACKS``.

    Each entry is either a directory containing ``template_pack.json`` or a
    direct manifest path.  Invalid optional packs are ignored so a broken
    third-party pack cannot take down the built-in catalog.
    """
    entries: list[dict[str, Any]] = []
    raw = os.environ.get(TEMPLATE_PACK_ENV, "")
    for item in raw.split(os.pathsep):
        candidate = item.strip()
        if not candidate:
            continue
        path = os.path.abspath(candidate)
        try:
            if os.path.isdir(path):
                manifest_path = os.path.join(path, "template_pack.json")
                if not os.path.isfile(manifest_path):
                    continue
                value = _read_json_path(manifest_path)
                root = path
            elif os.path.isfile(path) and path.lower().endswith(".json"):
                value = _read_json_path(path)
                root = os.path.dirname(path)
            else:
                continue
            if isinstance(value, Mapping) and isinstance(value.get("packs"), list):
                for pack in value["packs"]:
                    if isinstance(pack, Mapping):
                        item_pack = dict(pack)
                        item_pack.setdefault("root", root)
                        item_pack.setdefault("source", "external")
                        entries.append(item_pack)
            elif isinstance(value, Mapping):
                item_pack = dict(value)
                item_pack.setdefault("root", root)
                item_pack.setdefault("source", "external")
                entries.append(item_pack)
        except (OSError, ValueError, EngineError):
            continue
    return entries


def _template_map(value: Any) -> dict[str, Any]:
    if isinstance(value, Mapping):
        items = value.get("templates")
        if isinstance(items, Mapping):
            return {str(key).lower(): dict(item) for key, item in items.items() if isinstance(item, Mapping)}
        if isinstance(items, list):
            result: dict[str, Any] = {}
            for item in items:
                if isinstance(item, Mapping) and item.get("id"):
                    result[str(item["id"]).lower()] = dict(item)
            return result
    if isinstance(value, list):
        result = {}
        for item in value:
            if isinstance(item, Mapping) and item.get("id"):
                result[str(item["id"]).lower()] = dict(item)
        return result
    return {}


def template_packs() -> list[dict[str, Any]]:
    """List registered template packs and their actual asset counts."""
    manifest = _read_json("template_packs.json")
    packs = manifest.get("packs")
    if not isinstance(packs, list):
        raise EngineError("RESOURCE_INVALID", "template_packs.json must contain a packs array")
    result: list[dict[str, Any]] = []
    for raw in list(packs) + _external_pack_entries():
        if not isinstance(raw, Mapping):
            continue
        pack = dict(raw)
        pack.setdefault("enabled", True)
        assets = _pack_assets(pack)
        pack["themeCount"] = _asset_count(assets.get("themes"), "themes")
        pack["layoutCount"] = _asset_count(assets.get("layouts"), "layouts")
        pack["templateCount"] = _asset_count(assets.get("templates"), "templates")
        result.append(pack)
    return result


def load_template_pack(pack_id: str) -> dict[str, Any]:
    wanted = str(pack_id or "").strip().lower()
    for pack in template_packs():
        if str(pack.get("id") or "").lower() == wanted:
            assets = _pack_assets(pack)
            return {"manifest": pack, "assets": assets}
    raise EngineError("TEMPLATE_PACK_NOT_FOUND", f"unknown template pack: {pack_id}", {"allowed": [str(pack.get("id")) for pack in template_packs()]})


def _asset_items(value: Any, key: str) -> list[dict[str, Any]]:
    """Return the list of asset entries from a pack asset document."""
    items: Any = None
    if isinstance(value, Mapping):
        items = value.get(key)
        if items is None and key in ("themes", "layouts", "templates"):
            items = value
    elif isinstance(value, list):
        items = value
    if isinstance(items, Mapping):
        items = list(items.values())
    if not isinstance(items, list):
        return []
    return [dict(item) for item in items if isinstance(item, Mapping)]


def _pack_asset_issues(pack: Mapping[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    """Validate one pack's manifest and referenced asset files.

    Returns ``(assets, errors, warnings)``. Every error is a small JSON-safe
    record so callers can present a machine-readable repair list.
    """
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    def error(code: str, message: str, **extra: Any) -> None:
        errors.append({"code": code, "message": message, **extra})

    def warning(code: str, message: str, **extra: Any) -> None:
        warnings.append({"code": code, "message": message, **extra})

    if not str(pack.get("id") or "").strip():
        error("PACK_ID_MISSING", "template pack must declare a non-empty id")
    if not str(pack.get("label") or "").strip():
        error("PACK_LABEL_MISSING", "template pack must declare a non-empty label")
    if not str(pack.get("version") or "").strip():
        error("PACK_VERSION_MISSING", "template pack must declare a non-empty version")
    if pack.get("enabled") is False:
        warning("PACK_DISABLED", "template pack is present but disabled")

    assets: dict[str, Any] = {}
    for key in ("themes", "layouts", "templates"):
        raw_path = pack.get(key)
        if not raw_path:
            error("ASSET_NOT_CONFIGURED", f"{key} path is not configured")
            continue
        path = _pack_asset_path(pack, key)
        if not path or not os.path.isfile(path):
            error("ASSET_MISSING", f"{key} file does not exist", **{"path": path})
            continue
        try:
            value = _read_json_path(path)
        except EngineError as exc:
            error("ASSET_INVALID_JSON", f"{key} file is not valid JSON", **{"path": path, "error": str(exc)})
            continue
        assets[key] = value
        if not _asset_items(value, key):
            error("ASSET_EMPTY", f"{key} file contains no entries", **{"path": path})

    layout_ids = {
        str(item.get("id") or "").strip()
        for item in _asset_items(assets.get("layouts"), "layouts")
        if str(item.get("id") or "").strip()
    }
    theme_ids = {
        str(item.get("id") or "").strip()
        for item in _asset_items(assets.get("themes"), "themes")
        if str(item.get("id") or "").strip()
    }
    templates = _asset_items(assets.get("templates"), "templates")
    for index, template in enumerate(templates):
        template_id = str(template.get("id") or "").strip()
        if not template_id:
            error("TEMPLATE_ID_MISSING", "template entry must declare a non-empty id", **{"index": index})
        slides = template.get("slides")
        if not isinstance(slides, list) or not slides:
            error("TEMPLATE_SLIDES_MISSING", "template entry must contain a non-empty slides array", **{"template": template_id or index})
            continue
        for slide_index, slide in enumerate(slides):
            if not isinstance(slide, Mapping):
                error("TEMPLATE_SLIDE_INVALID", "template slide must be an object", **{"template": template_id or index, "index": slide_index})
                continue
            layout = str(slide.get("layout") or "").strip()
            if not layout:
                error("TEMPLATE_LAYOUT_MISSING", "template slide must declare a layout", **{"template": template_id or index, "index": slide_index})
            elif layout_ids and layout not in layout_ids:
                error("TEMPLATE_LAYOUT_UNKNOWN", "template slide references an unknown layout", **{"template": template_id or index, "layout": layout})
        theme = str(template.get("theme") or "").strip()
        if theme and theme_ids and theme not in theme_ids:
            warning("TEMPLATE_THEME_UNKNOWN", "template theme is not present in the pack themes", **{"template": template_id or index, "theme": theme})

    return assets, errors, warnings


def validate_template_pack(pack_id: str | None = None) -> dict[str, Any]:
    """Validate one pack, or every registered pack.

    Validation is read-only and never loads or executes template content. It
    reports manifest, asset-file, and cross-reference problems so pack authors
    can fix a pack before distributing it.
    """
    packs = template_packs()
    wanted = str(pack_id or "").strip().lower()
    if wanted:
        packs = [pack for pack in packs if str(pack.get("id") or "").lower() == wanted]
        if not packs:
            raise EngineError("TEMPLATE_PACK_NOT_FOUND", f"unknown template pack: {pack_id}", {"allowed": [str(pack.get("id")) for pack in template_packs()]})
    seen: set[str] = set()
    reports: list[dict[str, Any]] = []
    for pack in packs:
        pid = str(pack.get("id") or "").strip()
        if pid and pid.lower() in seen:
            reports.append({"id": pid, "valid": False, "errors": [{"code": "PACK_DUPLICATE_ID", "message": "template pack id is duplicated"}], "warnings": [], "counts": {"themes": 0, "layouts": 0, "templates": 0}})
            continue
        if pid:
            seen.add(pid.lower())
        assets, errors, warnings = _pack_asset_issues(pack)
        reports.append({
            "id": pid or None,
            "label": str(pack.get("label") or ""),
            "version": str(pack.get("version") or ""),
            "source": str(pack.get("source") or "builtin"),
            "valid": not errors,
            "errors": errors,
            "warnings": warnings,
            "counts": {
                "themes": len(_asset_items(assets.get("themes"), "themes")),
                "layouts": len(_asset_items(assets.get("layouts"), "layouts")),
                "templates": len(_asset_items(assets.get("templates"), "templates")),
            },
        })
    return {"valid": all(item["valid"] for item in reports), "checked": len(reports), "packs": reports}


def _resolve_template_from_pack(deck: Mapping[str, Any]) -> dict[str, Any]:
    pack_id = str(deck.get("templatePack") or "").strip()
    if not pack_id:
        return office_pptx._resolve_template_spec(deck)
    pack = load_template_pack(pack_id)
    templates = _template_map(pack["assets"].get("templates"))
    template_name = str(deck.get("template") or "").strip().lower()
    template = templates.get(template_name)
    if template is None:
        raise EngineError("TEMPLATE_NOT_FOUND", f"unknown template in pack {pack_id}: {template_name}", {"allowed": sorted(templates)})
    resolved = dict(deck)
    if not resolved.get("slides"):
        resolved["slides"] = [dict(item) for item in template.get("slides") or []]
    if not resolved.get("theme"):
        resolved["theme"] = template.get("theme")
    resolved["_templateName"] = template_name
    resolved["_templatePack"] = pack_id
    return resolved


def normalize_deck(spec: Any, *, resolve_template: bool = True, default_theme: str | None = None) -> dict[str, Any]:
    """Normalize any accepted deck-shaped input into Deck IR.

    ``spec`` may be either ``{slides: [...]}`` or ``{deck: {slides: [...]}}``.
    Existing Office Forge template ids are expanded before normalization so a
    template is just a reusable Deck IR fragment.
    """
    if not isinstance(spec, Mapping):
        raise EngineError("INVALID_ARGUMENT", "deck spec must be a JSON object")
    raw = spec.get("deck") if isinstance(spec.get("deck"), Mapping) else spec
    deck = dict(raw)
    if resolve_template and deck.get("template"):
        deck = _resolve_template_from_pack(deck) if deck.get("templatePack") else office_pptx._resolve_template_spec(deck)
    slides = deck.get("slides")
    if slides is None:
        slides = []
    if not isinstance(slides, list):
        raise EngineError("DECK_INVALID", "deck.slides must be an array")
    normalized_slides: list[dict[str, Any]] = []
    for index, raw_slide in enumerate(slides):
        if not isinstance(raw_slide, Mapping):
            raise EngineError("DECK_INVALID", f"deck.slides[{index}] must be an object")
        slide = dict(raw_slide)
        slide.setdefault("layout", "content")
        blocks = slide.get("blocks")
        if blocks is not None:
            if not isinstance(blocks, list):
                raise EngineError("DECK_INVALID", f"deck.slides[{index}].blocks must be an array")
            normalized_blocks: list[dict[str, Any]] = []
            for block_index, raw_block in enumerate(blocks):
                if not isinstance(raw_block, Mapping):
                    raise EngineError("DECK_INVALID", f"deck.slides[{index}].blocks[{block_index}] must be an object")
                block = dict(raw_block)
                block_type = str(block.get("type") or "").lower()
                if block_type not in DECK_BLOCK_TYPES:
                    raise EngineError(
                        "DECK_INVALID",
                        f"deck.slides[{index}].blocks[{block_index}] has unknown type: {block_type or '<empty>'}",
                    )
                block["type"] = block_type
                normalized_blocks.append(block)
            slide["blocks"] = normalized_blocks
        normalized_slides.append(slide)
    deck["slides"] = normalized_slides
    deck["schemaVersion"] = max(int(deck.get("schemaVersion") or DECK_SCHEMA_VERSION), DECK_SCHEMA_VERSION)
    if default_theme and not deck.get("theme"):
        deck["theme"] = default_theme
    return deck


def _block_text(block: Mapping[str, Any]) -> str:
    value = block.get("text")
    if value is None and isinstance(block.get("items"), list):
        value = " ".join(str(item) for item in block["items"])
    return str(value or "")


def _block_values(block: Mapping[str, Any]) -> list[Any]:
    values = block.get("items")
    if values is None:
        values = block.get("values")
    if isinstance(values, list):
        return values
    text = block.get("text")
    return [text] if text not in (None, "") else []


def _blocks_to_legacy_slide(slide: Mapping[str, Any]) -> dict[str, Any]:
    """Flatten Deck IR blocks for the legacy python-pptx renderer.

    The Studio adapter understands ``blocks`` directly; legacy accepts the
    older top-level fields.  Keeping one flattening function here means every
    adapter sees the same content and no block is silently dropped.
    """
    result = dict(slide)
    blocks = slide.get("blocks")
    if not isinstance(blocks, list):
        return result
    bullets: list[str] = []
    paragraphs: list[str] = []
    kpis: list[dict[str, Any]] = []
    table: dict[str, Any] | None = None
    chart: dict[str, Any] | None = None
    images: list[Any] = []
    shapes: list[dict[str, Any]] = []
    lines: list[dict[str, Any]] = []
    media: list[dict[str, Any]] = []
    comments: list[dict[str, Any]] = []
    smartart: dict[str, Any] | None = None
    transition: dict[str, Any] | None = None
    animation: dict[str, Any] | None = None
    for block in blocks:
        if not isinstance(block, Mapping):
            continue
        kind = str(block.get("type") or "").lower()
        if kind == "title":
            result["title"] = _block_text(block)
        elif kind == "subtitle":
            result["subtitle"] = _block_text(block)
        elif kind == "bullets":
            bullets.extend(str(item) for item in _block_values(block) if str(item).strip())
        elif kind in ("paragraphs", "paragraph", "text"):
            paragraphs.extend(str(item) for item in _block_values(block) if str(item).strip())
        elif kind in ("kpi", "kpis", "metrics", "stats", "timeline", "matrix"):
            for item in _block_values(block):
                if isinstance(item, Mapping):
                    kpis.append(dict(item))
                else:
                    kpis.append({"value": item})
        elif kind == "table":
            raw_rows = block.get("rows") or block.get("data") or []
            rows = [list(row) if isinstance(row, (list, tuple)) else [row] for row in raw_rows]
            headers = block.get("headers") or block.get("header") or block.get("columns")
            if isinstance(headers, list):
                table = {"headers": list(headers), "rows": rows}
            else:
                table = {"rows": rows}
        elif kind == "chart":
            chart = dict(block.get("chart") or block)
        elif kind in ("image", "images"):
            raw = block.get("images") or block.get("items")
            if isinstance(raw, list):
                images.extend(raw)
            elif block.get("image"):
                images.append(block["image"])
            elif block.get("path") or block.get("src"):
                images.append(block)
        elif kind in ("shape", "shapes"):
            shapes.append({k: v for k, v in block.items() if k != "type"} | {"shape": block.get("shape") or "RECTANGLE"})
        elif kind in ("line", "lines"):
            lines.append({k: v for k, v in block.items() if k != "type"})
        elif kind == "media":
            media.append({k: v for k, v in block.items() if k != "type"})
        elif kind in ("comment", "comments"):
            comments.append({k: v for k, v in block.items() if k != "type"})
        elif kind == "smartart":
            smartart = {k: v for k, v in block.items() if k != "type"}
        elif kind == "transition":
            raw = block.get("transition") if isinstance(block.get("transition"), Mapping) else block
            transition = {k: v for k, v in raw.items() if k != "type"}
        elif kind == "animation":
            raw = block.get("animation") if isinstance(block.get("animation"), Mapping) else block
            animation = {k: v for k, v in raw.items() if k != "type"}
        elif kind == "quote":
            result["text"] = _block_text(block)
        elif kind in ("note", "notes"):
            result["notes"] = _block_text(block)
    if bullets:
        result.setdefault("bullets", [])
        result["bullets"] = list(result["bullets"]) + bullets
    if paragraphs:
        result.setdefault("paragraphs", [])
        result["paragraphs"] = list(result["paragraphs"]) + paragraphs
    if kpis:
        result["kpis"] = kpis
    if table is not None:
        result["table"] = table
    if chart is not None:
        result["chart"] = chart
    if images:
        result["images"] = images
    if shapes:
        result.setdefault("shapes", [])
        result["shapes"] = list(result["shapes"]) + shapes
    if lines:
        result.setdefault("lines", [])
        result["lines"] = list(result["lines"]) + lines
    if media:
        existing_media = result.get("media")
        result["media"] = (list(existing_media) if isinstance(existing_media, list) else []) + media
    if comments:
        existing_comments = result.get("comments")
        result["comments"] = (list(existing_comments) if isinstance(existing_comments, list) else []) + comments
    if smartart is not None:
        result["smartArt"] = smartart
    if transition is not None:
        result["transition"] = transition
    if animation is not None:
        result["animation"] = animation
    return result


def deck_to_backend_spec(deck: Mapping[str, Any]) -> dict[str, Any]:
    """Convert normalized Deck IR to the legacy python-pptx spec shape.

    The legacy renderer ignores extra IR fields; the explicit copy keeps the
    adapter boundary honest and prevents backend mutating the caller's deck.
    """
    result = dict(deck)
    result.pop("schemaVersion", None)
    result.pop("meta", None)
    # Template resolution already happened in normalize_deck().  Leaving the
    # marker here makes the legacy renderer resolve the template a second time
    # against the built-in catalog, which would reject pack-local templates.
    result.pop("template", None)
    result.pop("templatePack", None)
    result["slides"] = [_blocks_to_legacy_slide(slide) for slide in deck.get("slides", [])]
    return result


def apply_deck_operations(deck: Mapping[str, Any], operations: Any) -> dict[str, Any]:
    """Apply a small, design-aware edit script to a Deck IR copy.

    Operations are intentionally declarative: they edit Deck IR blocks rather
    than a renderer-private shape tree, so an imported deck can be changed and
    re-rendered through any adapter without losing its theme/master context.
    """
    if operations is None:
        return copy.deepcopy(dict(deck))
    if not isinstance(operations, list):
        raise EngineError("INVALID_ARGUMENT", "operations must be an array")
    result = copy.deepcopy(dict(deck))
    slides = result.get("slides")
    if not isinstance(slides, list):
        slides = []
        result["slides"] = slides
    for index, raw in enumerate(operations):
        if not isinstance(raw, Mapping):
            raise EngineError("INVALID_ARGUMENT", f"operation {index} must be an object")
        op = str(raw.get("op") or raw.get("operation") or "").lower()
        if op in ("settheme", "set-theme"):
            theme = raw.get("theme")
            if not isinstance(theme, str) or not theme.strip():
                raise EngineError("INVALID_ARGUMENT", "setTheme requires theme")
            result["theme"] = theme
            continue
        if op in ("settitle", "set-title"):
            slide_index = int(raw.get("slide", raw.get("index", 0)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            title = raw.get("title")
            if title is None:
                raise EngineError("INVALID_ARGUMENT", "setTitle requires title")
            slide = dict(slides[slide_index])
            slide["title"] = str(title)
            blocks = list(slide.get("blocks") or [])
            replaced = False
            for block_index, block in enumerate(blocks):
                if isinstance(block, Mapping) and str(block.get("type") or "").lower() == "title":
                    blocks[block_index] = {**block, "text": str(title)}
                    replaced = True
                    break
            if not replaced:
                blocks.insert(0, {"type": "title", "text": str(title)})
            slide["blocks"] = blocks
            slides[slide_index] = slide
            continue
        if op == "addslide":
            slide = raw.get("slide") or raw.get("spec")
            if not isinstance(slide, Mapping):
                raise EngineError("INVALID_ARGUMENT", "addSlide requires slide")
            slides.append(copy.deepcopy(dict(slide)))
            continue
        if op == "removeslide":
            slide_index = int(raw.get("slide", raw.get("index", -1)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            slides.pop(slide_index)
            continue
        if op == "replaceslide":
            slide_index = int(raw.get("slide", raw.get("index", -1)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            slide = raw.get("slide") or raw.get("spec")
            if not isinstance(slide, Mapping):
                raise EngineError("INVALID_ARGUMENT", "replaceSlide requires slide")
            slides[slide_index] = copy.deepcopy(dict(slide))
            continue
        if op in ("appendblock", "append-block"):
            slide_index = int(raw.get("slide", raw.get("index", -1)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            block = raw.get("block")
            if not isinstance(block, Mapping):
                raise EngineError("INVALID_ARGUMENT", "appendBlock requires block")
            slide = dict(slides[slide_index])
            slide["blocks"] = list(slide.get("blocks") or []) + [copy.deepcopy(dict(block))]
            slides[slide_index] = slide
            continue
        if op in ("replaceblock", "replace-block"):
            slide_index = int(raw.get("slide", raw.get("index", -1)))
            block_index = int(raw.get("block", raw.get("blockIndex", -1)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            blocks = list(dict(slides[slide_index]).get("blocks") or [])
            if block_index < 0 or block_index >= len(blocks):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"block index {block_index} is out of range")
            block = raw.get("block") or raw.get("value")
            if not isinstance(block, Mapping):
                raise EngineError("INVALID_ARGUMENT", "replaceBlock requires block")
            blocks[block_index] = copy.deepcopy(dict(block))
            slide = dict(slides[slide_index])
            slide["blocks"] = blocks
            slides[slide_index] = slide
            continue
        if op in ("setnotes", "set-notes"):
            slide_index = int(raw.get("slide", raw.get("index", 0)))
            if slide_index < 0 or slide_index >= len(slides):
                raise EngineError("DECK_INDEX_OUT_OF_RANGE", f"slide index {slide_index} is out of range")
            slide = dict(slides[slide_index])
            slide["notes"] = str(raw.get("notes") or "")
            slides[slide_index] = slide
            continue
        raise EngineError("INVALID_ARGUMENT", f"unknown deck operation: {op}")
    return result


def summarize_deck(deck: Mapping[str, Any]) -> dict[str, Any]:
    slides = deck.get("slides") if isinstance(deck.get("slides"), list) else []
    layouts: dict[str, int] = {}
    block_types: dict[str, int] = {}
    for slide in slides:
        if not isinstance(slide, Mapping):
            continue
        layout = str(slide.get("layout") or "content")
        layouts[layout] = layouts.get(layout, 0) + 1
        blocks = slide.get("blocks")
        if isinstance(blocks, list):
            for block in blocks:
                if isinstance(block, Mapping):
                    kind = str(block.get("type") or "unknown")
                    block_types[kind] = block_types.get(kind, 0) + 1
    return {
        "slides": len(slides),
        "layouts": layouts,
        "blockTypes": block_types,
        "theme": deck.get("theme"),
        "template": deck.get("template"),
        "title": deck.get("title"),
        "schemaVersion": deck.get("schemaVersion", DECK_SCHEMA_VERSION),
    }


def catalog() -> dict[str, Any]:
    return {
        "schemaVersion": DECK_SCHEMA_VERSION,
        "blockTypes": list(DECK_BLOCK_TYPES),
        "layouts": [dict(item) for item in office_pptx._LAYOUT_CATALOG.values()],
        "templates": [dict(item) for item in office_pptx._TEMPLATE_CATALOG.values()],
        "themes": [{"name": key, **value} for key, value in office_pptx._THEMES.items()],
        "templatePacks": template_packs(),
    }
