"""Deck adapter registry.

The registry is the only place that chooses a PPT renderer.  Adapters use the
same Deck IR contract, so adding a new engine does not change ``officeforge_*``
tool names or the legacy ``create_pptx`` behavior.
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from typing import Any, Mapping

import office_deck_ir
import office_ppt_anon
import office_ppt_backends
import office_pptx
import office_webdeck
from pptx import Presentation
from office_common import EngineError, require_feature

_PLUGIN_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_EXTERNAL_ADAPTERS = (
    {
        "name": "zbsph",
        "label": "dsh-ppt-studio (zbsph)",
        "kind": "planner",
        "package": "dsh-ppt-studio",
        "capabilities": ["deck_plan", "layout", "template", "tables", "charts", "images", "notes", "shapes", "background", "office_render_check"],
        "hint": "npm install dsh-ppt-studio@1.2.7, then set OFFICEFORGE_ZBSPH_DIR to its package directory",
    },
    {
        "name": "anon1009",
        "label": "Anon1009 dsh-ppt-studio",
        "kind": "aesthetic",
        "package": "",
        "capabilities": ["aesthetic_audit", "beautify", "layout"],
        "hint": "Install Anon1009/dsh-ppt-studio and set OFFICEFORGE_ANON1009_DIR",
    },
    {
        "name": "guizang",
        "label": "歸藏设计 (skill-backed HTML deck)",
        "kind": "skill",
        "package": "",
        "capabilities": ["web_deck", "html_export", "design"],
        "hint": "Bundled guizang-ppt-skill + guizang-page-skill + guizang-social-card-skill; generate single-file HTML instead of native PPTX",
    },
    {
        "name": "composer",
        "label": "dsh-ppt-composer (bundled web deck)",
        "kind": "web_deck",
        "package": "",
        "capabilities": ["web_deck", "html_export", "storyboard"],
        "hint": "Bundled lightweight web deck renderer; optional ez118/dsh-ppt-composer adds its full storyboard UI",
    },
)


def _manifest_ok(path: str) -> bool:
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "r", encoding="utf-8") as handle:
            json.load(handle)
        return True
    except (OSError, ValueError):
        return False


def _package_manifests(package: str) -> list[str]:
    """Return package.json candidates without assuming where pnpm hoisted it."""
    candidates = [
        os.path.join(_PLUGIN_ROOT, "node_modules", package, "package.json"),
        # Installed plugins live in <profile>/node_modules/dsh-office-forge, so a
        # hoisted sibling package is one level above the plugin root.
        os.path.join(os.path.dirname(_PLUGIN_ROOT), package, "package.json"),
    ]
    for variable in ("OFFICEFORGE_PROFILE_DIR", "DSH_PROFILE_DIR"):
        root = os.environ.get(variable, "").strip()
        if root:
            candidates.append(os.path.join(root, "node_modules", package, "package.json"))
    return candidates


def _external_available(entry: Mapping[str, Any]) -> bool:
    if entry.get("name") == "anon1009":
        return os.path.isfile(os.path.join(_PLUGIN_ROOT, "lib", "vendor", "anon1009", "python", "ppt_audit.py"))
    if entry.get("name") == "guizang":
        return os.path.isfile(os.path.join(_PLUGIN_ROOT, "resources", "guizang-themes.json"))
    if entry.get("name") == "composer":
        return True
    env_name = "OFFICEFORGE_" + str(entry["name"]).upper() + "_DIR"
    if os.path.isdir(os.environ.get(env_name, "")):
        return True
    package = str(entry.get("package") or "")
    if package:
        for candidate in _package_manifests(package):
            if _manifest_ok(candidate):
                return True
    return False


class ExternalPptAdapter:
    """Descriptor-backed adapter for optional PPT tools.

    The adapter is deliberately non-executing until its owner package is
    installed and wired.  It fails with a stable ADAPTER_UNAVAILABLE error
    instead of silently pretending to render.
    """

    def __init__(self, entry: Mapping[str, Any]) -> None:
        self.entry = dict(entry)
        self.name = str(entry["name"])
        self.label = str(entry["label"])

    def _unavailable(self) -> EngineError:
        return EngineError(
            "ADAPTER_UNAVAILABLE",
            f"PPT adapter '{self.name}' is not available",
            {"adapter": self.name, "hint": self.entry.get("hint"), "package": self.entry.get("package")},
        )

    def create(self, *args: Any, **kwargs: Any) -> dict[str, Any]:
        raise self._unavailable()

    def edit(self, *args: Any, **kwargs: Any) -> dict[str, Any]:
        raise self._unavailable()

    def audit(self, *args: Any, **kwargs: Any) -> dict[str, Any]:
        raise self._unavailable()


class AnonPptAdapter:
    """Bundled Anon1009 aesthetic renderer plus beautifier.

    Creating with this adapter first builds a deterministic native deck, then
    applies the vendored Anon1009 beautifier so ``engine: "anon1009"`` is a
    real create/edit path rather than a descriptor-only placeholder.
    """

    name = "anon1009"
    label = "Anon1009 dsh-ppt-studio (bundled aesthetic rules)"

    def __init__(self) -> None:
        self._legacy = office_ppt_backends.LegacyPptBackend()

    def _theme(self, args: Mapping[str, Any] | None = None) -> str:
        raw = str((args or {}).get("theme") or "executive-navy").strip()
        aliases = {
            "forge": "executive-navy",
            "consulting": "executive-navy",
            "ocean": "tech-obsidian",
            "tech": "tech-obsidian",
            "academic": "academic-emerald",
            "minimal": "minimal-ivory",
            "swiss": "minimal-ivory",
            "editorial": "editorial-terracotta",
            "morandi": "morandi-slate",
        }
        return aliases.get(raw.lower(), raw)

    def create(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
        require_feature(features, "ppt", "beautify", "PowerPoint beautify")
        os.makedirs(workspace, exist_ok=True)
        temp_dir = tempfile.mkdtemp(prefix=".anon-create-", dir=workspace)
        temp_deck = os.path.join(temp_dir, "deck.pptx")
        try:
            native_spec = dict(spec)
            native_spec["theme"] = "forge"
            self._legacy.create(temp_deck, native_spec, limits, True, workspace, features)
            result = office_ppt_anon.beautify_anon(
                temp_deck,
                {"output": path, "theme": self._theme(spec), "overwrite": bool(overwrite)},
                limits,
                workspace,
                features,
            )
            result["type"] = "pptx"
            result["path"] = path
            result["bytes"] = os.path.getsize(path)
            result["slides"] = len(Presentation(path).slides)
            result["theme"] = self._theme(spec)
            return result
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def edit(self, path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None) -> dict[str, Any]:
        require_feature(features, "ppt", "beautify", "PowerPoint beautify")
        source = os.path.abspath(path)
        temp_dir = tempfile.mkdtemp(prefix=".anon-edit-", dir=workspace)
        temp_deck = os.path.join(temp_dir, "deck.pptx")
        try:
            shutil.copy2(source, temp_deck)
            self._legacy.edit(temp_deck, args, limits, workspace, features)
            result = office_ppt_anon.beautify_anon(
                temp_deck,
                {"output": path, "theme": self._theme(args), "overwrite": True},
                limits,
                workspace,
                features,
            )
            result["type"] = "pptx"
            result["path"] = path
            result["bytes"] = os.path.getsize(path)
            result["slides"] = len(Presentation(path).slides)
            result["theme"] = self._theme(args)
            return result
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def audit(self, path: str, mode: str = "aesthetic", args: Mapping[str, Any] | None = None, features: Any = None) -> dict[str, Any]:
        require_feature(features, "ppt", "audit", "PowerPoint deck audit")
        return office_ppt_anon.audit_anon(path, args)

    def repair(self, path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None) -> dict[str, Any]:
        require_feature(features, "ppt", "beautify", "PowerPoint beautify")
        return office_ppt_anon.beautify_anon(path, args, limits, workspace, features)


class GuizangPptAdapter:
    """Bundled 歸藏 design adapter.

    Guizang is HTML-first: its real skill ships a self-contained HTML deck
    template, motion runtime and presenter shell.  This adapter therefore
    emits real Guizang HTML for ``.html`` targets and never pretends a native
    PPTX theme swap is the full Guizang system.
    """

    name = "guizang"
    label = "歸藏设计 (skill-backed HTML deck)"

    def create(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
        if not str(path).lower().endswith((".html", ".htm")):
            raise EngineError(
                "ADAPTER_UNSUPPORTED",
                "Guizang is an HTML-first skill; use .html output or choose studio for editable PPTX",
                {"adapter": "guizang", "supported": ["html"], "hint": "officeforge_web or officeforge_deck with a .html path"},
            )
        import office_guizang
        result = office_guizang.render_html(path, spec, limits, overwrite, workspace, features)
        result["adapter"] = self.name
        return result

    def edit(self, path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None) -> dict[str, Any]:
        raise EngineError(
            "ADAPTER_UNSUPPORTED",
            "Guizang HTML is regenerated from Deck IR; use officeforge_deck edit with a .pptx source then render .html",
            {"adapter": "guizang", "operation": "edit"},
        )

    def audit(self, path: str, mode: str, args: Mapping[str, Any]) -> dict[str, Any]:
        return office_pptx.audit_pptx(path, mode, args)


class ComposerPptAdapter:
    """Bundled lightweight web-deck adapter.

    The optional ez118/dsh-ppt-composer package remains a richer external
    storyboard UI; this adapter provides a deterministic single-file HTML
    renderer with the same Deck IR contract and no extra dependency.
    """

    name = "composer"
    label = "dsh-ppt-composer (bundled web deck)"

    def create(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
        result = office_webdeck.create_html_deck(path, spec, limits, overwrite, workspace, features)
        result["engine"] = self.name
        return result

    def edit(self, *args: Any, **kwargs: Any) -> dict[str, Any]:
        raise EngineError(
            "ADAPTER_UNSUPPORTED",
            "composer is a web-deck renderer; edit the generated HTML or use officeforge_web",
            {"adapter": "composer", "operation": "edit"},
        )

    def audit(self, path: str, mode: str, args: Mapping[str, Any]) -> dict[str, Any]:
        return office_pptx.audit_pptx(path, mode, args)


def _capabilities(name: str) -> list[str]:
    if name == "studio":
        return [
            "native_pptx",
            "theme",
            "layout",
            "kpi",
            "bullets",
            "title",
            "section",
            "template",
            "tables",
            "charts",
            "images",
        ]
    return [
        "native_pptx",
        "charts",
        "images",
        "tables",
        "comments",
        "media",
        "master",
        "morph",
        "animations",
        "smartart",
        "sections",
    ]



_LEGACY_ONLY_FEATURES = {
    "comments",
    "media",
    "master",
    "morph",
    "animations",
    "smartart",
    "sections",
    "transitions",
}
_STUDIO_FEATURES = {"tables", "charts", "images", "design", "notes", "masterPreset"}
_ZBSPH_FEATURES = _STUDIO_FEATURES | {
    "shapes", "background", "safeArea", "overlaps", "tableHeader", "chartOptions",
}
_LEGACY_FEATURES = _STUDIO_FEATURES | _LEGACY_ONLY_FEATURES | {"shapes", "background"}
_ADAPTER_FEATURES = {
    "legacy": _LEGACY_FEATURES,
    "studio": _STUDIO_FEATURES,
    "zbsph": _ZBSPH_FEATURES,
    "guizang": _LEGACY_FEATURES,
    "composer": {"web_deck", "html_export", "storyboard", "design"},
    "anon1009": _LEGACY_FEATURES | {"aesthetic_audit", "beautify"},
}


def _deck_features(deck: Mapping[str, Any]) -> set[str]:
    features: set[str] = set()
    if deck.get("template") or deck.get("templatePack") or deck.get("theme"):
        features.add("design")
    for key in _LEGACY_ONLY_FEATURES:
        if deck.get(key):
            features.add(key)
    for key, feature in (("smartArt", "smartart"), ("animation", "animations"), ("transition", "transitions")):
        if deck.get(key):
            features.add(feature)
    # A string master is a named dsh-ppt preset (studio can render it); an
    # object master is the legacy raw master/footer/background shape.
    master = deck.get("master")
    features.discard("master")
    if isinstance(master, Mapping):
        features.add("master")
    elif master or deck.get("masterPreset"):
        features.add("masterPreset")
    for slide in deck.get("slides") or []:
        if not isinstance(slide, Mapping):
            continue
        for key in _LEGACY_ONLY_FEATURES:
            if slide.get(key):
                features.add(key)
        for key, feature in (("smartArt", "smartart"), ("animation", "animations")):
            if slide.get(key):
                features.add(feature)
        transition = slide.get("transition")
        if isinstance(transition, Mapping):
            features.add("transitions")
            if str(transition.get("type") or transition.get("effect") or "").lower() == "morph":
                features.add("morph")
        if slide.get("table"):
            features.add("tables")
            table = slide.get("table")
            if isinstance(table, Mapping) and (table.get("header") or table.get("headers") or table.get("columns")):
                features.add("tableHeader")
        if slide.get("chart"):
            features.add("charts")
            chart = slide.get("chart")
            if isinstance(chart, Mapping) and any(chart.get(key) is not None for key in ("legend", "labels", "axes")):
                features.add("chartOptions")
        if slide.get("image") or slide.get("images"):
            features.add("images")
        if slide.get("shapes") or slide.get("shape") or slide.get("lines") or slide.get("line"):
            features.add("shapes")
        if slide.get("background"):
            features.add("background")
        if slide.get("safeArea"):
            features.add("safeArea")
        if slide.get("overlapMode") or slide.get("expectedOverlaps") or slide.get("expectedOutOfSafeArea"):
            features.add("overlaps")
        if slide.get("notes"):
            features.add("notes")
        for block in slide.get("blocks") or []:
            if not isinstance(block, Mapping):
                continue
            kind = str(block.get("type") or "").lower()
            if kind == "table":
                features.add("tables")
                table = block.get("table")
                if isinstance(table, Mapping) and (table.get("header") or table.get("headers") or table.get("columns")):
                    features.add("tableHeader")
            elif kind == "chart":
                features.add("charts")
                chart = block.get("chart")
                if isinstance(chart, Mapping) and any(chart.get(key) is not None for key in ("legend", "labels", "axes")):
                    features.add("chartOptions")
            elif kind in ("image", "images"):
                features.add("images")
            elif kind in ("shape", "shapes", "line", "lines"):
                features.add("shapes")
            elif kind == "media":
                features.add("media")
            elif kind in ("comment", "comments"):
                features.add("comments")
            elif kind == "smartart":
                features.add("smartart")
            elif kind == "transition":
                features.add("transitions")
            elif kind == "animation":
                features.add("animations")
    return features


def _adapter_entry(name: str) -> Mapping[str, Any] | None:
    for entry in _EXTERNAL_ADAPTERS:
        if entry["name"] == name:
            return entry
    return None


def _auto_select(features: set[str], deck: Mapping[str, Any]) -> tuple[str, str]:
    theme = str(deck.get("theme") or "").lower()
    if theme.startswith("guizang-") and "guizang" in _ADAPTER_FEATURES:
        entry = _adapter_entry("guizang")
        if entry is not None and _external_available(entry):
            return "guizang", "theme is a bundled guizang design token"
    advanced = bool(features & {
        "tables", "charts", "images", "shapes", "background",
        "safeArea", "overlaps", "tableHeader", "chartOptions", "masterPreset",
    })
    if advanced and not (features & _LEGACY_ONLY_FEATURES):
        entry = _adapter_entry("zbsph")
        if entry is not None and _external_available(entry):
            return "zbsph", "table/chart/image/shape/background block is supported by the zbsph renderer"
    if features & {"design", "tables", "charts", "images", "masterPreset"} and not (features & _LEGACY_ONLY_FEATURES):
        return "studio", "designed or table/chart/image deck is supported by the studio renderer"
    if features & _LEGACY_ONLY_FEATURES:
        return "legacy", "legacy-only Office objects require the python-pptx renderer"
    return "legacy", "plain deck uses the default python-pptx renderer"


def select_adapter(engine: Any, deck: Mapping[str, Any]) -> tuple[Any, dict[str, Any]]:
    requested = str(engine or "auto").lower()
    features = _deck_features(deck)
    if requested == "auto":
        selected, reason = _auto_select(features, deck)
        decision = {
            "requested": requested,
            "selected": selected,
            "fallbackFrom": None,
            "requiredFeatures": sorted(features),
            "unsupportedFeatures": sorted(features - _ADAPTER_FEATURES.get(selected, set())),
            "reason": reason,
        }
        return resolve_adapter(selected, deck), decision
    adapter = resolve_adapter(requested, deck)
    decision = {
        "requested": requested,
        "selected": requested,
        "fallbackFrom": None,
        "requiredFeatures": sorted(features),
        "unsupportedFeatures": sorted(features - _ADAPTER_FEATURES.get(requested, set())),
        "reason": "explicit renderer request",
    }
    return adapter, decision

def _theme_rows(path: str) -> list[dict[str, Any]]:
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        rows = payload.get("themes") if isinstance(payload, dict) else payload
        return [dict(item) for item in (rows or []) if isinstance(item, dict) and item.get("id")]
    except Exception:
        return []


def design_catalog() -> dict[str, Any]:
    """Expose every bundled design token set without loading a renderer."""
    anon_path = os.path.join(_PLUGIN_ROOT, "lib", "vendor", "anon1009", "assets", "themes.json")
    guizang_path = os.path.join(_PLUGIN_ROOT, "resources", "guizang-themes.json")
    office_path = os.path.join(_PLUGIN_ROOT, "resources", "themes.json")
    anon = _theme_rows(anon_path)
    guizang = _theme_rows(guizang_path)
    office = _theme_rows(office_path)
    return {
        "type": "deck",
        "mode": "designCatalog",
        "anon1009": {
            "themes": [{"id": item.get("id"), "label": item.get("name") or item.get("label") or item.get("id"), "mode": item.get("mode")} for item in anon],
            "modes": ["aestheticAudit", "beautify"],
        },
        "guizang": {
            "themes": [{"id": item.get("id"), "label": item.get("label") or item.get("id"), "family": item.get("family")} for item in guizang],
            "families": sorted({str(item.get("family")) for item in guizang if item.get("family")}),
        },
        "composer": {
            "themes": [{"id": item.get("id"), "label": item.get("label") or item.get("id")} for item in office],
            "format": "html",
        },
        "office": {
            "themes": [{"id": key, "label": value.get("label") or key} for key, value in office_pptx._THEMES.items()],
        },
    }


def adapters() -> list[dict[str, Any]]:
    result = [
        {
            "name": "legacy",
            "label": "python-pptx (legacy)",
            "capabilities": _capabilities("legacy"),
            "default": True,
        },
        {
            "name": "studio",
            "label": "dsh-ppt deck engine",
            "capabilities": _capabilities("studio"),
            "default": False,
            "fallback": "legacy",
        },
    ]
    for entry in _EXTERNAL_ADAPTERS:
        result.append(
            {
                "name": entry["name"],
                "label": entry["label"],
                "kind": entry["kind"],
                "capabilities": list(entry["capabilities"]),
                "available": _external_available(entry),
                "hint": entry["hint"],
                "default": False,
            }
        )
    return result


def loss_audit(deck: Mapping[str, Any], engine: Any = "auto") -> dict[str, Any]:
    """Compare a Deck IR against the selected adapter's capability set.

    This is deliberately a static report: it never renders a file and never
    pretends that a fallback is lossless.
    """
    features = _deck_features(deck)
    adapter, decision = select_adapter(engine, deck)
    selected = str(decision.get("selected") or "legacy")
    supported = set(_ADAPTER_FEATURES.get(selected, set()))
    lost = sorted(features - supported)
    return {
        "type": "deck",
        "mode": "lossAudit",
        "requestedEngine": str(engine or "auto"),
        "selectedAdapter": selected,
        "adapterLabel": getattr(adapter, "label", selected),
        "requiredFeatures": sorted(features),
        "supportedFeatures": sorted(features & supported),
        "lostFeatures": lost,
        "lossless": not lost,
        "adapterDecision": decision,
    }


def resolve_adapter(engine: Any, deck: Mapping[str, Any]) -> Any:
    requested = str(engine or "auto")
    for entry in _EXTERNAL_ADAPTERS:
        if entry["name"] == requested:
            if requested == "anon1009" and _external_available(entry):
                return AnonPptAdapter()
            if requested == "guizang" and _external_available(entry):
                return GuizangPptAdapter()
            if requested == "composer" and _external_available(entry):
                return ComposerPptAdapter()
            if requested == "zbsph" and _external_available(entry):
                return office_ppt_backends.ZbsphPptBackend()
            return ExternalPptAdapter(entry)
    try:
        return office_ppt_backends.backend_for(requested, deck)
    except EngineError:
        raise
    except Exception as error:  # noqa: BLE001 - registry must produce a stable error
        raise EngineError("DECK_ADAPTER_ERROR", f"could not resolve PPT adapter: {error}")


def render(path: str, deck: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    adapter, adapter_decision = select_adapter(engine, deck)
    result = adapter.create(path, office_deck_ir.deck_to_backend_spec(deck), limits, overwrite, workspace, features)
    result["deckSchemaVersion"] = deck.get("schemaVersion", office_deck_ir.DECK_SCHEMA_VERSION)
    result["adapter"] = getattr(adapter, "name", "legacy")
    result["adapterDecision"] = adapter_decision
    return result


def audit(path: str, mode: str, args: Mapping[str, Any], features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    if str(engine or "auto") == "anon1009":
        return AnonPptAdapter().audit(path, mode, args, features)
    require_feature(features, "ppt", "audit", "PowerPoint deck audit")
    result = office_pptx.audit_pptx(path, mode or "deck", args)
    result["adapter"] = "audit"
    return result


def repair(path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None, engine: Any = "auto") -> dict[str, Any]:
    if str(engine or "auto") == "anon1009":
        return AnonPptAdapter().repair(path, args, limits, workspace, features)
    require_feature(features, "ppt", "beautify", "PowerPoint beautify")
    result = office_pptx.beautify_pptx(path, args, limits, workspace, features)
    result["adapter"] = "beautify"
    return result
