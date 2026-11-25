"""PowerPoint backend selection.

The plugin has always exposed a stable ``officeforge_*`` tool contract.  The
backend module keeps the rendering implementation swappable without changing
that contract.  ``legacy`` is the original python-pptx path; ``studio`` uses
the dsh-ppt deck renderer for the subset it supports and falls back to the
legacy path for Office-specific objects (comments, media, masters, morph,
animations, SmartArt, pictures, tables and charts).
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from typing import Any, Mapping

import office_pptx
from office_common import EngineError, atomic_save


ENGINES = ("auto", "legacy", "studio", "zbsph")
_BRIDGE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "office_ppt_studio_bridge.mjs")
_ZBSPH_BRIDGE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "office_ppt_zbsp_bridge.mjs")


class PptBackend:
    name = "legacy"

    def create(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
        result = office_pptx.create_pptx(path, spec, limits, overwrite, workspace, features)
        result["engine"] = self.name
        return result

    def edit(self, path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str, features: Any = None) -> dict[str, Any]:
        result = office_pptx.edit_pptx(path, args, limits, workspace, features)
        result["engine"] = self.name
        return result

    def audit(self, path: str, mode: str, args: Mapping[str, Any]) -> dict[str, Any]:
        return office_pptx.audit_pptx(path, mode, args)


class LegacyPptBackend(PptBackend):
    name = "legacy"


class StudioPptBackend(PptBackend):
    name = "studio"
    bridge = _BRIDGE
    unavailable_code = "STUDIO_UNAVAILABLE"
    unsupported_code = "STUDIO_UNSUPPORTED"
    error_code = "STUDIO_ERROR"

    def create(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any = None) -> dict[str, Any]:
        # Resolve Office Forge template ids before handing the deck to the
        # studio renderer; templates are part of the public tool contract.
        spec = office_pptx._resolve_template_spec(spec)
        # A designed default is useful, but never invent content.
        if not spec.get("theme"):
            spec = dict(spec)
            spec["theme"] = "forge"

        node = shutil.which("node") or os.environ.get("OFFICEFORGE_NODE")
        if not node or not os.path.isfile(self.bridge):
            return self._legacy_result(path, spec, limits, overwrite, workspace, features, self.unavailable_code)

        workspace = os.path.abspath(workspace)
        tmpdir = tempfile.mkdtemp(prefix=".officeforge-studio-", dir=workspace)
        tmp_output = os.path.join(tmpdir, "deck.pptx")
        try:
            request = {
                "spec": dict(spec),
                "output": tmp_output,
                "workspace": workspace,
            }
            proc = subprocess.run(
                [node, self.bridge],
                input=json.dumps(request, ensure_ascii=False),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=180,
            )
            if proc.returncode != 0:
                return self._legacy_result(path, spec, limits, overwrite, workspace, features, self.unavailable_code)
            try:
                parsed = json.loads(proc.stdout.strip().splitlines()[-1])
            except (ValueError, IndexError):
                return self._legacy_result(path, spec, limits, overwrite, workspace, features, self.error_code)
            if not parsed.get("ok"):
                error = parsed.get("error") or {}
                code = error.get("code") or self.error_code
                if code in (self.unsupported_code, self.unavailable_code):
                    return self._legacy_result(path, spec, limits, overwrite, workspace, features, code)
                raise EngineError(code, error.get("message") or "studio renderer failed", error.get("detail"))
            source = parsed.get("pptxPath")
            if not source or not os.path.isfile(source):
                return self._legacy_result(path, spec, limits, overwrite, workspace, features, self.error_code)
            with open(source, "rb") as handle:
                data = handle.read()
            size, _ = atomic_save(
                path,
                lambda temp: _write_bytes(temp, data),
                overwrite=overwrite,
                max_bytes=int(limits.get("maxOutputBytes", 104857600)),
                make_backup=False,
            )
            return {
                "type": "pptx",
                "path": path,
                "bytes": size,
                "slides": parsed.get("slides"),
                "theme": parsed.get("theme"),
                "engine": self.name,
            }
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def _legacy_result(self, path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str, features: Any, reason: str) -> dict[str, Any]:
        result = office_pptx.create_pptx(path, spec, limits, overwrite, workspace, features)
        result["engine"] = "legacy"
        result["studioFallback"] = reason
        return result


class ZbsphPptBackend(StudioPptBackend):
    """Adapter for the dsh-ppt-studio (zbsph) PPTD renderer."""

    name = "zbsph"
    bridge = _ZBSPH_BRIDGE
    unavailable_code = "ZBSPH_UNAVAILABLE"
    unsupported_code = "ZBSPH_UNSUPPORTED"
    error_code = "ZBSPH_ERROR"


def _write_bytes(path: str, data: bytes) -> None:
    with open(path, "wb") as handle:
        handle.write(data)


_LEGACY = LegacyPptBackend()
_STUDIO = StudioPptBackend()


def normalize_engine(value: Any) -> str:
    name = str(value or "auto").lower()
    if name not in ENGINES:
        raise EngineError("INVALID_ARGUMENT", f"unsupported pptx engine: {name}", {"allowed": list(ENGINES)})
    return name


def resolve_engine(requested: Any, spec: Mapping[str, Any] | None = None) -> str:
    name = normalize_engine(requested)
    if name != "auto":
        return name
    # `auto` uses the original renderer unless the caller explicitly asks for
    # a designed template/theme, which is the low-risk migration boundary.
    if spec and any(spec.get(key) for key in ("template", "theme", "studio")):
        return "studio"
    return "legacy"


def backend_for(requested: Any, spec: Mapping[str, Any] | None = None) -> PptBackend:
    name = resolve_engine(requested, spec)
    return _STUDIO if name == "studio" else _LEGACY
