"""Bundled Anon1009 aesthetic adapter.

This module vendors the MIT-licensed audit/beautify scripts from
Anon1009/dsh-ppt-studio so the adapter works without a separate install.
It performs file-level audit and non-destructive beautification only; it does
not register tools or prompt hooks of its own.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
from typing import Any, Mapping

from office_common import EngineError, assert_inside, atomic_save

_VENDOR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor", "anon1009", "python")
if _VENDOR not in sys.path:
    sys.path.insert(0, _VENDOR)

import ppt_audit as _anon_audit  # noqa: E402
import ppt_beautify as _anon_beautify  # noqa: E402


def _safe_path(path: str, workspace: str) -> str:
    candidate = os.path.abspath(os.path.join(workspace, path) if not os.path.isabs(path) else path)
    assert_inside(candidate, workspace, "Anon1009 path")
    return candidate


def audit_anon(path: str, args: Mapping[str, Any] | None = None) -> dict[str, Any]:
    source = os.path.abspath(path)
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"file not found: {path}")
    try:
        result = _anon_audit.audit_pptx(source)
    except Exception as error:  # noqa: BLE001 - stable engine error for the host
        raise EngineError("ANON1009_AUDIT_FAILED", f"Anon1009 audit failed: {error}")
    result["adapter"] = "anon1009"
    result["engine"] = "anon1009"
    return result


def beautify_anon(
    path: str,
    args: Mapping[str, Any],
    limits: Mapping[str, Any],
    workspace: str,
    features: Any = None,
) -> dict[str, Any]:
    source = _safe_path(path, workspace)
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"file not found: {path}")
    requested = args.get("output") or args.get("outputPath")
    if requested:
        target = _safe_path(str(requested), workspace)
    else:
        stem, extension = os.path.splitext(source)
        target = f"{stem}-beautified{extension or '.pptx'}"
    overwrite = bool(args.get("overwrite", False))
    if os.path.exists(target) and not overwrite:
        raise EngineError("OUTPUT_EXISTS", f"output already exists: {target}")

    temp_dir = tempfile.mkdtemp(prefix=".anon1009-", dir=workspace)
    temp_output = os.path.join(temp_dir, "beautified.pptx")
    try:
        theme = str(args.get("theme") or "executive-navy")
        options = args.get("options")
        if not isinstance(options, dict):
            options = {
                "normalizeFonts": args.get("normalizeFonts", True),
                "applyThemeColors": args.get("applyThemeColors", True),
                "fixContrast": args.get("fixContrast", True),
                "fixAlignment": args.get("fixAlignment", True),
                "upgradeTables": args.get("upgradeTables", True),
                "addHeaderAccent": args.get("addHeaderAccent", True),
                "addSlideNumbers": args.get("addSlideNumbers", True),
                "enforceWidescreen": args.get("enforceWidescreen", False),
            }
        result = _anon_beautify.beautify_pptx(source, temp_output, theme_id=theme, options=options)
        max_bytes = int(limits.get("maxOutputBytes", 100 * 1024 * 1024))
        atomic_save(
            target,
            lambda temporary: shutil.copyfile(temp_output, temporary),
            overwrite=overwrite,
            max_bytes=max_bytes,
            make_backup=True,
        )
    except EngineError:
        raise
    except Exception as error:  # noqa: BLE001
        raise EngineError("ANON1009_BEAUTIFY_FAILED", f"Anon1009 beautify failed: {error}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    result["adapter"] = "anon1009"
    result["engine"] = "anon1009"
    result["outputPath"] = target
    return result
