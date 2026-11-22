"""JSON stdin/stdout dispatcher for the Office Forge Python engine.

The Node layer spawns this file once per tool call.  It never talks to a
network service and never keeps state between calls, so a crashed document
operation cannot poison the host process.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import traceback
from typing import Any, Mapping

import office_backends
import office_docx
import office_migrations
import office_pdf
import office_pdf_tools
import office_pptx
import office_xlsx
from office_common import EngineError, assert_inside, jsonable, read_data_file, require_feature

LIMITS: dict[str, Any] = {
    "maxInputBytes": 52_428_800,
    "maxOutputBytes": 104_857_600,
    "maxTextChars": 200_000,
    "maxRows": 5_000,
    "maxCols": 300,
    "maxSlides": 300,
    "maxSheets": 60,
    "maxBatchItems": 20,
}

_ALLOWED = {
    "docx": (".docx", ".docm", ".dotx"),
    "xlsx": (".xlsx", ".xlsm"),
    "pptx": (".pptx", ".pptm"),
    "pdf": (".pdf",),
}


def _workspace() -> str:
    return os.environ.get("OFFICEFORGE_WORKSPACE") or os.getcwd()


def _infer_type(path: str, explicit: Any = None) -> str:
    if explicit:
        value = str(explicit).lower().lstrip(".")
        if value not in _ALLOWED:
            raise EngineError("UNSUPPORTED_FORMAT", f"unsupported document type: {value}", {"supported": sorted(_ALLOWED)})
        return value
    extension = os.path.splitext(path)[1].lower()
    for kind, extensions in _ALLOWED.items():
        if extension in extensions:
            return kind
    raise EngineError("UNSUPPORTED_FORMAT", f"cannot infer document type from {extension or '(none)'}", {"supported": sorted(_ALLOWED)})


def _checked(path: Any, label: str = "path") -> str:
    if not path:
        raise EngineError("INVALID_ARGUMENT", f"{label} is required")
    raw = str(path)
    workspace = _workspace()
    # Relative paths belong to the workspace, never to the engine's own cwd;
    # resolving them against the process cwd would report a spurious
    # PATH_OUTSIDE_WORKSPACE instead of the real problem.
    absolute = raw if os.path.isabs(raw) else os.path.join(workspace, raw)
    absolute = os.path.abspath(absolute)
    assert_inside(absolute, workspace, label)
    return absolute


def _module_version(name: str) -> str:
    module = __import__(name)
    return str(getattr(module, "__version__", "unknown"))


def probe() -> dict[str, Any]:
    return {
        "python": sys.version.split()[0],
        "executable": sys.executable,
        "workspace": _workspace(),
        "modules": {
            "python-docx": _module_version("docx"),
            "openpyxl": _module_version("openpyxl"),
            "python-pptx": _module_version("pptx"),
            "Pillow": _module_version("PIL"),
            "lxml": _module_version("lxml"),
        },
        "limits": LIMITS,
        "capabilities": {
            "docx": {"read": True, "create": True, "edit": True, "template": True, "loops": True, "images": True, "hyperlinks": True, "toc": True, "captions": True, "styles": True, "fields": True, "comments": True, "revisions": True, "mailMerge": True, "footnotes": True, "bookmarks": True, "crossReferences": True, "citations": True, "columns": True, "watermark": True, "contentControls": True, "dotx": True, "audit": ["health", "profile"]},
            "xlsx": {"read": True, "create": True, "edit": True, "template": True, "loops": True, "charts": True, "comboCharts": True, "conditionalFormatting": True, "dataValidation": True, "comments": True, "mergedCells": True, "images": True, "formulas": True, "arrayFormulas": True, "dataTools": True, "pivots": True, "tables": True, "protection": True, "import": ["csv", "tsv", "json"], "audit": ["health", "profile", "insight", "diff"]},
            "pptx": {"read": True, "create": True, "edit": True, "template": True, "loops": False, "charts": True, "images": True, "transitions": True, "shapes": True, "textBoxes": True, "connectors": True, "animations": True, "smartart": True, "notes": True, "sections": True, "comments": True, "media": True, "master": True, "layouts": sorted(office_pptx._LAYOUT_CATALOG), "templates": sorted(office_pptx._TEMPLATE_CATALOG), "morph": True, "audit": ["deck", "profile", "themes", "layouts", "templates"], "themes": sorted(office_pptx._THEMES)},
            "pdf": {"read": True, "exportViaOfficeKit": True, "toolkit": ["merge", "split", "rotate", "watermark", "pageNumbers", "formFields", "fillForm", "extractImages", "ocr", "encrypt", "decrypt"]},
            "native": {"backend": "win32com", "operations": ["convert", "exportPdf", "acceptRevisions", "recalculate", "pivot", "smartArt"]},
        },
        "backends": office_backends.status(),
    }


def doctor() -> dict[str, Any]:
    """Return the runtime facts needed to diagnose a clean install."""
    info = probe()
    vendor = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor", "pypdf")
    return {
        "engine": info,
        "backends": office_backends.status(),
        "pdfVendor": os.path.isdir(vendor),
        "pdfVendorPath": vendor if os.path.isdir(vendor) else None,
    }


def _read(payload: Mapping[str, Any]) -> Any:
    path = _checked(payload.get("path"))
    kind = _infer_type(path, payload.get("type"))
    info = office_common_stat(path)
    if info.st_size > LIMITS["maxInputBytes"]:
        raise EngineError("INPUT_TOO_LARGE", "input file exceeds the input limit", {"bytes": info.st_size, "limit": LIMITS["maxInputBytes"]})
    args = payload.get("args") if isinstance(payload.get("args"), Mapping) else {}
    if kind == "docx":
        return office_docx.read_docx(path, args, LIMITS)
    if kind == "xlsx":
        return office_xlsx.read_xlsx(path, args, LIMITS)
    if kind == "pdf":
        return office_pdf.read_pdf(path, args, LIMITS)
    return office_pptx.read_pptx(path, args, LIMITS)


def office_common_stat(path: str) -> os.stat_result:
    if not os.path.isfile(path):
        raise EngineError("INPUT_NOT_FOUND", f"input file does not exist: {path}", {"path": path})
    return os.stat(path)


def _create(payload: Mapping[str, Any]) -> Any:
    path = _checked(payload.get("path"), "path")
    kind = _infer_type(path, payload.get("type"))
    spec = payload.get("spec")
    if not isinstance(spec, Mapping):
        raise EngineError("INVALID_ARGUMENT", "spec must be a JSON object")
    if kind == "pdf":
        raise EngineError("UNSUPPORTED_OPERATION", "PDF is read-only in Office Forge")
    spec = office_migrations.migrate_spec(kind, spec, payload.get("schemaVersion"))
    overwrite = bool(payload.get("overwrite", False))
    features = payload.get("features")
    if kind == "docx":
        result = office_docx.create_docx(path, spec, LIMITS, overwrite, _workspace(), features)
    elif kind == "xlsx":
        result = office_xlsx.create_xlsx(path, spec, LIMITS, overwrite, _workspace(), features)
    else:
        result = office_pptx.create_pptx(path, spec, LIMITS, overwrite, _workspace(), features)
    result["schemaVersion"] = office_migrations.CURRENT_SCHEMA_VERSION
    if "migratedFrom" in spec:
        result["migratedFrom"] = spec["migratedFrom"]
    return result


def _edit(payload: Mapping[str, Any]) -> Any:
    path = _checked(payload.get("path"), "path")
    kind = _infer_type(path, payload.get("type"))
    office_common_stat(path)
    args = payload.get("args")
    if not isinstance(args, Mapping):
        raise EngineError("INVALID_ARGUMENT", "args must be a JSON object")
    if kind == "pdf":
        raise EngineError("UNSUPPORTED_OPERATION", "PDF is read-only in Office Forge")
    features = payload.get("features")
    if kind == "docx":
        return office_docx.edit_docx(path, args, LIMITS, _workspace(), features)
    if kind == "xlsx":
        return office_xlsx.edit_xlsx(path, args, LIMITS, _workspace(), features)
    return office_pptx.edit_pptx(path, args, LIMITS, _workspace(), features)


def _audit(payload: Mapping[str, Any]) -> Any:
    features = payload.get("features")
    requested_mode = str(payload.get("mode") or "").lower()
    if not payload.get("path") and requested_mode in ("themes", "layouts", "templates"):
        require_feature(features, "ppt", "themes" if requested_mode == "themes" else requested_mode, f"PowerPoint {requested_mode} catalog")
        return office_pptx.audit_pptx("", requested_mode, {})
    path = _checked(payload.get("path"), "path")
    kind = _infer_type(path, payload.get("type"))
    info = office_common_stat(path)
    if info.st_size > LIMITS["maxInputBytes"]:
        raise EngineError("INPUT_TOO_LARGE", "input file exceeds the input limit", {"bytes": info.st_size, "limit": LIMITS["maxInputBytes"]})
    args = payload.get("args") if isinstance(payload.get("args"), Mapping) else {}
    features = payload.get("features")
    mode = str(payload.get("mode") or ("health" if kind in ("docx", "xlsx") else "deck")).lower()
    if kind == "docx":
        require_feature(features, "word", "audit", "Word document audit")
        return office_docx.audit_docx(path, mode, args)
    if kind == "xlsx":
        require_feature(features, "excel", "audit", "Excel workbook audit")
        compare_path = _checked(payload.get("comparePath"), "comparePath") if payload.get("comparePath") else None
        return office_xlsx.audit_xlsx(path, mode, args, compare_path)
    if kind == "pptx":
        require_feature(features, "ppt", "audit", "PowerPoint deck audit")
        return office_pptx.audit_pptx(path, mode, args)
    raise EngineError("UNSUPPORTED_OPERATION", f"audit is not available for {kind}", {"supported": ["docx", "xlsx", "pptx"]})


def _template(payload: Mapping[str, Any]) -> Any:
    template = _checked(payload.get("template"), "template")
    output = _checked(payload.get("output"), "output")
    kind = _infer_type(template, payload.get("type"))
    office_common_stat(template)
    data = payload.get("data")
    if data is None and payload.get("dataPath"):
        data_path = _checked(payload["dataPath"], "dataPath")
        data, _ = read_data_file(data_path, LIMITS["maxInputBytes"])
    if isinstance(data, list):
        # CSV/TSV and JSON arrays are row collections; expose conventional
        # names so {{#each items}} and {{#each rows}} both work.
        data = {"items": data, "rows": data}
    if not isinstance(data, Mapping):
        raise EngineError("INVALID_ARGUMENT", "data must be a JSON object")
    options = payload.get("options") if isinstance(payload.get("options"), Mapping) else {}
    if kind == "docx":
        return office_docx.template_docx(template, output, data, options, LIMITS)
    if kind == "xlsx":
        return office_xlsx.template_xlsx(template, output, data, options, LIMITS)
    return office_pptx.template_pptx(template, output, data, options, LIMITS)


def _batch(payload: Mapping[str, Any]) -> Any:
    items = payload.get("items")
    if not isinstance(items, list) or not items:
        raise EngineError("INVALID_ARGUMENT", "batch.items must be a non-empty array")
    if len(items) > LIMITS["maxBatchItems"]:
        raise EngineError("LIMIT_EXCEEDED", f"batch has too many items (max {LIMITS['maxBatchItems']})", {"count": len(items)})
    results: list[dict[str, Any]] = []
    batch_features = payload.get("features")
    for index, item in enumerate(items):
        item_id = str(item.get("id") or index + 1) if isinstance(item, Mapping) else str(index + 1)
        if not isinstance(item, Mapping):
            results.append({"id": item_id, "ok": False, "error": {"code": "INVALID_ARGUMENT", "message": "batch item must be an object"}})
            continue
        action = str(item.get("action") or "")
        if action == "batch":
            results.append({"id": item_id, "ok": False, "error": {"code": "INVALID_OPERATION", "message": "nested batch is not allowed"}})
            continue
        try:
            item_payload = dict(item)
            if item_payload.get("features") is None and batch_features is not None:
                item_payload["features"] = batch_features
            result = _dispatch(item_payload)
            results.append({"id": item_id, "ok": True, "result": result})
        except EngineError as error:
            results.append({"id": item_id, "ok": False, "error": error.as_dict()})
        except Exception as error:  # noqa: BLE001 - one bad item must not abort the batch
            results.append({"id": item_id, "ok": False, "error": {"code": "ENGINE_ERROR", "message": f"{type(error).__name__}: {error}"}})
    succeeded = sum(1 for item in results if item["ok"])
    return {"items": results, "succeeded": succeeded, "failed": len(results) - succeeded}


def _copy(payload: Mapping[str, Any]) -> Any:
    """Copy a converted file from the trusted officeKit output directory into
    the session workspace, enforcing the same atomic-write rules as the
    document writers."""
    source_raw = str(payload.get("source") or "")
    if not source_raw:
        raise EngineError("INVALID_ARGUMENT", "source is required")
    source = os.path.abspath(source_raw)
    allowed_root = os.environ.get("OFFICEFORGE_OFFICEKIT_DIR")
    if allowed_root:
        assert_inside(source, os.path.abspath(allowed_root), "source")
    else:
        _checked(source, "source")
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"{source} does not exist")
    size = os.path.getsize(source)
    if size > LIMITS["maxOutputBytes"]:
        raise EngineError("OUTPUT_TOO_LARGE", "converted file exceeds the output limit", {"bytes": size, "limit": LIMITS["maxOutputBytes"]})
    target = _checked(payload.get("target"), "target")
    overwrite = bool(payload.get("overwrite", False))
    if os.path.exists(target) and not overwrite:
        raise EngineError("OUTPUT_EXISTS", f"{target} already exists", {"path": target})
    parent = os.path.dirname(target) or "."
    os.makedirs(parent, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".officeforge-copy-", dir=parent)
    os.close(fd)
    try:
        shutil.copyfile(source, temporary)
        os.replace(temporary, target)
    finally:
        if os.path.exists(temporary):
            try:
                os.unlink(temporary)
            except OSError:
                pass
    return {"type": os.path.splitext(target)[1].lstrip(".").lower() or "file", "path": target, "bytes": size}


def _pdfop(payload: Mapping[str, Any]) -> Any:
    return office_pdf_tools.pdf_operation(payload, _workspace(), LIMITS)


def _native(payload: Mapping[str, Any]) -> Any:
    operation = str(payload.get("operation") or "").lower()
    source = _checked(payload.get("path") or payload.get("source"), "path")
    target = _checked(payload.get("output") or payload.get("target"), "output")
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"input file does not exist: {source}")
    if os.path.getsize(source) > LIMITS["maxInputBytes"]:
        raise EngineError("INPUT_TOO_LARGE", "input file exceeds the input limit", {"path": source})
    if os.path.exists(target) and not bool(payload.get("overwrite", False)):
        raise EngineError("OUTPUT_EXISTS", f"{target} already exists", {"path": target})
    timeout = int(payload.get("timeout") or 120)
    if operation == "exportpdf":
        return office_backends.export_pdf(source, target, _workspace(), timeout)
    if operation not in {"convert", "acceptrevisions", "recalculate", "pivot", "smartart"}:
        raise EngineError("UNSUPPORTED_OPERATION", f"unsupported native operation: {operation or '(empty)'}", {"allowed": ["convert", "exportPdf", "acceptRevisions", "recalculate", "pivot", "smartArt"]})
    action = {"acceptrevisions": "acceptRevisions", "smartart": "smartArt"}.get(operation, operation)
    return office_backends.run_com(action, {"source": source, "target": target, "spec": payload.get("spec") or {}}, _workspace(), timeout)


def _convert(payload: Mapping[str, Any]) -> Any:
    source = _checked(payload.get("source") or payload.get("path"), "source")
    target = _checked(payload.get("target") or payload.get("output"), "target")
    if not os.path.isfile(source):
        raise EngineError("INPUT_NOT_FOUND", f"input file does not exist: {source}")
    info = os.stat(source)
    if info.st_size > LIMITS["maxInputBytes"]:
        raise EngineError("INPUT_TOO_LARGE", "input file exceeds the input limit", {"bytes": info.st_size, "limit": LIMITS["maxInputBytes"]})
    return office_backends.convert_legacy(source, target, _workspace(), timeout=int(payload.get("timeout") or 120))


def _dispatch(payload: Mapping[str, Any]) -> Any:
    action = str(payload.get("action") or "")
    if action == "read":
        return _read(payload)
    if action == "create":
        return _create(payload)
    if action == "edit":
        return _edit(payload)
    if action == "audit":
        return _audit(payload)
    if action == "template":
        return _template(payload)
    if action == "batch":
        return _batch(payload)
    if action == "copy":
        return _copy(payload)
    if action == "convert":
        return _convert(payload)
    if action == "pdfop":
        return _pdfop(payload)
    if action == "native":
        return _native(payload)
    if action == "doctor":
        return doctor()
    if action == "probe":
        return probe()
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported engine action: {action or '(empty)'}", {"allowed": ["read", "create", "edit", "audit", "template", "batch", "copy", "convert", "pdfop", "native", "doctor", "probe"]})


def _entry() -> dict[str, Any]:
    try:
        raw = sys.stdin.read()
        if not raw.strip():
            raise EngineError("INVALID_ARGUMENT", "empty stdin; expected a JSON request")
        payload = json.loads(raw)
        if not isinstance(payload, Mapping):
            raise EngineError("INVALID_ARGUMENT", "request must be a JSON object")
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            result = _dispatch(payload)
        return {"ok": True, "result": jsonable(result)}
    except EngineError as error:
        return {"ok": False, "error": error.as_dict()}
    except json.JSONDecodeError as error:
        return {"ok": False, "error": {"code": "INVALID_JSON", "message": f"invalid request JSON: {error}"}}
    except Exception as error:  # noqa: BLE001 - the Node layer needs a classified failure
        return {
            "ok": False,
            "error": {
                "code": "ENGINE_ERROR",
                "message": f"{type(error).__name__}: {error}",
                "detail": {"traceback": traceback.format_exc()[-4000:]},
            },
        }


def main() -> None:
    try:
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    response = _entry()
    sys.stdout.write(json.dumps(response, ensure_ascii=False, default=jsonable))
    sys.stdout.flush()


if __name__ == "__main__":
    main()
