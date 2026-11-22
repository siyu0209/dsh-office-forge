"""Shared, dependency-free helpers for the Office Forge Python engine.

The engine is deliberately a short-lived child process.  Every helper here is
bounded and side-effect explicit: paths are checked against the workspace,
writes go through a same-directory temporary file, and no function recurses
into another tool.
"""

from __future__ import annotations

import csv
import datetime as _dt
import decimal
import json
import math
import os
import re
import shutil
import uuid
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Sequence


class EngineError(Exception):
    """A classified, model-facing failure."""

    def __init__(self, code: str, message: str, detail: Any = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail

    def as_dict(self) -> dict[str, Any]:
        value: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.detail is not None:
            value["detail"] = jsonable(self.detail)
        return value


# ---------------------------------------------------------------------------
# JSON / scalar normalisation
# ---------------------------------------------------------------------------


def jsonable(value: Any) -> Any:
    """Convert Python document values into lossless-enough JSON values."""
    if value is None or isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    if isinstance(value, decimal.Decimal):
        return float(value)
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()
    if isinstance(value, (_dt.timedelta,)):
        return value.total_seconds()
    if isinstance(value, bytes):
        return {"encoding": "base64", "length": len(value)}
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [jsonable(v) for v in value]
    return str(value)


def scalar_text(value: Any) -> str:
    """Human-readable form for a cell/paragraph value."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()
    if isinstance(value, decimal.Decimal):
        return format(value, "f")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


def feature_enabled(features: Any, group: str, name: str, default: bool = True) -> bool:
    """Return a nested feature toggle, defaulting open for older payloads."""
    if not isinstance(features, Mapping):
        return default
    group_value = features.get(group)
    if not isinstance(group_value, Mapping):
        return default
    return bool(group_value.get(name, default))


def require_feature(features: Any, group: str, name: str, label: str) -> None:
    if not feature_enabled(features, group, name):
        raise EngineError(
            "FEATURE_DISABLED",
            f"{label} is disabled by the plugin feature settings",
            {"feature": f"{group}.{name}"},
        )


def get_path(data: Any, dotted: str, default: Any = None) -> Any:
    """Read ``a.b.c`` (and list indexes) from nested data."""
    current = data
    for part in str(dotted).split("."):
        if isinstance(current, Mapping):
            if part not in current:
                return default
            current = current[part]
        elif isinstance(current, (list, tuple)):
            try:
                current = current[int(part)]
            except (ValueError, IndexError):
                return default
        else:
            return default
    return current


_MISSING = object()
_PLACEHOLDER = re.compile(r"\{\{\s*([^{}#/][^{}]*?)\s*\}\}")


def render_placeholders(text: str, data: Any, missing: str = "keep") -> str:
    """Replace ``{{key}}`` placeholders; ``missing`` is keep|empty|error."""
    if not isinstance(text, str) or "{{" not in text:
        return text

    def replace(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        # Optional simple filters: {{key|upper}}, {{key|lower}}, {{key|json}}
        parts = [part.strip() for part in key.split("|")]
        key = parts[0]
        value = get_path(data, key, _MISSING)
        if value is _MISSING:
            if missing == "error":
                raise EngineError("TEMPLATE_MISSING_KEY", f"template placeholder {{{{{key}}}}} has no value", {"key": key})
            if missing == "empty":
                return ""
            return match.group(0)
        for operation in parts[1:]:
            if operation == "upper":
                value = str(value).upper()
            elif operation == "lower":
                value = str(value).lower()
            elif operation == "json":
                value = json.dumps(jsonable(value), ensure_ascii=False)
            elif operation == "title":
                value = str(value).title()
        if isinstance(value, (dict, list)):
            return json.dumps(jsonable(value), ensure_ascii=False)
        return scalar_text(value)

    return _PLACEHOLDER.sub(replace, text)


# ---------------------------------------------------------------------------
# Path and byte policy
# ---------------------------------------------------------------------------


def canonical(path: str) -> str:
    return os.path.normcase(os.path.realpath(os.path.abspath(path)))


def assert_inside(path: str, root: str, label: str = "path") -> None:
    """Reject traversal outside the workspace root."""
    if not root:
        return
    candidate = canonical(path)
    base = canonical(root)
    if candidate != base and not candidate.startswith(base + os.sep):
        raise EngineError("PATH_OUTSIDE_WORKSPACE", f"{label} is outside the session workspace", {"path": path, "workspace": root})


def assert_input_file(path: str, allowed_ext: Sequence[str], max_bytes: int) -> os.stat_result:
    if not path or not os.path.isfile(path):
        raise EngineError("INPUT_NOT_FOUND", f"input file does not exist: {path}", {"path": path})
    ext = os.path.splitext(path)[1].lower()
    if allowed_ext and ext not in allowed_ext:
        raise EngineError("UNSUPPORTED_FORMAT", f"unsupported file type {ext or '(none)'}", {"path": path, "allowed": list(allowed_ext)})
    info = os.stat(path)
    if info.st_size > max_bytes:
        raise EngineError("INPUT_TOO_LARGE", f"input file exceeds {max_bytes} bytes", {"path": path, "bytes": info.st_size, "limit": max_bytes})
    return info


def ensure_parent(path: str) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)


def backup_path(path: str) -> str:
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = f"{path}.bak-{stamp}"
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{path}.bak-{stamp}-{counter}"
        counter += 1
    return candidate


def atomic_save(
    path: str,
    writer: Callable[[str], None],
    *,
    overwrite: bool,
    max_bytes: int,
    make_backup: bool = False,
) -> tuple[int, str | None]:
    """Write through a same-directory temp file and atomically publish it."""
    path = os.path.abspath(path)
    ensure_parent(path)
    existed = os.path.exists(path)
    if existed and not overwrite:
        raise EngineError("OUTPUT_EXISTS", f"output already exists (set overwrite=true to replace it): {path}", {"path": path})
    saved_backup: str | None = None
    if existed and make_backup:
        saved_backup = backup_path(path)
        shutil.copy2(path, saved_backup)
    temp = os.path.join(os.path.dirname(path), f".{os.path.basename(path)}.tmp-{uuid.uuid4().hex}")
    try:
        writer(temp)
        if not os.path.isfile(temp):
            raise EngineError("WRITE_FAILED", "document writer produced no file", {"path": path})
        size = os.path.getsize(temp)
        if size > max_bytes:
            raise EngineError("OUTPUT_TOO_LARGE", f"output file exceeds {max_bytes} bytes", {"path": path, "bytes": size, "limit": max_bytes})
        os.replace(temp, path)
        return size, saved_backup
    finally:
        if os.path.exists(temp):
            try:
                os.remove(temp)
            except OSError:
                pass


def write_bytes(path: str, data: bytes, *, overwrite: bool, max_bytes: int, make_backup: bool = False) -> tuple[int, str | None]:
    return atomic_save(path, lambda temp: Path(temp).write_bytes(data), overwrite=overwrite, max_bytes=max_bytes, make_backup=make_backup)


# ---------------------------------------------------------------------------
# Document core properties
# ---------------------------------------------------------------------------

DEFAULT_DOCUMENT_AUTHOR = "Office Forge"

#: Canonical property -> accepted spellings in a ``properties`` mapping.
#: The three formats disagree on names (docx ``author`` vs xlsx ``creator``),
#: so one alias table keeps the create paths consistent.
CORE_PROPERTY_ALIASES: dict[str, tuple[str, ...]] = {
    "title": ("title",),
    "subject": ("subject",),
    "author": ("author", "creator"),
    "keywords": ("keywords",),
    "comments": ("comments", "description"),
    "category": ("category",),
    "lastModifiedBy": ("lastModifiedBy", "last_modified_by", "lastmodifiedby"),
}


def core_property_lookup(values: Any, name: str) -> Any:
    """Return the first non-empty value for a canonical property name."""
    if not isinstance(values, Mapping):
        return None
    for alias in CORE_PROPERTY_ALIASES.get(name, (name,)):
        if alias in values:
            candidate = values[alias]
            if candidate is None:
                continue
            if isinstance(candidate, str) and not candidate.strip():
                continue
            return candidate
    return None


def apply_core_properties(
    core: Any,
    values: Any = None,
    *,
    now: _dt.datetime | None = None,
    default_author: str = DEFAULT_DOCUMENT_AUTHOR,
    revision: int = 1,
    author_attr: str = "author",
    last_modified_attr: str = "last_modified_by",
    comments_attr: str = "comments",
) -> dict[str, Any]:
    """Overwrite the template's leftover metadata with explicit, neutral values.

    Office libraries ship a default template whose ``author`` is ``python-docx``
    / ``python-pptx`` / ``openpyxl`` and whose ``created`` is a 2013 timestamp.
    Leaving those in place makes generated files look like someone else's work,
    so every create path routes through here.  Caller-supplied values win; the
    rest fall back to a neutral author and the current UTC instant.

    The attribute names differ per library (openpyxl uses ``creator`` and
    ``lastModifiedBy``), so callers pass their own names instead of the helper
    guessing.
    """
    stamp = now or _dt.datetime.now(_dt.timezone.utc).replace(tzinfo=None)
    author = scalar_text(core_property_lookup(values, "author") or default_author)
    applied: dict[str, Any] = {"author": author, "created": stamp.isoformat(), "revision": int(revision)}
    setattr(core, author_attr, author)
    setattr(core, last_modified_attr, scalar_text(core_property_lookup(values, "lastModifiedBy") or author))
    core.created = stamp
    core.modified = stamp
    core.revision = int(revision)
    for name in ("title", "subject", "keywords", "category"):
        supplied = core_property_lookup(values, name)
        if supplied is not None:
            setattr(core, name, scalar_text(supplied))
    # The docx template ships ``comments = "generated by python-docx"``; that is
    # a library artefact, not document content, so an unsupplied value clears it
    # instead of inheriting it.
    comments = core_property_lookup(values, "comments")
    setattr(core, comments_attr, scalar_text(comments) if comments is not None else "")
    return applied


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------


def truncate(text: str, limit: int) -> tuple[str, bool]:
    if limit <= 0 or len(text) <= limit:
        return text, False
    return text[:limit], True


def word_counts(text: str) -> dict[str, int]:
    """Count CJK characters and whitespace-delimited words without external deps."""
    cjk = len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", text))
    latin = len(re.findall(r"[A-Za-z0-9_]+", text))
    return {"cjkChars": cjk, "latinWords": latin, "chars": len(text)}


def md_escape(value: Any) -> str:
    return scalar_text(value).replace("\\", "\\\\").replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def md_table(headers: Sequence[Any], rows: Iterable[Sequence[Any]]) -> str:
    header_cells = [md_escape(value) for value in headers]
    if not header_cells:
        return ""
    lines = ["| " + " | ".join(header_cells) + " |", "|" + "|".join([" --- "] * len(header_cells)) + "|"]
    for row in rows:
        cells = [md_escape(value) for value in row]
        if len(cells) < len(header_cells):
            cells += [""] * (len(header_cells) - len(cells))
        lines.append("| " + " | ".join(cells[: len(header_cells)]) + " |")
    return "\n".join(lines)


def safe_filename(value: Any, fallback: str = "item") -> str:
    text = scalar_text(value).strip() or fallback
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:120] or fallback


# ---------------------------------------------------------------------------
# Data files
# ---------------------------------------------------------------------------


def read_data_file(path: str, max_bytes: int) -> tuple[Any, str]:
    """Read JSON/CSV/TSV data used by batch templates."""
    info = os.stat(path)
    if info.st_size > max_bytes:
        raise EngineError("INPUT_TOO_LARGE", "data file exceeds the input limit", {"path": path, "bytes": info.st_size, "limit": max_bytes})
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, "r", encoding="utf-8-sig") as handle:
            return json.load(handle), "json"
    if ext in (".csv", ".tsv"):
        delimiter = "\t" if ext == ".tsv" else ","
        with open(path, "r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle, delimiter=delimiter)
            rows = [dict(row) for row in reader]
            return rows, "csv"
    raise EngineError("UNSUPPORTED_DATA_FORMAT", "data file must be .json, .csv or .tsv", {"path": path})


def materialize_spec(value: Any, label: str = "spec") -> dict[str, Any]:
    if not isinstance(value, dict):
        raise EngineError("INVALID_ARGUMENT", f"{label} must be a JSON object", {"type": type(value).__name__})
    return value


def require_list(value: Any, label: str, max_items: int | None = None) -> list[Any]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise EngineError("INVALID_ARGUMENT", f"{label} must be an array", {"type": type(value).__name__})
    if max_items is not None and len(value) > max_items:
        raise EngineError("LIMIT_EXCEEDED", f"{label} has too many items", {"count": len(value), "limit": max_items})
    return value


def require_dict(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise EngineError("INVALID_ARGUMENT", f"{label} must be an object", {"type": type(value).__name__})
    return value
