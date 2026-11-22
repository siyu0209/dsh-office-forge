"""Small, explicit schema migration layer for persisted Office Forge specs.

Specs are plain JSON objects supplied by callers.  A missing schemaVersion is
treated as version 1 (all pre-0.5 specs).  Migrations are deliberately
data-only and never execute document content.
"""

from __future__ import annotations

import copy
from typing import Any, Mapping

from office_common import EngineError

CURRENT_SCHEMA_VERSION = 2


def spec_version(spec: Mapping[str, Any]) -> int:
    value = spec.get("schemaVersion", 1)
    try:
        return int(value)
    except Exception as error:
        raise EngineError("INVALID_SCHEMA_VERSION", f"schemaVersion must be an integer: {value!r}") from error


def migrate_spec(kind: str, spec: Mapping[str, Any], requested: Any = None) -> dict[str, Any]:
    if not isinstance(spec, Mapping):
        raise EngineError("INVALID_ARGUMENT", "spec must be a JSON object")
    version = spec_version(spec)
    if requested is not None:
        try:
            version = int(requested)
        except Exception as error:
            raise EngineError("INVALID_SCHEMA_VERSION", f"schemaVersion must be an integer: {requested!r}") from error
    if version > CURRENT_SCHEMA_VERSION:
        raise EngineError("SCHEMA_VERSION_UNSUPPORTED", f"schemaVersion {version} is newer than this plugin supports", {"supported": CURRENT_SCHEMA_VERSION})
    migrated = copy.deepcopy(dict(spec))
    if version < 2:
        # v1 -> v2 is the compatibility boundary introduced in Office Forge
        # 0.5. Keep every caller field unchanged and record provenance.
        migrated["schemaVersion"] = 2
        migrated["migratedFrom"] = version
    else:
        migrated["schemaVersion"] = version
    migrated.setdefault("documentType", kind)
    return migrated


def migrate_options(options: Mapping[str, Any] | None) -> dict[str, Any]:
    result = copy.deepcopy(dict(options or {}))
    result.setdefault("schemaVersion", CURRENT_SCHEMA_VERSION)
    return result
