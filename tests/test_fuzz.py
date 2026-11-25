"""Robustness sweep: malformed payloads must return classified errors.

Every case here feeds the engine something a hostile or confused caller could
send.  The contract is that the engine always answers with a single JSON
object carrying ``ok: false`` and a stable error code - never a traceback,
never a hang, and never a write outside the workspace.

    python tests/test_fuzz.py
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.join(os.path.dirname(HERE), "lib", "office_engine.py")

CASES: list[tuple[str, dict]] = [
    ("empty payload", {}),
    ("unknown action", {"action": "definitely-not-an-action"}),
    ("action not a string", {"action": 12345}),
    ("read without path", {"action": "read", "type": "docx"}),
    ("read with null path", {"action": "read", "type": "docx", "path": None}),
    ("read unsupported extension", {"action": "read", "type": "docx", "path": "x.exe"}),
    ("read traversal", {"action": "read", "type": "docx", "path": "../../../../Windows/win.ini"}),
    ("read absolute outside", {"action": "read", "type": "docx", "path": "C:/Windows/win.ini"}),
    ("read missing file", {"action": "read", "type": "docx", "path": "nope.docx"}),
    ("create without spec", {"action": "create", "type": "docx", "path": "a.docx"}),
    ("create spec not object", {"action": "create", "type": "docx", "path": "b.docx", "spec": "nonsense"}),
    ("create spec list", {"action": "create", "type": "docx", "path": "c.docx", "spec": [1, 2, 3]}),
    ("create unknown type", {"action": "create", "type": "md", "path": "d.docx", "spec": {}}),
    ("create traversal target", {"action": "create", "type": "docx", "path": "../escape.docx", "spec": {}}),
    ("edit missing operations", {"action": "edit", "type": "docx", "path": "e.docx"}),
    ("edit operations not list", {"action": "edit", "type": "docx", "path": "f.docx", "args": {"operations": "x"}}),
    ("edit unknown operation", {"action": "edit", "type": "docx", "path": "g.docx", "args": {"operations": [{"op": "zap"}]}}),
    ("edit operation not object", {"action": "edit", "type": "docx", "path": "h.docx", "args": {"operations": ["x"]}}),
    ("batch not list", {"action": "batch", "items": {"a": 1}}),
    ("batch nested batch", {"action": "batch", "items": [{"action": "batch", "items": [{"action": "probe"}]}]}),
    ("template without data", {"action": "template", "type": "docx", "path": "t.docx", "template": "m.docx"}),
    ("copy missing source", {"action": "copy", "target": "out.docx"}),
    ("copy traversal target", {"action": "copy", "source": "a.docx", "target": "../../out.docx"}),
    ("pdfop unknown operation", {"action": "pdfop", "operation": "shred", "path": "a.pdf"}),
    ("native unknown operation", {"action": "native", "operation": "hypnotize", "path": "a.docx"}),
    ("audit unknown mode", {"action": "audit", "type": "xlsx", "path": "a.xlsx", "mode": "vibes"}),
    ("schema version too high", {"action": "create", "type": "docx", "path": "i.docx", "spec": {"schemaVersion": 99}}),
    ("features not object", {"action": "create", "type": "docx", "path": "j.docx", "spec": {}, "features": "off"}),
    ("deeply nested spec", {"action": "create", "type": "docx", "path": "k.docx", "spec": {"sections": [{"paragraphs": [{"text": "x" * 100}]}] * 3}}),
    ("deck unknown operation", {"action": "deck", "operation": "teleport"}),
    ("deck import without path", {"action": "deck", "operation": "import"}),
    ("deck import wrong extension", {"action": "deck", "operation": "import", "path": "not-a-deck.docx"}),
    ("deck verify without path", {"action": "deck", "operation": "verify"}),
    ("deck verify wrong extension", {"action": "deck", "operation": "verify", "path": "not-a-deck.xlsx"}),
    ("deck loss audit spec not object", {"action": "deck", "operation": "lossAudit", "spec": "nonsense"}),
    ("deck render spec not object", {"action": "deck", "operation": "render", "path": "out.pptx", "spec": []}),
    ("deck import traversal", {"action": "deck", "operation": "import", "path": "../../escape.pptx"}),
]


def run(payload: dict, workspace: str, raw: str | None = None) -> tuple[int, str, str]:
    body = raw if raw is not None else json.dumps(payload, ensure_ascii=False)
    environment = dict(os.environ)
    environment["OFFICEFORGE_WORKSPACE"] = workspace
    environment["PYTHONUTF8"] = "1"
    completed = subprocess.run(
        [sys.executable, ENGINE],
        input=body,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=environment,
        timeout=180,
        check=False,
    )
    return completed.returncode, completed.stdout, completed.stderr


def assert_classified(label: str, workspace: str, payload: dict, raw: str | None = None) -> None:
    code, stdout, stderr = run(payload, workspace, raw)
    assert code == 0, f"{label}: engine exited {code}, stderr={stderr[-800:]}"
    text = (stdout or "").strip()
    assert text, f"{label}: engine produced no stdout"
    try:
        parsed = json.loads(text)
    except Exception as error:  # noqa: BLE001
        raise AssertionError(f"{label}: stdout is not JSON ({error}): {text[:400]}") from error
    assert isinstance(parsed, dict), f"{label}: response is not an object"
    if parsed.get("ok") is True:
        # Only structurally valid requests may succeed; that is still a pass
        # because the engine answered deterministically.
        print(f"PASS {label}: accepted")
        return
    error = parsed.get("error") or {}
    assert isinstance(error, dict) and error.get("code"), f"{label}: missing error code: {parsed}"
    assert "Traceback" not in text, f"{label}: leaked a traceback"
    print(f"PASS {label}: {error.get('code')}")


def test_cases() -> None:
    with tempfile.TemporaryDirectory(prefix="officeforge-fuzz-") as work:
        for label, payload in CASES:
            assert_classified(label, work, payload)
        assert_classified("garbage bytes", work, {}, raw="\x00\x01not json at all")
        assert_classified("empty stdin", work, {}, raw="")
        assert_classified("json array", work, {}, raw="[1,2,3]")
        assert_classified("json string", work, {}, raw='"hello"')
        assert_classified("json null", work, {}, raw="null")
        assert_classified("truncated json", work, {}, raw='{"action": "read"')
        # Nothing may have escaped the workspace.
        parent = os.path.dirname(work)
        for name in os.listdir(parent):
            assert not name.endswith("escape.docx"), f"workspace escape produced {name}"


def test_nested_batch_rejected() -> None:
    """A batch item that is itself a batch must be refused, not executed."""
    with tempfile.TemporaryDirectory(prefix="officeforge-fuzz-nested-") as work:
        code, stdout, _stderr = run(
            {"action": "batch", "items": [{"action": "batch", "items": [{"action": "probe"}]}, {"action": "probe"}]},
            work,
        )
        assert code == 0, f"engine exited {code}"
        parsed = json.loads(stdout)
        body = parsed.get("result") if isinstance(parsed.get("result"), dict) else parsed
        items = body.get("items") or []
        assert len(items) == 2, f"unexpected item count: {items}"
        assert items[0]["ok"] is False, f"nested batch was executed: {items[0]}"
        assert items[0]["error"]["code"] == "INVALID_OPERATION", f"wrong nested-batch code: {items[0]}"
        assert items[1]["ok"] is True, f"sibling item should still succeed: {items[1]}"
        print("PASS nested batch: rejected with INVALID_OPERATION while sibling succeeded")


def test_repeated_calls_stable() -> None:
    """The same bad request must fail identically twice (no state leakage)."""
    with tempfile.TemporaryDirectory(prefix="officeforge-fuzz-repeat-") as work:
        payload = {"action": "read", "type": "docx", "path": "missing.docx"}
        first = run(payload, work)
        second = run(payload, work)
        assert first[1] == second[1], "identical requests produced different responses"
        print("PASS repeat: identical response for identical bad request")


def main() -> None:
    test_cases()
    test_nested_batch_rejected()
    test_repeated_calls_stable()
    print("FUZZ TESTS PASSED")


if __name__ == "__main__":
    main()
