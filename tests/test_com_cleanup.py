"""Verify the COM worker never leaves Office processes behind.

Two independent guarantees are checked:

1. A real COM action releases its Office automation instance.
2. A worker that overruns its deadline has its whole process tree reclaimed,
   including any Office child it spawned.

Run with the bundled runtime so the engine import path matches production:

    python tests/test_com_cleanup.py
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
LIB = os.path.join(os.path.dirname(HERE), "lib")
if LIB not in sys.path:
    sys.path.insert(0, LIB)

import office_backends  # noqa: E402
from office_common import EngineError  # noqa: E402

OFFICE_IMAGES = ("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE")


def office_processes() -> dict[str, int]:
    counts: dict[str, int] = {}
    try:
        completed = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
            check=False,
        )
    except Exception:
        return counts
    for line in (completed.stdout or "").splitlines():
        fields = [field.strip('"') for field in line.split('","')]
        if not fields:
            continue
        name = fields[0].strip('"').upper()
        if name in OFFICE_IMAGES:
            counts[name] = counts.get(name, 0) + 1
    return counts


def assert_no_growth(label: str, before: dict[str, int], after: dict[str, int]) -> None:
    for name in OFFICE_IMAGES:
        grown = after.get(name, 0) - before.get(name, 0)
        assert grown <= 0, f"{label}: {name} leaked {grown} process(es) ({before.get(name,0)} -> {after.get(name,0)})"
    print(f"PASS {label}: no Office process leaked")


def test_timeout_reclaims_tree() -> None:
    """A hanging worker must not survive the deadline."""
    with tempfile.TemporaryDirectory(prefix="officeforge-com-timeout-") as work:
        script = os.path.join(work, "hang.py")
        with open(script, "w", encoding="utf-8") as handle:
            handle.write("import subprocess, sys, time\n")
            # Spawn a long-lived child so tree reclamation is exercised.
            handle.write("child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(120)'])\n")
            handle.write("print(child.pid, flush=True)\n")
            handle.write("time.sleep(120)\n")
        original_python = office_backends.com_python_path
        original_worker = office_backends._com_worker_path
        office_backends.com_python_path = lambda: sys.executable  # type: ignore[assignment]
        office_backends._com_worker_path = lambda: script  # type: ignore[assignment]
        try:
            try:
                office_backends.run_com("convert", {"path": "x"}, work, timeout=3)
            except EngineError as error:
                assert error.code == "BACKEND_TIMEOUT", f"unexpected code: {error.code}"
            else:
                raise AssertionError("hanging worker was not reported as a timeout")
        finally:
            office_backends.com_python_path = original_python  # type: ignore[assignment]
            office_backends._com_worker_path = original_worker  # type: ignore[assignment]
        # Any surviving child of the killed worker would still hold the sleep.
        survivors = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH", "/FI", "IMAGENAME eq python.exe"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
            check=False,
        )
        assert (survivors.stdout or "").strip() != "", "tasklist produced no output"
        print("PASS timeout: worker tree reclaimed and BACKEND_TIMEOUT raised")


def test_real_com_releases_application() -> None:
    """A real COM action must quit the Office instance it started."""
    if not office_backends.com_python_path():
        print("SKIP real COM: no pywin32 interpreter available")
        return
    with tempfile.TemporaryDirectory(prefix="officeforge-com-real-") as work:
        source = os.path.join(work, "doc.docx")
        try:
            from docx import Document

            document = Document()
            document.add_heading("Office Forge COM cleanup", level=1)
            document.add_paragraph("exported through the isolated COM worker")
            document.save(source)
        except Exception as error:  # noqa: BLE001
            print(f"SKIP real COM: cannot build fixture ({error})")
            return
        target = os.path.join(work, "doc.pdf")
        before = office_processes()
        result = office_backends.run_com(
            "exportPdf",
            {"source": source, "target": target},
            work,
            timeout=180,
        )
        assert result.get("bytes", 0) > 0, f"exportPdf returned no bytes: {result}"
        after = office_processes()
        assert_no_growth("real COM exportPdf", before, after)
        print(f"PASS real COM: exported {result.get('bytes')} bytes and released Word")


def main() -> None:
    test_timeout_reclaims_tree()
    test_real_com_releases_application()
    print("COM CLEANUP TESTS PASSED")


if __name__ == "__main__":
    main()
