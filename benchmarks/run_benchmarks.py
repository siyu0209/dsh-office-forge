#!/usr/bin/env python3
"""Deterministic performance and robustness baseline for Office Forge.

Run with the bundled DSH Python:
    python benchmarks/run_benchmarks.py
    python benchmarks/run_benchmarks.py --full

Each case is a real engine invocation.  The harness records wall time, peak
working set on Windows, and verifies either success or a typed error.  The
thresholds live in benchmarks/baseline.json so a regression makes the full
``tests/run_regression.py`` run fail.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE = ROOT / "lib" / "office_engine.py"
BASELINE_PATH = Path(__file__).resolve().parent / "baseline.json"
PYTHON = sys.executable
WORK = Path(tempfile.mkdtemp(prefix="officeforge-bench-"))
sys.path.insert(0, str(ROOT / "lib" / "vendor"))


def _peak_memory_mb(pid: int) -> float:
    """Return peak working set for a Windows process, or 0 elsewhere."""
    if os.name != "nt":
        return 0.0

    class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
        _fields_ = [
            ("cb", ctypes.c_ulong),
            ("PageFaultCount", ctypes.c_ulong),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    handle = ctypes.windll.kernel32.OpenProcess(0x0410, False, pid)
    if not handle:
        return 0.0
    try:
        counters = PROCESS_MEMORY_COUNTERS()
        counters.cb = ctypes.sizeof(PROCESS_MEMORY_COUNTERS)
        if ctypes.windll.psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb):
            return counters.PeakWorkingSetSize / (1024.0 * 1024.0)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)
    return 0.0


def call_engine(payload: dict, timeout: int = 600) -> dict:
    env = os.environ.copy()
    env.update({
        "OFFICEFORGE_WORKSPACE": str(WORK),
        "PYTHONUTF8": "1",
        "PYTHONIOENCODING": "utf-8",
    })
    started = time.perf_counter()
    process = subprocess.Popen(
        [PYTHON, str(ENGINE)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(WORK),
        env=env,
    )
    peak = [0.0]
    stop = threading.Event()

    def sampler() -> None:
        while not stop.is_set() and process.poll() is None:
            current = _peak_memory_mb(process.pid)
            if current > peak[0]:
                peak[0] = current
            stop.wait(0.05)

    thread = threading.Thread(target=sampler, daemon=True)
    thread.start()
    try:
        stdout, stderr = process.communicate(
            json.dumps(payload, ensure_ascii=False), timeout=timeout
        )
    finally:
        stop.set()
        thread.join(timeout=1.0)
    elapsed = time.perf_counter() - started
    parsed = {}
    try:
        parsed = json.loads(stdout.strip() or "{}")
    except json.JSONDecodeError:
        parsed = {"ok": False, "error": {"code": "INVALID_JSON", "message": stdout[:500]}}
    return {
        "ok": process.returncode == 0 and bool(parsed.get("ok")),
        "result": parsed,
        "stderr": stderr[-1000:],
        "seconds": elapsed,
        "peakMB": peak[0],
    }


def _docx_spec(paragraphs: int) -> dict:
    return {
        "title": "Office Forge benchmark",
        "sections": [{
            "heading": "Large section",
            "paragraphs": [{"text": f"Benchmark paragraph {i}: the quick brown fox jumps over the lazy dog."} for i in range(paragraphs)],
            "table": {
                "headers": ["A", "B", "C", "D", "E"],
                "rows": [[i, i * 2, i * 3, i * 4, i * 5] for i in range(200)],
            },
        }],
    }


def _xlsx_spec(rows: int, cols: int) -> dict:
    return {
        "sheets": [{
            "name": "Data",
            "columns": [{"header": f"C{i}", "width": 12} for i in range(cols)],
            "rows": [[i * c for c in range(cols)] for i in range(rows)],
            "freeze": "A2",
        }],
    }


def _pptx_spec(slides: int) -> dict:
    return {
        "slideSize": "16:9",
        "slides": [
            {
                "layout": "title" if i == 0 else "content",
                "title": f"Slide {i + 1}",
                "bullets": [f"Point {j + 1} for slide {i + 1}" for j in range(5)],
                "notes": f"Speaker notes for slide {i + 1}",
            }
            for i in range(slides)
        ],
    }


def _write_case(name: str, payload: dict, expected_error: str | None = None) -> tuple[bool, float, float, str]:
    response = call_engine(payload)
    result = response["result"]
    if expected_error:
        code = result.get("error", {}).get("code") if isinstance(result, dict) else None
        ok = (not response["ok"]) and code == expected_error
    else:
        ok = response["ok"]
    detail = ""
    if not ok:
        detail = json.dumps(result, ensure_ascii=False)[:500]
    return ok, response["seconds"], response["peakMB"], detail


def _run_case(name: str, payload: dict, expected_error: str | None = None) -> tuple[bool, float, float, str]:
    ok, seconds, peak, detail = _write_case(name, payload, expected_error)
    print(f"{'PASS' if ok else 'FAIL'} {name} seconds={seconds:.3f} peakMB={peak:.1f}" + (f" detail={detail}" if detail else ""))
    return ok, seconds, peak, detail


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--full", action="store_true", help="use 50k-row Excel / 300-slide PPT fixtures")
    parser.add_argument("--update-baseline", action="store_true", help="record measured values as baseline")
    args = parser.parse_args()

    rows = 50000 if args.full else 12000
    cols = 20 if args.full else 10
    slides = 300 if args.full else 120
    paragraphs = 30000 if args.full else 8000

    results: dict[str, dict] = {}
    cases: list[tuple[str, dict, str | None]] = [
        ("docx_create", {"action": "create", "type": "docx", "path": "bench.docx", "spec": _docx_spec(paragraphs)}, None),
        ("docx_read", {"action": "read", "path": "bench.docx", "format": "outline"}, None),
        ("docx_audit", {"action": "audit", "type": "docx", "mode": "profile", "path": "bench.docx"}, None),
        ("xlsx_create", {"action": "create", "type": "xlsx", "path": "bench.xlsx", "spec": _xlsx_spec(rows, cols)}, None),
        ("xlsx_read", {"action": "read", "path": "bench.xlsx", "format": "outline"}, None),
        ("xlsx_audit", {"action": "audit", "type": "xlsx", "mode": "health", "path": "bench.xlsx"}, None),
        ("pptx_create", {"action": "create", "type": "pptx", "path": "bench.pptx", "spec": _pptx_spec(slides)}, None),
        ("pptx_read", {"action": "read", "path": "bench.pptx", "format": "outline"}, None),
        ("pptx_audit", {"action": "audit", "type": "pptx", "mode": "deck", "path": "bench.pptx"}, None),
        ("pdf_corrupt", {"action": "read", "path": "corrupt.pdf", "format": "text"}, "INVALID_PDF"),
        ("pdf_empty", {"action": "read", "path": "empty.pdf", "format": "text"}, "INVALID_PDF"),
        ("pdf_encrypted_no_password", {"action": "read", "path": "encrypted.pdf", "format": "text"}, "PDF_PASSWORD_REQUIRED"),
    ]

    # Prepare small robustness fixtures.
    (WORK / "corrupt.pdf").write_bytes(b"not a pdf at all\n" * 100)
    (WORK / "empty.pdf").write_bytes(b"")
    try:
        from pypdf import PdfWriter
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        writer.encrypt("secret")
        with open(WORK / "encrypted.pdf", "wb") as handle:
            writer.write(handle)
    except Exception as exc:  # pragma: no cover - fixture environment issue
        print(f"FAIL fixture_encrypted_pdf {exc}")
        return 1

    baseline = {}
    if BASELINE_PATH.exists() and not args.update_baseline:
        baseline = json.loads(BASELINE_PATH.read_text(encoding="utf-8")).get("cases", {})

    failures = 0
    for name, payload, expected_error in cases:
        ok, seconds, peak, detail = _run_case(name, payload, expected_error)
        limits = baseline.get(name, {}) if not args.full else {}
        if limits:
            if seconds > float(limits.get("maxSeconds", 1e9)):
                print(f"FAIL {name} exceeded maxSeconds={limits['maxSeconds']} actual={seconds:.3f}")
                ok = False
            if peak and peak > float(limits.get("maxPeakMB", 1e9)):
                print(f"FAIL {name} exceeded maxPeakMB={limits['maxPeakMB']} actual={peak:.1f}")
                ok = False
        if not ok:
            failures += 1
        results[name] = {"seconds": round(seconds, 3), "peakMB": round(peak, 1)}

    if args.update_baseline:
        recorded = {
            "cases": {
                name: {
                    "maxSeconds": round(max(1.0, values["seconds"] * 1.5), 3),
                    "maxPeakMB": round(max(64.0, values["peakMB"] * 1.5), 1),
                }
                for name, values in results.items()
            }
        }
        BASELINE_PATH.write_text(json.dumps(recorded, indent=2) + "\n", encoding="utf-8")
        print(f"BASELINE UPDATED {BASELINE_PATH}")

    if failures:
        print(f"BENCHMARK FAILED cases={failures}")
        return 1
    print("BENCHMARK PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
