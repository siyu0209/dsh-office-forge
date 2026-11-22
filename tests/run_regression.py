"""Run every Office Forge suite and compare against the recorded baseline.

The baseline is a floor, not an equality check: adding tests is fine, silently
losing assertions is a regression.  Nothing here touches a DSH profile - the
profile test always builds its own temporary DSH_HOME.

    python tests/run_regression.py
    python tests/run_regression.py --quick    # skip benchmark, COM and profile suites
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
BASELINE = json.loads((ROOT / "tests" / "regression_baseline.json").read_text(encoding="utf-8"))
PYTHON = sys.executable
QUICK = "--quick" in sys.argv

# suite -> (command, per-assertion PASS regex or None, completion marker, baseline key)
SUITES: list[tuple[str, list[str], str | None, str, str]] = [
    ("engine", [PYTHON, "tests/test_engine.py"], r"^  PASS", "ALL TESTS PASSED", "enginePass"),
    ("fuzz", [PYTHON, "tests/test_fuzz.py"], r"^PASS ", "FUZZ TESTS PASSED", "fuzzPass"),
    ("plugin", ["node", "tests/test_plugin.mjs"], r"^  PASS", "PLUGIN TESTS PASSED", "pluginPass"),
    ("config", ["node", "tests/test_plugin_config.mjs"], r"^  PASS", "PLUGIN CONFIG TESTS PASSED", "configPass"),
    ("compatibility", ["node", "tests/test_compatibility.mjs"], r"^  PASS", "COMPATIBILITY TESTS PASSED", "compatibilityPass"),
    ("client", ["node", "tests/test_client.mjs"], None, "CLIENT UI TESTS PASSED", "clientPass"),
    ("clientHost", ["node", "tests/test_client_host.mjs"], None, "CLIENT HOST MACHINERY TESTS PASSED", "clientHostPass"),
]

OPTIONAL_SUITES: list[tuple[str, list[str], str | None, str, str]] = [
    ("benchmarks", [PYTHON, "benchmarks/run_benchmarks.py"], r"^PASS ", "BENCHMARK PASSED", "benchmarkPass"),
    ("comCleanup", [PYTHON, "tests/test_com_cleanup.py"], r"^PASS ", "COM CLEANUP TESTS PASSED", "comCleanupPass"),
    ("dshProfile", ["node", "tests/test_dsh_profile.mjs"], r"^  PASS", "DSH PROFILE TESTS PASSED", "dshProfilePass"),
]


def run_suite(label: str, command: list[str], pattern: str | None, marker: str, key: str) -> tuple[int, str]:
    completed = subprocess.run(
        command,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=900,
        check=False,
    )
    output = (completed.stdout or "") + "\n" + (completed.stderr or "")
    count = len(re.findall(pattern, output, re.MULTILINE)) if pattern else 0
    expected = int(BASELINE.get(key, 0))
    # A suite passes only when it exits cleanly, prints its completion marker,
    # and does not fall below the recorded assertion floor.
    ok = completed.returncode == 0 and marker in output and count >= expected
    status = "OK" if ok else "FAIL"
    detail = f"pass={count} baseline={expected}" if pattern else "marker only"
    print(f"[{status}] {label}: {detail} exit={completed.returncode}")
    if not ok:
        print(output[-4000:])
    return count, status


def main() -> None:
    failures: list[str] = []
    totals: dict[str, int] = {}
    for label, command, pattern, marker, key in SUITES:
        count, status = run_suite(label, command, pattern, marker, key)
        totals[label] = count
        if status == "FAIL":
            failures.append(label)
    if QUICK:
        print("quick mode: skipped comCleanup and dshProfile")
    else:
        for label, command, pattern, marker, key in OPTIONAL_SUITES:
            count, status = run_suite(label, command, pattern, marker, key)
            totals[label] = count
            if status == "FAIL":
                failures.append(label)
    summary = ", ".join(f"{name}={count}" for name, count in totals.items())
    print(f"TOTALS {summary}")
    if failures:
        raise SystemExit(f"REGRESSION FAILED: {', '.join(failures)}")
    print("REGRESSION BASELINE PASSED")


if __name__ == "__main__":
    main()
