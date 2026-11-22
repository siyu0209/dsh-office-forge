"""Optional native backends (Windows COM / LibreOffice) for Office Forge.

Core document operations never require these backends.  They are discovered
at runtime and exposed only as optional capabilities; a missing backend is a
classified ``BACKEND_UNAVAILABLE`` result, never an import-time failure.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from typing import Any

from office_common import EngineError, assert_inside

_COM_PYTHON_CACHE: dict[str, Any] = {"checked": False, "path": None, "error": None}

_LEGACY_TARGET = {
    ".doc": ".docx",
    ".docm": ".docx",
    ".xls": ".xlsx",
    ".xlsm": ".xlsx",
    ".ppt": ".pptx",
    ".pptm": ".pptx",
}


def _com_python_candidates() -> list[str]:
    candidates: list[str] = []
    for variable in ("OFFICEFORGE_COM_PYTHON", "DSH_COM_PYTHON"):
        value = os.environ.get(variable)
        if value:
            candidates.append(value)
    found = shutil.which("python")
    if found:
        candidates.append(found)
    candidates.extend((
        r"C:\Program Files\Python311\python.exe",
        r"C:\Program Files\Python310\python.exe",
        r"C:\Program Files\Python312\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python312\python.exe",
    ))
    result: list[str] = []
    for candidate in candidates:
        absolute = os.path.abspath(candidate)
        if absolute not in result and os.path.isfile(absolute):
            result.append(absolute)
    return result


def com_python_path() -> str | None:
    """Find a Python interpreter that has pywin32 and can run COM workers."""
    if _COM_PYTHON_CACHE["checked"]:
        return _COM_PYTHON_CACHE["path"]
    _COM_PYTHON_CACHE["checked"] = True
    for candidate in _com_python_candidates():
        try:
            completed = subprocess.run(
                [candidate, "-c", "import sys, win32com.client; print(sys.executable)"],
                capture_output=True,
                text=True,
                timeout=20,
                check=False,
            )
        except Exception as error:  # noqa: BLE001
            _COM_PYTHON_CACHE["error"] = f"{type(error).__name__}: {error}"
            continue
        if completed.returncode == 0 and completed.stdout.strip():
            _COM_PYTHON_CACHE["path"] = completed.stdout.strip().splitlines()[-1]
            return _COM_PYTHON_CACHE["path"]
        _COM_PYTHON_CACHE["error"] = (completed.stderr or completed.stdout or "").strip()[-500:]
    return None


def _com_available() -> bool:
    return com_python_path() is not None


def _com_worker_path() -> str:
    """Location of the isolated COM worker script."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "office_com_worker.py")


def _kill_process_tree(pid: int) -> None:
    """Terminate a worker and every Office child it may have spawned."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True,
            text=True,
            timeout=20,
            check=False,
        )
    except Exception:
        pass


def run_com(action: str, payload: dict[str, Any], workspace: str, timeout: int = 120) -> dict[str, Any]:
    """Run one COM action in an isolated system-Python worker process."""
    python_path = com_python_path()
    if not python_path:
        raise EngineError(
            "BACKEND_UNAVAILABLE",
            "no Python interpreter with pywin32 (win32com) is available",
            {"action": action, "candidates": _com_python_candidates(), "error": _COM_PYTHON_CACHE.get("error")},
        )
    worker = _com_worker_path()
    request = dict(payload)
    request["action"] = action
    environment = dict(os.environ)
    environment["OFFICEFORGE_WORKSPACE"] = os.path.abspath(workspace)
    environment["PYTHONUTF8"] = "1"
    try:
        process = subprocess.Popen(
            [python_path, worker],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=environment,
        )
    except Exception as error:  # noqa: BLE001
        raise EngineError("COM_ERROR", f"failed to start COM worker: {type(error).__name__}: {error}", {"action": action}) from error
    try:
        stdout, stderr = process.communicate(json.dumps(request, ensure_ascii=False), timeout=timeout)
    except subprocess.TimeoutExpired as error:
        # Killing only the Python worker would orphan WINWORD/EXCEL/POWERPNT,
        # so the whole process tree is reclaimed before reporting the timeout.
        _kill_process_tree(process.pid)
        try:
            process.communicate(timeout=15)
        except Exception:
            pass
        raise EngineError("BACKEND_TIMEOUT", f"COM worker timed out after {timeout}s", {"action": action}) from error
    raw = (stdout or "").strip()
    response: dict[str, Any] = {}
    if raw:
        for line in reversed(raw.splitlines()):
            try:
                candidate = json.loads(line)
                if isinstance(candidate, dict):
                    response = candidate
                    break
            except Exception:
                continue
    if not response:
        raise EngineError(
            "COM_ERROR",
            "COM worker returned no JSON response",
            {"action": action, "returncode": process.returncode, "stderr": (stderr or "")[-2000:]},
        )
    if not response.get("ok"):
        error = response.get("error") if isinstance(response.get("error"), dict) else {}
        code = str(error.get("code") or "COM_ERROR")
        message = str(error.get("message") or "COM operation failed")
        raise EngineError(code, message, {"action": action, "detail": error.get("detail"), "returncode": process.returncode})
    return response.get("result") if isinstance(response.get("result"), dict) else {}


def _libreoffice_path() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if os.path.isfile(candidate):
            return candidate
    return None


def _pdf_backends() -> dict[str, bool]:
    result = {}
    vendor_pypdf = os.path.isdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor", "pypdf"))
    for name in ("pypdf", "fitz", "pdfplumber", "pdfminer"):
        try:
            __import__(name)
            result[name] = True
        except Exception:
            # The package ships pypdf in lib/vendor so the PDF toolkit remains
            # available on a clean DSH runtime without site-packages changes.
            result[name] = bool(name == "pypdf" and vendor_pypdf)
    return result


def _ocr_status() -> dict[str, Any]:
    """Report the optional OCR stack without importing heavy modules eagerly.

    OCR is deliberately optional: the core document tools must keep working on
    a clean DSH runtime.  The doctor surfaces the exact missing component so a
    user can install only what they need.
    """
    pytesseract_ok = False
    pytesseract_error = None
    version = None
    try:
        import pytesseract  # type: ignore

        pytesseract_ok = True
        try:
            version = str(pytesseract.get_tesseract_version())
        except Exception as error:  # noqa: BLE001
            pytesseract_error = f"{type(error).__name__}: {error}"
    except Exception as error:  # noqa: BLE001
        pytesseract_error = f"{type(error).__name__}: {error}"
    fitz_ok = False
    try:
        import fitz  # type: ignore  # noqa: F401

        fitz_ok = True
    except Exception:
        fitz_ok = False
    tesseract_path = shutil.which("tesseract")
    languages: list[str] = []
    if tesseract_path:
        try:
            completed = subprocess.run(
                [tesseract_path, "--list-langs"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=15,
                check=False,
            )
            if completed.returncode == 0:
                for line in (completed.stdout or "").splitlines():
                    value = line.strip()
                    if value and "list of available languages" not in value.lower():
                        languages.append(value)
        except Exception:
            languages = []
    return {
        "available": bool(pytesseract_ok and fitz_ok and tesseract_path),
        "pytesseract": pytesseract_ok,
        "fitz": fitz_ok,
        "tesseract": bool(tesseract_path),
        "tesseractPath": tesseract_path,
        "version": version,
        "languages": languages,
        "error": pytesseract_error,
    }


def _module_versions() -> dict[str, str | None]:
    """Report the Python packages the engine actually imports.

    Modules are imported defensively so this function can also run when a
    dependency is missing and the engine cannot start.
    """
    result: dict[str, str | None] = {}
    for name in ("docx", "openpyxl", "pptx", "pypdf", "fitz", "pdfplumber", "pdfminer"):
        try:
            module = __import__(name)
            result[name] = str(getattr(module, "__version__", "unknown"))
        except Exception:
            result[name] = None
    return result


def _office_apps() -> dict[str, bool]:
    """Check whether Word/Excel/PowerPoint COM registrations exist.

    This is a registry probe only; it never starts Office or creates a COM
    object.  On non-Windows hosts it returns an empty mapping.
    """
    if os.name != "nt":
        return {}
    try:
        import winreg
    except Exception:
        return {}
    result: dict[str, bool] = {}
    for name, prog_id in (
        ("word", "Word.Application"),
        ("excel", "Excel.Application"),
        ("powerpoint", "PowerPoint.Application"),
    ):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + "\\CLSID"):
                result[name] = True
        except OSError:
            result[name] = False
    return result


def status() -> dict[str, Any]:
    return {
        "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        "modules": _module_versions(),
        "com": {
            "available": _com_available(),
            "platform": "windows",
            "python": com_python_path(),
            "office": _office_apps(),
        },
        "libreoffice": {"available": _libreoffice_path() is not None, "path": _libreoffice_path()},
        "pdf": _pdf_backends(),
        "ocr": _ocr_status(),
    }


def convert_legacy(path: str, output: str, workspace: str, timeout: int = 120) -> dict[str, Any]:
    """Convert .doc/.xls/.ppt through COM or LibreOffice when available."""
    source = os.path.abspath(path)
    target = os.path.abspath(output)
    assert_inside(source, workspace, "source")
    assert_inside(target, workspace, "output")
    extension = os.path.splitext(source)[1].lower()
    target_extension = _LEGACY_TARGET.get(extension)
    if target_extension is None:
        raise EngineError("UNSUPPORTED_FORMAT", f"legacy conversion does not support {extension or '(none)'}", {"supported": sorted(_LEGACY_TARGET)})
    if os.path.exists(target):
        raise EngineError("OUTPUT_EXISTS", f"{target} already exists", {"path": target})
    if _com_available():
        try:
            return run_com("convert", {"source": source, "target": target}, workspace, timeout)
        except EngineError as error:
            # If Office itself is unavailable, still try LibreOffice below.
            if error.code not in {"COM_ERROR", "BACKEND_ERROR"}:
                raise
    soffice = _libreoffice_path()
    if soffice:
        command = [soffice, "--headless", "--convert-to", target_extension.lstrip("."), "--outdir", os.path.dirname(target) or ".", source]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
        except Exception as error:
            raise EngineError("BACKEND_ERROR", f"LibreOffice conversion failed: {type(error).__name__}: {error}") from error
        if completed.returncode != 0 or not os.path.isfile(target):
            raise EngineError("BACKEND_ERROR", "LibreOffice conversion did not produce the target", {"returncode": completed.returncode, "stderr": (completed.stderr or "")[-2000:]})
        return {"type": target_extension.lstrip("."), "path": target, "bytes": os.path.getsize(target), "backend": "libreoffice"}
    raise EngineError("BACKEND_UNAVAILABLE", "no COM or LibreOffice backend is available for legacy conversion", {"extension": extension, "backends": status()})


def export_pdf(path: str, output: str, workspace: str, timeout: int = 120) -> dict[str, Any]:
    source = os.path.abspath(path)
    target = os.path.abspath(output)
    assert_inside(source, workspace, "source")
    assert_inside(target, workspace, "output")
    if os.path.exists(target):
        raise EngineError("OUTPUT_EXISTS", f"{target} already exists", {"path": target})
    if _com_available():
        return run_com("exportPdf", {"source": source, "target": target}, workspace, timeout)
    soffice = _libreoffice_path()
    if soffice:
        target_extension = "pdf"
        command = [soffice, "--headless", "--convert-to", target_extension, "--outdir", os.path.dirname(target) or ".", source]
        completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
        if completed.returncode != 0 or not os.path.isfile(target):
            raise EngineError("BACKEND_ERROR", "LibreOffice PDF export did not produce the target", {"returncode": completed.returncode, "stderr": (completed.stderr or "")[-2000:]})
        return {"type": "pdf", "path": target, "bytes": os.path.getsize(target), "backend": "libreoffice"}
    raise EngineError("BACKEND_UNAVAILABLE", "no COM or LibreOffice backend is available for PDF export", {"backends": status()})
