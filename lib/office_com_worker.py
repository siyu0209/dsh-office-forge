"""Isolated Windows COM worker for Office Forge.

The document engine itself runs on the bundled Python runtime, which does not
ship pywin32.  This worker is launched only through an explicitly selected
system Python that has ``win32com`` and installed Microsoft Office.  It is a
separate process so COM crashes, modal dialogs and Office startup problems can
never take down the DSH host or the native document engine.

Protocol: one JSON object on stdin, one JSON object on stdout.
"""

from __future__ import annotations

import json
import os
import sys
import traceback
from typing import Any


def _write(result: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(result, ensure_ascii=False, default=str))
    sys.stdout.flush()


def _error(code: str, message: str, detail: Any = None) -> dict[str, Any]:
    return {"ok": False, "error": {"code": code, "message": message, "detail": detail}}


_APPLICATIONS: list[Any] = []


def _application_pid(application: Any) -> int | None:
    """Resolve the Office process id behind a COM automation object."""
    for attribute in ("Hwnd",):
        try:
            hwnd = getattr(application, attribute)
        except Exception:
            continue
        if not hwnd:
            continue
        try:
            import win32process  # type: ignore

            _thread, pid = win32process.GetWindowThreadProcessId(int(hwnd))
            if pid:
                return int(pid)
        except Exception:
            continue
    return None


def _force_kill(pid: int) -> bool:
    """Last-resort reclamation for an Office process that refused to Quit."""
    try:
        import subprocess

        completed = subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True,
            text=True,
            timeout=20,
            check=False,
        )
        return completed.returncode == 0
    except Exception:
        return False


def _safe_quit(application: Any) -> None:
    """Best-effort cleanup; Office may already have closed a document.

    A refused or hung Quit is escalated to a forced process kill so a failed
    conversion never leaves a headless WINWORD/EXCEL/POWERPNT behind.
    """
    pid = _application_pid(application)
    try:
        application.Quit()
    except Exception:
        pass
    if pid:
        try:
            import win32process  # type: ignore

            handle = None
            try:
                import win32api  # type: ignore
                import win32con  # type: ignore

                handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            except Exception:
                handle = None
            if handle is not None:
                try:
                    if win32process.GetExitCodeProcess(handle) == 259:  # STILL_ACTIVE
                        _force_kill(pid)
                finally:
                    try:
                        win32api.CloseHandle(handle)
                    except Exception:
                        pass
        except Exception:
            pass


def _register(application: Any) -> Any:
    _APPLICATIONS.append(application)
    return application


def _dispatch(kind: str, suffix: str) -> Any:
    import win32com.client  # type: ignore

    if suffix in (".doc", ".docm", ".docx"):
        application = win32com.client.DispatchEx("Word.Application")
        application.Visible = False
        application.DisplayAlerts = 0
        return _register(application)
    if suffix in (".xls", ".xlsm", ".xlsx"):
        application = win32com.client.DispatchEx("Excel.Application")
        application.Visible = False
        application.DisplayAlerts = False
        return _register(application)
    application = win32com.client.DispatchEx("PowerPoint.Application")
    # PowerPoint refuses to hide its main window; keep it visible but isolated.
    try:
        application.Visible = True
    except Exception:
        pass
    return _register(application)


def _convert(payload: dict[str, Any]) -> dict[str, Any]:
    import win32com.client  # type: ignore

    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    extension = os.path.splitext(source)[1].lower()
    target_extension = os.path.splitext(target)[1].lower()
    application = _dispatch("convert", extension)
    try:
        if extension in (".doc", ".docm", ".docx"):
            document = application.Documents.Open(source, ReadOnly=True)
            try:
                format_id = 16 if target_extension == ".docx" else 12
                document.SaveAs2(target, FileFormat=format_id)
            finally:
                document.Close(False)
        elif extension in (".xls", ".xlsm", ".xlsx"):
            workbook = application.Workbooks.Open(source, ReadOnly=True)
            try:
                format_id = 51 if target_extension == ".xlsx" else 56
                workbook.SaveAs(target, FileFormat=format_id)
            finally:
                workbook.Close(False)
        else:
            presentation = application.Presentations.Open(source, WithWindow=False)
            try:
                presentation.SaveAs(target, 24 if target_extension == ".pptx" else 1)
            finally:
                presentation.Close()
    finally:
        _safe_quit(application)
    if not os.path.isfile(target):
        raise RuntimeError(f"conversion did not produce {target}")
    return {"type": target_extension.lstrip("."), "path": target, "bytes": os.path.getsize(target), "backend": "com"}


def _export_pdf(payload: dict[str, Any]) -> dict[str, Any]:
    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    extension = os.path.splitext(source)[1].lower()
    application = _dispatch("export", extension)
    try:
        if extension in (".doc", ".docm", ".docx"):
            document = application.Documents.Open(source, ReadOnly=True)
            try:
                document.ExportAsFixedFormat(OutputFileName=target, ExportFormat=17)
            finally:
                document.Close(False)
        elif extension in (".xls", ".xlsm", ".xlsx"):
            workbook = application.Workbooks.Open(source, ReadOnly=True)
            try:
                workbook.ExportAsFixedFormat(0, target)
            finally:
                workbook.Close(False)
        else:
            presentation = application.Presentations.Open(source, WithWindow=False)
            try:
                presentation.SaveAs(target, 32)
            finally:
                presentation.Close()
    finally:
        _safe_quit(application)
    if not os.path.isfile(target):
        raise RuntimeError(f"PDF export did not produce {target}")
    return {"type": "pdf", "path": target, "bytes": os.path.getsize(target), "backend": "com"}


def _accept_revisions(payload: dict[str, Any]) -> dict[str, Any]:
    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    application = _dispatch("revisions", ".docx")
    try:
        document = application.Documents.Open(source, ReadOnly=False)
        try:
            document.Revisions.AcceptAll()
            document.SaveAs2(target, FileFormat=16)
        finally:
            document.Close(False)
    finally:
        _safe_quit(application)
    return {"type": "docx", "path": target, "bytes": os.path.getsize(target), "backend": "com", "accepted": True}


def _recalculate(payload: dict[str, Any]) -> dict[str, Any]:
    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    application = _dispatch("recalculate", ".xlsx")
    try:
        workbook = application.Workbooks.Open(source)
        try:
            application.CalculateFullRebuild()
            workbook.SaveAs(target, FileFormat=51)
        finally:
            workbook.Close(False)
    finally:
        _safe_quit(application)
    return {"type": "xlsx", "path": target, "bytes": os.path.getsize(target), "backend": "com", "recalculated": True}


def _pivot(payload: dict[str, Any]) -> dict[str, Any]:
    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    spec = payload.get("spec") or {}
    source_sheet = str(spec.get("sourceSheet") or "")
    source_range = str(spec.get("sourceRange") or "")
    if not source_sheet or not source_range:
        raise ValueError("pivot spec requires sourceSheet and sourceRange")
    row_fields = [str(item) for item in (spec.get("rowFields") or [])]
    column_fields = [str(item) for item in (spec.get("columnFields") or [])]
    value_fields = spec.get("valueFields") or []
    if not value_fields:
        raise ValueError("pivot spec requires at least one valueFields entry")
    application = _dispatch("pivot", ".xlsx")
    try:
        workbook = application.Workbooks.Open(source)
        try:
            sheet = workbook.Worksheets(source_sheet)
            target_name = str(spec.get("targetSheet") or "OfficeForgePivot")
            try:
                destination = workbook.Worksheets(target_name)
                destination.Cells.Clear()
            except Exception:
                destination = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
                destination.Name = target_name
            cache = workbook.PivotCaches().Create(SourceType=1, SourceData=sheet.Range(source_range))
            table = cache.CreatePivotTable(TableDestination=destination.Range("A3"), TableName=str(spec.get("tableName") or "OfficeForgePivotTable"))
            for name in row_fields:
                table.PivotFields(name).Orientation = 1
            for name in column_fields:
                table.PivotFields(name).Orientation = 2
            for item in value_fields:
                if isinstance(item, dict):
                    field_name = str(item.get("field") or item.get("name") or "")
                    function = str(item.get("function") or "sum").lower()
                else:
                    field_name = str(item)
                    function = "sum"
                if not field_name:
                    continue
                function_id = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}.get(function, -4157)
                table.AddDataField(table.PivotFields(field_name), f"{function.upper()} of {field_name}", function_id)
            workbook.SaveAs(target, FileFormat=51)
        finally:
            workbook.Close(False)
    finally:
        _safe_quit(application)
    return {"type": "xlsx", "path": target, "bytes": os.path.getsize(target), "backend": "com", "nativePivot": True}


def _smart_art(payload: dict[str, Any]) -> dict[str, Any]:
    source = os.path.abspath(str(payload["source"]))
    target = os.path.abspath(str(payload["target"]))
    spec = payload.get("spec") or {}
    application = _dispatch("smartart", ".pptx")
    try:
        presentation = application.Presentations.Open(source, WithWindow=False)
        try:
            slide_index = max(1, int(spec.get("slide") or 1))
            slide = presentation.Slides(slide_index)
            layout_name = str(spec.get("layout") or "").strip()
            layout = None
            layouts = application.SmartArtLayouts
            for index in range(1, int(layouts.Count) + 1):
                candidate = layouts.Item(index)
                if layout_name and str(candidate.Name).lower() == layout_name.lower():
                    layout = candidate
                    break
            if layout is None:
                # Office localizations expose translated layout names; the first
                # layout is a stable fallback when the caller's English name
                # does not match the installed UI language.
                layout = layouts.Item(1)
            left = float(spec.get("left") or 80)
            top = float(spec.get("top") or 100)
            width = float(spec.get("width") or 600)
            height = float(spec.get("height") or 300)
            shape = slide.Shapes.AddSmartArt(layout, left, top, width, height)
            text = str(spec.get("text") or "")
            text_applied = False
            if text:
                for setter in (
                    lambda: setattr(shape.TextFrame2.TextRange, "Text", text),
                    lambda: setattr(shape.TextFrame.TextRange, "Text", text),
                ):
                    try:
                        setter()
                        text_applied = True
                        break
                    except Exception:
                        continue
            presentation.SaveAs(target, 24)
        finally:
            presentation.Close()
    finally:
        _safe_quit(application)
    return {"type": "pptx", "path": target, "bytes": os.path.getsize(target), "backend": "com", "nativeSmartArt": True}


_ACTIONS = {
    "convert": _convert,
    "exportPdf": _export_pdf,
    "acceptRevisions": _accept_revisions,
    "recalculate": _recalculate,
    "pivot": _pivot,
    "smartArt": _smart_art,
}


def _execute(payload: dict[str, Any]) -> dict[str, Any]:
    action = str(payload.get("action") or "")
    function = _ACTIONS.get(action)
    if function is None:
        return _error("UNSUPPORTED_OPERATION", f"unsupported COM action: {action or '(empty)'}", {"allowed": sorted(_ACTIONS)})
    result = function(payload)
    return {"ok": True, "result": result}


def _release_all() -> None:
    """Quit every automation application this worker started."""
    while _APPLICATIONS:
        application = _APPLICATIONS.pop()
        _safe_quit(application)


def main() -> None:
    try:
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        raw = sys.stdin.read()
        payload = json.loads(raw) if raw.strip() else {}
        if not isinstance(payload, dict):
            raise ValueError("request must be a JSON object")
        # Import lazily so a machine without pywin32 gets a clean error instead
        # of an import traceback before the protocol starts.
        import pythoncom  # type: ignore  # noqa: F401
        import win32com.client  # type: ignore  # noqa: F401
        pythoncom.CoInitialize()
        try:
            result = _execute(payload)
        finally:
            try:
                _release_all()
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        _write(result)
    except ImportError as error:
        _write(_error("BACKEND_UNAVAILABLE", f"pywin32 is unavailable: {error}"))
    except Exception as error:  # noqa: BLE001 - protocol must always return JSON
        try:
            _release_all()
        except Exception:
            pass
        _write(_error("COM_ERROR", f"{type(error).__name__}: {error}", {"traceback": traceback.format_exc()[-4000:]}))


if __name__ == "__main__":
    main()
