"""Excel (.xlsx) reading, creation, editing and template expansion."""

from __future__ import annotations

import copy
import csv
import datetime as _datetime
import json
import os
import re
from typing import Any, Mapping

from openpyxl import Workbook, load_workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference, ScatterChart
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.table import Table, TableStyleInfo

from office_common import (
    EngineError,
    apply_core_properties,
    assert_input_file,
    assert_inside,
    atomic_save,
    get_path,
    jsonable,
    md_table,
    render_placeholders,
    require_dict,
    require_feature,
    require_list,
    scalar_text,
    truncate,
)

_ALLOWED = (".xlsx",)
_HEADER_FILL = "1F4E78"
_HEADER_FONT = "FFFFFF"


def _select_sheets(workbook: Any, selector: Any) -> list[Any]:
    if selector in (None, "", "all"):
        return list(workbook.worksheets)
    sheets = list(workbook.worksheets)
    if isinstance(selector, (int, float)):
        index = int(selector)
        if index < 0 or index >= len(sheets):
            raise EngineError("SHEET_NOT_FOUND", f"sheet index {index} is out of range", {"sheets": [sheet.title for sheet in sheets]})
        return [sheets[index]]
    name = str(selector)
    for sheet in sheets:
        if sheet.title == name:
            return [sheet]
    raise EngineError("SHEET_NOT_FOUND", f"sheet {name!r} was not found", {"sheets": [sheet.title for sheet in sheets]})


def _sheet_summary(sheet: Any) -> dict[str, Any]:
    return {
        "name": sheet.title,
        "maxRow": sheet.max_row,
        "maxColumn": sheet.max_column,
        "state": sheet.sheet_state,
        "freezePanes": getattr(sheet, "freeze_panes", None),
        "autoFilter": bool(getattr(getattr(sheet, "auto_filter", None), "ref", None)),
    }


def _cell_value(cell: Any, include_formulas: bool) -> Any:
    value = cell.value
    if value is None:
        return None
    if isinstance(value, ArrayFormula):
        return {"arrayFormula": value.text, "ref": value.ref}
    if isinstance(value, str) and value.startswith("=") and not include_formulas:
        return ""
    return jsonable(value)


def _read_sheet(sheet: Any, args: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    max_rows = min(int(args.get("maxRows") or 200), int(limits["maxRows"]))
    max_cols = min(int(args.get("maxCols") or 100), int(limits["maxCols"]))
    include_formulas = bool(args.get("includeFormulas", False))
    rows: list[list[Any]] = []
    total_rows = 0
    truncated_rows = False
    for row_index, row in enumerate(sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, max_rows + 1), max_col=min(sheet.max_column, max_cols), values_only=False), start=1):
        total_rows = row_index
        values = [_cell_value(cell, include_formulas) for cell in row]
        if any(value not in (None, "") for value in values):
            rows.append(values)
    if sheet.max_row > max_rows:
        truncated_rows = True
    return {
        "name": sheet.title,
        "rows": rows,
        "totalRows": sheet.max_row,
        "totalColumns": sheet.max_column,
        "truncatedRows": truncated_rows,
        "truncatedColumns": sheet.max_column > max_cols,
    }


def _sheet_markdown(data: Mapping[str, Any], header_row: bool) -> str:
    rows = data.get("rows") or []
    if not rows:
        return f"## {data['name']}\n\n(empty)"
    if header_row and len(rows) > 1:
        headers = [scalar_text(value) for value in rows[0]]
        body = [[scalar_text(value) for value in row] for row in rows[1:]]
        return f"## {data['name']}\n\n" + md_table(headers, body)
    body = [[scalar_text(value) for value in row] for row in rows]
    width = max(len(row) for row in body)
    return f"## {data['name']}\n\n" + md_table([f"Column {index + 1}" for index in range(width)], body)


def _sheet_text(data: Mapping[str, Any]) -> str:
    lines = [f"# {data['name']}"]
    for row in data.get("rows") or []:
        lines.append("\t".join(scalar_text(value) for value in row))
    return "\n".join(lines)


def read_xlsx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    data_only = not bool(args.get("includeFormulas", False))
    workbook = load_workbook(path, read_only=True, data_only=data_only)
    try:
        sheets = _select_sheets(workbook, args.get("sheet"))
        summaries = [_sheet_summary(sheet) for sheet in workbook.worksheets]
        mode = str(args.get("mode") or "content")
        if mode == "outline":
            return {"type": "xlsx", "mode": "outline", "sheets": summaries, "totalSheets": len(summaries)}
        if mode == "search":
            needle = str(args.get("search") or "")
            if not needle:
                raise EngineError("INVALID_ARGUMENT", "search requires a non-empty search string")
            match_case = bool(args.get("matchCase", False))
            needle_cmp = needle if match_case else needle.casefold()
            matches: list[dict[str, Any]] = []
            for sheet in sheets:
                for row in sheet.iter_rows(values_only=False):
                    for cell in row:
                        value = cell.value
                        if value is None:
                            continue
                        text = scalar_text(value)
                        if needle_cmp in (text if match_case else text.casefold()):
                            matches.append({"sheet": sheet.title, "cell": cell.coordinate, "value": jsonable(value)})
                            if len(matches) >= 100:
                                break
                    if len(matches) >= 100:
                        break
                if len(matches) >= 100:
                    break
            return {"type": "xlsx", "mode": "search", "matches": matches, "sheets": summaries}
        records = [_read_sheet(sheet, args, limits) for sheet in sheets]
        output_format = str(args.get("format") or "markdown")
        if output_format == "json":
            payload = {"sheets": records, "workbook": summaries}
            import json
            full_text = json.dumps(payload, ensure_ascii=False, default=jsonable)
        elif output_format == "text":
            full_text = "\n\n".join(_sheet_text(record) for record in records)
        else:
            full_text = "\n\n".join(_sheet_markdown(record, bool(args.get("headerRow", True))) for record in records)
        text, truncated_flag = truncate(full_text, int(args.get("maxChars") or limits["maxTextChars"]))
        return {
            "type": "xlsx",
            "mode": "content",
            "format": output_format,
            "text": text,
            "totalChars": len(full_text),
            "truncated": truncated_flag,
            "sheets": records,
            "workbook": summaries,
        }
    finally:
        workbook.close()


# ---------------------------------------------------------------------------
# Creating
# ---------------------------------------------------------------------------


def _style_header_row(sheet: Any, row: int, column_count: int, fill: str = _HEADER_FILL) -> None:
    for column in range(1, column_count + 1):
        cell = sheet.cell(row=row, column=column)
        cell.font = Font(bold=True, color=_HEADER_FONT, name="Microsoft YaHei", size=10.5)
        cell.fill = PatternFill("solid", fgColor=fill.lstrip("#"))
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="BFBFBF"))


def _normalise_columns(columns: Any) -> list[dict[str, Any]]:
    if columns is None:
        return []
    result: list[dict[str, Any]] = []
    for index, column in enumerate(require_list(columns, "columns", max_items=200)):
        if isinstance(column, Mapping):
            header = scalar_text(column.get("header") or column.get("title") or f"Column {index + 1}")
            result.append(
                {
                    "header": header,
                    "key": scalar_text(column.get("key") or header),
                    "width": column.get("width"),
                    "numberFormat": column.get("numberFormat"),
                    "align": column.get("align"),
                }
            )
        else:
            header = scalar_text(column)
            result.append({"header": header, "key": header, "width": None, "numberFormat": None, "align": None})
    return result


def _normalise_rows(rows: Any, columns: list[dict[str, Any]]) -> list[list[Any]]:
    result: list[list[Any]] = []
    for row in require_list(rows, "rows", max_items=100000):
        if isinstance(row, Mapping):
            result.append([row.get(column["key"], row.get(column["header"])) for column in columns] if columns else [row])
        elif isinstance(row, (list, tuple)):
            result.append(list(row))
        else:
            result.append([row])
    return result


def _write_sheet(sheet: Any, spec: Mapping[str, Any]) -> None:
    columns = _normalise_columns(spec.get("columns"))
    rows = _normalise_rows(spec.get("rows"), columns)
    header = spec.get("header")
    start_row = 1
    if columns:
        for index, column in enumerate(columns, start=1):
            cell = sheet.cell(row=1, column=index, value=column["header"])
            if column.get("width"):
                sheet.column_dimensions[get_column_letter(index)].width = float(column["width"])
        _style_header_row(sheet, 1, len(columns), str(spec.get("headerFill") or _HEADER_FILL))
        start_row = 2
    elif header and rows:
        for index, value in enumerate(rows[0], start=1):
            sheet.cell(row=1, column=index, value=value)
        _style_header_row(sheet, 1, len(rows[0]), str(spec.get("headerFill") or _HEADER_FILL))
        rows = rows[1:]
        start_row = 2
    for row_index, row in enumerate(rows, start=start_row):
        for column_index, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_index, column=column_index, value=value)
            if columns and column_index <= len(columns):
                column = columns[column_index - 1]
                if column.get("numberFormat"):
                    cell.number_format = str(column["numberFormat"])
                if column.get("align"):
                    cell.alignment = Alignment(horizontal=str(column["align"]), vertical="center")
            cell.font = Font(name="Microsoft YaHei", size=10.5)
    if spec.get("freeze"):
        sheet.freeze_panes = str(spec["freeze"])
    if spec.get("autoFilter") and (columns or rows):
        last_column = len(columns) if columns else max((len(row) for row in rows), default=1)
        last_row = start_row + len(rows) - 1
        sheet.auto_filter.ref = f"A1:{get_column_letter(max(1, last_column))}{max(1, last_row)}"
    if spec.get("tabColor"):
        sheet.sheet_properties.tabColor = str(spec["tabColor"]).lstrip("#")
    if spec.get("zoom"):
        sheet.sheet_view.zoomScale = int(spec["zoom"])


def _chart_data_sheet(workbook: Any) -> Any:
    """Create a fresh hidden data sheet without ever reusing a user sheet."""
    base = "_officeforge_chart_data"
    title = base
    index = 1
    while title in workbook.sheetnames:
        title = f"{base}_{index}"
        index += 1
    sheet = workbook.create_sheet(title)
    sheet.sheet_state = "hidden"
    return sheet


def _chart_references(workbook: Any, chart_spec: Mapping[str, Any], sheet: Any = None, data_sheet: Any = None) -> tuple[Any, Any]:
    data_range = chart_spec.get("dataRange")
    if data_range:
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(str(data_range))
        data = Reference(sheet or workbook.active, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
        categories = None
        if chart_spec.get("categoriesRange"):
            c_min_col, c_min_row, c_max_col, c_max_row = range_boundaries(str(chart_spec["categoriesRange"]))
            categories = Reference(sheet or workbook.active, min_col=c_min_col, min_row=c_min_row, max_col=c_max_col, max_row=c_max_row)
        return data, categories
    categories = [scalar_text(value) for value in require_list(chart_spec.get("categories"), "chart.categories", max_items=500)]
    series = require_list(chart_spec.get("series"), "chart.series", max_items=20)
    data_sheet = data_sheet or _chart_data_sheet(workbook)
    start_row = data_sheet.max_row + 1
    if start_row == 1:
        start_row = 1
    data_sheet.cell(row=start_row, column=1, value="Category")
    series_values: list[list[Any]] = []
    for series_index, item in enumerate(series, start=2):
        data = require_dict(item, "chart.series item")
        data_sheet.cell(row=start_row, column=series_index, value=scalar_text(data.get("name") or f"Series {series_index - 1}"))
        series_values.append(list(require_list(data.get("values"), "chart.series.values", max_items=500)))
    if not categories:
        categories = [str(index + 1) for index in range(max((len(values) for values in series_values), default=0))]
    for row_offset, category in enumerate(categories, start=1):
        data_sheet.cell(row=start_row + row_offset, column=1, value=category)
        for series_index, values in enumerate(series_values, start=2):
            data_sheet.cell(row=start_row + row_offset, column=series_index, value=values[row_offset - 1] if row_offset - 1 < len(values) else None)
    data = Reference(data_sheet, min_col=1, min_row=start_row, max_row=start_row + len(categories), max_col=len(series) + 1)
    categories_ref = Reference(data_sheet, min_col=1, min_row=start_row + 1, max_row=start_row + len(categories))
    return data, categories_ref


def _combo_chart(workbook: Any, chart_spec: Mapping[str, Any]) -> Any:
    series_specs = require_list(chart_spec.get("series"), "chart.series", max_items=40)
    if not series_specs:
        raise EngineError("INVALID_ARGUMENT", "combo chart requires at least one series")
    categories = [scalar_text(value) for value in require_list(chart_spec.get("categories"), "chart.categories", max_items=500)]
    data_sheet = _chart_data_sheet(workbook)
    start_row = 1
    data_sheet.cell(row=start_row, column=1, value="Category")
    values_columns: list[list[Any]] = []
    for index, item in enumerate(series_specs, start=2):
        spec = require_dict(item, "chart.series item")
        data_sheet.cell(row=start_row, column=index, value=scalar_text(spec.get("name") or f"Series {index - 1}"))
        values_columns.append([value for value in require_list(spec.get("values"), "chart.series.values", max_items=500)])
    if not categories:
        categories = [str(index + 1) for index in range(max((len(values) for values in values_columns), default=0))]
    for offset, category in enumerate(categories, start=1):
        data_sheet.cell(row=start_row + offset, column=1, value=category)
        for index, values in enumerate(values_columns, start=2):
            data_sheet.cell(row=start_row + offset, column=index, value=values[offset - 1] if offset - 1 < len(values) else None)
    categories_ref = Reference(data_sheet, min_col=1, min_row=start_row + 1, max_row=start_row + len(categories))
    type_map = {"bar": BarChart, "column": BarChart, "line": LineChart, "area": AreaChart, "pie": PieChart, "scatter": ScatterChart}
    groups: dict[tuple[str, str], Any] = {}
    for index, item in enumerate(series_specs, start=2):
        spec = require_dict(item, "chart.series item")
        series_type = str(spec.get("type") or "column").lower()
        axis = str(spec.get("axis") or "primary").lower()
        key = (series_type, axis)
        if key not in groups:
            chart_class = type_map.get(series_type, BarChart)
            chart = chart_class()
            chart.title = scalar_text(chart_spec.get("title") or "")
            chart.set_categories(categories_ref)
            groups[key] = chart
        data_ref = Reference(data_sheet, min_col=index, min_row=start_row, max_row=start_row + len(categories), max_col=index)
        groups[key].add_data(data_ref, titles_from_data=True)
    primary = next((chart for (type_name, axis), chart in groups.items() if axis != "secondary"), None)
    if primary is None:
        primary = next(iter(groups.values()))
    for (type_name, axis), chart in groups.items():
        if chart is primary:
            continue
        if axis == "secondary":
            chart.y_axis.axId = 200
            chart.y_axis.title = scalar_text(chart_spec.get("secondaryAxisTitle") or "")
        primary += chart
    if any(axis == "secondary" for (_, axis) in groups):
        primary.y_axis.crosses = "max"
    return primary


def _apply_charts(sheet: Any, workbook: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "charts", "Excel charts")
    count = 0
    data_sheet = None
    for item in require_list(values, "charts", max_items=40):
        chart_spec = require_dict(item, "chart")
        chart_type = str(chart_spec.get("type") or "bar").lower()
        chart_class = {
            "bar": BarChart,
            "column": BarChart,
            "line": LineChart,
            "pie": PieChart,
            "area": AreaChart,
            "scatter": ScatterChart,
        }.get(chart_type, BarChart)
        if chart_type == "combo":
            chart = _combo_chart(workbook, chart_spec)
        else:
            chart = chart_class()
            chart.title = scalar_text(chart_spec.get("title") or "")
            if not chart_spec.get("dataRange") and data_sheet is None:
                data_sheet = _chart_data_sheet(workbook)
            data, categories = _chart_references(workbook, chart_spec, sheet, data_sheet)
            chart.add_data(data, titles_from_data=bool(chart_spec.get("titlesFromData", True)))
            if categories is not None:
                chart.set_categories(categories)
        if chart_spec.get("height"):
            chart.height = float(chart_spec["height"])
        if chart_spec.get("width"):
            chart.width = float(chart_spec["width"])
        sheet.add_chart(chart, str(chart_spec.get("anchor") or "H2"))
        count += 1
    return count


def _solid_fill(color: Any) -> Any:
    value = str(color or "FFF2CC").lstrip("#")
    return PatternFill(start_color=value, end_color=value, fill_type="solid")


def _apply_conditional_formatting(sheet: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "conditionalFormatting", "Excel conditional formatting")
    count = 0
    for item in require_list(values, "conditionalFormatting", max_items=200):
        rule_spec = require_dict(item, "conditionalFormatting item")
        cell_range = str(rule_spec.get("range") or "")
        if not cell_range:
            raise EngineError("INVALID_ARGUMENT", "conditionalFormatting.range is required")
        rule_type = str(rule_spec.get("type") or "cellIs").lower()
        if rule_type == "colorscale":
            rule = ColorScaleRule(
                start_type=str(rule_spec.get("startType") or "min"),
                start_color=str(rule_spec.get("startColor") or "F8696B").lstrip("#"),
                mid_type=str(rule_spec.get("midType") or "percentile"),
                mid_value=float(rule_spec.get("midValue") or 50),
                mid_color=str(rule_spec.get("midColor") or "FFEB84").lstrip("#"),
                end_type=str(rule_spec.get("endType") or "max"),
                end_color=str(rule_spec.get("endColor") or "63BE7B").lstrip("#"),
            )
        elif rule_type == "databar":
            rule = DataBarRule(
                start_type=str(rule_spec.get("startType") or "num"),
                start_value=rule_spec.get("startValue", 0),
                end_type=str(rule_spec.get("endType") or "num"),
                end_value=rule_spec.get("endValue", 100),
                color=str(rule_spec.get("color") or "638EC6").lstrip("#"),
            )
        elif rule_type == "formula":
            formula = rule_spec.get("formula")
            formulas = formula if isinstance(formula, list) else [formula or "FALSE"]
            rule = FormulaRule(formula=[str(value) for value in formulas], fill=_solid_fill(rule_spec.get("fill")))
        else:
            formula = rule_spec.get("formula")
            formulas = formula if isinstance(formula, list) else [formula or ""]
            rule = CellIsRule(
                operator=str(rule_spec.get("operator") or "greaterThan"),
                formula=[str(value) for value in formulas],
                fill=_solid_fill(rule_spec.get("fill")),
                font=Font(color=str(rule_spec.get("fontColor") or "9C0006").lstrip("#"), bold=bool(rule_spec.get("bold", False))),
            )
        sheet.conditional_formatting.add(cell_range, rule)
        count += 1
    return count


def _apply_data_validations(sheet: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "dataValidation", "Excel data validation")
    count = 0
    for item in require_list(values, "dataValidation", max_items=200):
        spec = require_dict(item, "dataValidation item")
        cell_range = str(spec.get("range") or "")
        if not cell_range:
            raise EngineError("INVALID_ARGUMENT", "dataValidation.range is required")
        validation = DataValidation(
            type=str(spec.get("type") or "whole"),
            operator=str(spec.get("operator") or "between") if spec.get("operator") else None,
            formula1=scalar_text(spec.get("formula1") or ""),
            formula2=scalar_text(spec.get("formula2") or "") if spec.get("formula2") is not None else None,
            allow_blank=bool(spec.get("allowBlank", True)),
        )
        if spec.get("promptTitle"):
            validation.promptTitle = scalar_text(spec["promptTitle"])
        if spec.get("prompt"):
            validation.prompt = scalar_text(spec["prompt"])
        if spec.get("errorTitle"):
            validation.errorTitle = scalar_text(spec["errorTitle"])
        if spec.get("error"):
            validation.error = scalar_text(spec["error"])
        sheet.add_data_validation(validation)
        validation.add(cell_range)
        count += 1
    return count


def _apply_comments(sheet: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "comments", "Excel comments")
    count = 0
    for item in require_list(values, "comments", max_items=500):
        spec = require_dict(item, "comment")
        cell = sheet[str(spec.get("cell") or "A1")]
        cell.comment = Comment(scalar_text(spec.get("text") or ""), scalar_text(spec.get("author") or "Office Forge"))
        count += 1
    return count


def _apply_merged_cells(sheet: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "mergedCells", "Excel merged cells")
    count = 0
    for item in require_list(values, "mergedCells", max_items=500):
        cell_range = str(item.get("range") if isinstance(item, Mapping) else item)
        if cell_range:
            sheet.merge_cells(cell_range)
            count += 1
    return count


def _apply_images(sheet: Any, values: Any, workspace: str | None, features: Any) -> int:
    require_feature(features, "excel", "images", "Excel images")
    if not workspace:
        raise EngineError("INVALID_ARGUMENT", "workspace is required for image insertion")
    count = 0
    for item in require_list(values, "images", max_items=40):
        spec = require_dict(item, "image")
        image_path = os.path.abspath(str(spec.get("path") or ""))
        assert_inside(image_path, workspace, "image.path")
        if not os.path.isfile(image_path):
            raise EngineError("INPUT_NOT_FOUND", f"image not found: {spec.get('path')}")
        image = ExcelImage(image_path)
        if spec.get("width"):
            image.width = float(spec["width"])
        if spec.get("height"):
            image.height = float(spec["height"])
        sheet.add_image(image, str(spec.get("anchor") or "A1"))
        count += 1
    return count


def _range_bounds(sheet: Any, range_ref: Any) -> tuple[int, int, int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(str(range_ref or "A1"))
    except Exception as error:
        raise EngineError("INVALID_RANGE", f"invalid cell range: {range_ref!r}") from error
    if max_col is None:
        max_col = min_col
    if max_row is None:
        max_row = sheet.max_row
    return int(min_col), int(min_row), int(max_col), int(max_row)


def _clean_text(value: str, collapse_spaces: bool) -> str:
    text = value.strip()
    if collapse_spaces:
        text = re.sub(r"\s+", " ", text)
    return text


def _parse_number_text(value: str) -> Any:
    text = value.strip().replace(",", "").replace("，", "")
    text = re.sub(r"^[¥￥$€£]\s*", "", text)
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return value
    if re.fullmatch(r"[+-]?\d+", text):
        try:
            return int(text)
        except ValueError:
            return value
    if re.fullmatch(r"[+-]?\d*\.\d+", text):
        try:
            return float(text)
        except ValueError:
            return value
    return value


def _parse_date_text(value: str) -> Any:
    text = value.strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%m/%d/%Y"):
        try:
            return _datetime.datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return value


def _sort_key(value: Any) -> tuple[int, Any]:
    if value is None or value == "":
        return (2, "")
    if isinstance(value, bool):
        return (0, float(value))
    if isinstance(value, (int, float)):
        return (0, float(value))
    if isinstance(value, _datetime.datetime):
        return (0, value.timestamp())
    if isinstance(value, _datetime.date):
        return (0, _datetime.datetime(value.year, value.month, value.day).timestamp())
    return (1, str(value).casefold())


def _sort_range(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "excel", "dataTools", "Excel data tools")
    min_col, min_row, max_col, max_row = _range_bounds(sheet, operation.get("range"))
    header = bool(operation.get("header", True))
    first_data_row = min_row + 1 if header else min_row
    keys = require_list(operation.get("keys"), "sortRange.keys", max_items=8)
    if not keys:
        raise EngineError("INVALID_ARGUMENT", "sortRange.keys must contain at least one key")
    parsed_keys: list[tuple[int, bool]] = []
    for key in keys:
        spec = key if isinstance(key, Mapping) else {"column": key}
        raw_column = spec.get("column")
        if isinstance(raw_column, str):
            column_index = column_index_from_string(raw_column) if raw_column.isalpha() else int(raw_column)
        else:
            column_index = int(raw_column)
        if column_index < min_col or column_index > max_col:
            raise EngineError("INVALID_ARGUMENT", f"sort column {raw_column!r} is outside the selected range")
        parsed_keys.append((column_index - min_col, bool(spec.get("descending", False))))

    rows = []
    for row_index in range(first_data_row, max_row + 1):
        cells = [sheet.cell(row=row_index, column=column) for column in range(min_col, max_col + 1)]
        rows.append((list(cells), [copy.copy(cell._style) for cell in cells], [cell.value for cell in cells]))

    def key_for(row: tuple[Any, Any, Any]) -> tuple[Any, ...]:
        values = row[2]
        return tuple(_sort_key(values[index]) for index, _ in parsed_keys)

    # Sorting is stable; reverse only the requested direction for each key is
    # approximated by sorting the last key first and applying reverse from the
    # final key backward.
    ordered = list(rows)
    for column_offset, descending in reversed(parsed_keys):
        ordered.sort(key=lambda row, offset=column_offset: _sort_key(row[2][offset]), reverse=descending)

    changed = 0
    for target_offset, (source_cells, source_styles, source_values) in enumerate(ordered):
        target_row = first_data_row + target_offset
        for column_offset in range(len(source_values)):
            cell = sheet.cell(row=target_row, column=min_col + column_offset)
            cell.value = source_values[column_offset]
            cell._style = copy.copy(source_styles[column_offset])
            changed += 1
    return {"range": f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}", "rows": len(rows), "cells": changed}


def _remove_duplicate_rows(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "excel", "dataTools", "Excel data tools")
    min_col, min_row, max_col, max_row = _range_bounds(sheet, operation.get("range"))
    header = bool(operation.get("header", True))
    first_data_row = min_row + 1 if header else min_row
    raw_columns = require_list(operation.get("columns"), "removeDuplicates.columns", max_items=20) if operation.get("columns") else list(range(min_col, max_col + 1))
    key_columns: list[int] = []
    for column in raw_columns:
        index = column_index_from_string(column) if isinstance(column, str) and column.isalpha() else int(column)
        if index < min_col or index > max_col:
            raise EngineError("INVALID_ARGUMENT", f"duplicate key column {column!r} is outside the selected range")
        key_columns.append(index)
    seen: set[tuple[Any, ...]] = set()
    to_delete: list[int] = []
    for row_index in range(first_data_row, max_row + 1):
        key = tuple(sheet.cell(row=row_index, column=column).value for column in key_columns)
        if key in seen:
            to_delete.append(row_index)
        else:
            seen.add(key)
    for row_index in reversed(to_delete):
        sheet.delete_rows(row_index, 1)
    return {"removed": len(to_delete), "kept": max(0, max_row - first_data_row + 1 - len(to_delete))}


def _fill_blanks(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "excel", "dataTools", "Excel data tools")
    min_col, min_row, max_col, max_row = _range_bounds(sheet, operation.get("range"))
    method = str(operation.get("method") or "value").lower()
    fill_value = operation.get("value")
    changed = 0
    if method == "up":
        for column in range(min_col, max_col + 1):
            last = None
            for row in range(max_row, min_row - 1, -1):
                cell = sheet.cell(row=row, column=column)
                if cell.__class__.__name__ == "MergedCell":
                    continue
                if cell.value not in (None, ""):
                    last = cell.value
                elif last is not None:
                    cell.value = last
                    changed += 1
    else:
        for column in range(min_col, max_col + 1):
            last = None
            for row in range(min_row, max_row + 1):
                cell = sheet.cell(row=row, column=column)
                if cell.__class__.__name__ == "MergedCell":
                    continue
                if cell.value not in (None, ""):
                    last = cell.value
                    continue
                if method == "down":
                    if last is not None:
                        cell.value = last
                        changed += 1
                else:
                    cell.value = fill_value
                    changed += 1
    return {"filled": changed, "method": method}


def _clean_range(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "excel", "dataTools", "Excel data tools")
    min_col, min_row, max_col, max_row = _range_bounds(sheet, operation.get("range"))
    trim = bool(operation.get("trim", True))
    collapse = bool(operation.get("collapseSpaces", True))
    number_text = bool(operation.get("numberText", True))
    date_text = bool(operation.get("dateText", False))
    case_mode = str(operation.get("case") or "").lower()
    changed = 0
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            cell = sheet.cell(row=row, column=column)
            if cell.__class__.__name__ == "MergedCell" or not isinstance(cell.value, str) or cell.value.startswith("="):
                continue
            original = cell.value
            value = _clean_text(original, collapse) if trim else original
            if case_mode == "lower":
                value = value.lower()
            elif case_mode == "upper":
                value = value.upper()
            if number_text:
                value = _parse_number_text(value) if isinstance(value, str) else value
            if date_text and isinstance(value, str):
                value = _parse_date_text(value)
            if value != original:
                cell.value = value
                changed += 1
    return {"cleaned": changed, "range": f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"}


def _add_formula_column(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    require_feature(features, "excel", "formulas", "Excel formulas")
    raw_column = operation.get("column")
    if raw_column in (None, ""):
        column_index = sheet.max_column + 1
    elif isinstance(raw_column, str) and raw_column.isalpha():
        column_index = column_index_from_string(raw_column)
    else:
        column_index = int(raw_column)
    header_row = int(operation.get("headerRow") or 1)
    if operation.get("header") not in (None, ""):
        sheet.cell(row=header_row, column=column_index, value=scalar_text(operation["header"]))
    start_row = int(operation.get("startRow") or header_row + 1)
    end_row = int(operation.get("endRow") or sheet.max_row)
    formula = scalar_text(operation.get("formula") or "")
    if not formula:
        raise EngineError("INVALID_ARGUMENT", "addFormulaColumn.formula must not be empty")
    if start_row <= end_row and "{r}" not in formula and "{row}" not in formula:
        raise EngineError("INVALID_ARGUMENT", "addFormulaColumn.formula must contain {r} or {row} when filling multiple rows")
    number_format = scalar_text(operation.get("numberFormat") or "")
    for row in range(start_row, end_row + 1):
        cell = sheet.cell(row=row, column=column_index, value=formula.replace("{r}", str(row)).replace("{row}", str(row)))
        if number_format:
            cell.number_format = number_format
    if operation.get("width"):
        sheet.column_dimensions[get_column_letter(column_index)].width = float(operation["width"])
    return {"column": get_column_letter(column_index), "rows": max(0, end_row - start_row + 1)}


def _apply_data_operation(sheet: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    name = str(operation.get("op") or "")
    if name == "addFormulaColumn":
        return _add_formula_column(sheet, operation, features)
    if name == "sortRange":
        return _sort_range(sheet, operation, features)
    if name == "removeDuplicates":
        return _remove_duplicate_rows(sheet, operation, features)
    if name == "fillBlanks":
        return _fill_blanks(sheet, operation, features)
    if name == "cleanRange":
        return _clean_range(sheet, operation, features)
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported Excel data operation: {name}")


def _table_name(value: Any, index: int) -> str:
    raw = scalar_text(value).strip()
    cleaned = re.sub(r"[^0-9A-Za-z_]", "_", raw)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned:
        cleaned = f"Table{index}"
    if not re.match(r"[A-Za-z_]", cleaned[0]):
        cleaned = f"Table_{cleaned}"
    return cleaned[:200]


def _apply_tables(sheet: Any, values: Any, features: Any) -> int:
    require_feature(features, "excel", "tables", "Excel table objects")
    count = 0
    existing = set(sheet.tables.keys()) if hasattr(sheet, "tables") else set()
    for index, item in enumerate(require_list(values, "tables", max_items=20), start=1):
        spec = require_dict(item, "table")
        ref = scalar_text(spec.get("range") or spec.get("ref") or "")
        if not ref:
            raise EngineError("INVALID_ARGUMENT", "table.range must not be empty")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except (TypeError, ValueError) as error:
            raise EngineError("INVALID_RANGE", f"invalid table range: {ref}", {"error": str(error)}) from error
        if min_row is None or max_row is None or min_row == max_row:
            raise EngineError("INVALID_RANGE", "table.range must span a header row and at least one data row", {"range": ref})
        name = _table_name(spec.get("name"), index)
        if name in existing:
            raise EngineError("TABLE_EXISTS", f"table already exists on this sheet: {name}", {"table": name})
        table = Table(displayName=name, ref=f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}")
        style = spec.get("style")
        if style:
            table.tableStyleInfo = TableStyleInfo(
                name=scalar_text(style),
                showFirstColumn=bool(spec.get("showFirstColumn", False)),
                showLastColumn=bool(spec.get("showLastColumn", False)),
                showRowStripes=bool(spec.get("showRowStripes", True)),
                showColumnStripes=bool(spec.get("showColumnStripes", False)),
            )
        sheet.add_table(table)
        existing.add(name)
        count += 1
    return count


_PROTECTION_OPTIONS = (
    "formatCells",
    "formatColumns",
    "formatRows",
    "insertColumns",
    "insertRows",
    "insertHyperlinks",
    "deleteColumns",
    "deleteRows",
    "sort",
    "autoFilter",
    "pivotTables",
)


def _protection_password(value: Any) -> str | None:
    text = scalar_text(value).strip()
    return text or None


def _apply_sheet_protection(sheet: Any, spec: Mapping[str, Any], features: Any) -> bool:
    require_feature(features, "excel", "protection", "Excel sheet protection")
    protection = SheetProtection()
    protection.sheet = True
    password = _protection_password(spec.get("password"))
    if password:
        protection.password = password
    for key in _PROTECTION_OPTIONS:
        if key in spec:
            setattr(protection, key, bool(spec[key]))
    sheet.protection = protection
    return True


def _apply_workbook_protection(workbook: Any, spec: Mapping[str, Any], features: Any) -> bool:
    require_feature(features, "excel", "protection", "Excel workbook protection")
    protection = WorkbookProtection()
    protection.lockStructure = bool(spec.get("lockStructure", True))
    protection.lockWindows = bool(spec.get("lockWindows", False))
    password = _protection_password(spec.get("password"))
    if password:
        protection.workbookPassword = password
    workbook.security = protection
    return True


def _apply_sheet_features(sheet: Any, workbook: Any, spec: Mapping[str, Any], workspace: str | None, features: Any) -> dict[str, int]:
    return {
        "charts": _apply_charts(sheet, workbook, spec["charts"], features) if spec.get("charts") else 0,
        "conditionalFormatting": _apply_conditional_formatting(sheet, spec["conditionalFormatting"], features) if spec.get("conditionalFormatting") else 0,
        "dataValidation": _apply_data_validations(sheet, spec["dataValidation"], features) if spec.get("dataValidation") else 0,
        "comments": _apply_comments(sheet, spec["comments"], features) if spec.get("comments") else 0,
        "mergedCells": _apply_merged_cells(sheet, spec["mergedCells"], features) if spec.get("mergedCells") else 0,
        "images": _apply_images(sheet, spec["images"], workspace, features) if spec.get("images") else 0,
        "tables": _apply_tables(sheet, spec["tables"], features) if spec.get("tables") else 0,
        "protection": 1 if spec.get("protection") and _apply_sheet_protection(sheet, require_dict(spec["protection"], "protection"), features) else 0,
    }


def _unique_sheet_title(workbook: Any, wanted: str) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", wanted).strip() or "Sheet"
    base = base[:31]
    title = base
    suffix = 2
    existing = {sheet.title for sheet in workbook.worksheets}
    while title in existing:
        ending = f" ({suffix})"
        title = base[: 31 - len(ending)] + ending
        suffix += 1
    return title


def _import_source_path(value: Any, workspace: str | None) -> str:
    if value in (None, ""):
        raise EngineError("INVALID_ARGUMENT", "import.source is required")
    raw = os.fspath(value)
    absolute = raw if os.path.isabs(raw) else os.path.join(workspace or os.getcwd(), raw)
    absolute = os.path.abspath(absolute)
    if workspace:
        assert_inside(absolute, workspace)
    return absolute


def _import_sheets(import_spec: Mapping[str, Any], workspace: str | None, limits: Mapping[str, Any]) -> list[dict[str, Any]]:
    source = _import_source_path(import_spec.get("source") or import_spec.get("path"), workspace)
    fmt = str(import_spec.get("format") or os.path.splitext(source)[1].lstrip(".")).lower()
    if fmt not in {"csv", "tsv", "json"}:
        raise EngineError("UNSUPPORTED_FORMAT", f"unsupported import format: {fmt}", {"allowed": ["csv", "tsv", "json"]})
    assert_input_file(source, (".csv", ".tsv", ".json"), int(limits["maxInputBytes"]))
    if fmt in {"csv", "tsv"}:
        delimiter = "\t" if fmt == "tsv" else ","
        with open(source, "r", encoding="utf-8-sig", newline="") as handle:
            rows = [list(row) for row in csv.reader(handle, delimiter=delimiter)]
        if not rows:
            rows = []
        raw_header = bool(import_spec.get("header", True))
        columns = [scalar_text(value) for value in rows[0]] if raw_header and rows else []
        data_rows = rows[1:] if raw_header and rows else rows
        coerced: list[list[Any]] = []
        for row in data_rows:
            values: list[Any] = []
            for value in row:
                text = scalar_text(value)
                parsed = _parse_number_text(text) if text.strip() else None
                values.append(parsed if parsed is not None else (None if not text.strip() else value))
            coerced.append(values)
        return [{"name": scalar_text(import_spec.get("sheet") or "Imported"), "columns": columns, "rows": coerced}]
    with open(source, "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    if isinstance(payload, list):
        if payload and all(isinstance(item, Mapping) for item in payload):
            columns: list[str] = []
            for item in payload:
                for key in item:
                    if str(key) not in columns:
                        columns.append(str(key))
            return [{"name": scalar_text(import_spec.get("sheet") or "Imported"), "columns": columns, "rows": payload}]
        if payload and all(isinstance(item, (list, tuple)) for item in payload):
            raw_header = bool(import_spec.get("header", True))
            columns = [scalar_text(value) for value in payload[0]] if raw_header else []
            data_rows = [list(row) for row in (payload[1:] if raw_header else payload)]
            return [{"name": scalar_text(import_spec.get("sheet") or "Imported"), "columns": columns, "rows": data_rows}]
        return [{"name": scalar_text(import_spec.get("sheet") or "Imported"), "columns": [], "rows": [[value] for value in payload]}]
    if isinstance(payload, Mapping) and isinstance(payload.get("sheets"), list):
        return [require_dict(item, "import sheet") for item in payload["sheets"]]
    if isinstance(payload, Mapping):
        columns = [str(key) for key in payload.keys()]
        return [{"name": scalar_text(import_spec.get("sheet") or "Imported"), "columns": columns, "rows": [payload]}]
    raise EngineError("INVALID_ARGUMENT", "JSON import must be an object or array")


def create_xlsx(path: str, spec: Mapping[str, Any], limits: Mapping[str, Any], overwrite: bool, workspace: str | None = None, features: Any = None) -> dict[str, Any]:
    workbook = Workbook()
    workbook.remove(workbook.active)
    if spec.get("import"):
        require_feature(features, "excel", "import", "Excel CSV/JSON import")
        sheets = _import_sheets(require_dict(spec["import"], "import"), workspace, limits)
        if spec.get("sheets"):
            sheets = list(sheets) + list(require_list(spec["sheets"], "sheets", max_items=100))
    else:
        sheets = require_list(spec.get("sheets"), "sheets", max_items=100)
    if not sheets:
        sheets = [{"name": "Sheet1", "rows": []}]
    created: list[str] = []
    feature_counts: dict[str, int] = {}
    for sheet_spec in sheets:
        data = require_dict(sheet_spec, "sheet")
        title = _unique_sheet_title(workbook, scalar_text(data.get("name") or "Sheet1"))
        sheet = workbook.create_sheet(title)
        _write_sheet(sheet, data)
        counts = _apply_sheet_features(sheet, workbook, data, workspace, features)
        for key, value in counts.items():
            feature_counts[key] = feature_counts.get(key, 0) + value
        created.append(title)
    if spec.get("properties"):
        properties = require_dict(spec["properties"], "properties")
        for key in ("title", "subject", "creator", "keywords", "description", "category"):
            if key in properties:
                setattr(workbook.properties, key, scalar_text(properties[key]))
    # Normalise the template metadata after caller values so a missing author
    # never leaves openpyxl's own placeholder behind.
    applied_properties = apply_core_properties(
        workbook.properties,
        spec.get("properties"),
        author_attr="creator",
        last_modified_attr="lastModifiedBy",
        comments_attr="description",
    )
    if spec.get("workbookProtection"):
        _apply_workbook_protection(workbook, require_dict(spec["workbookProtection"], "workbookProtection"), features)
        feature_counts["workbookProtection"] = feature_counts.get("workbookProtection", 0) + 1

    def writer(temp: str) -> None:
        workbook.save(temp)

    size, _ = atomic_save(path, writer, overwrite=overwrite, max_bytes=int(limits["maxOutputBytes"]))
    return {
        "type": "xlsx",
        "path": path,
        "bytes": size,
        "sheets": created,
        "featureCounts": feature_counts,
        "properties": applied_properties,
    }


# ---------------------------------------------------------------------------
# Editing
# ---------------------------------------------------------------------------


def _resolve_sheet(workbook: Any, name: Any, create: bool = False) -> Any:
    if name in (None, ""):
        return workbook.active
    for sheet in workbook.worksheets:
        if sheet.title == str(name):
            return sheet
    if create:
        return workbook.create_sheet(_unique_sheet_title(workbook, str(name)))
    raise EngineError("SHEET_NOT_FOUND", f"sheet {name!r} was not found", {"sheets": [sheet.title for sheet in workbook.worksheets]})


def _replace_in_sheet(sheet: Any, find: str, replacement: str, regex: bool, match_case: bool) -> int:
    count = 0
    flags = 0 if match_case else re.IGNORECASE
    pattern = re.compile(find if regex else re.escape(find), flags)
    for row in sheet.iter_rows():
        for cell in row:
            value = cell.value
            if not isinstance(value, str):
                continue
            if regex:
                new_value, number = pattern.subn(replacement, value)
            elif match_case:
                number = value.count(find)
                new_value = value.replace(find, replacement)
            else:
                new_value, number = pattern.subn(lambda _m: replacement, value)
            if number:
                cell.value = new_value
                count += number
    return count



def _pivot_summary(workbook: Any, operation: Mapping[str, Any], features: Any) -> dict[str, Any]:
    """Create a deterministic pivot-style summary sheet without Excel COM.

    This is a file-level grouped aggregation, not a native Excel PivotTable
    object. It intentionally writes values so the result is portable.
    """
    require_feature(features, "excel", "pivots", "Excel pivot summaries")
    source = _resolve_sheet(workbook, operation.get("sheet"))
    header_row = int(operation.get("headerRow") or 1)
    if header_row < 1 or header_row > source.max_row:
        raise EngineError("INVALID_RANGE", "pivotSummary.headerRow is outside the sheet")
    headers = [scalar_text(source.cell(row=header_row, column=column).value) for column in range(1, source.max_column + 1)]
    header_index = {name: index for index, name in enumerate(headers) if name}
    row_keys = operation.get("rows") or operation.get("rowFields") or []
    if isinstance(row_keys, str):
        row_keys = [row_keys]
    row_keys = [scalar_text(key) for key in row_keys]
    if not row_keys:
        raise EngineError("INVALID_ARGUMENT", "pivotSummary.rows must contain at least one field")
    for key in row_keys:
        if key not in header_index:
            raise EngineError("COLUMN_NOT_FOUND", f"pivot field was not found: {key}")
    value_specs = operation.get("values") or operation.get("valueFields") or []
    if isinstance(value_specs, Mapping):
        value_specs = [value_specs]
    if not value_specs:
        raise EngineError("INVALID_ARGUMENT", "pivotSummary.values must contain at least one aggregation")
    parsed_values: list[tuple[str, str, str]] = []
    for item in value_specs:
        spec = require_dict(item, "pivot value")
        column = scalar_text(spec.get("column") or spec.get("field") or "")
        if column not in header_index:
            raise EngineError("COLUMN_NOT_FOUND", f"pivot value field was not found: {column}")
        aggregate = scalar_text(spec.get("aggregate") or spec.get("function") or "sum").lower()
        if aggregate not in {"sum", "count", "avg", "average", "min", "max"}:
            raise EngineError("UNSUPPORTED_OPERATION", f"unsupported pivot aggregate: {aggregate}")
        label = scalar_text(spec.get("label") or f"{aggregate}_{column}")
        parsed_values.append((column, aggregate, label))

    groups: dict[tuple[Any, ...], list[list[Any]]] = {}
    for row in range(header_row + 1, source.max_row + 1):
        key = tuple(source.cell(row=row, column=header_index[name] + 1).value for name in row_keys)
        if all(value is None for value in key):
            continue
        groups.setdefault(key, []).append([source.cell(row=row, column=header_index[column] + 1).value for column, _, _ in parsed_values])

    target_name = _unique_sheet_title(workbook, scalar_text(operation.get("target") or operation.get("name") or "Pivot"))
    target = workbook.create_sheet(target_name)
    target.append([*row_keys, *[label for _, _, label in parsed_values]])
    summary_rows = 0
    for key, value_rows in sorted(groups.items(), key=lambda pair: tuple(scalar_text(value) for value in pair[0])):
        output: list[Any] = list(key)
        for value_index, (_, aggregate, _) in enumerate(parsed_values):
            numbers = []
            for row_values in value_rows:
                raw = row_values[value_index]
                if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                    numbers.append(float(raw))
                else:
                    parsed = _parse_number_text(raw)
                    if parsed is not None:
                        numbers.append(float(parsed))
            if aggregate == "count":
                output.append(len(numbers))
            elif aggregate in {"avg", "average"}:
                output.append(sum(numbers) / len(numbers) if numbers else None)
            elif aggregate == "min":
                output.append(min(numbers) if numbers else None)
            elif aggregate == "max":
                output.append(max(numbers) if numbers else None)
            else:
                output.append(sum(numbers) if numbers else 0)
        target.append(output)
        summary_rows += 1
    for column in range(1, len(row_keys) + len(parsed_values) + 1):
        target.column_dimensions[get_column_letter(column)].width = 16
    return {"target": target_name, "rows": summary_rows, "groups": len(groups)}

def edit_xlsx(path: str, args: Mapping[str, Any], limits: Mapping[str, Any], workspace: str | None = None, features: Any = None) -> dict[str, Any]:
    workbook = load_workbook(path, data_only=False, keep_vba=str(path).lower().endswith(".xlsm"))
    operations = require_list(args.get("operations"), "operations", max_items=300)
    if not operations:
        raise EngineError("INVALID_ARGUMENT", "operations must contain at least one edit")
    applied: list[dict[str, Any]] = []
    for item in operations:
        operation = require_dict(item, "operation")
        name = str(operation.get("op") or "")
        if name == "setCells":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            changed = 0
            for entry in require_list(operation.get("cells"), "cells", max_items=50000):
                cell_spec = require_dict(entry, "cell")
                coordinate = cell_spec.get("cell")
                if coordinate:
                    cell = sheet[str(coordinate)]
                else:
                    row = int(cell_spec.get("row") or 1)
                    column = cell_spec.get("column")
                    if isinstance(column, str) and column.isalpha():
                        from openpyxl.utils import column_index_from_string
                        column = column_index_from_string(column)
                    cell = sheet.cell(row=row, column=int(column or 1))
                if "arrayFormula" in cell_spec:
                    formula = str(cell_spec["arrayFormula"])
                    if not formula.startswith("="):
                        formula = "=" + formula
                    ref = str(cell_spec.get("ref") or cell_spec.get("range") or cell.coordinate)
                    cell.value = ArrayFormula(ref=ref, text=formula)
                elif "formula" in cell_spec:
                    cell.value = str(cell_spec["formula"] if str(cell_spec["formula"]).startswith("=") else "=" + str(cell_spec["formula"]))
                elif "value" in cell_spec:
                    cell.value = cell_spec["value"]
                if cell_spec.get("numberFormat"):
                    cell.number_format = str(cell_spec["numberFormat"])
                changed += 1
            applied.append({"op": "setCells", "sheet": sheet.title, "cells": changed})
        elif name == "appendRows":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            rows = _normalise_rows(operation.get("rows"), [])
            for row in rows:
                sheet.append(row)
            applied.append({"op": "appendRows", "sheet": sheet.title, "rows": len(rows)})
        elif name == "replace":
            find = scalar_text(operation.get("find") or "")
            if not find:
                raise EngineError("INVALID_ARGUMENT", "replace.find must not be empty")
            selected = _select_sheets(workbook, operation.get("sheet")) if operation.get("sheet") else list(workbook.worksheets)
            total = sum(_replace_in_sheet(sheet, find, scalar_text(operation.get("replace") or ""), bool(operation.get("regex", False)), bool(operation.get("matchCase", False))) for sheet in selected)
            applied.append({"op": "replace", "matches": total})
        elif name == "createSheet":
            title = _unique_sheet_title(workbook, scalar_text(operation.get("name") or "Sheet"))
            workbook.create_sheet(title)
            applied.append({"op": "createSheet", "sheet": title})
        elif name == "deleteSheet":
            sheet = _resolve_sheet(workbook, operation.get("name"))
            if len(workbook.worksheets) <= 1:
                raise EngineError("INVALID_OPERATION", "cannot delete the only worksheet")
            workbook.remove(sheet)
            applied.append({"op": "deleteSheet", "sheet": sheet.title})
        elif name == "renameSheet":
            sheet = _resolve_sheet(workbook, operation.get("name"))
            new_title = _unique_sheet_title(workbook, scalar_text(operation.get("newName") or sheet.title))
            old_title = sheet.title
            sheet.title = new_title
            applied.append({"op": "renameSheet", "from": old_title, "to": new_title})
        elif name == "setColumnWidths":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            widths = require_dict(operation.get("widths"), "widths")
            for column, width in widths.items():
                sheet.column_dimensions[str(column)].width = float(width)
            applied.append({"op": "setColumnWidths", "sheet": sheet.title, "columns": len(widths)})
        elif name == "setFreeze":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            sheet.freeze_panes = operation.get("cell") or None
            applied.append({"op": "setFreeze", "sheet": sheet.title, "cell": sheet.freeze_panes})
        elif name == "setAutoFilter":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            sheet.auto_filter.ref = operation.get("range") or None
            applied.append({"op": "setAutoFilter", "sheet": sheet.title, "range": sheet.auto_filter.ref})
        elif name == "addChart":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            chart_spec = operation.get("chart") if operation.get("chart") is not None else operation
            count = _apply_charts(sheet, workbook, [chart_spec], features)
            applied.append({"op": "addChart", "sheet": sheet.title, "count": count})
        elif name == "addConditionalFormatting":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            rules = operation.get("rules") if operation.get("rules") is not None else [operation]
            count = _apply_conditional_formatting(sheet, rules, features)
            applied.append({"op": "addConditionalFormatting", "sheet": sheet.title, "count": count})
        elif name == "addDataValidation":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            validations = operation.get("validations") if operation.get("validations") is not None else [operation]
            count = _apply_data_validations(sheet, validations, features)
            applied.append({"op": "addDataValidation", "sheet": sheet.title, "count": count})
        elif name == "addComment":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            comments = operation.get("comments") if operation.get("comments") is not None else [operation]
            count = _apply_comments(sheet, comments, features)
            applied.append({"op": "addComment", "sheet": sheet.title, "count": count})
        elif name == "mergeCells":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            ranges = operation.get("ranges") if operation.get("ranges") is not None else [{"range": operation.get("range")}]
            count = _apply_merged_cells(sheet, ranges, features)
            applied.append({"op": "mergeCells", "sheet": sheet.title, "count": count})
        elif name == "addImage":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            images = operation.get("images") if operation.get("images") is not None else [operation]
            count = _apply_images(sheet, images, workspace, features)
            applied.append({"op": "addImage", "sheet": sheet.title, "count": count})
        elif name in {"addFormulaColumn", "sortRange", "removeDuplicates", "fillBlanks", "cleanRange"}:
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            result = _apply_data_operation(sheet, operation, features)
            applied.append({"op": name, "sheet": sheet.title, **result})
        elif name == "pivotSummary":
            result = _pivot_summary(workbook, operation, features)
            applied.append({"op": "pivotSummary", **result})
        elif name == "addTable":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            tables = operation.get("tables") if operation.get("tables") is not None else [operation]
            count = _apply_tables(sheet, tables, features)
            applied.append({"op": "addTable", "sheet": sheet.title, "count": count})
        elif name == "protectSheet":
            sheet = _resolve_sheet(workbook, operation.get("sheet"))
            _apply_sheet_protection(sheet, operation, features)
            applied.append({"op": "protectSheet", "sheet": sheet.title})
        elif name == "protectWorkbook":
            _apply_workbook_protection(workbook, operation, features)
            applied.append({"op": "protectWorkbook", "lockStructure": bool(operation.get("lockStructure", True)), "lockWindows": bool(operation.get("lockWindows", False))})
        else:
            raise EngineError("UNSUPPORTED_OPERATION", f"unsupported xlsx edit operation: {name or '(empty)'}", {"allowed": ["setCells", "appendRows", "replace", "createSheet", "deleteSheet", "renameSheet", "setColumnWidths", "setFreeze", "setAutoFilter", "addChart", "addConditionalFormatting", "addDataValidation", "addComment", "mergeCells", "addImage", "addFormulaColumn", "sortRange", "removeDuplicates", "fillBlanks", "cleanRange", "pivotSummary", "addTable", "protectSheet", "protectWorkbook"]})

    size, backup = atomic_save(path, lambda temp: workbook.save(temp), overwrite=True, max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(args.get("backup", True)))
    workbook.close()
    return {"type": "xlsx", "path": path, "bytes": size, "backupPath": backup, "operations": applied}


# ---------------------------------------------------------------------------
# Template expansion
# ---------------------------------------------------------------------------


def _copy_row_style(source: Any, target: Any) -> None:
    for source_cell, target_cell in zip(source, target):
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)
        target_cell.number_format = source_cell.number_format
        target_cell.font = copy.copy(source_cell.font)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.alignment = copy.copy(source_cell.alignment)
        target_cell.protection = copy.copy(source_cell.protection)


def _expand_sheet_loops(sheet: Any, data: Any, missing: str) -> int:
    count = 0
    row_index = 1
    while row_index <= sheet.max_row:
        values = [sheet.cell(row=row_index, column=column).value for column in range(1, sheet.max_column + 1)]
        joined = " ".join(scalar_text(value) for value in values if value is not None)
        match = re.search(r"\{\{\s*#(?:each\s+)?([\w.]+)\s*\}\}", joined)
        if not match:
            row_index += 1
            continue
        key = match.group(1)
        end_row = row_index + 1
        end_pattern = re.compile(r"\{\{\s*/(?:each(?:\s+([\w.]+))?|([\w.]+))\s*\}\}")
        while end_row <= sheet.max_row:
            end_values = [sheet.cell(row=end_row, column=column).value for column in range(1, sheet.max_column + 1)]
            end_joined = " ".join(scalar_text(value) for value in end_values if value is not None)
            end_match = end_pattern.search(end_joined)
            end_key = (end_match.group(1) or end_match.group(2)) if end_match else None
            if end_match and (end_key is None or end_key == key):
                break
            end_row += 1
        if end_row > sheet.max_row:
            row_index += 1
            continue
        items = get_path(data, key, [])
        if not isinstance(items, list):
            items = [] if items in (None, "") else [items]
        template_count = end_row - row_index - 1
        if template_count <= 0:
            row_index = end_row + 1
            continue
        template_rows = [[sheet.cell(row=row_index + offset, column=column) for column in range(1, sheet.max_column + 1)] for offset in range(1, template_count + 1)]
        template_snapshot = [
            [
                {
                    "value": template_cell.value,
                    "style": copy.copy(template_cell._style) if template_cell.has_style else None,
                    "numberFormat": template_cell.number_format,
                    "font": copy.copy(template_cell.font),
                    "fill": copy.copy(template_cell.fill),
                    "border": copy.copy(template_cell.border),
                    "alignment": copy.copy(template_cell.alignment),
                    "protection": copy.copy(template_cell.protection),
                }
                for template_cell in template_row
            ]
            for template_row in template_rows
        ]
        needed_rows = len(items) * template_count
        if needed_rows > template_count:
            sheet.insert_rows(end_row + 1, amount=needed_rows - template_count)
        elif needed_rows < template_count:
            sheet.delete_rows(row_index + 1 + needed_rows, amount=template_count - needed_rows)
        template_start = row_index + 1
        cursor = template_start
        for item_index, item in enumerate(items):
            context = dict(data) if isinstance(data, Mapping) else {}
            if isinstance(item, Mapping):
                context.update(item)
            context["this"] = item
            context["index"] = item_index + 1
            for template_row in template_snapshot:
                for column_index, snapshot in enumerate(template_row, start=1):
                    target = sheet.cell(row=cursor, column=column_index)
                    value = snapshot["value"]
                    if isinstance(value, str):
                        value = render_placeholders(value, context, missing)
                    target.value = value
                    if snapshot["style"] is not None:
                        target._style = copy.copy(snapshot["style"])
                    target.number_format = snapshot["numberFormat"]
                    target.font = copy.copy(snapshot["font"])
                    target.fill = copy.copy(snapshot["fill"])
                    target.border = copy.copy(snapshot["border"])
                    target.alignment = copy.copy(snapshot["alignment"])
                    target.protection = copy.copy(snapshot["protection"])
                cursor += 1
        # Blank template markers after the expanded block.
        for marker_row in (row_index, template_start + needed_rows):
            for column in range(1, sheet.max_column + 1):
                sheet.cell(row=marker_row, column=column).value = None
        count += len(items)
        row_index = template_start + needed_rows + 1
    return count


def _replace_workbook_text(workbook: Any, data: Any, missing: str) -> int:
    count = 0
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "{{" in cell.value:
                    cell.value = render_placeholders(cell.value, data, missing)
                    count += 1
    return count


def template_xlsx(template: str, output: str, data: Mapping[str, Any], options: Mapping[str, Any], limits: Mapping[str, Any]) -> dict[str, Any]:
    workbook = load_workbook(template, data_only=False)
    loops = 0
    if options.get("expandLoops", True):
        for sheet in workbook.worksheets:
            loops += _expand_sheet_loops(sheet, data, str(options.get("missing") or "keep"))
    replacements = _replace_workbook_text(workbook, data, str(options.get("missing") or "keep"))
    size, backup = atomic_save(output, lambda temp: workbook.save(temp), overwrite=bool(options.get("overwrite", False)), max_bytes=int(limits["maxOutputBytes"]), make_backup=bool(options.get("backup", False)))
    workbook.close()
    return {"type": "xlsx", "path": output, "bytes": size, "rowsExpanded": loops, "cellsReplaced": replacements, "backupPath": backup}


# ---------------------------------------------------------------------------
# Workbook audit and insight (v0.4)
# ---------------------------------------------------------------------------

_ERROR_VALUES = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def _formula_pattern(formula: str) -> str:
    text = re.sub(r'"[^"]*"', '"{}"', formula)
    text = re.sub(r"(?<![A-Za-z0-9_])([A-Z]{1,3})\$?(\d+)", r"\1{r}", text)
    return text


def _formula_refs(formula: str) -> set[str]:
    return {match.replace("$", "").upper() for match in re.findall(r"\b([A-Za-z]{1,3}\$?\d+)\b", formula)}


def _cell_summary(cell: Any) -> dict[str, Any]:
    return {"sheet": cell.parent.title, "cell": cell.coordinate}


def _health_xlsx(path: str, args: Mapping[str, Any]) -> dict[str, Any]:
    workbook = load_workbook(path, data_only=False, keep_vba=str(path).lower().endswith(".xlsm"))
    try:
        cached = load_workbook(path, data_only=True)
    except Exception:
        cached = None
    issues: list[dict[str, Any]] = []
    max_issues = max(1, min(int(args.get("maxIssues") or 200), 1000))
    try:
        for sheet in workbook.worksheets:
            sheet_cached = cached[sheet.title] if cached is not None and sheet.title in cached.sheetnames else None
            max_row = min(sheet.max_row, 5000)
            max_col = min(sheet.max_column, 100)
            formulas: dict[int, list[Any]] = {}
            for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                for cell in row:
                    value = cell.value
                    if not isinstance(value, str) or not value.startswith("="):
                        continue
                    formulas.setdefault(cell.column, []).append(cell)
                    computed = sheet_cached[cell.coordinate].value if sheet_cached is not None else None
                    if (isinstance(computed, str) and computed in _ERROR_VALUES) or any(error in value for error in _ERROR_VALUES):
                        issues.append({
                            "code": "formula_error",
                            "severity": "high",
                            "sheet": sheet.title,
                            "cell": cell.coordinate,
                            "formula": value,
                            "message": f"{cell.coordinate} contains an error value or an error literal",
                        })
            for column, cells in formulas.items():
                if len(cells) < 3:
                    continue
                buckets: dict[str, list[Any]] = {}
                for cell in cells:
                    buckets.setdefault(_formula_pattern(str(cell.value)), []).append(cell)
                if len(buckets) > 1:
                    dominant = max(buckets, key=lambda pattern: len(buckets[pattern]))
                    if len(buckets[dominant]) >= 2:
                        for pattern, group in buckets.items():
                            if pattern == dominant:
                                continue
                            for cell in group:
                                issues.append({
                                    "code": "inconsistent_formula",
                                    "severity": "medium",
                                    "sheet": sheet.title,
                                    "cell": cell.coordinate,
                                    "formula": cell.value,
                                    "expectedPattern": dominant,
                                    "message": "formula pattern differs from the dominant pattern in this column",
                                })
                min_row = min(cell.row for cell in cells)
                max_row_in_column = max(cell.row for cell in cells)
                for row_number in range(min_row, max_row_in_column + 1):
                    cell = sheet.cell(row=row_number, column=column)
                    if cell.value is None or (isinstance(cell.value, str) and cell.value.startswith("=")):
                        continue
                    if isinstance(cell.value, (int, float)):
                        issues.append({
                            "code": "hardcoded_value",
                            "severity": "low",
                            "sheet": sheet.title,
                            "cell": cell.coordinate,
                            "value": jsonable(cell.value),
                            "message": "numeric constant sits inside a formula column",
                        })
            # Blank rows inside a populated block are usually broken table links.
            for row_number in range(2, max_row):
                values = [sheet.cell(row=row_number, column=column).value for column in range(1, min(max_col, 20) + 1)]
                if any(value not in (None, "") for value in values):
                    continue
                above = any(sheet.cell(row=row_number - 1, column=column).value not in (None, "") for column in range(1, min(max_col, 20) + 1))
                below = any(sheet.cell(row=row_number + 1, column=column).value not in (None, "") for column in range(1, min(max_col, 20) + 1))
                if above and below:
                    issues.append({
                        "code": "blank_row_break",
                        "severity": "low",
                        "sheet": sheet.title,
                        "cell": f"{row_number}:{row_number}",
                        "message": "blank row interrupts an otherwise populated table",
                    })
            # Lightweight circular-reference check across formula cells.
            coordinates = {cell.coordinate for group in formulas.values() for cell in group}
            edges = {
                cell.coordinate: [reference for reference in _formula_refs(str(cell.value)) if reference in coordinates]
                for group in formulas.values()
                for cell in group
            }
            reported: set[str] = set()
            for start in edges:
                if start in reported:
                    continue
                stack = list(edges.get(start, []))
                seen: set[str] = set()
                while stack:
                    node = stack.pop()
                    if node == start:
                        issues.append({
                            "code": "circular_reference",
                            "severity": "high",
                            "sheet": sheet.title,
                            "cell": start,
                            "message": "formula dependency cycle reaches back to this cell",
                        })
                        reported.add(start)
                        break
                    if node in seen:
                        continue
                    seen.add(node)
                    stack.extend(edges.get(node, []))
        summary: dict[str, int] = {}
        for issue in issues:
            summary[issue["code"]] = summary.get(issue["code"], 0) + 1
        high = sum(1 for issue in issues if issue["severity"] == "high")
        medium = sum(1 for issue in issues if issue["severity"] == "medium")
        low = sum(1 for issue in issues if issue["severity"] == "low")
        score = max(0, 100 - high * 10 - medium * 5 - low * 2)
        return {
            "type": "xlsx",
            "mode": "health",
            "score": score,
            "issueCount": len(issues),
            "summary": summary,
            "issues": issues[:max_issues],
            "truncated": len(issues) > max_issues,
        }
    finally:
        workbook.close()
        if cached is not None:
            cached.close()


def profile_xlsx(path: str) -> dict[str, Any]:
    workbook = load_workbook(path, data_only=False, keep_vba=str(path).lower().endswith(".xlsm"))
    try:
        sheets: list[dict[str, Any]] = []
        for sheet in workbook.worksheets:
            formula_count = 0
            error_count = 0
            comment_count = 0
            max_row = min(sheet.max_row, 5000)
            max_col = min(sheet.max_column, 100)
            for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_count += 1
                    if isinstance(cell.value, str) and cell.value in _ERROR_VALUES:
                        error_count += 1
                    if cell.comment is not None:
                        comment_count += 1
            sheets.append({
                "name": sheet.title,
                "maxRow": sheet.max_row,
                "maxColumn": sheet.max_column,
                "state": sheet.sheet_state,
                "freezePanes": getattr(sheet, "freeze_panes", None),
                "autoFilter": bool(getattr(getattr(sheet, "auto_filter", None), "ref", None)),
                "formulaCount": formula_count,
                "errorCount": error_count,
                "mergedRanges": len(sheet.merged_cells.ranges),
                "tables": len(sheet.tables),
                "charts": len(sheet._charts),
                "images": len(sheet._images),
                "conditionalFormatting": len(sheet.conditional_formatting),
                "dataValidation": len(sheet.data_validations.dataValidation),
                "comments": comment_count,
            })
        return {
            "type": "xlsx",
            "mode": "profile",
            "sheetCount": len(sheets),
            "definedNames": list(workbook.defined_names.keys()),
            "sheets": sheets,
        }
    finally:
        workbook.close()


def insight_xlsx(path: str, args: Mapping[str, Any]) -> dict[str, Any]:
    workbook = load_workbook(path, data_only=True)
    try:
        sheets: list[dict[str, Any]] = []
        max_rows = min(int(args.get("maxRows") or 2000), 5000)
        for sheet in workbook.worksheets:
            rows = list(sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, max_rows), max_col=min(sheet.max_column, 60), values_only=True))
            if not rows:
                sheets.append({"name": sheet.title, "columns": [], "rowCount": 0})
                continue
            headers = [scalar_text(value) if value not in (None, "") else f"Column {index + 1}" for index, value in enumerate(rows[0])]
            columns: list[dict[str, Any]] = []
            for index, header in enumerate(headers):
                values = [row[index] for row in rows[1:] if index < len(row) and row[index] not in (None, "")]
                numeric = [value for value in values if isinstance(value, (int, float)) and not isinstance(value, bool)]
                text_values = [scalar_text(value) for value in values if not isinstance(value, (int, float))]
                item: dict[str, Any] = {
                    "header": header,
                    "nonEmpty": len(values),
                    "missing": max(0, len(rows) - 1 - len(values)),
                }
                if numeric:
                    item.update({
                        "type": "number",
                        "count": len(numeric),
                        "sum": round(sum(numeric), 6),
                        "average": round(sum(numeric) / len(numeric), 6),
                        "min": min(numeric),
                        "max": max(numeric),
                    })
                elif text_values:
                    counts: dict[str, int] = {}
                    for value in text_values:
                        counts[value] = counts.get(value, 0) + 1
                    item.update({
                        "type": "text",
                        "count": len(text_values),
                        "topValues": [{"value": value, "count": count} for value, count in sorted(counts.items(), key=lambda pair: (-pair[1], pair[0]))[:5]],
                    })
                else:
                    item["type"] = "empty"
                columns.append(item)
            sheets.append({"name": sheet.title, "rowCount": max(0, len(rows) - 1), "columns": columns})
        return {"type": "xlsx", "mode": "insight", "sheets": sheets}
    finally:
        workbook.close()


def diff_xlsx(path: str, compare_path: str, args: Mapping[str, Any]) -> dict[str, Any]:
    left = load_workbook(path, data_only=False)
    right = load_workbook(compare_path, data_only=False)
    try:
        changes: list[dict[str, Any]] = []
        max_changes = max(1, min(int(args.get("maxChanges") or 200), 1000))
        left_names = set(left.sheetnames)
        right_names = set(right.sheetnames)
        summary = {
            "addedSheets": sorted(right_names - left_names),
            "removedSheets": sorted(left_names - right_names),
            "changedCells": 0,
        }
        for name in sorted(left_names & right_names):
            left_sheet = left[name]
            right_sheet = right[name]
            max_row = max(left_sheet.max_row, right_sheet.max_row)
            max_col = max(left_sheet.max_column, right_sheet.max_column)
            for row in range(1, min(max_row, 5000) + 1):
                for column in range(1, min(max_col, 100) + 1):
                    left_cell = left_sheet.cell(row=row, column=column)
                    right_cell = right_sheet.cell(row=row, column=column)
                    if left_cell.value != right_cell.value:
                        summary["changedCells"] += 1
                        if len(changes) < max_changes:
                            changes.append({
                                "sheet": name,
                                "cell": left_cell.coordinate,
                                "left": jsonable(left_cell.value),
                                "right": jsonable(right_cell.value),
                            })
        return {"type": "xlsx", "mode": "diff", "summary": summary, "changes": changes, "truncated": summary["changedCells"] > len(changes)}
    finally:
        left.close()
        right.close()


def audit_xlsx(path: str, mode: str, args: Mapping[str, Any], compare_path: str | None = None) -> dict[str, Any]:
    if mode == "health":
        return _health_xlsx(path, args)
    if mode == "profile":
        return profile_xlsx(path)
    if mode == "insight":
        return insight_xlsx(path, args)
    if mode == "diff":
        if not compare_path:
            raise EngineError("INVALID_ARGUMENT", "diff mode requires comparePath")
        return diff_xlsx(path, compare_path, args)
    raise EngineError("UNSUPPORTED_OPERATION", f"unsupported xlsx audit mode: {mode}", {"allowed": ["health", "profile", "insight", "diff"]})
