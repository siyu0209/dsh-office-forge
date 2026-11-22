"""End-to-end checks for the Office Forge Python engine.

Run with the bundled DSH runtime:
    python.exe -m tests.test_engine
or:
    python.exe tests/test_engine.py
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE = ROOT / "lib" / "office_engine.py"
WORK = Path(tempfile.mkdtemp(prefix="officeforge-test-"))


def call(payload: dict, extra_env: dict | None = None) -> dict:
    env = os.environ.copy()
    env["OFFICEFORGE_WORKSPACE"] = str(WORK)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    if extra_env:
        env.update(extra_env)
    process = subprocess.run(
        [sys.executable, str(ENGINE)],
        input=json.dumps(payload, ensure_ascii=False),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        cwd=str(WORK),
        env=env,
        timeout=120,
    )
    if not process.stdout.strip():
        raise AssertionError(f"engine returned no stdout (exit {process.returncode})\nstderr:\n{process.stderr[-4000:]}")
    data = json.loads(process.stdout)
    return data


def ok(payload: dict, extra_env: dict | None = None) -> dict:
    data = call(payload, extra_env)
    if not data.get("ok"):
        raise AssertionError(f"engine error for {payload.get('action')}: {json.dumps(data, ensure_ascii=False, indent=2)}")
    return data["result"]


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)
    print(f"  PASS  {message}")


def test_probe() -> None:
    print("[probe]")
    result = ok({"action": "probe"})
    check(result["python"].startswith("3.12"), f"python={result['python']}")
    for module in ("python-docx", "openpyxl", "python-pptx"):
        check(module in result["modules"], f"module {module} is importable")


def test_docx() -> None:
    print("[docx]")
    path = WORK / "report.docx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "季度报告",
            "subtitle": "Office Forge 测试文档",
            "author": "Office Forge",
            "sections": [
                {
                    "heading": "一、概览",
                    "level": 1,
                    "paragraphs": ["这是第一段正文。", "This is the second paragraph."],
                },
                {
                    "heading": "二、数据",
                    "level": 2,
                    "bullets": ["要点 A", "要点 B"],
                    "table": {
                        "headers": ["项目", "数值"],
                        "rows": [["收入", 120], ["成本", 80]],
                    },
                },
                {"heading": "三、结论", "level": 1, "paragraphs": ["结论段落。"]},
            ],
            "pageNumbers": True,
        },
    })
    check(created["bytes"] > 1000, f"docx created, {created['bytes']} bytes")

    read = ok({"action": "read", "path": str(path), "type": "docx", "args": {"mode": "content", "format": "markdown"}})
    check("季度报告" in read["text"] and "收入" in read["text"], f"docx markdown contains title and table text; got {read.get('text', '')[:400]!r}")
    check("一、概览" in read["text"], "docx markdown contains headings")

    outline = ok({"action": "read", "path": str(path), "type": "docx", "args": {"mode": "outline", "format": "json"}})
    check(bool(outline.get("headings")), "docx outline is non-empty")

    edit = ok({
        "action": "edit",
        "path": str(path),
        "type": "docx",
        "args": {
            "operations": [
                {"op": "replace", "find": "季度报告", "replace": "年度报告"},
                {"op": "append", "sections": [{"paragraphs": ["追加的结尾段落。"]}]},
                {"op": "setCore", "values": {"title": "年度报告", "author": "Office Forge"}},
            ],
            "backup": True,
        },
    })
    check(edit["operations"][0]["matches"] == 1, "docx replace matched once")
    check(Path(edit["backupPath"]).exists(), "docx backup file exists")
    reread = ok({"action": "read", "path": str(path), "type": "docx", "args": {"mode": "content", "format": "text"}})
    check("年度报告" in reread["text"] and "追加的结尾段落" in reread["text"], "docx edit persisted")

    # Template with a one-level table loop.
    template = WORK / "template.docx"
    ok({
        "action": "create",
        "path": str(template),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "项目：{{title}}",
            "sections": [
                {
                    "heading": "客户：{{customer}}",
                    "paragraphs": ["负责人：{{owner}}"],
                    "table": {
                        "headers": ["姓名", "部门"],
                        "rows": [
                            ["{{#each people}}", ""],
                            ["{{name}}", "{{dept}}"],
                            ["{{/each}}", ""],
                        ],
                    },
                }
            ],
        },
    })
    output = WORK / "template-out.docx"
    filled = ok({
        "action": "template",
        "template": str(template),
        "output": str(output),
        "type": "docx",
        "data": {
            "title": "项目甲",
            "customer": "客户乙",
            "owner": "王五",
            "people": [
                {"name": "张三", "dept": "研发"},
                {"name": "李四", "dept": "市场"},
            ],
        },
        "options": {"overwrite": True},
    })
    check(filled["tablesExpanded"] == 2, "docx template expanded two loop items")
    filled_read = ok({"action": "read", "path": str(output), "type": "docx", "args": {"mode": "content", "format": "markdown"}})
    check("张三" in filled_read["text"] and "李四" in filled_read["text"] and "客户乙" in filled_read["text"], "docx template values present")
    check("{{" not in filled_read["text"], "docx template left no unresolved placeholders")


def test_xlsx() -> None:
    print("[xlsx]")
    path = WORK / "sales.xlsx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "xlsx",
        "overwrite": True,
        "spec": {
            "sheets": [
                {
                    "name": "销售",
                    "columns": [
                        {"header": "月份", "key": "month", "width": 14},
                        {"header": "金额", "key": "amount", "numberFormat": "#,##0"},
                    ],
                    "rows": [
                        {"month": "一月", "amount": 100},
                        {"month": "二月", "amount": 200},
                    ],
                    "freeze": "A2",
                    "autoFilter": True,
                },
                {"name": "备注", "rows": [["说明"], ["Office Forge 测试"]]},
            ]
        },
    })
    check(created["bytes"] > 1000 and created["sheets"] == ["销售", "备注"], "xlsx created with two sheets")

    read = ok({"action": "read", "path": str(path), "type": "xlsx", "args": {"mode": "content", "format": "markdown"}})
    check("一月" in read["text"] and "200" in read["text"], "xlsx markdown contains rows")
    check("销售" in read["text"] and "备注" in read["text"], "xlsx markdown contains both sheets")

    edit = ok({
        "action": "edit",
        "path": str(path),
        "type": "xlsx",
        "args": {
            "operations": [
                {"op": "setCells", "sheet": "销售", "cells": [{"cell": "A4", "value": "三月"}, {"cell": "B4", "value": 300}]},
                {"op": "appendRows", "sheet": "销售", "rows": [["四月", 400]]},
                {"op": "replace", "find": "Office Forge 测试", "replace": "Office Forge 已修改"},
            ],
            "backup": True,
        },
    })
    check(edit["operations"][0]["cells"] == 2, "xlsx setCells changed two cells")
    check(Path(edit["backupPath"]).exists(), "xlsx backup file exists")

    template = WORK / "template.xlsx"
    ok({
        "action": "create",
        "path": str(template),
        "type": "xlsx",
        "overwrite": True,
        "spec": {
            "sheets": [
                {
                    "name": "明细",
                    "rows": [
                        ["姓名", "部门"],
                        ["{{#each people}}", ""],
                        ["{{name}}", "{{dept}}"],
                        ["{{/each}}", ""],
                    ],
                }
            ]
        },
    })
    output = WORK / "template-out.xlsx"
    filled = ok({
        "action": "template",
        "template": str(template),
        "output": str(output),
        "type": "xlsx",
        "data": {"people": [{"name": "张三", "dept": "研发"}, {"name": "李四", "dept": "市场"}, {"name": "王五", "dept": "财务"}]},
        "options": {"overwrite": True},
    })
    check(filled["rowsExpanded"] == 3, "xlsx template expanded three loop rows")
    filled_read = ok({"action": "read", "path": str(output), "type": "xlsx", "args": {"mode": "content", "format": "markdown"}})
    check("张三" in filled_read["text"] and "王五" in filled_read["text"], "xlsx template values present")
    check("{{" not in filled_read["text"], "xlsx template left no unresolved placeholders")


def test_data_path() -> None:
    print("[dataPath]")
    data_file = WORK / "people.csv"
    data_file.write_text("name,dept" + chr(10) + "张三,研发" + chr(10) + "李四,市场" + chr(10), encoding="utf-8")
    template = WORK / "csv-template.docx"
    ok({
        "action": "create",
        "path": str(template),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "sections": [{
                "table": {
                    "headers": ["姓名", "部门"],
                    "rows": [
                        ["{{#each items}}", ""],
                        ["{{name}}", "{{dept}}"],
                        ["{{/each}}", ""],
                    ],
                },
            }],
        },
    })
    output = WORK / "csv-out.docx"
    filled = ok({
        "action": "template",
        "template": str(template),
        "output": str(output),
        "type": "docx",
        "dataPath": str(data_file),
        "options": {"overwrite": True},
    })
    check(filled["tablesExpanded"] == 2, "CSV dataPath expanded two rows")
    text = ok({"action": "read", "path": str(output), "type": "docx", "args": {"mode": "content", "format": "markdown"}})["text"]
    check("张三" in text and "李四" in text, "CSV dataPath values present")


def test_pptx() -> None:
    print("[pptx]")
    path = WORK / "deck.pptx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "pptx",
        "overwrite": True,
        "spec": {
            "slideSize": "16:9",
            "slides": [
                {"layout": "title", "title": "Office Forge 演示", "subtitle": "自动生成的演示文稿"},
                {
                    "layout": "content",
                    "title": "关键数据",
                    "bullets": ["第一项", {"text": "第二项", "level": 1}],
                    "table": {"rows": [["指标", "值"], ["收入", "120"], ["成本", "80"]]},
                },
                {
                    "layout": "content",
                    "title": "趋势",
                    "chart": {"type": "bar", "categories": ["一月", "二月"], "series": [{"name": "收入", "values": [100, 200]}]},
                    "notes": "演讲者备注",
                },
            ],
        },
    })
    check(created["slides"] == 3 and created["bytes"] > 1000, "pptx created with three slides")

    read = ok({"action": "read", "path": str(path), "type": "pptx", "args": {"mode": "content", "format": "markdown"}})
    check("Office Forge 演示" in read["text"] and "关键数据" in read["text"], "pptx markdown contains slide titles")
    check("演讲者备注" in read["text"], "pptx markdown contains notes")

    edit = ok({
        "action": "edit",
        "path": str(path),
        "type": "pptx",
        "args": {
            "operations": [
                {"op": "replace", "find": "关键数据", "replace": "核心数据"},
                {"op": "setNotes", "slide": 1, "text": "更新后的备注"},
                {"op": "appendSlides", "slides": [{"layout": "section", "title": "附录"}]},
            ],
            "backup": True,
        },
    })
    check(edit["operations"][0]["matches"] >= 1 and Path(edit["backupPath"]).exists(), "pptx edit and backup succeeded")


def _pixel() -> Path:
    from PIL import Image
    path = WORK / "pixel.png"
    if not path.exists():
        Image.new("RGB", (32, 32), (31, 78, 120)).save(path)
    return path


def test_docx_features() -> None:
    print("[docx features]")
    _pixel()
    path = WORK / "features.docx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "Word 增强功能",
            "toc": True,
            "tocTitle": "目录",
            "sections": [{
                "heading": "链接与图片",
                "paragraphs": [
                    {"text": "Office Forge 链接", "link": "https://example.com/officeforge"},
                    {"text": "图片：", "image": {"path": "pixel.png", "width": 0.5, "height": 0.5}},
                ],
                "images": [{"path": "pixel.png", "width": 0.4}],
            }],
        },
    })
    check(created["bytes"] > 1000, "docx feature document created")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
    check("w:hyperlink" in document_xml, "docx contains hyperlink")
    check("w:drawing" in document_xml, "docx contains image drawing")
    check("TOC" in document_xml and "w:fldChar" in document_xml, "docx contains TOC field")
    disabled = call({
        "action": "create",
        "path": str(WORK / "disabled-image.docx"),
        "type": "docx",
        "overwrite": True,
        "features": {"word": {"images": False}},
        "spec": {"sections": [{"paragraphs": [{"text": "x", "image": {"path": "pixel.png"}}]}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "docx image feature can be disabled")


def test_xlsx_features() -> None:
    print("[xlsx features]")
    _pixel()
    path = WORK / "features.xlsx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "xlsx",
        "overwrite": True,
        "spec": {
            "sheets": [{
                "name": "数据",
                "columns": [{"header": "月份"}, {"header": "收入"}],
                "rows": [["一月", 100], ["二月", 200], ["三月", 150]],
                "charts": [{
                    "type": "bar",
                    "title": "收入趋势",
                    "categories": ["一月", "二月", "三月"],
                    "series": [{"name": "收入", "values": [100, 200, 150]}],
                    "anchor": "E2",
                }],
                "conditionalFormatting": [{
                    "range": "B2:B4", "type": "cellIs", "operator": "greaterThan",
                    "formula": "150", "fill": "FFC7CE",
                }],
                "dataValidation": [{
                    "range": "B2:B4", "type": "whole", "operator": "between",
                    "formula1": "0", "formula2": "1000",
                }],
                "comments": [{"cell": "A1", "text": "表头备注", "author": "测试"}],
                "mergedCells": [{"range": "D1:E1"}],
                "images": [{"path": "pixel.png", "anchor": "G2", "width": 40, "height": 40}],
            }],
        },
    })
    counts = created["featureCounts"]
    check(counts["charts"] == 1 and counts["conditionalFormatting"] == 1, "xlsx chart and conditional formatting created")
    check(counts["dataValidation"] == 1 and counts["comments"] == 1, "xlsx validation and comment created")
    check(counts["mergedCells"] == 1 and counts["images"] == 1, "xlsx merged cell and image created")
    from openpyxl import load_workbook
    workbook = load_workbook(path)
    sheet = workbook["数据"]
    check(len(sheet._charts) == 1, "xlsx chart persisted")
    check(len(sheet.conditional_formatting._cf_rules) > 0, "xlsx conditional formatting persisted")
    check(len(sheet.data_validations.dataValidation) == 1, "xlsx data validation persisted")
    check(sheet["A1"].comment is not None and sheet["A1"].comment.text == "表头备注", "xlsx comment persisted")
    check("D1:E1" in [str(item) for item in sheet.merged_cells.ranges], "xlsx merged range persisted")
    check(len(sheet._images) == 1, "xlsx image persisted")
    disabled = call({
        "action": "create",
        "path": str(WORK / "disabled-chart.xlsx"),
        "type": "xlsx",
        "overwrite": True,
        "features": {"excel": {"charts": False}},
        "spec": {"sheets": [{"name": "S", "rows": [[1], [2]], "charts": [{"type": "bar", "categories": ["a"], "series": [{"values": [1]}]}]}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "xlsx chart feature can be disabled")


def test_pptx_features() -> None:
    print("[pptx features]")
    path = WORK / "features.pptx"
    created = ok({
        "action": "create",
        "path": str(path),
        "type": "pptx",
        "overwrite": True,
        "spec": {
            "slides": [{
                "layout": "title",
                "title": "PPT 增强功能",
                "subtitle": "切换与形状",
                "transition": {"effect": "fade", "speed": "med", "advanceAfter": 1.5},
                "shapes": [{
                    "shape": "rounded_rectangle", "fill": "1F4E78", "line": "FFFFFF",
                    "text": "Office Forge", "left": 1, "top": 3, "width": 4, "height": 1.2, "rotation": 5,
                }],
            }],
        },
    })
    check(created["bytes"] > 1000, "pptx feature deck created")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
    check("p:transition" in slide_xml, "pptx transition persisted")
    check("Office Forge" in slide_xml and "p:sp" in slide_xml, "pptx shape persisted")
    disabled = call({
        "action": "create",
        "path": str(WORK / "disabled-transition.pptx"),
        "type": "pptx",
        "overwrite": True,
        "features": {"ppt": {"transitions": False}},
        "spec": {"slides": [{"layout": "title", "title": "x", "transition": {"effect": "fade"}}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "pptx transition feature can be disabled")


def test_docx_v03() -> None:
    print("[docx v0.3]")
    path = WORK / "v03.docx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "v0.3 Word",
            "sections": [{"heading": "章节", "paragraphs": ["正文一", "正文二"]}],
        },
    })
    result = ok({
        "action": "edit",
        "path": str(path),
        "type": "docx",
        "args": {
            "operations": [
                {"op": "addCaption", "paragraph": 2, "label": "图", "text": "示例图", "position": "below"},
                {"op": "setStyle", "style": "Normal", "bold": True, "size": 12, "color": "1F4E78"},
                {"op": "updateFields"},
            ],
            "backup": True,
        },
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addCaption", "setStyle", "updateFields"], "docx v0.3 operations applied")
    check(ops[0]["label"] == "图" and ops[0]["position"] == "below", "docx caption metadata")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        settings_xml = archive.read("word/settings.xml").decode("utf-8")
    check("SEQ" in document_xml and "w:fldSimple" in document_xml, "docx caption field persisted")
    check("updateFields" in settings_xml, "docx updateFields persisted")
    disabled = call({
        "action": "edit",
        "path": str(path),
        "type": "docx",
        "features": {"word": {"captions": False}},
        "args": {"operations": [{"op": "addCaption", "text": "x"}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "docx captions can be disabled")


def test_xlsx_v03() -> None:
    print("[xlsx v0.3]")
    path = WORK / "v03.xlsx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "xlsx",
        "overwrite": True,
        "spec": {"sheets": [{
            "name": "数据",
            "columns": [{"header": "名称"}, {"header": "金额"}],
            "rows": [["b", " 20 "], ["a", "10"], ["a", "10"], ["c", ""]],
        }]},
    })
    result = ok({
        "action": "edit",
        "path": str(path),
        "type": "xlsx",
        "args": {
            "operations": [
                {"op": "addFormulaColumn", "sheet": "数据", "column": "C", "header": "双倍", "formula": "=B{r}*2", "startRow": 2, "endRow": 5},
                {"op": "sortRange", "sheet": "数据", "range": "A1:B5", "header": True, "keys": [{"column": "A"}]},
                {"op": "removeDuplicates", "sheet": "数据", "range": "A1:B5", "columns": ["A"], "header": True},
                {"op": "fillBlanks", "sheet": "数据", "range": "B1:B5", "method": "down"},
                {"op": "cleanRange", "sheet": "数据", "range": "A1:B5", "numberText": True},
            ],
            "backup": True,
        },
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addFormulaColumn", "sortRange", "removeDuplicates", "fillBlanks", "cleanRange"], "xlsx v0.3 operations applied")
    from openpyxl import load_workbook
    workbook = load_workbook(path)
    sheet = workbook["数据"]
    check(sheet["C2"].value == "=B2*2", "xlsx formula column persisted")
    check(any(isinstance(sheet.cell(row=row, column=2).value, (int, float)) for row in range(2, 7)), "xlsx cleanRange converted number text")
    check(any(sheet.cell(row=row, column=1).value == "a" for row in range(2, 7)), "xlsx sortRange retained data")
    disabled = call({
        "action": "edit",
        "path": str(path),
        "type": "xlsx",
        "features": {"excel": {"dataTools": False}},
        "args": {"operations": [{"op": "sortRange", "range": "A1:B5", "keys": [{"column": "A"}]}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "xlsx data tools can be disabled")


def test_pptx_v03() -> None:
    print("[pptx v0.3]")
    path = WORK / "v03.pptx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "pptx",
        "overwrite": True,
        "spec": {"slides": [{"layout": "title", "title": "v0.3 PPT", "subtitle": "textbox"}]},
    })
    result = ok({
        "action": "edit",
        "path": str(path),
        "type": "pptx",
        "args": {
            "operations": [
                {"op": "addTextBox", "slide": 1, "text": "Office Forge 文本框", "left": 1, "top": 2, "width": 4, "height": 1, "fill": "EAF1F8"},
                {"op": "addConnector", "slide": 1, "connector": "straight", "begin": {"x": 1, "y": 1}, "end": {"x": 3, "y": 2}, "color": "1F4E78"},
            ],
            "backup": True,
        },
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addTextBox", "addConnector"], "pptx v0.3 operations applied")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
    check("Office Forge 文本框" in slide_xml and "p:sp" in slide_xml, "pptx textbox persisted")
    check("p:cxnSp" in slide_xml, "pptx connector persisted")
    disabled = call({
        "action": "edit",
        "path": str(path),
        "type": "pptx",
        "features": {"ppt": {"shapes": False}},
        "args": {"operations": [{"op": "addTextBox", "slide": 1, "text": "x"}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "pptx shapes can be disabled")



def test_docx_v05_plus() -> None:
    print("[docx v0.5 plus]")
    path = WORK / "v05plus.docx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "v0.5 Word",
            "sections": [{"heading": "客户", "paragraphs": ["客户：{{name}}", "第二篇"]}],
        },
    })
    result = ok({
        "action": "edit",
        "path": str(path),
        "type": "docx",
        "args": {
            "operations": [
                {"op": "addComment", "find": "第二篇", "text": "请核对", "author": "Office Forge"},
                {"op": "trackRevision", "find": "第二篇", "replace": "第二篇（修订）", "author": "Office Forge"},
                {"op": "mailMerge", "data": {"name": "张三"}},
            ],
            "backup": True,
        },
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addComment", "trackRevision", "mailMerge"], "docx v0.5 plus operations applied")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        comments_xml = archive.read("word/comments.xml").decode("utf-8")
    check("w:commentRangeStart" in document_xml and "请核对" in comments_xml, "docx comment persisted")
    check("w:del" in document_xml and "w:ins" in document_xml and "第二篇（修订）" in document_xml, "docx tracked revision persisted")
    check("张三" in document_xml, "docx mail merge replaced placeholder")
    disabled = call({
        "action": "edit", "path": str(path), "type": "docx",
        "features": {"word": {"comments": False}},
        "args": {"operations": [{"op": "addComment", "text": "x"}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "docx comments can be disabled")


def test_xlsx_v05_plus() -> None:
    print("[xlsx v0.5 plus]")
    path = WORK / "v05plus.xlsx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "xlsx",
        "overwrite": True,
        "spec": {"sheets": [{
            "name": "Data",
            "columns": ["Region", "Product", "Amount"],
            "rows": [["East", "A", 10], ["East", "B", 20], ["West", "A", 5]],
        }]},
    })
    result = ok({
        "action": "edit", "path": str(path), "type": "xlsx",
        "args": {"operations": [{
            "op": "pivotSummary", "sheet": "Data", "rows": ["Region"],
            "values": [{"column": "Amount", "aggregate": "sum", "label": "Total"}],
            "target": "Pivot",
        }], "backup": True},
    })
    ops = result["operations"]
    check(ops[0]["op"] == "pivotSummary" and ops[0]["rows"] == 2 and ops[0]["groups"] == 2, "xlsx pivot summary applied")
    from openpyxl import load_workbook
    workbook = load_workbook(path)
    pivot = workbook["Pivot"]
    check(pivot["A1"].value == "Region" and pivot["B1"].value == "Total", "xlsx pivot headers persisted")
    values = {pivot.cell(row=row, column=1).value: pivot.cell(row=row, column=2).value for row in range(2, pivot.max_row + 1)}
    check(values.get("East") == 30 and values.get("West") == 5, "xlsx pivot aggregates persisted")
    disabled = call({
        "action": "edit", "path": str(path), "type": "xlsx",
        "features": {"excel": {"pivots": False}},
        "args": {"operations": [{"op": "pivotSummary", "rows": ["Region"], "values": [{"column": "Amount"}]}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "xlsx pivots can be disabled")


def test_pptx_v05_plus() -> None:
    print("[pptx v0.5 plus]")
    path = WORK / "v05plus.pptx"
    ok({
        "action": "create", "path": str(path), "type": "pptx", "overwrite": True,
        "spec": {"slides": [{"layout": "title", "title": "v0.5 PPT", "subtitle": "animation"}]},
    })
    result = ok({
        "action": "edit", "path": str(path), "type": "pptx",
        "args": {"operations": [
            {"op": "addAnimation", "slide": 1, "animation": {"effect": "fade", "duration": 1.2}},
            {"op": "addSmartArt", "slide": 1, "smartArt": {
                "items": [{"text": "计划"}, {"text": "执行"}, {"text": "检查"}],
                "layout": "horizontal", "left": 1, "top": 3, "width": 2, "height": 1,
            }},
        ], "backup": True},
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addAnimation", "addSmartArt"], "pptx v0.5 plus operations applied")
    check(ops[1]["nativeSmartArt"] is False and ops[1]["items"] == 3, "pptx smartart fallback metadata")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
    check("p:timing" in slide_xml and "p:animEffect" in slide_xml, "pptx animation persisted")
    check(all(text in slide_xml for text in ("计划", "执行", "检查")), "pptx smartart text persisted")
    disabled = call({
        "action": "edit", "path": str(path), "type": "pptx",
        "features": {"ppt": {"animations": False}},
        "args": {"operations": [{"op": "addAnimation", "slide": 1, "animation": {"effect": "fade"}}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "pptx animations can be disabled")

def test_pptx_deepening() -> None:
    print("[pptx deepening]")
    layouts = ok({"action": "audit", "type": "pptx", "mode": "layouts"})
    check(len(layouts["layouts"]) >= 30, "pptx layout catalog lists at least 30 layouts")
    check(any(item["id"] == "kpi" for item in layouts["layouts"]), "pptx layout catalog includes kpi")
    templates = ok({"action": "audit", "type": "pptx", "mode": "templates"})
    check(len(templates["templates"]) >= 10, "pptx template catalog lists at least 10 decks")
    check(any(item["id"] == "board" for item in templates["templates"]), "pptx template catalog includes board")

    from PIL import Image
    poster = WORK / "poster.png"
    Image.new("RGB", (320, 180), (30, 80, 140)).save(poster)
    clip = WORK / "clip.mp4"
    clip.write_bytes(b"\x00\x00\x00\x18ftypmp42" + b"\x00" * 2048)

    deck = WORK / "deepening.pptx"
    created = ok({
        "action": "create",
        "path": str(deck),
        "type": "pptx",
        "overwrite": True,
        "spec": {
            "template": "board",
            "master": {"background": "0B1F33", "footer": "Office Forge"},
        },
    })
    check(created["template"] == "board" and created["slides"] == 5, "pptx template deck expands five slides")

    custom = WORK / "custom-layouts.pptx"
    ok({
        "action": "create",
        "path": str(custom),
        "type": "pptx",
        "overwrite": True,
        "spec": {
            "theme": "swiss",
            "slides": [
                {"layout": "cover", "title": "PPT 深化", "subtitle": "版式库"},
                {"layout": "kpi", "title": "关键指标", "kpis": [{"label": "收入", "value": "1.2亿"}, {"label": "增长", "value": "+38%"}]},
                {"layout": "timeline", "title": "路线", "items": ["Q1", "Q2", "Q3", "Q4"]},
                {"layout": "cards", "title": "卡片", "cards": [{"title": "A", "text": "内容一"}, {"title": "B", "text": "内容二"}]},
                {"layout": "comparison", "title": "对比", "items": ["方案 A", "方案 B"]},
                {"layout": "matrix", "title": "矩阵", "items": ["高价值", "高成本", "低价值", "低成本"]},
                {"layout": "two-col", "title": "双栏", "items": ["左一", "右一"]},
                {"layout": "quote", "title": "引用", "quote": "简单是可靠的先决条件。"},
                {"layout": "bullets", "title": "要点", "items": ["第一条", "第二条"]},
                {"layout": "closing", "title": "谢谢", "subtitle": "欢迎提问"},
            ],
        },
    })
    import zipfile
    from pptx import Presentation

    presentation = Presentation(custom)
    check(len(presentation.slides) == 10, "pptx catalog layouts create ten slides")
    texts = [shape.text_frame.text for slide in presentation.slides for shape in slide.shapes if shape.has_text_frame]
    check(any("1.2亿" in text for text in texts), "pptx kpi layout renders values")
    check(any("Q1" in text for text in texts), "pptx timeline layout renders items")

    edited = ok({
        "action": "edit",
        "path": str(custom),
        "type": "pptx",
        "args": {
            "operations": [
                {"op": "addComment", "slide": 2, "text": "KPI 需要补充口径", "author": "审阅人"},
                {"op": "listComments"},
                {"op": "addMedia", "slide": 3, "media": {"kind": "video", "path": "clip.mp4", "poster": "poster.png", "autoplay": True}},
                {"op": "setTransition", "slide": 1, "transition": {"effect": "morph", "option": "byObject"}},
                {"op": "applyMaster", "master": {"footer": "Confidential"}},
            ],
            "backup": False,
        },
    })
    ops = edited["operations"]
    check([item["op"] for item in ops] == ["addComment", "listComments", "addMedia", "setTransition", "applyMaster"], "pptx deepening operations applied")
    check(ops[0]["commentIndex"] == 1 and ops[0]["author"] == "审阅人", "pptx comment records author and index")
    check(ops[1]["comments"][0]["text"] == "KPI 需要补充口径", "pptx comments can be listed")
    check(ops[2]["kind"] == "video" and ops[2]["autoplay"] is True, "pptx media records kind and autoplay")
    check(ops[4]["footer"] is True, "pptx master footer applied")

    with zipfile.ZipFile(custom) as archive:
        names = archive.namelist()
        content_types = archive.read("[Content_Types].xml").decode("utf-8")
        slide_one = archive.read("ppt/slides/slide1.xml").decode("utf-8")
        comment_xml = archive.read("ppt/comments/comment1.xml").decode("utf-8")
        author_xml = archive.read("ppt/commentAuthors.xml").decode("utf-8")
    check(any(name.startswith("ppt/media/") and name.endswith(".mp4") for name in names), "pptx media part embedded")
    check("comments+xml" in content_types and "commentAuthors+xml" in content_types, "pptx comment content types declared")
    check("p:morph" in slide_one and 'option="byObject"' in slide_one, "pptx morph transition persisted")
    check("KPI 需要补充口径" in comment_xml and "审阅人" in author_xml, "pptx comment text and author persisted")
    reopened = Presentation(custom)
    check(len(reopened.slides) == 10, "pptx with comments reopens cleanly")

    disabled = call({
        "action": "edit", "path": str(custom), "type": "pptx",
        "features": {"ppt": {"comments": False}},
        "args": {"operations": [{"op": "addComment", "slide": 1, "text": "x", "author": "y"}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "pptx comments can be disabled")
    missing = call({"action": "create", "path": str(WORK / "nope.pptx"), "type": "pptx", "overwrite": True, "spec": {"template": "nope"}})
    check(missing.get("ok") is False and missing["error"]["code"] == "TEMPLATE_NOT_FOUND", "pptx unknown template rejected")


def test_xlsx_audit() -> None:
    print("[xlsx audit]")
    path = WORK / "audit.xlsx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "xlsx",
        "overwrite": True,
        "spec": {"sheets": [{
            "name": "Data",
            "columns": ["Item", "Qty", "Amount", "Cycle", "Check"],
            "rows": [
                ["A", 10, "=B2*2", None, "=#REF!"],
                ["B", 20, 99, None, None],
                ["C", 30, "=B4*3", None, None],
                ["D", 40, "=B5*2", "=D6", None],
                [None, None, None, "=D5", None],
            ],
        }]},
    })
    health = ok({"action": "audit", "path": str(path), "type": "xlsx", "mode": "health"})
    codes = set(health["summary"])
    check("inconsistent_formula" in codes, "xlsx audit detects inconsistent formulas")
    check("hardcoded_value" in codes, "xlsx audit detects hardcoded values in formula columns")
    check("formula_error" in codes, "xlsx audit detects error literals")
    check("circular_reference" in codes, "xlsx audit detects circular references")
    profile = ok({"action": "audit", "path": str(path), "type": "xlsx", "mode": "profile"})
    check(profile["sheetCount"] == 1 and profile["sheets"][0]["formulaCount"] >= 4, "xlsx profile counts formulas")
    insight = ok({"action": "audit", "path": str(path), "type": "xlsx", "mode": "insight"})
    check(insight["sheets"][0]["columns"][1]["type"] == "number", "xlsx insight returns numeric statistics")
    other = WORK / "audit-other.xlsx"
    ok({
        "action": "create",
        "path": str(other),
        "type": "xlsx",
        "overwrite": True,
        "spec": {"sheets": [{"name": "Data", "columns": ["Item", "Qty"], "rows": [["A", 999]]}]},
    })
    diff = ok({"action": "audit", "path": str(path), "type": "xlsx", "mode": "diff", "comparePath": str(other)})
    check(diff["summary"]["changedCells"] > 0, "xlsx diff reports changed cells")
    disabled = call({"action": "audit", "path": str(path), "type": "xlsx", "mode": "health", "features": {"excel": {"audit": False}}})
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "xlsx audit can be disabled")


def test_pptx_audit() -> None:
    print("[pptx audit]")
    themes = ok({"action": "audit", "type": "pptx", "mode": "themes"})
    check(len(themes["themes"]) >= 24, "pptx audit lists at least 24 themes")
    check(any(item.get("name") == "swiss" for item in themes["themes"]), "pptx audit includes externalized swiss theme")
    path = WORK / "audit.pptx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "pptx",
        "overwrite": True,
        "spec": {"slides": [{"layout": "title", "title": "v0.4 PPT", "subtitle": "audit"}, {"layout": "blank"}]},
    })
    deck = ok({"action": "audit", "path": str(path), "type": "pptx", "mode": "deck"})
    check("empty_slide" in deck["summary"], "pptx deck check detects empty slides")
    themed = WORK / "themed.pptx"
    ok({
        "action": "create",
        "path": str(themed),
        "type": "pptx",
        "overwrite": True,
        "spec": {"theme": "forge", "slides": [{"layout": "title", "title": "Themed", "subtitle": "forge"}]},
    })
    from pptx import Presentation
    presentation = Presentation(themed)
    check(presentation.slides[0].background.fill.fore_color.rgb is not None, "pptx create applies theme background")
    edited = ok({"action": "edit", "path": str(themed), "type": "pptx", "args": {"operations": [{"op": "applyTheme", "theme": "ocean"}]}})
    check(edited["operations"][0]["theme"] == "ocean", "pptx applyTheme operation applied")
    disabled = call({"action": "edit", "path": str(themed), "type": "pptx", "features": {"ppt": {"themes": False}}, "args": {"operations": [{"op": "applyTheme", "theme": "ocean"}]}})
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "pptx themes can be disabled")


def test_docx_audit() -> None:
    print("[docx audit]")
    path = WORK / "audit.docx"
    ok({
        "action": "create",
        "path": str(path),
        "type": "docx",
        "overwrite": True,
        "spec": {
            "title": "Word 审计",
            "sections": [
                {"heading": "第一章", "level": 3, "paragraphs": ["正文一"]},
                {"heading": "重复标题", "level": 1, "paragraphs": ["正文二"]},
                {"heading": "重复标题", "level": 1, "paragraphs": ["正文三"]},
                {"heading": "孤立标题", "level": 1},
                {"heading": "下一节", "level": 1, "paragraphs": [
                    {"text": "占位符 {{name}}"},
                    {"text": "TODO: 补充"},
                    {"text": ""},
                    {"text": ""},
                    "长" * 1600,
                ], "table": {"headers": ["名称", "数量"], "rows": [["A", 1]]}},
            ],
        },
    })
    health = ok({"action": "audit", "path": str(path), "type": "docx"})
    check(health["mode"] == "health", "docx audit default mode is health")
    codes = set(health["summary"])
    for code in ("heading_jump", "duplicate_heading", "orphan_heading", "placeholder_left", "long_paragraph", "empty_paragraph_run"):
        check(code in codes, f"docx audit detects {code}")
    check(health["score"] < 100 and health["issueCount"] >= 6, "docx health computes score and issue count")
    profile = ok({"action": "audit", "path": str(path), "type": "docx", "mode": "profile"})
    check(profile["stats"]["headings"] >= 6, "docx profile counts headings")
    check(profile["stats"]["tables"] == 1, "docx profile counts tables")
    check(profile["metadata"]["author"] == "Office Forge", "docx profile returns metadata")
    disabled = call({"action": "audit", "path": str(path), "type": "docx", "features": {"word": {"audit": False}}})
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "docx audit can be disabled")


def test_pdf_v05() -> None:
    print("[pdf v0.5]")
    sample = WORK / "sample.pdf"
    shutil.copyfile(ROOT / "tests" / "sample.pdf", sample)
    result = ok({"action": "read", "path": str(sample), "type": "pdf", "args": {}})
    check(result["type"] == "pdf", "pdf read returns pdf type")
    check(result["pageCount"] >= 1, "pdf read reports at least one page")
    check("Hello PDF from Office Forge" in result["text"], "pdf fallback extracts sample text")


def test_schema_migration_v05() -> None:
    print("[schema migration v0.5]")
    result = ok({
        "action": "create",
        "path": str(WORK / "legacy-spec.docx"),
        "type": "docx",
        "overwrite": True,
        "spec": {"schemaVersion": 1, "title": "旧版 spec", "sections": []},
    })
    check(result["schemaVersion"] == 2, "schema migration upgrades spec to version 2")
    check(result["migratedFrom"] == 1, "schema migration records migratedFrom=1")


def test_pdf_tools_v05() -> None:
    print("[pdf tools v0.5]")
    sys.path.insert(0, str(ROOT / "lib" / "vendor"))
    from pypdf import PdfWriter
    a = WORK / "pdf-a.pdf"
    b = WORK / "pdf-b.pdf"
    for target in (a, b):
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        with open(target, "wb") as handle:
            writer.write(handle)
    merged = ok({"action": "pdfop", "operation": "merge", "inputs": [str(a), str(b)], "output": str(WORK / "pdf-merged.pdf")})
    check(merged["pages"] == 2 and (WORK / "pdf-merged.pdf").is_file(), "pdf merge combines two files")
    split = ok({"action": "pdfop", "operation": "split", "path": str(a), "outputDir": str(WORK / "pdf-split"), "everyPage": True})
    check(len(split["outputs"]) == 1 and (WORK / "pdf-split" / "pdf-a-001.pdf").is_file(), "pdf split writes page files")
    rotated = ok({"action": "pdfop", "operation": "rotate", "path": str(a), "output": str(WORK / "pdf-rotated.pdf"), "angle": 90})
    check(rotated["angle"] == 90 and (WORK / "pdf-rotated.pdf").is_file(), "pdf rotate applies angle")
    watermarked = ok({"action": "pdfop", "operation": "watermark", "path": str(a), "watermarkPath": str(b), "output": str(WORK / "pdf-watermarked.pdf")})
    check(watermarked["pages"] == 1 and (WORK / "pdf-watermarked.pdf").is_file(), "pdf watermark merges overlay")


def test_pdf_tools_v090() -> None:
    print("[pdf tools v0.9]")
    sys.path.insert(0, str(ROOT / "lib" / "vendor"))
    from PIL import Image
    from pypdf import PdfReader, PdfWriter

    blank = WORK / "pdf-three-pages.pdf"
    writer = PdfWriter()
    for _ in range(3):
        writer.add_blank_page(width=200, height=200)
    with open(blank, "wb") as handle:
        writer.write(handle)

    numbered = ok({
        "action": "pdfop",
        "operation": "pageNumbers",
        "path": str(blank),
        "output": str(WORK / "pdf-numbered.pdf"),
        "format": "Page {page} of {total}",
        "position": "bottom-center",
    })
    check(numbered["numbered"] == 3, "pdf pageNumbers numbers every page")
    numbered_text = ok({"action": "read", "path": str(WORK / "pdf-numbered.pdf"), "type": "pdf", "args": {}})["text"]
    check("Page 1 of 3" in numbered_text and "Page 3 of 3" in numbered_text, "pdf pageNumbers text is readable")

    protected = ok({
        "action": "pdfop",
        "operation": "encrypt",
        "path": str(blank),
        "output": str(WORK / "pdf-encrypted.pdf"),
        "userPassword": "secret",
        "ownerPassword": "owner",
    })
    check(protected["pages"] == 3, "pdf encrypt preserves pages")
    denied = call({"action": "read", "path": str(WORK / "pdf-encrypted.pdf"), "type": "pdf", "args": {}})
    check(denied.get("ok") is False and denied["error"]["code"] == "PDF_PASSWORD_REQUIRED", "pdf read requires password when encrypted")
    decrypted = ok({
        "action": "pdfop",
        "operation": "decrypt",
        "path": str(WORK / "pdf-encrypted.pdf"),
        "output": str(WORK / "pdf-decrypted.pdf"),
        "password": "secret",
    })
    check(decrypted["pages"] == 3, "pdf decrypt restores page count")

    fixture = WORK / "form_fixture.pdf"
    shutil.copyfile(ROOT / "tests" / "form_fixture.pdf", fixture)
    listed = ok({"action": "pdfop", "operation": "formFields", "path": str(fixture)})
    check(listed["count"] == 1 and listed["fields"][0]["name"] == "name", "pdf formFields lists fixture field")
    filled = ok({
        "action": "pdfop",
        "operation": "fillForm",
        "path": str(fixture),
        "output": str(WORK / "pdf-filled.pdf"),
        "fields": {"name": "John"},
    })
    check(filled["fields"] == 1 and (WORK / "pdf-filled.pdf").is_file(), "pdf fillForm writes output")
    filled_fields = PdfReader(str(WORK / "pdf-filled.pdf")).get_fields() or {}
    check(filled_fields.get("name", {}).get("/V") == "John", "pdf fillForm persists field value")

    image = Image.new("RGB", (48, 48), (220, 40, 40))
    image_pdf = WORK / "pdf-image.pdf"
    image.save(image_pdf, "PDF")
    extracted = ok({
        "action": "pdfop",
        "operation": "extractImages",
        "path": str(image_pdf),
        "outputDir": str(WORK / "pdf-images"),
    })
    check(extracted["count"] >= 1 and Path(extracted["images"][0]["path"]).is_file(), "pdf extractImages writes image files")

    ocr = call({
        "action": "pdfop",
        "operation": "ocr",
        "path": str(blank),
        "pages": [1],
        "language": "eng",
    })
    if ocr.get("ok"):
        check(isinstance(ocr["result"].get("text"), list), "pdf ocr returns page text when backend exists")
    else:
        check(ocr["error"]["code"] == "BACKEND_UNAVAILABLE", "pdf ocr fails cleanly without tesseract")


def test_native_v05() -> None:
    print("[native v0.5]")
    source = WORK / "native-v05.docx"
    ok({"action": "create", "path": str(source), "type": "docx", "overwrite": True, "spec": {"title": "Native", "sections": [{"paragraphs": ["native backend"]}]}})
    target = WORK / "native-v05.pdf"
    result = call({"action": "native", "operation": "exportPdf", "path": str(source), "output": str(target)})
    if result.get("ok"):
        check(target.exists() and target.stat().st_size > 1000, "native exportPdf produced a PDF when COM is available")
        check(result["result"]["backend"] == "com", "native exportPdf reports the COM backend")
    else:
        check(result.get("error", {}).get("code") in {"BACKEND_UNAVAILABLE", "COM_ERROR"}, "native exportPdf reports a classified backend error")


def test_convert_v05() -> None:
    print("[convert v0.5]")
    source = WORK / "legacy.doc"
    source.write_bytes(b"not a real doc")
    target = WORK / "legacy.docx"
    result = call({"action": "convert", "source": str(source), "target": str(target)})
    if result.get("ok"):
        check(target.exists(), "convert used the available COM backend")
    else:
        check(result.get("error", {}).get("code") in {"BACKEND_UNAVAILABLE", "COM_ERROR", "BACKEND_ERROR"}, "convert fails cleanly when no usable backend exists")


def test_batch() -> None:
    print("[batch]")
    result = ok({
        "action": "batch",
        "items": [
            {"id": "a", "action": "create", "path": str(WORK / "batch-a.docx"), "type": "docx", "overwrite": True, "spec": {"title": "批量 A", "sections": []}},
            {"id": "b", "action": "create", "path": str(WORK / "batch-b.xlsx"), "type": "xlsx", "overwrite": True, "spec": {"sheets": [{"name": "S1", "rows": [["x"], ["1"]]}]}},
            {"id": "bad", "action": "read", "path": str(WORK / "missing.docx"), "type": "docx"},
        ],
    })
    check(result["succeeded"] == 2 and result["failed"] == 1, "batch isolated one bad item from two good items")
    check((WORK / "batch-a.docx").exists() and (WORK / "batch-b.xlsx").exists(), "batch outputs exist")


def test_copy() -> None:
    print("[copy]")
    kit_dir = WORK / "_fake_officekit"
    kit_dir.mkdir(exist_ok=True)
    source = kit_dir / "converted.pdf"
    source.write_bytes(b"%PDF-1.4\n% Office Forge copy test\n")
    target = WORK / "exported.pdf"
    result = ok(
        {"action": "copy", "source": str(source), "target": str(target), "overwrite": True},
        {"OFFICEFORGE_OFFICEKIT_DIR": str(kit_dir)},
    )
    check(target.exists() and target.read_bytes() == source.read_bytes(), "copy action copied bytes into workspace")
    check(result["bytes"] == source.stat().st_size, "copy action reported byte count")


def test_docx_v08() -> None:
    print("[docx v0.8]")
    path = WORK / "v08.docx"
    ok({
        "action": "create", "path": str(path), "type": "docx", "overwrite": True,
        "spec": {"title": "v0.8 Word", "sections": [{"heading": "章节", "paragraphs": ["第一段", "第二段"]}]},
    })
    result = ok({
        "action": "edit", "path": str(path), "type": "docx",
        "args": {"operations": [
            {"op": "addBookmark", "paragraph": 3, "name": "目标段落"},
            {"op": "addFootnote", "paragraph": 3, "text": "这是脚注内容"},
            {"op": "addCrossReference", "paragraph": 4, "bookmark": "目标段落"},
        ], "backup": True},
    })
    ops = result["operations"]
    check([item["op"] for item in ops] == ["addBookmark", "addFootnote", "addCrossReference"], "docx v0.8 operations applied")
    bookmark = ops[0]["bookmark"]
    check(bookmark.startswith("bm_"), "docx non-ASCII bookmark hashed to ASCII name")
    check(ops[1]["footnoteId"] == 1, "docx footnote id assigned")
    check(ops[2]["bookmark"] == bookmark, "docx cross reference targets the bookmark")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        body = archive.read("word/document.xml").decode("utf-8")
        footnotes = archive.read("word/footnotes.xml").decode("utf-8") if "word/footnotes.xml" in names else ""
        content_types = archive.read("[Content_Types].xml").decode("utf-8")
    check("word/footnotes.xml" in names, "docx footnotes part created")
    check("footnotes" in content_types, "docx content types declares footnotes")
    check("这是脚注内容" in footnotes, "docx footnote text persisted")
    check("footnoteReference" in body and "bookmarkStart" in body, "docx body has footnote reference and bookmark")
    check("REF " in body and "w:fldChar" in body, "docx cross reference field written")
    disabled = call({
        "action": "edit", "path": str(path), "type": "docx",
        "features": {"word": {"footnotes": False}},
        "args": {"operations": [{"op": "addFootnote", "paragraph": 3, "text": "x"}]},
    })
    check(disabled.get("ok") is False and disabled["error"]["code"] == "FEATURE_DISABLED", "docx footnotes can be disabled")
    disabled_bookmark = call({
        "action": "edit", "path": str(path), "type": "docx",
        "features": {"word": {"bookmarks": False}},
        "args": {"operations": [{"op": "addBookmark", "paragraph": 3, "name": "b2"}]},
    })
    check(disabled_bookmark.get("ok") is False and disabled_bookmark["error"]["code"] == "FEATURE_DISABLED", "docx bookmarks can be disabled")
    duplicate = call({
        "action": "edit", "path": str(path), "type": "docx",
        "args": {"operations": [{"op": "addBookmark", "paragraph": 3, "name": "目标段落"}]},
    })
    check(duplicate.get("ok") is False and duplicate["error"]["code"] == "BOOKMARK_EXISTS", "docx duplicate bookmark rejected")
    missing = call({
        "action": "edit", "path": str(path), "type": "docx",
        "args": {"operations": [{"op": "addCrossReference", "paragraph": 4, "bookmark": "不存在的书签"}]},
    })
    check(missing.get("ok") is False and missing["error"]["code"] == "BOOKMARK_NOT_FOUND", "docx missing bookmark rejected")


def test_docx_deepening() -> None:
    print("[docx deepening]")
    path = WORK / "deep.docx"
    created = ok({
        "action": "create", "path": str(path), "type": "docx", "overwrite": True,
        "spec": {
            "title": "Word 深化",
            "sections": [{"heading": "正文", "paragraphs": ["第一段", "第二段"]}],
            "columns": {"count": 2, "space": 18},
            "watermark": {"text": "内部资料", "color": "C0C0C0", "fontSize": 40},
            "contentControls": [{"tag": "client", "alias": "客户名", "text": "示例客户"}],
            "bibliography": ["张三. 文档工程. 2024.", "李四. 自动化排版. 2025."],
        },
    })
    check(created["bytes"] > 0, "docx deepening created")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        body = archive.read("word/document.xml").decode("utf-8")
        headers = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("word/header") and name.endswith(".xml")
        )
    check('w:num="2"' in body, "docx section columns written")
    check("内部资料" in headers and "v:textpath" in headers, "docx watermark written to header")
    check("<w:sdt" in body and "客户名" in body, "docx content control written")
    check("参考文献" in body and "文档工程" in body, "docx bibliography written")

    edited = ok({
        "action": "edit", "path": str(path), "type": "docx",
        "args": {"backup": False, "operations": [
            {"op": "addCitation", "paragraph": 2, "index": 1},
            {"op": "addBibliography", "entries": ["王五. 引用测试. 2026."], "title": "引用来源"},
            {"op": "saveAsTemplate", "output": str(WORK / "deep.dotx")},
        ]},
    })
    check(edited["operations"][0]["citation"] == "[1]", "docx citation marker inserted")
    template_path = WORK / "deep.dotx"
    check(template_path.is_file(), "docx template file created")
    if template_path.is_file():
        with zipfile.ZipFile(template_path) as archive:
            content_types = archive.read("[Content_Types].xml").decode("utf-8")
        check("wordprocessingml.template.main+xml" in content_types, "docx template content type written")
    disabled_citation = call({
        "action": "edit", "path": str(path), "type": "docx",
        "features": {"word": {"citations": False}},
        "args": {"operations": [{"op": "addCitation", "paragraph": 2, "index": 2}]},
    })
    check(disabled_citation.get("ok") is False and disabled_citation["error"]["code"] == "FEATURE_DISABLED", "docx citations can be disabled")
    disabled_columns = call({
        "action": "edit", "path": str(path), "type": "docx",
        "features": {"word": {"columns": False}},
        "args": {"operations": [{"op": "setColumns", "count": 1}]},
    })
    check(disabled_columns.get("ok") is False and disabled_columns["error"]["code"] == "FEATURE_DISABLED", "docx columns can be disabled")


def test_xlsx_v08() -> None:
    print("[xlsx v0.8]")
    path = WORK / "v08.xlsx"
    result = ok({
        "action": "create", "path": str(path), "type": "xlsx", "overwrite": True,
        "spec": {
            "sheets": [{
                "name": "Sales",
                "rows": [["Region", "Amount"], ["North", 100], ["South", 250]],
                "tables": [{"range": "A1:B3", "name": "SalesTable", "style": "TableStyleMedium9"}],
                "protection": {"password": "pw1"},
            }],
            "workbookProtection": {"password": "pw2", "lockStructure": True},
        },
    })
    counts = result["featureCounts"]
    check(counts["tables"] == 1 and counts["protection"] == 1, "xlsx table and sheet protection created")
    check(counts["workbookProtection"] == 1, "xlsx workbook protection created")
    edited = ok({
        "action": "edit", "path": str(path), "type": "xlsx",
        "args": {"operations": [
            {"op": "addTable", "sheet": "Sales", "range": "D1:E2", "name": "Extra"},
            {"op": "protectSheet", "sheet": "Sales", "password": "pw3", "sort": True},
            {"op": "protectWorkbook", "lockStructure": True},
        ], "backup": True},
    })
    check([item["op"] for item in edited["operations"]] == ["addTable", "protectSheet", "protectWorkbook"], "xlsx v0.8 operations applied")
    import openpyxl
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Sales"]
    check(set(sheet.tables.keys()) == {"SalesTable", "Extra"}, "xlsx table objects persisted")
    check(sheet.protection.sheet is True and sheet.protection.sort is True, "xlsx sheet protection persisted")
    check(workbook.security.lockStructure is True, "xlsx workbook protection persisted")
    disabled_tables = call({
        "action": "edit", "path": str(path), "type": "xlsx",
        "features": {"excel": {"tables": False}},
        "args": {"operations": [{"op": "addTable", "sheet": "Sales", "range": "G1:H2", "name": "Nope"}]},
    })
    check(disabled_tables.get("ok") is False and disabled_tables["error"]["code"] == "FEATURE_DISABLED", "xlsx tables can be disabled")
    disabled_protection = call({
        "action": "edit", "path": str(path), "type": "xlsx",
        "features": {"excel": {"protection": False}},
        "args": {"operations": [{"op": "protectSheet", "sheet": "Sales"}]},
    })
    check(disabled_protection.get("ok") is False and disabled_protection["error"]["code"] == "FEATURE_DISABLED", "xlsx protection can be disabled")
    duplicate = call({
        "action": "edit", "path": str(path), "type": "xlsx",
        "args": {"operations": [{"op": "addTable", "sheet": "Sales", "range": "A1:B3", "name": "SalesTable"}]},
    })
    check(duplicate.get("ok") is False and duplicate["error"]["code"] == "TABLE_EXISTS", "xlsx duplicate table rejected")


def test_xlsx_v10() -> None:
    print("[xlsx v1.0]")
    csv_path = WORK / "v10.csv"
    csv_path.write_text("name,dept,score\nAlice,R&D,10\nBob,Sales,20\n", encoding="utf-8")
    imported = ok({
        "action": "create", "path": str(WORK / "v10-import.xlsx"), "type": "xlsx", "overwrite": True,
        "spec": {"import": {"source": "v10.csv", "header": True}},
    })
    check(imported["sheets"] == ["Imported"], "xlsx CSV import creates sheet")
    import openpyxl
    from openpyxl.worksheet.formula import ArrayFormula
    wb = openpyxl.load_workbook(WORK / "v10-import.xlsx", data_only=False)
    ws = wb["Imported"]
    check(ws["A1"].value == "name" and ws["A2"].value == "Alice" and ws["C3"].value == 20, "xlsx CSV import rows and header")
    json_path = WORK / "v10.json"
    json_path.write_text(json.dumps([{"name": "Carol", "score": 30}, {"name": "Dave", "score": 40}]), encoding="utf-8")
    imported_json = ok({
        "action": "create", "path": str(WORK / "v10-json.xlsx"), "type": "xlsx", "overwrite": True,
        "spec": {"import": {"source": "v10.json", "format": "json", "header": True}},
    })
    check(imported_json["sheets"] == ["Imported"], "xlsx JSON import creates sheet")
    wb_json = openpyxl.load_workbook(WORK / "v10-json.xlsx", data_only=False)
    check(wb_json["Imported"]["A1"].value == "name" and wb_json["Imported"]["B3"].value == 40, "xlsx JSON dict rows imported")

    array_path = WORK / "v10-array.xlsx"
    ok({"action": "create", "path": str(array_path), "type": "xlsx", "overwrite": True, "spec": {"sheets": [{"name": "S", "rows": [["A", "B"], [1, 2], [3, 4]]}]}})
    ok({"action": "edit", "path": str(array_path), "type": "xlsx", "args": {"operations": [{"op": "setCells", "sheet": "S", "cells": [{"cell": "C1", "arrayFormula": "=SUM(A2:B2)", "ref": "C1:C2"}]}]}})
    wb_array = openpyxl.load_workbook(array_path, data_only=False)
    formula = wb_array["S"]["C1"].value
    check(isinstance(formula, ArrayFormula) and formula.ref == "C1:C2" and formula.text == "=SUM(A2:B2)", "xlsx array formula persisted")

    combo_path = WORK / "v10-combo.xlsx"
    combo = ok({
        "action": "create", "path": str(combo_path), "type": "xlsx", "overwrite": True,
        "spec": {"sheets": [{
            "name": "Combo", "columns": ["Cat", "A", "B"], "rows": [["x", 1, 5], ["y", 2, 6], ["z", 3, 7]],
            "charts": [{"type": "combo", "title": "Combo", "categories": ["x", "y", "z"],
                        "series": [{"name": "A", "values": [1, 2, 3], "type": "column", "axis": "primary"},
                                   {"name": "B", "values": [5, 6, 7], "type": "line", "axis": "secondary"}]}],
        }]},
    })
    check(combo["featureCounts"]["charts"] == 1, "xlsx combo chart created")
    import zipfile
    with zipfile.ZipFile(combo_path) as archive:
        chart_xml = archive.read("xl/charts/chart1.xml").decode("utf-8", "replace")
    check("<barChart" in chart_xml and "<lineChart" in chart_xml and 'axId val="200"' in chart_xml, "xlsx combo chart has primary and secondary axes")

    disabled_import = call({
        "action": "create", "path": str(WORK / "v10-no-import.xlsx"), "type": "xlsx", "overwrite": True,
        "features": {"excel": {"import": False}},
        "spec": {"import": {"source": "v10.csv"}},
    })
    check(disabled_import.get("ok") is False and disabled_import["error"]["code"] == "FEATURE_DISABLED", "xlsx import can be disabled")


def test_pptx_v08() -> None:
    print("[pptx v0.8]")
    path = WORK / "v08.pptx"
    ok({
        "action": "create", "path": str(path), "type": "pptx", "overwrite": True,
        "spec": {"slides": [
            {"layout": "title", "title": "A"},
            {"layout": "content", "title": "B", "bullets": ["x"]},
            {"layout": "content", "title": "C", "bullets": ["y"]},
        ]},
    })
    result = ok({
        "action": "edit", "path": str(path), "type": "pptx",
        "args": {"operations": [
            {"op": "setNotesBatch", "notes": [{"slide": 1, "text": "讲稿一"}, {"slide": 2, "text": "讲稿二"}, {"slide": 3, "text": ""}]},
            {"op": "addSection", "name": "开场", "slides": [1]},
            {"op": "addSection", "name": "主体", "slides": [2, 3]},
            {"op": "listSections"},
        ], "backup": True},
    })
    ops = result["operations"]
    check(ops[0]["applied"] == 2 and ops[0]["cleared"] == 1, "pptx notes batch applied and cleared")
    check(ops[3]["sections"] == [{"name": "开场", "slides": [1]}, {"name": "主体", "slides": [2, 3]}], "pptx sections listed")
    from pptx import Presentation
    presentation = Presentation(str(path))
    texts = [slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else "" for slide in presentation.slides]
    check(texts[0] == "讲稿一" and texts[1] == "讲稿二" and texts[2] == "", "pptx notes persisted")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        manifest = archive.read("ppt/presentation.xml").decode("utf-8")
    check("sectionLst" in manifest and "开场" in manifest and "主体" in manifest, "pptx sections persisted")
    disabled_notes = call({
        "action": "edit", "path": str(path), "type": "pptx",
        "features": {"ppt": {"notes": False}},
        "args": {"operations": [{"op": "setNotesBatch", "notes": [{"slide": 1, "text": "z"}]}]},
    })
    check(disabled_notes.get("ok") is False and disabled_notes["error"]["code"] == "FEATURE_DISABLED", "pptx notes can be disabled")
    disabled_sections = call({
        "action": "edit", "path": str(path), "type": "pptx",
        "features": {"ppt": {"sections": False}},
        "args": {"operations": [{"op": "addSection", "name": "x", "slides": [1]}]},
    })
    check(disabled_sections.get("ok") is False and disabled_sections["error"]["code"] == "FEATURE_DISABLED", "pptx sections can be disabled")
    conflict = call({
        "action": "edit", "path": str(path), "type": "pptx",
        "args": {"operations": [{"op": "addSection", "name": "重复", "slides": [1]}]},
    })
    check(conflict.get("ok") is False and conflict["error"]["code"] == "SECTION_CONFLICT", "pptx section conflict rejected")


def test_document_properties() -> None:
    print("[properties v0.9]")
    docx_path = WORK / "props.docx"
    created = ok({
        "action": "create", "type": "docx", "path": str(docx_path),
        "spec": {"title": "元数据", "sections": [{"paragraphs": ["正文"]}]},
    })
    check(created["properties"]["author"] == "Office Forge", "docx default author is neutral")
    check(created["properties"]["revision"] == 1, "docx revision starts at 1")
    read = ok({"action": "read", "path": str(docx_path), "args": {"mode": "content"}})
    metadata = read["metadata"]
    check(metadata["author"] == "Office Forge", "docx stored author is neutral")
    check(metadata["lastModifiedBy"] == "Office Forge", "docx lastModifiedBy is neutral")
    check(not metadata["created"].startswith("2013"), "docx created is not the 2013 template stamp")
    check(metadata["created"].startswith("20"), "docx created is a current timestamp")
    check(metadata["comments"] == "", "docx template comment artefact cleared")

    explicit_path = WORK / "props-explicit.docx"
    ok({
        "action": "create", "type": "docx", "path": str(explicit_path),
        "spec": {"properties": {"author": "张三", "title": "报告"}, "sections": [{"paragraphs": ["正文"]}]},
    })
    explicit = ok({"action": "read", "path": str(explicit_path), "args": {"mode": "content"}})["metadata"]
    check(explicit["author"] == "张三", "docx explicit author wins over default")
    check(explicit["title"] == "报告", "docx explicit title applied")

    legacy_path = WORK / "props-legacy.docx"
    ok({
        "action": "create", "type": "docx", "path": str(legacy_path),
        "spec": {"author": "李四", "sections": [{"paragraphs": ["正文"]}]},
    })
    legacy = ok({"action": "read", "path": str(legacy_path), "args": {"mode": "content"}})["metadata"]
    check(legacy["author"] == "李四", "docx top-level author still supported")

    xlsx_path = WORK / "props.xlsx"
    xlsx_created = ok({
        "action": "create", "type": "xlsx", "path": str(xlsx_path),
        "spec": {"sheets": [{"name": "S", "rows": [[1, 2]]}]},
    })
    check(xlsx_created["properties"]["author"] == "Office Forge", "xlsx default creator is neutral")
    from openpyxl import load_workbook

    workbook = load_workbook(xlsx_path)
    check(workbook.properties.creator == "Office Forge", "xlsx stored creator is neutral")
    check(workbook.properties.lastModifiedBy == "Office Forge", "xlsx stored lastModifiedBy is neutral")
    check(not str(workbook.properties.created).startswith("2013"), "xlsx created is not a template stamp")

    xlsx_named = WORK / "props-named.xlsx"
    ok({
        "action": "create", "type": "xlsx", "path": str(xlsx_named),
        "spec": {"sheets": [{"name": "S", "rows": [[1]]}], "properties": {"creator": "王五", "title": "账"}},
    })
    named = load_workbook(xlsx_named)
    check(named.properties.creator == "王五", "xlsx explicit creator wins")
    check(named.properties.title == "账", "xlsx explicit title applied")

    pptx_path = WORK / "props.pptx"
    pptx_created = ok({
        "action": "create", "type": "pptx", "path": str(pptx_path),
        "spec": {"slides": [{"title": "封面"}]},
    })
    check(pptx_created["properties"]["author"] == "Office Forge", "pptx default author is neutral")
    from pptx import Presentation

    presentation = Presentation(str(pptx_path))
    check(presentation.core_properties.author == "Office Forge", "pptx stored author is neutral")
    check(presentation.core_properties.revision == 1, "pptx stored revision is 1")
    check(not str(presentation.core_properties.created).startswith("2013"), "pptx created is not a template stamp")


def test_pdf_invalid() -> None:
    print("[pdf-invalid]")
    corrupt = WORK / "corrupt.pdf"
    corrupt.write_bytes(b"not a pdf at all\n" * 100)
    data = call({"action": "read", "path": str(corrupt), "format": "text"})
    check(not data.get("ok") and data.get("error", {}).get("code") == "INVALID_PDF", "corrupt PDF rejected as INVALID_PDF")

    empty = WORK / "empty.pdf"
    empty.write_bytes(b"")
    data = call({"action": "read", "path": str(empty), "format": "text"})
    check(not data.get("ok") and data.get("error", {}).get("code") == "INVALID_PDF", "empty PDF rejected as INVALID_PDF")


def main() -> int:
    print(f"workspace: {WORK}")
    try:
        test_probe()
        test_docx()
        test_xlsx()
        test_data_path()
        test_pptx()
        test_docx_features()
        test_xlsx_features()
        test_pptx_features()
        test_docx_v03()
        test_xlsx_v03()
        test_pptx_v03()
        test_docx_v05_plus()
        test_xlsx_v05_plus()
        test_pptx_v05_plus()
        test_docx_v08()
        test_docx_deepening()
        test_xlsx_v08()
        test_xlsx_v10()
        test_pptx_v08()
        test_pptx_deepening()
        test_document_properties()
        test_docx_audit()
        test_xlsx_audit()
        test_pptx_audit()
        test_pdf_v05()
        test_pdf_tools_v05()
        test_pdf_tools_v090()
        test_pdf_invalid()
        test_schema_migration_v05()
        test_native_v05()
        test_convert_v05()
        test_batch()
        test_copy()
        print("\nALL TESTS PASSED")
        return 0
    except Exception as error:
        print(f"\nFAILED: {type(error).__name__}: {error}")
        return 1
    finally:
        if os.environ.get("OFFICEFORGE_KEEP_TEST_DIR") != "1":
            shutil.rmtree(WORK, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
