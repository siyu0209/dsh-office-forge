"""Non-destructive PPTX beautifier and pixel-aligned layout archetype builder."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from ppt_audit import audit_pptx
from themes import (
    get_theme,
    hex_to_pptx_rgb,
    is_dark_color,
    pick_contrast_color,
    pptx_rgb_to_hex,
    wcag_contrast_ratio,
)

EMU_PER_INCH = 914400
WIDESCREEN_W = 13.333
WIDESCREEN_H = 7.5


def _extract_fill_hex(fill: Any) -> str | None:
    try:
        if fill.type == MSO_FILL_TYPE.SOLID:
            return pptx_rgb_to_hex(fill.fore_color.rgb)
    except Exception:
        return None
    return None


def _is_cover_slide(slide: Any, idx: int) -> bool:
    if idx != 1:
        return False
    if any(s.has_chart or s.has_table for s in slide.shapes):
        return False
    text_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    return len(text_shapes) <= 4


def _set_slide_bg(slide: Any, hex_color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_pptx_rgb(hex_color)


def _add_rounded_card(
    slide: Any,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_hex: str,
    border_hex: str | None = None,
    radius: float = 0.04,
    name: str | None = None,
) -> Any:
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    if name:
        shape.name = name
    try:
        shape.adjustments[0] = radius
    except Exception:
        pass
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_pptx_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_to_pptx_rgb(border_hex)
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def _add_rect(
    slide: Any,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_hex: str,
    name: str | None = None,
) -> Any:
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    if name:
        shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_pptx_rgb(fill_hex)
    shape.line.fill.background()
    return shape


def _add_textbox(
    slide: Any,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str,
    font_size_pt: float,
    color_hex: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    line_spacing: float = 1.22,
    name: str | None = None,
) -> Any:
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)

    lines = text.split("\n") if text else [""]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if i > 0:
            p.space_before = Pt(4)
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = bold
        run.font.color.rgb = hex_to_pptx_rgb(color_hex)
    return tb


def _add_slide_header(
    slide: Any,
    theme: dict[str, Any],
    title: str,
    subtitle: str | None = None,
    kicker: str | None = None,
) -> float:
    colors = theme["colors"]
    typo = theme["typography"]
    layout = theme["layout"]
    ml = layout["marginLeftIn"]
    mr = layout["marginRightIn"]
    content_w = WIDESCREEN_W - ml - mr

    # Top accent bar
    _add_rect(slide, 0, 0, WIDESCREEN_W, layout["accentBarHeightIn"], colors["primary"], name="PPTStudioHeaderBar")

    y = layout["marginTopIn"] - 0.1
    if kicker:
        _add_textbox(
            slide,
            ml,
            y,
            content_w,
            0.28,
            kicker.upper(),
            typo["headingFont"],
            10.5,
            colors["secondary"],
            bold=True,
            name="PPTStudioKicker",
        )
        y += 0.25

    _add_textbox(
        slide,
        ml,
        y,
        content_w,
        0.55,
        title,
        typo["headingFont"],
        typo["slideTitlePt"],
        colors["textPrimary"],
        bold=True,
        name="PPTStudioTitle",
    )
    y += 0.56

    if subtitle:
        _add_textbox(
            slide,
            ml,
            y,
            content_w,
            0.36,
            subtitle,
            typo["bodyFont"],
            typo["bodyPt"] - 1,
            colors["textSecondary"],
            bold=False,
            name="PPTStudioSubtitle",
        )
        y += 0.40

    # Subtle divider line under header
    _add_rect(slide, ml, y, content_w, 0.015, colors["border"], name="PPTStudioDivider")
    return y + 0.18


def _add_footer_page_num(slide: Any, theme: dict[str, Any], index: int, total: int) -> None:
    colors = theme["colors"]
    typo = theme["typography"]
    ml = theme["layout"]["marginLeftIn"]
    mr = theme["layout"]["marginRightIn"]
    _add_textbox(
        slide,
        WIDESCREEN_W - mr - 1.2,
        WIDESCREEN_H - 0.42,
        1.2,
        0.28,
        f"{index:02d} / {total:02d}",
        typo["bodyFont"],
        10.5,
        colors["textSecondary"],
        align=PP_ALIGN.RIGHT,
        name="PPTStudioPageNumber",
    )
    _add_rect(slide, ml, WIDESCREEN_H - 0.45, 0.45, 0.03, colors["accent"], name="PPTStudioFooterAccent")


def _snap_near_aligned_shapes(slide: Any, sw: float, sh: float, threshold_in: float = 0.15) -> int:
    """Snap sibling shapes with small alignment drift onto a shared coordinate and clamp bounds."""
    snapped = 0
    candidates = []
    for s in slide.shapes:
        if s.name.startswith("PPTStudio"):
            continue
        w = s.width / EMU_PER_INCH
        h = s.height / EMU_PER_INCH
        if w * h >= 0.72 * sw * sh:
            continue
        if w < 0.08 or h < 0.08:
            continue
        candidates.append(s)

    # Clamp out-of-bounds shapes cleanly back inside slide
    for s in candidates:
        left_in = s.left / EMU_PER_INCH
        top_in = s.top / EMU_PER_INCH
        w_in = s.width / EMU_PER_INCH
        h_in = s.height / EMU_PER_INCH
        changed = False
        if left_in < 0:
            s.left = Inches(0.25)
            changed = True
        elif left_in + w_in > sw:
            new_left = max(0.25, sw - w_in - 0.25)
            s.left = Inches(new_left)
            changed = True
        if top_in < 0:
            s.top = Inches(0.25)
            changed = True
        elif top_in + h_in > sh:
            new_top = max(0.25, sh - h_in - 0.25)
            s.top = Inches(new_top)
            changed = True
        if changed:
            snapped += 1

    # Cluster left coordinates and top coordinates
    for attr in ("left", "top"):
        coords = sorted((getattr(s, attr) / EMU_PER_INCH, idx) for idx, s in enumerate(candidates))
        i = 0
        while i < len(coords):
            group = [coords[i]]
            j = i + 1
            while j < len(coords) and abs(coords[j][0] - group[0][0]) <= threshold_in:
                group.append(coords[j])
                j += 1
            if len(group) >= 2:
                anchor = round(group[0][0], 3)
                for val, s_idx in group:
                    if abs(val - anchor) >= 0.015:
                        setattr(candidates[s_idx], attr, Inches(anchor))
                        snapped += 1
            i = j

    return snapped


def beautify_pptx(
    input_path: str,
    output_path: str | None = None,
    theme_id: str = "executive-navy",
    options: dict[str, Any] | None = None,
) -> dict[str, Any]:
    opts = options or {}
    normalize_fonts = opts.get("normalizeFonts", True)
    apply_theme_colors = opts.get("applyThemeColors", True)
    fix_contrast = opts.get("fixContrast", True)
    fix_alignment = opts.get("fixAlignment", True)
    upgrade_tables = opts.get("upgradeTables", True)
    add_header_accent = opts.get("addHeaderAccent", True)
    add_slide_numbers = opts.get("addSlideNumbers", True)
    enforce_widescreen = opts.get("enforceWidescreen", False)

    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"Input PPTX not found: {input_path}")

    out_path = Path(output_path) if output_path else in_path.with_name(f"{in_path.stem}-beautified{in_path.suffix}")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    before_audit = audit_pptx(str(in_path))
    theme = get_theme(theme_id)
    colors = theme["colors"]
    typo = theme["typography"]
    layout = theme["layout"]

    prs = Presentation(str(in_path))
    sw = prs.slide_width / EMU_PER_INCH
    sh = prs.slide_height / EMU_PER_INCH

    changes_applied: list[str] = []
    aligned_count = 0
    contrast_fixed = 0
    tables_upgraded = 0
    fonts_normalized = 0

    if enforce_widescreen and sw > 0 and sh > 0 and (sw / sh) < 1.5:
        scale_x = WIDESCREEN_W / sw
        scale_y = WIDESCREEN_H / sh
        prs.slide_width = Inches(WIDESCREEN_W)
        prs.slide_height = Inches(WIDESCREEN_H)
        for slide in prs.slides:
            for shape in slide.shapes:
                shape.left = int(shape.left * scale_x)
                shape.top = int(shape.top * scale_y)
                shape.width = int(shape.width * scale_x)
                shape.height = int(shape.height * scale_y)
        sw, sh = WIDESCREEN_W, WIDESCREEN_H
        changes_applied.append("已转换为 16:9 宽屏比例 (13.333\" × 7.5\")")

    total_slides = len(prs.slides)
    for idx, slide in enumerate(prs.slides, start=1):
        is_cover = _is_cover_slide(slide, idx)
        slide_bg = colors["coverBackground"] if is_cover else colors["background"]

        if apply_theme_colors:
            _set_slide_bg(slide, slide_bg)

        existing_names = {s.name for s in slide.shapes}

        for shape in slide.shapes:
            if shape.name.startswith("PPTStudio"):
                continue

            left_in = shape.left / EMU_PER_INCH
            top_in = shape.top / EMU_PER_INCH
            w_in = shape.width / EMU_PER_INCH
            h_in = shape.height / EMU_PER_INCH
            is_full_bleed = (w_in * h_in) >= 0.72 * sw * sh and left_in <= 0.3 and top_in <= 0.3

            if is_full_bleed and apply_theme_colors and not shape.has_text_frame and hasattr(shape, "fill"):
                shape.fill.solid()
                shape.fill.fore_color.rgb = hex_to_pptx_rgb(slide_bg)
                continue

            # Upgrade solid non-picture auto shapes to theme surface
            if (
                apply_theme_colors
                and hasattr(shape, "fill")
                and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                and not is_full_bleed
                and w_in >= 0.5
                and h_in >= 0.35
            ):
                current_fill = _extract_fill_hex(shape.fill)
                if current_fill:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = hex_to_pptx_rgb(colors["surface"])
                    shape.line.color.rgb = hex_to_pptx_rgb(colors["border"])
                    shape.line.width = Pt(0.75)

            shape_fill_hex = (_extract_fill_hex(shape.fill) if hasattr(shape, "fill") else None) or slide_bg

            # Tables
            if upgrade_tables and shape.has_table:
                tbl = shape.table
                for r_idx, row in enumerate(tbl.rows):
                    for cell in row.cells:
                        cell.fill.solid()
                        if r_idx == 0:
                            cell.fill.fore_color.rgb = hex_to_pptx_rgb(colors["primary"])
                            cell_fg = colors["textOnDark"]
                            cell_bold = True
                        else:
                            bg_c = colors["surface"] if r_idx % 2 == 1 else colors["surfaceAlt"]
                            cell.fill.fore_color.rgb = hex_to_pptx_rgb(bg_c)
                            cell_fg = pick_contrast_color(bg_c, colors["textOnDark"], colors["textPrimary"])
                            cell_bold = False
                        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if normalize_fonts:
                                    run.font.name = typo["bodyFont"]
                                    run.font.size = Pt(max(11.0, typo["bodyPt"] - 1.5))
                                run.font.bold = cell_bold
                                run.font.color.rgb = hex_to_pptx_rgb(cell_fg)
                tables_upgraded += 1

            # Text frames
            if shape.has_text_frame:
                tf = shape.text_frame
                is_title = (
                    (shape.is_placeholder and getattr(shape.placeholder_format, "idx", -1) == 0)
                    or (top_in < 1.35 and h_in <= 1.4 and len(tf.text.strip()) <= 80)
                )
                is_footer = top_in > sh - 0.65 and h_in <= 0.45

                for p_idx, p in enumerate(tf.paragraphs):
                    if p.line_spacing is None:
                        p.line_spacing = 1.22
                    for run in p.runs:
                        if not run.text:
                            continue
                        if normalize_fonts:
                            target_font = typo["headingFont"] if is_title else typo["bodyFont"]
                            run.font.name = target_font
                            if is_cover and is_title:
                                run.font.size = Pt(typo["coverTitlePt"])
                                run.font.bold = True
                            elif is_title:
                                run.font.size = Pt(typo["slideTitlePt"])
                                run.font.bold = True
                            elif is_footer:
                                run.font.size = Pt(10.5)
                            else:
                                cur_pt = run.font.size.pt if run.font.size else typo["bodyPt"]
                                if cur_pt < 11.5:
                                    run.font.size = Pt(typo["bodyPt"])
                            fonts_normalized += 1

                        if apply_theme_colors or fix_contrast:
                            try:
                                cur_fg = pptx_rgb_to_hex(run.font.color.rgb) if run.font.color and run.font.color.rgb else None
                            except Exception:
                                cur_fg = None

                            if is_cover:
                                desired = colors["textOnDark"] if is_dark_color(shape_fill_hex) else colors["textPrimary"]
                            elif is_title:
                                desired = colors["primary"] if not is_dark_color(shape_fill_hex) else colors["textOnDark"]
                            else:
                                desired = colors["textPrimary"] if not is_dark_color(shape_fill_hex) else colors["textOnDark"]

                            if apply_theme_colors:
                                if wcag_contrast_ratio(desired, shape_fill_hex) < 4.5:
                                    desired = pick_contrast_color(shape_fill_hex, colors["textOnDark"], colors["textPrimary"])
                                run.font.color.rgb = hex_to_pptx_rgb(desired)
                                contrast_fixed += 1
                            elif fix_contrast and (cur_fg is None or wcag_contrast_ratio(cur_fg, shape_fill_hex) < 4.5):
                                fixed_fg = pick_contrast_color(shape_fill_hex, colors["textOnDark"], colors["textPrimary"])
                                run.font.color.rgb = hex_to_pptx_rgb(fixed_fg)
                                contrast_fixed += 1

        if fix_alignment:
            aligned_count += _snap_near_aligned_shapes(slide, sw, sh)

        if add_header_accent and not is_cover and "PPTStudioHeaderBar" not in existing_names:
            _add_rect(slide, 0, 0, sw, layout["accentBarHeightIn"], colors["primary"], name="PPTStudioHeaderBar")

        if add_slide_numbers and not is_cover and "PPTStudioPageNumber" not in existing_names:
            _add_textbox(
                slide,
                sw - layout["marginRightIn"] - 1.2,
                sh - 0.42,
                1.2,
                0.28,
                f"{idx:02d} / {total_slides:02d}",
                typo["bodyFont"],
                10.5,
                colors["textSecondary"],
                align=PP_ALIGN.RIGHT,
                name="PPTStudioPageNumber",
            )

    if apply_theme_colors:
        changes_applied.append(f"已应用「{theme['name']}」主题配色体系")
    if fonts_normalized:
        changes_applied.append(f"统一并规范 {fonts_normalized} 处文本字号与字体层级")
    if contrast_fixed:
        changes_applied.append(f"校验并优化 {contrast_fixed} 处文本 WCAG 对比度")
    if tables_upgraded:
        changes_applied.append(f"重绘 {tables_upgraded} 个表格为现代斑马纹与主题表头样式")
    if aligned_count:
        changes_applied.append(f"自动吸附对齐/修正越界 {aligned_count} 次")
    if add_header_accent:
        changes_applied.append("为内容页添加品牌顶部强调色条")
    if add_slide_numbers:
        changes_applied.append("添加统一右下角页码")

    prs.save(str(out_path))
    after_audit = audit_pptx(str(out_path))

    return {
        "inputPath": str(in_path),
        "outputPath": str(out_path),
        "theme": theme["id"],
        "slideCount": total_slides,
        "beforeScore": before_audit["score"],
        "afterScore": after_audit["score"],
        "scoreDelta": after_audit["score"] - before_audit["score"],
        "beforeIssueCounts": before_audit["issueCounts"],
        "afterIssueCounts": after_audit["issueCounts"],
        "changesApplied": changes_applied,
    }


def apply_layout(
    pptx_path: str,
    layout_type: str,
    data: dict[str, Any],
    theme_id: str = "executive-navy",
    output_path: str | None = None,
) -> dict[str, Any]:
    """Create or append a pixel-aligned, design-token-driven slide to a PPTX presentation."""
    target_in = Path(pptx_path)
    target_out = Path(output_path) if output_path else target_in
    target_out.parent.mkdir(parents=True, exist_ok=True)

    if target_in.exists():
        prs = Presentation(str(target_in))
    else:
        prs = Presentation()
        prs.slide_width = Inches(WIDESCREEN_W)
        prs.slide_height = Inches(WIDESCREEN_H)

    theme = get_theme(theme_id)
    colors = theme["colors"]
    typo = theme["typography"]
    layout_cfg = theme["layout"]
    ml = layout_cfg["marginLeftIn"]
    mr = layout_cfg["marginRightIn"]
    gutter = layout_cfg["gutterIn"]
    content_w = WIDESCREEN_W - ml - mr

    blank_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]
    slide = prs.slides.add_slide(blank_layout)
    slide_index = len(prs.slides)

    if layout_type == "cover":
        _set_slide_bg(slide, colors["coverBackground"])
        # Left vertical decorative stripe + subtle card accent
        _add_rect(slide, 0, 0, 0.22, WIDESCREEN_H, colors["accent"], name="PPTStudioCoverStripe")
        _add_rect(slide, 0.22, 0, 0.08, WIDESCREEN_H, colors["secondary"], name="PPTStudioCoverStripe2")

        eyebrow = data.get("eyebrow", "PRESENTATION")
        _add_textbox(
            slide,
            1.1,
            1.75,
            10.8,
            0.35,
            eyebrow.upper(),
            typo["headingFont"],
            13,
            colors["accent"],
            bold=True,
        )
        _add_textbox(
            slide,
            1.1,
            2.2,
            10.8,
            1.6,
            data.get("title", "演示文稿主标题"),
            typo["headingFont"],
            typo["coverTitlePt"],
            colors["textOnDark"],
            bold=True,
        )
        if data.get("subtitle"):
            _add_textbox(
                slide,
                1.1,
                3.95,
                10.2,
                0.9,
                data["subtitle"],
                typo["bodyFont"],
                18,
                colors["textSecondary"] if not is_dark_color(colors["coverBackground"]) else "#CBD5E1",
            )
        meta = data.get("meta", "")
        if meta:
            _add_rect(slide, 1.1, 5.45, 2.2, 0.04, colors["accent"])
            _add_textbox(slide, 1.1, 5.65, 9.0, 0.45, meta, typo["bodyFont"], 13, "#CBD5E1")

    elif layout_type == "section_divider":
        _set_slide_bg(slide, colors["coverBackground"])
        _add_rect(slide, 0, 0, WIDESCREEN_W, 0.08, colors["accent"])
        num = str(data.get("sectionNumber", "01"))
        _add_textbox(slide, 1.1, 2.0, 2.0, 1.1, num, typo["headingFont"], 64, colors["accent"], bold=True)
        _add_textbox(slide, 1.1, 3.1, 10.5, 1.0, data.get("title", "章节标题"), typo["headingFont"], 34, colors["textOnDark"], bold=True)
        if data.get("subtitle"):
            _add_textbox(slide, 1.1, 4.2, 9.5, 0.9, data["subtitle"], typo["bodyFont"], 16, "#CBD5E1")

    elif layout_type == "kpi_cards":
        _set_slide_bg(slide, colors["background"])
        top_y = _add_slide_header(slide, theme, data.get("title", "核心指标概览"), data.get("subtitle"), data.get("kicker"))
        cards = data.get("cards", [])[:4]
        if not cards:
            cards = [{"value": "98.5%", "label": "指标示例", "delta": "+12.4%", "note": "补充说明"}]
        n = len(cards)
        card_w = (content_w - gutter * (n - 1)) / n
        card_h = 3.25
        card_y = top_y + 0.15

        for i, card in enumerate(cards):
            cx = ml + i * (card_w + gutter)
            _add_rounded_card(slide, cx, card_y, card_w, card_h, colors["surface"], colors["border"], layout_cfg["cardRadiusRatio"])
            stripe_color = colors["chartPalette"][i % len(colors["chartPalette"])]
            _add_rect(slide, cx + 0.05, card_y, card_w - 0.10, 0.07, stripe_color)

            _add_textbox(slide, cx + 0.24, card_y + 0.28, card_w - 0.48, 0.35, card.get("label", ""), typo["headingFont"], 14, colors["textSecondary"], bold=True)
            _add_textbox(slide, cx + 0.24, card_y + 0.78, card_w - 0.48, 0.95, str(card.get("value", "")), typo["headingFont"], 34 if n <= 3 else 28, colors["primary"], bold=True)
            if card.get("delta"):
                delta_str = str(card["delta"])
                delta_color = colors["danger"] if delta_str.startswith("-") else colors["success"]
                _add_textbox(slide, cx + 0.24, card_y + 1.82, card_w - 0.48, 0.35, delta_str, typo["headingFont"], 14, delta_color, bold=True)
            if card.get("note"):
                _add_textbox(slide, cx + 0.24, card_y + 2.25, card_w - 0.48, 0.80, card["note"], typo["bodyFont"], 12.5, colors["textSecondary"])

        if data.get("takeaway"):
            takeaway_y = card_y + card_h + 0.28
            _add_rounded_card(slide, ml, takeaway_y, content_w, 0.85, colors["surfaceAlt"], colors["border"], 0.06)
            _add_rect(slide, ml, takeaway_y, 0.08, 0.85, colors["accent"])
            _add_textbox(slide, ml + 0.24, takeaway_y + 0.16, content_w - 0.4, 0.55, data["takeaway"], typo["bodyFont"], 13.5, colors["textPrimary"], bold=True, valign=MSO_ANCHOR.MIDDLE)

    elif layout_type == "two_column_comparison":
        _set_slide_bg(slide, colors["background"])
        top_y = _add_slide_header(slide, theme, data.get("title", "方案对比分析"), data.get("subtitle"), data.get("kicker"))
        col_w = (content_w - gutter) / 2
        col_h = WIDESCREEN_H - top_y - 0.85

        left_data = data.get("left", {})
        right_data = data.get("right", {})
        for idx_col, (col_data, header_color) in enumerate([(left_data, colors["secondary"]), (right_data, colors["primary"])]):
            cx = ml + idx_col * (col_w + gutter)
            _add_rounded_card(slide, cx, top_y, col_w, col_h, colors["surface"], colors["border"], 0.03)
            _add_rect(slide, cx, top_y, col_w, 0.58, header_color)
            header_fg = pick_contrast_color(header_color, colors["textOnDark"], "#0F172A")
            _add_textbox(slide, cx + 0.25, top_y, col_w - 0.5, 0.58, col_data.get("title", f"方案 {idx_col + 1}"), typo["headingFont"], 16, header_fg, bold=True, valign=MSO_ANCHOR.MIDDLE)
            items = col_data.get("items", [])
            bullet_text = "\n".join(f"▪  {item}" for item in items)
            _add_textbox(slide, cx + 0.28, top_y + 0.78, col_w - 0.56, col_h - 1.0, bullet_text, typo["bodyFont"], typo["bodyPt"], colors["textPrimary"], line_spacing=1.35)

    elif layout_type == "timeline_steps":
        _set_slide_bg(slide, colors["background"])
        top_y = _add_slide_header(slide, theme, data.get("title", "实施路线图"), data.get("subtitle"), data.get("kicker"))
        steps = data.get("steps", [])[:5]
        if not steps:
            steps = [{"step": "01", "title": "阶段一", "desc": "明确目标与基线"}]
        n = len(steps)
        card_w = (content_w - gutter * (n - 1)) / n
        track_y = top_y + 0.55

        # Connecting horizontal track
        _add_rect(slide, ml + card_w * 0.5, track_y + 0.22, content_w - card_w, 0.04, colors["border"])

        for i, st in enumerate(steps):
            cx = ml + i * (card_w + gutter)
            badge_x = cx + card_w / 2 - 0.25
            badge_color = colors["chartPalette"][i % len(colors["chartPalette"])]
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(badge_x), Inches(track_y), Inches(0.5), Inches(0.5))
            circle.fill.solid()
            circle.fill.fore_color.rgb = hex_to_pptx_rgb(badge_color)
            circle.line.fill.background()
            tf = circle.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = str(st.get("step", f"{i + 1:02d}"))
            r.font.name = typo["headingFont"]
            r.font.size = Pt(13)
            r.font.bold = True
            r.font.color.rgb = hex_to_pptx_rgb(pick_contrast_color(badge_color, colors["textOnDark"], "#0F172A"))

            card_y = track_y + 0.75
            card_h = WIDESCREEN_H - card_y - 0.85
            _add_rounded_card(slide, cx, card_y, card_w, card_h, colors["surface"], colors["border"], 0.04)
            _add_textbox(slide, cx + 0.18, card_y + 0.22, card_w - 0.36, 0.55, st.get("title", ""), typo["headingFont"], 15, colors["primary"], bold=True, align=PP_ALIGN.CENTER)
            _add_textbox(slide, cx + 0.18, card_y + 0.85, card_w - 0.36, card_h - 1.05, st.get("desc", ""), typo["bodyFont"], 13, colors["textSecondary"], align=PP_ALIGN.LEFT, line_spacing=1.3)

    elif layout_type == "chart_with_insights":
        _set_slide_bg(slide, colors["background"])
        top_y = _add_slide_header(slide, theme, data.get("title", "数据洞察与核心结论"), data.get("subtitle"), data.get("kicker"))
        avail_h = WIDESCREEN_H - top_y - 0.85
        left_w = content_w * 0.61
        right_w = content_w - left_w - gutter
        right_x = ml + left_w + gutter

        _add_rounded_card(slide, ml, top_y, left_w, avail_h, colors["surface"], colors["border"], 0.03)
        img_path = data.get("imagePath")
        if img_path and Path(img_path).exists():
            slide.shapes.add_picture(str(img_path), Inches(ml + 0.15), Inches(top_y + 0.15), width=Inches(left_w - 0.3), height=Inches(avail_h - 0.3))
        else:
            _add_textbox(slide, ml + 0.4, top_y + avail_h / 2 - 0.3, left_w - 0.8, 0.6, "图表占位区（可搭配 ppt_generate_chart 嵌入高清图）", typo["bodyFont"], 13, colors["textSecondary"], align=PP_ALIGN.CENTER)

        insights = data.get("insights", [])[:3]
        if not insights:
            insights = [{"title": "核心洞察", "detail": "在此填写关键数据结论与业务建议。"}]
        n_ins = len(insights)
        ins_gap = 0.18
        ins_h = (avail_h - ins_gap * (n_ins - 1)) / n_ins
        for i, ins in enumerate(insights):
            iy = top_y + i * (ins_h + ins_gap)
            _add_rounded_card(slide, right_x, iy, right_w, ins_h, colors["surface"], colors["border"], 0.04)
            _add_rect(slide, right_x, iy + 0.12, 0.07, ins_h - 0.24, colors["chartPalette"][i % len(colors["chartPalette"])])
            _add_textbox(slide, right_x + 0.22, iy + 0.16, right_w - 0.36, 0.36, ins.get("title", ""), typo["headingFont"], 14.5, colors["primary"], bold=True)
            _add_textbox(slide, right_x + 0.22, iy + 0.54, right_w - 0.36, ins_h - 0.70, ins.get("detail", ""), typo["bodyFont"], 12.5, colors["textSecondary"])

    else:
        raise ValueError(
            f"Unsupported layout_type '{layout_type}'. Supported: cover, section_divider, kpi_cards, two_column_comparison, timeline_steps, chart_with_insights"
        )

    if layout_type not in {"cover", "section_divider"}:
        _add_footer_page_num(slide, theme, slide_index, slide_index)

    prs.save(str(target_out))
    return {
        "outputPath": str(target_out),
        "layoutType": layout_type,
        "theme": theme["id"],
        "slideIndex": slide_index,
        "totalSlides": len(prs.slides),
    }
