"""PPTX layout, typography, WCAG contrast, and visual balance auditor."""

from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE

from themes import pptx_rgb_to_hex, wcag_contrast_ratio

EMU_PER_INCH = 914400


@dataclass
class AuditIssue:
    slideIndex: int
    severity: str  # "error" | "warning" | "info"
    category: str  # "bounds" | "overlap" | "typography" | "contrast" | "alignment" | "density"
    shapeName: str
    message: str
    suggestion: str


def _emu_to_in(emu: int | None) -> float:
    if emu is None:
        return 0.0
    return round(emu / EMU_PER_INCH, 3)


def _extract_solid_fill_hex(fill: Any) -> str | None:
    try:
        if fill is not None and fill.type == MSO_FILL_TYPE.SOLID:
            rgb = fill.fore_color.rgb
            return pptx_rgb_to_hex(rgb)
    except Exception:
        return None
    return None


def _shape_fill_hex(shape: Any) -> str | None:
    if not hasattr(shape, "fill"):
        return None
    return _extract_solid_fill_hex(shape.fill)


def _slide_background_hex(slide: Any, presentation: Any) -> str:
    bg = _extract_solid_fill_hex(slide.background.fill)
    if bg:
        return bg
    try:
        layout_bg = _extract_solid_fill_hex(slide.slide_layout.background.fill)
        if layout_bg:
            return layout_bg
    except Exception:
        pass
    try:
        master_bg = _extract_solid_fill_hex(slide.slide_layout.slide_master.background.fill)
        if master_bg:
            return master_bg
    except Exception:
        pass
    return "#FFFFFF"


def _is_full_bleed_background(left: float, top: float, width: float, height: float, sw: float, sh: float) -> bool:
    return (width * height) >= 0.72 * (sw * sh) and left <= 0.3 and top <= 0.3


def _intersection_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    ax1, ay1, aw, ah = a
    bx1, by1, bw, bh = b
    ax2, ay2 = ax1 + aw, ay1 + ah
    bx2, by2 = bx1 + bw, by1 + bh
    ix1 = max(ax1, bx1)
    iy1 = max(ay1, by1)
    ix2 = min(ax2, bx2)
    iy2 = min(ay2, by2)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    return (ix2 - ix1) * (iy2 - iy1)


def _contains(outer: tuple[float, float, float, float], inner: tuple[float, float, float, float], tol: float = 0.08) -> bool:
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner
    return (
        ix >= ox - tol
        and iy >= oy - tol
        and (ix + iw) <= (ox + ow + tol)
        and (iy + ih) <= (oy + oh + tol)
    )


def audit_pptx(input_path: str) -> dict[str, Any]:
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"PPTX file not found: {input_path}")

    prs = Presentation(str(path))
    sw = _emu_to_in(prs.slide_width)
    sh = _emu_to_in(prs.slide_height)
    aspect = round(sw / sh, 3) if sh > 0 else 0.0

    issues: list[AuditIssue] = []
    all_font_families: set[str] = set()
    all_font_sizes: set[float] = set()
    all_colors: set[str] = set()
    text_only_streak = 0

    margin_x = 0.45
    margin_y = 0.35

    for idx, slide in enumerate(prs.slides, start=1):
        slide_bg = _slide_background_hex(slide, prs)
        content_boxes: list[tuple[str, tuple[float, float, float, float], bool]] = []
        has_visual = False
        has_title_like = False
        total_chars = 0
        slide_area = sw * sh
        covered_area = 0.0

        filled_boxes: list[tuple[tuple[float, float, float, float], str]] = []

        for shape in slide.shapes:
            sname = shape.name or f"Shape-{shape.shape_id}"
            left = _emu_to_in(shape.left)
            top = _emu_to_in(shape.top)
            width = _emu_to_in(shape.width)
            height = _emu_to_in(shape.height)
            box = (left, top, width, height)

            if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE) or shape.has_chart or shape.has_table:
                has_visual = True

            is_bg = _is_full_bleed_background(left, top, width, height, sw, sh)
            if is_bg:
                bg_fill = _shape_fill_hex(shape)
                if bg_fill and not shape.has_text_frame:
                    slide_bg = bg_fill
                continue

            # Ignore tiny decorative accent lines/bars (< 0.35 in thinness without text) for margin/overlap checks
            is_thin_accent = (width < 0.15 or height < 0.15) or (
                not (shape.has_text_frame and shape.text_frame.text.strip())
                and (width <= 0.32 or height <= 0.32)
                and (left <= 0.35 or top <= 0.35)
            )

            # 1. Bounds & margin overflow checks
            if left < -0.02 or top < -0.02 or (left + width) > sw + 0.02 or (top + height) > sh + 0.02:
                issues.append(AuditIssue(
                    slideIndex=idx,
                    severity="error",
                    category="bounds",
                    shapeName=sname,
                    message=f"元素超出幻灯片边界 (x={left:.2f}\", y={top:.2f}\", w={width:.2f}\", h={height:.2f}\")。",
                    suggestion="调整元素坐标或宽高，确保完全位于画布内。",
                ))
            elif not is_thin_accent and (
                left < margin_x
                or top < margin_y
                or (left + width) > (sw - margin_x)
                or (top + height) > (sh - margin_y)
            ):
                # Allow small footer/page numbers near bottom
                is_footer = top > sh - 0.55 and height <= 0.4
                if not is_footer:
                    issues.append(AuditIssue(
                        slideIndex=idx,
                        severity="warning",
                        category="bounds",
                        shapeName=sname,
                        message=f"元素贴近边缘安全区外 (建议左右留白 ≥{margin_x}\", 上下 ≥{margin_y}\")。",
                        suggestion="适当内缩边距，增加页面呼吸感。",
                    ))

            has_text = bool(shape.has_text_frame and shape.text_frame.text.strip())
            if not is_thin_accent:
                content_boxes.append((sname, box, has_text))

            shape_fill_hex = _shape_fill_hex(shape)
            if shape_fill_hex:
                all_colors.add(shape_fill_hex)
                filled_boxes.append((box, shape_fill_hex))

            effective_bg = shape_fill_hex
            if effective_bg is None:
                # Inherit the topmost enclosing filled shape behind this transparent textbox
                for fbox, fhex in reversed(filled_boxes):
                    if _contains(fbox, box, tol=0.08):
                        effective_bg = fhex
                        break
            if effective_bg is None:
                effective_bg = slide_bg

            # 2. Typography & WCAG contrast checks
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                total_chars += len(txt)
                line_count = len([p for p in shape.text_frame.paragraphs if p.text.strip()])

                if len(txt) > 220 or line_count > 8:
                    issues.append(AuditIssue(
                        slideIndex=idx,
                        severity="warning",
                        category="density",
                        shapeName=sname,
                        message=f"单个文本框内容过密（{len(txt)} 字符 / {line_count} 段），易形成“文字墙”。",
                        suggestion="拆分为 2-3 个信息卡片或精简为要点短句（建议单框 ≤140 字）。",
                    ))

                checked_contrast = False
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if not run.text.strip():
                            continue
                        if run.font.name:
                            all_font_families.add(run.font.name)
                        if run.font.size:
                            pt = round(run.font.size.pt, 1)
                            all_font_sizes.add(pt)
                            if pt >= 20 or (top < 1.3 and pt >= 18):
                                has_title_like = True
                            is_footer_box = top > sh - 0.65 and height <= 0.45
                            if pt < 11.0 and not is_footer_box:
                                issues.append(AuditIssue(
                                    slideIndex=idx,
                                    severity="warning",
                                    category="typography",
                                    shapeName=sname,
                                    message=f"正文字号偏小 ({pt}pt)，投影或远距离阅读吃力。",
                                    suggestion="正文建议 ≥13pt，注释/页脚建议 ≥10.5pt。",
                                ))
                        # Contrast check once per shape (or per distinct text color)
                        if not checked_contrast:
                            try:
                                fg_hex = pptx_rgb_to_hex(run.font.color.rgb) if run.font.color and run.font.color.rgb else None
                            except Exception:
                                fg_hex = None
                            if fg_hex:
                                all_colors.add(fg_hex)
                                ratio = wcag_contrast_ratio(fg_hex, effective_bg)
                                pt_val = run.font.size.pt if run.font.size else 14.0
                                min_ratio = 3.0 if pt_val >= 18.0 else 4.5
                                if ratio < min_ratio:
                                    issues.append(AuditIssue(
                                        slideIndex=idx,
                                        severity="error",
                                        category="contrast",
                                        shapeName=sname,
                                        message=f"文字颜色 {fg_hex} 与背景 {effective_bg} 对比度过低 ({ratio:.2f}:1 < {min_ratio}:1)。",
                                        suggestion="更换为高对比度文字颜色或使用 ppt_beautify 自动反色修复。",
                                    ))
                                checked_contrast = True

        # Top-level boxes (not strictly contained inside a larger card)
        top_level_indices: set[int] = set()
        for i in range(len(content_boxes)):
            box_i = content_boxes[i][1]
            area_i = box_i[2] * box_i[3]
            is_child = any(
                j != i
                and (content_boxes[j][1][2] * content_boxes[j][1][3]) > area_i * 1.15
                and _contains(content_boxes[j][1], box_i)
                for j in range(len(content_boxes))
            )
            if not is_child:
                top_level_indices.add(i)
                covered_area += min(area_i, slide_area)

        # 3. Overlap check between sibling boxes (excluding card container + inner label nesting)
        for i in range(len(content_boxes)):
            name_a, box_a, text_a = content_boxes[i]
            for j in range(i + 1, len(content_boxes)):
                name_b, box_b, text_b = content_boxes[j]
                if _contains(box_a, box_b) or _contains(box_b, box_a):
                    continue
                inter = _intersection_area(box_a, box_b)
                if inter <= 0.01:
                    continue
                smaller_area = min(box_a[2] * box_a[3], box_b[2] * box_b[3])
                if smaller_area > 0 and (inter / smaller_area) > 0.12:
                    if text_a or text_b:
                        issues.append(AuditIssue(
                            slideIndex=idx,
                            severity="error",
                            category="overlap",
                            shapeName=f"{name_a} ↔ {name_b}",
                            message=f"元素发生明显重叠（重叠面积占比 {inter / smaller_area:.0%}）。",
                            suggestion="调整元素位置或拆分布局，避免文字或卡片相互遮挡。",
                        ))

        # 4. Near-miss alignment check among top-level sibling boxes
        near_miss_reported = False
        top_list = sorted(top_level_indices)
        for p in range(len(top_list)):
            if near_miss_reported:
                break
            i = top_list[p]
            name_a, box_a, _ = content_boxes[i]
            for q in range(p + 1, len(top_list)):
                j = top_list[q]
                name_b, box_b, _ = content_boxes[j]
                dl = abs(box_a[0] - box_b[0])
                dt = abs(box_a[1] - box_b[1])
                if 0.03 <= dl <= 0.16 or 0.03 <= dt <= 0.16:
                    issues.append(AuditIssue(
                        slideIndex=idx,
                        severity="info",
                        category="alignment",
                        shapeName=f"{name_a} / {name_b}",
                        message="相邻元素存在轻微未对齐偏移（0.03\"–0.16\"），影响精致感。",
                        suggestion="启用 ppt_beautify 的对齐吸附（fixAlignment）统一基准线。",
                    ))
                    near_miss_reported = True
                    break

        # 5. Density & visual balance
        coverage = covered_area / slide_area if slide_area > 0 else 0.0
        if coverage > 0.85 and len(content_boxes) > 2:
            issues.append(AuditIssue(
                slideIndex=idx,
                severity="warning",
                category="density",
                shapeName="Slide",
                message=f"页面元素覆盖率偏高 ({coverage:.0%})，留白不足（建议留白 25%–40%）。",
                suggestion="精简冗余边框或减少单页承载的模块数量。",
            ))

        if idx > 1 and not has_title_like and total_chars > 40:
            issues.append(AuditIssue(
                slideIndex=idx,
                severity="info",
                category="typography",
                shapeName="Slide",
                message="未检测到明显的页面主标题（顶部 ≥18pt）。",
                suggestion="每页顶部建议设置 24–28pt 的结论式标题（Action Title）。",
            ))

        if not has_visual and len(content_boxes) <= 2 and idx > 1 and total_chars > 80:
            text_only_streak += 1
            if text_only_streak >= 2:
                issues.append(AuditIssue(
                    slideIndex=idx,
                    severity="warning",
                    category="density",
                    shapeName="Slide",
                    message=f"已连续 {text_only_streak} 页为纯文字布局，缺少图表、卡片或图示。",
                    suggestion="使用 ppt_generate_chart、ppt_generate_diagram 或 ppt_apply_layout 增加视觉锚点。",
                ))
        else:
            text_only_streak = 0

    # Deck-wide typography & color sprawl checks
    if aspect < 1.5:
        issues.append(AuditIssue(
            slideIndex=0,
            severity="info",
            category="bounds",
            shapeName="Deck",
            message=f"当前画布比例为 {sw:.2f}\"×{sh:.2f}\"（非 16:9 宽屏），现代演示推荐使用 16:9 (13.333\"×7.5\")。",
            suggestion="可在 ppt_beautify 中开启 enforceWidescreen。",
        ))
    if len(all_font_families) > 3:
        issues.append(AuditIssue(
            slideIndex=0,
            severity="warning",
            category="typography",
            shapeName="Deck",
            message=f"全稿使用了 {len(all_font_families)} 种字体（{', '.join(sorted(all_font_families)[:5])}），视觉不够统一。",
            suggestion="整套 PPT 建议控制在 1–2 种字体（中英各一）。",
        ))
    if len(all_font_sizes) > 7:
        issues.append(AuditIssue(
            slideIndex=0,
            severity="info",
            category="typography",
            shapeName="Deck",
            message=f"全稿字号种类过多（{len(all_font_sizes)} 种），层级略显杂乱。",
            suggestion="收敛至 4 级字号体系：封面标题 / 页标题 / 小标题 / 正文。",
        ))

    # Score calculation
    deductions = 0
    for issue in issues:
        if issue.severity == "error":
            deductions += 10
        elif issue.severity == "warning":
            deductions += 5
        else:
            deductions += 2
    score = max(0, min(100, 100 - deductions))
    grade = (
        "A" if score >= 90
        else "B" if score >= 80
        else "C" if score >= 70
        else "D" if score >= 60
        else "F"
    )

    counts = {
        "error": sum(1 for i in issues if i.severity == "error"),
        "warning": sum(1 for i in issues if i.severity == "warning"),
        "info": sum(1 for i in issues if i.severity == "info"),
    }

    return {
        "path": str(path),
        "slideCount": len(prs.slides),
        "slideWidthIn": sw,
        "slideHeightIn": sh,
        "aspectRatio": aspect,
        "score": score,
        "grade": grade,
        "issueCounts": counts,
        "fontFamilies": sorted(all_font_families),
        "fontSizesPt": sorted(all_font_sizes),
        "colorCount": len(all_colors),
        "issues": [asdict(i) for i in issues],
    }
