"""Theme token loading, WCAG 2.1 contrast calculation, and font helpers."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
THEMES_FILE = ASSETS_DIR / "themes.json"


def load_themes() -> list[dict[str, Any]]:
    with THEMES_FILE.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data["themes"]


def get_theme(theme_id: str | None = None) -> dict[str, Any]:
    themes = load_themes()
    if not theme_id:
        return themes[0]
    for theme in themes:
        if theme["id"] == theme_id:
            return theme
    valid = ", ".join(t["id"] for t in themes)
    raise ValueError(f"Unknown theme '{theme_id}'. Available themes: {valid}")


def hex_to_rgb_tuple(hex_color: str) -> tuple[int, int, int]:
    cleaned = hex_color.strip().lstrip("#")
    if len(cleaned) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    return (
        int(cleaned[0:2], 16),
        int(cleaned[2:4], 16),
        int(cleaned[4:6], 16),
    )


def hex_to_pptx_rgb(hex_color: str) -> RGBColor:
    r, g, b = hex_to_rgb_tuple(hex_color)
    return RGBColor(r, g, b)


def pptx_rgb_to_hex(rgb: RGBColor | None) -> str | None:
    if rgb is None:
        return None
    return f"#{int(rgb[0]):02X}{int(rgb[1]):02X}{int(rgb[2]):02X}"


def _linearize_channel(channel_8bit: int) -> float:
    srgb = channel_8bit / 255.0
    if srgb <= 0.03928:
        return srgb / 12.92
    return ((srgb + 0.055) / 1.055) ** 2.4


def relative_luminance(hex_color: str) -> float:
    r, g, b = hex_to_rgb_tuple(hex_color)
    return (
        0.2126 * _linearize_channel(r)
        + 0.7152 * _linearize_channel(g)
        + 0.0722 * _linearize_channel(b)
    )


def wcag_contrast_ratio(color_a: str, color_b: str) -> float:
    l1 = relative_luminance(color_a)
    l2 = relative_luminance(color_b)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def is_dark_color(hex_color: str) -> bool:
    return relative_luminance(hex_color) < 0.35


def pick_contrast_color(
    bg_hex: str,
    light_candidate: str = "#F8FAFC",
    dark_candidate: str = "#0F172A",
) -> str:
    light_ratio = wcag_contrast_ratio(bg_hex, light_candidate)
    dark_ratio = wcag_contrast_ratio(bg_hex, dark_candidate)
    return light_candidate if light_ratio >= dark_ratio else dark_candidate


def configure_matplotlib_fonts(preferred_font: str | None = None) -> str:
    """Configure matplotlib to use an installed CJK-capable sans-serif font when available."""
    import matplotlib
    from matplotlib import font_manager

    matplotlib.use("Agg")
    available = {f.name for f in font_manager.fontManager.ttflist}
    candidates = [
        preferred_font,
        "Microsoft YaHei",
        "PingFang SC",
        "Hiragino Sans GB",
        "Noto Sans CJK SC",
        "Noto Sans SC",
        "WenQuanYi Micro Hei",
        "SimHei",
        "Source Han Sans CN",
        "DejaVu Sans",
    ]
    chosen_list = [name for name in candidates if name and name in available]
    if "DejaVu Sans" not in chosen_list:
        chosen_list.append("DejaVu Sans")
    matplotlib.rcParams["font.sans-serif"] = chosen_list
    matplotlib.rcParams["axes.unicode_minus"] = False
    return chosen_list[0]


def env_truthy(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes"}
