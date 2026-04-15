"""
Future Partners Brand Configuration
Extracted from FuturePartners_BrandGuidelines2023

Import this module in pptx_exporter.py and any other file that needs brand values.
Never hardcode colors, fonts, or sizing elsewhere.
"""

from pptx.dml.color import RGBColor
from pptx.util import Pt
from typing import List


# ---------------------------------------------------------------------------
# Core Brand Colors (Brand Palette)
# Use for slide backgrounds, text, logos, UI elements
# Black and Cream are the ONLY colors approved for copy/text
# ---------------------------------------------------------------------------

FP_CREAM   = RGBColor(0xFC, 0xFF, 0xEE)   # #FCFFEE — primary slide background
FP_BLACK   = RGBColor(0x00, 0x00, 0x00)   # #000000 — primary text
FP_TAN     = RGBColor(0xE4, 0xD6, 0xBB)   # #E4D6BB — warm neutral
FP_BLUE    = RGBColor(0x00, 0x6A, 0xFF)   # #006AFF — primary accent
FP_RED     = RGBColor(0xF2, 0x54, 0x2A)   # #F2542A — secondary accent
FP_PURPLE  = RGBColor(0x88, 0x94, 0xFF)   # #8894FF — secondary accent
FP_GREEN   = RGBColor(0xF3, 0xFF, 0xAF)   # #F3FFAF — soft accent (same as Green #5)


# ---------------------------------------------------------------------------
# Infographic Palettes
# 4 palettes × 5 tones each. NEVER mix colors across palettes in one chart.
# Use get_palette() to retrieve the right set for a given chart.
# Palette index 1 = darkest/most saturated, 5 = lightest
# ---------------------------------------------------------------------------

PALETTE_BLUE = [
    RGBColor(0x00, 0x6A, 0xFF),   # FP Blue #1  — #006AFF
    RGBColor(0x25, 0x8D, 0xFF),   # FP Blue #2  — #258DFF
    RGBColor(0x49, 0xB1, 0xFF),   # FP Blue #3  — #49B1FF
    RGBColor(0x6E, 0xD4, 0xFF),   # FP Blue #4  — #6ED4FF
    RGBColor(0x92, 0xF7, 0xFF),   # FP Blue #5  — #92F7FF
]

PALETTE_GREEN = [
    RGBColor(0xC7, 0xE7, 0x0D),   # FP Green #1 — #C7E70D
    RGBColor(0xD5, 0xF3, 0x2B),   # FP Green #2 — #D5F32B
    RGBColor(0xE2, 0xFA, 0x58),   # FP Green #3 — #E2FA58
    RGBColor(0xEA, 0xFD, 0x83),   # FP Green #4 — #EAFD83
    RGBColor(0xF3, 0xFF, 0xAF),   # FP Green #5 — #F3FFAF
]

PALETTE_PURPLE = [
    RGBColor(0x88, 0x94, 0xFF),   # FP Purple #1 — #8894FF
    RGBColor(0x9C, 0xA6, 0xFF),   # FP Purple #2 — #9CA6FF
    RGBColor(0xAF, 0xB7, 0xFF),   # FP Purple #3 — #AFB7FF
    RGBColor(0xC3, 0xC9, 0xFF),   # FP Purple #4 — #C3C9FF
    RGBColor(0xD6, 0xDA, 0xFF),   # FP Purple #5 — #D6DAFF
]

PALETTE_RED = [
    RGBColor(0xF2, 0x54, 0x2A),   # FP Red #1 — #F2542A
    RGBColor(0xF5, 0x76, 0x4F),   # FP Red #2 — #F5764F
    RGBColor(0xF9, 0x97, 0x75),   # FP Red #3 — #F99775
    RGBColor(0xFC, 0xB9, 0x9A),   # FP Red #4 — #FCB99A
    RGBColor(0xFF, 0xDA, 0xBF),   # FP Red #5 — #FFDABF
]

# Ordered list for cycling across charts in a single deck.
# Each chart picks one palette and uses tones top-down by number of series.
ALL_PALETTES = [PALETTE_BLUE, PALETTE_GREEN, PALETTE_PURPLE, PALETTE_RED]
PALETTE_NAMES = ["blue", "green", "purple", "red"]


def get_palette(name: str = "blue") -> List[RGBColor]:
    """
    Return an infographic palette by name.
    Valid names: 'blue', 'green', 'purple', 'red'
    Defaults to blue if name is unrecognised.
    """
    mapping = {
        "blue":   PALETTE_BLUE,
        "green":  PALETTE_GREEN,
        "purple": PALETTE_PURPLE,
        "red":    PALETTE_RED,
    }
    return mapping.get(name.lower(), PALETTE_BLUE)


def get_chart_colors(n_series: int, palette_name: str = "blue") -> List[RGBColor]:
    """
    Return exactly n_series colors from the named palette, following the
    brand guideline rules for tone selection:
      1 series  → shade #1 only
      2 series  → shades #1, #5
      3 series  → shades #1, #3, #5
      4 series  → shades #1, #2, #4, #5
      5 series  → shades #1, #2, #3, #4, #5
    For n > 5, cycle back from the start (edge case, rarely needed).
    """
    palette = get_palette(palette_name)   # list of 5 RGBColor objects (index 0–4)

    selection_map = {
        1: [0],
        2: [0, 4],
        3: [0, 2, 4],
        4: [0, 1, 3, 4],
        5: [0, 1, 2, 3, 4],
    }

    if n_series <= 0:
        return []
    if n_series <= 5:
        indices = selection_map[n_series]
    else:
        # More than 5 series: repeat palette cyclically
        indices = [(i % 5) for i in range(n_series)]

    return [palette[i] for i in indices]


# ---------------------------------------------------------------------------
# Typography
# Brand fonts: GT America Extended Bold (headlines), GT America Regular (body)
# System font fallbacks used here since custom fonts require installation
# ---------------------------------------------------------------------------

# Primary: GT America Extended Bold → system fallback: Arial Black
FONT_HEADLINE   = "Arial Black"

# Primary: GT America Regular → system fallback: Arial
FONT_BODY       = "Arial"

# Primary: Self Modern (narrative/quotes) → system fallback: Times New Roman
FONT_NARRATIVE  = "Times New Roman"

# Font sizes (points) — matching brand hierarchy
FONT_SIZE = {
    "slide_title":    Pt(24),   # Slide-level title (scaled down from web 128pt)
    "section_title":  Pt(18),   # Section/category headers
    "subhead":        Pt(14),   # Chart subheadings
    "body":           Pt(11),   # Body copy, base text, question text
    "data_label":     Pt(10),   # Chart data labels
    "axis":           Pt(9),    # Axis tick labels
    "footnote":       Pt(8),    # Base N, footnotes
    "callout":        Pt(12),   # Text callouts on charts
}


# ---------------------------------------------------------------------------
# Slide Layout & Backgrounds
# ---------------------------------------------------------------------------

# Default slide background: FP Cream
SLIDE_BG_COLOR = FP_CREAM

# Text color on cream background
TEXT_COLOR_PRIMARY   = FP_BLACK
TEXT_COLOR_SECONDARY = RGBColor(0x55, 0x55, 0x55)  # FP Black @ ~33% opacity approx

# Gridline color for charts: FP Black @ 30%
GRIDLINE_COLOR = RGBColor(0xB3, 0xB3, 0xB3)   # approximate 30% black on cream

# Slide dimensions: standard 16:9 widescreen
SLIDE_WIDTH_IN  = 13.333
SLIDE_HEIGHT_IN = 7.5


# ---------------------------------------------------------------------------
# Chart Style Defaults
# Per guidelines: single-color per chart, use one palette set, no mixing.
# Bar charts: ~20% gap width between columns.
# Data labels: GT America Medium (use Arial bold as system fallback).
# ---------------------------------------------------------------------------

CHART_DEFAULTS = {
    "default_palette":      "blue",
    "data_label_font":      FONT_BODY,
    "data_label_bold":      True,
    "data_label_size":      FONT_SIZE["data_label"],
    "axis_font":            FONT_BODY,
    "axis_size":            FONT_SIZE["axis"],
    "show_legend":          False,
    "show_gridlines":       True,
    "gridline_color":       GRIDLINE_COLOR,
    "gap_width":            20,           # ~20% of column width per brand guidelines
    "overlap":              -100,         # 200% spacing between groups for grouped charts
}


def get_data_label_format(values_sample) -> str:
    """Return the correct number format string based on the data range.
    Proportions (0-1) → '0.0%', counts/raw numbers → '#,##0.0'."""
    nums = []
    for v in values_sample:
        if v is not None:
            try:
                nums.append(float(v))
            except (TypeError, ValueError):
                pass
    if nums and all(0.0 <= n <= 1.01 for n in nums) and max(nums, default=0) <= 1.01:
        return "0.0%"
    return "#,##0.0"

# Chart palette rotation for multi-slide decks.
# Assign palette by index to spread color across the report.
def get_palette_for_table_index(idx: int) -> str:
    """Cycle through palettes so consecutive slides use different palette families."""
    return PALETTE_NAMES[idx % len(PALETTE_NAMES)]


# ---------------------------------------------------------------------------
# Executive Summary Slide Styling
# ---------------------------------------------------------------------------

EXEC_SUMMARY = {
    "slide_bg":         FP_CREAM,
    "title_text":       "Executive Summary",
    "title_font":       FONT_HEADLINE,
    "title_size":       Pt(28),
    "title_color":      FP_BLACK,
    "bullet_font":      FONT_BODY,
    "bullet_size":      Pt(11),
    "bullet_color":     FP_BLACK,
    "accent_color":     FP_BLUE,          # Use FP Blue for highlight labels
    "max_bullets":      12,               # Cap before wrapping to a second summary slide
}


# ---------------------------------------------------------------------------
# AI Insight Text Box Styling
# Per guidelines: body copy uses GT America Regular; keep it clean and readable
# ---------------------------------------------------------------------------

AI_INSIGHT = {
    "font":         FONT_BODY,
    "size":         FONT_SIZE["body"],
    "color":        TEXT_COLOR_PRIMARY,
    "bold":         False,
    "bg_color":     None,                 # transparent; sits on cream slide bg
    # Position (x, y, w, h) in inches — bottom strip of each slide
    "position":     (0.5, 6.2, 12.3, 0.9),
}


# ---------------------------------------------------------------------------
# AI Insight Generation Settings
# Controls the LLM calls in ai_insights.py.  Environment variable AI_MODEL
# takes precedence over the "model" value below.
# ---------------------------------------------------------------------------

AI_GENERATION = {
    "model":          "gpt-4o-mini",
    "max_tokens":     250,
    "temperature":    None,               # None → use API default
    "max_concurrent": 5,
    "cache_dir":      ".ai_cache",
    "system_prompt": (
        "You are a senior research analyst at a creative insights firm specializing in "
        "travel, tourism, and hospitality. You write with the confident, clear voice of "
        "an experienced analyst — professional yet conversational. "
        "Ground every statement in specific data: percentages, averages, and segment "
        "differences. Interpret the numbers — explain what they mean and why they matter "
        "for destination marketers and tourism stakeholders. "
        "Write in present tense. Use active language. "
        "Never use filler phrases like 'the data shows', 'it is clear that', or "
        "'interestingly'. Never mention that you are an AI. "
        "Avoid jargon. Highlight differences by segment where meaningful."
    ),
    "system_prompt_wave": (
        " When previous-wave data is provided, explicitly compare old and new values. "
        "Highlight meaningful shifts (e.g. 'increased from 28% to 34%') and interpret "
        "what the change signals for strategy."
    ),
}


# ---------------------------------------------------------------------------
# Convenience: flat BRAND dict for drop-in replacement in pptx_exporter.py
# Replace the hardcoded BRAND dict at the top of pptx_exporter.py with:
#   from brand_config import BRAND
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Template Layout Indices (from Template_ReportSlides.pptx)
# ---------------------------------------------------------------------------

import os as _os

_BASE_DIR = _os.path.dirname(_os.path.abspath(__file__))
TEMPLATE_PATH = _os.path.join(_BASE_DIR, "templates", "Template_ReportSlides.pptx")
CHART_TEMPLATES_DIR = _os.path.join(_BASE_DIR, "templates", "chart_templates")

LAYOUT = {
    "title_slide":        0,   # Title Slide
    "section_header":     4,   # Section Header
    "primary_chart":     16,   # Primary Chart Template (chart + left analysis)
    "three_chart":       17,   # Three Chart Layout
    "one_two_third":     18,   # 1 and 2/3 Chart
    "one_two_third_alt": 19,   # 1 and 2/3 Chart Alt (chart + table appendix)
    "fifty_fifty":       20,   # Fifty Fifty Chart
    "big_chart":         21,   # Big Chart Template (full-width chart)
    "overview":          15,   # Overview
    "blank":             28,   # Blank
}

# Placeholder indices for the Primary Chart Template layout (layout 16)
PH = {
    "title":       0,    # Long Action Takeaway (Title)
    "analysis":    1,    # Supporting Analysis (left body, OBJECT)
    "chart":       2,    # Chart area (right content, OBJECT)
    "footer":     11,    # Footer
    "slide_num":  12,    # Slide Number
    "punch":      13,    # Punch Point placeholder (always removed)
    "qbase":      14,    # Question/Base (bottom right)
    "chart_title":15,    # Chart Title (above chart, right side)
    "note":       16,    # Note / appendix reference (bottom left)
}


# ---------------------------------------------------------------------------
# Convenience: flat BRAND dict for backward compatibility
# ---------------------------------------------------------------------------

BRAND = {
    "font_family_head":  FONT_HEADLINE,
    "font_family_body":  FONT_BODY,
    "title_size":        24,              # points as int (pptx_exporter uses Pt() internally)
    "axis_size":         9,
    "label_size":        10,
    "bg_color":          SLIDE_BG_COLOR,
    # Default single-series chart colors = Blue palette shades 1–5
    "colors":            PALETTE_BLUE,
}
