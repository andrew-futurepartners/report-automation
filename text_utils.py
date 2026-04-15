"""
Shared text utilities for Report Relay.

Centralises base-text parsing/formatting, the "preserve formatting by
updating first run" pattern, and number formatting used across
app.py, deck_update.py, and pptx_exporter.py.
"""

import logging
import re
from typing import Dict, Optional, Union

logger = logging.getLogger("report_relay.text_utils")


# ---------------------------------------------------------------------------
# Number formatting
# ---------------------------------------------------------------------------

def format_number_with_commas(number) -> Optional[str]:
    """Format a number with comma separators for thousands places."""
    if number is None:
        return None
    return f"{number:,}"


# ---------------------------------------------------------------------------
# Base text parsing / formatting
# ---------------------------------------------------------------------------

def parse_base_text(text: str) -> Dict[str, object]:
    """Extract description and N value from a base string.

    Handles patterns like:
        "Base: Total respondents. 1,448 complete surveys."
        "Base: Total respondents. 1448"
        "Base: Total respondents 123"

    Returns:
        {"description": str, "n_value": int | None}
    """
    result: Dict[str, object] = {"description": "", "n_value": None}
    if not text or "Base:" not in text:
        return result

    parts_by_period = text.split(".")
    if len(parts_by_period) >= 2:
        desc = parts_by_period[0].replace("Base:", "").strip()
        desc = desc.rstrip(" =").strip()
        result["description"] = desc

        remainder = ".".join(parts_by_period[1:])
        nums = re.findall(r"[\d,]+", remainder)
        for n in nums:
            cleaned = n.replace(",", "")
            if cleaned.isdigit():
                result["n_value"] = int(cleaned)
                break
    else:
        tokens = text.split()
        if len(tokens) >= 3:
            try:
                base_idx = tokens.index("Base:")
            except ValueError:
                return result
            desc_tokens = []
            for i in range(base_idx + 1, len(tokens)):
                if tokens[i].replace(",", "").isdigit():
                    result["n_value"] = int(tokens[i].replace(",", ""))
                    break
                desc_tokens.append(tokens[i])
            result["description"] = " ".join(desc_tokens).rstrip(" =").strip()

    return result


def format_base_text(description: str, n_value: Optional[int] = None) -> str:
    """Build a canonical base string from description and N value."""
    desc = description or "Total respondents"
    if n_value is not None:
        return f"Base: {desc}. {format_number_with_commas(n_value)} complete surveys."
    return f"Base: {desc}."


# ---------------------------------------------------------------------------
# Shape / cell text update with format preservation
# ---------------------------------------------------------------------------

def _get_text_frame(target):
    """Return the text_frame for a shape or table cell."""
    if hasattr(target, "text_frame"):
        return target.text_frame
    return None


def safe_update_text(target, new_text: str, *, preserve_font: bool = False) -> bool:
    """Update text content while preserving paragraph/run formatting.

    Works with both shapes (via shape.text_frame) and table cells
    (via cell.text_frame).

    When *preserve_font* is True and the first paragraph has multiple runs,
    the paragraph is cleared and a single new run is created with font
    properties copied from the original first run (name, size, bold,
    italic, colour).

    Returns True if the update was applied, False otherwise.
    """
    tf = _get_text_frame(target)
    if tf is None:
        # Last resort for table cells with a .text attribute
        if hasattr(target, "text"):
            target.text = new_text
            return True
        return False

    if not hasattr(tf, "paragraphs") or len(tf.paragraphs) == 0:
        tf.text = new_text
        return True

    paragraph = tf.paragraphs[0]

    if not paragraph.runs:
        run = paragraph.add_run()
        run.text = new_text
        return True

    first_run = paragraph.runs[0]

    if preserve_font and len(paragraph.runs) > 1:
        font_props = _snapshot_font(first_run.font)
        paragraph.clear()
        new_run = paragraph.add_run()
        new_run.text = new_text
        _apply_font_snapshot(new_run.font, font_props)
    else:
        first_run.text = new_text
        # Clear extra runs to prevent concatenation
        if len(paragraph.runs) > 1:
            font_props = _snapshot_font(first_run.font)
            paragraph.clear()
            new_run = paragraph.add_run()
            new_run.text = new_text
            _apply_font_snapshot(new_run.font, font_props)

    return True


def _snapshot_font(font) -> dict:
    """Capture font properties that we want to preserve across a clear."""
    props: dict = {}
    for attr in ("name", "size", "bold", "italic"):
        try:
            props[attr] = getattr(font, attr)
        except Exception:
            props[attr] = None
    try:
        if hasattr(font.color, "rgb") and font.color.rgb is not None:
            props["color_rgb"] = font.color.rgb
    except Exception:
        pass
    return props


def _apply_font_snapshot(font, props: dict):
    """Re-apply previously captured font properties."""
    for attr in ("name", "size", "bold", "italic"):
        val = props.get(attr)
        if val is not None:
            try:
                setattr(font, attr, val)
            except Exception:
                pass
    if "color_rgb" in props:
        try:
            font.color.rgb = props["color_rgb"]
        except Exception:
            pass
