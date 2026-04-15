"""
chart_data_patcher — Patch chart data at the XML level, preserving 100% of formatting.

Instead of python-pptx's ``chart.replace_data()`` (which rebuilds the chart XML
from scratch, wiping fills, gradients, shadows, etc.), this module directly
modifies ``c:cat`` and ``c:val`` caches inside existing ``c:ser`` elements.
All formatting nodes (``c:spPr``, ``c:dPt``, ``c:dLbl``, …) remain untouched.
"""

import logging
from copy import deepcopy
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from lxml import etree

logger = logging.getLogger("report_relay.chart_data_patcher")

# ---------------------------------------------------------------------------
# Chart XML namespaces (same as pptx_exporter)
# ---------------------------------------------------------------------------

_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_C = "{%s}" % _C_NS

PLOT_TAGS = [
    f"{_C}barChart",
    f"{_C}bar3DChart",
    f"{_C}lineChart",
    f"{_C}line3DChart",
    f"{_C}areaChart",
    f"{_C}area3DChart",
    f"{_C}pieChart",
    f"{_C}pie3DChart",
    f"{_C}doughnutChart",
    f"{_C}radarChart",
    f"{_C}scatterChart",
    f"{_C}bubbleChart",
]


# ---------------------------------------------------------------------------
# Value-format detection (extracted from deck_update._looks_percent)
# ---------------------------------------------------------------------------

def _looks_percent_heuristic(values: Sequence) -> bool:
    """Return True if *values* look like decimal fractions (0–1 range)."""
    try:
        numeric = [float(v) for v in values if v is not None]
    except (ValueError, TypeError):
        return False
    if not numeric:
        return False
    mn, mx = min(numeric), max(numeric)
    return 0.0 <= mn <= 1.0 and 0.0 <= mx <= 1.1 and mx >= 0.05


def detect_value_format(
    values: Sequence,
    alt_metadata: Optional[Dict[str, Any]] = None,
) -> str:
    """Determine whether *values* represent percentages or plain numbers.

    Checks for an explicit ``value_format`` key in *alt_metadata* first
    (``"percentage"`` or ``"number"``).  Falls back to the 0-1 range
    heuristic when no override is present.

    Returns ``"percentage"`` or ``"number"``.
    """
    if alt_metadata:
        explicit = alt_metadata.get("value_format") or alt_metadata.get("valueformat")
        if explicit and explicit.strip().lower() in ("percentage", "number"):
            return explicit.strip().lower()
    if _looks_percent_heuristic(values):
        return "percentage"
    return "number"


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def _find_plot(chart_tree: etree._Element) -> Optional[etree._Element]:
    """Locate the first plot element (``c:barChart``, etc.) inside ``c:plotArea``."""
    plot_area = chart_tree.find(f".//{_C}plotArea")
    if plot_area is None:
        return None
    for tag in PLOT_TAGS:
        el = plot_area.find(tag)
        if el is not None:
            return el
    return None


def _format_code_for(value_format: str) -> str:
    if value_format == "percentage":
        return "0.0%"
    return "General"


def _rebuild_str_cache(parent: etree._Element, categories: List[str]):
    """Replace (or create) the ``c:strCache`` under *parent* with new categories."""
    # parent is c:cat or c:tx
    str_ref = parent.find(f"{_C}strRef")
    target = str_ref if str_ref is not None else parent

    old_cache = target.find(f"{_C}strCache")
    if old_cache is not None:
        target.remove(old_cache)

    cache = etree.SubElement(target, f"{_C}strCache")
    pt_count = etree.SubElement(cache, f"{_C}ptCount")
    pt_count.set("val", str(len(categories)))
    for idx, cat in enumerate(categories):
        pt = etree.SubElement(cache, f"{_C}pt")
        pt.set("idx", str(idx))
        v = etree.SubElement(pt, f"{_C}v")
        v.text = str(cat) if cat is not None else ""


def _rebuild_num_cache(
    parent: etree._Element,
    values: List,
    format_code: str = "General",
):
    """Replace (or create) the ``c:numCache`` under *parent* with new values."""
    num_ref = parent.find(f"{_C}numRef")
    target = num_ref if num_ref is not None else parent

    old_cache = target.find(f"{_C}numCache")
    if old_cache is not None:
        target.remove(old_cache)

    cache = etree.SubElement(target, f"{_C}numCache")
    fc = etree.SubElement(cache, f"{_C}formatCode")
    fc.text = format_code
    pt_count = etree.SubElement(cache, f"{_C}ptCount")
    pt_count.set("val", str(len(values)))
    for idx, val in enumerate(values):
        if val is None:
            continue
        pt = etree.SubElement(cache, f"{_C}pt")
        pt.set("idx", str(idx))
        v_el = etree.SubElement(pt, f"{_C}v")
        v_el.text = str(val)


def _rebuild_tx_cache(ser_el: etree._Element, name: str):
    """Update the series name in the ``c:tx`` element."""
    tx = ser_el.find(f"{_C}tx")
    if tx is None:
        tx = etree.SubElement(ser_el, f"{_C}tx")
    _rebuild_str_cache(tx, [name])


def _ensure_cat(ser_el: etree._Element) -> etree._Element:
    """Return the ``c:cat`` element, creating it if absent."""
    cat = ser_el.find(f"{_C}cat")
    if cat is None:
        cat = etree.SubElement(ser_el, f"{_C}cat")
    return cat


def _ensure_val(ser_el: etree._Element) -> etree._Element:
    """Return the ``c:val`` element, creating it if absent."""
    val = ser_el.find(f"{_C}val")
    if val is None:
        val = etree.SubElement(ser_el, f"{_C}val")
    return val


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def patch_chart_data(
    chart,
    categories: List[str],
    values: List,
    value_format: str = "auto",
):
    """Patch a single-series chart's data in-place at the XML level.

    *chart* is a ``python-pptx`` ``Chart`` object (``shape.chart``).
    All existing formatting (fills, gradients, data-label styles, etc.)
    is preserved because only ``c:cat`` / ``c:val`` caches are touched.
    """
    if value_format == "auto":
        value_format = detect_value_format(values)
    fmt_code = _format_code_for(value_format)

    chart_tree = chart.part._element
    plot = _find_plot(chart_tree)
    if plot is None:
        logger.warning("patch_chart_data: no plot element found — skipping")
        return

    series_list = plot.findall(f"{_C}ser")
    if not series_list:
        logger.warning("patch_chart_data: no c:ser elements found — skipping")
        return

    ser = series_list[0]

    cat_el = _ensure_cat(ser)
    _rebuild_str_cache(cat_el, categories)

    val_el = _ensure_val(ser)
    _rebuild_num_cache(val_el, values, fmt_code)

    logger.debug(
        "patch_chart_data: patched %d categories, %d values (format=%s)",
        len(categories), len(values), value_format,
    )


def patch_chart_series(
    chart,
    series_data: List[Tuple[str, List[str], List]],
    value_format: str = "auto",
):
    """Patch a multi-series chart's data in-place at the XML level.

    *series_data* is a list of ``(series_name, categories, values)`` tuples.
    Existing series are updated; extra series are added by cloning the last
    existing one (preserving its formatting); surplus series are removed.
    """
    if not series_data:
        return

    all_values = []
    for _, _, vals in series_data:
        all_values.extend(v for v in vals if v is not None)
    if value_format == "auto":
        value_format = detect_value_format(all_values)
    fmt_code = _format_code_for(value_format)

    chart_tree = chart.part._element
    plot = _find_plot(chart_tree)
    if plot is None:
        logger.warning("patch_chart_series: no plot element found — skipping")
        return

    existing = plot.findall(f"{_C}ser")
    n_needed = len(series_data)

    # Expand: clone last series as formatting template
    while len(existing) < n_needed:
        template = existing[-1] if existing else None
        if template is None:
            break
        new_ser = deepcopy(template)
        idx_val = len(existing)
        idx_el = new_ser.find(f"{_C}idx")
        if idx_el is not None:
            idx_el.set("val", str(idx_val))
        order_el = new_ser.find(f"{_C}order")
        if order_el is not None:
            order_el.set("val", str(idx_val))
        plot.append(new_ser)
        existing = plot.findall(f"{_C}ser")

    # Shrink: remove trailing surplus series
    while len(existing) > n_needed:
        plot.remove(existing[-1])
        existing = plot.findall(f"{_C}ser")

    # Update each series' data
    for i, (name, cats, vals) in enumerate(series_data):
        ser = existing[i]

        idx_el = ser.find(f"{_C}idx")
        if idx_el is not None:
            idx_el.set("val", str(i))
        order_el = ser.find(f"{_C}order")
        if order_el is not None:
            order_el.set("val", str(i))

        _rebuild_tx_cache(ser, name)

        cat_el = _ensure_cat(ser)
        _rebuild_str_cache(cat_el, cats)

        val_el = _ensure_val(ser)
        _rebuild_num_cache(val_el, vals, fmt_code)

    logger.debug(
        "patch_chart_series: patched %d series (format=%s)",
        len(series_data), value_format,
    )
