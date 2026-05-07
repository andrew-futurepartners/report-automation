"""
auto_mapper.py — AI-powered bootstrap pass that maps un-annotated PowerPoint
shapes to crosstab tables by extracting structural fingerprints from existing
chart/table data and matching them against parsed crosstab tables.

Produces a copy of the PPTX with alt text written onto matched shapes so that
the normal deck_update pipeline can process it without manual pre-mapping.
"""

import json
import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation

from crosstab_parser import parse_workbook

logger = logging.getLogger("report_relay.auto_mapper")

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    union = a | b
    if not union:
        return 0.0
    return len(a & b) / len(union)


# ---------------------------------------------------------------------------
# Detection helpers
# ---------------------------------------------------------------------------

_QCODE_RE = re.compile(r"\(Q\d+[_\d]*\)")

_DATE_PATTERNS = [
    re.compile(
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4}",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{1,2}",
        re.IGNORECASE,
    ),
]

_UNINFORMATIVE_SERIES = frozenset({"%", "series 1", "series1", ""})


def _extract_qcode(text: str) -> Optional[str]:
    """Extract a question code like ``(Q1369)`` from arbitrary text."""
    m = _QCODE_RE.search(text or "")
    return m.group() if m else None


def _is_uninformative_series(names: List[str]) -> bool:
    """True when every series name is a generic placeholder."""
    if not names:
        return True
    return all(_norm(n) in _UNINFORMATIVE_SERIES for n in names)


def _is_trend_chart(categories: List[str]) -> bool:
    """True when most categories look like dates (historical trend data)."""
    if len(categories) < 3:
        return False
    hits = 0
    for cat in categories:
        if any(p.search(cat) for p in _DATE_PATTERNS):
            hits += 1
    return hits / len(categories) >= 0.5


def _fuzzy_match_label(
    label: str,
    candidates: List[str],
    threshold: float = 0.6,
) -> Optional[str]:
    """Find the best fuzzy match for *label* among *candidates*.

    Returns the best matching candidate string or None.
    """
    if not label or not candidates:
        return None
    norm_label = _norm(label)
    best_score = 0.0
    best_match: Optional[str] = None
    for cand in candidates:
        ratio = SequenceMatcher(None, norm_label, _norm(cand)).ratio()
        if ratio > best_score:
            best_score = ratio
            best_match = cand
    if best_score >= threshold:
        return best_match
    return None


def _fuzzy_containment(subset_labels: set, superset_labels: set, threshold: float = 0.7) -> float:
    """Like _containment but with fuzzy matching for individual labels."""
    if not subset_labels:
        return 0.0
    hits = 0
    for label in subset_labels:
        if label in superset_labels:
            hits += 1
            continue
        for cand in superset_labels:
            if SequenceMatcher(None, label, cand).ratio() >= threshold:
                hits += 1
                break
    return hits / len(subset_labels)


# ---------------------------------------------------------------------------
# Fingerprinting
# ---------------------------------------------------------------------------

@dataclass
class ShapeFingerprint:
    """Structural data extracted from a PowerPoint shape (no alt text needed)."""
    slide_idx: int
    shape_idx: int
    shape_name: str
    shape_type: str                        # "chart" | "table" | "text"
    categories: List[str] = field(default_factory=list)
    series_names: List[str] = field(default_factory=list)
    col_headers: List[str] = field(default_factory=list)
    row_labels: List[str] = field(default_factory=list)
    text_content: str = ""
    values_sample: List[float] = field(default_factory=list)
    existing_alt: str = ""


def _extract_chart_fingerprint(shape, slide_idx: int, shape_idx: int) -> Optional[ShapeFingerprint]:
    """Pull categories, series names, and sample values from a chart shape."""
    try:
        chart = shape.chart
    except (ValueError, AttributeError):
        return None

    categories = []
    try:
        for pt in chart.plots[0].categories:
            if pt is not None:
                categories.append(str(pt))
    except Exception:
        pass

    series_names = []
    values_sample = []
    try:
        for s in chart.series:
            try:
                series_names.append(str(s.name) if s.name else "")
            except Exception:
                series_names.append("")
            try:
                for v in s.values:
                    if v is not None:
                        values_sample.append(float(v))
            except Exception:
                pass
    except Exception:
        pass

    if not categories and not series_names:
        return None

    return ShapeFingerprint(
        slide_idx=slide_idx,
        shape_idx=shape_idx,
        shape_name=shape.name or "",
        shape_type="chart",
        categories=categories,
        series_names=series_names,
        values_sample=values_sample[:20],
        existing_alt=_read_existing_alt(shape),
    )


def _extract_table_fingerprint(shape, slide_idx: int, shape_idx: int) -> Optional[ShapeFingerprint]:
    """Pull header row and first-column labels from a table shape."""
    if not shape.has_table:
        return None
    tbl = shape.table
    if len(tbl.rows) < 2 or len(tbl.columns) < 2:
        return None

    col_headers = []
    for c in range(1, len(tbl.columns)):
        col_headers.append(tbl.cell(0, c).text_frame.text.strip())

    row_labels = []
    for r in range(1, len(tbl.rows)):
        row_labels.append(tbl.cell(r, 0).text_frame.text.strip())

    if not col_headers and not row_labels:
        return None

    return ShapeFingerprint(
        slide_idx=slide_idx,
        shape_idx=shape_idx,
        shape_name=shape.name or "",
        shape_type="table",
        col_headers=col_headers,
        row_labels=row_labels,
        existing_alt=_read_existing_alt(shape),
    )


def _extract_text_fingerprint(shape, slide_idx: int, shape_idx: int) -> Optional[ShapeFingerprint]:
    """Pull text content from a text box shape."""
    if not hasattr(shape, "text_frame"):
        return None
    text = shape.text_frame.text.strip()
    if not text or len(text) < 5:
        return None
    return ShapeFingerprint(
        slide_idx=slide_idx,
        shape_idx=shape_idx,
        shape_name=shape.name or "",
        shape_type="text",
        text_content=text,
    )


def extract_all_fingerprints(prs: Presentation) -> List[ShapeFingerprint]:
    """Walk the entire presentation and extract fingerprints from all shapes."""
    fingerprints = []
    for slide_idx, slide in enumerate(prs.slides):
        for shape_idx, shape in enumerate(slide.shapes):
            fp = _extract_chart_fingerprint(shape, slide_idx, shape_idx)
            if fp:
                fingerprints.append(fp)
                continue
            fp = _extract_table_fingerprint(shape, slide_idx, shape_idx)
            if fp:
                fingerprints.append(fp)
                continue
            fp = _extract_text_fingerprint(shape, slide_idx, shape_idx)
            if fp:
                fingerprints.append(fp)
    return fingerprints


# ---------------------------------------------------------------------------
# Structural scoring
# ---------------------------------------------------------------------------

ROW_WEIGHT = 0.65
COL_WEIGHT = 0.35


@dataclass
class MapCandidate:
    """A candidate table match for a fingerprint."""
    table: Dict[str, Any]
    score: float
    row_score: float = 0.0
    col_score: float = 0.0
    orientation: str = "normal"            # "normal" | "flipped"


def _containment(subset: set, superset: set) -> float:
    """What fraction of *subset* appears in *superset*?

    Charts typically show a subset of table rows (excluding Base/Mean) and
    only one of many columns, so containment is a better measure than Jaccard.
    """
    if not subset:
        return 0.0
    return len(subset & superset) / len(subset)


def _sim_blend(fp_set: set, t_set: set, use_fuzzy: bool = False) -> float:
    """Containment/Jaccard blend; optionally fuzzy for abbreviated labels."""
    if use_fuzzy:
        contain = _fuzzy_containment(fp_set, t_set, threshold=0.7)
    else:
        contain = _containment(fp_set, t_set)
    jacc = _jaccard(fp_set, t_set)
    return 0.6 * contain + 0.4 * jacc


def score_fingerprint_against_tables(
    fp: ShapeFingerprint,
    tables: List[Dict[str, Any]],
) -> List[MapCandidate]:
    """Score a single fingerprint against all tables using label overlap.

    Tries both axis orientations for charts (normal and flipped) and keeps
    the higher score.  When series names are uninformative (``%``,
    ``Series 1``), the column axis is ignored so the score is based
    entirely on row/category similarity.
    """
    if fp.shape_type == "chart":
        fp_cats = {_norm(c) for c in fp.categories if c}
        fp_ser = {_norm(s) for s in fp.series_names if s}
        uninformative = _is_uninformative_series(fp.series_names)
    elif fp.shape_type == "table":
        fp_cats = {_norm(r) for r in fp.row_labels if r}
        fp_ser = {_norm(c) for c in fp.col_headers if c}
        uninformative = False
    else:
        return []

    candidates: List[MapCandidate] = []
    for t in tables:
        t_rows = {_norm(r) for r in t.get("row_labels", []) if isinstance(r, str)}
        t_cols = {_norm(c) for c in t.get("col_labels", []) if isinstance(c, str)}

        # --- Normal orientation: categories↔rows, series↔cols ---
        row_sim_n = _sim_blend(fp_cats, t_rows)
        col_sim_n = _sim_blend(fp_ser, t_cols) if not uninformative else 0.0
        rw = 1.0 if uninformative else ROW_WEIGHT
        cw = 0.0 if uninformative else COL_WEIGHT
        score_n = rw * row_sim_n + cw * col_sim_n

        best_score = score_n
        best_row = row_sim_n
        best_col = col_sim_n
        best_orient = "normal"

        # --- Flipped orientation: categories↔cols, series↔rows ---
        if fp.shape_type == "chart" and not uninformative:
            row_sim_f = _sim_blend(fp_cats, t_cols)
            col_sim_f = _sim_blend(fp_ser, t_rows, use_fuzzy=True)
            score_f = ROW_WEIGHT * row_sim_f + COL_WEIGHT * col_sim_f
            if score_f > best_score:
                best_score = score_f
                best_row = row_sim_f
                best_col = col_sim_f
                best_orient = "flipped"

        candidates.append(MapCandidate(
            table=t, score=best_score,
            row_score=best_row, col_score=best_col,
            orientation=best_orient,
        ))

    candidates.sort(key=lambda c: c.score, reverse=True)
    return candidates


# ---------------------------------------------------------------------------
# LLM disambiguation
# ---------------------------------------------------------------------------

def _llm_disambiguate(
    fp: ShapeFingerprint,
    candidates: List[MapCandidate],
    max_candidates: int = 5,
) -> Optional[Tuple[MapCandidate, str]]:
    """Use an LLM to pick the best candidate when structural scoring is ambiguous.

    Returns (chosen_candidate, reason) or None.
    """
    try:
        from openai import OpenAI
    except ImportError:
        logger.error("openai package not installed — skipping LLM disambiguation")
        return None

    top = candidates[:max_candidates]
    if not top:
        return None

    if fp.shape_type == "chart":
        shape_info = {
            "shape_name": fp.shape_name,
            "type": "chart",
            "categories": fp.categories[:15],
            "series_names": fp.series_names,
        }
    elif fp.shape_type == "table":
        shape_info = {
            "shape_name": fp.shape_name,
            "type": "table",
            "column_headers": fp.col_headers,
            "row_labels": fp.row_labels[:15],
        }
    else:
        return None

    prompt_candidates = []
    for i, c in enumerate(top):
        prompt_candidates.append({
            "index": i,
            "title": c.table.get("title", ""),
            "row_labels": c.table.get("row_labels", [])[:15],
            "col_labels": c.table.get("col_labels", []),
            "structural_score": round(c.score, 3),
        })

    user_msg = (
        "I have a PowerPoint shape that needs to be matched to one of several "
        "crosstab data tables. The shape's existing data was extracted as follows:\n\n"
        f"Shape: {json.dumps(shape_info)}\n\n"
        "Candidate tables from the new crosstab data:\n"
        f"{json.dumps(prompt_candidates)}\n\n"
        "Which candidate table is the best match for this shape? "
        "Consider overlap in row labels (categories) and column labels (series/banners). "
        'Return JSON: {"best_index": <int or null>, "reason": "<one sentence>"}'
    )

    try:
        client = OpenAI()
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=150,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You are a data-matching assistant that maps PowerPoint chart/table shapes to their corresponding crosstab data tables."},
                {"role": "user", "content": user_msg},
            ],
        )
        body = json.loads(response.choices[0].message.content)
        idx = body.get("best_index")
        reason = body.get("reason", "")
        if idx is not None and 0 <= idx < len(top):
            return top[idx], reason
    except Exception as e:
        logger.error("LLM disambiguation failed: %s", e)

    return None


# ---------------------------------------------------------------------------
# Mapping result
# ---------------------------------------------------------------------------

@dataclass
class AutoMapResult:
    """Result of auto-mapping one shape."""
    fingerprint: ShapeFingerprint
    table_title: Optional[str]
    col_key: Optional[str]
    confidence: float
    method: str  # "qcode" | "structural" | "llm" | "slide_context" | "trend_skip" | "unmatched"
    reason: str = ""
    candidates: List[MapCandidate] = field(default_factory=list)
    orientation: str = "normal"
    row_key: Optional[str] = None
    column_keys: Optional[List[str]] = None


# ---------------------------------------------------------------------------
# Alt text writer
# ---------------------------------------------------------------------------

def _build_alt_text(
    table_title: str,
    col_key: Optional[str],
    shape_type: str,
    *,
    orientation: str = "normal",
    row_key: Optional[str] = None,
    column_keys: Optional[List[str]] = None,
) -> str:
    """Build the standard alt text format that deck_update expects."""
    lines = [f"table_title: {table_title}"]
    if orientation == "flipped" and row_key:
        lines.append(f"row_key: {row_key}")
    if column_keys:
        lines.append(f"column: {','.join(column_keys)}")
    elif col_key:
        lines.append(f"column: {col_key}")
    if shape_type == "chart":
        lines.append("type: chart")
    elif shape_type == "table":
        lines.append("type: table")
    if orientation == "flipped":
        lines.append("orientation: flipped")
    lines.append("auto_update: yes")
    return "\n".join(lines)


def _build_text_alt_text(table_title: str, text_type: str) -> str:
    """Build alt text for associated text shapes."""
    return f"type: {text_type}\ntable_title: {table_title}"


def _write_alt_text(shape, alt_text: str):
    """Write alt text onto a shape's cNvPr descr attribute."""
    el = shape.element

    c_nv_pr = el.find(
        ".//p:cNvPr",
        namespaces={"p": P_NS},
    )
    if c_nv_pr is not None:
        c_nv_pr.set("descr", alt_text)
        return

    try:
        shape.alternative_text = alt_text
    except (AttributeError, ValueError):
        logger.warning("Could not write alt text to shape '%s'", shape.name)


def _read_existing_alt(shape) -> str:
    """Read existing alt text from a shape."""
    try:
        el = shape.element
        c_nv_pr = el.find(
            ".//p:cNvPr",
            namespaces={"p": P_NS},
        )
        if c_nv_pr is not None:
            return c_nv_pr.get("descr", "")
    except Exception:
        pass
    try:
        return shape.alternative_text or ""
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Text shape association (slide context)
# ---------------------------------------------------------------------------

def _classify_text_shape(text: str) -> Optional[str]:
    """Guess the type of a text shape based on its content.

    Returns "text_question", "text_base", "text_title", or None.
    """
    lower = text.strip().lower()
    if lower.startswith("question:") or lower.startswith("q:"):
        return "text_question"
    if lower.startswith("base:") or lower.startswith("base n"):
        return "text_base"
    if "respondent" in lower and ("base" in lower or "n=" in lower or "n =" in lower):
        return "text_base"
    return None


def associate_text_shapes(
    prs: Presentation,
    map_results: List[AutoMapResult],
) -> List[AutoMapResult]:
    """For each slide that has a matched chart/table, associate nearby text shapes
    with the same table using content heuristics.

    Returns new AutoMapResult entries for the associated text shapes.
    """
    slide_to_table: Dict[int, str] = {}
    for r in map_results:
        if r.table_title and r.fingerprint.shape_type in ("chart", "table"):
            slide_idx = r.fingerprint.slide_idx
            if slide_idx not in slide_to_table:
                slide_to_table[slide_idx] = r.table_title

    text_results: List[AutoMapResult] = []
    slides = list(prs.slides)

    for slide_idx, table_title in slide_to_table.items():
        if slide_idx >= len(slides):
            continue
        slide = slides[slide_idx]
        for shape_idx, shape in enumerate(slide.shapes):
            existing_alt = _read_existing_alt(shape)
            if existing_alt and "table_title" in existing_alt.lower():
                continue

            is_data_shape = False
            try:
                _ = shape.chart
                is_data_shape = True
            except (ValueError, AttributeError):
                pass
            if shape.has_table:
                is_data_shape = True
            if is_data_shape:
                continue

            if not hasattr(shape, "text_frame"):
                continue
            text = shape.text_frame.text.strip()
            if not text or len(text) < 3:
                continue

            text_type = _classify_text_shape(text)
            if text_type is None:
                continue

            fp = ShapeFingerprint(
                slide_idx=slide_idx,
                shape_idx=shape_idx,
                shape_name=shape.name or "",
                shape_type="text",
                text_content=text,
            )
            text_results.append(AutoMapResult(
                fingerprint=fp,
                table_title=table_title,
                col_key=None,
                confidence=0.80,
                method="slide_context",
                reason=f"Text shape classified as '{text_type}' on same slide as '{table_title}'",
            ))

    return text_results


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

HIGH_CONFIDENCE_THRESHOLD = 0.70
LLM_THRESHOLD = 0.40


def auto_map_presentation(
    pptx_in: str,
    crosstab_xlsx: str,
    pptx_out: str,
    *,
    use_llm: bool = True,
    write_alt: bool = True,
    progress_callback=None,
) -> List[AutoMapResult]:
    """Run the full auto-mapping pipeline.

    1. Parse the crosstab workbook.
    2. Extract fingerprints from all PPTX shapes.
    3. Score each fingerprint against tables (structural).
    4. Use LLM for ambiguous matches if enabled.
    5. Associate text shapes via slide context.
    6. Optionally write alt text onto matched shapes and save.

    Returns the list of AutoMapResult for UI review.
    """
    prs = Presentation(pptx_in)
    data = parse_workbook(crosstab_xlsx)
    tables = data["tables"]

    return auto_map_presentation_obj(
        prs, tables,
        pptx_out=pptx_out,
        use_llm=use_llm,
        write_alt=write_alt,
        progress_callback=progress_callback,
    )


def auto_map_presentation_obj(
    prs: Presentation,
    tables: List[Dict[str, Any]],
    *,
    pptx_out: Optional[str] = None,
    use_llm: bool = True,
    write_alt: bool = True,
    progress_callback=None,
) -> List[AutoMapResult]:
    """Core auto-mapping logic operating on already-loaded objects.

    Useful when the Streamlit app already has the Presentation and tables loaded.
    """
    fingerprints = extract_all_fingerprints(prs)
    data_fps = [fp for fp in fingerprints if fp.shape_type in ("chart", "table")]

    # Build Q-code → table index for Tier 0 matching
    qcode_index: Dict[str, List[Dict[str, Any]]] = {}
    for t in tables:
        qc = _extract_qcode(t.get("title", ""))
        if qc:
            qcode_index.setdefault(qc, []).append(t)

    results: List[AutoMapResult] = []
    total = len(data_fps) or 1

    for i, fp in enumerate(data_fps):
        # --- Trend chart detection: skip historical time-series charts ---
        if fp.shape_type == "chart" and _is_trend_chart(fp.categories):
            results.append(AutoMapResult(
                fingerprint=fp, table_title=None, col_key=None,
                confidence=0.0, method="trend_skip",
                reason="Historical trend chart — not updatable from single-wave crosstab",
            ))
            if progress_callback:
                progress_callback((i + 1) / total)
            continue

        # --- Tier 0: Q-code match from existing alt text ---
        fp_qcode = _extract_qcode(fp.existing_alt)
        if fp_qcode and fp_qcode in qcode_index:
            matched_tables = qcode_index[fp_qcode]
            qc_table = matched_tables[0]
            col_key = _infer_col_key(fp, qc_table)
            results.append(AutoMapResult(
                fingerprint=fp, table_title=qc_table["title"], col_key=col_key,
                confidence=0.95, method="qcode",
                reason=f"Q-code {fp_qcode} matched table title",
            ))
            if progress_callback:
                progress_callback((i + 1) / total)
            continue

        # --- Structural scoring ---
        candidates = score_fingerprint_against_tables(fp, tables)
        if not candidates:
            results.append(AutoMapResult(
                fingerprint=fp, table_title=None, col_key=None,
                confidence=0.0, method="unmatched",
                reason="No candidate tables available",
            ))
            if progress_callback:
                progress_callback((i + 1) / total)
            continue

        best = candidates[0]

        if best.score >= HIGH_CONFIDENCE_THRESHOLD:
            orient = best.orientation
            col_key, row_key, column_keys = _infer_keys(fp, best.table, orient)
            title = best.table["title"]
            results.append(AutoMapResult(
                fingerprint=fp, table_title=title, col_key=col_key,
                confidence=best.score, method="structural",
                reason=f"Row overlap {best.row_score:.0%}, col overlap {best.col_score:.0%}",
                candidates=candidates[:5],
                orientation=orient,
                row_key=row_key,
                column_keys=column_keys,
            ))
        elif best.score >= LLM_THRESHOLD and use_llm:
            llm_result = _llm_disambiguate(fp, candidates)
            if llm_result:
                chosen, reason = llm_result
                orient = chosen.orientation
                col_key, row_key, column_keys = _infer_keys(fp, chosen.table, orient)
                title = chosen.table["title"]
                results.append(AutoMapResult(
                    fingerprint=fp, table_title=title, col_key=col_key,
                    confidence=max(chosen.score, 0.75), method="llm",
                    reason=reason,
                    candidates=candidates[:5],
                    orientation=orient,
                    row_key=row_key,
                    column_keys=column_keys,
                ))
            else:
                results.append(AutoMapResult(
                    fingerprint=fp, table_title=None, col_key=None,
                    confidence=best.score, method="unmatched",
                    reason="LLM could not resolve ambiguity",
                    candidates=candidates[:5],
                ))
        else:
            results.append(AutoMapResult(
                fingerprint=fp, table_title=None, col_key=None,
                confidence=best.score, method="unmatched",
                reason=f"Best score {best.score:.2f} below threshold",
                candidates=candidates[:5],
            ))

        if progress_callback:
            progress_callback((i + 1) / total)

    # --- Text shape association ---
    text_results = associate_text_shapes(prs, results)
    results.extend(text_results)

    # --- Write alt text onto shapes ---
    if write_alt:
        slides = list(prs.slides)
        for r in results:
            if r.table_title is None:
                continue
            fp = r.fingerprint
            if fp.slide_idx >= len(slides):
                continue
            slide = slides[fp.slide_idx]
            shapes = list(slide.shapes)
            if fp.shape_idx >= len(shapes):
                continue
            shape = shapes[fp.shape_idx]

            if fp.shape_type in ("chart", "table"):
                alt = _build_alt_text(
                    r.table_title, r.col_key, fp.shape_type,
                    orientation=r.orientation,
                    row_key=r.row_key,
                    column_keys=r.column_keys,
                )
            elif fp.shape_type == "text":
                text_type = _classify_text_shape(fp.text_content)
                if text_type:
                    alt = _build_text_alt_text(r.table_title, text_type)
                else:
                    continue
            else:
                continue

            _write_alt_text(shape, alt)
            logger.info(
                "Wrote alt text for shape '%s' → table '%s' (method=%s, conf=%.2f)",
                fp.shape_name, r.table_title, r.method, r.confidence,
            )

    if pptx_out:
        prs.save(pptx_out)
        logger.info("Saved auto-mapped presentation to %s", pptx_out)

    return results


def _infer_col_key(fp: ShapeFingerprint, table: Dict[str, Any]) -> Optional[str]:
    """Try to infer the column key (data series) from the fingerprint."""
    col_labels = table.get("col_labels", [])
    if not col_labels:
        return None

    if fp.shape_type == "chart" and len(fp.series_names) == 1:
        sn = fp.series_names[0]
        if sn in col_labels:
            return sn

    for preferred in ["Total", "Overall", "All", "Base"]:
        if preferred in col_labels:
            return preferred

    return col_labels[0] if col_labels else None


def _infer_row_key(fp: ShapeFingerprint, table: Dict[str, Any]) -> Optional[str]:
    """For flipped charts, resolve the series name to an actual table row label."""
    row_labels = table.get("row_labels", [])
    if not row_labels or not fp.series_names:
        return None
    for sn in fp.series_names:
        if not sn or _norm(sn) in _UNINFORMATIVE_SERIES:
            continue
        if sn in row_labels:
            return sn
        fuzzy = _fuzzy_match_label(sn, row_labels, threshold=0.55)
        if fuzzy:
            return fuzzy
    return None


def _infer_column_keys(fp: ShapeFingerprint, table: Dict[str, Any]) -> Optional[List[str]]:
    """For flipped charts, map chart categories back to table column labels."""
    col_labels = table.get("col_labels", [])
    if not col_labels or not fp.categories:
        return None
    keys: List[str] = []
    for cat in fp.categories:
        if cat in col_labels:
            keys.append(cat)
        else:
            fuzzy = _fuzzy_match_label(cat, col_labels, threshold=0.6)
            if fuzzy:
                keys.append(fuzzy)
    return keys if keys else None


def _infer_keys(
    fp: ShapeFingerprint,
    table: Dict[str, Any],
    orientation: str,
) -> Tuple[Optional[str], Optional[str], Optional[List[str]]]:
    """Return ``(col_key, row_key, column_keys)`` appropriate for *orientation*."""
    if orientation == "flipped":
        row_key = _infer_row_key(fp, table)
        column_keys = _infer_column_keys(fp, table)
        return None, row_key, column_keys
    return _infer_col_key(fp, table), None, None


# ---------------------------------------------------------------------------
# Convenience: get a summary report for the UI
# ---------------------------------------------------------------------------

def results_to_report(results: List[AutoMapResult]) -> List[dict]:
    """Convert AutoMapResult list into a serializable report for the Streamlit UI."""
    report = []
    for r in results:
        entry = {
            "shape_name": r.fingerprint.shape_name,
            "shape_type": r.fingerprint.shape_type,
            "slide_idx": r.fingerprint.slide_idx,
            "table_title": r.table_title,
            "col_key": r.col_key,
            "confidence": round(r.confidence, 3),
            "method": r.method,
            "reason": r.reason,
            "orientation": r.orientation,
            "row_key": r.row_key,
            "candidates": [
                {
                    "title": c.table.get("title", ""),
                    "score": round(c.score, 3),
                    "row_score": round(c.row_score, 3),
                    "col_score": round(c.col_score, 3),
                }
                for c in r.candidates[:5]
            ],
        }
        report.append(entry)
    return report
