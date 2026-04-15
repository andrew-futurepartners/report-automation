"""
crosstab_parser — Parse Q-style crosstab workbooks into structured table dicts.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger("report_relay.crosstab_parser")

# ---------------------------------------------------------------------------
# Defaults & constants
# ---------------------------------------------------------------------------

DEFAULT_PARSE_OPTIONS: Dict[str, Any] = {
    "min_non_null_cells": 4,
    "footnote_threshold": 0.10,
    "min_block_rows": 2,
    "min_block_cols": 2,
}

FOOTNOTE_SUBSTRING_PATTERNS = [
    "total sample", "unweighted", "weighted", "base n =", "n =",
    "multiple comparison", "false discovery rate", "fdr", "p =", "p<", "p>",
    "significance", "statistical", "confidence", "margin of error",
    "fieldwork", "survey", "methodology", "data collection",
]

FOOTNOTE_PREFIXES = [
    "source:", "note:", "notes:", "n =", "n=", "*", "#", "\u2020", "\u2021",
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _merge_options(user_opts: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    opts = dict(DEFAULT_PARSE_OPTIONS)
    if user_opts:
        opts.update(user_opts)
    return opts


def _find_blocks(df: pd.DataFrame, opts: Dict[str, Any]) -> List[Tuple[int, int]]:
    min_cells = opts["min_non_null_cells"]
    min_rows = opts["min_block_rows"]
    min_cols = opts["min_block_cols"]

    empty_row = df.isna().all(axis=1)
    blocks: List[Tuple[int, int]] = []
    start = None
    for i, is_empty in enumerate(list(empty_row) + [True]):
        if start is None and not is_empty:
            start = i
        elif start is not None and is_empty:
            end = i - 1
            sub = df.iloc[start:end + 1]
            if (sub.notna().sum().sum() >= min_cells
                    and sub.shape[0] >= min_rows
                    and sub.shape[1] >= min_cols):
                blocks.append((start, end))
            start = None
    return blocks


def _strip_edges(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    while df2.shape[0] > 0 and df2.iloc[0].isna().all():
        df2 = df2.iloc[1:]
    while df2.shape[0] > 0 and df2.iloc[-1].isna().all():
        df2 = df2.iloc[:-1]
    while df2.shape[1] > 0 and df2.iloc[:, 0].isna().all():
        df2 = df2.iloc[:, 1:]
    while df2.shape[1] > 0 and df2.iloc[:, -1].isna().all():
        df2 = df2.iloc[:, :-1]
    return df2


def _is_footnote_row(label: str, row_data: pd.Series, threshold: float) -> bool:
    """Determine whether a body row is a footnote.

    Checks prefix kill-list first (always fires), then substring patterns,
    then falls back to the sparse-data percentage threshold.
    """
    label_lower = label.lower().strip()

    for prefix in FOOTNOTE_PREFIXES:
        if label_lower.startswith(prefix):
            return True

    for pattern in FOOTNOTE_SUBSTRING_PATTERNS:
        if pattern in label_lower:
            return True

    total_cols = len(row_data)
    if total_cols > 0:
        non_null_count = row_data.count()
        if (non_null_count / total_cols) < threshold:
            return True

    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_workbook(path: str, parse_options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Parse a Q-style crosstab workbook into a dict.

    Returns::

        {
          "tables": [
            {
              "id": "Sheet1#1",
              "sheet": "Sheet1",
              "title": "Q Age",
              "row_labels": [...],
              "col_labels": [...],
              "values": [[...], ...],
              "meta": {"block_start": 12, "block_end": 35, ...}
            },
            ...
          ]
        }

    *parse_options* can override any key in ``DEFAULT_PARSE_OPTIONS``.
    """
    opts = _merge_options(parse_options)
    footnote_threshold = opts["footnote_threshold"]

    xls = pd.ExcelFile(path)
    tables: List[Dict[str, Any]] = []

    for s in xls.sheet_names:
        raw = xls.parse(s, header=None)
        blocks = _find_blocks(raw, opts)
        logger.debug("Sheet '%s': found %d block(s)", s, len(blocks))

        for bi, (st, en) in enumerate(blocks, start=1):
            sub = raw.iloc[st:en + 1, :]
            sub = _strip_edges(sub)
            if sub.shape[0] < 2 or sub.shape[1] < 2:
                continue

            # --- Title detection fallback (first-row single-cell title) ---
            # Before running header detection, check whether the first row of
            # the block is a standalone title (single non-empty cell in col A).
            title_consumed_first_row = False
            first_row = sub.iloc[0]
            non_empty_first = first_row.dropna()
            if (len(non_empty_first) == 1
                    and non_empty_first.index[0] == sub.columns[0]):
                detected_title = str(non_empty_first.values[0]).strip()
                if detected_title:
                    title_consumed_first_row = True
                    sub = sub.iloc[1:].reset_index(drop=True)
                    if sub.shape[0] < 2:
                        continue
                    logger.debug(
                        "Sheet '%s' block %d: consumed first-row title '%s'",
                        s, bi, detected_title,
                    )

            # --- Header detection ---
            lookahead = min(5, sub.shape[0])
            nn_counts = [int(sub.iloc[r].notna().sum()) for r in range(lookahead)]
            if nn_counts:
                banner_row_idx = int(max(range(len(nn_counts)), key=lambda i: nn_counts[i]))
            else:
                banner_row_idx = 0

            metric_row_idx: Optional[int] = None
            if banner_row_idx > 0:
                above_count = nn_counts[banner_row_idx - 1]
                banner_count = nn_counts[banner_row_idx]
                if above_count >= 2 and above_count <= max(2, int(banner_count * 0.8)):
                    metric_row_idx = banner_row_idx - 1

            # Build header labels
            banner_row = sub.iloc[banner_row_idx].fillna("").astype(str).tolist()
            if metric_row_idx is not None:
                raw_metric_row = sub.iloc[metric_row_idx].astype(str).tolist()
                metric_row_ff = []
                last = ""
                for cell in raw_metric_row:
                    c = "" if cell is None or str(cell).strip() == "nan" else str(cell)
                    if c.strip():
                        last = c
                    metric_row_ff.append(last)
            else:
                metric_row_ff = [""] * len(banner_row)

            # Body starts after the banner row
            body = sub.iloc[banner_row_idx + 1:].reset_index(drop=True)

            # First column = row labels; remaining columns = numeric data
            row_labels_raw = body.iloc[:, 0].fillna("").astype(str).tolist()
            data_part_raw = body.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")

            # --- Footnote filtering ---
            rows_to_keep: List[int] = []
            footnotes_removed = 0
            for i, label in enumerate(row_labels_raw):
                if not label.strip():
                    continue
                if _is_footnote_row(label, data_part_raw.iloc[i], footnote_threshold):
                    footnotes_removed += 1
                    continue
                rows_to_keep.append(i)

            if footnotes_removed:
                logger.debug(
                    "Sheet '%s' block %d: removed %d footnote row(s)",
                    s, bi, footnotes_removed,
                )

            if rows_to_keep:
                row_labels = [row_labels_raw[i] for i in rows_to_keep]
                data_part = data_part_raw.iloc[rows_to_keep].reset_index(drop=True)
            else:
                row_labels = row_labels_raw
                data_part = data_part_raw

            data_part = data_part.where(pd.notna(data_part), None)

            # --- Column labels ---
            col_banners = [str(b).strip() for b in banner_row[1:len(data_part.columns) + 1]]
            col_groups = metric_row_ff[1:len(data_part.columns) + 1]
            col_groups = [g.strip() if isinstance(g, str) and g.strip() != "" else "" for g in col_groups]
            col_labels = [
                (f"{b} | {g}" if g else str(b))
                for g, b in zip(col_groups, col_banners)
            ]

            # --- Title resolution ---
            title = None
            if title_consumed_first_row:
                title = detected_title
            else:
                top_header_idx = metric_row_idx if metric_row_idx is not None else banner_row_idx
                for r in range(int(top_header_idx)):
                    row_vals = sub.iloc[r].dropna().astype(str).tolist()
                    if row_vals:
                        title = row_vals[0]
                        break
            if not title:
                title = f"{s} table {bi}"

            table_id = f"{s}#{bi}"
            tdict = {
                "id": table_id,
                "sheet": s,
                "title": title,
                "row_labels": row_labels,
                "col_labels": col_labels,
                "values": data_part.values.tolist(),
                "meta": {
                    "block_start": int(st),
                    "block_end": int(en),
                    "col_banners": col_banners,
                    "col_groups": col_groups,
                },
            }
            tables.append(tdict)

    logger.info("Parsed %d table(s) from '%s'", len(tables), path)
    return {"tables": tables}


def to_json(data: Dict[str, Any]) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)
