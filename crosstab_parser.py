
import pandas as pd
from typing import List, Dict, Any, Tuple, Optional
import json

def _find_blocks(df: pd.DataFrame) -> List[Tuple[int, int]]:
    empty_row = df.isna().all(axis=1)
    blocks = []
    start = None
    for i, is_empty in enumerate(list(empty_row) + [True]):
        if start is None and not is_empty:
            start = i
        elif start is not None and is_empty:
            end = i - 1
            sub = df.iloc[start:end+1]
            if sub.notna().sum().sum() >= 10 and sub.shape[0] >= 2 and sub.shape[1] >= 2:
                blocks.append((start, end))
            start = None
    return blocks

def _strip_edges(df: pd.DataFrame) -> pd.DataFrame:
    # Drop fully empty rows and columns at edges
    df2 = df.copy()
    while df2.shape[0] > 0 and df2.iloc[0].isna().all():
        df2 = df2.iloc[1:]
    while df2.shape[0] > 0 and df2.iloc[-1].isna().all():
        df2 = df2.iloc[:-1]
    while df2.shape[1] > 0 and df2.iloc[:,0].isna().all():
        df2 = df2.iloc[:,1:]
    while df2.shape[1] > 0 and df2.iloc[:,-1].isna().all():
        df2 = df2.iloc[:,:-1]
    return df2

def parse_workbook(path: str) -> Dict[str, Any]:
    """
    Parse a Q-style crosstab workbook into a dict:
    {
      "tables": [
        {
          "id": "Sheet1#1",
          "sheet": "Sheet1",
          "title": "Q Age",
          "row_labels": [...],
          "col_labels": [...],
          "values": [[...],[...],...],
          "meta": {"source_range":"A12:N35"}
        },
        ...
      ]
    }
    Heuristics:
      - Split by empty rows
      - Assume first non-empty row contains column headers
      - Assume first column contains row labels
    """
    xls = pd.ExcelFile(path)
    tables = []
    for s in xls.sheet_names:
        raw = xls.parse(s, header=None)
        blocks = _find_blocks(raw)
        for bi, (st, en) in enumerate(blocks, start=1):
            sub = raw.iloc[st:en+1, :]
            sub = _strip_edges(sub)
            if sub.shape[0] < 2 or sub.shape[1] < 2:
                continue

            # Detect header structure (one-row vs two-row: Metric row + Banner row)
            # Strategy: within first 5 rows, choose banner row as the row with max
            # non-null cells; if the row above has significantly fewer values, treat it
            # as a metric row and forward-fill its labels.
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
                # Treat the row above as a metric row only if it looks like a
                # true header (not a single-cell question line): require >=2
                # non-empty cells and not almost as dense as the banner row.
                if above_count >= 2 and above_count <= max(2, int(banner_count * 0.8)):
                    metric_row_idx = banner_row_idx - 1

            # Build header labels
            banner_row = sub.iloc[banner_row_idx].fillna("").astype(str).tolist()
            if metric_row_idx is not None:
                raw_metric_row = sub.iloc[metric_row_idx].astype(str).tolist()
                # Forward-fill across columns to simulate Excel merged cells
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
            body = sub.iloc[banner_row_idx+1:].reset_index(drop=True)

            # First column are row labels
            row_labels_raw = body.iloc[:,0].fillna("").astype(str).tolist()
            data_part_raw = body.iloc[:,1:].apply(pd.to_numeric, errors="coerce")
            
            # Filter out footnote rows - these typically contain text patterns like:
            # "Total sample", "Multiple comparison", "base n =", "Unweighted", etc.
            footnote_patterns = [
                "total sample", "unweighted", "weighted", "base n =", "n =",
                "multiple comparison", "false discovery rate", "fdr", "p =", "p<", "p>",
                "significance", "statistical", "confidence", "margin of error",
                "fieldwork", "survey", "methodology", "data collection"
            ]
            
            # Identify rows to keep (exclude footnotes)
            rows_to_keep = []
            for i, label in enumerate(row_labels_raw):
                label_lower = label.lower().strip()
                
                # Skip empty labels
                if not label_lower:
                    continue
                    
                # Check if this looks like a footnote
                is_footnote = False
                for pattern in footnote_patterns:
                    if pattern in label_lower:
                        is_footnote = True
                        break
                
                
                # Also check if the row has mostly NaN values (typical of footnote rows)
                if not is_footnote:
                    row_data = data_part_raw.iloc[i]
                    non_null_count = row_data.count()
                    total_cols = len(row_data)
                    # If less than 20% of columns have data, likely a footnote
                    if total_cols > 0 and (non_null_count / total_cols) < 0.2:
                        is_footnote = True
                
                if not is_footnote:
                    rows_to_keep.append(i)
            
            # Filter the data to keep only non-footnote rows
            if rows_to_keep:
                row_labels = [row_labels_raw[i] for i in rows_to_keep]
                data_part = data_part_raw.iloc[rows_to_keep].reset_index(drop=True)
            else:
                # Fallback if all rows were filtered out
                row_labels = row_labels_raw
                data_part = data_part_raw
            
            # Replace NaN values with None to avoid chart errors
            data_part = data_part.where(pd.notna(data_part), None)
            # Column label construction
            # Always expose flat col_labels for downstream (optionally combined "Banner | Metric")
            col_banners = [str(b).strip() for b in banner_row[1:len(data_part.columns)+1]]
            col_groups = metric_row_ff[1:len(data_part.columns)+1]
            # Normalize blanks
            col_groups = [g.strip() if isinstance(g, str) and g.strip() != "" else "" for g in col_groups]
            # Combined labels for unique identification (Banner | Metric)
            col_labels = [
                (f"{b} | {g}" if g else str(b))
                for g, b in zip(col_groups, col_banners)
            ]

            # Title guess: first non-empty cell above the header area (metric row if present, else banner row)
            title = None
            top_header_idx = metric_row_idx if metric_row_idx is not None else banner_row_idx
            for r in range(int(top_header_idx)):
                row_vals = sub.iloc[r].dropna().astype(str).tolist()
                if row_vals:
                    title = row_vals[0]
                    break
            if not title:
                title = f"{s} table {bi}"

            # Keep a stable key based on title
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
                    "col_groups": col_groups
                }
            }
            tables.append(tdict)

    return {"tables": tables}

def to_json(data: Dict[str, Any]) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)
