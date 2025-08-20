from pptx import Presentation
from pptx.chart.data import ChartData
from typing import Dict, Any, List, Optional
import re, json

from crosstab_parser import parse_workbook

EXCLUDE_PREFIXES = ("base", "mean", "average", "avg")

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()

def _parse_alt_text(shape) -> Dict[str, str]:
    out: Dict[str, str] = {}
    try:
        alt = shape.alternative_text or ""
    except Exception:
        alt = ""
    for line in alt.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            out[_norm(k)] = v.strip()
    return out

def _exclude_indices(labels: List[str]) -> set:
    ex = set()
    for i, lab in enumerate(labels):
        if isinstance(lab, str) and _norm(lab).startswith(EXCLUDE_PREFIXES):
            ex.add(i)
    return ex

def _row_index_map(labels: List[str]) -> Dict[str, int]:
    return {_norm(l): i for i, l in enumerate(labels)}

def _choose_col_idx(col_labels: List[str], col_key: Optional[str]) -> Optional[int]:
    if not col_labels:
        return None
    if col_key and col_key in col_labels:
        return col_labels.index(col_key)
    for cand in ["Total", "Overall", "All", "Base"]:
        if cand in col_labels:
            return col_labels.index(cand)
    return 0

def _series_from_table(table: Dict[str, Any], col_idx: Optional[int], exclude_rows: set):
    cats, vals = [], []
    row_labels = table["row_labels"]
    values = table["values"]
    for i, lab in enumerate(row_labels):
        if i in exclude_rows:
            continue
        cats.append(lab)
        if col_idx is None:
            vals.append(None)
        else:
            row = values[i]
            vals.append(row[col_idx] if col_idx < len(row) else None)
    return cats, vals

def _update_chart(shape, table: Dict[str, Any], col_key: Optional[str], explicit_rows: Optional[List[str]]):
    chart = shape.chart
    col_idx = _choose_col_idx(table["col_labels"], col_key)
    ex = _exclude_indices(table["row_labels"])

    if explicit_rows:
        idx_map = _row_index_map(table["row_labels"])
        cats, vals = [], []
        for lab in explicit_rows:
            j = idx_map.get(_norm(lab))
            if j is None or j in ex:
                continue
            cats.append(lab)
            if col_idx is None:
                vals.append(None)
            else:
                row = table["values"][j]
                vals.append(row[col_idx] if col_idx < len(row) else None)
    else:
        cats, vals = _series_from_table(table, col_idx, ex)

    cd = ChartData()
    cd.categories = cats
    cd.add_series(col_key if col_idx is not None else "Series", vals)
    chart.replace_data(cd)

def _update_table(shape, table: Dict[str, Any]):
    if not shape.has_table:
        return
    tbl = shape.table
    hdrs = []
    for c in range(1, len(tbl.columns)):
        txt = tbl.cell(0, c).text_frame.text.strip()
        hdrs.append(txt)
    col_labels = table["col_labels"]
    col_map = [col_labels.index(h) if h in col_labels else None for h in hdrs]
    idx_map = _row_index_map(table["row_labels"])

    for r in range(1, len(tbl.rows)):
        rlab = tbl.cell(r, 0).text_frame.text.strip()
        j = idx_map.get(_norm(rlab))
        for c in range(1, len(tbl.columns)):
            txt = ""
            ci = col_map[c - 1]
            if j is not None and ci is not None:
                try:
                    val = table["values"][j][ci]
                    txt = "" if val is None else f"{float(val):.1f}"
                except Exception:
                    txt = ""
            tbl.cell(r, c).text_frame.text = txt

def _find_shape(slide, name: str):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None

def _update_question_and_base(slide, table: Dict[str, Any], q_name: Optional[str], b_name: Optional[str]):
    if q_name:
        q = _find_shape(slide, q_name)
        if q is not None and hasattr(q, "text_frame"):
            q.text_frame.text = f"Question: {table.get('title','')}"
    if b_name:
        base_n = None
        row_labels = table["row_labels"]
        values = table["values"]
        col_labels = table["col_labels"]
        base_idx = None
        for i, lab in enumerate(row_labels):
            if _norm(lab).startswith("base"):
                base_idx = i
                break
        if base_idx is not None:
            ci = _choose_col_idx(col_labels, "Total")
            if ci is not None and base_idx < len(values) and ci < len(values[base_idx]):
                try:
                    base_n = int(round(float(values[base_idx][ci])))
                except Exception:
                    base_n = None
        b = _find_shape(slide, b_name)
        if b is not None and hasattr(b, "text_frame"):
            b.text_frame.text = (
                f"Base: Total respondents. {base_n} complete surveys."
                if base_n is not None else
                "Base: Total respondents."
            )

def update_presentation(pptx_in: str, crosstab_xlsx: str, pptx_out: str) -> str:
    prs = Presentation(pptx_in)
    data = parse_workbook(crosstab_xlsx)
    tbl_by_norm = {_norm(t.get("title")): t for t in data["tables"]}

    for slide in prs.slides:
        for shp in slide.shapes:
            name = shp.name or ""
            alt = _parse_alt_text(shp)

            # tags from name
            table_key = None
            col = None
            if name.startswith("CHART:"):
                parts = name.split(":", 2)
                if len(parts) >= 2:
                    table_key = parts[1].strip()
                if len(parts) == 3:
                    col = parts[2].strip()
            elif name.startswith("TABLE:"):
                parts = name.split(":", 1)
                if len(parts) == 2:
                    table_key = parts[1].strip()

            # overrides from Alt Text
            if not table_key and "table_key" in alt:
                table_key = alt["table_key"]
            if not col and "col" in alt:
                col = alt["col"]
            q_bind = alt.get("bind_question")
            b_bind = alt.get("bind_base")

            if not table_key:
                continue

            t = tbl_by_norm.get(_norm(table_key))
            if not t:
                t = next((x for x in data["tables"] if x.get("title") == table_key), None)
            if not t:
                continue

            if hasattr(shp, "chart"):
                _update_chart(shp, t, col, explicit_rows=None)
                _update_question_and_base(slide, t, q_bind, b_bind)
            elif shp.has_table:
                _update_table(shp, t)
                _update_question_and_base(slide, t, q_bind, b_bind)

    prs.save(pptx_out)
    return pptx_out
