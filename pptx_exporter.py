from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
import json
from typing import Dict, Any, List

# Minimal brand theme
BRAND = {
    "font_family_head": "GT America Extended Regular",
    "font_family_body": "GT America Regular",
    "title_size": 28,
    "axis_size": 11,
    "label_size": 12,
    "bg_color": RGBColor(247, 247, 234),
    "colors": [
        RGBColor(33, 117, 243),
        RGBColor(0, 170, 114),
        RGBColor(247, 148, 30),
        RGBColor(153, 102, 255),
        RGBColor(255, 99, 132),
    ]
}

def _apply_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BRAND["bg_color"]

def add_title(slide, text):
    tx = slide.shapes.title if slide.shapes.title else slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(BRAND["title_size"])
    run.font.bold = True
    run.font.name = BRAND["font_family_head"]

def _chart_type_map(kind: str):
    kind = (kind or "").lower()
    if kind in ["bar_h", "bar horizontal", "horizontal"]:
        return XL_CHART_TYPE.BAR_CLUSTERED
    if kind in ["bar_v", "bar", "column", "vertical"]:
        return XL_CHART_TYPE.COLUMN_CLUSTERED
    if kind in ["donut", "doughnut"]:
        return XL_CHART_TYPE.DOUGHNUT
    if kind in ["line"]:
        return XL_CHART_TYPE.LINE_MARKERS
    return XL_CHART_TYPE.COLUMN_CLUSTERED

def _apply_series_colors(chart):
    for i, series in enumerate(chart.series):
        if i < len(BRAND["colors"]):
            fill = series.format.fill
            fill.solid()
            rgb = BRAND["colors"][i]
            fill.fore_color.rgb = rgb

def _set_alt_text(shape, mapping: dict):
    """Write simple key: value lines into Alt Text Description, easy for humans to edit."""
    try:
        lines = [f"{k}: {v}" for k, v in mapping.items()]
        shape.alternative_text = "\n".join(lines)
    except Exception:
        pass

def _tag_shape(shape, obj_type: str, table_title: str, col_key: str | None = None,
               bind_question: str = "OBJ_QUESTION", bind_base: str = "OBJ_BASE"):
    # Shape name pattern that the updater reads
    if obj_type == "chart":
        shape.name = f"CHART:{table_title}:{col_key or 'Total'}"
        _set_alt_text(shape, {
            "table_key": table_title,
            "col": col_key or "Total",
            "exclude": "base, mean, average, avg",
            "bind_question": bind_question,
            "bind_base": bind_base,
        })
    elif obj_type == "table":
        shape.name = f"TABLE:{table_title}"
        _set_alt_text(shape, {
            "table_key": table_title,
            "col": "*",
            "exclude": "base, mean, average, avg",
        })
    elif obj_type == "text_question":
        shape.name = "OBJ_QUESTION"
    elif obj_type == "text_base":
        shape.name = "OBJ_BASE"

def _add_table(slide, col_labels: List[str], row_labels: List[str], values: List[List[float]]):
    rows = 1 + len(row_labels)
    cols = 1 + len(col_labels)
    left, top, width, height = Inches(0.5), Inches(4.5), Inches(9.0), Inches(3.0)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # Header
    table.cell(0, 0).text = ""
    for j, c in enumerate(col_labels, start=1):
        table.cell(0, j).text = str(c)

    # Body
    for i, rlab in enumerate(row_labels, start=1):
        table.cell(i, 0).text = str(rlab)
        for j, v in enumerate(values[i - 1][:len(col_labels)], start=1):
            table.cell(i, j).text = "" if v is None else f"{v:.1f}"

    # Simple font styling
    for r in range(rows):
        for c in range(cols):
            tf = table.cell(r, c).text_frame
            for p in tf.paragraphs:
                for run in p.runs:
                    run.font.name = BRAND["font_family_body"]
                    run.font.size = Pt(10)

    return table_shape

def add_chart_slide(prs, table: Dict[str, Any], layout=5, chart_kind="bar_h", include_table=False, chart_title=None, base_text: str | None = None):
    slide = prs.slides.add_slide(prs.slide_layouts[layout])
    _apply_background(slide)
    add_title(slide, chart_title or table["title"])

    # Exclusions for charts
    base_row_idx = None
    for i, rlab in enumerate(table["row_labels"]):
        if isinstance(rlab, str) and rlab.strip().lower().startswith("base"):
            base_row_idx = i
            break

    exclude_indices = set()
    if base_row_idx is not None:
        exclude_indices.add(base_row_idx)
    for i, rlab in enumerate(table["row_labels"]):
        if isinstance(rlab, str):
            lab = rlab.strip().lower()
            if lab.startswith("mean") or lab.startswith("average") or lab.startswith("avg"):
                exclude_indices.add(i)

    col_labels = table["col_labels"]
    values = table["values"]
    row_labels = table["row_labels"]

    # Series, prefer Total
    if "Total" in col_labels:
        total_idx = col_labels.index("Total")
    else:
        total_idx = 0 if col_labels else None

    # Chart data arrays
    if total_idx is not None:
        categories = [lab for i, lab in enumerate(row_labels) if i not in exclude_indices]
        series_values = [row[total_idx] if total_idx < len(row) else None for i, row in enumerate(values) if i not in exclude_indices]
    else:
        categories = [lab for i, lab in enumerate(row_labels) if i not in exclude_indices]
        series_values = [0] * len(categories)

    # Add chart
    x, y, w, h = Inches(0.5), Inches(1.2), Inches(9.0), Inches(3.0 if include_table else 6.0)
    chart_type = _chart_type_map(chart_kind)
    chart_data = ChartData()
    chart_data.categories = categories
    chart_data.add_series("Total" if total_idx is not None else "Series", series_values)

    chart_shape = slide.shapes.add_chart(chart_type, x, y, w, h, chart_data)
    chart = chart_shape.chart

    # Formatting
    chart.has_legend = False
    for s in chart.series:
        s.data_labels.show_value = True
        s.data_labels.number_format = "0.0"
        try:
            s.data_labels.position = 2  # end
        except Exception:
            pass
    try:
        gl = chart.value_axis.major_gridlines
        gl.format.line.width = Pt(0.5)
        gl.format.line.fore_color.rgb = RGBColor(210, 210, 210)
    except Exception:
        pass

    chart.category_axis.has_title = False
    chart.value_axis.has_title = False
    chart.category_axis.tick_labels.font.size = Pt(BRAND["axis_size"])
    chart.category_axis.tick_labels.font.name = BRAND["font_family_body"]
    chart.value_axis.tick_labels.font.size = Pt(BRAND["axis_size"])
    chart.value_axis.tick_labels.font.name = BRAND["font_family_body"]

    _apply_series_colors(chart)

    # Tag the chart shape so the updater can find it
    series_name = "Total" if total_idx is not None else "Series"
    _tag_shape(chart_shape, "chart", table.get("title"), series_name)

    # Optional table on the same slide
    if include_table:
        tbl = _add_table(slide, col_labels, row_labels, values)
        _tag_shape(tbl, "table", table.get("title"))

    # Optional Base text
    if base_text:
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(7.0 if include_table else 7.0), Inches(9.0), Inches(0.4))
        _tag_shape(tb, "text_base", table.get("title"))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = base_text
        run.font.size = Pt(10)
        run.font.name = BRAND["font_family_body"]

    # Embed small metadata box, can be ignored by updater
    meta = {
        "table_key": table.get("title"),
        "row_count": len(row_labels),
        "col_labels": col_labels,
        "chart_kind": chart_kind,
        "include_table": include_table,
    }
    meta_json = json.dumps(meta)
    meta_box = slide.shapes.add_textbox(Inches(0.0), Inches(7.45), Inches(0.1), Inches(0.2))
    meta_box.name = "DATA_META"
    tfm = meta_box.text_frame
    tfm.clear()
    r = tfm.paragraphs[0].add_run()
    r.text = meta_json
    r.font.size = Pt(1)

def export_pptx(tables: List[Dict[str, Any]], selections: Dict[str, Dict[str, Any]], out_path: str):
    prs = Presentation()
    # 16:9 slide size
    from pptx.util import Inches as _In
    prs.slide_width = _In(13.333)
    prs.slide_height = _In(7.5)

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    _apply_background(title_slide)
    add_title(title_slide, "Automated Crosstab Report")

    for t in tables:
        sel = selections.get(t["id"], {})
        ctype = sel.get("chart_type", "bar_h")
        include_table = False
        if ctype.lower() in ["chart+table", "chart_with_table", "chart table"]:
            include_table = True
            ctype = "bar_h"
        title = sel.get("title") or t["title"]
        add_chart_slide(
            prs,
            t,
            layout=5,
            chart_kind=ctype,
            include_table=include_table,
            chart_title=title,
            base_text=sel.get("base_text"),
        )

    prs.save(out_path)
    return out_path
