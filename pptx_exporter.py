from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
import json
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

# Minimal brand theme
BRAND = {
    "font_family_head": "Arial",  # Fallback to common fonts
    "font_family_body": "Arial",
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

@dataclass
class TextCallout:
    """Represents a text callout that can be associated with table data."""
    
    # Core identification
    table_title: str   # The table this callout belongs to
    column_key: str    # The column this callout references (e.g., "Total", "Age 18-24")
    row_label: str     # Specific row label this callout targets
    
    # Display properties
    text: Optional[str] = None          # Custom text override (if None, uses callout_type)
    position: tuple = (0.5, 7.0, 9.0, 0.4)  # (x, y, width, height) in inches
    font_size: int = 12
    font_bold: bool = True
    font_color: Optional[RGBColor] = None
    metric_type: str = "percentage"    # percentage | number | currency
    
    # Metadata for mapping
    bind_question: str = "TEXT_QUESTION"
    bind_base: str = "TEXT_BASE"
    auto_update: bool = True
    
    def __post_init__(self):
        """Set default text if not provided."""
        if self.text is None:
            self.text = f"{self.row_label}: [Value]"
    
    def get_display_text(self, table_data: Optional[Dict[str, Any]] = None) -> str:
        """Get the text to display, incorporating data from the table."""
        if table_data:
            # Try to find the actual value for this callout
            row_idx = self._find_row_index(table_data)
            col_idx = self._find_column_index(table_data)
            
            if row_idx is not None and col_idx is not None:
                try:
                    value = table_data["values"][row_idx][col_idx]
                    if value is not None:
                        # Format the value appropriately
                        formatted_value = ""
                        if isinstance(value, (int, float)):
                            if (self.metric_type or "").lower() == "percentage":
                                formatted_value = f"{float(value) * 100:.1f}%"
                            elif (self.metric_type or "").lower() == "currency":
                                formatted_value = f"${float(value):,.0f}"
                            else:
                                # number
                                formatted_value = f"{float(value):,.1f}"
                        else:
                            formatted_value = str(value)
                        
                        # Replace [Value] placeholder with actual value in custom text
                        if self.text and "[Value]" in self.text:
                            return self.text.replace("[Value]", formatted_value)
                        # If user provided custom text without placeholder, respect it
                        if self.text:
                            return self.text
                        # Default
                        return f"{self.row_label}: {formatted_value}"
                except (IndexError, TypeError):
                    pass
        
        # Return custom text if available, otherwise fallback to default
        return self.text if self.text else f"{self.row_label}: [Value]"
    
    def _find_row_index(self, table_data: Dict[str, Any]) -> Optional[int]:
        """Find the row index for this callout."""
        row_labels = table_data.get("row_labels", [])
        for i, label in enumerate(row_labels):
            if isinstance(label, str) and self.row_label.lower() in label.lower():
                return i
        return None
    
    def _find_column_index(self, table_data: Dict[str, Any]) -> Optional[int]:
        """Find the column index for this callout."""
        col_labels = table_data.get("col_labels", [])
        if self.column_key in col_labels:
            return col_labels.index(self.column_key)
        
        # Fallback to common column names
        for fallback in ["Total", "Overall", "All", "Base"]:
            if fallback in col_labels:
                return col_labels.index(fallback)
        
        return 0 if col_labels else None
    
    def to_mapping_dict(self) -> Dict[str, Any]:
        """Convert to mapping dictionary for alt text."""
        return {
            "type": "text_callout",
            "table_title": self.table_title,
            "column": self.column_key,
            "row": self.row_label,
            "metric_type": self.metric_type,
            "auto_update": "yes" if self.auto_update else "no"
        }

def _apply_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BRAND["bg_color"]

def add_title(slide, text, table_title=None):
    """Add a title to the slide with optional alt text for mapping."""
    tx = slide.shapes.title if slide.shapes.title else slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(BRAND["title_size"])
    run.font.bold = True
    run.font.name = BRAND["font_family_head"]
    
    # Add alt text for mapping if table_title is provided
    if table_title:
        _tag_shape(tx, "text_title", table_title, "Total")

def add_question_text(slide, question_text, table_title, position=(0.5, 1.0, 9.0, 0.4)):
    """Add question text with proper alt text for mapping."""
    x, y, w, h = [Inches(pos) for pos in position]
    qb = slide.shapes.add_textbox(x, y, w, h)
    
    # Set alt text for mapping using _tag_shape
    _tag_shape(qb, "text_question", table_title, "Total")
    
    # Set the text content
    tf = qb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f"Question: {question_text}"
    run.font.size = Pt(12)
    run.font.name = BRAND["font_family_body"]
    
    return qb

def add_text_callout(slide, callout: TextCallout, table_data: Optional[Dict[str, Any]] = None):
    """Add a text callout to the slide with proper alt text for mapping."""
    x, y, w, h = [Inches(pos) for pos in callout.position]
    callout_box = slide.shapes.add_textbox(x, y, w, h)
    
    # Set alt text for mapping using _tag_shape
    _tag_shape(callout_box, "text_callout", callout.table_title, callout.column_key, 
               callout.bind_question, callout.bind_base, callout.row_label, callout.metric_type)
    
    # Set the text content
    tf = callout_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    
    # Get the display text (with data if available)
    display_text = callout.get_display_text(table_data)
    run.text = display_text
    
    # Apply formatting
    run.font.size = Pt(callout.font_size)
    run.font.bold = callout.font_bold
    run.font.name = BRAND["font_family_body"]
    
    if callout.font_color:
        run.font.color.rgb = callout.font_color
    
    return callout_box

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
    """Write simple key: value lines into Alt Text Description using the correct XML structure."""
    try:
        # Create human-readable format that's easy to edit in PowerPoint
        lines = []
        for k, v in mapping.items():
            if v is not None and v != "":
                lines.append(f"{k}: {v}")
        
        alt_text_content = "\n".join(lines)
        
        # Method 1: Try to set via alternative_text property (if it exists)
        try:
            if hasattr(shape, 'alternative_text'):
                shape.alternative_text = alt_text_content
                return
        except Exception:
            pass
        
        # Method 2: Set via XML using the correct cNvPr structure
        try:
            if hasattr(shape, 'element'):
                # Look for the cNvPr element which contains name and description
                c_nv_pr = None
                
                # For GraphicFrame (charts/tables)
                if 'graphicFrame' in shape.element.tag:
                    c_nv_pr = shape.element.find('.//p:cNvPr', namespaces={'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
                # For Shape (text boxes, etc.)
                elif 'sp' in shape.element.tag:
                    c_nv_pr = shape.element.find('.//p:cNvPr', namespaces={'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
                
                if c_nv_pr is not None:
                    # Set the description attribute
                    c_nv_pr.set('descr', alt_text_content)
                    return
                    
        except Exception:
            pass
        
        # Method 3: Try to set via shape name as a fallback
        try:
            # Use a special naming convention that includes the mapping info
            safe_mapping = alt_text_content.replace('\n', ' | ').replace(':', '_')[:100]
            shape.name = f"MAPPED_{safe_mapping}"
        except Exception:
            pass
            
    except Exception:
        pass

def _tag_shape(shape, obj_type: str, table_title: str, col_key: str | None = None,
               bind_question: str = "TEXT_QUESTION", bind_base: str = "TEXT_BASE", row_label: str | None = None, metric_type: str | None = None):
    """Tag shapes with alt text only - no shape naming."""
    
    if obj_type == "chart":
        _set_alt_text(shape, {
            "type": "chart",
            "table_title": table_title,
            "column": col_key or "Total",
            "exclude_rows": "base, mean, average, avg",
            "auto_update": "yes"
        })
        
    elif obj_type == "table":
        _set_alt_text(shape, {
            "type": "table", 
            "table_title": table_title,
            "columns": "*",
            "column": col_key or "Total",  # Add column attribute
            "exclude_rows": "base, mean, average, avg",
            "auto_update": "yes"
        })
        
    elif obj_type in ["question_text", "text_question"]:
        _set_alt_text(shape, {
            "type": "text_question",
            "table_title": table_title,
            "column": col_key or "Total",  # Add column attribute
            "auto_update": "yes"
        })
        
    elif obj_type == "text_base":
        _set_alt_text(shape, {
            "type": "text_base",
            "table_title": table_title,
            "column": col_key or "Total",  # Add column attribute
            "auto_update": "yes"
        })
        
    elif obj_type == "text_title":
        _set_alt_text(shape, {
            "type": "text_title",
            "table_title": table_title,
            "column": col_key or "Total",  # Add column attribute
            "auto_update": "yes"
        })
        
    elif obj_type == "text_callout":
        _set_alt_text(shape, {
            "type": "text_callout",
            "table_title": table_title,
            "column": col_key or "Total",
            "row": row_label or "",
            "metric_type": (metric_type or "percentage"),
            "auto_update": "yes"
        })

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

def sort_table_rows(table: Dict[str, Any], column_key: str = "Total", excluded_rows: List[str] = None) -> Dict[str, Any]:
    """Sort table rows by values in descending order, keeping excluded rows at bottom."""
    if excluded_rows is None:
        excluded_rows = []
    
    # Create a copy to avoid modifying the original
    sorted_table = table.copy()
    row_labels = table["row_labels"].copy()
    values = [row.copy() for row in table["values"]]
    
    # Find the column index to sort by
    col_labels = table["col_labels"]
    if column_key in col_labels:
        sort_col_idx = col_labels.index(column_key)
    elif "Total" in col_labels:
        sort_col_idx = col_labels.index("Total")
    else:
        sort_col_idx = 0 if col_labels else None
    
    if sort_col_idx is None:
        return sorted_table  # Can't sort without a valid column
    
    # Separate rows into sortable and excluded
    sortable_data = []
    excluded_data = []
    
    for i, (label, row) in enumerate(zip(row_labels, values)):
        if label in excluded_rows:
            excluded_data.append((label, row, i))
        else:
            # Get the value to sort by
            sort_value = row[sort_col_idx] if sort_col_idx < len(row) and row[sort_col_idx] is not None else 0
            sortable_data.append((label, row, i, sort_value))
    
    # Sort the sortable data by value (descending)
    sortable_data.sort(key=lambda x: x[3], reverse=True)
    
    # Rebuild the table with sorted rows
    new_row_labels = []
    new_values = []
    
    # Add sorted rows first
    for label, row, _, _ in sortable_data:
        new_row_labels.append(label)
        new_values.append(row)
    
    # Add excluded rows at the bottom
    for label, row, _ in excluded_data:
        new_row_labels.append(label)
        new_values.append(row)
    
    # Update the sorted table
    sorted_table["row_labels"] = new_row_labels
    sorted_table["values"] = new_values
    
    return sorted_table

def add_chart_slide(prs, table: Dict[str, Any], layout=5, chart_kind="bar_h", include_table=False, chart_title=None, base_text: str | None = None, question_text: str | None = None, callouts: List[TextCallout] | None = None, enable_sorting: bool = False, excluded_rows: List[str] = None, column_key: str | None = None):
    slide = prs.slides.add_slide(prs.slide_layouts[layout])
    _apply_background(slide)
    
    # Apply row sorting if enabled
    working_table = table
    if enable_sorting:
        # Determine which column to sort by
        sort_column = "Total"  # Default sort column
        if callouts and len(callouts) > 0:
            # Use the column from the first callout if available
            sort_column = callouts[0].column_key
        working_table = sort_table_rows(table, sort_column, excluded_rows or [])
    
    # Add title with alt text for mapping
    title_to_use = chart_title or working_table["title"]
    add_title(slide, title_to_use, working_table.get("title"))
    
    # Add question text if provided
    if question_text:
        add_question_text(slide, question_text, working_table.get("title"), position=(0.5, 0.8, 9.0, 0.4))
        chart_y_offset = 1.4  # Move chart down if question text is present
    else:
        chart_y_offset = 1.2  # Default chart position

    # Exclusions for charts
    base_row_idx = None
    for i, rlab in enumerate(working_table["row_labels"]):
        if isinstance(rlab, str) and rlab.strip().lower().startswith("base"):
            base_row_idx = i
            break

    exclude_indices = set()
    if base_row_idx is not None:
        exclude_indices.add(base_row_idx)
    for i, rlab in enumerate(working_table["row_labels"]):
        if isinstance(rlab, str):
            lab = rlab.strip().lower()
            if lab.startswith("mean") or lab.startswith("average") or lab.startswith("avg"):
                exclude_indices.add(i)

    col_labels = working_table["col_labels"]
    values = working_table["values"]
    row_labels = working_table["row_labels"]

    # Series to plot: use explicit column_key if provided, else prefer Total
    selected_idx = None
    if column_key and column_key in col_labels:
        selected_idx = col_labels.index(column_key)
    elif "Total" in col_labels:
        selected_idx = col_labels.index("Total")
    else:
        selected_idx = 0 if col_labels else None

    # Chart data arrays
    if selected_idx is not None:
        categories = [lab for i, lab in enumerate(row_labels) if i not in exclude_indices]
        series_values = [row[selected_idx] if selected_idx < len(row) else None for i, row in enumerate(values) if i not in exclude_indices]
    else:
        categories = [lab for i, lab in enumerate(row_labels) if i not in exclude_indices]
        series_values = [0] * len(categories)

    # Add chart
    x, y, w, h = Inches(0.5), Inches(chart_y_offset), Inches(9.0), Inches(3.0 if include_table else 6.0)
    chart_type = _chart_type_map(chart_kind)
    chart_data = ChartData()
    chart_data.categories = categories
    series_name = column_key if (column_key and column_key in col_labels) else ("Total" if "Total" in col_labels else "Series")
    chart_data.add_series(series_name, series_values)

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
    _tag_shape(chart_shape, "chart", working_table.get("title"), series_name)

    # Optional table on the same slide
    if include_table:
        tbl = _add_table(slide, col_labels, row_labels, values)
        _tag_shape(tbl, "table", working_table.get("title"))

    # Optional Base text
    if base_text:
        base_y = 7.0 if include_table else 7.0
        if question_text:
            base_y += 0.4  # Move base text down if question text is present
            
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(base_y), Inches(9.0), Inches(0.4))
        _tag_shape(tb, "text_base", working_table.get("title"))
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = base_text
        run.font.size = Pt(10)
        run.font.name = BRAND["font_family_body"]

    # Add text callouts if provided
    if callouts:
        # Calculate starting position for callouts
        callout_start_y = 7.0
        if base_text:
            callout_start_y += 0.5  # Move callouts down if base text is present
        if question_text:
            callout_start_y += 0.4  # Move callouts down if question text is present
        
        # Position callouts vertically, adjusting for multiple callouts
        for i, callout in enumerate(callouts):
            # Adjust position for this specific callout
            callout_position = list(callout.position)
            callout_position[1] = callout_start_y + (i * 0.5)  # Stack vertically with 0.5" spacing
            
            # Create a copy of the callout with adjusted position
            adjusted_callout = TextCallout(
                table_title=callout.table_title,
                column_key=callout.column_key,
                row_label=callout.row_label,
                text=callout.text,
                position=tuple(callout_position),
                font_size=callout.font_size,
                font_bold=callout.font_bold,
                font_color=callout.font_color,
                bind_question=callout.bind_question,
                bind_base=callout.bind_base,
                auto_update=callout.auto_update
            )
            
            add_text_callout(slide, adjusted_callout, working_table)

    # Embed small metadata box, can be ignored by updater
    meta = {
        "table_key": working_table.get("title"),
        "row_count": len(row_labels),
        "col_labels": col_labels,
        "chart_kind": chart_kind,
        "include_table": include_table,
        "callout_count": len(callouts) if callouts else 0,
        "enable_sorting": enable_sorting,
        "excluded_rows": excluded_rows or []
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
        
        # Process callouts if specified in selections
        callouts = None
        if "callouts" in sel:
            callouts = []
            for callout_data in sel["callouts"]:
                # Create TextCallout object from selection data
                callout = TextCallout(
                    table_title=t.get("title", ""),
                    column_key=callout_data.get("column_key", "Total"),
                    row_label=callout_data.get("row_label"),
                    text=callout_data.get("text"),
                    position=callout_data.get("position", (0.5, 7.0, 9.0, 0.4)),
                    font_size=callout_data.get("font_size", 12),
                    font_bold=callout_data.get("font_bold", True),
                    font_color=callout_data.get("font_color"),
                    bind_question=callout_data.get("bind_question", "TEXT_QUESTION"),
                    bind_base=callout_data.get("bind_base", "TEXT_BASE"),
                    metric_type=callout_data.get("metric_type", "percentage"),
                    auto_update=callout_data.get("auto_update", True)
                )
                callouts.append(callout)
        
        add_chart_slide(
            prs,
            t,
            layout=5,
            chart_kind=ctype,
            include_table=include_table,
            chart_title=title,
            base_text=sel.get("base_text"),
            question_text=sel.get("question_text"),
            callouts=callouts,
            enable_sorting=sel.get("enable_sorting", False),
            excluded_rows=sel.get("excluded_rows", []),
            column_key=sel.get("column_key")
        )

    prs.save(out_path)
    return out_path

def create_row_callout(table_title: str, row_label: str, column_key: str = "Total", 
                      custom_text: str = None, position: tuple = (0.5, 7.0, 9.0, 0.4)) -> TextCallout:
    """Create a callout for a specific row in a table."""
    return TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label=row_label,
        text=custom_text,
        position=position,
        font_size=12,
        font_bold=True
    )

def create_common_callouts(table_title: str, column_key: str = "Total") -> List[TextCallout]:
    """Create a list of common callout types for a table."""
    callouts = []
    
    # Top Box callouts
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Top 3 Box",
        text="Top 3 Box",
        position=(0.5, 7.0, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Top 2 Box", 
        text="Top 2 Box",
        position=(5.0, 7.0, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    # Bottom Box callouts
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Bottom 2 Box",
        text="Bottom 2 Box",
        position=(0.5, 7.5, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Bottom 3 Box",
        text="Bottom 3 Box",
        position=(5.0, 7.5, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    return callouts

def create_statistical_callouts(table_title: str, column_key: str = "Total") -> List[TextCallout]:
    """Create statistical callouts like Average, Mean, etc."""
    callouts = []
    
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Average",
        text="Average",
        position=(0.5, 7.0, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Mean",
        text="Mean",
        position=(5.0, 7.0, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    callouts.append(TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label="Total Spend",
        text="Total Spend",
        position=(0.5, 7.5, 4.0, 0.4),
        font_size=12,
        font_bold=True
    ))
    
    return callouts

def create_custom_callout(table_title: str, column_key: str = "Total", 
                         row_label: str = None, custom_text: str = None, 
                         position: tuple = (0.5, 7.0, 9.0, 0.4)) -> TextCallout:
    """Create a custom callout with specific parameters."""
    return TextCallout(
        table_title=table_title,
        column_key=column_key,
        row_label=row_label or "Custom",
        text=custom_text,
        position=position,
        font_size=12,
        font_bold=True
    )
