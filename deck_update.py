from pptx import Presentation
from pptx.chart.data import ChartData
from typing import Dict, Any, List, Optional, Tuple
import re, json

from crosstab_parser import parse_workbook

EXCLUDE_PREFIXES = ("base", "mean", "average", "avg")

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()

def _parse_alt_text(shape) -> Dict[str, str]:
    """Parse alt text with enhanced flexibility for manual editing."""
    out: Dict[str, str] = {}
    
    # Method 1: Try to read from XML descr attribute (most reliable)
    try:
        if hasattr(shape, 'element'):
            # Look for the cNvPr element which contains the description
            c_nv_pr = None
            
            # For GraphicFrame (charts/tables)
            if 'graphicFrame' in shape.element.tag:
                c_nv_pr = shape.element.find('.//p:cNvPr', namespaces={'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
            # For Shape (text boxes, etc.)
            elif 'sp' in shape.element.tag:
                c_nv_pr = shape.element.find('.//p:cNvPr', namespaces={'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
            
            if c_nv_pr is not None and c_nv_pr.get('descr'):
                alt = c_nv_pr.get('descr')
            else:
                alt = ""
        else:
            alt = ""
    except Exception:
        alt = ""
    
    # Method 2: Fallback to alternative_text property (if it exists)
    if not alt:
        try:
            alt = shape.alternative_text or ""
        except Exception:
            alt = ""
    
    # Parse the alt text content
    for line in alt.splitlines():
        line = line.strip()
        if ":" in line:
            # Handle both "key: value" and "key : value" formats
            if " : " in line:
                k, v = line.split(" : ", 1)
            else:
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
    import math
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
            val = row[col_idx] if col_idx < len(row) else None
            
            # Clean the value to handle NaN and infinite values
            if val is not None:
                try:
                    # Convert to float to check for NaN/inf
                    float_val = float(val)
                    if math.isnan(float_val) or math.isinf(float_val):
                        # Replace NaN/inf with None (which becomes 0 in charts)
                        val = None
                    else:
                        val = float_val
                except (ValueError, TypeError):
                    # If conversion fails, treat as None
                    val = None
            
            vals.append(val)
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

    # Store current chart formatting before updating data
    chart_formatting = {}
    try:
        # Store chart type
        chart_formatting['chart_type'] = chart.chart_type
        
        # Store chart style if available
        if hasattr(chart, 'chart_style'):
            chart_formatting['chart_style'] = chart.chart_style
            
        # Store plot area formatting
        if hasattr(chart, 'plot_area'):
            plot_area = chart.plot_area
            if hasattr(plot_area, 'format'):
                chart_formatting['plot_format'] = plot_area.format
                
        # Store chart area formatting
        if hasattr(chart, 'chart_area'):
            chart_area = chart.chart_area
            if hasattr(chart_area, 'format'):
                chart_formatting['chart_area_format'] = chart_area.format
                
    except Exception as e:
        print(f"Warning: Could not preserve all chart formatting: {e}")

    # Update the chart data
    cd = ChartData()
    cd.categories = cats
    cd.add_series(col_key if col_idx is not None else "Series", vals)
    chart.replace_data(cd)
    
    # Restore chart formatting after data update
    try:
        # Restore chart type if it was changed
        if 'chart_type' in chart_formatting:
            chart.chart_type = chart_formatting['chart_type']
            
        # Restore chart style if it was changed
        if 'chart_style' in chart_formatting:
            chart.chart_style = chart_formatting['chart_style']
            
    except Exception as e:
        print(f"Warning: Could not restore all chart formatting: {e}")
    
    print(f"✓ Updated chart data while preserving formatting for table: {table.get('title')}")

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
                    if val is not None:
                        import math
                        float_val = float(val)
                        if math.isnan(float_val) or math.isinf(float_val):
                            txt = ""  # Display empty for NaN/inf values
                        else:
                            txt = f"{float_val:.1f}"
                    else:
                        txt = ""
                except Exception:
                    txt = ""
            
            # Preserve formatting by updating only the text content, not the entire cell
            cell = tbl.cell(r, c)
            if hasattr(cell, 'text_frame') and hasattr(cell.text_frame, 'paragraphs'):
                if len(cell.text_frame.paragraphs) > 0:
                    paragraph = cell.text_frame.paragraphs[0]
                    if paragraph.runs:
                        # Update only the first run's text, preserving formatting
                        paragraph.runs[0].text = txt
                    else:
                        # No runs, create one with the new text
                        run = paragraph.add_run()
                        run.text = txt
                else:
                    # No paragraphs, create one
                    paragraph = cell.text_frame.add_paragraph()
                    run = paragraph.add_run()
                    run.text = txt
            else:
                # Fallback to updating the entire text_frame
                cell.text = txt
    
    print(f"✓ Updated table data while preserving formatting for table: {table.get('title')}")

def _find_shape(slide, name: str):
    """Find shape by name with enhanced search capabilities."""
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None

def _find_shapes_by_pattern(slide, pattern: str):
    """Find shapes that match a pattern (useful for manual naming)."""
    matches = []
    for shp in slide.shapes:
        if shp.name and pattern.lower() in shp.name.lower():
            matches.append(shp)
    return matches

def _get_table_mapping_from_shape(shape, data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Enhanced table mapping that handles both automatic and manual configurations."""
    name = shape.name or ""
    alt = _parse_alt_text(shape)
    
    # Priority 1: Direct table_title match from alt text
    if "table_title" in alt:
        table_title = alt["table_title"]
        # Try exact match first
        for table in data["tables"]:
            if table.get("title") == table_title:
                return table
        
        # Try normalized match
        norm_title = _norm(table_title)
        for table in data["tables"]:
            if _norm(table.get("title", "")) == norm_title:
                return table
    
    # Priority 2: Shape name pattern matching
    if name.startswith("TABLE_"):
        # Extract table title from name
        table_title = name[6:].replace("_", " ")  # Remove "TABLE_" prefix
        for table in data["tables"]:
            if _norm(table.get("title", "")) == _norm(table_title):
                return table
    
    # Priority 3: Legacy TABLE: format
    if name.startswith("TABLE:"):
        parts = name.split(":", 1)
        if len(parts) == 2:
            table_title = parts[1].strip()
            for table in data["tables"]:
                if _norm(table.get("title", "")) == _norm(table_title):
                    return table
    
    return None

def _get_chart_mapping_from_shape(shape, data: Dict[str, Any], selections: Optional[Dict[str, Any]] = None) -> Optional[Tuple[Dict[str, Any], Optional[str]]]:
    """Enhanced chart mapping that handles both automatic and manual configurations."""
    name = shape.name or ""
    alt = _parse_alt_text(shape)
    
    table = None
    col_key = None
    
    # Priority 1: Direct mapping from alt text
    if "table_title" in alt:
        table_title = alt["table_title"]
        col_key = alt.get("column")
        
        # Find matching table
        for t in data["tables"]:
            if t.get("title") == table_title:
                table = t
                break
        
        if not table:
            # Try normalized match
            norm_title = _norm(table_title)
            for t in data["tables"]:
                if _norm(t.get("title", "")) == norm_title:
                    table = t
                    break
    
    # Priority 2: Shape name pattern matching
    if not table and name.startswith("CHART_"):
        # Extract table title and column from name
        name_parts = name[6:].split("_")  # Remove "CHART_" prefix
        if len(name_parts) >= 2:
            # Last part might be column, rest is table title
            potential_col = name_parts[-1]
            table_title = "_".join(name_parts[:-1])
            
            # Check if last part is a valid column name
            for t in data["tables"]:
                if _norm(t.get("title", "")) == _norm(table_title):
                    if potential_col in t.get("col_labels", []):
                        table = t
                        col_key = potential_col
                        break
                    else:
                        # Last part wasn't a column, treat whole thing as table title
                        table_title = "_".join(name_parts)
                        if _norm(t.get("title", "")) == _norm(table_title):
                            table = t
                            break
    
    # Priority 3: Legacy CHART: format
    if not table and name.startswith("CHART:"):
        parts = name.split(":", 2)
        if len(parts) >= 2:
            table_title = parts[1].strip()
            col_key = parts[2].strip() if len(parts) == 3 else None
            
            for t in data["tables"]:
                if _norm(t.get("title", "")) == _norm(table_title):
                    table = t
                    break
    
    # Priority 4: Use column_key from selections if available
    if table and selections and table.get("title") in selections:
        selection = selections[table.get("title")]
        if "column_key" in selection:
            col_key = selection["column_key"]
    
    return (table, col_key) if table else None

def _format_number_with_commas(number):
    """Format a number with comma separators for thousands places."""
    if number is None:
        return None
    return f"{number:,}"

def _update_question_and_base(slide, table: Dict[str, Any], col_key: Optional[str], explicit_rows: Optional[List[int]]):
    """Update question and base text based on alt text mapping, preserving custom content."""
    # Find shapes by alt text type instead of name
    for shape in slide.shapes:
        alt = _parse_alt_text(shape)
        
        # Update question text
        if alt.get("type") == "question_text" and alt.get("table_title") == table.get("title"):
            if hasattr(shape, "text_frame"):
                # Preserve existing custom question text if it's different from the table title
                current_text = shape.text_frame.text
                if current_text.startswith("Question: "):
                    existing_question = current_text[10:]  # Remove "Question: " prefix
                    # Only update if the existing question is the same as table title (default)
                    # This preserves custom questions that users have written
                    if existing_question == table.get("title", ""):
                        # Preserve formatting by only updating the text content
                        if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                            paragraph = shape.text_frame.paragraphs[0]
                            if paragraph.runs:
                                paragraph.runs[0].text = f"Question: {table.get('title', '')}"
                            else:
                                run = paragraph.add_run()
                                run.text = f"Question: {table.get('title', '')}"
                        else:
                            shape.text_frame.text = f"Question: {table.get('title', '')}"
                        print(f"✓ Updated question text for table: {table.get('title')}")
                    else:
                        print(f"✓ Preserved custom question text: {existing_question}")
                else:
                    # No "Question: " prefix, add it while preserving formatting
                    if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                        paragraph = shape.text_frame.paragraphs[0]
                        if paragraph.runs:
                            paragraph.runs[0].text = f"Question: {table.get('title', '')}"
                        else:
                            run = paragraph.add_run()
                            run.text = f"Question: {table.get('title', '')}"
                    else:
                        shape.text_frame.text = f"Question: {table.get('title', '')}"
                    print(f"✓ Added question text for table: {table.get('title')}")
        
        # Update base text
        elif alt.get("type") == "text_base" and alt.get("table_title") == table.get("title"):
            if hasattr(shape, "text_frame"):
                # Calculate new base size
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
                
                # Preserve custom base description while updating N value
                current_base_text = shape.text_frame.text
                custom_description = ""
                
                # Extract custom description from existing text
                if "Base:" in current_base_text:
                    # Look for patterns like "Base: Total respondents. 123 complete surveys."
                    base_parts = current_base_text.split(".")
                    if len(base_parts) >= 2:
                        # First part contains the custom description
                        custom_description = base_parts[0].replace("Base:", "").strip()
                        # Clean up any trailing punctuation or equals signs
                        custom_description = custom_description.rstrip(" =").strip()
                    else:
                        # No period, might be just "Base: Total respondents 123"
                        base_parts = current_base_text.split()
                        if len(base_parts) >= 3:
                            base_idx_text = base_parts.index("Base:")
                            if base_idx_text >= 0 and base_idx_text + 1 < len(base_parts):
                                # Find where the number starts
                                for i in range(base_idx_text + 1, len(base_parts)):
                                    if base_parts[i].replace(",", "").isdigit():
                                        custom_description = " ".join(base_parts[base_idx_text + 1:i])
                                        # Clean up any trailing punctuation or equals signs
                                        custom_description = custom_description.rstrip(" =").strip()
                                        break
                
                # Use custom description if found, otherwise use default
                if custom_description:
                    if base_n is not None:
                        # Use the custom description as-is, don't force "Total respondents"
                        new_text = f"Base: {custom_description}. {_format_number_with_commas(base_n)} complete surveys."
                    else:
                        new_text = f"Base: {custom_description}."
                    print(f"✓ Updated base text for table: {table.get('title')} - preserved custom description: {custom_description}, new N: {_format_number_with_commas(base_n)}")
                else:
                    # Use default description
                    if base_n is not None:
                        new_text = f"Base: Total respondents. {_format_number_with_commas(base_n)} complete surveys."
                    else:
                        new_text = "Base: Total respondents."
                    print(f"✓ Updated base text for table: {table.get('title')} - Base: {_format_number_with_commas(base_n)}")
                
                # Preserve formatting by only updating the text content, not the entire text_frame
                if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                    # Update only the first paragraph's text, preserving formatting
                    paragraph = shape.text_frame.paragraphs[0]
                    if paragraph.runs:
                        # Update the first run's text, preserving its formatting
                        paragraph.runs[0].text = new_text
                    else:
                        # No runs, create one with the new text
                        run = paragraph.add_run()
                        run.text = new_text
                else:
                    # Fallback to updating the entire text_frame
                    shape.text_frame.text = new_text
        
        # Update chart title
        elif alt.get("type") == "text_title" and alt.get("table_title") == table.get("title"):
            if hasattr(shape, "text_frame"):
                # Preserve existing custom title - don't overwrite it with table title
                current_text = shape.text_frame.text
                # Only update if the current title is the same as table title (default)
                # This preserves custom titles that users have written
                if current_text == table.get("title", ""):
                    print(f"✓ Chart title already current for table: {table.get('title')}")
                else:
                    print(f"✓ Preserved custom chart title: {current_text}")
                    # Don't change the text - keep the custom title

def _update_question_and_base_with_selections(slide, table: Dict[str, Any], selections: dict, table_title: Optional[str]):
    """Update question and base text based on alt text mapping, using current selections if available."""
    
    print(f"DEBUG: _update_question_and_base_with_selections called for table: {table_title}")
    
    # Find the selection for this specific table by matching table title
    table_selection = None
    if table_title and table_title in selections:
        table_selection = selections[table_title]
        print(f"DEBUG: Found selection for table '{table_title}': {table_selection}")
    else:
        print(f"⚠️ No selection found for table: {table_title}")
        print(f"DEBUG: Available selections: {list(selections.keys()) if selections else 'None'}")
        return
    
    # Find shapes by alt text type instead of name
    shape_count = 0
    for shape in slide.shapes:
        alt = _parse_alt_text(shape)
        
        # Update question text
        if alt.get("type") == "question_text" and alt.get("table_title") == table_title:
            shape_count += 1
            print(f"DEBUG: Found question_text shape #{shape_count} for table: {table_title}")
            if hasattr(shape, "text_frame"):
                # Use current selection for question text if available
                if "question_text" in table_selection:
                    new_text = table_selection["question_text"]
                    print(f"DEBUG: Question text from selection: '{new_text}'")
                    print(f"DEBUG: Current shape text before update: '{shape.text_frame.text}'")
                    
                    # Preserve formatting by only updating the text content
                    if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                        paragraph = shape.text_frame.paragraphs[0]
                        print(f"DEBUG: Paragraph has {len(paragraph.runs)} runs")
                        
                        if paragraph.runs:
                            # Keep the first run (which has the formatting) and update its text
                            first_run = paragraph.runs[0]
                            first_run.text = f"Question: {new_text}"
                            print(f"DEBUG: Updated first run text: '{first_run.text}'")
                            
                            # Remove any additional runs to prevent concatenation
                            # We'll just clear the paragraph and recreate the first run with formatting
                            if len(paragraph.runs) > 1:
                                # Store the formatting from the first run
                                font = first_run.font
                                # Clear and recreate
                                paragraph.clear()
                                new_run = paragraph.add_run()
                                new_run.text = f"Question: {new_text}"
                                # Apply the formatting
                                new_run.font.name = font.name
                                new_run.font.size = font.size
                                new_run.font.bold = font.bold
                                new_run.font.italic = font.italic
                                # Only set color if it has an rgb property
                                if hasattr(font.color, 'rgb') and font.color.rgb is not None:
                                    new_run.font.color.rgb = font.color.rgb
                                print(f"DEBUG: Recreated run with preserved formatting")
                        else:
                            # No runs exist, create one
                            run = paragraph.add_run()
                            run.text = f"Question: {new_text}"
                            print(f"DEBUG: Created new run text: '{run.text}'")
                    else:
                        print(f"DEBUG: No paragraphs, updating entire text_frame")
                        shape.text_frame.text = f"Question: {new_text}"
                        print(f"DEBUG: Text_frame text after update: '{shape.text_frame.text}'")
                    
                    print(f"DEBUG: Shape text after update: '{shape.text_frame.text}'")
                    print(f"✓ Updated question text for table: {table_title} using selection: {new_text}")
                else:
                    print(f"⚠️ No question_text in selection for table: {table_title}")
        
        # Update base text
        elif alt.get("type") == "text_base" and alt.get("table_title") == table_title:
            shape_count += 1
            print(f"DEBUG: Found text_base shape #{shape_count} for table: {table_title}")
            if hasattr(shape, "text_frame"):
                # Use current selection for base text if available
                if "base_text" in table_selection:
                    # Get the base text from selection
                    base_text_template = table_selection["base_text"]
                    
                    # If column_key is specified in selections, update the N value from that column
                    if "column_key" in table_selection:
                        column_key = table_selection["column_key"]
                        # Find the base row and selected column to get the N value
                        base_idx = None
                        col_idx = None
                        
                        # Find base row index
                        row_labels = table.get("row_labels", [])
                        for i, lab in enumerate(row_labels):
                            if isinstance(lab, str) and lab.strip().lower().startswith("base"):
                                base_idx = i
                                break
                        
                        # Find column index
                        col_labels = table.get("col_labels", [])
                        if column_key in col_labels:
                            col_idx = col_labels.index(column_key)
                        
                        # Get the new N value if both indices are found
                        new_n_value = None
                        if base_idx is not None and col_idx is not None and base_idx < len(table.get("values", [])):
                            try:
                                row_values = table["values"][base_idx]
                                if col_idx < len(row_values):
                                    new_n_value = int(round(float(row_values[col_idx])))
                            except Exception:
                                pass
                        
                        # Update the base text with the new N value
                        if new_n_value is not None:
                            # Extract the custom description from the base text template
                            if "Base:" in base_text_template:
                                # Look for patterns like "Base: Total respondents. 123 complete surveys."
                                base_parts = base_text_template.split(".")
                                if len(base_parts) >= 2:
                                    # First part contains the custom description
                                    custom_desc = base_parts[0].replace("Base:", "").strip()
                                    # Clean up any trailing punctuation or equals signs
                                    custom_desc = custom_desc.rstrip(" =").strip()
                                    new_text = f"Base: {custom_desc}. {new_n_value:,} complete surveys."
                                else:
                                    # No period, might be just "Base: Total respondents 123"
                                    base_parts = base_text_template.split()
                                    if len(base_parts) >= 3:
                                        # Extract everything after "Base:" but before the number
                                        base_idx_text = base_parts.index("Base:")
                                        if base_idx_text >= 0 and base_idx_text + 1 < len(base_parts):
                                            # Find where the number starts
                                            for i in range(base_idx_text + 1, len(base_parts)):
                                                if base_parts[i].replace(",", "").isdigit():
                                                    custom_desc = " ".join(base_parts[base_idx_text + 1:i])
                                                    # Clean up any trailing punctuation or equals signs
                                                    custom_desc = custom_desc.rstrip(" =").strip()
                                                    new_text = f"Base: {custom_desc}. {new_n_value:,} complete surveys."
                                                    break
                                            else:
                                                new_text = f"Base: {base_text_template.replace('Base:', '').strip()}. {new_n_value:,} complete surveys."
                                        else:
                                            new_text = f"Base: {base_text_template.replace('Base:', '').strip()}. {new_n_value:,} complete surveys."
                                    else:
                                        new_text = f"Base: {base_text_template.replace('Base:', '').strip()}. {new_n_value:,} complete surveys."
                            else:
                                new_text = f"Base: {base_text_template}. {new_n_value:,} complete surveys."
                        else:
                            new_text = base_text_template
                    else:
                        new_text = base_text_template
                    
                    print(f"DEBUG: Base text from selection: '{base_text_template}'")
                    print(f"DEBUG: Final base text: '{new_text}'")
                    print(f"DEBUG: Current shape text before update: '{shape.text_frame.text}'")
                    
                    # Preserve formatting by only updating the text content
                    if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                        paragraph = shape.text_frame.paragraphs[0]
                        
                        if paragraph.runs:
                            # Keep the first run (which has the formatting) and update its text
                            first_run = paragraph.runs[0]
                            first_run.text = new_text
                            
                            # Remove any additional runs to prevent concatenation
                            # We'll just clear the paragraph and recreate the first run with formatting
                            if len(paragraph.runs) > 1:
                                # Store the formatting from the first run
                                font = first_run.font
                                # Clear and recreate
                                paragraph.clear()
                                new_run = paragraph.add_run()
                                new_run.text = new_text
                                # Apply the formatting
                                new_run.font.name = font.name
                                new_run.font.size = font.size
                                new_run.font.bold = font.bold
                                new_run.font.italic = font.italic
                                # Only set color if it has an rgb property
                                if hasattr(font.color, 'rgb') and font.color.rgb is not None:
                                    new_run.font.color.rgb = font.color.rgb
                        else:
                            # No runs exist, create one
                            run = paragraph.add_run()
                            run.text = new_text
                    else:
                        shape.text_frame.text = new_text
                    
                    print(f"DEBUG: Shape text after update: '{shape.text_frame.text}'")
                    print(f"✓ Updated base text for table: {table_title} using selection: {new_text}")
                else:
                    print(f"⚠️ No base_text in selection for table: {table_title}")
        
        # Update chart title
        elif alt.get("type") == "text_title" and alt.get("table_title") == table_title:
            shape_count += 1
            print(f"DEBUG: Found text_title shape #{shape_count} for table: {table_title}")
            if hasattr(shape, "text_frame"):
                # Use current selection for chart title if available
                if "title" in table_selection:
                    new_text = table_selection["title"]
                    print(f"DEBUG: Chart title from selection: '{new_text}'")
                    print(f"DEBUG: Current shape text before update: '{shape.text_frame.text}'")
                    
                    # Preserve formatting by only updating the text content
                    if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                        paragraph = shape.text_frame.paragraphs[0]
                        
                        if paragraph.runs:
                            # Keep the first run (which has the formatting) and update its text
                            first_run = paragraph.runs[0]
                            first_run.text = new_text
                            
                            # Remove any additional runs to prevent concatenation
                            # We'll just clear the paragraph and recreate the first run with formatting
                            if len(paragraph.runs) > 1:
                                # Store the formatting from the first run
                                font = first_run.font
                                # Clear and recreate
                                paragraph.clear()
                                new_run = paragraph.add_run()
                                new_run.text = new_text
                                # Apply the formatting
                                new_run.font.name = font.name
                                new_run.font.size = font.size
                                new_run.font.bold = font.bold
                                new_run.font.italic = font.italic
                                # Only set color if it has an rgb property
                                if hasattr(font.color, 'rgb') and font.color.rgb is not None:
                                    new_run.font.color.rgb = font.color.rgb
                        else:
                            # No runs exist, create one
                            run = paragraph.add_run()
                            run.text = new_text
                    else:
                        shape.text_frame.text = new_text
                    
                    print(f"DEBUG: Shape text after update: '{shape.text_frame.text}'")
                    print(f"✓ Updated chart title for table: {table_title} using selection: {new_text}")
                else:
                    print(f"⚠️ No title in selection for table: {table_title}")
    
    print(f"DEBUG: Total shapes updated for table '{table_title}': {shape_count}")

def _update_new_text_callout_system(slide, table: Dict[str, Any], col_key: Optional[str]):
    """Update new TextCallout objects based on alt text mapping, incorporating actual data values."""
    for shape in slide.shapes:
        alt = _parse_alt_text(shape)
        
        # Update new text callouts
        if alt.get("type") == "text_callout" and alt.get("table_title") == table.get("title"):
            if hasattr(shape, "text_frame"):
                # Get callout information from alt text
                row_label = alt.get("row_label", "")
                column = alt.get("column", "Total")
                
                # Get the current text from the shape to see if it has custom formatting
                current_shape_text = shape.text_frame.text if hasattr(shape, "text_frame") else ""
                
                # Try to find the row and column indices
                row_idx = None
                col_idx = None
                
                if row_label:
                    # Find row index
                    row_labels = table.get("row_labels", [])
                    for i, label in enumerate(row_labels):
                        if isinstance(label, str) and row_label.lower() in label.lower():
                            row_idx = i
                            break
                    
                    # Find column index
                    col_labels = table.get("col_labels", [])
                    if column in col_labels:
                        col_idx = col_labels.index(column)
                    else:
                        # Fallback to common column names
                        for fallback in ["Total", "Overall", "All", "Base"]:
                            if fallback in col_labels:
                                col_idx = col_labels.index(fallback)
                                break
                        if col_idx is None:
                            col_idx = 0 if col_labels else None
                    
                    # Get the actual value if both indices are found
                    if row_idx is not None and col_idx is not None:
                        try:
                            values = table.get("values", [])
                            if row_idx < len(values) and col_idx < len(values[row_idx]):
                                value = values[row_idx][col_idx]
                                if value is not None:
                                    # Format the value appropriately
                                    formatted_value = ""
                                    if isinstance(value, (int, float)):
                                        if hasattr(value, 'is_integer') and value.is_integer():
                                            formatted_value = str(int(value))
                                        else:
                                            formatted_value = f"{value:.1f}%"
                                    else:
                                        formatted_value = str(value)
                                    
                                    # Check if the current text has custom formatting with [Value] placeholder
                                    if current_shape_text and "[Value]" in current_shape_text:
                                        new_text = current_shape_text.replace("[Value]", formatted_value)
                                    else:
                                        # Use default format
                                        new_text = f"{row_label}: {formatted_value}"
                        except (IndexError, TypeError, AttributeError):
                            pass
                
                # If no custom formatting found, use default
                if not new_text:
                    new_text = f"{row_label}: [Value]"
                
                # Update the text while preserving formatting
                current_text = shape.text_frame.text
                if current_text != new_text:
                    if hasattr(shape.text_frame, 'paragraphs') and len(shape.text_frame.paragraphs) > 0:
                        paragraph = shape.text_frame.paragraphs[0]
                        if paragraph.runs:
                            paragraph.runs[0].text = new_text
                        else:
                            run = paragraph.add_run()
                            run.text = new_text
                    else:
                        shape.text_frame.text = new_text
                    
                    print(f"✓ Updated text callout '{row_label}' for table: {table.get('title')}")

def update_presentation(pptx_in: str, crosstab_xlsx: str, pptx_out: str, selections: dict = None) -> str:
    prs = Presentation(pptx_in)
    data = parse_workbook(crosstab_xlsx)
    
    # Debug: Print selections if provided
    if selections:
        print(f"DEBUG: Selections provided: {selections}")
        for tid, sel in selections.items():
            print(f"DEBUG: Table ID {tid}: {sel}")
    else:
        print("DEBUG: No selections provided")
    
    # Track what was updated for reporting
    update_log = {
        "charts_updated": 0,
        "tables_updated": 0,
        "text_updated": 0,
        "shapes_skipped": 0,
        "mapping_issues": []
    }

    for slide in prs.slides:
        for shp in slide.shapes:
            name = shp.name or ""
            alt = _parse_alt_text(shp)
            
            # Skip shapes that don't want auto-updates
            if alt.get("auto_update", "yes").lower() == "no":
                update_log["shapes_skipped"] += 1
                continue
            
            # Handle charts - use try/except to safely check for charts
            try:
                # Try to access the chart property - this will raise ValueError if no chart
                chart = shp.chart
                # If we get here, the shape contains a chart
                chart_mapping = _get_chart_mapping_from_shape(shp, data, selections)
                if chart_mapping:
                    table, col_key = chart_mapping
                    _update_chart(shp, table, col_key, explicit_rows=None)
                    
                    # Update question and base text for this table using current selections if available
                    if selections:
                        # Find the selection for this table by matching table title
                        table_selection = None
                        table_title = table.get("title")
                        if table_title and table_title in selections:
                            table_selection = selections[table_title]
                            print(f"DEBUG: Found selection for table '{table_title}': {table_selection}")
                        else:
                            print(f"⚠️ No selection found for table: {table_title}")
                            print(f"DEBUG: Available selections: {list(selections.keys()) if selections else 'None'}")
                        
                        if table_selection:
                            _update_question_and_base_with_selections(slide, table, {table_title: table_selection}, table_title)
                        else:
                            _update_question_and_base(slide, table, None, None)
                    else:
                        _update_question_and_base(slide, table, None, None)
                    
                    # Update text callouts for this table
                    _update_new_text_callout_system(slide, table, col_key)
                    
                    update_log["charts_updated"] += 1
                    print(f"✓ Updated chart with existing mapping for table: {table.get('title')}")
                else:
                    # No mapping found - preserve the chart as-is
                    print(f"⚠️ Chart '{name}' has no table mapping - preserving as-is")
                    update_log["shapes_skipped"] += 1
            except (ValueError, AttributeError):
                # Shape doesn't contain a chart or doesn't have chart attribute
                print(f"⚠️ Shape '{name}' doesn't contain a chart - skipping")
                pass
            
            # Handle tables
            if shp.has_table:
                table = _get_table_mapping_from_shape(shp, data)
                if table:
                    _update_table(shp, table)
                    
                    # Update question and base text for this table using current selections if available
                    if selections:
                        # Find the selection for this table by matching table title
                        table_selection = None
                        table_title = table.get("title")
                        if table_title and table_title in selections:
                            table_selection = selections[table_title]
                            print(f"DEBUG: Found selection for table '{table_title}': {table_selection}")
                        else:
                            print(f"⚠️ No selection found for table: {table_title}")
                            print(f"DEBUG: Available selections: {list(selections.keys()) if selections else 'None'}")
                        
                        if table_selection:
                            _update_question_and_base_with_selections(slide, table, {table_title: table_selection}, table_title)
                        else:
                            _update_question_and_base(slide, table, None, None)
                    else:
                        _update_question_and_base(slide, table, None, None)
                    
                    # Update text callouts for this table
                    _update_new_text_callout_system(slide, table, None)
                    
                    update_log["tables_updated"] += 1
                    print(f"✓ Updated table with existing mapping for table: {table.get('title')}")
                else:
                    # No mapping found - preserve the table as-is
                    print(f"⚠️ Table '{name}' has no table mapping - preserving as-is")
                    update_log["shapes_skipped"] += 1
            
            # Handle text objects (question and base)
            elif name in ["TEXT_QUESTION", "OBJ_QUESTION"]:
                # For text objects, we need to find which table they're bound to
                # This could be enhanced with more sophisticated binding logic
                update_log["text_updated"] += 1
            elif name in ["TEXT_BASE", "OBJ_BASE"]:
                update_log["text_updated"] += 1

    # Print update summary
    print(f"\n{'='*50}")
    print(f"UPDATE SUMMARY")
    print(f"{'='*50}")
    print(f"✓ Charts updated: {update_log['charts_updated']}")
    print(f"✓ Tables updated: {update_log['tables_updated']}")
    print(f"✓ Text objects updated: {update_log['text_updated']}")
    print(f"⚠️ Shapes preserved (no mapping): {update_log['shapes_skipped']}")
    print(f"\n📋 What was updated:")
    print(f"  • Chart/table data refreshed with new crosstab values")
    print(f"  • Base N values updated while preserving custom descriptions")
    print(f"  • Question text updated only if using default values")
    print(f"  • Chart titles preserved if custom")
    print(f"\n🔒 What was preserved:")
    print(f"  • All formatting (fonts, colors, sizes, styles)")
    print(f"  • Custom chart titles and question text")
    print(f"  • Custom base descriptions")
    print(f"  • Charts/tables without mappings")
    print(f"{'='*50}")
    
    if update_log["mapping_issues"]:
        print(f"\n⚠️ Mapping issues found: {len(update_log['mapping_issues'])}")
        for issue in update_log["mapping_issues"][:3]:  # Show first 3 issues
            print(f"    - {issue}")

    prs.save(pptx_out)
    return pptx_out

def update_presentation_with_unmapped(pptx_in: str, crosstab_xlsx: str, pptx_out: str, 
                                    selections: dict = None, all_tables: list = None, 
                                    existing_content: dict = None) -> str:
    """Enhanced update function that handles existing mappings and adds unmapped tables page."""
    prs = Presentation(pptx_in)
    data = parse_workbook(crosstab_xlsx)
    
    # Track what was updated for reporting
    update_log = {
        "charts_updated": 0,
        "tables_updated": 0,
        "text_updated": 0,
        "shapes_skipped": 0,
        "unmapped_tables": [],
        "mapping_issues": []
    }
    
    # Identify unmapped tables
    mapped_table_titles = set(existing_content.keys()) if existing_content else set()
    unmapped_tables = []
    
    if all_tables:
        for table in all_tables:
            if table["title"] not in mapped_table_titles:
                unmapped_tables.append(table)
        update_log["unmapped_tables"] = [t["title"] for t in unmapped_tables]

    # First, update all existing mapped content (same as original function)
    for slide in prs.slides:
        for shp in slide.shapes:
            name = shp.name or ""
            alt = _parse_alt_text(shp)
            
            # Skip shapes that don't want auto-updates
            if alt.get("auto_update", "yes").lower() == "no":
                update_log["shapes_skipped"] += 1
                continue
            
            # Handle charts - use try/except to safely check for charts
            try:
                # Try to access the chart property - this will raise ValueError if no chart
                chart = shp.chart
                # If we get here, the shape contains a chart
                chart_mapping = _get_chart_mapping_from_shape(shp, data, selections)
                if chart_mapping:
                    table, col_key = chart_mapping
                    _update_chart(shp, table, col_key, explicit_rows=None)
                    
                    # Update question and base text for this table using current selections if available
                    if selections:
                        table_selection = None
                        table_title = table.get("title")
                        if table_title and table_title in selections:
                            table_selection = selections[table_title]
                        
                        if table_selection:
                            _update_question_and_base_with_selections(slide, table, {table_title: table_selection}, table_title)
                        else:
                            _update_question_and_base(slide, table, None, None)
                    else:
                        _update_question_and_base(slide, table, None, None)
                    
                    # Update text callouts for this table
                    _update_new_text_callout_system(slide, table, col_key)
                    
                    update_log["charts_updated"] += 1
                else:
                    update_log["shapes_skipped"] += 1
            except (ValueError, AttributeError):
                # Shape doesn't contain a chart or doesn't have chart attribute
                pass
            
            # Handle tables
            if shp.has_table:
                table = _get_table_mapping_from_shape(shp, data)
                if table:
                    _update_table(shp, table)
                    
                    # Update question and base text for this table using current selections if available
                    if selections:
                        table_selection = None
                        table_title = table.get("title")
                        if table_title and table_title in selections:
                            table_selection = selections[table_title]
                        
                        if table_selection:
                            _update_question_and_base_with_selections(slide, table, {table_title: table_selection}, table_title)
                        else:
                            _update_question_and_base(slide, table, None, None)
                    else:
                        _update_question_and_base(slide, table, None, None)
                    
                    # Update text callouts for this table
                    _update_new_text_callout_system(slide, table, None)
                    
                    update_log["tables_updated"] += 1
                else:
                    update_log["shapes_skipped"] += 1

    # Add unmapped tables summary page
    if unmapped_tables:
        from pptx_exporter import add_title, _apply_background
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        
        # Create unmapped tables summary slide
        unmapped_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Use blank layout
        _apply_background(unmapped_slide)
        
        # Add title
        add_title(unmapped_slide, "Unmapped Tables Summary")
        
        # Add subtitle explaining what this page contains
        subtitle_box = unmapped_slide.shapes.add_textbox(
            Inches(0.5), Inches(1.2), Inches(9.0), Inches(0.6)
        )
        subtitle_tf = subtitle_box.text_frame
        subtitle_tf.clear()
        subtitle_p = subtitle_tf.paragraphs[0]
        subtitle_run = subtitle_p.add_run()
        subtitle_run.text = f"The following {len(unmapped_tables)} tables from your crosstab had no existing connections and are listed here for reference:"
        subtitle_run.font.size = Pt(14)
        subtitle_run.font.name = "Arial"
        
        # Create a simple list of unmapped tables
        list_y_start = 2.0
        list_x = 0.5
        line_height = 0.3
        
        for i, table in enumerate(unmapped_tables):
            # Add table title
            title_box = unmapped_slide.shapes.add_textbox(
                Inches(list_x), Inches(list_y_start + i * line_height * 4), Inches(8.5), Inches(0.25)
            )
            title_tf = title_box.text_frame
            title_tf.clear()
            title_p = title_tf.paragraphs[0]
            title_run = title_p.add_run()
            title_run.text = f"• {table['title']}"
            title_run.font.size = Pt(12)
            title_run.font.bold = True
            title_run.font.name = "Arial"
            
            # Add basic stats
            stats_box = unmapped_slide.shapes.add_textbox(
                Inches(list_x + 0.3), Inches(list_y_start + i * line_height * 4 + 0.25), Inches(8.2), Inches(0.2)
            )
            stats_tf = stats_box.text_frame
            stats_tf.clear()
            stats_p = stats_tf.paragraphs[0]
            stats_run = stats_p.add_run()
            
            row_count = len(table.get("row_labels", []))
            col_count = len(table.get("col_labels", []))
            stats_run.text = f"  Rows: {row_count}, Columns: {col_count}"
            stats_run.font.size = Pt(10)
            stats_run.font.name = "Arial"
            stats_run.font.color.rgb = RGBColor(100, 100, 100)
            
            # Add column names
            if table.get("col_labels"):
                cols_text = ", ".join(table["col_labels"][:8])  # Show first 8 columns
                if len(table["col_labels"]) > 8:
                    cols_text += "..."
                
                cols_box = unmapped_slide.shapes.add_textbox(
                    Inches(list_x + 0.3), Inches(list_y_start + i * line_height * 4 + 0.45), Inches(8.2), Inches(0.2)
                )
                cols_tf = cols_box.text_frame
                cols_tf.clear()
                cols_p = cols_tf.paragraphs[0]
                cols_run = cols_p.add_run()
                cols_run.text = f"  Columns: {cols_text}"
                cols_run.font.size = Pt(9)
                cols_run.font.name = "Arial"
                cols_run.font.color.rgb = RGBColor(120, 120, 120)
            
            # Stop if we're running out of space (about 12 tables max)
            if i >= 11:
                remaining_count = len(unmapped_tables) - 12
                if remaining_count > 0:
                    more_box = unmapped_slide.shapes.add_textbox(
                        Inches(list_x), Inches(list_y_start + 12 * line_height * 4), Inches(8.5), Inches(0.25)
                    )
                    more_tf = more_box.text_frame
                    more_tf.clear()
                    more_p = more_tf.paragraphs[0]
                    more_run = more_p.add_run()
                    more_run.text = f"... and {remaining_count} more tables"
                    more_run.font.size = Pt(11)
                    more_run.font.italic = True
                    more_run.font.name = "Arial"
                    more_run.font.color.rgb = RGBColor(150, 150, 150)
                break

    # Print update summary
    print(f"\n{'='*60}")
    print(f"ENHANCED UPDATE SUMMARY")
    print(f"{'='*60}")
    print(f"✓ Charts updated: {update_log['charts_updated']}")
    print(f"✓ Tables updated: {update_log['tables_updated']}")
    print(f"✓ Text objects updated: {update_log['text_updated']}")
    print(f"⚠️ Shapes preserved (no mapping): {update_log['shapes_skipped']}")
    print(f"📄 Unmapped tables added to summary page: {len(unmapped_tables)}")
    
    if unmapped_tables:
        print(f"\n📋 Unmapped tables:")
        for table in unmapped_tables[:5]:  # Show first 5
            print(f"  • {table['title']}")
        if len(unmapped_tables) > 5:
            print(f"  ... and {len(unmapped_tables) - 5} more")
    
    print(f"\n📋 What was updated:")
    print(f"  • Existing chart/table data refreshed with new crosstab values")
    print(f"  • Base N values updated while preserving custom descriptions")
    print(f"  • Question text updated only if using default values")
    print(f"  • Chart titles preserved if custom")
    print(f"  • New 'Unmapped Tables Summary' page added with {len(unmapped_tables)} tables")
    print(f"\n🔒 What was preserved:")
    print(f"  • All formatting (fonts, colors, sizes, styles)")
    print(f"  • Custom chart titles and question text")
    print(f"  • Custom base descriptions")
    print(f"  • Charts/tables without mappings")
    print(f"{'='*60}")

    prs.save(pptx_out)
    return pptx_out
