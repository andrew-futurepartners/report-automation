import streamlit as st
from crosstab_parser import parse_workbook
from pptx_exporter import export_pptx
from deck_update import update_presentation, update_presentation_with_unmapped
from pptx import Presentation
from deck_update import _parse_alt_text

def _format_number_with_commas(number):
    """Format a number with comma separators for thousands places."""
    if number is None:
        return None
    return f"{number:,}"

def parse_existing_powerpoint(pptx_file):
    """Parse existing PowerPoint to extract current content and settings."""
    try:
        prs = Presentation(pptx_file)
        existing_content = {}
        
        for slide in prs.slides:
            for shape in slide.shapes:
                alt = _parse_alt_text(shape)
                
                if alt.get("type") in ["chart", "table", "question_text", "text_question", "text_base", "text_title", "text_callout"]:
                    table_title = alt.get("table_title")
                    if table_title:
                        if table_title not in existing_content:
                            existing_content[table_title] = {
                                "title": table_title,
                                "question_text": "",
                                "base_text": "",
                                "chart_type": "bar_h",
                                "custom_base_description": "",  # Store custom base description
                                "custom_question": "",  # Store custom question text
                                "callouts": [],  # Initialize callouts list
                                # Presence flags for connected objects
                                "has_chart": False,
                                "has_table": False,
                                "has_title": False,
                                "has_base": False,
                                "has_question": False,
                                "has_callouts": False
                            }
                        
                        # Track object presence
                        if alt.get("type") == "chart":
                            existing_content[table_title]["has_chart"] = True
                            if "column" in alt:
                                existing_content[table_title]["chart_column"] = alt.get("column")
                        elif alt.get("type") == "table":
                            existing_content[table_title]["has_table"] = True

                        # Extract question text
                        if alt.get("type") in ["question_text", "text_question"] and hasattr(shape, "text_frame"):
                            question_text = shape.text_frame.text
                            if question_text.startswith("Question: "):
                                question_text = question_text[10:]  # Remove "Question: " prefix
                            existing_content[table_title]["question_text"] = question_text
                            existing_content[table_title]["custom_question"] = question_text
                            existing_content[table_title]["has_question"] = True
                        
                        # Extract base text
                        elif alt.get("type") == "text_base" and hasattr(shape, "text_frame"):
                            base_text = shape.text_frame.text
                            existing_content[table_title]["base_text"] = base_text
                            existing_content[table_title]["has_base"] = True
                            
                            # Extract custom base description (everything before the N count)
                            if "Base:" in base_text:
                                # Look for patterns like "Base: Total respondents. 123 complete surveys."
                                # or "Base: Total respondents. 123"
                                base_parts = base_text.split(".")
                                if len(base_parts) >= 2:
                                    # First part contains the custom description
                                    custom_desc = base_parts[0].replace("Base:", "").strip()
                                    # Clean up any trailing punctuation or equals signs
                                    custom_desc = custom_desc.rstrip(" =").strip()
                                    existing_content[table_title]["custom_base_description"] = custom_desc
                                else:
                                    # No period, might be just "Base: Total respondents 123"
                                    base_parts = base_text.split()
                                    if len(base_parts) >= 3:
                                        # Extract everything after "Base:" but before the number
                                        base_idx = base_parts.index("Base:")
                                        if base_idx >= 0 and base_idx + 1 < len(base_parts):
                                            # Find where the number starts
                                            for i in range(base_idx + 1, len(base_parts)):
                                                if base_parts[i].replace(",", "").isdigit():
                                                    custom_desc = " ".join(base_parts[base_idx + 1:i])
                                                    # Clean up any trailing punctuation or equals signs
                                                    custom_desc = custom_desc.rstrip(" =").strip()
                                                    existing_content[table_title]["custom_base_description"] = custom_desc
                                                    break
                        
                        # Extract chart title
                        elif alt.get("type") == "text_title" and hasattr(shape, "text_frame"):
                            existing_content[table_title]["title"] = shape.text_frame.text
                            existing_content[table_title]["has_title"] = True
                        
                        # Extract text callouts
                        elif alt.get("type") == "text_callout" and hasattr(shape, "text_frame"):
                            if "callouts" not in existing_content[table_title]:
                                existing_content[table_title]["callouts"] = []
                            
                            callout_text = shape.text_frame.text
                            callout_info = {
                                "row_label": alt.get("row", alt.get("row_label", "")),
                                "column_key": alt.get("column", "Total"),
                                "text": callout_text,
                                "position": (0.5, 7.0, 9.0, 0.4),  # Default position
                                "font_size": 12,
                                "font_bold": True,
                                "metric_type": alt.get("metric_type", "percentage")
                            }
                            existing_content[table_title]["callouts"].append(callout_info)
                            existing_content[table_title]["has_callouts"] = True
        
        return existing_content
    except Exception as e:
        st.error(f"Error parsing PowerPoint: {e}")
        return {}

st.set_page_config(page_title="Report Relay", layout="wide")

st.title("Report Relay")
st.write("An efficient handoff from crosstab to report.")

# Step 1: Choose workflow type
st.subheader("Step 1: Choose your workflow")
workflow_type = st.radio(
    "Are you creating a new report or updating an existing one?",
    ["Create New Report", "Update Existing Report"],
    index=0,
    help="""
    **Create New Report**: Upload crosstab Excel → Choose chart types and customize → Export new PowerPoint
    
    **Update Existing Report**: Upload PowerPoint → Detect connections → Upload crosstab → Review mappings → Update PowerPoint
    """
)

# Add helpful workflow information
with st.expander("💡 How each workflow works", expanded=False):
    if workflow_type == "Create New Report":
        st.write("""
        **New Report Workflow:**
        1. Upload crosstab Excel file
        2. Choose chart types and customize titles for each table
        3. Export new PowerPoint presentation
        
        **Perfect for:** First-time report creation, one-off reports
        """)
    else:
        st.write("""
        **Update Existing Report Workflow:**
        1. Upload existing PowerPoint presentation
        2. System detects existing table connections
        3. Upload new crosstab Excel file  
        4. Review tables with/without existing connections
        5. Configure column selection (switch segments quickly)
        6. Update presentation (preserves formatting, adds unmapped tables summary)
        
        **Perfect for:** Regular report updates, preserving custom formatting and content
        
        **Key Benefits:**
        - Custom work is preserved when updating reports
        - Base N values are automatically refreshed with new data
        - Chart titles, questions, and base descriptions can be customized and preserved
        - New tables are automatically listed on a summary page
        - Quick segment switching (Total → Male → Female → Gen Z, etc.)
        """)

# Initialize session state for workflow management
if "workflow_step" not in st.session_state:
    st.session_state.workflow_step = "start"
if "existing_content" not in st.session_state:
    st.session_state.existing_content = {}
if "data" not in st.session_state:
    st.session_state.data = None

# Workflow branching
if workflow_type == "Create New Report":
    # NEW REPORT WORKFLOW
    st.subheader("Step 2: Upload crosstab data")
    uploaded = st.file_uploader("Upload crosstab Excel", type=["xlsx", "xls"], key="new_report_excel")
    
    if uploaded:
        with st.spinner("Parsing workbook..."):
            with open("uploaded.xlsx", "wb") as f:
                f.write(uploaded.getbuffer())
            data = parse_workbook("uploaded.xlsx")
            st.session_state.data = data

        st.success(f"Found {len(data['tables'])} tables")
        st.divider()
        
        # Show default choice controls for new reports
        st.subheader("Step 3: Configure visualizations")
        
        # Build banner list preserving datafile order (from the first table that has banners)
        all_columns = []
        found_order = False
        for table in data["tables"]:
            meta = table.get("meta", {})
            banners = meta.get("col_banners") or table.get("col_labels", [])
            if banners and not found_order:
                for b in banners:
                    if b is None:
                        continue
                    norm_b = str(b).replace("\xa0", " ").strip()
                    if norm_b and norm_b not in all_columns:
                        all_columns.append(norm_b)
                found_order = len(all_columns) > 0
        # Fallback: aggregate in encountered order if no table provided banners
        if not all_columns:
            for table in data["tables"]:
                banners = table.get("col_labels", [])
                for b in banners:
                    if b is None:
                        continue
                    norm_b = str(b).replace("\xa0", " ").strip()
                    if norm_b and norm_b not in all_columns:
                        all_columns.append(norm_b)
        
        # Default selections
        col1, col2 = st.columns(2)
        with col1:
            default_choice = st.selectbox("Default visualization", ["Bar Horizontal", "Bar Vertical", "Donut", "Line", "Chart + Table"], index=0)
            apply_all_btn = st.button("Apply chart type to all")
        
        with col2:
            default_column = st.selectbox("Default column", all_columns, index=all_columns.index("Total") if "Total" in all_columns else 0)
            apply_column_btn = st.button("Apply column to all")

else:
    # UPDATE EXISTING REPORT WORKFLOW
    st.subheader("Step 2: Upload existing PowerPoint")
    existing_ppt = st.file_uploader("Upload the PowerPoint to update", type=["pptx"], key="existing_ppt")
    
    if existing_ppt:
        with st.spinner("Parsing existing PowerPoint..."):
            existing_content = parse_existing_powerpoint(existing_ppt)
            st.session_state.existing_content = existing_content
            # Save the PowerPoint file for later use
            with open("to_update.pptx", "wb") as pf:
                pf.write(existing_ppt.getbuffer())

        if existing_content:
            st.success(f"Found connections for {len(existing_content)} tables")
            
            # Show found connections
            with st.expander(f"📊 Found connections for {len(existing_content)} tables", expanded=True):
                for table_title in existing_content.keys():
                    st.write(f"• {table_title}")
        else:
            st.info("No existing table connections found in this PowerPoint.")
        
        st.divider()
        
        # Step 3: Upload crosstab for updates
        st.subheader("Step 3: Upload new crosstab data")
        uploaded = st.file_uploader("Upload crosstab Excel", type=["xlsx", "xls"], key="update_report_excel")
        
        if uploaded:
            with st.spinner("Parsing workbook..."):
                with open("uploaded.xlsx", "wb") as f:
                    f.write(uploaded.getbuffer())
                data = parse_workbook("uploaded.xlsx")
                st.session_state.data = data

            st.success(f"Found {len(data['tables'])} tables in crosstab")
            
            # Categorize tables into connected vs unconnected
            connected_tables = []
            unconnected_tables = []
            
            for table in data["tables"]:
                if table["title"] in existing_content:
                    connected_tables.append(table)
                else:
                    unconnected_tables.append(table)
            
            st.divider()
            
            # Show categorized results
            st.subheader("Step 4: Review table connections")
            
            # Connected tables section (collapsible)
            if connected_tables:
                with st.expander(f"✅ Tables with existing connections ({len(connected_tables)})", expanded=True):
                    st.success("These tables will be updated with new data while preserving your custom formatting.")
                    for table in connected_tables:
                        st.write(f"• {table['title']}")
            
            # Unconnected tables section (collapsible)  
            if unconnected_tables:
                with st.expander(f"➕ Tables with NO connections ({len(unconnected_tables)})", expanded=True):
                    st.info("These tables are new or don't have existing connections. They will be added to a summary page.")
                    for table in unconnected_tables:
                        st.write(f"• {table['title']}")
            
            # Show column selection for update workflow
            if data["tables"]:
                # Build banner list preserving datafile order (from the first table that has banners)
                all_columns = []
                found_order = False
                for table in data["tables"]:
                    meta = table.get("meta", {})
                    banners = meta.get("col_banners") or table.get("col_labels", [])
                    if banners and not found_order:
                        for b in banners:
                            if b is None:
                                continue
                            norm_b = str(b).replace("\xa0", " ").strip()
                            if norm_b and norm_b not in all_columns:
                                all_columns.append(norm_b)
                        found_order = len(all_columns) > 0
                if not all_columns:
                    for table in data["tables"]:
                        banners = table.get("col_labels", [])
                        for b in banners:
                            if b is None:
                                continue
                            norm_b = str(b).replace("\xa0", " ").strip()
                            if norm_b and norm_b not in all_columns:
                                all_columns.append(norm_b)
                
                st.subheader("Step 5: Configure column selection")
                
                col1, _ = st.columns([1, 1])
                with col1:
                    default_column = st.selectbox("Default column for updates", all_columns, 
                                                index=all_columns.index("Total") if "Total" in all_columns else 0,
                                                key="update_default_column")
                    apply_column_btn = st.button("Apply column to all connected tables", key="update_apply_column")
                
            
            # Set defaults for the rest of the logic
            default_choice = "Bar Horizontal"
            apply_all_btn = False
        else:
            # If no data yet, set default values
            default_column = "Total"
            apply_column_btn = False

    # Selections state
    if "selections" not in st.session_state:
        st.session_state["selections"] = {}

# Only show table configuration if we have data loaded
if st.session_state.data is not None:
    data = st.session_state.data
    existing_content = st.session_state.existing_content
    
    # Ensure default_column and apply_column_btn are defined for both workflows
    if workflow_type == "Create New Report":
        # These were defined earlier in the new report section
        pass
    else:
        # For update workflow, check if they were defined in the update section
        if 'default_column' not in locals():
            default_column = "Total"
            apply_column_btn = False
    
    # Determine which tables to show configuration for
    if workflow_type == "Create New Report":
        # Show all tables for new reports
        tables_to_configure = data["tables"]
        config_title = "Configure all tables"
    else:
        # For updates, only show connected tables that have existing settings
        # These can be modified if needed
        tables_to_configure = []
        for table in data["tables"]:
            if table["title"] in existing_content:
                tables_to_configure.append(table)
        config_title = f"Configure existing tables ({len(tables_to_configure)})"
    
    # Only show configuration section if there are tables to configure
    if tables_to_configure:
        st.divider()
        st.subheader(config_title)

    # Per-table controls
    for t in tables_to_configure:
        tid = t["id"]
        with st.expander(f"{t['title']}  ({tid})", expanded=False):
            cols = st.columns([2, 1, 2])
            # Lookup any existing mapping info for this table (from PPT)
            existing_table = existing_content.get(t["title"], {})
            with cols[0]:
                st.write("**Row labels**")
                # Show row labels as a concise bullet list; limit to first 12
                _all_rows = [rl for rl in t["row_labels"] if isinstance(rl, str) and rl.strip()]
                _preview = _all_rows[:12]
                if _preview:
                    st.markdown("\n".join([f"- {rl}" for rl in _preview]))
                _remaining = max(0, len(_all_rows) - len(_preview))
                if _remaining > 0:
                    st.markdown(f"- … and {_remaining} more")

            with cols[1]:
                options = ["Bar Horizontal", "Bar Vertical", "Donut", "Line", "Chart + Table"]
                if apply_all_btn:
                    choice = default_choice
                else:
                    choice = st.session_state["selections"].get(tid, {}).get("chart_type_label", default_choice)
                choice = st.selectbox("Chart type", options, key=f"ctype_{tid}", index=options.index(choice))
                
                # Column selection for this table
                # Keep Data column as banners only, add Metric selector when available
                meta = t.get("meta", {})
                banners = meta.get("col_banners") or t.get("col_labels", [])
                groups = meta.get("col_groups") or ["" for _ in banners]
                combined_labels = t.get("col_labels", [])

                # Derive metric options (unique, order-preserving, non-empty)
                metric_options = []
                seen_metrics = set()
                for g in groups:
                    if isinstance(g, str):
                        g2 = g.strip()
                        if g2 and g2 not in seen_metrics:
                            seen_metrics.add(g2)
                            metric_options.append(g2)
                has_metrics = len(metric_options) > 0
                if not has_metrics:
                    metric_options = [""]

                # Default metric: prefer session state; else metric parsed from mapped chart column; else left-most
                leftmost_metric = next((g for g in groups if isinstance(g, str) and g.strip()), "")
                current_metric = st.session_state["selections"].get(tid, {}).get("metric_key")
                if not current_metric:
                    mapped_col = existing_table.get("chart_column")
                    if isinstance(mapped_col, str) and "|" in mapped_col:
                        try:
                            banner_part, metric_part = [p.strip() for p in mapped_col.split("|", 1)]
                            current_metric = metric_part
                        except Exception:
                            current_metric = None
                if not current_metric:
                    current_metric = leftmost_metric
                if current_metric not in metric_options:
                    current_metric = metric_options[0]

                if has_metrics:
                    current_metric = st.selectbox("Metric", metric_options, key=f"metric_{tid}", index=metric_options.index(current_metric))
                
                # Handle column application for both new reports and updates
                if workflow_type == "Create New Report" and apply_column_btn:
                    # Apply default banner to all tables for new reports
                    current_col = default_column if default_column in banners else ("Total" if "Total" in banners else banners[0] if banners else "Total")
                elif workflow_type == "Update Existing Report" and apply_column_btn:
                    # Apply default banner to all connected tables for updates
                    current_col = default_column if default_column in banners else ("Total" if "Total" in banners else banners[0] if banners else "Total")
                else:
                    # Use existing banner selection or mapped chart banner; else default
                    current_col = st.session_state["selections"].get(tid, {}).get("banner_key")
                    if not current_col:
                        mapped_col = existing_table.get("chart_column")
                        if isinstance(mapped_col, str):
                            if "|" in mapped_col:
                                try:
                                    banner_part, metric_part = [p.strip() for p in mapped_col.split("|", 1)]
                                    current_col = banner_part
                                except Exception:
                                    current_col = mapped_col.strip()
                            else:
                                current_col = mapped_col.strip()
                    if not current_col:
                        current_col = "Total"
                    if current_col not in banners:
                        current_col = "Total" if "Total" in banners else banners[0] if banners else "Total"
                
                selected_col = st.selectbox("Data column", banners, 
                                         index=banners.index(current_col) if current_col in banners else 0,
                                         key=f"col_{tid}")

                # Resolve combined label for selected (metric, banner)
                pair_to_combined = {}
                for i, b in enumerate(banners):
                    g = groups[i] if i < len(groups) else ""
                    lab = combined_labels[i] if i < len(combined_labels) else b
                    pair_to_combined[(g or "", b)] = lab
                combined_selected = pair_to_combined.get(((current_metric or ""), selected_col))
                if not combined_selected:
                    # Fallback to the first column with this banner
                    for i, b in enumerate(banners):
                        if b == selected_col:
                            combined_selected = combined_labels[i]
                            break
                if not combined_selected and combined_labels:
                    combined_selected = combined_labels[0]
            with cols[2]:
                # Use existing content if available, otherwise use defaults
                has_title_obj = existing_table.get("has_title", False)
                has_base_obj = existing_table.get("has_base", False)
                has_question_obj = existing_table.get("has_question", False)
                
                # Chart title - prioritize existing custom title, then session state, then table title
                existing_title = existing_table.get("title", "")
                session_title = st.session_state["selections"].get(tid, {}).get("title", "")
                
                # Priority: existing content > session state > table title
                if existing_title and existing_title != t["title"]:
                    default_title = existing_title
                elif session_title:
                    default_title = session_title
                else:
                    default_title = t["title"]
                
                # Show indicator if using existing custom title
                title_label = "Chart title"
                if existing_title and existing_title != t["title"]:
                    title_label += " (Previously: " + existing_title + ")"
                
                title_val = None
                if has_title_obj:
                    title_val = st.text_input(title_label, value=default_title, key=f"title_{tid}")

                # Base text logic - preserve custom descriptions while updating N values
                def _find_base_idx(labels):
                    for i, lab in enumerate(labels):
                        if isinstance(lab, str) and lab.strip().lower().startswith("base"):
                            return i
                    return None
                
                base_idx = _find_base_idx(t["row_labels"])
                # Determine the Total column under the selected metric if applicable
                total_idx = None
                total_candidates = []
                if (" | " in (combined_selected or "")):
                    # Prefer banner-qualified by current metric first (Banner | Metric)
                    metric_name = (current_metric or "").strip()
                    if metric_name:
                        total_candidates.append(f"Total | {metric_name}")
                total_candidates.append("Total")
                for cand in total_candidates:
                    if cand in t["col_labels"]:
                        total_idx = t["col_labels"].index(cand)
                        break
                if total_idx is None:
                    total_idx = 0 if t["col_labels"] else None
                
                # Calculate new base N value
                new_base_n = None
                if base_idx is not None and total_idx is not None and base_idx < len(t["values"]) and total_idx < len(t["values"][base_idx]):
                    try:
                        new_base_n = int(round(float(t["values"][base_idx][total_idx])))
                    except Exception:
                        new_base_n = None
                
                # Determine base text default - preserve custom descriptions
                # Priority: existing content > session state > calculated default
                if existing_table.get("custom_base_description"):
                    # Use existing custom description with new N value
                    custom_desc = existing_table["custom_base_description"]
                    if new_base_n is not None:
                        # Use the custom description as-is, don't force "Total respondents"
                        default_base = f"Base: {custom_desc}. {_format_number_with_commas(new_base_n)} complete surveys."
                    else:
                        default_base = f"Base: {custom_desc}."
                elif existing_table.get("base_text"):
                    # Use existing base text as-is
                    default_base = existing_table["base_text"]
                else:
                    # Check session state as fallback
                    default_base = st.session_state["selections"].get(tid, {}).get("base_text")
                    if default_base is None:
                        # Calculate default from crosstab
                        if new_base_n is not None:
                            default_base = f"Base: Total respondents. {_format_number_with_commas(new_base_n)} complete surveys."
                        else:
                            default_base = "Base: Total respondents."
                
                # Show indicator if using existing custom base description
                base_label = "Base text"
                if existing_table.get("custom_base_description"):
                    base_label += f" (Previously: {existing_table['custom_base_description']})"
                
                base_text_val = None
                if has_base_obj:
                    base_text_val = st.text_input(base_label, value=default_base, key=f"base_{tid}")
                
                # Question text - preserve custom questions
                existing_q = existing_table.get("custom_question", "")
                session_q = st.session_state["selections"].get(tid, {}).get("question_text", "")
                
                # Priority: existing content > session state > table title
                if existing_q and existing_q != t["title"]:
                    default_q = existing_q
                elif session_q:
                    default_q = session_q
                else:
                    default_q = t["title"]
                
                # Show indicator if using existing custom question
                question_label = "Question text"
                if existing_q and existing_q != t["title"]:
                    question_label += " (Previously: " + existing_q + ")"
                
                question_text_val = None
                if has_question_obj:
                    question_text_val = st.text_input(question_label, value=default_q, key=f"qtext_{tid}")

            # Row sorting management section
            st.write("**Row Sorting**")
            
            # Toggle checkbox for row sorting
            enable_sorting = st.checkbox(f"Enable row sorting for this table", value=False, key=f"enable_sorting_{tid}")
            
            if enable_sorting:
                st.info("💡 Rows will be sorted by their values (descending). You can exclude certain rows from sorting.")
                
                # Get available rows for exclusion
                available_rows = [row for row in t["row_labels"] if isinstance(row, str) and row.strip()]
                
                # Common rows that users might want to exclude from sorting
                suggested_excludes = []
                for row in available_rows:
                    row_lower = row.lower().strip()
                    if any(pattern in row_lower for pattern in ["other", "none", "don't know", "no answer", "prefer not", "n/a"]):
                        suggested_excludes.append(row)
                
                # Row exclusion selection
                exclude_options = ["None (sort all rows)"] + available_rows
                default_excludes = ["None (sort all rows)"]
                if suggested_excludes:
                    default_excludes = suggested_excludes
                
                excluded_rows = st.multiselect(
                    "Rows to exclude from sorting (will appear at bottom)",
                    exclude_options,
                    default=default_excludes,
                    key=f"sort_exclude_{tid}",
                    help="Select rows that should remain at the bottom and not be sorted by value"
                )
                
                # Remove "None" option if other selections are made
                if "None (sort all rows)" in excluded_rows and len(excluded_rows) > 1:
                    excluded_rows = [row for row in excluded_rows if row != "None (sort all rows)"]
                elif excluded_rows == ["None (sort all rows)"]:
                    excluded_rows = []
                
                # Show preview of what will be sorted vs excluded
                if excluded_rows:
                    sortable_rows = [row for row in available_rows if row not in excluded_rows]
                    st.write(f"**Will be sorted:** {len(sortable_rows)} rows")
                    st.write(f"**Will stay at bottom:** {len(excluded_rows)} rows ({', '.join(excluded_rows[:3])}{'...' if len(excluded_rows) > 3 else ''})")
                else:
                    st.write(f"**Will be sorted:** All {len(available_rows)} rows")
            else:
                excluded_rows = []
            
            st.divider()
            
            # Callout management section (only if connected in PPT)
            # Check if there are existing callouts from PowerPoint or session state
            existing_table = existing_content.get(t["title"], {})
            existing_callouts = existing_table.get("callouts", []) if existing_table else []
            has_callouts_obj = existing_table.get("has_callouts", False)
            current_callouts = st.session_state["selections"].get(tid, {}).get("callouts", [])
            total_callout_count = len(existing_callouts) + len(current_callouts)
            
            # Determine if callouts should be enabled by default
            # Enable if there are existing callouts or if user has previously enabled them
            default_enabled = (total_callout_count > 0 or 
                             st.session_state.get(f"enable_callouts_{tid}", False))
            
            enable_callouts = False
            if has_callouts_obj:
                st.write("**Callouts**")
                # Toggle checkbox for callouts
                toggle_label = f"Enable callouts for this table"
                if total_callout_count > 0:
                    toggle_label += f" ({total_callout_count} active)"
                enable_callouts = st.checkbox(toggle_label, value=default_enabled, key=f"enable_callouts_{tid}")
            
            if not enable_callouts:
                # Clear any existing callouts when disabled
                if "callouts" in st.session_state["selections"].get(tid, {}):
                    st.session_state["selections"][tid]["callouts"] = []
            
            if has_callouts_obj and enable_callouts:
                # Only show callout controls when enabled
                if total_callout_count == 0:
                    cc1, cc2, cc3 = st.columns([2, 1, 1])
                    
                    with cc1:
                        # Row selection for callouts
                        available_rows = [row for row in t["row_labels"] if isinstance(row, str) and row.strip()]
                        selected_row = st.selectbox("Select row for callout", available_rows, key=f"callout_row_{tid}")
                    
                    with cc2:
                        # Callout selectors: Banner only to users, Metric shown only if multiple
                        callout_metric = current_metric
                        if has_metrics:
                            callout_metric = st.selectbox("Callout metric", metric_options, key=f"callout_metric_{tid}", index=metric_options.index(current_metric) if current_metric in metric_options else 0)
                        callout_banner_default = selected_col if selected_col in banners else (banners[0] if banners else "")
                        callout_banner = st.selectbox("Callout banner", banners, key=f"callout_banner_{tid}", index=banners.index(callout_banner_default) if callout_banner_default in banners else 0)
                        # Metric type selector (display capitalized, store lowercase)
                        metric_type_display = ["Percentage", "Number", "Currency"]
                        default_metric_type = st.session_state["selections"].get(tid, {}).get("callout_metric_type", "percentage")
                        def_idx = metric_type_display.index(default_metric_type.capitalize()) if default_metric_type and default_metric_type.capitalize() in metric_type_display else 0
                        selected_metric_display = st.selectbox("Callout metric type", metric_type_display, key=f"callout_metric_type_{tid}", index=def_idx)
                        callout_metric_type = selected_metric_display.lower()
                        # Resolve to combined column label internally
                        callout_col = pair_to_combined.get(((callout_metric or ""), callout_banner))
                        if not callout_col:
                            # Fallback to first occurrence of the banner
                            for i, b in enumerate(banners):
                                if b == callout_banner:
                                    callout_col = combined_labels[i]
                                    break
                        if not callout_col and combined_labels:
                            callout_col = combined_labels[0]
                    
                    with cc3:
                        if st.button("Add Callout", key=f"add_callout_{tid}"):
                            if "callouts" not in st.session_state["selections"][tid]:
                                st.session_state["selections"][tid]["callouts"] = []
                            
                            # Create new callout with default text format
                            default_text = f"{selected_row}: [Value]"
                            new_callout = {
                                "row_label": selected_row,
                                "column_key": callout_col,
                                "text": default_text,  # Start with default "Row Name: Value" format
                                "position": (0.5, 7.0, 9.0, 0.4),
                                "font_size": 12,
                                "font_bold": True,
                                "metric_type": callout_metric_type
                            }
                            
                            st.session_state["selections"][tid]["callouts"].append(new_callout)
                            st.success(f"Added callout for '{selected_row}'")
                            st.rerun()  # Refresh to update the UI
                else:
                    st.info("📝 Callouts already exist for this table")
                
                # Display existing callouts with editable text boxes
                
                # Show existing callouts from PowerPoint with "Previously:" indicator
                if has_callouts_obj and (existing_table.get("callouts", []) ):
                    st.write("**Existing Callouts from PowerPoint:**")
                    ecs = existing_table.get("callouts", [])
                    for i, existing_callout in enumerate(ecs):
                        callout_display_cols = st.columns([4, 1, 1, 1])
                        
                        with callout_display_cols[0]:
                            # Show the callout info with "Previously:" indicator
                            st.write(f"• {existing_callout['row_label']} → {existing_callout['column_key']}")
                            
                            # Editable text box showing existing callout text with "Previously:" label
                            callout_label = "Callout text (Previously: " + existing_callout['text'] + ")"
                            updated_text = st.text_input(
                                callout_label,
                                value=existing_callout['text'],
                                key=f"existing_callout_text_{tid}_{i}",
                                help="Customize your callout text. Use [Value] as a placeholder for the actual data value."
                            )
                            # Metric type dropdown for existing callout (display capitalized, store lowercase)
                            mt_display = ["Percentage", "Number", "Currency"]
                            mt_current = existing_callout.get('metric_type', 'percentage')
                            mt_idx = mt_display.index(mt_current.capitalize()) if mt_current and mt_current.capitalize() in mt_display else 0
                            mt_updated_display = st.selectbox("Metric type", mt_display, index=mt_idx, key=f"existing_callout_metric_{tid}_{i}")
                            mt_updated = mt_updated_display.lower()
                            
                            # Update the existing callout text in real-time
                            if updated_text != existing_callout['text']:
                                ecs[i]['text'] = updated_text
                            if mt_updated != existing_callout.get('metric_type'):
                                ecs[i]['metric_type'] = mt_updated
                        
                        with callout_display_cols[1]:
                            if st.button("Remove", key=f"remove_existing_callout_{tid}_{i}"):
                                ecs.pop(i)
                                st.rerun()
                        
                        with callout_display_cols[2]:
                            if st.button("Edit", key=f"edit_existing_callout_{tid}_{i}"):
                                st.session_state["editing_existing_callout"] = (tid, i)
                                st.rerun()
                        
                        # Add some spacing between callouts
                        st.divider()
                
                # Show current callouts from session state
                if current_callouts:
                    st.write("**Current Callouts:**")
                    for i, callout in enumerate(current_callouts):
                        callout_display_cols = st.columns([4, 1, 1, 1])
                        
                        with callout_display_cols[0]:
                            # Show the callout info and editable text box
                            st.write(f"• {callout['row_label']} → {callout['column_key']}")
                            
                            # Editable text box for customizing callout text
                            current_text = callout.get('text', f"{callout['row_label']}: [Value]")
                            updated_text = st.text_input(
                                "Callout text (customize as needed):",
                                value=current_text,
                                key=f"callout_text_{tid}_{i}",
                                help="Customize your callout text. Use [Value] as a placeholder for the actual data value."
                            )
                            mt_display = ["Percentage", "Number", "Currency"]
                            mt_current = callout.get('metric_type', 'percentage')
                            mt_idx = mt_display.index(mt_current.capitalize()) if mt_current and mt_current.capitalize() in mt_display else 0
                            mt_updated_display = st.selectbox("Metric type", mt_display, index=mt_idx, key=f"callout_metric_{tid}_{i}")
                            mt_updated = mt_updated_display.lower()
                            
                            # Update the callout text in real-time
                            if updated_text != current_text:
                                st.session_state["selections"][tid]["callouts"][i]["text"] = updated_text
                            if mt_updated != mt_current:
                                st.session_state["selections"][tid]["callouts"][i]["metric_type"] = mt_updated
                        
                        with callout_display_cols[1]:
                            if st.button("Remove", key=f"remove_callout_{tid}_{i}"):
                                st.session_state["selections"][tid]["callouts"].pop(i)
                                st.rerun()
                        
                        with callout_display_cols[2]:
                            if st.button("Edit", key=f"edit_callout_{tid}_{i}"):
                                st.session_state["editing_callout"] = (tid, i)
                                st.rerun()
                        
                        # Add some spacing between callouts
                        st.divider()

            # Build selections dictionary
            selection_dict = {
                "chart_type_label": choice,
                "chart_type": {
                    "Bar Horizontal": "bar_h",
                    "Bar Vertical":   "bar_v",
                    "Donut":          "donut",
                    "Line":           "line",
                    "Chart + Table":  "chart+table"
                }[choice],
                # Persist both banner and metric plus the resolved combined label for export
                "banner_key": selected_col,
                "metric_key": current_metric,
                "column_key": combined_selected,
                "enable_sorting": enable_sorting,
                "excluded_rows": excluded_rows if enable_sorting else []
            }

            # Only include text fields if those objects are linked in PPT
            if title_val is not None:
                selection_dict["title"] = title_val
            if base_text_val is not None:
                selection_dict["base_text"] = base_text_val
            if question_text_val is not None:
                selection_dict["question_text"] = question_text_val
            
            # Include only user-created callouts in selections; do NOT merge in existing PPT callouts
            if enable_callouts:
                user_callouts = st.session_state["selections"].get(tid, {}).get("callouts", [])
                if user_callouts:
                    selection_dict["callouts"] = user_callouts
            
            st.session_state["selections"][tid] = selection_dict

    # Export/Update section
    if st.session_state.data is not None:
        st.divider()
        
        if workflow_type == "Create New Report":
            # NEW REPORT EXPORT
            st.subheader("Step 4: Export new PowerPoint")
            if st.button("Export PowerPoint", type="primary"):
                sels = {tid: {
                    "chart_type": v["chart_type"], 
                    "column_key": v.get("column_key", "Total"), 
                    "title": v["title"], 
                    "base_text": v.get("base_text"), 
                    "question_text": v.get("question_text"), 
                    "callouts": v.get("callouts", []),
                    "enable_sorting": v.get("enable_sorting", False),
                    "excluded_rows": v.get("excluded_rows", [])
                } for tid, v in st.session_state["selections"].items()}
                out = "report.pptx"
                export_pptx(data["tables"], sels, out)
            with open(out, "rb") as f:
                st.download_button("Download report.pptx", f, file_name="report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        
        else:
            # UPDATE EXISTING REPORT
            st.subheader("Step 6: Update PowerPoint")
            
            # Show what will happen
            connected_count = len([t for t in data["tables"] if t["title"] in existing_content])
            unconnected_count = len([t for t in data["tables"] if t["title"] not in existing_content])
            
            st.info(f"""
            **Update Summary:**
            - {connected_count} existing tables will be updated with new data
            - {unconnected_count} new tables will be added to an "Unmapped Tables" summary page
            - Custom formatting and content will be preserved where possible
            """)
            
            if st.button("Update PowerPoint", type="primary"):
                # Build selections for new tables only
                new_table_sels = {}
                for tid, v in st.session_state["selections"].items():
                    new_table_sels[tid] = {
                        "chart_type": v["chart_type"],
                        "column_key": v.get("column_key", "Total"),
                        "title": v.get("title"),
                        "base_text": v.get("base_text"),
                        "question_text": v.get("question_text"),
                        "callouts": v.get("callouts", []),
                        "enable_sorting": v.get("enable_sorting", False),
                        "excluded_rows": v.get("excluded_rows", [])
                    }
                
                # Convert selections to use table titles as keys for proper matching
                table_selections = {}
                for tid, v in st.session_state["selections"].items():
                    # Find the table title for this tid
                    for t in data["tables"]:
                        if t["id"] == tid:
                            table_selections[t["title"]] = {
                                "chart_type": v["chart_type"],
                                "title": v.get("title"),
                                "base_text": v.get("base_text"),
                                "question_text": v.get("question_text")
                            }
                            break
                
                # Update presentation with enhanced functionality
                updated = update_presentation_with_unmapped("to_update.pptx", "uploaded.xlsx", "updated_report.pptx", 
                                                          table_selections, data["tables"], existing_content)
                
                with open(updated, "rb") as f:
                    st.download_button("Download updated_report.pptx", f, file_name="updated_report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# Show instructions when no data is loaded
if st.session_state.data is None:
    if workflow_type == "Create New Report":
        st.info("👆 Upload a crosstab Excel file to begin creating your report.")
