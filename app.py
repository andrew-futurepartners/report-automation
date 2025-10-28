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
                
                if alt.get("type") in ["chart", "table", "question_text", "text_base", "text_title", "text_callout"]:
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
                                "callouts": []  # Initialize callouts list
                            }
                        
                        # Extract question text
                        if alt.get("type") == "question_text" and hasattr(shape, "text_frame"):
                            question_text = shape.text_frame.text
                            if question_text.startswith("Question: "):
                                question_text = question_text[10:]  # Remove "Question: " prefix
                            existing_content[table_title]["question_text"] = question_text
                            existing_content[table_title]["custom_question"] = question_text
                        
                        # Extract base text
                        elif alt.get("type") == "text_base" and hasattr(shape, "text_frame"):
                            base_text = shape.text_frame.text
                            existing_content[table_title]["base_text"] = base_text
                            
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
                        
                        # Extract text callouts
                        elif alt.get("type") == "text_callout" and hasattr(shape, "text_frame"):
                            if "callouts" not in existing_content[table_title]:
                                existing_content[table_title]["callouts"] = []
                            
                            callout_text = shape.text_frame.text
                            callout_info = {
                                "row_label": alt.get("row_label", ""),
                                "column_key": alt.get("column", "Total"),
                                "text": callout_text,
                                "position": (0.5, 7.0, 9.0, 0.4),  # Default position
                                "font_size": 12,
                                "font_bold": True
                            }
                            existing_content[table_title]["callouts"].append(callout_info)
        
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
        
        # Get all unique column labels across all tables for column selection
        all_columns = set()
        for table in data["tables"]:
            all_columns.update(table.get("col_labels", []))
        all_columns = sorted(list(all_columns))
        
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
                # Get all unique column labels across all tables for column selection
                all_columns = set()
                for table in data["tables"]:
                    all_columns.update(table.get("col_labels", []))
                all_columns = sorted(list(all_columns))
                
                st.subheader("Step 5: Configure column selection")
                st.info("💡 Quickly switch to different segments (e.g., Male, Female, Gen Z) for all connected tables")
                
                col1, col2 = st.columns(2)
                with col1:
                    default_column = st.selectbox("Default column for updates", all_columns, 
                                                index=all_columns.index("Total") if "Total" in all_columns else 0,
                                                key="update_default_column")
                    apply_column_btn = st.button("Apply column to all connected tables", key="update_apply_column")
                
                with col2:
                    st.write("**Available segments:**")
                    # Show first 8 columns as examples
                    segment_preview = all_columns[:8]
                    if len(all_columns) > 8:
                        segment_preview.append(f"... and {len(all_columns) - 8} more")
                    for col in segment_preview:
                        st.write(f"• {col}")
            
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
            with cols[0]:
                st.write("**Row labels**")
                st.write(", ".join(t["row_labels"][:12]) + ("..." if len(t["row_labels"]) > 12 else ""))
                st.write("**Columns**")
                st.write(", ".join(t["col_labels"]))

            with cols[1]:
                options = ["Bar Horizontal", "Bar Vertical", "Donut", "Line", "Chart + Table"]
                if apply_all_btn:
                    choice = default_choice
                else:
                    choice = st.session_state["selections"].get(tid, {}).get("chart_type_label", default_choice)
                choice = st.selectbox("Chart type", options, key=f"ctype_{tid}", index=options.index(choice))
                
                # Column selection for this table
                col_options = t["col_labels"]
                
                # Handle column application for both new reports and updates
                if workflow_type == "Create New Report" and apply_column_btn:
                    # Apply default column to all tables for new reports
                    current_col = default_column if default_column in col_options else ("Total" if "Total" in col_options else col_options[0] if col_options else "Total")
                elif workflow_type == "Update Existing Report" and apply_column_btn:
                    # Apply default column to all connected tables for updates
                    current_col = default_column if default_column in col_options else ("Total" if "Total" in col_options else col_options[0] if col_options else "Total")
                else:
                    # Use existing selection or default
                    current_col = st.session_state["selections"].get(tid, {}).get("column_key", "Total")
                    if current_col not in col_options:
                        current_col = "Total" if "Total" in col_options else col_options[0] if col_options else "Total"
                
                selected_col = st.selectbox("Data column", col_options, 
                                         index=col_options.index(current_col) if current_col in col_options else 0,
                                         key=f"col_{tid}")
            with cols[2]:
                # Use existing content if available, otherwise use defaults
                existing_table = existing_content.get(t["title"], {})
                
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
                
                title_val = st.text_input(title_label, value=default_title, key=f"title_{tid}")

                # Base text logic - preserve custom descriptions while updating N values
                def _find_base_idx(labels):
                    for i, lab in enumerate(labels):
                        if isinstance(lab, str) and lab.strip().lower().startswith("base"):
                            return i
                    return None
                
                base_idx = _find_base_idx(t["row_labels"])
                total_idx = t["col_labels"].index("Total") if "Total" in t["col_labels"] else (0 if t["col_labels"] else None)
                
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
                
                question_text_val = st.text_input(question_label, value=default_q, key=f"qtext_{tid}")

            # Callout management section
            st.write("**Callouts**")
            
            # Check if there are existing callouts from PowerPoint or session state
            existing_table = existing_content.get(t["title"], {})
            existing_callouts = existing_table.get("callouts", []) if existing_table else []
            current_callouts = st.session_state["selections"].get(tid, {}).get("callouts", [])
            total_callout_count = len(existing_callouts) + len(current_callouts)
            
            # Determine if callouts should be enabled by default
            # Enable if there are existing callouts or if user has previously enabled them
            default_enabled = (total_callout_count > 0 or 
                             st.session_state.get(f"enable_callouts_{tid}", False))
            
            # Toggle checkbox for callouts
            toggle_label = f"Enable callouts for this table"
            if total_callout_count > 0:
                toggle_label += f" ({total_callout_count} active)"
            
            enable_callouts = st.checkbox(toggle_label, value=default_enabled, key=f"enable_callouts_{tid}")
            
            if enable_callouts:
                st.info("💡 Select a row and column, then click 'Add Callout' to create text callouts")
            else:
                # Clear any existing callouts when disabled
                if "callouts" in st.session_state["selections"].get(tid, {}):
                    st.session_state["selections"][tid]["callouts"] = []
            
            if enable_callouts:
                # Only show callout controls when enabled
                callout_cols = st.columns([2, 1, 1])
                
                with callout_cols[0]:
                    # Row selection for callouts
                    available_rows = [row for row in t["row_labels"] if isinstance(row, str) and row.strip()]
                    selected_row = st.selectbox("Select row for callout", available_rows, key=f"callout_row_{tid}")
                
                with callout_cols[1]:
                    # Column selection for callout (defaults to selected column for chart)
                    callout_col = st.selectbox("Callout column", col_options, 
                                             index=col_options.index(selected_col) if selected_col in col_options else 0,
                                             key=f"callout_col_{tid}")
                
                with callout_cols[2]:
                    # Only show add callout button if no callouts exist yet
                    if total_callout_count == 0:
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
                                "font_bold": True
                            }
                            
                            st.session_state["selections"][tid]["callouts"].append(new_callout)
                            st.success(f"Added callout for '{selected_row}'")
                            st.rerun()  # Refresh to update the UI
                    else:
                        st.info("📝 Callouts already exist for this table")
                
                # Display existing callouts with editable text boxes
                
                # Show existing callouts from PowerPoint with "Previously:" indicator
                if existing_callouts:
                    st.write("**Existing Callouts from PowerPoint:**")
                    for i, existing_callout in enumerate(existing_callouts):
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
                            
                            # Update the existing callout text in real-time
                            if updated_text != existing_callout['text']:
                                existing_callouts[i]['text'] = updated_text
                        
                        with callout_display_cols[1]:
                            if st.button("Remove", key=f"remove_existing_callout_{tid}_{i}"):
                                existing_callouts.pop(i)
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
                            
                            # Update the callout text in real-time
                            if updated_text != current_text:
                                st.session_state["selections"][tid]["callouts"][i]["text"] = updated_text
                        
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
                "column_key": selected_col,  # Add selected column to selections
                "title": title_val,
                "base_text": base_text_val,
                "question_text": question_text_val
            }
            
            # Include callouts if the checkbox is enabled
            if enable_callouts:
                # Combine existing callouts from PowerPoint with new callouts from session state
                all_callouts = []
                
                # Add existing callouts from PowerPoint (if any)
                if existing_table and "callouts" in existing_table:
                    all_callouts.extend(existing_table["callouts"])
                
                # Add new callouts from session state (if any)
                if "callouts" in st.session_state["selections"].get(tid, {}):
                    all_callouts.extend(st.session_state["selections"][tid]["callouts"])
                
                if all_callouts:
                    selection_dict["callouts"] = all_callouts
            
            st.session_state["selections"][tid] = selection_dict

    # Export/Update section
    if st.session_state.data is not None:
        st.divider()
        
        if workflow_type == "Create New Report":
            # NEW REPORT EXPORT
            st.subheader("Step 4: Export new PowerPoint")
            if st.button("Export PowerPoint", type="primary"):
                sels = {tid: {"chart_type": v["chart_type"], "column_key": v.get("column_key", "Total"), "title": v["title"], "base_text": v.get("base_text"), "question_text": v.get("question_text"), "callouts": v.get("callouts", [])}
                        for tid, v in st.session_state["selections"].items()}
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
                        "title": v["title"],
                        "base_text": v.get("base_text"),
                        "question_text": v.get("question_text"),
                        "callouts": v.get("callouts", [])
                    }
                
                # Convert selections to use table titles as keys for proper matching
                table_selections = {}
                for tid, v in st.session_state["selections"].items():
                    # Find the table title for this tid
                    for t in data["tables"]:
                        if t["id"] == tid:
                            table_selections[t["title"]] = {
                                "chart_type": v["chart_type"],
                                "title": v["title"],
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
else:
        st.info("👆 Upload an existing PowerPoint file to begin the update process.")
