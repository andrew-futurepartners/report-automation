import streamlit as st
from crosstab_parser import parse_workbook
from pptx_exporter import export_pptx
from deck_update import update_presentation
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

st.set_page_config(page_title="Crosstab to PowerPoint", layout="wide")

st.title("Crosstab to PowerPoint")
st.write("Upload a Q-style crosstab Excel, pick chart types and titles, then export a branded PowerPoint.")

# Add helpful workflow information
with st.expander("💡 How to use this tool", expanded=False):
    st.write("""
    **New Report Workflow:**
    1. Upload crosstab Excel → Choose chart types and customize titles → Export new PowerPoint
    
    **Update Existing Report Workflow:**
    1. Upload crosstab Excel → Upload existing PowerPoint → Review what will be preserved/updated → Update PowerPoint
    
    **Key Benefits:**
    - Custom work is preserved when updating reports
    - Base N values are automatically refreshed with new data
    - Chart titles, questions, and base descriptions can be customized and preserved
    """)

uploaded = st.file_uploader("Upload crosstab Excel", type=["xlsx", "xls"])

default_choice = st.selectbox("Default visualization", ["Bar Horizontal", "Bar Vertical", "Donut", "Line", "Chart + Table"], index=0)
apply_all_btn = st.button("Apply default to all")

if uploaded:
    with st.spinner("Parsing workbook..."):
        with open("uploaded.xlsx", "wb") as f:
            f.write(uploaded.getbuffer())
        data = parse_workbook("uploaded.xlsx")

    st.success(f"Found {len(data['tables'])} tables")
    st.divider()

    # Choose action: Export new vs Update existing
    action = st.radio("What do you want to do?", ["Export new PowerPoint", "Update existing PowerPoint"], index=0)
    existing_ppt = None
    existing_content = {}
    
    if action == "Update existing PowerPoint":
        existing_ppt = st.file_uploader("Upload the PowerPoint to update", type=["pptx"], key="ppt_to_update")
        if existing_ppt:
            with st.spinner("Parsing existing PowerPoint..."):
                existing_content = parse_existing_powerpoint(existing_ppt)
            if existing_content:
                st.success(f"Found existing content for {len(existing_content)} tables")
                
            st.info("We will refresh tagged charts, tables, Question, and Base using the crosstab you just uploaded. **Custom content will be preserved where possible.**")
            
            # Add helpful information about what gets preserved
            st.info("""
            **What gets updated:** Chart data, table values, base N counts
            **What gets preserved:** Custom chart titles, custom question text, custom base descriptions (e.g., "Total respondents" vs "All participants")
            **What gets refreshed:** Base N values from the new crosstab data
            """)

    # Selections state
    if "selections" not in st.session_state:
        st.session_state["selections"] = {}

    # Per-table controls
    for t in data["tables"]:
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

    st.divider()
    if action == "Export new PowerPoint":
        if st.button("Export PowerPoint"):
            sels = {tid: {"chart_type": v["chart_type"], "column_key": v.get("column_key", "Total"), "title": v["title"], "base_text": v.get("base_text"), "question_text": v.get("question_text"), "callouts": v.get("callouts", [])}
                    for tid, v in st.session_state["selections"].items()}
            out = "report.pptx"
            export_pptx(data["tables"], sels, out)
            with open(out, "rb") as f:
                st.download_button("Download report.pptx", f, file_name="report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        if existing_ppt is None:
            st.info("Upload a PowerPoint to update.")
        else:
            if st.button("Update PowerPoint"):
                # For updates, we want to use the current text box values from the UI
                # This allows users to edit the text and have those edits reflected in the updated report
                # Build selections including existing callouts from PowerPoint
                sels = {}
                for tid, v in st.session_state["selections"].items():
                    table_selection = {
                        "chart_type": v["chart_type"],
                        "column_key": v.get("column_key", "Total"),
                        "title": v["title"],
                        "base_text": v.get("base_text"),
                        "question_text": v.get("question_text")
                    }
                    
                    # Include callouts if enabled
                    if st.session_state.get(f"enable_callouts_{tid}", False):
                        # Combine existing callouts from PowerPoint with new callouts
                        all_callouts = []
                        
                        # Add existing callouts from PowerPoint (if any)
                        existing_table = existing_content.get(tid, {})
                        if existing_table and "callouts" in existing_table:
                            all_callouts.extend(existing_table["callouts"])
                        
                        # Add new callouts from session state (if any)
                        if "callouts" in v:
                            all_callouts.extend(v["callouts"])
                        
                        if all_callouts:
                            table_selection["callouts"] = all_callouts
                    
                    sels[tid] = table_selection
                
                with open("to_update.pptx", "wb") as pf:
                    pf.write(existing_ppt.getbuffer())
                
                # Update the presentation with current selections (including user edits)
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
                
                print(f"DEBUG: Table selections: {table_selections}")
                
                # Debug: Print the exact text values being passed
                for table_title, selection in table_selections.items():
                    print(f"DEBUG: Table '{table_title}' selections:")
                    print(f"  - title: '{selection.get('title')}'")
                    print(f"  - question_text: '{selection.get('question_text')}'")
                    print(f"  - base_text: '{selection.get('base_text')}'")
                
                updated = update_presentation("to_update.pptx", "uploaded.xlsx", "updated_report.pptx", table_selections)
                with open(updated, "rb") as f:
                    st.download_button("Download updated_report.pptx", f, file_name="updated_report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
else:
    st.info("Upload a workbook to begin.")
