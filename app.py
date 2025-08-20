
import streamlit as st
from crosstab_parser import parse_workbook
from pptx_exporter import export_pptx
from deck_update import update_presentation

st.set_page_config(page_title="Crosstab to PowerPoint", layout="wide")

st.title("Crosstab to PowerPoint")
st.write("Upload a Q-style crosstab Excel, pick chart types and titles, then export a branded PowerPoint.")

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
    if action == "Update existing PowerPoint":
        existing_ppt = st.file_uploader("Upload the PowerPoint to update", type=["pptx"], key="ppt_to_update")
        st.info("We will refresh tagged charts, tables, Question, and Base using the crosstab you just uploaded.")

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
            with cols[2]:
                default_title = t["title"]
                title_val = st.text_input("Chart title", value=st.session_state["selections"].get(tid, {}).get("title", default_title), key=f"title_{tid}")

                # Base text default
                def _find_base_idx(labels):
                    for i, lab in enumerate(labels):
                        if isinstance(lab, str) and lab.strip().lower().startswith("base"):
                            return i
                    return None
                base_idx = _find_base_idx(t["row_labels"])
                total_idx = t["col_labels"].index("Total") if "Total" in t["col_labels"] else (0 if t["col_labels"] else None)
                default_base = st.session_state["selections"].get(tid, {}).get("base_text")
                if default_base is None:
                    if base_idx is not None and total_idx is not None and base_idx < len(t["values"]) and total_idx < len(t["values"][base_idx]):
                        try:
                            n_int = int(round(float(t["values"][base_idx][total_idx])))
                            default_base = f"Base: Total respondents. {n_int} complete surveys."
                        except Exception:
                            default_base = "Base: Total respondents."
                    else:
                        default_base = "Base: Total respondents."
                base_text_val = st.text_input("Base text", value=default_base, key=f"base_{tid}")

                # Question text
                default_q = st.session_state["selections"].get(tid, {}).get("question_text", t["title"])
                question_text_val = st.text_input("Question text", value=default_q, key=f"qtext_{tid}")

            st.session_state["selections"][tid] = {
                "chart_type_label": choice,
                "chart_type": {
                    "Bar Horizontal": "bar_h",
                    "Bar Vertical":   "bar_v",
                    "Donut":          "donut",
                    "Line":           "line",
                    "Chart + Table":  "chart+table"
                }[choice],
                "title": title_val,
                "base_text": base_text_val,
                "question_text": question_text_val
            }

    st.divider()
    if action == "Export new PowerPoint":
        if st.button("Export PowerPoint"):
            sels = {tid: {"chart_type": v["chart_type"], "title": v["title"], "base_text": v.get("base_text"), "question_text": v.get("question_text")}
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
                with open("to_update.pptx", "wb") as pf:
                    pf.write(existing_ppt.getbuffer())
                updated = update_presentation("to_update.pptx", "uploaded.xlsx", "updated_report.pptx")
                with open(updated, "rb") as f:
                    st.download_button("Download updated_report.pptx", f, file_name="updated_report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
else:
    st.info("Upload a workbook to begin.")
