"""
AI Insight Generator — Two-Tier Analysis
Calls the OpenAI API to produce two levels of insight for each crosstab table:
  1. takeaway — strategic, actionable headline (~15 words)
  2. analysis — 2-3 factual sentences with specific numbers (~60 words)

Usage:
    from ai_insights import generate_table_insights, generate_all_insights
"""

import json
import os
from dotenv import load_dotenv
from typing import Dict, Any, List, Optional

load_dotenv()


# ---------------------------------------------------------------------------
# Core two-tier generator
# ---------------------------------------------------------------------------

def generate_table_insights(table: Dict[str, Any], column_key: str = "Total") -> Dict[str, str]:
    """
    Generate two tiers of insight for a single crosstab table.

    Returns:
        {"takeaway": str, "analysis": str}
        Falls back to empty strings on API failure.
    """
    empty = {"takeaway": "", "analysis": ""}

    try:
        from openai import OpenAI
    except ImportError:
        print("openai package not installed. Run: pip install openai")
        return empty

    table_text = _format_table_for_prompt(table, column_key)
    if not table_text:
        return empty

    system_msg = (
        "You are a senior research analyst at a creative insights firm specializing in "
        "travel, tourism, and hospitality. You write with the confident, clear voice of "
        "an experienced analyst — professional yet conversational. "
        "Ground every statement in specific data: percentages, averages, and segment "
        "differences. Interpret the numbers — explain what they mean and why they matter "
        "for destination marketers and tourism stakeholders. "
        "Write in present tense. Use active language. "
        "Never use filler phrases like 'the data shows', 'it is clear that', or "
        "'interestingly'. Never mention that you are an AI. "
        "Avoid jargon. Highlight differences by segment where meaningful."
    )

    user_msg = f"""Below is data from a crosstab table titled "{table['title']}".
Column shown: {column_key}

{table_text}

Produce exactly two outputs as a JSON object with these keys:

1. "takeaway" — One strategic, actionable sentence (~15 words) aimed at destination marketers or tourism stakeholders. Frame it as clear guidance on what to do with this finding. Use a directive or declarative statement.

2. "analysis" — 2-3 sentences (~60 words) that go beyond reporting numbers. Include specific metrics (percentages, segment comparisons, changes). Interpret what the figures mean and why they matter for traveler behavior or destination strategy. Note any emerging patterns or demographic distinctions. Write with confident, clear language.

Return ONLY the JSON object, no other text."""

    try:
        client = OpenAI()
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            max_tokens=250,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
        )
        result = json.loads(response.choices[0].message.content)
        return {
            "takeaway": result.get("takeaway", ""),
            "analysis": result.get("analysis", ""),
        }
    except Exception as e:
        print(f"Warning: AI insight generation failed for '{table.get('title')}': {e}")
        return empty


def generate_all_insights(
    tables: List[Dict[str, Any]],
    column_key: str = "Total",
    selections: Optional[Dict[str, Dict]] = None,
) -> Dict[str, Dict[str, str]]:
    """
    Generate two-tier insights for every table.

    Returns:
        Dict mapping table title -> {"takeaway", "analysis"}
    """
    insights: Dict[str, Dict[str, str]] = {}
    sel_by_id = selections or {}

    for table in tables:
        col = column_key
        if sel_by_id:
            tid = table.get("id", "")
            sel = sel_by_id.get(tid, {})
            col = sel.get("column_key") or sel.get("banner_key") or column_key

        title = table.get("title", "")
        print(f"Generating insights for: {title} ({col})")
        insights[title] = generate_table_insights(table, col)

    return insights


# ---------------------------------------------------------------------------
# Backward-compatible wrapper
# ---------------------------------------------------------------------------

def generate_table_summary(table: Dict[str, Any], column_key: str = "Total") -> str:
    """Legacy wrapper — returns only the analysis tier as a plain string."""
    result = generate_table_insights(table, column_key)
    return result.get("analysis", "")


def generate_all_summaries(
    tables: List[Dict[str, Any]],
    column_key: str = "Total",
    selections: Optional[Dict[str, Dict]] = None,
) -> Dict[str, str]:
    """Legacy wrapper — returns {title: analysis_string}."""
    full = generate_all_insights(tables, column_key, selections)
    return {title: tiers.get("analysis", "") for title, tiers in full.items()}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _format_table_for_prompt(table: Dict[str, Any], column_key: str) -> str:
    """Render the relevant column of a table as readable text for the prompt."""
    row_labels = table.get("row_labels", [])
    col_labels = table.get("col_labels", [])
    values     = table.get("values", [])

    if not row_labels or not col_labels or not values:
        return ""

    col_idx = None
    if column_key in col_labels:
        col_idx = col_labels.index(column_key)
    else:
        for fallback in ["Total", "Overall", "All"]:
            if fallback in col_labels:
                col_idx = col_labels.index(fallback)
                break
    if col_idx is None:
        col_idx = 0

    numeric_vals = []
    for i, row in enumerate(values):
        label = row_labels[i] if i < len(row_labels) else ""
        if _is_metadata_row(str(label)):
            continue
        v = row[col_idx] if col_idx < len(row) else None
        if v is not None:
            try:
                numeric_vals.append(float(v))
            except (TypeError, ValueError):
                pass

    is_proportion = (
        len(numeric_vals) > 0
        and all(0.0 <= v <= 1.01 for v in numeric_vals)
        and max(numeric_vals, default=0) <= 1.01
    )

    lines = [f"{'Row':<35} {column_key}"]
    lines.append("-" * 50)

    for i, label in enumerate(row_labels):
        if _is_metadata_row(str(label)):
            continue
        row = values[i] if i < len(values) else []
        v = row[col_idx] if col_idx < len(row) else None
        if v is None:
            formatted = "n/a"
        else:
            try:
                fv = float(v)
                if is_proportion:
                    formatted = f"{fv * 100:.1f}%"
                elif fv == int(fv) and fv > 100:
                    formatted = f"{int(fv):,}"
                else:
                    formatted = f"{fv:,.1f}"
            except (TypeError, ValueError):
                formatted = str(v)
        lines.append(f"{str(label):<35} {formatted}")

    base_line = _extract_base_n(table, col_idx)
    if base_line:
        lines.append(f"\n{base_line}")

    return "\n".join(lines)


def _is_metadata_row(label: str) -> bool:
    lower = label.strip().lower()
    return any(lower.startswith(p) for p in ("base", "mean", "avg", "average", "median"))


def _extract_base_n(table: Dict[str, Any], col_idx: int) -> str:
    for i, label in enumerate(table.get("row_labels", [])):
        if str(label).strip().lower().startswith("base"):
            row = table["values"][i] if i < len(table["values"]) else []
            v = row[col_idx] if col_idx < len(row) else None
            if v is not None:
                try:
                    return f"Base n = {int(round(float(v))):,}"
                except (TypeError, ValueError):
                    pass
    return ""


# ---------------------------------------------------------------------------
# Quick standalone test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    test_table = {
        "id": "test#1",
        "title": "Generation by Traveler Type",
        "row_labels": ["Gen Z", "Millennials", "Gen X", "Boomers+", "Base"],
        "col_labels": ["Total", "Hotel", "VFR", "Day-Tripper"],
        "values": [
            [0.146, 0.154, 0.153, 0.119],
            [0.321, 0.340, 0.296, 0.337],
            [0.244, 0.237, 0.228, 0.286],
            [0.289, 0.269, 0.323, 0.259],
            [1448,  548,   600,   300],
        ],
    }
    result = generate_table_insights(test_table, "Total")
    print("\nGenerated insights:")
    for k, v in result.items():
        print(f"  {k}: {v}")
