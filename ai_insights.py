"""
AI Insight Generator — Two-Tier Analysis
Calls the OpenAI API to produce two levels of insight for each crosstab table:
  1. takeaway — strategic, actionable headline (~15 words)
  2. analysis — 2-3 factual sentences with specific numbers (~60 words)

Usage:
    from ai_insights import generate_table_insights, generate_all_insights
"""

import hashlib
import json
import logging
import os
import time
from dotenv import load_dotenv
from typing import Dict, Any, List, Optional

from brand_config import AI_GENERATION

logger = logging.getLogger("report_relay.ai_insights")

load_dotenv()


# ---------------------------------------------------------------------------
# File-based insight cache
# ---------------------------------------------------------------------------

def _cache_dir() -> str:
    return AI_GENERATION.get("cache_dir", ".ai_cache")


def _cache_key(table: Dict[str, Any], column_key: str) -> str:
    """Deterministic hash of table data + column key for cache lookup."""
    payload = json.dumps(
        {
            "title": table.get("title", ""),
            "row_labels": table.get("row_labels", []),
            "col_labels": table.get("col_labels", []),
            "values": table.get("values", []),
            "column_key": column_key,
        },
        sort_keys=True,
        default=str,
    )
    return hashlib.sha256(payload.encode()).hexdigest()


def _read_cache(key: str) -> Optional[Dict[str, str]]:
    path = os.path.join(_cache_dir(), f"{key}.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {"takeaway": data.get("takeaway", ""), "analysis": data.get("analysis", "")}
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        return None


def _write_cache(key: str, insights: Dict[str, str], model: str, column_key: str) -> None:
    cdir = _cache_dir()
    os.makedirs(cdir, exist_ok=True)
    path = os.path.join(cdir, f"{key}.json")
    try:
        payload = {
            **insights,
            "_model": model,
            "_column_key": column_key,
            "_timestamp": time.time(),
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except OSError as e:
        logger.warning("Failed to write insight cache %s: %s", path, e)


# ---------------------------------------------------------------------------
# Core two-tier generator
# ---------------------------------------------------------------------------

def generate_table_insights(
    table: Dict[str, Any],
    column_key: str = "Total",
    use_cache: bool = True,
    old_table: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    """
    Generate two tiers of insight for a single crosstab table.

    Args:
        use_cache: When True, check the file cache before calling the API
            and persist new results.
        old_table: Optional previous-wave table dict.  When provided the
            prompt includes both old and new data so the LLM can highlight
            wave-over-wave changes.

    Returns:
        {"takeaway": str, "analysis": str}
        Falls back to empty strings on API failure.
    """
    empty = {"takeaway": "", "analysis": ""}

    # Skip cache when wave-over-wave context is supplied (results are unique)
    cache_hit_key = _cache_key(table, column_key) if (use_cache and not old_table) else None
    if cache_hit_key:
        cached = _read_cache(cache_hit_key)
        if cached:
            logger.debug("Cache hit for '%s' (%s)", table.get("title"), column_key)
            return cached

    try:
        from openai import OpenAI
    except ImportError:
        logger.error("openai package not installed. Run: pip install openai")
        return empty

    table_text = _format_table_for_prompt(table, column_key)
    if not table_text:
        return empty

    system_msg = AI_GENERATION.get("system_prompt", "")
    if old_table:
        system_msg += AI_GENERATION.get("system_prompt_wave", "")

    wave_section = ""
    if old_table:
        old_text = _format_table_for_prompt(old_table, column_key)
        if old_text:
            wave_section = (
                f"\n\n--- Previous wave data (same table) ---\n{old_text}\n"
                "--- End previous wave ---\n"
            )

    user_msg = f"""Below is data from a crosstab table titled "{table['title']}".
Column shown: {column_key}

{table_text}{wave_section}

Produce exactly two outputs as a JSON object with these keys:

1. "takeaway" — One strategic, actionable sentence (~15 words) aimed at destination marketers or tourism stakeholders. Frame it as clear guidance on what to do with this finding. Use a directive or declarative statement.

2. "analysis" — 2-3 sentences (~60 words) that go beyond reporting numbers. Include specific metrics (percentages, segment comparisons, changes). Interpret what the figures mean and why they matter for traveler behavior or destination strategy. Note any emerging patterns or demographic distinctions. Write with confident, clear language.

Return ONLY the JSON object, no other text."""

    model = os.getenv("AI_MODEL") or AI_GENERATION.get("model", "gpt-4o-mini")
    max_tokens = AI_GENERATION.get("max_tokens", 250)
    temperature = AI_GENERATION.get("temperature")

    try:
        client = OpenAI()
        api_kwargs: Dict[str, Any] = {
            "model": model,
            "max_tokens": max_tokens,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
        }
        if temperature is not None:
            api_kwargs["temperature"] = temperature
        response = client.chat.completions.create(**api_kwargs)
        result = json.loads(response.choices[0].message.content)
        insights = {
            "takeaway": result.get("takeaway", ""),
            "analysis": result.get("analysis", ""),
        }
        if cache_hit_key:
            _write_cache(cache_hit_key, insights, model, column_key)
        return insights
    except Exception as e:
        logger.error("AI insight generation failed for '%s': %s", table.get("title"), e)
        return empty


def generate_all_insights(
    tables: List[Dict[str, Any]],
    column_key: str = "Total",
    selections: Optional[Dict[str, Dict]] = None,
    use_cache: bool = True,
    old_tables: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Dict[str, str]]:
    """
    Generate two-tier insights for every table using parallel API calls.

    Args:
        old_tables: Optional dict keyed by table title mapping to previous-wave
            table dicts.  When supplied, the matching old table is forwarded to
            ``generate_table_insights`` for wave-over-wave commentary.

    Returns:
        Dict mapping table title -> {"takeaway", "analysis"}
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    insights: Dict[str, Dict[str, str]] = {}
    sel_by_id = selections or {}
    old_map = old_tables or {}

    tasks: List[tuple] = []
    for table in tables:
        col = column_key
        if sel_by_id:
            tid = table.get("id", "")
            sel = sel_by_id.get(tid, {})
            col = sel.get("column_key") or sel.get("banner_key") or column_key
        title = table.get("title", "")
        old_tbl = old_map.get(title)
        tasks.append((table, col, title, old_tbl))

    max_workers = AI_GENERATION.get("max_concurrent", 5)
    logger.info("Generating insights for %d tables (max_concurrent=%d)", len(tasks), max_workers)

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {
            pool.submit(generate_table_insights, tbl, col, use_cache, old_tbl): title
            for tbl, col, title, old_tbl in tasks
        }
        for future in as_completed(futures):
            title = futures[future]
            try:
                insights[title] = future.result()
            except Exception as e:
                logger.error("Insight generation failed for '%s': %s", title, e)
                insights[title] = {"takeaway": "", "analysis": ""}

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
    use_cache: bool = True,
) -> Dict[str, str]:
    """Legacy wrapper — returns {title: analysis_string}."""
    full = generate_all_insights(tables, column_key, selections, use_cache=use_cache)
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
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(name)s | %(message)s")
    result = generate_table_insights(test_table, "Total")
    logger.info("Generated insights:")
    for k, v in result.items():
        logger.info("  %s: %s", k, v)
