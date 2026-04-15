# Report Relay: Attack Plan for Pointed Improvements

## Priority 1: Replace Hard-Coded Alt-Text Matching with AI-Led Smart Match System

### The Problem
The current update system (`deck_update.py`) matches PowerPoint shapes to crosstab data using **exact string matching** on `table_title` stored in alt text. If the crosstab title changes even slightly between waves (e.g., `"Q3 Age"` becomes `"Q3. Age"`, or `"Q Age - Recode"` becomes `"Q Age (Recoded)"`), the match silently fails and the shape gets skipped. The team has to manually re-tag shapes or ensure titles never change, which is fragile and error-prone.

### The Fix: Three-Tier Intelligent Matching

Build a new `smart_match.py` module that replaces the current `_get_chart_mapping_from_shape` and `_get_table_mapping_from_shape` functions with a three-tier matching pipeline:

**Tier 1: Exact match** (current behavior, zero cost)
- Normalize whitespace/case, then compare `table_title` from alt text against crosstab table titles.
- If exact match found, use it. Done.

**Tier 2: Fuzzy structural match** (no AI, fast)
- When exact match fails, use a scoring algorithm that combines:
  - **Title similarity** (Levenshtein ratio or token overlap, e.g. `thefuzz` library) weighted at 40%
  - **Row label overlap** (Jaccard similarity of row label sets) weighted at 35%
  - **Column structure overlap** (Jaccard similarity of column label sets) weighted at 25%
- Accept matches above a configurable confidence threshold (default: 0.75).
- This catches renamed tables, reordered questions, minor wording changes.

**Tier 3: LLM-assisted match** (optional, for ambiguous cases)
- When Tier 2 produces multiple candidates above threshold or all candidates are below threshold, present the top 3 candidates to the LLM with a structured prompt:
  - "Given this shape was previously mapped to table titled X with rows [A, B, C] and columns [D, E], which of these new tables is the best match?"
- Cache match decisions per session so the LLM is only called once per ambiguous pairing.
- This is the "AI-led" match system that handles edge cases like completely renamed questions that still have the same row/column structure.

### Implementation Steps

1. **Create `smart_match.py`** with a `SmartMatcher` class:
   - `__init__(self, tables, confidence_threshold=0.75, use_llm=False)`
   - `match(self, alt_text_metadata) -> MatchResult` (returns table, confidence, tier_used)
   - `match_all(self, shapes_metadata) -> List[MatchResult]` (batch operation with deduplication)
   - Internal: `_exact_match()`, `_fuzzy_match()`, `_llm_match()`

2. **Add a match report** that gets surfaced in the Streamlit UI:
   - Show each shape's match result with confidence score and tier used.
   - Flag low-confidence matches (0.60-0.75) in yellow for user review.
   - Flag failed matches in red.
   - Let users override any match with a dropdown before running the update.

3. **Store richer metadata in alt text** during report creation:
   - In addition to `table_title`, store a hash of row labels and column labels.
   - Store the sheet name and block index as fallback identifiers.
   - This gives Tier 2 more signal to work with on future updates.

4. **Refactor `deck_update.py`** to use `SmartMatcher` instead of inline matching logic.

### Files to Change
- **New**: `smart_match.py`
- **Modify**: `deck_update.py` (replace `_get_chart_mapping_from_shape`, `_get_table_mapping_from_shape`)
- **Modify**: `pptx_exporter.py` (`_tag_shape` to write richer alt text metadata)
- **Modify**: `app.py` (add match review UI step in update workflow)

---

## Priority 2: Eliminate Massive Code Duplication in `deck_update.py`

### The Problem
`deck_update.py` is 1,462 lines and contains enormous amounts of copy-pasted logic:
- `update_presentation()` and `update_presentation_with_unmapped()` are ~90% identical. The second one just adds an unmapped tables summary slide.
- `_update_question_and_base()` and `_update_question_and_base_with_selections()` do nearly the same thing with slightly different input shapes.
- Base text parsing logic (extracting custom description from "Base: X. N complete surveys.") is duplicated in **4 places**: `app.py` lines 70-94, `deck_update.py` lines 676-700, `deck_update.py` lines 862-892, and `app.py` lines 570-589.
- The "preserve formatting by updating only first run" pattern is copy-pasted ~15 times.

### The Fix

1. **Merge the two `update_presentation` functions** into one, with an `include_unmapped_summary=True` parameter.

2. **Extract a `BaseTextParser` utility** that handles all base text parsing/formatting in one place:
   - `parse(text) -> {"description": str, "n_value": int | None}`
   - `format(description, n_value) -> str`

3. **Extract a `_safe_update_text(shape, new_text)` helper** that encapsulates the "preserve formatting, update first run, clear extra runs" pattern. Replace all 15+ instances.

4. **Consolidate question/base/title update logic** into a single `_update_text_shapes(slide, table, selections)` function that handles all text shape types in one pass.

### Files to Change
- **New**: `text_utils.py` (BaseTextParser, safe_update_text)
- **Modify**: `deck_update.py` (consolidate functions)
- **Modify**: `app.py` (use BaseTextParser)

---

## Priority 3: Fix the Broken Selections Passthrough in Update Workflow

### The Problem
The update workflow has a critical bug in how selections are passed to the update engine. In `app.py` (lines 946-958), selections are converted from table ID keys to table title keys, but only `chart_type`, `title`, `base_text`, and `question_text` are included. **`column_key` is not passed**, which means the update engine can't switch data columns during an update. The `column_key` the user selects in Step 5 of the update workflow is effectively ignored.

Additionally, the `_update_question_and_base_with_selections()` function receives selections wrapped in an extra dict layer (`{table_title: table_selection}`) but the calling code in `update_presentation` (line 1144) already passes it this way, creating confusing nesting.

### The Fix

1. **Include `column_key` in the table_selections dict** passed to the update engine.
2. **Standardize the selections dict shape** - always use `{table_title: {key: value}}` and document it.
3. **Pass `column_key` through to `_update_chart`** when available in selections, overriding the alt-text column.
4. **Add callout data** to the passed selections so callout updates also reflect column changes.

### Files to Change
- **Modify**: `app.py` (lines 946-958, include column_key and callouts)
- **Modify**: `deck_update.py` (simplify selection handling, use column_key from selections)

---

## Priority 4: Harden the Crosstab Parser Against Edge Cases

### The Problem
`crosstab_parser.py` uses heuristics that break on non-standard crosstabs:
- The 20% threshold for footnote detection (line 142) is too aggressive. A table with 10 columns where a row only has data in 1 (e.g., a "Total" only row) gets incorrectly classified as a footnote.
- The block detection minimum of 10 non-null cells (line 16) misses small but valid tables.
- Title detection looks at rows above the header but doesn't handle cases where the title IS the first row of the block (common in some Q-style exports).
- There's no handling of merged cells in the Excel file beyond forward-filling the metric row.

### The Fix

1. **Make the footnote threshold configurable** and raise the default to 0.10 (10%).
2. **Add a "known footnote patterns" kill list** that's more aggressive than the percentage check, so the percentage check can be more permissive.
3. **Lower the minimum block size** to 4 non-null cells (2x2 minimum).
4. **Add title detection fallback**: if no title is found above the header, check whether the first row of the block itself looks like a title (single non-empty cell in column A, no data in other columns).
5. **Add a `parse_options` dict parameter** to `parse_workbook()` so downstream callers can tune thresholds without modifying the module.

### Files to Change
- **Modify**: `crosstab_parser.py`

---

## Priority 5: Add Error Handling, Logging, and a Dry-Run Mode

### The Problem
- The codebase uses bare `except Exception` everywhere (deck_update.py alone has 40+ except blocks). Errors are silently swallowed, making debugging nearly impossible.
- Debug output is scattered `print()` statements with emoji prefixes. There's no structured logging.
- There's no way to preview what an update will do without actually doing it. If something goes wrong, you only find out after the file is saved.

### The Fix

1. **Replace `print()` with Python `logging` module**:
   - Use `logger = logging.getLogger("report_relay")` in each module.
   - Debug-level for shape-by-shape details, info-level for summary, warning-level for skipped shapes, error-level for failures.
   - Configure a StreamHandler for Streamlit console output and optionally a FileHandler for persistent logs.

2. **Narrow the except clauses**:
   - Replace `except Exception` with specific exceptions (`ValueError`, `KeyError`, `AttributeError`).
   - At minimum, log the exception type and message before continuing.
   - Add a `--strict` mode that raises instead of swallowing.

3. **Add a dry-run mode** to `update_presentation()`:
   - Walk all shapes and compute matches, but don't modify anything.
   - Return a structured report: `{matched: [...], unmatched: [...], warnings: [...]}`.
   - Surface this in the Streamlit UI as a preview step before the user clicks "Update."

### Files to Change
- **All Python files** (logging migration)
- **Modify**: `deck_update.py` (dry-run mode)
- **Modify**: `app.py` (preview step UI)

---

## Priority 6: Decouple Chart Formatting Preservation from Data Update

### The Problem
`_update_chart()` in `deck_update.py` (lines 131-420) is a 290-line monster that tries to:
1. Read current chart formatting (80 lines of try/except)
2. Replace chart data (5 lines)
3. Restore chart formatting (130 lines of try/except)
4. Apply fallback percentage formatting (another 50 lines)

The formatting preservation is extremely fragile. `chart.replace_data()` from python-pptx wipes formatting, and the restoration code tries to re-apply it field by field. But it misses many properties (fill colors, line styles, gradient fills, shadow effects), and the fallback logic for percentage formatting can override intentional non-percentage formatting.

### The Fix

1. **Extract chart XML before data replacement, patch data into it, and write it back**:
   - Instead of using `chart.replace_data()` (which rebuilds the chart XML from scratch), directly modify the `c:val` and `c:cat` references in the existing chart XML.
   - This preserves 100% of formatting because you're only touching the data nodes.
   - This is the same approach `_apply_crtx_template` in `pptx_exporter.py` already uses (graft data refs into template XML).

2. **Create a `chart_data_patcher.py` module** with:
   - `patch_chart_data(chart_part, categories, values)` - modifies data in-place
   - `patch_chart_series(chart_part, series_data)` - for multi-series updates

3. **Move the percentage detection heuristic** to a separate, testable function and make it overridable via alt-text metadata (e.g., `value_format: percentage` or `value_format: number`).

### Files to Change
- **New**: `chart_data_patcher.py`
- **Modify**: `deck_update.py` (replace `_update_chart` internals)
- **Modify**: `pptx_exporter.py` (add `value_format` to alt text tags)

---

## Priority 7: Improve the Streamlit UI for Update Workflow

### The Problem
- The update workflow crams everything into a single scrolling page. For a report with 30+ tables, configuring each one is overwhelming.
- Connected tables are listed as plain text bullets with no useful information about what changed.
- There's no diff view showing old values vs new values.
- Column selection applies globally but there's no way to set different columns per table in the update workflow.
- The app writes uploaded files to the working directory with hardcoded names ("uploaded.xlsx", "to_update.pptx"), which means concurrent users would overwrite each other's files.

### The Fix

1. **Add a match review step** (ties to Priority 1) showing a table with columns: Shape Type | Old Title | Matched Title | Confidence | Action (dropdown: Update / Skip / Manual).

2. **Add a data diff preview** for connected tables: show a side-by-side or highlight of which values changed, which rows were added/removed.

3. **Use `tempfile.NamedTemporaryFile`** instead of hardcoded filenames for uploads.

4. **Allow per-table column override** in the update workflow, not just global.

5. **Add a progress bar** for the update operation (it can be slow with many tables + AI insights).

### Files to Change
- **Modify**: `app.py` (UI restructure, temp files, progress)

---

## Priority 8: Make AI Insights More Robust and Configurable

### The Problem
- AI insights are hardcoded to `gpt-4o-mini` with no way to switch models.
- The system prompt is baked into the code, not configurable.
- There's no caching. If you re-export the same data, you pay for API calls again.
- The insight generator doesn't receive context about what changed between waves, so it can't highlight trends.
- Insights are generated sequentially. For 30 tables, this adds significant latency.

### The Fix

1. **Make the model configurable** via environment variable or `brand_config.py`.
2. **Add a simple file-based cache**: hash the table data + column key, store insights in a JSON file. Check cache before calling API.
3. **Add wave-over-wave context**: when updating, pass both old and new data to the insight generator so it can say "increased from X% to Y%."
4. **Use `asyncio` or `concurrent.futures`** to parallelize API calls (batch of 5-10 at a time).
5. **Move the system prompt to a configurable template** in `brand_config.py` or a separate `prompts/` directory.

### Files to Change
- **Modify**: `ai_insights.py` (caching, parallelism, model config)
- **Modify**: `brand_config.py` (AI config section)
- **New**: `prompts/insight_system.txt` (externalized prompt)

---

## Priority 9: Security Fix in `mapping_helper.py`

### The Problem
`apply_mapping_from_file()` on line 146 uses `exec(f.read(), globals())` to load a Python mapping file. This executes arbitrary code. Anyone who shares a malicious mapping file can run any code on the machine.

### The Fix
Replace `exec()` with a safe loader:
- Use `ast.literal_eval()` if the file only contains a dict literal.
- Or switch to JSON/YAML format for mapping files.
- Or use `importlib` with a restricted namespace.

### Files to Change
- **Modify**: `mapping_helper.py`

---

## Priority 10: Add Tests

### The Problem
There are zero tests. Every change risks breaking something else, and the only way to verify is manual end-to-end testing with real PowerPoint files.

### The Fix

1. **Unit tests for `crosstab_parser.py`**: Create sample Excel files with known structures and verify parsed output.
2. **Unit tests for `smart_match.py`** (new module): Test each tier with known inputs.
3. **Unit tests for `BaseTextParser`** (new utility): Test parsing and formatting edge cases.
4. **Integration test for the update pipeline**: Create a known PPTX, update it with known data, verify shapes contain expected values.
5. **Snapshot tests for chart formatting**: Export a chart, update it, compare XML before/after to ensure formatting preservation.

### Files to Change
- **New**: `tests/` directory with test files

---

## Execution Order

| Phase | Priorities | Estimated Effort | Impact |
|-------|-----------|-----------------|--------|
| **Phase 1: Foundation** | P2 (deduplication), P5 (logging), P9 (security) | 2-3 days | Reduces codebase by ~400 lines, makes everything else easier |
| **Phase 2: Core Match System** | P1 (smart match), P3 (selections fix) | 3-4 days | Biggest user-facing improvement: updates actually work reliably |
| **Phase 3: Data Integrity** | P4 (parser hardening), P6 (chart patching) | 2-3 days | Fewer silent data errors, better formatting preservation |
| **Phase 4: UX + Polish** | P7 (UI improvements), P8 (AI improvements) | 2-3 days | Better team experience, faster workflows |
| **Phase 5: Safety Net** | P10 (tests) | 2-3 days | Prevents regressions, enables confident iteration |
