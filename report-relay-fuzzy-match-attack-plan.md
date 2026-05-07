# Report Relay — AI Fuzzy Match Upgrade Attack Plan

**Date:** 2026-05-07
**Status:** Ready for review
**Goal:** Eliminate manual alt-text mapping by upgrading the auto-mapping engine to use value correlation, slide context, and fuzzy text matching to dynamically connect PowerPoint charts to crosstab tables.

---

## The Problem

Today's matching pipeline relies on alt-text metadata embedded in each PowerPoint shape. This works fine for decks that Report Relay *created* (where alt text is auto-generated), but breaks down for existing, manually-built reports. The SATS deck has 80 slides, 30+ charts, and zero alt-text. The crosstab has 277 tables across 4 sheets. Without alt text, the auto-mapper must figure out which chart maps to which table purely from the data itself.

The current `auto_mapper.py` structural scorer uses label overlap (Jaccard/containment on category and series names), which works for simple distribution charts but fails on segment breakout charts where categories are segments ("Total", "Gen Z"...) that appear in *every* table. These charts are distinguishable only by their **values** and **series names** (which are often abbreviated row labels).

---

## Key Data Findings

### Three chart patterns in the deck

| Pattern | Example | Categories | Series Name | How to Match |
|---|---|---|---|---|
| **Distribution** | Slide 12 | Response options | "%" (uninformative) | Categories ↔ row labels + value correlation |
| **Segment breakout** | Slide 13 | Segments (Total, Gen Z...) | "Better or Much Better Off" | Series name ↔ row label (fuzzy) + values across columns |
| **Time series** | Slide 14 | Dates | "Series 1" | Skip — not updatable from single-wave crosstab |

### Value correlation is the strongest signal

Even across different data waves, Pearson correlation between chart values and the correct crosstab column is **0.9942**. The absolute numbers differ (different wave), but the relative pattern is nearly identical. This means:

- For **distribution charts**: correlate chart values against each table's Total column values → the right table will have r ≈ 0.99
- For **segment breakout charts**: correlate chart values against a specific row's values across column banners → the right table+row combination will have r ≈ 0.99
- **No other table will come close** — random correlations among 277 tables will be well below 0.5

### Slide context provides disambiguation

Each chart slide has text boxes containing the question text and a descriptive title. These can be matched against crosstab table titles (which contain Q-codes and full question text). This is especially useful as a secondary signal when value correlation produces multiple high candidates.

---

## Architecture: Multi-Signal Scoring Engine

### Scoring signals (in priority order)

| Signal | Weight | Where it lives | What it does |
|---|---|---|---|
| **Value correlation** | 0.45 | New: `_value_correlation_score()` in `auto_mapper.py` | Pearson correlation of chart values vs crosstab values. Primary signal. |
| **Label overlap** | 0.25 | Existing: `score_fingerprint_against_tables()` | Jaccard/containment of categories ↔ row labels (already works). |
| **Slide text context** | 0.20 | New: `_slide_context_score()` in `auto_mapper.py` | Q-code extraction + fuzzy match of slide question/title text against table titles. |
| **Series-to-row fuzzy** | 0.10 | New: `_series_name_score()` in `auto_mapper.py` | Fuzzy match chart series name against all row labels across all tables. |

For segment breakout charts (where categories are segments, not response options), the engine automatically detects this pattern and switches to **flipped correlation mode**: it correlates chart values against each row's values *across columns* instead of each table's column values *across rows*.

### Segment detection heuristic

A chart is a "segment breakout" when its categories overlap heavily with the column banners (Total, Gen Z, Millennial, etc.) and its series name is NOT in the set of uninformative names (%, Series 1, etc.). The current `_is_uninformative_series()` check already exists — we'll use its inverse to trigger flipped matching.

---

## Implementation Plan

### Phase 1: Value Correlation Engine (auto_mapper.py)

**New function: `_value_correlation_score()`**

```
Input:  ShapeFingerprint (with values_sample), list of tables
Output: List[MapCandidate] with correlation-based scores
```

**Logic for distribution charts** (categories = response options):
1. Extract chart values as a vector V_chart (length = number of categories)
2. For each crosstab table:
   a. Find the "Total" column (or first column)
   b. Extract values for non-base/non-mean rows → V_table
   c. If len(V_chart) != len(V_table), try fuzzy-aligning by matching category labels to row labels
   d. Compute Pearson correlation r = corr(V_chart, V_table)
   e. Score = max(0, r) — negative correlations get 0

**Logic for segment breakout charts** (categories = segments):
1. Extract chart values V_chart (one value per segment)
2. For each crosstab table, for each non-base/non-mean row:
   a. Extract that row's values across the segment columns → V_row
   b. Map chart categories to crosstab column banners (fuzzy match "Baby Boomer+" → "Boomer or older")
   c. Compute Pearson correlation r = corr(V_chart, V_row)
   d. Track the best (table, row) combination
3. Score = max correlation found; also record which row matched (needed for alt-text writing)

**Segment-to-banner mapping** (reusable helper):
- Build a mapping of chart category names → crosstab column banner names
- Use SequenceMatcher with threshold 0.6 for fuzzy matching
- Cache the mapping per chart since it's the same for all tables
- Handle common variations: "Baby Boomer+" ↔ "Boomer or older", "Northeast" ↔ "NORTHEAST"

**Changes to `ShapeFingerprint`:**
- Already has `values_sample: List[float]` — ensure this is populated for all chart types (currently capped at 20 values, which is enough)

**Changes to `score_fingerprint_against_tables()`:**
- After existing structural scoring, compute value correlation score
- Blend: `final_score = 0.45 * value_corr + 0.25 * label_overlap + 0.20 * slide_context + 0.10 * series_match`
- Return enhanced `MapCandidate` with all sub-scores for debugging/UI

### Phase 2: Slide Context Extraction (auto_mapper.py)

**New function: `_extract_slide_context()`**

For each slide that has a chart/table shape, extract text from all text boxes on that slide and classify them:
- **Question text**: starts with "Question:" or contains question-like phrasing
- **Chart title**: typically in a TextBox named "TextBox 1" or "TextBox 2", shorter text
- **Base text**: starts with "Base:" or contains "respondents"
- **Q-code**: regex extract `(Q\d+)` from any text on the slide

**New function: `_slide_context_score()`**

```
Input:  SlideContext (question text, title, q-code), table dict
Output: float 0.0–1.0
```

Logic:
1. If Q-code is found in slide text AND in table title → score = 1.0 (exact structural match)
2. Else: SequenceMatcher ratio of slide question text vs table title (normalized, lowercased)
3. Boost: if chart title keywords appear in table row labels (e.g., "Gasoline" in title matches "Gasoline was too expensive" row)

**New dataclass: `SlideContext`**

```python
@dataclass
class SlideContext:
    slide_idx: int
    question_text: str = ""
    chart_title: str = ""
    base_text: str = ""
    q_code: Optional[str] = None
```

**Changes to `extract_all_fingerprints()`:**
- Build a `Dict[int, SlideContext]` (slide_idx → context) alongside fingerprints
- Pass slide context into the scoring pipeline

### Phase 3: Series-to-Row Fuzzy Matching (auto_mapper.py)

**New function: `_series_name_score()`**

```
Input:  series_name (str), table dict
Output: (float score, str matched_row_label)
```

Logic:
1. If series name is uninformative (%, Series 1, etc.) → return 0.0
2. Normalize both series name and all row labels
3. Try exact substring match first (e.g., "Gas too expensive" in "Gasoline was too expensive")
4. Fall back to SequenceMatcher ratio with threshold 0.55
5. Return best match score and the matched row label (needed for flipped chart alt-text)

Special handling for composite labels:
- "Better or Much Better Off" should match "Top 2 Box" — this requires understanding that the series name describes the *combination* of the top two response options. Approach: also check if the series name is a substring of or similar to the *table title* keywords, and if so, look for "Top 2 Box" or "Top 3 Box" rows.

### Phase 4: Blended Scoring Orchestrator (auto_mapper.py)

**Modify: `auto_map_presentation_obj()`**

Current flow:
1. Extract fingerprints
2. Score structurally
3. LLM disambiguate if ambiguous
4. Associate text shapes
5. Write alt text

New flow:
1. Extract fingerprints **+ slide contexts** (new)
2. For each chart/table fingerprint:
   a. Detect chart type: distribution vs segment-breakout vs time-series
   b. **Value correlation scoring** (new — Phase 1)
   c. Structural label overlap scoring (existing, keep as-is)
   d. **Slide context scoring** (new — Phase 2)
   e. **Series-to-row scoring** (new — Phase 3)
   f. **Blend all signals** into final score with configurable weights
   g. If top candidate score < LLM_THRESHOLD and use_llm=True → **LLM fallback** (existing, but pass richer context)
3. Associate text shapes (existing)
4. Write alt text (existing, but include matched row for segment breakouts)

**New: `AutoMapResult` additions**

Add fields:
- `value_corr_score: float` — for the UI/debugging
- `label_score: float`
- `context_score: float`
- `series_score: float`
- `matched_row: Optional[str]` — which row the series name mapped to (for segment breakouts)

### Phase 5: Enhanced LLM Fallback (auto_mapper.py)

**Modify: `_llm_disambiguate()`**

Current prompt only sends shape info + candidate titles/labels. Upgrade to include:
- **Slide context** (question text, chart title) — huge help for the LLM
- **Value correlation scores** for each candidate — let the LLM see which tables are numerically closest
- **Series name** and its best fuzzy match per candidate

This makes the LLM fallback much more effective for the remaining hard cases where algorithmic scoring is ambiguous.

### Phase 6: SmartMatcher Upgrade (smart_match.py)

**For the update flow** (not auto-mapping — this is when alt text already exists from a previous run):

The existing SmartMatcher uses title matching + row/col Jaccard. Upgrade Tier 2 fuzzy matching to also consider:
- Value correlation (when the shape has chart data accessible)
- Response-option text similarity (match row labels in alt text against new table's row labels)

This handles the case where table titles changed between crosstab versions but the underlying data is the same.

**Add new method: `_value_enhanced_fuzzy()`**

Same Pearson correlation approach as Phase 1, but applied to chart data extracted at match time (not from fingerprints). This is a lighter-weight check used only when title matching fails.

### Phase 7: UI Enhancements (app.py)

**Auto-Map Results display upgrades:**
- Show sub-scores (value correlation, label overlap, context, series match) in the results table
- Color-code: green (r > 0.90), yellow (0.70-0.90), red (< 0.70)
- Show matched row label for segment breakout charts
- Show slide context (question text) alongside each match for human verification

**New column in match review table:**
- "Value Corr" column showing the Pearson r for each match
- "Matched Row" column for segment breakout charts

### Phase 8: Generalization Safeguards

The entire design must work for any crosstab + PPTX combination, not just SATS. Key safeguards:

1. **No hardcoded segment names**: The segment detection heuristic looks for overlap between chart categories and column banners dynamically — it doesn't know or care that "Gen Z" is a generation.

2. **No hardcoded response option patterns**: The value correlation approach is completely content-agnostic — it just compares numbers.

3. **Configurable weights**: The signal weights (0.45/0.25/0.20/0.10) should be constants at the top of the file, easy to tune.

4. **Graceful degradation**: If value correlation can't compute (chart has no numeric data), fall back to existing structural scoring. If slide context is empty, that signal gets 0 and the other signals compensate.

5. **Multi-sheet support**: The crosstab parser already handles multiple sheets. The scoring engine iterates all 277 tables regardless of which sheet they came from.

6. **Mixed chart types**: Each chart is independently classified and scored. A deck can mix distribution charts, segment breakouts, and time series — each gets the right matching strategy.

---

## File Change Summary

| File | Changes | Scope |
|---|---|---|
| `auto_mapper.py` | Value correlation engine, slide context extraction, series-to-row matching, blended scoring orchestrator, enhanced LLM prompt | Major — bulk of new code |
| `smart_match.py` | Value-enhanced fuzzy matching in Tier 2 for update flow | Moderate |
| `app.py` | UI display of sub-scores, matched row column, value correlation column | Minor |
| `deck_update.py` | No changes needed — it already uses SmartMatcher/AutoMapper results | None |
| `crosstab_parser.py` | No changes needed — already parses all sheets correctly | None |
| `chart_data_patcher.py` | No changes needed | None |
| `text_utils.py` | No changes needed | None |
| `brand_config.py` | No changes needed | None |
| `requirements.txt` | Add `numpy` (for Pearson correlation) | One line |

---

## Implementation Order & Dependencies

```
Phase 1 (Value Correlation) ──┐
Phase 2 (Slide Context)   ────┤
Phase 3 (Series-to-Row)   ────┼──▶ Phase 4 (Blended Orchestrator) ──▶ Phase 5 (LLM) ──▶ Phase 7 (UI)
                               │
                               └──▶ Phase 6 (SmartMatcher upgrade)
                               
Phase 8 (Generalization) — runs throughout, not a separate step
```

Phases 1–3 are independent and can be built in any order. Phase 4 integrates them. Phase 5 enhances the LLM fallback with the new signals. Phase 6 is a parallel track for the update-flow matcher. Phase 7 is UI polish.

---

## Testing Strategy

### Unit tests for value correlation

- Distribution chart with known crosstab match → r > 0.95
- Segment breakout chart with known match → r > 0.90
- Random/unrelated data → r < 0.30
- Edge case: chart has only 2 data points (low n for correlation)
- Edge case: chart values are all identical (no variance → undefined correlation)

### Integration test with SATS data

- Run auto-mapper on the 80-slide SATS deck + 277-table crosstab
- Verify: distribution charts on slides 2, 12, 15, 19, 23, 26, 29, 30 all match correctly
- Verify: segment breakout charts on slides 4–8, 13, 16, 20, 24, 27, 31 all match correctly
- Verify: time series charts on slides 3, 14, 17, 21, 25, 28 are all skipped
- Target: 90%+ of charts matched with confidence > 0.80

### Generalization test

- Run on a different project's PPTX + crosstab (if available) to confirm no SATS-specific assumptions

---

## Risk & Mitigation

| Risk | Mitigation |
|---|---|
| Pearson correlation undefined for constant vectors (all same value) | Check variance > 0 before computing; fall back to label overlap |
| Multiple tables with very similar data patterns | Slide context and series-name matching break ties; LLM fallback for remaining ambiguity |
| Chart shows only a subset of rows (e.g., top 5 of 20) | Fuzzy-align chart categories to row labels before correlation; only correlate aligned pairs |
| Crosstab values are percentages (0-1) but chart displays as whole numbers (0-100) | Normalize both to same scale before correlation |
| New projects with radically different crosstab format | Graceful degradation — if no signal scores above threshold, shape is left unmatched (current behavior) |

---

## Estimated Effort

| Phase | New Lines of Code (approx) | Complexity |
|---|---|---|
| Phase 1: Value Correlation | ~180 | High — core algorithm |
| Phase 2: Slide Context | ~80 | Medium |
| Phase 3: Series-to-Row | ~60 | Medium |
| Phase 4: Blended Orchestrator | ~100 | Medium — integration |
| Phase 5: LLM Enhancement | ~40 | Low — prompt upgrade |
| Phase 6: SmartMatcher | ~80 | Medium |
| Phase 7: UI | ~50 | Low |
| **Total** | **~590** | |

---

## Open Questions for Andrew

1. **Confidence threshold for auto-writing alt text**: Currently 0.70. With value correlation, should we raise it to 0.80 to be more conservative, or is 0.70 fine?

2. **"Top 2 Box" / "Top 3 Box" matching**: When a chart's series name is "Better or Much Better Off" and the table has a "Top 2 Box" row, should we treat this as a match? It requires knowing that the series name describes the sum of the top two options. I can build this as a heuristic (detect "or" in series name → look for Top N Box rows), but want to confirm this pattern is common across your projects.

3. **Should time series charts remain fully skipped**, or should they be tagged with the matching table title even though their values won't update? (This would let future features like "update time series by appending new wave data" work.)
