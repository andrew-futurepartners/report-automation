"""
smart_match.py — Three-tier intelligent matching for PowerPoint shapes to crosstab tables.

Tier 1: Exact match (normalize whitespace/case, compare table_title)
Tier 2: Fuzzy structural match (title similarity + row/col Jaccard overlap)
Tier 3: Optional LLM-assisted match for ambiguous cases (cached per session)
"""

import hashlib
import json
import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional

logger = logging.getLogger("report_relay.smart_match")

# ---------------------------------------------------------------------------
# Normalisation helpers (shared with deck_update)
# ---------------------------------------------------------------------------

def _norm(s: str) -> str:
    """Normalize: collapse whitespace, strip punctuation, lowercase."""
    t = re.sub(r"\s+", " ", (s or "")).strip().lower()
    return re.sub(r"[^\w\s]", "", t).strip()


def label_hash(labels: list) -> str:
    """Short deterministic hash of a sorted label set — stored in alt text."""
    canonical = "|".join(sorted(str(l).lower().strip() for l in labels))
    return hashlib.md5(canonical.encode()).hexdigest()[:8]


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class MatchCandidate:
    """A single candidate with its score breakdown."""
    table: Dict[str, Any]
    score: float
    title_score: float = 0.0
    row_score: float = 0.0
    col_score: float = 0.0


@dataclass
class MatchResult:
    """Outcome of matching one shape to crosstab tables."""
    shape_name: str
    shape_alt_title: str
    table: Optional[Dict[str, Any]]
    confidence: float
    tier: int                               # 1, 2, or 3
    candidates: List[MatchCandidate] = field(default_factory=list)
    col_key: Optional[str] = None
    exclude_terms: Optional[List[str]] = None
    status: str = "matched"                 # matched | low_confidence | failed
    shape_type: str = "unknown"             # chart | table | unknown


# ---------------------------------------------------------------------------
# Jaccard helper
# ---------------------------------------------------------------------------

def _jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    union = a | b
    if not union:
        return 0.0
    return len(a & b) / len(union)


# ---------------------------------------------------------------------------
# SmartMatcher
# ---------------------------------------------------------------------------

TITLE_WEIGHT = 0.40
ROW_WEIGHT   = 0.35
COL_WEIGHT   = 0.25


class SmartMatcher:
    """Three-tier shape-to-table matching engine.

    Args:
        tables: list of parsed crosstab table dicts (each with "title",
                "row_labels", "col_labels", etc.)
        confidence_threshold: minimum combined score to accept a Tier-2 match.
        use_llm: if True, engage Tier-3 LLM disambiguation for ambiguous cases.
    """

    def __init__(
        self,
        tables: List[Dict[str, Any]],
        confidence_threshold: float = 0.75,
        use_llm: bool = False,
        overrides: Optional[Dict[str, str]] = None,
    ):
        self._tables = tables
        self._threshold = confidence_threshold
        self._use_llm = use_llm

        # Pre-compute normalised title map for Tier 1
        self._norm_map: Dict[str, Dict[str, Any]] = {}
        for t in tables:
            key = _norm(t.get("title", ""))
            if key:
                self._norm_map[key] = t

        # User overrides: maps shape alt-title (or name) → forced table title
        self._overrides: Dict[str, str] = {}
        if overrides:
            for shape_label, table_title in overrides.items():
                if table_title == "__skip__":
                    self._overrides[shape_label] = table_title
                elif _norm(table_title) in self._norm_map:
                    self._overrides[shape_label] = table_title

        # LLM decision cache (keyed by alt-text hash → MatchResult)
        self._llm_cache: Dict[str, MatchResult] = {}

        # Accumulate results for the match report
        self._report: List[MatchResult] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def match(self, shape_metadata: dict) -> Optional[MatchResult]:
        """Match a single shape to a table using the three-tier pipeline.

        *shape_metadata* is expected to have at least::

            {"name": str, "alt": dict}   # alt from _parse_alt_text

        An optional ``"shape_type"`` key (``"chart"`` or ``"table"``) is
        carried through to the ``MatchResult`` for UI display.

        Returns ``MatchResult`` or ``None`` if nothing matched.
        """
        name = shape_metadata.get("name", "")
        alt  = shape_metadata.get("alt", {})
        alt_title = alt.get("table_title", "")
        shape_type = shape_metadata.get("shape_type", "unknown")

        col_key = alt.get("column")
        exclude_terms = None
        if "exclude_rows" in alt and isinstance(alt["exclude_rows"], str):
            exclude_terms = [p.strip() for p in alt["exclude_rows"].split(",")]

        # --- Tier 0: user override (from match-review UI) ---
        override_label = alt_title or name
        if override_label in self._overrides:
            forced_title = self._overrides[override_label]
            if forced_title == "__skip__":
                skipped = MatchResult(
                    shape_name=name,
                    shape_alt_title=alt_title,
                    table=None,
                    confidence=0.0,
                    tier=0,
                    status="skipped",
                    shape_type=shape_type,
                )
                self._report.append(skipped)
                return None
            forced_table = self._norm_map.get(_norm(forced_title))
            if forced_table:
                result = MatchResult(
                    shape_name=name,
                    shape_alt_title=alt_title,
                    table=forced_table,
                    confidence=1.0,
                    tier=0,
                    col_key=col_key,
                    exclude_terms=exclude_terms,
                    status="override",
                    shape_type=shape_type,
                )
                self._report.append(result)
                return result

        # --- Tier 1: exact (normalised) match ---
        result = self._exact_match(name, alt, alt_title, col_key, exclude_terms)
        if result:
            result.shape_type = shape_type
            self._report.append(result)
            return result

        # --- Tier 2: fuzzy structural match ---
        result = self._fuzzy_match(name, alt, alt_title, col_key, exclude_terms)
        if result:
            result.shape_type = shape_type
            self._report.append(result)
            return result

        # --- Tier 3: LLM-assisted (optional) ---
        if self._use_llm:
            result = self._llm_match(name, alt, alt_title, col_key, exclude_terms)
            if result:
                result.shape_type = shape_type
                self._report.append(result)
                return result

        # Nothing matched
        failed = MatchResult(
            shape_name=name,
            shape_alt_title=alt_title,
            table=None,
            confidence=0.0,
            tier=0,
            status="failed",
            shape_type=shape_type,
        )
        self._report.append(failed)
        return None

    def match_all(self, shapes_metadata: List[dict]) -> List[Optional[MatchResult]]:
        """Batch-match with deduplication — highest-confidence claim wins per table."""
        raw = [self.match(sm) for sm in shapes_metadata]

        # First pass: find the best confidence per table title
        best_per_table: Dict[str, float] = {}
        for r in raw:
            if r is None or r.table is None:
                continue
            title = r.table.get("title", "")
            if title and r.confidence > best_per_table.get(title, -1.0):
                best_per_table[title] = r.confidence

        # Second pass: downgrade duplicates (lower-confidence claims on same table)
        seen: Dict[str, bool] = {}
        for r in raw:
            if r is None or r.table is None:
                continue
            title = r.table.get("title", "")
            if not title:
                continue
            if title in seen:
                r.table = None
                r.status = "duplicate"
                r.confidence = 0.0
                logger.debug("Duplicate match for '%s' downgraded: %s", title, r.shape_name)
            elif r.confidence >= best_per_table[title]:
                seen[title] = True
            else:
                r.table = None
                r.status = "duplicate"
                r.confidence = 0.0
                logger.debug("Duplicate match for '%s' downgraded: %s", title, r.shape_name)

        return raw

    def get_report(self) -> List[dict]:
        """Return a serialisable match report for the Streamlit UI."""
        out = []
        for r in self._report:
            entry = {
                "shape_name": r.shape_name,
                "shape_alt_title": r.shape_alt_title,
                "shape_type": r.shape_type,
                "matched_table": r.table.get("title") if r.table else None,
                "confidence": round(r.confidence, 3),
                "tier": r.tier,
                "status": r.status,
                "candidates": [
                    {
                        "title": c.table.get("title"),
                        "score": round(c.score, 3),
                        "title_score": round(c.title_score, 3),
                        "row_score": round(c.row_score, 3),
                        "col_score": round(c.col_score, 3),
                    }
                    for c in r.candidates
                ],
            }
            out.append(entry)
        return out

    # ------------------------------------------------------------------
    # Tier 1 — exact match (normalised whitespace/case)
    # ------------------------------------------------------------------

    def _exact_match(
        self, name: str, alt: dict, alt_title: str,
        col_key: Optional[str], exclude_terms: Optional[List[str]],
    ) -> Optional[MatchResult]:
        table = None

        # Alt-text title lookup
        if alt_title:
            norm_key = _norm(alt_title)
            table = self._norm_map.get(norm_key)

        # Shape name fallbacks: CHART_ / TABLE_ / CHART: / TABLE:
        if not table:
            table, col_key = self._name_fallback(name, col_key)

        if table:
            return MatchResult(
                shape_name=name,
                shape_alt_title=alt_title,
                table=table,
                confidence=1.0,
                tier=1,
                col_key=col_key,
                exclude_terms=exclude_terms,
            )
        return None

    # ------------------------------------------------------------------
    # Tier 2 — fuzzy structural match
    # ------------------------------------------------------------------

    def _fuzzy_match(
        self, name: str, alt: dict, alt_title: str,
        col_key: Optional[str], exclude_terms: Optional[List[str]],
    ) -> Optional[MatchResult]:
        if not alt_title:
            return None

        norm_alt = _norm(alt_title)
        alt_row_hash = alt.get("row_hash")
        alt_col_hash = alt.get("col_hash")

        candidates: List[MatchCandidate] = []

        for t in self._tables:
            t_title = _norm(t.get("title", ""))
            if not t_title:
                continue

            # Title similarity (SequenceMatcher)
            title_sim = SequenceMatcher(None, norm_alt, t_title).ratio()

            # Row label Jaccard
            row_sim = 0.0
            if alt_row_hash:
                t_row_hash = label_hash(t.get("row_labels", []))
                row_sim = 1.0 if alt_row_hash == t_row_hash else 0.0
            else:
                alt_rows = {_norm(r) for r in alt.get("row_labels", [])}
                t_rows   = {_norm(r) for r in t.get("row_labels", [])}
                if alt_rows or t_rows:
                    row_sim = _jaccard(alt_rows, t_rows)

            # Column label Jaccard
            col_sim = 0.0
            if alt_col_hash:
                t_col_hash = label_hash(t.get("col_labels", []))
                col_sim = 1.0 if alt_col_hash == t_col_hash else 0.0
            else:
                alt_cols = {_norm(c) for c in alt.get("col_labels", [])}
                t_cols   = {_norm(c) for c in t.get("col_labels", [])}
                if alt_cols or t_cols:
                    col_sim = _jaccard(alt_cols, t_cols)

            combined = (
                TITLE_WEIGHT * title_sim
                + ROW_WEIGHT * row_sim
                + COL_WEIGHT * col_sim
            )

            candidates.append(MatchCandidate(
                table=t, score=combined,
                title_score=title_sim, row_score=row_sim, col_score=col_sim,
            ))

        if not candidates:
            return None

        candidates.sort(key=lambda c: c.score, reverse=True)
        best = candidates[0]

        if best.score < self._threshold:
            # Check if multiple ambiguous candidates exist → potential Tier 3
            above_low = [c for c in candidates if c.score >= 0.60]
            if self._use_llm and len(above_low) >= 2:
                return None  # fall through to Tier 3
            if best.score >= 0.60:
                return MatchResult(
                    shape_name=name,
                    shape_alt_title=alt_title,
                    table=best.table,
                    confidence=best.score,
                    tier=2,
                    candidates=candidates[:5],
                    col_key=col_key,
                    exclude_terms=exclude_terms,
                    status="low_confidence",
                )
            return None

        return MatchResult(
            shape_name=name,
            shape_alt_title=alt_title,
            table=best.table,
            confidence=best.score,
            tier=2,
            candidates=candidates[:5],
            col_key=col_key,
            exclude_terms=exclude_terms,
        )

    # ------------------------------------------------------------------
    # Tier 3 — LLM-assisted disambiguation
    # ------------------------------------------------------------------

    def _llm_match(
        self, name: str, alt: dict, alt_title: str,
        col_key: Optional[str], exclude_terms: Optional[List[str]],
    ) -> Optional[MatchResult]:
        cache_key = _norm(alt_title)
        if cache_key in self._llm_cache:
            cached = self._llm_cache[cache_key]
            return MatchResult(
                shape_name=name,
                shape_alt_title=alt_title,
                table=cached.table,
                confidence=cached.confidence,
                tier=3,
                candidates=cached.candidates,
                col_key=col_key,
                exclude_terms=exclude_terms,
            )

        # Build candidate list for the prompt (top 3 from fuzzy)
        norm_alt = _norm(alt_title)
        scored: List[MatchCandidate] = []
        for t in self._tables:
            t_title = _norm(t.get("title", ""))
            if not t_title:
                continue
            title_sim = SequenceMatcher(None, norm_alt, t_title).ratio()
            scored.append(MatchCandidate(table=t, score=title_sim, title_score=title_sim))

        scored.sort(key=lambda c: c.score, reverse=True)
        top3 = scored[:3]

        if not top3:
            return None

        try:
            from openai import OpenAI
        except ImportError:
            logger.error("openai package not installed — skipping Tier 3 LLM match")
            return None

        prompt_candidates = []
        for i, c in enumerate(top3):
            t = c.table
            prompt_candidates.append({
                "index": i,
                "title": t.get("title", ""),
                "row_labels": t.get("row_labels", [])[:10],
                "col_labels": t.get("col_labels", []),
            })

        shape_info = {
            "alt_title": alt_title,
            "row_labels": alt.get("row_labels", []),
            "col_labels": alt.get("col_labels", []),
        }

        user_msg = (
            "A PowerPoint shape was previously mapped to a table titled "
            f'"{alt_title}". The new crosstab does not have an exact match. '
            "Which of these candidate tables is the best match?\n\n"
            f"Shape metadata: {json.dumps(shape_info)}\n\n"
            f"Candidates: {json.dumps(prompt_candidates)}\n\n"
            'Return JSON: {"best_index": <int or null>, "reason": "<one sentence>"}'
        )

        try:
            client = OpenAI()
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                max_tokens=120,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": "You are a data-matching assistant."},
                    {"role": "user", "content": user_msg},
                ],
            )
            body = json.loads(response.choices[0].message.content)
            idx = body.get("best_index")
            if idx is not None and 0 <= idx < len(top3):
                chosen = top3[idx]
                result = MatchResult(
                    shape_name=name,
                    shape_alt_title=alt_title,
                    table=chosen.table,
                    confidence=max(chosen.score, 0.80),
                    tier=3,
                    candidates=top3,
                    col_key=col_key,
                    exclude_terms=exclude_terms,
                )
                self._llm_cache[cache_key] = result
                logger.info(
                    "LLM match for '%s' → '%s' (reason: %s)",
                    alt_title, chosen.table.get("title"), body.get("reason", ""),
                )
                return result
        except Exception as e:
            logger.error("Tier 3 LLM match failed for '%s': %s", alt_title, e)

        return None

    # ------------------------------------------------------------------
    # Shape-name fallback helpers (CHART_*, TABLE_*, CHART:, TABLE:)
    # ------------------------------------------------------------------

    def _name_fallback(
        self, name: str, col_key: Optional[str],
    ) -> tuple:
        """Try to resolve a table from the shape name conventions.

        Returns ``(table_or_None, col_key)``.
        """
        if name.startswith("CHART_"):
            return self._parse_chart_underscore(name, col_key)
        if name.startswith("TABLE_"):
            return self._parse_table_underscore(name)
        if name.startswith("CHART:"):
            return self._parse_chart_colon(name, col_key)
        if name.startswith("TABLE:"):
            return self._parse_table_colon(name)
        return None, col_key

    def _parse_chart_underscore(self, name: str, col_key: Optional[str]) -> tuple:
        name_parts = name[6:].split("_")
        if len(name_parts) >= 2:
            potential_col = name_parts[-1]
            table_title = "_".join(name_parts[:-1])
            norm_title = _norm(table_title)
            for t in self._tables:
                if _norm(t.get("title", "")) == norm_title:
                    if potential_col in t.get("col_labels", []):
                        return t, potential_col
                    full_title = "_".join(name_parts)
                    if _norm(t.get("title", "")) == _norm(full_title):
                        return t, col_key
        return None, col_key

    def _parse_table_underscore(self, name: str) -> tuple:
        table_title = name[6:].replace("_", " ")
        norm_title = _norm(table_title)
        for t in self._tables:
            if _norm(t.get("title", "")) == norm_title:
                return t, None
        return None, None

    def _parse_chart_colon(self, name: str, col_key: Optional[str]) -> tuple:
        parts = name.split(":", 2)
        if len(parts) >= 2:
            table_title = parts[1].strip()
            ck = parts[2].strip() if len(parts) == 3 else col_key
            norm_title = _norm(table_title)
            for t in self._tables:
                if _norm(t.get("title", "")) == norm_title:
                    return t, ck
        return None, col_key

    def _parse_table_colon(self, name: str) -> tuple:
        parts = name.split(":", 1)
        if len(parts) == 2:
            table_title = parts[1].strip()
            norm_title = _norm(table_title)
            for t in self._tables:
                if _norm(t.get("title", "")) == norm_title:
                    return t, None
        return None, None
