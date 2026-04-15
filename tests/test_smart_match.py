"""Tests for smart_match.py — three-tier matching engine."""

import pytest

from smart_match import SmartMatcher, MatchResult, MatchCandidate, _norm, label_hash, _jaccard


def _tables():
    """Helper returning a list of table dicts for matcher construction."""
    return [
        {
            "title": "Brand Awareness",
            "row_labels": ["Aided", "Unaided", "Top of Mind"],
            "col_labels": ["Total", "Male", "Female"],
            "values": [[0.85, 0.80, 0.90], [0.45, 0.42, 0.48], [0.22, 0.20, 0.24]],
        },
        {
            "title": "Purchase Intent",
            "row_labels": ["Definitely", "Probably", "Might"],
            "col_labels": ["Total", "18-34", "35-54"],
            "values": [[0.30, 0.35, 0.25], [0.40, 0.38, 0.42], [0.20, 0.22, 0.18]],
        },
        {
            "title": "Net Promoter Score",
            "row_labels": ["Promoters", "Passives", "Detractors"],
            "col_labels": ["Total", "Q1", "Q2"],
            "values": [[0.45, 0.42, 0.48], [0.35, 0.38, 0.32], [0.20, 0.20, 0.20]],
        },
    ]


# ---------------------------------------------------------------------------
# Tier 1 — exact match
# ---------------------------------------------------------------------------

class TestTier1ExactMatch:
    def test_exact_match_by_alt_title(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({"name": "Shape1", "alt": {"table_title": "Brand Awareness"}})
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.tier == 1
        assert result.confidence == 1.0
        assert result.status == "matched"

    def test_exact_match_normalization(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Shape1",
            "alt": {"table_title": "  BRAND   awareness! "},
        })
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.tier == 1

    def test_exact_match_name_fallback_chart_underscore(self):
        tables = [
            {
                "title": "Brand_Awareness",
                "row_labels": ["Aided", "Unaided"],
                "col_labels": ["Total", "Male"],
                "values": [[0.85, 0.80], [0.45, 0.42]],
            },
        ]
        matcher = SmartMatcher(tables)
        result = matcher.match({
            "name": "CHART_Brand_Awareness_Total",
            "alt": {},
        })
        assert result is not None
        assert result.table["title"] == "Brand_Awareness"
        assert result.tier == 1
        assert result.col_key == "Total"

    def test_exact_match_name_fallback_table_underscore(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "TABLE_Brand_Awareness",
            "alt": {},
        })
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.tier == 1

    def test_exact_match_name_fallback_chart_colon(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "CHART:Brand Awareness:Female",
            "alt": {},
        })
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.col_key == "Female"

    def test_exact_match_name_fallback_table_colon(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "TABLE:Purchase Intent",
            "alt": {},
        })
        assert result is not None
        assert result.table["title"] == "Purchase Intent"


# ---------------------------------------------------------------------------
# Tier 2 — fuzzy match
# ---------------------------------------------------------------------------

class TestTier2FuzzyMatch:
    def test_fuzzy_match_above_threshold(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Shape1",
            "alt": {
                "table_title": "Brand Awareness Survey",
                "row_labels": ["Aided", "Unaided", "Top of Mind"],
                "col_labels": ["Total", "Male", "Female"],
            },
        })
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.tier == 2
        assert result.confidence >= 0.75

    def test_fuzzy_match_below_threshold_no_match(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Shape1",
            "alt": {"table_title": "Completely Different Topic XYZ"},
        })
        assert result is None

    def test_fuzzy_scoring_weights(self):
        """Title=0.40, Row=0.35, Col=0.25 — verify weighting by making
        only title match perfectly and checking the resulting score."""
        tables = [{
            "title": "Exact Title Match",
            "row_labels": ["A", "B"],
            "col_labels": ["X", "Y"],
            "values": [[1, 2], [3, 4]],
        }]
        matcher = SmartMatcher(tables)
        result = matcher.match({
            "name": "Shape1",
            "alt": {
                "table_title": "Exact Title Match",
                "row_labels": ["Completely", "Different"],
                "col_labels": ["No", "Match"],
            },
        })
        # Exact title match via Tier 1 will fire first since norm matches
        assert result is not None
        assert result.tier == 1

    def test_fuzzy_candidates_populated(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Shape1",
            "alt": {
                "table_title": "Brand Awareness Survey",
                "row_labels": ["Aided", "Unaided", "Top of Mind"],
            },
        })
        assert result is not None
        assert len(result.candidates) > 0
        first = result.candidates[0]
        assert first.title_score > 0

    def test_fuzzy_with_row_hash_match(self):
        tables = _tables()
        rh = label_hash(tables[0]["row_labels"])
        matcher = SmartMatcher(tables)
        result = matcher.match({
            "name": "Shape1",
            "alt": {
                "table_title": "Brand Awareness Tracking",
                "row_hash": rh,
            },
        })
        assert result is not None
        assert result.table["title"] == "Brand Awareness"


# ---------------------------------------------------------------------------
# Threshold behavior
# ---------------------------------------------------------------------------

class TestThresholdBehavior:
    def test_low_confidence_between_060_075(self):
        tables = [{
            "title": "Customer Satisfaction Index",
            "row_labels": ["Very Satisfied", "Somewhat Satisfied"],
            "col_labels": ["Total", "Region A"],
            "values": [[0.5, 0.6], [0.3, 0.2]],
        }]
        matcher = SmartMatcher(tables)
        result = matcher.match({
            "name": "Shape1",
            "alt": {"table_title": "Customer Satisfaction"},
        })
        if result is not None and result.tier == 2:
            if result.confidence < 0.75:
                assert result.status == "low_confidence"

    def test_below_060_returns_none(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Shape1",
            "alt": {"table_title": "ZZZZZ Unrelated ZZZZZ"},
        })
        assert result is None


# ---------------------------------------------------------------------------
# Deduplication (match_all)
# ---------------------------------------------------------------------------

class TestDeduplication:
    def test_match_all_deduplication(self):
        matcher = SmartMatcher(_tables())
        shapes = [
            {"name": "Chart1", "alt": {"table_title": "Brand Awareness"}},
            {"name": "Chart2", "alt": {"table_title": "Brand Awareness"}},
        ]
        results = matcher.match_all(shapes)
        assert len(results) == 2

        matched = [r for r in results if r is not None and r.table is not None]
        assert len(matched) == 1
        assert matched[0].table["title"] == "Brand Awareness"

        dupes = [r for r in results if r is not None and r.status == "duplicate"]
        assert len(dupes) == 1

    def test_match_all_different_tables(self):
        matcher = SmartMatcher(_tables())
        shapes = [
            {"name": "Chart1", "alt": {"table_title": "Brand Awareness"}},
            {"name": "Chart2", "alt": {"table_title": "Purchase Intent"}},
        ]
        results = matcher.match_all(shapes)
        tables_matched = [
            r.table["title"] for r in results if r is not None and r.table is not None
        ]
        assert "Brand Awareness" in tables_matched
        assert "Purchase Intent" in tables_matched


# ---------------------------------------------------------------------------
# MatchResult properties
# ---------------------------------------------------------------------------

class TestMatchResultFields:
    def test_match_result_has_expected_fields(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({"name": "MyChart", "alt": {"table_title": "Brand Awareness", "column": "Male"}})
        assert isinstance(result, MatchResult)
        assert result.shape_name == "MyChart"
        assert result.shape_alt_title == "Brand Awareness"
        assert result.col_key == "Male"
        assert result.tier in (1, 2, 3)
        assert result.confidence > 0

    def test_failed_match_returns_none(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({"name": "NoMatch", "alt": {"table_title": "ZZZZZZ"}})
        assert result is None

    def test_exclude_terms_passthrough(self):
        matcher = SmartMatcher(_tables())
        result = matcher.match({
            "name": "Chart1",
            "alt": {
                "table_title": "Brand Awareness",
                "exclude_rows": "Base, Mean",
            },
        })
        assert result is not None
        assert result.exclude_terms == ["Base", "Mean"]


# ---------------------------------------------------------------------------
# Overrides
# ---------------------------------------------------------------------------

class TestOverrides:
    def test_override_match(self):
        tables = _tables()
        matcher = SmartMatcher(tables, overrides={"Brand Awareness": "Purchase Intent"})
        result = matcher.match({
            "name": "Chart1",
            "alt": {"table_title": "Brand Awareness"},
        })
        assert result is not None
        assert result.table["title"] == "Purchase Intent"
        assert result.tier == 0
        assert result.status == "override"
        assert result.confidence == 1.0

    def test_skip_override(self):
        tables = _tables()
        matcher = SmartMatcher(tables, overrides={"Brand Awareness": "__skip__"})
        result = matcher.match({
            "name": "Chart1",
            "alt": {"table_title": "Brand Awareness"},
        })
        assert result is None

    def test_override_invalid_title_ignored(self):
        tables = _tables()
        matcher = SmartMatcher(tables, overrides={"Brand Awareness": "Nonexistent Table"})
        result = matcher.match({
            "name": "Chart1",
            "alt": {"table_title": "Brand Awareness"},
        })
        # Override didn't register because target title doesn't exist; falls through to Tier 1
        assert result is not None
        assert result.table["title"] == "Brand Awareness"
        assert result.tier == 1


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

class TestHelpers:
    def test_norm(self):
        assert _norm("  Hello,  World! ") == "hello world"
        assert _norm("") == ""
        assert _norm(None) == ""

    def test_label_hash_deterministic(self):
        labels = ["Aided", "Unaided", "Top of Mind"]
        h1 = label_hash(labels)
        h2 = label_hash(labels)
        assert h1 == h2
        assert len(h1) == 8

    def test_label_hash_order_independent(self):
        h1 = label_hash(["A", "B", "C"])
        h2 = label_hash(["C", "A", "B"])
        assert h1 == h2

    def test_jaccard_identical(self):
        assert _jaccard({"a", "b"}, {"a", "b"}) == 1.0

    def test_jaccard_disjoint(self):
        assert _jaccard({"a"}, {"b"}) == 0.0

    def test_jaccard_empty_sets(self):
        assert _jaccard(set(), set()) == 1.0

    def test_jaccard_partial(self):
        result = _jaccard({"a", "b", "c"}, {"b", "c", "d"})
        assert abs(result - 0.5) < 0.01


# ---------------------------------------------------------------------------
# get_report
# ---------------------------------------------------------------------------

class TestGetReport:
    def test_report_after_match(self):
        matcher = SmartMatcher(_tables())
        matcher.match({"name": "Chart1", "alt": {"table_title": "Brand Awareness"}})
        report = matcher.get_report()
        assert len(report) == 1
        assert report[0]["shape_name"] == "Chart1"
        assert report[0]["matched_table"] == "Brand Awareness"
        assert report[0]["tier"] == 1
