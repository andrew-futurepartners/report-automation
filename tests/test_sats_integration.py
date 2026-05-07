"""Integration test: auto-map pipeline against the real SATS dataset.

Requires the sample data files in sample_data/SATS/.  Marked ``slow``
so the fast unit-test suite can skip it (run with ``pytest -m slow``).
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from pptx import Presentation
from crosstab_parser import parse_workbook
from auto_mapper import (
    auto_map_presentation_obj,
    extract_all_fingerprints,
    results_to_report,
)

_SATS_DIR = os.path.join(os.path.dirname(__file__), "..", "sample_data", "SATS")
_PPTX = os.path.join(_SATS_DIR, "The State of the American Traveler - March 2026 Report.pptx")
_XLSX = os.path.join(_SATS_DIR, "SATS March 2026 - Internal Crosstab 4 3 2026.xlsx")

_HAS_DATA = os.path.isfile(_PPTX) and os.path.isfile(_XLSX)
_SKIP_REASON = "SATS sample data not present in sample_data/SATS/"

pytestmark = pytest.mark.slow


@pytest.fixture(scope="module")
def sats_data():
    """Parse the SATS crosstab once for the whole module."""
    return parse_workbook(_XLSX)


@pytest.fixture(scope="module")
def sats_prs():
    """Load the SATS PPTX once for the whole module."""
    return Presentation(_PPTX)


@pytest.fixture(scope="module")
def sats_map_results(sats_prs, sats_data):
    """Run auto_map_presentation_obj (no LLM, no alt-text writes)."""
    return auto_map_presentation_obj(
        sats_prs, sats_data["tables"],
        use_llm=False, write_alt=False,
    )


# ---------------------------------------------------------------------------
# Crosstab parsing
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _HAS_DATA, reason=_SKIP_REASON)
class TestCrosstabParsing:
    def test_parses_multiple_sheets(self, sats_data):
        tables = sats_data["tables"]
        assert len(tables) >= 100, f"Expected 100+ tables, got {len(tables)}"

    def test_tables_have_required_fields(self, sats_data):
        for t in sats_data["tables"][:10]:
            assert "title" in t
            assert "row_labels" in t
            assert "col_labels" in t
            assert "values" in t
            assert len(t["row_labels"]) >= 1
            assert len(t["col_labels"]) >= 1


# ---------------------------------------------------------------------------
# Fingerprint extraction
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _HAS_DATA, reason=_SKIP_REASON)
class TestFingerprinting:
    def test_extracts_fingerprints_from_all_slides(self, sats_prs):
        fps = extract_all_fingerprints(sats_prs)
        chart_fps = [f for f in fps if f.shape_type == "chart"]
        table_fps = [f for f in fps if f.shape_type == "table"]
        text_fps = [f for f in fps if f.shape_type == "text"]
        assert len(chart_fps) >= 20, f"Expected 20+ chart fingerprints, got {len(chart_fps)}"
        assert len(fps) >= 30, f"Expected 30+ total fingerprints, got {len(fps)}"

    def test_chart_fingerprints_have_values(self, sats_prs):
        fps = extract_all_fingerprints(sats_prs)
        chart_fps = [f for f in fps if f.shape_type == "chart"]
        with_vals = [f for f in chart_fps if f.values_sample]
        assert len(with_vals) >= 15, (
            f"Expected 15+ charts with values, got {len(with_vals)}"
        )


# ---------------------------------------------------------------------------
# Auto-mapping
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _HAS_DATA, reason=_SKIP_REASON)
class TestAutoMap:
    def test_no_crash_on_full_deck(self, sats_map_results):
        """The pipeline should complete without exceptions."""
        assert sats_map_results is not None
        assert len(sats_map_results) >= 1

    def test_match_rate_above_minimum(self, sats_map_results):
        data_results = [
            r for r in sats_map_results
            if r.fingerprint.shape_type in ("chart", "table")
        ]
        matched = [r for r in data_results if r.table_title is not None]
        rate = len(matched) / max(len(data_results), 1)
        assert rate >= 0.40, (
            f"Match rate {rate:.0%} is below 40% minimum "
            f"({len(matched)}/{len(data_results)} shapes matched)"
        )

    def test_high_confidence_matches_exist(self, sats_map_results):
        high_conf = [
            r for r in sats_map_results
            if r.confidence >= 0.80 and r.table_title is not None
        ]
        assert len(high_conf) >= 5, (
            f"Expected 5+ high-confidence matches, got {len(high_conf)}"
        )

    def test_multiple_match_methods_used(self, sats_map_results):
        methods = {r.method for r in sats_map_results if r.table_title}
        assert len(methods) >= 1, "Expected at least one match method"

    def test_value_correlation_produces_scores(self, sats_map_results):
        with_val_corr = [
            r for r in sats_map_results
            if r.value_corr_score > 0.5
        ]
        assert len(with_val_corr) >= 3, (
            f"Expected 3+ shapes with strong value correlation, got {len(with_val_corr)}"
        )


# ---------------------------------------------------------------------------
# Report serialisation
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _HAS_DATA, reason=_SKIP_REASON)
class TestReportSerialization:
    def test_results_to_report_works(self, sats_map_results):
        report = results_to_report(sats_map_results)
        assert isinstance(report, list)
        assert len(report) >= 1
        for entry in report:
            assert "shape_name" in entry
            assert "confidence" in entry
            assert "method" in entry
