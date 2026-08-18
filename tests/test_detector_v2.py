"""Regression tests for KPI integrity and broken-reference detection."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest
from typing import Any, Mapping, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
TEST_ROOT = Path(__file__).resolve().parent
for import_root in (SRC_ROOT, TEST_ROOT):
    if str(import_root) not in sys.path:
        sys.path.insert(0, str(import_root))

from pops_anomaly_detector import AnalysisConfig, analyze_directories  # noqa: E402
from pops_anomaly_detector.models import CountryResult, Finding, RunResult  # noqa: E402
from xlsx_factory import (  # noqa: E402
    Cell,
    CellLike,
    error_cell,
    make_grid,
    make_kpi_sheet,
    replace_cell,
    write_xlsx,
)


Sheets = Mapping[str, Sequence[Sequence[CellLike]]]


class SemanticDetectorTests(unittest.TestCase):
    def _pair(
        self,
        sent_sheets: Sheets,
        received_sheets: Sheets,
        *,
        sent_options: Mapping[str, Any] | None = None,
        received_options: Mapping[str, Any] | None = None,
    ) -> tuple[RunResult, CountryResult]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_dir.mkdir()
            received_dir.mkdir()
            write_xlsx(
                sent_dir / "France.xlsx",
                sent_sheets,
                **dict(sent_options or {}),
            )
            write_xlsx(
                received_dir / "France.xlsx",
                received_sheets,
                **dict(received_options or {}),
            )
            run = analyze_directories(
                AnalysisConfig(
                    sent_dir=sent_dir,
                    received_dir=received_dir,
                    output_dir=root / "reports",
                )
            )
        self.assertEqual(len(run.countries), 1)
        country = run.countries[0]
        self.assertEqual(country.comparison_state, "PAIRED", country.errors)
        return run, country

    @staticmethod
    def _findings(country: CountryResult, code: str) -> list[Finding]:
        return [finding for finding in country.findings if finding.code == code]

    def _one_finding(self, country: CountryResult, code: str) -> Finding:
        matches = self._findings(country, code)
        self.assertEqual(
            len(matches),
            1,
            f"Expected one {code}, got {[item.code for item in country.findings]}",
        )
        return matches[0]

    @staticmethod
    def _semantic_codes(country: CountryResult) -> set[str]:
        return {
            finding.code
            for finding in country.findings
            if finding.category in {"KPI_INTEGRITY", "FORMULA_INTEGRITY"}
        }

    @staticmethod
    def _sheet(country: CountryResult, name: str):
        matches = [
            comparison
            for comparison in country.sheets
            if comparison.sent_name == name or comparison.received_name == name
        ]
        if len(matches) != 1:
            raise AssertionError(f"Expected one comparison for {name!r}; got {len(matches)}")
        return matches[0]

    def test_exact_sheet_names_are_compared_when_counts_are_equal(self) -> None:
        common = make_grid(prefix="common")
        _run, country = self._pair(
            {"Overview": common, "Budget": make_grid(prefix="budget")},
            {"Overview": common, "Forecast": make_grid(prefix="forecast")},
        )

        self.assertEqual(country.metrics.sent_sheet_count, 2)
        self.assertEqual(country.metrics.received_sheet_count, 2)
        self.assertEqual(country.metrics.sheet_net_delta, 0)
        self.assertFalse(country.metrics.sheet_names_match)
        self.assertEqual(country.sent_sheet_names, ["Overview", "Budget"])
        self.assertEqual(country.received_sheet_names, ["Overview", "Forecast"])
        self.assertEqual(
            {finding.code for finding in country.findings},
            {"SHEET_ADDED", "SHEET_DELETED"},
        )

    def test_identical_kpi_identifiers_compare_across_string_encodings(self) -> None:
        kpis = ["Revenue", "Headcount", "Customer satisfaction"]
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(kpis)},
            {"KPI": make_kpi_sheet(kpis)},
            sent_options={"use_shared_strings": True},
            received_options={"use_shared_strings": False},
        )

        self.assertEqual(self._semantic_codes(country), set())
        self.assertEqual(country.metrics.kpi_sent_count, 3)
        self.assertEqual(country.metrics.kpi_received_count, 3)
        comparison = self._sheet(country, "KPI").kpi_comparison
        self.assertIsNotNone(comparison)
        assert comparison is not None
        self.assertEqual(comparison.missing_count, 0)
        self.assertEqual(comparison.unexpected_count, 0)

    def test_missing_kpi_identifier_is_occurrence_aware(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Margin"])},
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", None])},
        )

        finding = self._one_finding(country, "KPI_IDENTIFIER_MISSING")
        self.assertEqual(finding.category, "KPI_INTEGRITY")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.kpi_missing_count, 1)
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_UNEXPECTED"))

    def test_unexpected_kpi_identifier_is_occurrence_aware(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", None])},
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Margin"])},
        )

        finding = self._one_finding(country, "KPI_IDENTIFIER_UNEXPECTED")
        self.assertEqual(finding.category, "KPI_INTEGRITY")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.kpi_unexpected_count, 1)
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_MISSING"))

    def test_replaced_kpi_identifier_is_missing_and_unexpected(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Margin"])},
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Cash flow"])},
        )

        missing = self._one_finding(country, "KPI_IDENTIFIER_MISSING")
        unexpected = self._one_finding(country, "KPI_IDENTIFIER_UNEXPECTED")
        self.assertEqual((missing.unit_count, unexpected.unit_count), (1, 1))
        structural = {
            "ROWS_INSERTED",
            "ROWS_DELETED",
            "COLUMNS_INSERTED",
            "COLUMNS_DELETED",
        }
        self.assertTrue(structural.isdisjoint(item.code for item in country.findings))
        self.assertFalse(self._findings(country, "PREFILLED_VALUE_CHANGED"))

    def test_duplicate_kpi_identifiers_are_not_collapsed_to_a_set(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Headcount"])},
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", None])},
        )

        missing = self._one_finding(country, "KPI_IDENTIFIER_MISSING")
        self.assertEqual(missing.unit_count, 1)
        comparison = self._sheet(country, "KPI").kpi_comparison
        self.assertIsNotNone(comparison)
        assert comparison is not None and comparison.sent is not None
        self.assertTrue(comparison.sent.duplicate_keys)

    def test_kpi_reorder_is_medium_and_not_add_delete(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount", "Margin"])},
            {"KPI": make_kpi_sheet(["Headcount", "Revenue", "Margin"])},
        )

        finding = self._one_finding(country, "KPI_ORDER_CHANGED")
        self.assertEqual(finding.category, "KPI_INTEGRITY")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_MISSING"))
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_UNEXPECTED"))

    def test_missing_kpi_header_does_not_cascade_into_identifier_findings(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(["Revenue", "Headcount"])},
            {
                "KPI": make_kpi_sheet(
                    ["Revenue", "Headcount"],
                    header="Metric",
                )
            },
        )

        self._one_finding(country, "KPI_HEADER_MISSING")
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_MISSING"))
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_UNEXPECTED"))
        comparison = self._sheet(country, "KPI").kpi_comparison
        self.assertIsNotNone(comparison)
        assert comparison is not None and comparison.received is not None
        self.assertIsNone(comparison.received.header_coordinate)

    def test_ambiguous_kpi_header_does_not_guess_a_column(self) -> None:
        sent = make_kpi_sheet(["Revenue", "Headcount"])
        received = replace_cell(sent, 1, 2, "KPI")

        _run, country = self._pair({"KPI": sent}, {"KPI": received})

        self._one_finding(country, "KPI_HEADER_AMBIGUOUS")
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_MISSING"))
        self.assertFalse(self._findings(country, "KPI_IDENTIFIER_UNEXPECTED"))
        comparison = self._sheet(country, "KPI").kpi_comparison
        self.assertIsNotNone(comparison)
        assert comparison is not None and comparison.received is not None
        self.assertGreaterEqual(len(comparison.received.header_candidates), 2)

    def test_kpi_header_on_non_kpi_sheet_is_ignored(self) -> None:
        sent = make_kpi_sheet(["Revenue", "Headcount"])
        received = make_kpi_sheet(["Revenue", "Different identifier"])

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertFalse(
            [item for item in country.findings if item.category == "KPI_INTEGRITY"]
        )
        self.assertIsNone(self._sheet(country, "Data").kpi_comparison)

    def test_internal_blank_does_not_end_the_kpi_list(self) -> None:
        kpis = ["Revenue", None, "Headcount"]
        _run, country = self._pair(
            {"KPI": make_kpi_sheet(kpis)},
            {"KPI": make_kpi_sheet(kpis)},
        )

        self.assertFalse(
            [item for item in country.findings if item.category == "KPI_INTEGRITY"]
        )
        self.assertEqual(country.metrics.kpi_sent_count, 2)
        self.assertEqual(country.metrics.kpi_received_count, 2)
        comparison = self._sheet(country, "KPI").kpi_comparison
        assert comparison is not None and comparison.sent is not None
        self.assertEqual(
            [entry.display_value for entry in comparison.sent.entries],
            ["Revenue", "Headcount"],
        )

    def test_numeric_kpi_identifiers_compare_by_numeric_value_not_lexeme(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet([1])},
            {"KPI": make_kpi_sheet([1.0])},
        )

        self.assertFalse(
            [item for item in country.findings if item.category == "KPI_INTEGRITY"]
        )

    def test_numeric_and_text_kpi_identifiers_remain_distinct(self) -> None:
        _run, country = self._pair(
            {"KPI": make_kpi_sheet([1])},
            {"KPI": make_kpi_sheet(["1"])},
        )

        self.assertEqual(
            {
                item.code
                for item in country.findings
                if item.category == "KPI_INTEGRITY"
            },
            {"KPI_IDENTIFIER_MISSING", "KPI_IDENTIFIER_UNEXPECTED"},
        )

    def test_formula_cached_kpi_identifier_uses_its_stored_semantic_value(self) -> None:
        _run, country = self._pair(
            {
                "KPI": make_kpi_sheet(
                    [Cell("Revenue", formula='"Revenue"')]
                )
            },
            {"KPI": make_kpi_sheet(["Revenue"])},
        )

        self.assertFalse(
            [item for item in country.findings if item.category == "KPI_INTEGRITY"]
        )
        comparison = self._sheet(country, "KPI").kpi_comparison
        assert comparison is not None and comparison.sent is not None
        self.assertEqual(comparison.sent.entries[0].value_kind, "FORMULA_TEXT")
        self.assertEqual(comparison.sent.entries[0].confidence, "MEDIUM")

    def test_formula_kpi_identifier_without_cache_is_unresolved(self) -> None:
        unresolved = Cell(value=None, formula="A1")
        _run, country = self._pair(
            {"KPI": make_kpi_sheet([unresolved])},
            {"KPI": make_kpi_sheet([unresolved])},
        )

        finding = self._one_finding(country, "KPI_IDENTIFIER_UNRESOLVED")
        self.assertEqual(finding.category, "KPI_INTEGRITY")
        self.assertGreaterEqual(finding.unit_count, 1)

    def test_ref_error_union_counts_cells_once_and_ignores_quoted_tokens(self) -> None:
        sheet = make_grid(rows=9, columns=6, prefix="reference")
        sheet = replace_cell(sheet, 2, 2, error_cell())
        sheet = replace_cell(
            sheet,
            3,
            3,
            error_cell(formula="SUM(#REF!)"),
        )
        sheet = replace_cell(sheet, 4, 4, Cell(0, formula="SUM(#REF!)"))
        sheet = replace_cell(sheet, 5, 5, "#REF!")
        sheet = replace_cell(
            sheet,
            6,
            6,
            Cell(0, formula='IF(A1=1,"#REF!",0)'),
        )
        sheet = replace_cell(
            sheet,
            7,
            1,
            Cell(0, formula="SUM('#REF!'!A1)"),
        )
        sheet = replace_cell(sheet, 8, 2, error_cell("#N/A"))

        _run, country = self._pair({"Data": sheet}, {"Data": sheet})

        metrics = self._sheet(country, "Data").sent_metrics
        self.assertIsNotNone(metrics)
        assert metrics is not None
        self.assertEqual(metrics.ref_error_count, 3)
        self.assertEqual(metrics.cached_ref_error_count, 2)
        self.assertEqual(metrics.formula_ref_error_count, 2)

    def test_ref_error_counts_are_per_sheet_and_increase_is_aggregated(self) -> None:
        sent_data = make_grid(prefix="data")
        sent_data = replace_cell(sent_data, 2, 2, error_cell())
        received_data = replace_cell(
            sent_data,
            4,
            4,
            Cell(0, formula="SUM(#REF!)"),
        )
        clean = make_grid(prefix="clean")

        _run, country = self._pair(
            {"Data": sent_data, "Clean": clean},
            {"Data": received_data, "Clean": clean},
        )

        self.assertEqual(country.metrics.sent_ref_errors, 1)
        self.assertEqual(country.metrics.received_ref_errors, 2)
        data = self._sheet(country, "Data")
        clean_comparison = self._sheet(country, "Clean")
        assert data.sent_metrics is not None and data.received_metrics is not None
        assert clean_comparison.sent_metrics is not None
        assert clean_comparison.received_metrics is not None
        self.assertEqual(
            (data.sent_metrics.ref_error_count, data.received_metrics.ref_error_count),
            (1, 2),
        )
        self.assertEqual(
            (
                clean_comparison.sent_metrics.ref_error_count,
                clean_comparison.received_metrics.ref_error_count,
            ),
            (0, 0),
        )
        finding = self._one_finding(country, "REFERENCE_ERRORS_INCREASED")
        self.assertEqual(finding.category, "FORMULA_INTEGRITY")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(finding.sent_sheet_name, "Data")
        self.assertEqual(finding.received_sheet_name, "Data")


if __name__ == "__main__":
    unittest.main()
