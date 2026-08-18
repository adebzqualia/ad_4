"""HTML/JSON contract tests for the expanded anomaly report."""

from __future__ import annotations

from html.parser import HTMLParser
import importlib.util
import json
from pathlib import Path
import re
import sys
import tempfile
import unittest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
TEST_ROOT = Path(__file__).resolve().parent
for import_root in (SRC_ROOT, TEST_ROOT):
    if str(import_root) not in sys.path:
        sys.path.insert(0, str(import_root))

from pops_anomaly_detector import AnalysisConfig, analyze_directories  # noqa: E402
from pops_anomaly_detector.models import CountryResult, RunResult  # noqa: E402
from xlsx_factory import (  # noqa: E402
    Cell,
    error_cell,
    insert_column,
    insert_row,
    make_grid,
    make_kpi_sheet,
    replace_cell,
    write_xlsx,
)


REPORTING_AVAILABLE = importlib.util.find_spec("pops_anomaly_detector.reporting") is not None
if REPORTING_AVAILABLE:
    from pops_anomaly_detector.reporting import (  # type: ignore[import-not-found]  # noqa: E402
        render_country_report,
        write_reports,
    )


class _ReportStructureProbe(HTMLParser):
    """Capture durable report semantics without depending on attribute order."""

    def __init__(self) -> None:
        super().__init__()
        self._details_stack: list[dict[str, str]] = []
        self.sheet_details_inside_type = 0
        self.finding_attributes: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs) -> None:  # type: ignore[no-untyped-def]
        attributes = dict(attrs)
        classes = set(attributes.get("class", "").split())
        if tag == "details":
            if "anomaly-sheet" in classes and any(
                "anomaly-type" in set(ancestor.get("class", "").split())
                for ancestor in self._details_stack
            ):
                self.sheet_details_inside_type += 1
            self._details_stack.append(attributes)
        if "data-finding" in attributes:
            self.finding_attributes.append(attributes)

    def handle_endtag(self, tag: str) -> None:
        if tag == "details" and self._details_stack:
            self._details_stack.pop()


@unittest.skipUnless(REPORTING_AVAILABLE, "reporting.py is not present")
class ExpandedReportTests(unittest.TestCase):
    def _analyze(self, sent_sheets, received_sheets) -> tuple[RunResult, CountryResult]:  # type: ignore[no-untyped-def]
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_dir.mkdir()
            received_dir.mkdir()
            write_xlsx(sent_dir / "France.xlsx", sent_sheets)
            write_xlsx(received_dir / "France.xlsx", received_sheets)
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

    def _mixed_run(self) -> tuple[RunResult, CountryResult]:
        data = make_grid(rows=9, columns=7, prefix="data")
        sent_errors = make_grid(prefix="errors")
        received_errors = replace_cell(
            sent_errors,
            3,
            3,
            error_cell(formula="SUM(#REF!)"),
        )
        run, country = self._analyze(
            {
                "Data": data,
                "KPI": make_kpi_sheet(["Revenue", "Headcount", "Margin"]),
                "Errors": sent_errors,
            },
            {
                "Data": insert_row(data, 4),
                "KPI": make_kpi_sheet(["Headcount", "Revenue", "Margin"]),
                "Errors": received_errors,
            },
        )
        country.warnings.append("Synthetic analysis note for report placement.")
        return run, country

    def test_anomaly_filters_and_nested_type_sheet_dropdowns_have_data_contract(self) -> None:
        run, country = self._mixed_run()
        html = render_country_report(run, country)

        for identifier in (
            'id="anomaly-filters"',
            'id="anomaly-type-filter"',
            'id="anomaly-severity-filter"',
            'id="anomaly-filter-reset"',
            'id="anomaly-filter-count"',
        ):
            self.assertIn(identifier, html)
        self.assertIn("data-type-group", html)
        self.assertIn("data-sheet-group", html)
        for code in (
            "ROWS_INSERTED",
            "KPI_ORDER_CHANGED",
            "REFERENCE_ERRORS_INCREASED",
        ):
            self.assertIn(f'data-type="{code}"', html)

        probe = _ReportStructureProbe()
        probe.feed(html)
        self.assertGreaterEqual(probe.sheet_details_inside_type, 3)
        self.assertEqual(len(probe.finding_attributes), len(country.findings))
        for attributes in probe.finding_attributes:
            for attribute in (
                "data-category",
                "data-type",
                "data-severity",
                "data-sheet",
            ):
                self.assertIn(attribute, attributes)

    def test_report_section_order_notes_inventory_and_wide_sheet_evidence(self) -> None:
        run, country = self._mixed_run()
        html = render_country_report(run, country)

        summary = html.index('id="summary-heading"')
        inventory = html.index('id="sheet-inventory-section"')
        anomalies = html.index('class="section anomaly-category"')
        comparison = html.index('id="sheets-heading"')
        evidence = html.index('id="sources-heading"')
        methodology = html.index('id="method-heading"')
        audit = html.index('id="audit-heading"')
        notes = html.index('id="analysis-notes"')
        self.assertEqual(
            sorted(
                [
                    summary,
                    inventory,
                    anomalies,
                    comparison,
                    evidence,
                    methodology,
                    audit,
                    notes,
                ]
            ),
            [
                summary,
                inventory,
                anomalies,
                comparison,
                evidence,
                methodology,
                audit,
                notes,
            ],
        )
        self.assertIn('details id="sheet-inventory" class="card sheet-inventory"', html)
        self.assertIn("Synthetic analysis note for report placement.", html)
        notes_details = re.search(
            r'<details\s+class="[^"]*\banalysis-notes\b[^"]*"([^>]*)>',
            html,
        )
        self.assertIsNotNone(notes_details)
        assert notes_details is not None
        self.assertNotRegex(notes_details.group(1), r"\bopen(?:\s|=|$)")

        self.assertRegex(
            html,
            r'<tr\s+class="sheet-evidence-row">\s*'
            r'<td\s+colspan="8">\s*'
            r'<details\s+class="sheet-evidence"',
        )
        self.assertIn("#REF! cells", html)
        self.assertIn('class="reference-errors"', html)

    def test_column_locations_render_excel_letters_while_model_stays_numeric(self) -> None:
        baseline = make_grid(rows=9, columns=30, prefix="wide")
        received = insert_column(baseline, 28)
        run, country = self._analyze({"Data": baseline}, {"Data": received})

        findings = [item for item in country.findings if item.code == "COLUMNS_INSERTED"]
        self.assertEqual(len(findings), 1)
        finding = findings[0]
        self.assertEqual((finding.start, finding.end), (28, 28))
        self.assertIn("AB", finding.message)
        html = render_country_report(run, country)
        self.assertRegex(
            html,
            r"<dt>Location</dt>\s*<dd>AB \(received coordinates\)</dd>",
        )

    def test_sheet_inventory_lists_exact_names_even_when_counts_match(self) -> None:
        clean = make_grid(prefix="clean")
        received_errors = replace_cell(clean, 2, 2, Cell(0, formula="SUM(#REF!)"))
        run, country = self._analyze(
            {"Overview": clean, "Budget": clean},
            {"Overview": clean, "Forecast": received_errors},
        )

        html = render_country_report(run, country)
        inventory_start = html.index('id="sheet-inventory-section"')
        anomaly_start = html.index('class="section anomaly-category"')
        inventory_html = html[inventory_start:anomaly_start]
        for name in ("Overview", "Budget", "Forecast"):
            self.assertIn(name, inventory_html)
        self.assertIn("Sheet names differ", inventory_html)
        self.assertIn("Deleted", inventory_html)
        self.assertIn("Added", inventory_html)

    def test_json_category_results_are_derived_from_each_finding_category(self) -> None:
        run, country = self._mixed_run()
        expected_counts: dict[str, int] = {}
        for finding in country.findings:
            expected_counts[finding.category] = expected_counts.get(finding.category, 0) + 1

        payload = run.to_dict()
        serialized_country = payload["countries"][0]
        category_results = serialized_country["category_results"]
        actual_counts = {
            item["code"]: item["finding_count"] for item in category_results
        }
        for category, expected in expected_counts.items():
            self.assertEqual(actual_counts.get(category), expected)
        self.assertTrue(
            all(
                category in expected_counts or count == 0
                for category, count in actual_counts.items()
            )
        )
        self.assertTrue(
            {"STRUCTURAL", "KPI_INTEGRITY", "FORMULA_INTEGRITY"}.issubset(
                actual_counts
            )
        )

        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "reports"
            write_reports(run, output)
            country_json = output / "countries" / Path(country.report_filename).with_suffix(
                ".json"
            ).name
            serialized = json.loads(country_json.read_text(encoding="utf-8"))
        written_counts = {
            item["code"]: item["finding_count"]
            for item in serialized["category_results"]
        }
        for category, expected in expected_counts.items():
            self.assertEqual(written_counts.get(category), expected)
        self.assertTrue(
            all(
                category in expected_counts or count == 0
                for category, count in written_counts.items()
            )
        )


if __name__ == "__main__":
    unittest.main()
