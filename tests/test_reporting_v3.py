"""Report contracts for formula and prefilled-value integrity findings."""

from __future__ import annotations

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
from xlsx_factory import Cell, make_grid, replace_cell, write_xlsx  # noqa: E402


REPORTING_AVAILABLE = importlib.util.find_spec("pops_anomaly_detector.reporting") is not None
if REPORTING_AVAILABLE:
    from pops_anomaly_detector.reporting import (  # type: ignore[import-not-found]  # noqa: E402
        render_country_report,
        render_global_report,
        write_reports,
    )


@unittest.skipUnless(REPORTING_AVAILABLE, "reporting.py is not present")
class CellIntegrityReportTests(unittest.TestCase):
    def _run_with_cell_findings(self) -> tuple[RunResult, CountryResult]:
        sent_data = make_grid(prefix="integrity")
        sent_data = replace_cell(sent_data, 3, 1, 2)
        sent_data = replace_cell(sent_data, 3, 2, 4)
        sent_data = replace_cell(
            sent_data,
            3,
            3,
            Cell(6, formula="A3+B3", style="accent"),
        )
        sent_data = replace_cell(
            sent_data,
            5,
            5,
            Cell("Approved baseline", style="bold"),
        )
        received_data = replace_cell(
            sent_data,
            3,
            3,
            Cell(6, formula="A3-B3", style="accent"),
        )
        received_data = replace_cell(
            received_data,
            5,
            5,
            Cell("Changed by country", style="bold"),
        )
        clean = make_grid(prefix="clean")

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_dir.mkdir()
            received_dir.mkdir()
            write_xlsx(
                sent_dir / "France.xlsx",
                {"Data": sent_data, "Clean": clean},
            )
            write_xlsx(
                received_dir / "France.xlsx",
                {"Data": received_data, "Clean": clean},
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
    def _sheet(country: CountryResult, name: str):
        matches = [
            comparison
            for comparison in country.sheets
            if comparison.sent_name == name or comparison.received_name == name
        ]
        if len(matches) != 1:
            raise AssertionError(f"Expected one comparison for {name!r}; got {len(matches)}")
        return matches[0]

    def test_per_sheet_and_country_cell_integrity_metrics_are_exact(self) -> None:
        run, country = self._run_with_cell_findings()

        data = self._sheet(country, "Data")
        clean = self._sheet(country, "Clean")
        self.assertEqual(data.formula_changed_count, 1)
        self.assertEqual(data.formula_unresolved_count, 0)
        self.assertEqual(data.value_changed_count, 1)
        self.assertEqual(data.value_unresolved_count, 0)
        self.assertEqual(clean.formula_changed_count, 0)
        self.assertEqual(clean.formula_unresolved_count, 0)
        self.assertEqual(clean.value_changed_count, 0)
        self.assertEqual(clean.value_unresolved_count, 0)
        self.assertEqual(country.metrics.formula_changed_count, 1)
        self.assertEqual(country.metrics.formula_unresolved_count, 0)
        self.assertEqual(country.metrics.value_changed_count, 1)
        self.assertEqual(country.metrics.value_unresolved_count, 0)
        self.assertEqual(run.summary.medium_findings, 2)
        self.assertEqual(run.summary.high_findings, 0)
        self.assertEqual(country.max_anomaly_severity, "MEDIUM")

    def test_html_exposes_cell_integrity_categories_filters_and_sheet_counts(self) -> None:
        run, country = self._run_with_cell_findings()
        html = render_country_report(run, country)

        self.assertIn('id="anomaly-type-filter"', html)
        self.assertIn('id="anomaly-severity-filter"', html)
        self.assertIn('<option value="FORMULA_MODIFIED">', html)
        self.assertIn('<option value="PREFILLED_VALUE_CHANGED">', html)
        self.assertIn('<option value="MEDIUM">MEDIUM</option>', html)
        self.assertIn('data-category="FORMULA_INTEGRITY"', html)
        self.assertIn('data-category="VALUE_INTEGRITY"', html)
        self.assertIn('data-type="FORMULA_MODIFIED"', html)
        self.assertIn('data-type="PREFILLED_VALUE_CHANGED"', html)
        self.assertIsNotNone(
            re.search(r'class="[^"]*\bcell-integrity\b[^"]*"', html)
        )
        self.assertIn("Formula changes", html)
        self.assertIn("Value changes", html)

        global_html = render_global_report(run)
        self.assertIn("FORMULA_MODIFIED", global_html)
        self.assertIn("PREFILLED_VALUE_CHANGED", global_html)
        self.assertIn("Formula integrity", global_html)
        self.assertIn("Value integrity", global_html)

    def test_json_keeps_dynamic_categories_and_per_sheet_metrics(self) -> None:
        run, country = self._run_with_cell_findings()
        payload = run.to_dict()
        serialized_country = payload["countries"][0]
        categories = {
            item["code"]: item["finding_count"]
            for item in serialized_country["category_results"]
        }
        self.assertEqual(categories.get("FORMULA_INTEGRITY"), 1)
        self.assertEqual(categories.get("VALUE_INTEGRITY"), 1)
        self.assertTrue(
            all(
                category in {"FORMULA_INTEGRITY", "VALUE_INTEGRITY"}
                or count == 0
                for category, count in categories.items()
            )
        )
        sheets = {
            item["sent_name"] or item["received_name"]: item
            for item in serialized_country["sheets"]
        }
        self.assertEqual(sheets["Data"]["formula_changed_count"], 1)
        self.assertEqual(sheets["Data"]["value_changed_count"], 1)
        self.assertEqual(sheets["Clean"]["formula_changed_count"], 0)
        self.assertEqual(sheets["Clean"]["value_changed_count"], 0)

        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "reports"
            write_reports(run, output)
            country_json = output / "countries" / Path(country.report_filename).with_suffix(
                ".json"
            ).name
            written = json.loads(country_json.read_text(encoding="utf-8"))
        written_sheets = {
            item["sent_name"] or item["received_name"]: item
            for item in written["sheets"]
        }
        self.assertEqual(written_sheets["Data"]["formula_changed_count"], 1)
        self.assertEqual(written_sheets["Data"]["value_changed_count"], 1)


if __name__ == "__main__":
    unittest.main()
