"""End-to-end tests for structural comparison using raw OOXML fixtures."""

from __future__ import annotations

import importlib.util
import json
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
from pops_anomaly_detector.models import CountryResult, RunResult  # noqa: E402
from xlsx_factory import (  # noqa: E402
    Cell,
    STYLE_KEYS,
    delete_column,
    delete_row,
    insert_column,
    insert_row,
    make_grid,
    replace_cell,
    write_xlsx,
)


REPORTING_AVAILABLE = (
    importlib.util.find_spec("pops_anomaly_detector.reporting") is not None
)
if REPORTING_AVAILABLE:
    from pops_anomaly_detector.reporting import write_reports  # type: ignore[import-not-found]  # noqa: E402


Sheets = Mapping[str, Sequence[Sequence[Cell | str | int | float | bool | None]]]


class StructuralDetectorTests(unittest.TestCase):
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
                sent_dir / "France.xlsx", sent_sheets, **dict(sent_options or {})
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
        return run, run.countries[0]

    @staticmethod
    def _finding(country: CountryResult, code: str):
        matches = [finding for finding in country.findings if finding.code == code]
        if len(matches) != 1:
            raise AssertionError(
                f"Expected one {code} finding, got {len(matches)}: "
                f"{[finding.code for finding in country.findings]}"
            )
        return matches[0]

    def test_unchanged_workbook_is_ok(self) -> None:
        grid = make_grid()

        run, country = self._pair({"Data": grid}, {"Data": grid})

        self.assertEqual(country.overall_status, "OK")
        self.assertEqual(country.comparison_state, "PAIRED")
        self.assertEqual(country.findings, [])
        self.assertEqual(country.sheets[0].status, "UNCHANGED")
        self.assertEqual(country.metrics.sent_sheet_count, 1)
        self.assertEqual(country.metrics.received_sheet_count, 1)
        self.assertEqual(run.summary.ok, 1)
        self.assertEqual(run.summary.error, 0)

    def test_value_and_formula_edits_at_same_coordinates_are_not_structural(self) -> None:
        # A long consecutive edit block is deliberate: a sequence aligner must
        # not find it cheaper to call content substitutions a delete+insert.
        sent = make_grid(rows=16, columns=6)
        received = make_grid(rows=16, columns=6)
        for row in range(4, 14):
            for column in range(1, 7):
                style = STYLE_KEYS[(row + column) % len(STYLE_KEYS)]
                if column in {2, 5}:
                    sent_value = Cell(
                        value=row + column,
                        formula=f"A{row}+{column}",
                        style=style,
                    )
                    received_value = Cell(
                        value=(row + column) * 100,
                        formula=f"SUM(C1:D{row})*{column + 3}",
                        style=style,
                    )
                else:
                    sent_value = Cell(f"sent-r{row}-c{column}", style=style)
                    received_value = Cell(f"received-r{row}-c{column}", style=style)
                sent = replace_cell(sent, row, column, sent_value)
                received = replace_cell(received, row, column, received_value)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        axis_codes = {
            "ROWS_INSERTED",
            "ROWS_DELETED",
            "COLUMNS_INSERTED",
            "COLUMNS_DELETED",
        }
        self.assertTrue(axis_codes.isdisjoint(finding.code for finding in country.findings))
        self.assertEqual(country.metrics.rows_added, 0)
        self.assertEqual(country.metrics.rows_deleted, 0)
        self.assertEqual(country.metrics.columns_added, 0)
        self.assertEqual(country.metrics.columns_deleted, 0)
        self.assertEqual(country.overall_status, "OK")

    def test_sheet_addition_and_deletion_are_both_reported(self) -> None:
        overview = make_grid(prefix="overview")
        data = make_grid(prefix="data")
        forecast = make_grid(prefix="forecast")

        _run, country = self._pair(
            {"Overview": overview, "Data": data},
            {"Overview": overview, "Forecast": forecast},
        )

        self.assertEqual(
            {finding.code for finding in country.findings},
            {"SHEET_ADDED", "SHEET_DELETED"},
        )
        self.assertEqual(country.metrics.sheets_added, 1)
        self.assertEqual(country.metrics.sheets_deleted, 1)
        self.assertEqual(country.metrics.sheet_net_delta, 0)
        self.assertEqual(country.overall_status, "ERROR")
        self.assertTrue(all(finding.severity == "HIGH" for finding in country.findings))

    def test_row_insert_is_located_in_received_coordinates(self) -> None:
        baseline = make_grid(rows=9, columns=6)
        received = insert_row(baseline, 4)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        finding = self._finding(country, "ROWS_INSERTED")
        self.assertEqual((finding.start, finding.end), (4, 4))
        self.assertEqual(finding.coordinate_space, "RECEIVED")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.rows_added, 1)
        self.assertEqual(country.metrics.columns_added, 0)
        self.assertEqual(country.metrics.columns_deleted, 0)

    def test_row_delete_is_located_in_sent_coordinates(self) -> None:
        baseline = make_grid(rows=9, columns=6)
        received = delete_row(baseline, 4)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        finding = self._finding(country, "ROWS_DELETED")
        self.assertEqual((finding.start, finding.end), (4, 4))
        self.assertEqual(finding.coordinate_space, "SENT")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.rows_deleted, 1)
        self.assertEqual(country.metrics.columns_added, 0)
        self.assertEqual(country.metrics.columns_deleted, 0)

    def test_column_insert_is_located_in_received_coordinates(self) -> None:
        baseline = make_grid(rows=9, columns=7)
        received = insert_column(baseline, 3)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        finding = self._finding(country, "COLUMNS_INSERTED")
        self.assertEqual((finding.start, finding.end), (3, 3))
        self.assertEqual(finding.coordinate_space, "RECEIVED")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.columns_added, 1)
        self.assertEqual(country.metrics.rows_added, 0)
        self.assertEqual(country.metrics.rows_deleted, 0)

    def test_column_delete_is_located_in_sent_coordinates(self) -> None:
        baseline = make_grid(rows=9, columns=7)
        received = delete_column(baseline, 3)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        finding = self._finding(country, "COLUMNS_DELETED")
        self.assertEqual((finding.start, finding.end), (3, 3))
        self.assertEqual(finding.coordinate_space, "SENT")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(country.metrics.columns_deleted, 1)
        self.assertEqual(country.metrics.rows_added, 0)
        self.assertEqual(country.metrics.rows_deleted, 0)

    def test_net_zero_row_replacement_reports_delete_and_insert(self) -> None:
        baseline = make_grid(rows=10, columns=6)
        received = insert_row(delete_row(baseline, 3), 7)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        deleted = self._finding(country, "ROWS_DELETED")
        inserted = self._finding(country, "ROWS_INSERTED")
        self.assertEqual((deleted.start, deleted.end), (3, 3))
        self.assertEqual((inserted.start, inserted.end), (7, 7))
        self.assertEqual(country.metrics.rows_deleted, 1)
        self.assertEqual(country.metrics.rows_added, 1)
        self.assertEqual(country.metrics.row_net_delta, 0)
        self.assertEqual(country.overall_status, "ERROR")

    def test_simultaneous_row_and_column_insertions_are_separated(self) -> None:
        baseline = make_grid(rows=9, columns=7)
        received = insert_column(insert_row(baseline, 4), 3)

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        row_finding = self._finding(country, "ROWS_INSERTED")
        column_finding = self._finding(country, "COLUMNS_INSERTED")
        self.assertEqual((row_finding.start, row_finding.end), (4, 4))
        self.assertEqual((column_finding.start, column_finding.end), (3, 3))
        self.assertEqual(country.metrics.rows_added, 1)
        self.assertEqual(country.metrics.columns_added, 1)
        self.assertEqual(country.metrics.rows_deleted, 0)
        self.assertEqual(country.metrics.columns_deleted, 0)

    def test_local_cell_shift_is_unresolved_not_a_whole_row_edit(self) -> None:
        baseline = make_grid(rows=10, columns=6)
        received: list[list[Cell | None]] = [
            [None for _column in range(6)] for _row in range(11)
        ]
        for source_row, values in enumerate(baseline, start=1):
            for source_column, value in enumerate(values, start=1):
                target_row = (
                    source_row + 1
                    if 2 <= source_column <= 5 and source_row >= 4
                    else source_row
                )
                received[target_row - 1][source_column - 1] = value
        for column in range(2, 6):
            received[3][column - 1] = Cell(
                f"local-shift-r4-c{column}",
                style="accent",
            )

        _run, country = self._pair({"Data": baseline}, {"Data": received})

        codes = {finding.code for finding in country.findings}
        self.assertIn("STRUCTURE_UNRESOLVED", codes)
        self.assertNotIn("ROWS_INSERTED", codes)
        self.assertNotIn("ROWS_DELETED", codes)
        self.assertNotIn("COLUMNS_INSERTED", codes)
        self.assertNotIn("COLUMNS_DELETED", codes)
        self.assertEqual(country.overall_status, "ERROR")

    def test_stale_worksheet_dimension_is_diagnostic_only(self) -> None:
        baseline = make_grid(rows=8, columns=6)

        _run, country = self._pair(
            {"Data": baseline},
            {"Data": baseline},
            sent_options={"dimensions": {"Data": "A1:F8"}},
            received_options={"dimensions": {"Data": "A1:XFD1048576"}},
        )

        self.assertEqual(country.overall_status, "OK")
        self.assertEqual(country.findings, [])
        comparison = country.sheets[0]
        self.assertEqual(comparison.sent_metrics.declared_dimension, "A1:F8")
        self.assertEqual(
            comparison.received_metrics.declared_dimension, "A1:XFD1048576"
        )
        self.assertEqual(comparison.sent_metrics.active_rows, 8)
        self.assertEqual(comparison.received_metrics.active_rows, 8)
        self.assertEqual(comparison.sent_metrics.active_columns, 6)
        self.assertEqual(comparison.received_metrics.active_columns, 6)

    def test_default_empty_cell_and_full_column_format_do_not_inflate_extent(self) -> None:
        baseline = make_grid(rows=8, columns=6)

        _run, country = self._pair(
            {"Data": baseline},
            {"Data": baseline},
            received_options={
                "empty_default_cell": "A50000",
                "full_column_format": True,
            },
        )

        self.assertEqual(country.overall_status, "OK")
        self.assertEqual(country.findings, [])
        comparison = country.sheets[0]
        self.assertEqual(comparison.received_metrics.active_rows, 8)
        self.assertEqual(comparison.received_metrics.active_columns, 6)

    def test_reordered_shared_strings_and_style_ids_are_semantically_unchanged(self) -> None:
        baseline = make_grid(rows=8, columns=6)

        _run, country = self._pair(
            {"Data": baseline},
            {"Data": baseline},
            sent_options={
                "use_shared_strings": True,
                "style_order": STYLE_KEYS,
            },
            received_options={
                "use_shared_strings": True,
                "reverse_shared_strings": True,
                "style_order": tuple(reversed(STYLE_KEYS)),
            },
        )

        self.assertNotEqual(country.sent_file.sha256, country.received_file.sha256)
        self.assertEqual(country.overall_status, "OK")
        self.assertEqual(country.findings, [])
        self.assertEqual(country.sheets[0].status, "UNCHANGED")

    def test_missing_and_unexpected_files_are_counted_separately(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_dir.mkdir()
            received_dir.mkdir()
            write_xlsx(sent_dir / "Missing.xlsx", {"Data": make_grid()})
            write_xlsx(received_dir / "Unexpected.xlsx", {"Data": make_grid()})

            run = analyze_directories(
                AnalysisConfig(
                    sent_dir=sent_dir,
                    received_dir=received_dir,
                    output_dir=root / "reports",
                )
            )

        self.assertEqual(run.summary.sent_files, 1)
        self.assertEqual(run.summary.received_files, 1)
        self.assertEqual(run.summary.matched_pairs, 0)
        self.assertEqual(run.summary.missing_received, 1)
        self.assertEqual(run.summary.unexpected_received, 1)
        self.assertEqual(run.summary.error, 2)
        self.assertEqual(
            {country.comparison_state for country in run.countries},
            {"MISSING_RECEIVED", "UNEXPECTED_RECEIVED"},
        )
        self.assertTrue(all(country.overall_status == "ERROR" for country in run.countries))

    @unittest.skipUnless(REPORTING_AVAILABLE, "reporting.py is not present")
    def test_reports_include_global_country_html_and_machine_readable_json(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            report_dir = root / "reports"
            sent_dir.mkdir()
            received_dir.mkdir()
            baseline = make_grid()
            write_xlsx(sent_dir / "France.xlsx", {"Data": baseline})
            write_xlsx(received_dir / "France.xlsx", {"Data": insert_row(baseline, 4)})
            run = analyze_directories(
                AnalysisConfig(
                    sent_dir=sent_dir,
                    received_dir=received_dir,
                    output_dir=report_dir,
                )
            )

            returned = Path(write_reports(run, report_dir))
            self.assertEqual(returned.resolve(), report_dir.resolve())
            global_report = report_dir / "index.html"
            country_report = (
                report_dir / "countries" / run.countries[0].report_filename
            )
            self.assertTrue(global_report.is_file())
            self.assertTrue(country_report.is_file())
            global_html = global_report.read_text(encoding="utf-8")
            country_html = country_report.read_text(encoding="utf-8")
            self.assertIn("France", global_html)
            self.assertIn("France", country_html)
            self.assertIn("structural", country_html.casefold())

            json_paths = sorted(report_dir.glob("*.json"))
            self.assertTrue(json_paths, "Report generation did not emit JSON output.")
            payloads = [json.loads(path.read_text(encoding="utf-8")) for path in json_paths]
            self.assertTrue(
                any(isinstance(payload, dict) and "summary" in payload for payload in payloads),
                "No JSON report contained the run summary.",
            )


if __name__ == "__main__":
    unittest.main()
