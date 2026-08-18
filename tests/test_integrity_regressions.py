"""Cross-cutting regressions for conservative cell-integrity detection.

These cases cover OOXML serialization differences that Excel can introduce
without a country changing workbook logic, plus report/discovery contracts
that need to stay consistent as the detector grows.
"""

from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess
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
from pops_anomaly_detector.reporting import write_reports  # noqa: E402
from xlsx_factory import (  # noqa: E402
    Cell,
    CellLike,
    error_cell,
    make_grid,
    replace_cell,
    shared_formula_cell,
    write_xlsx,
)


Sheets = Mapping[str, Sequence[Sequence[CellLike]]]
MANUAL_FORMULA_CODES = {
    "FORMULA_REPLACED_WITH_VALUE",
    "FORMULA_REMOVED",
    "FORMULA_MODIFIED",
}


class IntegrityRegressionTests(unittest.TestCase):
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

    @staticmethod
    def _formula_sheet(formula: str, cached_value: int | str = 10):
        grid = make_grid(rows=9, columns=7, prefix="formula-regression")
        grid = replace_cell(grid, 4, 2, 4)
        return replace_cell(
            grid,
            4,
            4,
            Cell(cached_value, formula=formula, style="accent"),
        )

    @staticmethod
    def _shared_sheet(*, split_formula: str | None = None):
        grid = make_grid(rows=8, columns=6, prefix="shared-split")
        for row in (2, 3, 4):
            grid = replace_cell(grid, row, 1, row)
            grid = replace_cell(
                grid,
                row,
                2,
                shared_formula_cell(
                    row * 2,
                    shared_index=0,
                    formula="A2*2" if row == 2 else None,
                    formula_ref=("B2:B3" if split_formula else "B2:B4")
                    if row == 2
                    else None,
                    style="bold",
                ),
            )
        if split_formula is not None:
            grid = replace_cell(
                grid,
                4,
                2,
                Cell(8, formula=split_formula, style="bold"),
            )
        return grid

    def test_cached_ref_error_is_counted_but_not_a_manual_formula_change(self) -> None:
        sent = self._formula_sheet("B4+1", cached_value=5)
        received = replace_cell(
            sent,
            4,
            4,
            error_cell(formula="B4+1", style="accent"),
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_changed_count, 0)
        self.assertEqual(comparison.formula_unresolved_count, 0)
        self.assertEqual(country.findings, [])
        self.assertTrue(
            MANUAL_FORMULA_CODES.isdisjoint(
                finding.code for finding in country.findings
            ),
            [finding.code for finding in country.findings],
        )
        assert comparison.received_metrics is not None
        self.assertEqual(comparison.received_metrics.ref_error_count, 1)
        self.assertEqual(comparison.received_metrics.cached_ref_error_count, 1)
        self.assertEqual(comparison.received_metrics.formula_ref_error_count, 0)
        self.assertEqual(country.metrics.received_ref_errors, 1)
        self.assertFalse(self._findings(country, "REFERENCE_ERRORS_INCREASED"))

    def test_dynamic_array_anchorarray_serialization_is_not_modified(self) -> None:
        sent = self._formula_sheet("SUM(B4#)", cached_value=4)
        received = self._formula_sheet(
            "SUM(_xlfn.ANCHORARRAY(B4))",
            cached_value=4,
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertFalse(self._findings(country, "FORMULA_MODIFIED"))
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_changed_count, 0)
        self.assertEqual(comparison.formula_unresolved_count, 0)

    def test_dynamic_array_compatibility_rewrite_does_not_hide_constant_edit(self) -> None:
        sent = self._formula_sheet("SUM(B4#)+1", cached_value=5)
        received = self._formula_sheet(
            "SUM(_xlfn.ANCHORARRAY(B4))+2",
            cached_value=6,
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        findings = self._findings(country, "FORMULA_MODIFIED")
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].unit_count, 1)
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_shared_follower_split_to_equivalent_normal_formula_is_unchanged(self) -> None:
        sent = self._shared_sheet()
        received = self._shared_sheet(split_formula="A4*2")

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertFalse(self._findings(country, "FORMULA_MODIFIED"))
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_changed_count, 0)
        self.assertEqual(comparison.formula_unresolved_count, 0)

    def test_shared_follower_split_to_changed_normal_formula_is_modified(self) -> None:
        sent = self._shared_sheet()
        received = self._shared_sheet(split_formula="A4*3")

        _run, country = self._pair({"Data": sent}, {"Data": received})

        findings = self._findings(country, "FORMULA_MODIFIED")
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].unit_count, 1)
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_textless_orphan_formula_removal_has_direct_evidence(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="orphan-removal")
        sent = replace_cell(
            sent,
            4,
            4,
            shared_formula_cell(None, shared_index=9, style="accent"),
        )
        received = replace_cell(sent, 4, 4, None)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        findings = self._findings(country, "FORMULA_REMOVED")
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].unit_count, 1)
        self.assertTrue(any("D4" in item for item in findings[0].evidence))
        self.assertTrue(
            any("<unresolved formula>" in item for item in findings[0].evidence)
        )
        self.assertTrue(any("<blank>" in item for item in findings[0].evidence))
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_textless_orphan_formula_replacement_has_direct_evidence(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="orphan-replacement")
        sent = replace_cell(
            sent,
            4,
            4,
            shared_formula_cell(None, shared_index=9, style="accent"),
        )
        received = replace_cell(sent, 4, 4, Cell(77, style="accent"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        findings = self._findings(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].unit_count, 1)
        self.assertTrue(any("D4" in item for item in findings[0].evidence))
        self.assertTrue(
            any("<unresolved formula>" in item for item in findings[0].evidence)
        )
        self.assertTrue(any("77" in item for item in findings[0].evidence))
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_to_dict_category_results_match_written_country_json_exactly(self) -> None:
        sent = self._formula_sheet("B4+1", cached_value=5)
        received = self._formula_sheet("B4+2", cached_value=6)
        run, country = self._pair({"Data": sent}, {"Data": received})

        in_memory = run.to_dict()["countries"][0]["category_results"]
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "reports"
            write_reports(run, output)
            country_json = (
                output
                / "countries"
                / Path(country.report_filename).with_suffix(".json").name
            )
            written = json.loads(country_json.read_text(encoding="utf-8"))

        self.assertEqual(in_memory, written["category_results"])
        categories = {item["code"]: item for item in in_memory}
        self.assertEqual(
            categories["FORMULA_INTEGRITY"],
            {
                "code": "FORMULA_INTEGRITY",
                "label": "Formula integrity anomalies",
                "severity": "MEDIUM",
                "status": "WARNING",
                "finding_count": 1,
            },
        )
        self.assertEqual(categories["STRUCTURAL"]["severity"], None)
        self.assertEqual(categories["STRUCTURAL"]["status"], "OK")

    def test_recursive_discovery_preserves_input_relative_evidence_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_book = sent_dir / "emea" / "west" / "France.xlsx"
            received_book = received_dir / "returns" / "France.xlsx"
            sent_book.parent.mkdir(parents=True)
            received_book.parent.mkdir(parents=True)
            grid = make_grid(prefix="recursive-evidence")
            write_xlsx(sent_book, {"Data": grid})
            write_xlsx(received_book, {"Data": grid})

            run = analyze_directories(
                AnalysisConfig(
                    sent_dir=sent_dir,
                    received_dir=received_dir,
                    output_dir=root / "reports",
                    recursive=True,
                )
            )

        self.assertEqual(len(run.countries), 1)
        country = run.countries[0]
        assert country.sent_file is not None and country.received_file is not None
        self.assertEqual(
            country.sent_file.relative_path,
            "sent/emea/west/France.xlsx",
        )
        self.assertEqual(
            country.received_file.relative_path,
            "received/returns/France.xlsx",
        )


@unittest.skipUnless(shutil.which("git"), "git is required to evaluate ignore rules")
class GitIgnoreRegressionTests(unittest.TestCase):
    def test_sensitive_recursive_inputs_are_ignored_at_any_depth(self) -> None:
        sensitive_paths = (
            "data/sent/emea/west/France.xlsx",
            "data/received/returns/France.xlsx",
            "data/sent/emea/west/PORTUGAL.XLSX",
            "data/received/returns/PORTUGAL.XLSX",
        )
        for relative_path in sensitive_paths:
            with self.subTest(path=relative_path):
                result = subprocess.run(
                    [
                        "git",
                        "check-ignore",
                        "--no-index",
                        "--quiet",
                        "--",
                        relative_path,
                    ],
                    cwd=PROJECT_ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(
                    result.returncode,
                    0,
                    f"{relative_path!r} is not protected by .gitignore: {result.stderr}",
                )


if __name__ == "__main__":
    unittest.main()
