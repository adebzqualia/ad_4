"""Adversarial tests for conservative formula and prefilled-value checks."""

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
    delete_column,
    delete_row,
    insert_column,
    insert_row,
    make_grid,
    replace_cell,
    shared_formula_cell,
    write_xlsx,
)


Sheets = Mapping[str, Sequence[Sequence[CellLike]]]
FORMULA_CODES = {
    "FORMULA_REPLACED_WITH_VALUE",
    "FORMULA_REMOVED",
    "FORMULA_MODIFIED",
}
VALUE_CODES = {"PREFILLED_VALUE_CHANGED"}
CELL_INTEGRITY_CODES = FORMULA_CODES | VALUE_CODES


class CellIntegrityTests(unittest.TestCase):
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

    def _assert_no_cell_integrity_findings(self, country: CountryResult) -> None:
        self.assertTrue(
            CELL_INTEGRITY_CODES.isdisjoint(item.code for item in country.findings),
            [item.code for item in country.findings],
        )

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
    def _formula_sheet(
        formula: str,
        *,
        cached_value: int | float | str | None = 10,
        row: int = 4,
        column: int = 4,
    ):
        grid = make_grid(rows=9, columns=7, prefix="formula")
        grid = replace_cell(grid, row, 2, 4)
        grid = replace_cell(grid, row, 3, 6)
        return replace_cell(
            grid,
            row,
            column,
            Cell(cached_value, formula=formula, style="accent"),
        )

    @staticmethod
    def _shared_formula_sheet(
        multiplier: int,
        caches: Sequence[int],
        *,
        master_row: int = 2,
    ):
        grid = make_grid(rows=8, columns=6, prefix="shared")
        for row in (2, 3, 4):
            grid = replace_cell(grid, row, 1, row)
        if master_row not in {2, 3, 4}:
            raise ValueError("The shared formula master must lie inside B2:B4.")
        for row, cache in zip((2, 3, 4), caches, strict=True):
            is_master = row == master_row
            grid = replace_cell(
                grid,
                row,
                2,
                shared_formula_cell(
                    cache,
                    shared_index=0,
                    formula=f"A{row}*{multiplier}" if is_master else None,
                    formula_ref="B2:B4" if is_master else None,
                    style="bold",
                ),
            )
        return grid

    def test_same_formula_and_cache_has_no_formula_finding(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)

        _run, country = self._pair({"Data": sent}, {"Data": sent})

        self._assert_no_cell_integrity_findings(country)

    def test_formula_cache_change_is_ignored(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)
        received = self._formula_sheet("B4+C4", cached_value=999_999)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)

    def test_formula_replaced_with_hardcoded_value_is_aggregated(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)
        received = replace_cell(sent, 4, 4, Cell(10, style="accent"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(finding.category, "FORMULA_INTEGRITY")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)
        self.assertNotEqual(finding.severity, "HIGH")

    def test_formula_to_absent_or_semantic_blank_is_removed(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)
        for label, replacement in (
            ("absent", None),
            ("empty text", Cell("", style="accent")),
            ("whitespace text", Cell("   ", style="accent")),
        ):
            with self.subTest(replacement=label):
                received = replace_cell(sent, 4, 4, replacement)
                _run, country = self._pair({"Data": sent}, {"Data": received})

                finding = self._one_finding(country, "FORMULA_REMOVED")
                self.assertEqual(finding.category, "FORMULA_INTEGRITY")
                self.assertEqual(finding.severity, "MEDIUM")
                self.assertEqual(finding.unit_count, 1)
                self.assertFalse(
                    self._findings(country, "FORMULA_REPLACED_WITH_VALUE")
                )

    def test_logical_formula_adjustment_after_row_insert_is_not_modified(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)
        sent = replace_cell(sent, 6, 6, Cell("Locked value", style="bold"))
        received = insert_row(sent, 2)
        received = replace_cell(
            received,
            5,
            4,
            Cell(10, formula="B5+C5", style="accent"),
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "ROWS_INSERTED")), 1)
        self._assert_no_cell_integrity_findings(country)

    def test_unadjusted_same_formula_text_after_row_insert_is_modified(self) -> None:
        sent = self._formula_sheet("B4", cached_value=4)
        sent = replace_cell(sent, 6, 6, Cell("Stable anchor", style="bold"))
        # The formula cell moves from D4 to D5, but its reference is deliberately
        # left at B4.  The raw formula text is unchanged while the logical target
        # has changed from sent row 4 to sent row 3.
        received = insert_row(sent, 2)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "ROWS_INSERTED")), 1)
        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)

    def test_logical_formula_adjustment_after_column_insert_is_not_modified(self) -> None:
        sent = self._formula_sheet("B4+C4", cached_value=10)
        sent = replace_cell(sent, 6, 6, Cell("Locked value", style="bold"))
        received = insert_column(sent, 2)
        received = replace_cell(
            received,
            4,
            5,
            Cell(10, formula="C4+D4", style="accent"),
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "COLUMNS_INSERTED")), 1)
        self._assert_no_cell_integrity_findings(country)

    def test_mixed_absolute_formula_adjusts_after_row_and_column_insert(self) -> None:
        sent = self._formula_sheet("$B4+C$3+$B$2", cached_value=10)
        sent = replace_cell(sent, 7, 6, Cell("Stable anchor", style="bold"))
        received = insert_column(insert_row(sent, 2), 2)
        received = replace_cell(
            received,
            5,
            5,
            Cell(10, formula="$C5+D$4+$C$3", style="accent"),
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "ROWS_INSERTED")), 1)
        self.assertEqual(len(self._findings(country, "COLUMNS_INSERTED")), 1)
        self._assert_no_cell_integrity_findings(country)

    def test_fundamental_formula_changes_are_medium_never_high(self) -> None:
        cases = {
            "operator": ("B4+C4", "B4-C4"),
            "function": ("SUM(B4:C4)", "AVERAGE(B4:C4)"),
            "constant": ("B4*2", "B4*3"),
            "reference": ("B4+C4", "B4+E4"),
        }
        for label, (before, after) in cases.items():
            with self.subTest(change=label):
                sent = self._formula_sheet(before, cached_value=10)
                received = self._formula_sheet(after, cached_value=10)
                _run, country = self._pair({"Data": sent}, {"Data": received})

                finding = self._one_finding(country, "FORMULA_MODIFIED")
                self.assertEqual(finding.category, "FORMULA_INTEGRITY")
                self.assertEqual(finding.severity, "MEDIUM")
                self.assertEqual(finding.unit_count, 1)
                self.assertNotEqual(finding.severity, "HIGH")

    def test_compatibility_marker_does_not_hide_fundamental_edit(self) -> None:
        sent = self._formula_sheet("@B4+1", cached_value=5)
        received = self._formula_sheet("@B4-1", cached_value=3)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_compatibility_marker_delta_does_not_hide_independent_edit(self) -> None:
        sent = self._formula_sheet("@B4+1", cached_value=5)
        received = self._formula_sheet("B4+2", cached_value=6)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 0)

    def test_compatibility_marker_text_inside_string_is_not_neutralized(self) -> None:
        sent = self._formula_sheet('"@submitted"', cached_value="@submitted")
        received = self._formula_sheet('"submitted"', cached_value="submitted")

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)

    def test_compatibility_marker_only_delta_is_unresolved(self) -> None:
        sent = self._formula_sheet("@B4", cached_value=4)
        received = self._formula_sheet("B4", cached_value=4)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_changed_count, 0)
        self.assertEqual(comparison.formula_unresolved_count, 1)

    def test_cross_sheet_formula_replaced_with_value_is_detected(self) -> None:
        budget = make_grid(prefix="budget")
        budget = replace_cell(budget, 2, 2, 25)
        sent_data = self._formula_sheet("'Budget'!B2", cached_value=25)
        received_data = replace_cell(sent_data, 4, 4, Cell(25, style="accent"))

        _run, country = self._pair(
            {"Data": sent_data, "Budget": budget},
            {"Data": received_data, "Budget": budget},
        )

        finding = self._one_finding(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(finding.sent_sheet_name, "Data")
        self.assertTrue(any("D4" in item for item in finding.evidence))

    def test_numeric_sheet_name_component_is_not_normalized_as_constant(self) -> None:
        sent_data = self._formula_sheet("'Plan-01'!A1", cached_value=1)
        received_data = self._formula_sheet("'Plan-1'!A1", cached_value=1)
        plan_01 = make_grid(prefix="plan-01")
        plan_1 = make_grid(prefix="plan-1")

        _run, country = self._pair(
            {"Data": sent_data, "Plan-01": plan_01, "Plan-1": plan_1},
            {"Data": received_data, "Plan-01": plan_01, "Plan-1": plan_1},
        )

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)

    def test_unchanged_shared_formula_group_ignores_cache_changes(self) -> None:
        sent = self._shared_formula_sheet(2, [4, 6, 8])
        received = self._shared_formula_sheet(2, [400, 600, 800])

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)

    def test_changed_shared_formula_group_is_one_aggregated_finding(self) -> None:
        sent = self._shared_formula_sheet(2, [4, 6, 8])
        received = self._shared_formula_sheet(3, [6, 9, 12])

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 3)

    def test_follower_first_shared_formula_group_is_resolved_and_compared(self) -> None:
        sent = self._shared_formula_sheet(2, [4, 6, 8], master_row=4)
        received = self._shared_formula_sheet(3, [6, 9, 12], master_row=4)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.unit_count, 3)
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_unresolved_count, 0)

    def test_unresolvable_shared_formula_replacement_is_directly_detected(self) -> None:
        sent = make_grid(prefix="orphan-shared")
        sent = replace_cell(
            sent,
            4,
            4,
            shared_formula_cell(10, shared_index=9, style="accent"),
        )
        received = replace_cell(sent, 4, 4, Cell(10, style="accent"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(finding.unit_count, 1)
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_unresolved_count, 0)

    def test_formula_cells_on_deleted_row_do_not_duplicate_structural_finding(self) -> None:
        sent = make_grid(rows=9, columns=7, prefix="row-delete")
        for column in (2, 3, 4):
            sent = replace_cell(
                sent,
                4,
                column,
                Cell(8, formula=f"A4*{column}", style="accent"),
            )
        received = delete_row(sent, 4)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "ROWS_DELETED")), 1)
        self._assert_no_cell_integrity_findings(country)
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_unresolved_count, 3)

    def test_formula_cells_on_deleted_column_do_not_duplicate_structural_finding(self) -> None:
        sent = make_grid(rows=9, columns=7, prefix="column-delete")
        for row in range(2, 7):
            sent = replace_cell(
                sent,
                row,
                3,
                Cell(row * 2, formula=f"A{row}+B{row}", style="accent"),
            )
        received = delete_column(sent, 3)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "COLUMNS_DELETED")), 1)
        self._assert_no_cell_integrity_findings(country)
        comparison = self._sheet(country, "Data")
        self.assertEqual(comparison.formula_unresolved_count, 5)

    def test_deleted_sheet_counts_skipped_sent_cells_as_unresolved(self) -> None:
        deleted = [
            [Cell(3, formula="1+2", style="accent"), None],
            [None, Cell("Locked value", style="bold")],
        ]
        keep = make_grid(prefix="keep")

        _run, country = self._pair(
            {"Deleted": deleted, "Keep": keep},
            {"Keep": keep},
        )

        self.assertEqual(len(self._findings(country, "SHEET_DELETED")), 1)
        self._assert_no_cell_integrity_findings(country)
        comparison = self._sheet(country, "Deleted")
        self.assertEqual(comparison.formula_unresolved_count, 1)
        self.assertEqual(comparison.value_unresolved_count, 1)
        self.assertEqual(country.metrics.formula_unresolved_count, 1)
        self.assertEqual(country.metrics.value_unresolved_count, 1)

    def test_unresolved_mapping_suppresses_formula_and_value_findings(self) -> None:
        sent = make_grid(rows=10, columns=6, prefix="local-shift")
        sent = replace_cell(
            sent,
            5,
            3,
            Cell(10, formula="A5+B5", style="accent"),
        )
        sent = replace_cell(sent, 7, 4, Cell("Locked value", style="bold"))
        received: list[list[Cell | None]] = [
            [None for _column in range(6)] for _row in range(11)
        ]
        for source_row, values in enumerate(sent, start=1):
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
        received[5][2] = Cell(10, style="accent")
        received[7][3] = Cell("Changed value", style="bold")

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "STRUCTURE_UNRESOLVED")), 1)
        self._assert_no_cell_integrity_findings(country)
        comparison = self._sheet(country, "Data")
        self.assertGreaterEqual(comparison.formula_unresolved_count, 1)
        self.assertGreaterEqual(comparison.value_unresolved_count, 1)

    def test_meaningful_prefilled_literal_change_is_medium(self) -> None:
        sent = make_grid(prefix="value")
        sent = replace_cell(sent, 4, 4, Cell("Approved baseline", style="bold"))
        received = replace_cell(sent, 4, 4, Cell("Changed by country", style="bold"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(finding.category, "VALUE_INTEGRITY")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)
        self.assertNotEqual(finding.severity, "HIGH")

    def test_invalid_array_output_range_does_not_exempt_unrelated_prefill(self) -> None:
        sent = make_grid(prefix="invalid-array-range")
        sent = replace_cell(sent, 1, 1, Cell("Locked value", style="bold"))
        sent = replace_cell(
            sent,
            4,
            4,
            Cell(
                "Locked value",
                formula="A1",
                formula_type="array",
                formula_ref="A1:A2",
                style="accent",
            ),
        )
        received = replace_cell(sent, 1, 1, Cell("Country edit", style="bold"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(finding.unit_count, 1)
        self.assertTrue(any("A1" in item for item in finding.evidence))

    def test_placeholder_prefilled_values_may_be_completed(self) -> None:
        placeholders: list[tuple[str, CellLike]] = [
            ("numeric integer zero", 0),
            ("numeric float zero", 0.0),
            ("text integer zero", "0"),
            ("text decimal zero", "0.0"),
            ("text padded decimal zero", "0.00"),
            ("hyphen", "-"),
            ("whitespace", "   "),
            ("absent", None),
        ]
        for label, placeholder in placeholders:
            with self.subTest(placeholder=label):
                sent = make_grid(prefix=f"placeholder-{label}")
                sent = replace_cell(sent, 4, 4, placeholder)
                received = replace_cell(sent, 4, 4, "Submitted value")
                _run, country = self._pair({"Data": sent}, {"Data": received})

                self.assertFalse(
                    self._findings(country, "PREFILLED_VALUE_CHANGED"),
                    [item.code for item in country.findings],
                )
                self.assertEqual(
                    self._sheet(country, "Data").value_unresolved_count,
                    0,
                )

    def test_typed_numeric_one_and_one_point_zero_are_equal(self) -> None:
        sent = make_grid(prefix="numeric")
        sent = replace_cell(sent, 4, 4, 1)
        received = replace_cell(sent, 4, 4, 1.0)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)

    def test_many_formula_changes_are_aggregated_with_capped_evidence(self) -> None:
        sent = make_grid(rows=55, columns=6, prefix="many-formulas")
        received = make_grid(rows=55, columns=6, prefix="many-formulas")
        for row in range(2, 52):
            sent = replace_cell(sent, row, 1, row)
            received = replace_cell(received, row, 1, row)
            sent = replace_cell(
                sent,
                row,
                4,
                Cell(row * 2, formula=f"A{row}*2", style="accent"),
            )
            received = replace_cell(
                received,
                row,
                4,
                Cell(row * 2, formula=f"A{row}*3", style="accent"),
            )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.unit_count, 50)
        self.assertEqual(country.metrics.formula_changed_count, 50)
        self.assertEqual(self._sheet(country, "Data").formula_changed_count, 50)
        self.assertLess(len(finding.evidence), finding.unit_count)
        self.assertLessEqual(
            sum("omitted" not in item.casefold() for item in finding.evidence),
            40,
        )
        self.assertTrue(any("omitted" in item.casefold() for item in finding.evidence))

    def test_many_value_changes_are_aggregated_with_capped_evidence(self) -> None:
        sent = make_grid(rows=55, columns=6, prefix="many-values")
        received = make_grid(rows=55, columns=6, prefix="many-values")
        for row in range(2, 52):
            sent = replace_cell(sent, row, 4, Cell(f"Locked {row}", style="bold"))
            received = replace_cell(
                received,
                row,
                4,
                Cell(f"Country edit {row}", style="bold"),
            )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(finding.unit_count, 50)
        self.assertEqual(country.metrics.value_changed_count, 50)
        self.assertEqual(self._sheet(country, "Data").value_changed_count, 50)
        self.assertLess(len(finding.evidence), finding.unit_count)
        self.assertLessEqual(
            sum("omitted" not in item.casefold() for item in finding.evidence),
            40,
        )
        self.assertTrue(any("omitted" in item.casefold() for item in finding.evidence))

    def test_cross_sheet_auto_shift_with_quoted_name_is_not_modified(self) -> None:
        budget = make_grid(rows=9, columns=6, prefix="budget-plan")
        sent_data = self._formula_sheet("'Budget Plan'!B4", cached_value=25)
        received_budget = insert_row(budget, 2)
        received_data = replace_cell(
            sent_data,
            4,
            4,
            Cell(25, formula="'Budget Plan'!B5", style="accent"),
        )

        _run, country = self._pair(
            {"Data": sent_data, "Budget Plan": budget},
            {"Data": received_data, "Budget Plan": received_budget},
        )

        self.assertEqual(len(self._findings(country, "ROWS_INSERTED")), 1)
        self._assert_no_cell_integrity_findings(country)

    def test_unadjusted_cross_sheet_reference_after_insert_is_modified(self) -> None:
        budget = make_grid(rows=9, columns=6, prefix="budget-plan-retarget")
        sent_data = self._formula_sheet("'Budget Plan'!B4", cached_value=25)
        received_budget = insert_row(budget, 2)
        # The referenced sheet shifts, but the stored reference is deliberately
        # left at B4 instead of Excel's logical-preserving B5 adjustment.
        received_data = self._formula_sheet("'Budget Plan'!B4", cached_value=25)

        _run, country = self._pair(
            {"Data": sent_data, "Budget Plan": budget},
            {"Data": received_data, "Budget Plan": received_budget},
        )

        self.assertEqual(len(self._findings(country, "ROWS_INSERTED")), 1)
        finding = self._one_finding(country, "FORMULA_MODIFIED")
        self.assertEqual(finding.severity, "MEDIUM")
        self.assertEqual(finding.unit_count, 1)

    def test_shared_whole_column_auto_shift_is_not_modified(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="shared-whole-column")
        for row, cache in zip((2, 3, 4), (12, 18, 24), strict=True):
            sent = replace_cell(
                sent,
                row,
                2,
                shared_formula_cell(
                    cache,
                    shared_index=4,
                    formula="SUM(A:A)*2" if row == 2 else None,
                    formula_ref="B2:B4" if row == 2 else None,
                    style="bold",
                ),
            )
        received = insert_column(sent, 1)
        received = replace_cell(
            received,
            2,
            3,
            shared_formula_cell(
                12,
                shared_index=4,
                formula="SUM(B:B)*2",
                formula_ref="C2:C4",
                style="bold",
            ),
        )

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "COLUMNS_INSERTED")), 1)
        self._assert_no_cell_integrity_findings(country)

    def test_array_formula_output_caches_are_not_prefilled_values(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="array-output")
        sent = replace_cell(
            sent,
            2,
            2,
            Cell(
                4,
                formula="A2:A4*2",
                formula_type="array",
                formula_ref="B2:B4",
                style="accent",
            ),
        )
        sent = replace_cell(sent, 3, 2, Cell(6, style="accent"))
        sent = replace_cell(sent, 4, 2, Cell(8, style="accent"))
        received = replace_cell(
            sent,
            2,
            2,
            Cell(
                40,
                formula="A2:A4*2",
                formula_type="array",
                formula_ref="B2:B4",
                style="accent",
            ),
        )
        received = replace_cell(received, 3, 2, Cell(60, style="accent"))
        received = replace_cell(received, 4, 2, Cell(80, style="accent"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)
        self.assertEqual(self._sheet(country, "Data").value_unresolved_count, 0)

    def test_ref_error_formula_is_not_duplicated_as_manual_modification(self) -> None:
        sent = self._formula_sheet("SUM(B4)", cached_value=4)
        received = self._formula_sheet("SUM(#REF!)", cached_value=4)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self.assertEqual(len(self._findings(country, "REFERENCE_ERRORS_INCREASED")), 1)
        self.assertFalse(self._findings(country, "FORMULA_MODIFIED"))

    def test_explicit_unparsed_formula_replaced_with_value_is_still_direct_evidence(self) -> None:
        sent = self._formula_sheet("XFE1+1", cached_value=10)
        received = replace_cell(sent, 4, 4, Cell(10, style="accent"))

        _run, country = self._pair({"Data": sent}, {"Data": received})

        finding = self._one_finding(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(finding.unit_count, 1)

    def test_external_link_index_change_is_unresolved_not_modified(self) -> None:
        sent = self._formula_sheet("[1]Budget!B4", cached_value=10)
        received = self._formula_sheet("[2]Budget!B4", cached_value=10)

        _run, country = self._pair({"Data": sent}, {"Data": received})

        self._assert_no_cell_integrity_findings(country)
        self.assertEqual(self._sheet(country, "Data").formula_unresolved_count, 1)


if __name__ == "__main__":
    unittest.main()
