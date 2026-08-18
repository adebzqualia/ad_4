"""Focused tests for conservative formula lexing and normalization."""

from __future__ import annotations

from pathlib import Path
import sys
import unittest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from pops_anomaly_detector.formula_logic import (  # noqa: E402
    FormulaComparisonUnresolved,
    canonicalize_formula,
    compact_formula,
    formula_has_ref_error,
)


class FormulaLogicTests(unittest.TestCase):
    @staticmethod
    def _resolver(sheet: str | None, row: int | None, col: int | None):
        return (f"sheet:{(sheet or 'current').casefold()}", row, col)

    def test_quoted_sheet_spaces_are_preserved_and_resolved(self) -> None:
        seen: list[str | None] = []

        def resolver(sheet: str | None, row: int | None, col: int | None):
            seen.append(sheet)
            return (f"sheet:{(sheet or 'current').casefold()}", row, col)

        identity = canonicalize_formula("'Budget Plan'!B2+1", resolver)

        self.assertEqual(seen, ["Budget Plan"])
        self.assertIn("sheet:budget plan", identity)
        self.assertNotEqual(
            compact_formula("'Budget Plan'!B2"),
            compact_formula("'BudgetPlan'!B2"),
        )

    def test_spaces_around_operators_are_not_intersections(self) -> None:
        spaced = canonicalize_formula(" B2 + C2 ", self._resolver)
        compact = canonicalize_formula("B2+C2", self._resolver)
        self.assertEqual(spaced, compact)

    def test_unsupported_reference_grammars_are_unresolved(self) -> None:
        formulas = (
            "SUM(Sheet1:Sheet3!A1)",
            "SUM('Sheet 1:Sheet 3'!A1)",
            "SUM(Table1[A1])",
            "SUM(A:A 1:1)",
            "XFE1+1",
        )
        for formula in formulas:
            with self.subTest(formula=formula):
                with self.assertRaises(FormulaComparisonUnresolved):
                    canonicalize_formula(formula, self._resolver)

    def test_ref_error_scan_ignores_strings_and_quoted_sheet_names(self) -> None:
        self.assertTrue(formula_has_ref_error("SUM(#REF!)"))
        self.assertFalse(formula_has_ref_error('IFERROR(A1,"#REF!")'))
        self.assertFalse(formula_has_ref_error("'#REF!'!A1"))


if __name__ == "__main__":
    unittest.main()
