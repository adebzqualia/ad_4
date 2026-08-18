"""OOXML regression tests for Excel-generated worksheet output ranges."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
import sys
import tempfile
import unittest
from xml.etree import ElementTree as ET


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
TEST_ROOT = Path(__file__).resolve().parent
for import_root in (SRC_ROOT, TEST_ROOT):
    if str(import_root) not in sys.path:
        sys.path.insert(0, str(import_root))

from pops_anomaly_detector import AnalysisConfig, analyze_directories  # noqa: E402
from pops_anomaly_detector.coordinates import CellRange, parse_range_reference  # noqa: E402
from pops_anomaly_detector.models import CountryResult  # noqa: E402
from pops_anomaly_detector.ooxml import (  # noqa: E402
    WorkbookReadError,
    WorkbookStructure,
    read_workbook,
)
from xlsx_factory import (  # noqa: E402
    CONTENT_TYPES_NS,
    DOCUMENT_REL_NS,
    PACKAGE_REL_NS,
    SPREADSHEET_NS,
    Cell,
    CellLike,
    make_grid,
    make_kpi_sheet,
    read_xlsx_part,
    replace_cell,
    update_xlsx_parts,
    write_xlsx,
)


TABLE_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
)
PIVOT_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
)
QUERY_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
)


def _tag(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _add_override(root: ET.Element, part_name: str, content_type: str) -> None:
    if any(child.attrib.get("PartName") == part_name for child in root):
        return
    ET.SubElement(
        root,
        _tag(CONTENT_TYPES_NS, "Override"),
        PartName=part_name,
        ContentType=content_type,
    )


def _worksheet_relationships(*relations: tuple[str, str, str]) -> bytes:
    root = ET.Element(_tag(PACKAGE_REL_NS, "Relationships"))
    for relation_id, relation_type, target in relations:
        ET.SubElement(
            root,
            _tag(PACKAGE_REL_NS, "Relationship"),
            Id=relation_id,
            Type=relation_type,
            Target=target,
        )
    return _xml_bytes(root)


def _attach_pivot_output(path: Path, reference: str) -> Path:
    worksheet = ET.fromstring(read_xlsx_part(path, "xl/worksheets/sheet1.xml"))
    parts = ET.SubElement(
        worksheet,
        _tag(SPREADSHEET_NS, "pivotTableParts"),
        count="1",
    )
    ET.SubElement(
        parts,
        _tag(SPREADSHEET_NS, "pivotTablePart"),
        {_tag(DOCUMENT_REL_NS, "id"): "rIdPivot"},
    )

    pivot = ET.Element(
        _tag(SPREADSHEET_NS, "pivotTableDefinition"),
        name="Pivot1",
        cacheId="1",
        dataCaption="Values",
    )
    ET.SubElement(
        pivot,
        _tag(SPREADSHEET_NS, "location"),
        ref=reference,
        firstHeaderRow="1",
        firstDataRow="1",
        firstDataCol="1",
    )

    content_types = ET.fromstring(read_xlsx_part(path, "[Content_Types].xml"))
    _add_override(
        content_types,
        "/xl/pivotTables/pivotTable1.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    )
    return update_xlsx_parts(
        path,
        {
            "[Content_Types].xml": _xml_bytes(content_types),
            "xl/worksheets/sheet1.xml": _xml_bytes(worksheet),
            "xl/worksheets/_rels/sheet1.xml.rels": _worksheet_relationships(
                ("rIdPivot", PIVOT_REL, "../pivotTables/pivotTable1.xml")
            ),
            "xl/pivotTables/pivotTable1.xml": _xml_bytes(pivot),
        },
    )


def _attach_table_output(
    path: Path,
    reference: str,
    *,
    query_backed: bool = False,
    calculated_columns: Sequence[int] = (),
    totals_columns: Sequence[int] = (),
    totals_row_count: int = 0,
) -> Path:
    bounds = parse_range_reference(reference)
    width = bounds.max_col - bounds.min_col + 1
    worksheet = ET.fromstring(read_xlsx_part(path, "xl/worksheets/sheet1.xml"))
    parts = ET.SubElement(
        worksheet,
        _tag(SPREADSHEET_NS, "tableParts"),
        count="1",
    )
    ET.SubElement(
        parts,
        _tag(SPREADSHEET_NS, "tablePart"),
        {_tag(DOCUMENT_REL_NS, "id"): "rIdTable"},
    )

    table_attributes = {
        "id": "1",
        "name": "Table1",
        "displayName": "Table1",
        "ref": reference,
        "headerRowCount": "1",
        "totalsRowCount": str(totals_row_count),
        "totalsRowShown": "1" if totals_row_count else "0",
    }
    table = ET.Element(_tag(SPREADSHEET_NS, "table"), table_attributes)
    columns = ET.SubElement(
        table,
        _tag(SPREADSHEET_NS, "tableColumns"),
        count=str(width),
    )
    for index in range(1, width + 1):
        attributes = {"id": str(index), "name": f"Column {index}"}
        if index in totals_columns:
            attributes["totalsRowFunction"] = "sum"
        column = ET.SubElement(
            columns,
            _tag(SPREADSHEET_NS, "tableColumn"),
            attributes,
        )
        if index in calculated_columns:
            formula = ET.SubElement(
                column,
                _tag(SPREADSHEET_NS, "calculatedColumnFormula"),
            )
            formula.text = "[@[Column 1]]*2"

    content_types = ET.fromstring(read_xlsx_part(path, "[Content_Types].xml"))
    _add_override(
        content_types,
        "/xl/tables/table1.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )
    updates = {
        "[Content_Types].xml": _xml_bytes(content_types),
        "xl/worksheets/sheet1.xml": _xml_bytes(worksheet),
        "xl/worksheets/_rels/sheet1.xml.rels": _worksheet_relationships(
            ("rIdTable", TABLE_REL, "../tables/table1.xml")
        ),
        "xl/tables/table1.xml": _xml_bytes(table),
    }
    if query_backed:
        query_table = ET.Element(
            _tag(SPREADSHEET_NS, "queryTable"),
            name="Query1",
            connectionId="1",
        )
        _add_override(
            content_types,
            "/xl/queryTables/queryTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
        )
        updates["[Content_Types].xml"] = _xml_bytes(content_types)
        updates["xl/tables/_rels/table1.xml.rels"] = _worksheet_relationships(
            ("rIdQuery", QUERY_REL, "../queryTables/queryTable1.xml")
        )
        updates["xl/queryTables/queryTable1.xml"] = _xml_bytes(query_table)
    return update_xlsx_parts(path, updates)


Sheets = dict[str, list[list[CellLike]]]
Decorator = Callable[[Path], Path]


class GeneratedOutputTests(unittest.TestCase):
    def _pair(
        self,
        sent_sheets: Sheets,
        received_sheets: Sheets,
        sent_decorator: Decorator,
        received_decorator: Decorator,
    ) -> tuple[CountryResult, WorkbookStructure, WorkbookStructure]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sent_dir = root / "sent"
            received_dir = root / "received"
            sent_dir.mkdir()
            received_dir.mkdir()
            sent_path = sent_decorator(
                write_xlsx(sent_dir / "France.xlsx", sent_sheets)
            )
            received_path = received_decorator(
                write_xlsx(received_dir / "France.xlsx", received_sheets)
            )
            config = AnalysisConfig(
                sent_dir=sent_dir,
                received_dir=received_dir,
                output_dir=root / "reports",
            )
            sent_structure = read_workbook(sent_path, config)
            received_structure = read_workbook(received_path, config)
            run = analyze_directories(config)
        self.assertEqual(len(run.countries), 1)
        country = run.countries[0]
        self.assertEqual(country.comparison_state, "PAIRED", country.errors)
        return country, sent_structure, received_structure

    @staticmethod
    def _findings(country: CountryResult, code: str):
        return [finding for finding in country.findings if finding.code == code]

    @staticmethod
    def _formula_output_sheet(formula_type: str):
        sheet = make_grid(rows=8, columns=6, prefix=f"{formula_type}-output")
        sheet = replace_cell(
            sheet,
            2,
            2,
            Cell(
                100,
                formula="SUM(A2:A4)*2",
                formula_type=formula_type,  # type: ignore[arg-type]
                formula_ref="B2:C4",
                style="accent",
            ),
        )
        follower_values: dict[tuple[int, int], int | str] = {
            (2, 3): "North",
            (3, 2): 30,
            (3, 3): "South",
            (4, 2): 40,
            (4, 3): "West",
        }
        for (row, column), value in follower_values.items():
            sheet = replace_cell(
                sheet,
                row,
                column,
                Cell(value, style=sheet[row - 1][column - 1].style),
            )
        return sheet, follower_values

    def test_pivot_location_masks_generated_values_and_content_anchors(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="pivot")
        received = sent
        for row in range(2, 6):
            for column in range(2, 5):
                received = replace_cell(
                    received,
                    row,
                    column,
                    Cell(
                        f"refreshed-pivot-{row}-{column}",
                        style=sent[row - 1][column - 1].style,
                    ),
                )
        received = replace_cell(
            received,
            7,
            6,
            Cell("Country edit outside pivot", style=sent[6][5].style),
        )
        decorate = lambda path: _attach_pivot_output(path, "B2:D5")

        country, sent_book, received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            decorate,
            decorate,
        )

        sent_sheet = sent_book.sheets[0]
        received_sheet = received_book.sheets[0]
        self.assertEqual(
            sent_sheet.generated_output_ranges,
            [CellRange(2, 2, 5, 4)],
        )
        self.assertFalse(
            any(
                2 <= cell.row <= 5 and 2 <= cell.col <= 4
                for cell in sent_sheet.material_cells
            )
        )
        self.assertEqual(
            [row.digest for row in sent_sheet.rows[1:5]],
            [row.digest for row in received_sheet.rows[1:5]],
        )
        self.assertEqual(
            [column.digest for column in sent_sheet.columns[1:4]],
            [column.digest for column in received_sheet.columns[1:4]],
        )
        value_finding = self._findings(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(len(value_finding), 1)
        self.assertEqual(value_finding[0].unit_count, 1)
        self.assertFalse(self._findings(country, "ROWS_INSERTED"))
        self.assertFalse(self._findings(country, "ROWS_DELETED"))

    def test_pivot_output_expansion_is_not_a_structural_row_or_column_edit(self) -> None:
        sent = make_grid(rows=9, columns=6, prefix="pivot-expansion")
        for column in range(2, 5):
            sent = replace_cell(sent, 6, column, None)
        received = sent
        for row in range(2, 7):
            for column in range(2, 5):
                received = replace_cell(
                    received,
                    row,
                    column,
                    Cell(
                        f"refreshed-pivot-{row}-{column}",
                        style=sent[row - 1][column - 1].style,
                    ),
                )

        country, sent_book, received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            lambda path: _attach_pivot_output(path, "B2:D5"),
            lambda path: _attach_pivot_output(path, "B2:D6"),
        )

        self.assertEqual(
            sent_book.sheets[0].generated_output_ranges,
            [CellRange(2, 2, 5, 4)],
        )
        self.assertEqual(
            received_book.sheets[0].generated_output_ranges,
            [CellRange(2, 2, 6, 4)],
        )
        structural_axis_codes = {
            "ROWS_INSERTED",
            "ROWS_DELETED",
            "COLUMNS_INSERTED",
            "COLUMNS_DELETED",
            "ROW_ALIGNMENT_UNRESOLVED",
            "COLUMN_ALIGNMENT_UNRESOLVED",
        }
        self.assertTrue(
            structural_axis_codes.isdisjoint(
                finding.code for finding in country.findings
            ),
            [finding.code for finding in country.findings],
        )
        self.assertEqual(country.findings, [])

    def test_query_backed_table_masks_its_entire_range(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="query")
        received = sent
        for row in range(1, 7):
            for column in range(1, 4):
                received = replace_cell(
                    received,
                    row,
                    column,
                    Cell(
                        f"refreshed-query-{row}-{column}",
                        style=sent[row - 1][column - 1].style,
                    ),
                )
        decorate = lambda path: _attach_table_output(
            path,
            "A1:C6",
            query_backed=True,
        )

        country, sent_book, received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            decorate,
            decorate,
        )

        self.assertEqual(
            sent_book.sheets[0].generated_output_ranges,
            [CellRange(1, 1, 6, 3)],
        )
        self.assertFalse(self._findings(country, "PREFILLED_VALUE_CHANGED"))
        self.assertEqual(
            [row.digest for row in sent_book.sheets[0].rows],
            [row.digest for row in received_book.sheets[0].rows],
        )

    def test_identical_query_backed_kpi_table_retains_kpi_semantics(self) -> None:
        kpis = ["Revenue", "Headcount", "Customer satisfaction"]
        sheet = make_kpi_sheet(
            kpis,
            header_row=1,
            kpi_column=2,
            columns=4,
        )
        decorate = lambda path: _attach_table_output(
            path,
            "A1:D4",
            query_backed=True,
        )

        country, sent_book, received_book = self._pair(
            {"KPI": sheet},
            {"KPI": sheet},
            decorate,
            decorate,
        )

        self.assertEqual(
            sent_book.sheets[0].generated_output_ranges,
            [CellRange(1, 1, 4, 4)],
        )
        self.assertFalse(self._findings(country, "KPI_HEADER_MISSING"))
        self.assertEqual(country.findings, [])
        comparison = country.sheets[0].kpi_comparison
        self.assertIsNotNone(comparison)
        assert comparison is not None
        self.assertEqual(comparison.status, "OK")
        assert comparison.sent is not None and comparison.received is not None
        self.assertEqual(comparison.sent.status, "FOUND")
        self.assertEqual(comparison.received.status, "FOUND")
        self.assertEqual(
            [entry.display_value for entry in comparison.sent.entries],
            kpis,
        )
        self.assertEqual(
            [entry.display_value for entry in comparison.received.entries],
            kpis,
        )

    def test_table_calculated_column_and_total_are_generated_outputs(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="calculated")
        received = sent
        for row in range(2, 6):
            received = replace_cell(
                received,
                row,
                3,
                Cell(row * 1_000, style=sent[row - 1][2].style),
            )
        received = replace_cell(
            received,
            6,
            2,
            Cell(99_999, style=sent[5][1].style),
        )
        received = replace_cell(
            received,
            3,
            1,
            Cell("Country edit in input column", style=sent[2][0].style),
        )
        decorate = lambda path: _attach_table_output(
            path,
            "A1:C6",
            calculated_columns=(3,),
            totals_columns=(2,),
            totals_row_count=1,
        )

        country, sent_book, _received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            decorate,
            decorate,
        )

        sheet = sent_book.sheets[0]
        self.assertEqual(
            sheet.generated_output_ranges,
            [CellRange(2, 3, 5, 3), CellRange(6, 2, 6, 2)],
        )
        material_coordinates = {cell.coordinate for cell in sheet.material_cells}
        self.assertTrue({"C2", "C3", "C4", "C5", "B6"}.isdisjoint(material_coordinates))
        value_finding = self._findings(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(len(value_finding), 1)
        self.assertEqual(value_finding[0].unit_count, 1)

    def test_array_and_data_table_follower_refreshes_do_not_change_structure(self) -> None:
        for formula_type in ("array", "dataTable"):
            with self.subTest(formula_type=formula_type):
                sent, followers = self._formula_output_sheet(formula_type)
                received = sent
                for offset, ((row, column), value) in enumerate(
                    followers.items(),
                    start=1,
                ):
                    refreshed: int | str
                    if isinstance(value, int):
                        refreshed = value * 1_000
                    else:
                        refreshed = f"Refreshed {offset}"
                    received = replace_cell(
                        received,
                        row,
                        column,
                        Cell(
                            refreshed,
                            style=sent[row - 1][column - 1].style,
                        ),
                    )

                country, sent_book, received_book = self._pair(
                    {"Data": sent},
                    {"Data": received},
                    lambda path: path,
                    lambda path: path,
                )

                sent_sheet = sent_book.sheets[0]
                received_sheet = received_book.sheets[0]
                expected_range = [CellRange(2, 2, 4, 3)]
                self.assertEqual(sent_sheet.generated_output_ranges, expected_range)
                self.assertEqual(received_sheet.generated_output_ranges, expected_range)
                self.assertEqual(sent_sheet.metrics.active_ref, received_sheet.metrics.active_ref)
                self.assertEqual(sent_sheet.metrics.content_ref, received_sheet.metrics.content_ref)
                self.assertEqual(
                    [row.digest for row in sent_sheet.rows],
                    [row.digest for row in received_sheet.rows],
                )
                self.assertEqual(
                    [column.digest for column in sent_sheet.columns],
                    [column.digest for column in received_sheet.columns],
                )
                sent_material = {cell.coordinate for cell in sent_sheet.material_cells}
                self.assertIn("B2", sent_material, "The formula anchor must remain auditable.")
                self.assertTrue(
                    {"C2", "B3", "C3", "B4", "C4"}.isdisjoint(sent_material)
                )
                self.assertFalse(self._findings(country, "PREFILLED_VALUE_CHANGED"))
                self.assertFalse(self._findings(country, "ROWS_INSERTED"))
                self.assertFalse(self._findings(country, "ROWS_DELETED"))
                self.assertFalse(self._findings(country, "COLUMNS_INSERTED"))
                self.assertFalse(self._findings(country, "COLUMNS_DELETED"))
                self.assertEqual(country.findings, [])

    def test_array_anchor_replaced_with_value_remains_directly_detectable(self) -> None:
        sent, _followers = self._formula_output_sheet("array")
        received = replace_cell(sent, 2, 2, Cell(999, style="accent"))

        country, sent_book, _received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            lambda path: path,
            lambda path: path,
        )

        self.assertEqual(
            sent_book.sheets[0].generated_output_ranges,
            [CellRange(2, 2, 4, 3)],
        )
        findings = self._findings(country, "FORMULA_REPLACED_WITH_VALUE")
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].unit_count, 1)
        self.assertTrue(any("B2" in item for item in findings[0].evidence))
        self.assertTrue(any("999" in item for item in findings[0].evidence))

    def test_static_table_values_remain_protected(self) -> None:
        sent = make_grid(rows=8, columns=6, prefix="static-table")
        received = replace_cell(
            sent,
            3,
            2,
            Cell("Country edit", style=sent[2][1].style),
        )
        decorate = lambda path: _attach_table_output(path, "A1:C6")

        country, sent_book, _received_book = self._pair(
            {"Data": sent},
            {"Data": received},
            decorate,
            decorate,
        )

        self.assertEqual(sent_book.sheets[0].generated_output_ranges, [])
        value_finding = self._findings(country, "PREFILLED_VALUE_CHANGED")
        self.assertEqual(len(value_finding), 1)
        self.assertEqual(value_finding[0].unit_count, 1)

    def test_invalid_pivot_location_is_not_trusted_as_a_mask(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = write_xlsx(
                Path(temporary) / "Broken.xlsx",
                {"Data": make_grid(prefix="broken-pivot")},
            )
            _attach_pivot_output(path, "D5:B2")
            with self.assertRaisesRegex(WorkbookReadError, "invalid range"):
                read_workbook(path, AnalysisConfig())


if __name__ == "__main__":
    unittest.main()
