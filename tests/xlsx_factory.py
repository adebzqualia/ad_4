"""Small OOXML workbook factory used by the structural detector tests.

The fixtures deliberately do not depend on openpyxl.  They write the package
parts that matter to the detector directly, making it possible to exercise
shared-string indices, style indices, dimensions, and physical cell addresses
without a spreadsheet library normalising those details for us.
"""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import tempfile
from typing import Literal, Mapping, Sequence, TypeAlias
from xml.etree import ElementTree as ET
import zipfile


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

WORKBOOK_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
WORKSHEET_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
SHARED_STRINGS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)

ET.register_namespace("", SPREADSHEET_NS)
ET.register_namespace("r", DOCUMENT_REL_NS)


Scalar: TypeAlias = str | int | float | bool


@dataclass(frozen=True, slots=True)
class Cell:
    """One stored cell; ``value`` is also the cached value for a formula."""

    value: Scalar | None = None
    formula: str | None = None
    style: str = "default"
    data_type: Literal["auto", "e"] = "auto"
    formula_type: Literal["normal", "shared", "array", "dataTable"] = "normal"
    shared_index: int | None = None
    formula_ref: str | None = None


CellLike: TypeAlias = Cell | Scalar | None
Matrix: TypeAlias = list[list[Cell]]

STYLE_KEYS = ("default", "accent", "bold")
_STYLE_XFS = {
    "default": {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"},
    "accent": {
        "numFmtId": "0",
        "fontId": "0",
        "fillId": "2",
        "borderId": "0",
        "xfId": "0",
        "applyFill": "1",
    },
    "bold": {
        "numFmtId": "0",
        "fontId": "1",
        "fillId": "0",
        "borderId": "0",
        "xfId": "0",
        "applyFont": "1",
    },
}


def _tag(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def read_xlsx_part(path: Path, part_name: str) -> bytes:
    """Read one package part for a focused OOXML fixture extension."""

    with zipfile.ZipFile(path, "r") as archive:
        return archive.read(part_name)


def update_xlsx_parts(path: Path, updates: Mapping[str, bytes]) -> Path:
    """Replace or add package parts without normalizing the workbook."""

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.stem}-",
        suffix=".xlsx",
        dir=path.parent,
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
            temporary_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as target:
            for info in source.infolist():
                if info.filename not in updates:
                    target.writestr(info, source.read(info.filename))
            for part_name, payload in updates.items():
                target.writestr(part_name, payload)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
    return path


def _as_cell(value: CellLike) -> Cell:
    return value if isinstance(value, Cell) else Cell(value=value)


def _matrix(rows: Sequence[Sequence[CellLike]]) -> Matrix:
    result = [[_as_cell(value) for value in row] for row in rows]
    if result and any(len(row) != len(result[0]) for row in result):
        raise ValueError("Worksheet matrices must be rectangular.")
    return result


def make_grid(rows: int = 8, columns: int = 6, prefix: str = "anchor") -> Matrix:
    """Return a dense grid whose values provide stable, unique row/column anchors."""

    if rows < 1 or columns < 1:
        raise ValueError("A grid must have at least one row and one column.")
    styles = STYLE_KEYS
    return [
        [
            Cell(
                value=f"{prefix}-r{row}-c{column}",
                style=styles[(row + column) % len(styles)],
            )
            for column in range(1, columns + 1)
        ]
        for row in range(1, rows + 1)
    ]


def make_kpi_sheet(
    kpis: Sequence[CellLike],
    *,
    header_row: int = 3,
    kpi_column: int = 3,
    header: CellLike = "KPI",
    columns: int = 5,
) -> Matrix:
    """Return a dense KPI worksheet with stable context around its KPI column.

    The surrounding populated cells are intentional.  They let tests exercise
    blank KPI identifiers, value replacements, and header edits without those
    content changes also looking like physical row or column operations.
    """

    if header_row < 1 or kpi_column < 1:
        raise ValueError("KPI header coordinates are one-based.")
    width = max(columns, kpi_column)
    result = make_grid(
        rows=header_row + len(kpis),
        columns=width,
        prefix="kpi-context",
    )
    result[header_row - 1][kpi_column - 1] = _as_cell(header)
    for offset, value in enumerate(kpis, start=1):
        result[header_row + offset - 1][kpi_column - 1] = _as_cell(value)
    return result


def error_cell(
    value: str = "#REF!",
    *,
    formula: str | None = None,
    style: str = "default",
) -> Cell:
    """Return an OOXML error-typed cell, optionally with a formula/cache pair."""

    if not value.startswith("#"):
        raise ValueError("Excel error values must start with '#'.")
    return Cell(value=value, formula=formula, style=style, data_type="e")


def shared_formula_cell(
    value: Scalar | None,
    *,
    shared_index: int,
    formula: str | None = None,
    formula_ref: str | None = None,
    style: str = "default",
) -> Cell:
    """Return a shared-formula master or follower with an optional cache."""

    if shared_index < 0:
        raise ValueError("Shared formula indices cannot be negative.")
    return Cell(
        value=value,
        formula=formula,
        style=style,
        formula_type="shared",
        shared_index=shared_index,
        formula_ref=formula_ref,
    )


def _has_formula(cell: Cell) -> bool:
    return cell.formula is not None or cell.formula_type != "normal"


def replace_cell(
    rows: Sequence[Sequence[CellLike]], row: int, column: int, value: CellLike
) -> Matrix:
    result = _matrix(rows)
    result[row - 1][column - 1] = _as_cell(value)
    return result


def insert_row(
    rows: Sequence[Sequence[CellLike]],
    at: int,
    values: Sequence[CellLike] | None = None,
) -> Matrix:
    result = _matrix(rows)
    width = len(result[0]) if result else len(values or ())
    if not 1 <= at <= len(result) + 1:
        raise ValueError("Row insertion position is outside the matrix.")
    inserted = (
        [_as_cell(value) for value in values]
        if values is not None
        else [Cell(f"inserted-row-{at}-c{column}", style="accent") for column in range(1, width + 1)]
    )
    if len(inserted) != width:
        raise ValueError("Inserted row has the wrong width.")
    result.insert(at - 1, inserted)
    return result


def delete_row(rows: Sequence[Sequence[CellLike]], at: int) -> Matrix:
    result = _matrix(rows)
    if not 1 <= at <= len(result):
        raise ValueError("Row deletion position is outside the matrix.")
    del result[at - 1]
    return result


def insert_column(
    rows: Sequence[Sequence[CellLike]],
    at: int,
    values: Sequence[CellLike] | None = None,
) -> Matrix:
    result = _matrix(rows)
    width = len(result[0]) if result else 0
    if not 1 <= at <= width + 1:
        raise ValueError("Column insertion position is outside the matrix.")
    inserted = (
        [_as_cell(value) for value in values]
        if values is not None
        else [Cell(f"inserted-column-{at}-r{row}", style="bold") for row in range(1, len(result) + 1)]
    )
    if len(inserted) != len(result):
        raise ValueError("Inserted column has the wrong height.")
    for row, value in zip(result, inserted, strict=True):
        row.insert(at - 1, value)
    return result


def delete_column(rows: Sequence[Sequence[CellLike]], at: int) -> Matrix:
    result = _matrix(rows)
    if not result or not 1 <= at <= len(result[0]):
        raise ValueError("Column deletion position is outside the matrix.")
    for row in result:
        del row[at - 1]
    return result


def column_name(index: int) -> str:
    if index < 1:
        raise ValueError("Column indices are one-based.")
    letters: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def _content_types(sheet_count: int, include_shared_strings: bool) -> bytes:
    root = ET.Element(_tag(CONTENT_TYPES_NS, "Types"))
    ET.SubElement(
        root,
        _tag(CONTENT_TYPES_NS, "Default"),
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    ET.SubElement(
        root,
        _tag(CONTENT_TYPES_NS, "Default"),
        Extension="xml",
        ContentType="application/xml",
    )
    ET.SubElement(
        root,
        _tag(CONTENT_TYPES_NS, "Override"),
        PartName="/xl/workbook.xml",
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    )
    ET.SubElement(
        root,
        _tag(CONTENT_TYPES_NS, "Override"),
        PartName="/xl/styles.xml",
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    )
    if include_shared_strings:
        ET.SubElement(
            root,
            _tag(CONTENT_TYPES_NS, "Override"),
            PartName="/xl/sharedStrings.xml",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )
    for index in range(1, sheet_count + 1):
        ET.SubElement(
            root,
            _tag(CONTENT_TYPES_NS, "Override"),
            PartName=f"/xl/worksheets/sheet{index}.xml",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        )
    return _xml_bytes(root)


def _package_relationships() -> bytes:
    root = ET.Element(_tag(PACKAGE_REL_NS, "Relationships"))
    ET.SubElement(
        root,
        _tag(PACKAGE_REL_NS, "Relationship"),
        Id="rId1",
        Type=WORKBOOK_REL,
        Target="xl/workbook.xml",
    )
    return _xml_bytes(root)


def _workbook(sheet_names: Sequence[str]) -> bytes:
    root = ET.Element(_tag(SPREADSHEET_NS, "workbook"))
    sheets = ET.SubElement(root, _tag(SPREADSHEET_NS, "sheets"))
    for index, name in enumerate(sheet_names, start=1):
        ET.SubElement(
            sheets,
            _tag(SPREADSHEET_NS, "sheet"),
            {
                "name": name,
                "sheetId": str(index),
                _tag(DOCUMENT_REL_NS, "id"): f"rId{index}",
            },
        )
    return _xml_bytes(root)


def _workbook_relationships(sheet_count: int, include_shared_strings: bool) -> bytes:
    root = ET.Element(_tag(PACKAGE_REL_NS, "Relationships"))
    for index in range(1, sheet_count + 1):
        ET.SubElement(
            root,
            _tag(PACKAGE_REL_NS, "Relationship"),
            Id=f"rId{index}",
            Type=WORKSHEET_REL,
            Target=f"worksheets/sheet{index}.xml",
        )
    relation_index = sheet_count + 1
    ET.SubElement(
        root,
        _tag(PACKAGE_REL_NS, "Relationship"),
        Id=f"rId{relation_index}",
        Type=STYLES_REL,
        Target="styles.xml",
    )
    if include_shared_strings:
        ET.SubElement(
            root,
            _tag(PACKAGE_REL_NS, "Relationship"),
            Id=f"rId{relation_index + 1}",
            Type=SHARED_STRINGS_REL,
            Target="sharedStrings.xml",
        )
    return _xml_bytes(root)


def _styles(style_order: Sequence[str]) -> bytes:
    if tuple(sorted(style_order)) != tuple(sorted(STYLE_KEYS)):
        raise ValueError(f"style_order must be a permutation of {STYLE_KEYS!r}.")

    root = ET.Element(_tag(SPREADSHEET_NS, "styleSheet"))
    fonts = ET.SubElement(root, _tag(SPREADSHEET_NS, "fonts"), count="2")
    normal_font = ET.SubElement(fonts, _tag(SPREADSHEET_NS, "font"))
    ET.SubElement(normal_font, _tag(SPREADSHEET_NS, "sz"), val="11")
    ET.SubElement(normal_font, _tag(SPREADSHEET_NS, "name"), val="Calibri")
    bold_font = ET.SubElement(fonts, _tag(SPREADSHEET_NS, "font"))
    ET.SubElement(bold_font, _tag(SPREADSHEET_NS, "b"))
    ET.SubElement(bold_font, _tag(SPREADSHEET_NS, "sz"), val="11")
    ET.SubElement(bold_font, _tag(SPREADSHEET_NS, "name"), val="Calibri")

    fills = ET.SubElement(root, _tag(SPREADSHEET_NS, "fills"), count="3")
    for pattern in ("none", "gray125"):
        fill = ET.SubElement(fills, _tag(SPREADSHEET_NS, "fill"))
        ET.SubElement(fill, _tag(SPREADSHEET_NS, "patternFill"), patternType=pattern)
    accent_fill = ET.SubElement(fills, _tag(SPREADSHEET_NS, "fill"))
    pattern_fill = ET.SubElement(
        accent_fill, _tag(SPREADSHEET_NS, "patternFill"), patternType="solid"
    )
    ET.SubElement(pattern_fill, _tag(SPREADSHEET_NS, "fgColor"), rgb="FFFFFF00")
    ET.SubElement(pattern_fill, _tag(SPREADSHEET_NS, "bgColor"), indexed="64")

    borders = ET.SubElement(root, _tag(SPREADSHEET_NS, "borders"), count="1")
    ET.SubElement(borders, _tag(SPREADSHEET_NS, "border"))
    style_xfs = ET.SubElement(root, _tag(SPREADSHEET_NS, "cellStyleXfs"), count="1")
    ET.SubElement(
        style_xfs,
        _tag(SPREADSHEET_NS, "xf"),
        numFmtId="0",
        fontId="0",
        fillId="0",
        borderId="0",
    )
    cell_xfs = ET.SubElement(
        root, _tag(SPREADSHEET_NS, "cellXfs"), count=str(len(style_order))
    )
    for style in style_order:
        ET.SubElement(cell_xfs, _tag(SPREADSHEET_NS, "xf"), _STYLE_XFS[style])
    cell_styles = ET.SubElement(root, _tag(SPREADSHEET_NS, "cellStyles"), count="1")
    ET.SubElement(
        cell_styles,
        _tag(SPREADSHEET_NS, "cellStyle"),
        name="Normal",
        xfId="0",
        builtinId="0",
    )
    return _xml_bytes(root)


def _shared_string_values(sheets: Sequence[tuple[str, Matrix]]) -> list[str]:
    values: list[str] = []
    seen: set[str] = set()
    for _name, rows in sheets:
        for row in rows:
            for cell in row:
                if (
                    cell.data_type == "auto"
                    and not _has_formula(cell)
                    and isinstance(cell.value, str)
                    and cell.value not in seen
                ):
                    seen.add(cell.value)
                    values.append(cell.value)
    return values


def _shared_strings(values: Sequence[str], occurrence_count: int) -> bytes:
    root = ET.Element(
        _tag(SPREADSHEET_NS, "sst"),
        count=str(occurrence_count),
        uniqueCount=str(len(values)),
    )
    for value in values:
        item = ET.SubElement(root, _tag(SPREADSHEET_NS, "si"))
        text = ET.SubElement(item, _tag(SPREADSHEET_NS, "t"))
        if value != value.strip():
            text.set(_tag(XML_NS, "space"), "preserve")
        text.text = value
    return _xml_bytes(root)


def _worksheet(
    rows: Matrix,
    declared_dimension: str,
    style_ids: Mapping[str, int],
    shared_string_ids: Mapping[str, int] | None,
    empty_default_cell: str | None = None,
    full_column_format: bool = False,
) -> bytes:
    root = ET.Element(_tag(SPREADSHEET_NS, "worksheet"))
    ET.SubElement(root, _tag(SPREADSHEET_NS, "dimension"), ref=declared_dimension)
    if full_column_format:
        columns = ET.SubElement(root, _tag(SPREADSHEET_NS, "cols"))
        ET.SubElement(
            columns,
            _tag(SPREADSHEET_NS, "col"),
            min="1",
            max="16384",
            width="12",
            customWidth="1",
        )
    sheet_data = ET.SubElement(root, _tag(SPREADSHEET_NS, "sheetData"))
    for row_index, values in enumerate(rows, start=1):
        row = ET.SubElement(sheet_data, _tag(SPREADSHEET_NS, "row"), r=str(row_index))
        for column_index, cell in enumerate(values, start=1):
            has_formula = _has_formula(cell)
            if cell.value is None and not has_formula:
                continue
            attributes = {
                "r": f"{column_name(column_index)}{row_index}",
                "s": str(style_ids[cell.style]),
            }
            if cell.data_type == "e":
                attributes["t"] = "e"
            elif not has_formula and isinstance(cell.value, str):
                attributes["t"] = "s" if shared_string_ids is not None else "inlineStr"
            elif not has_formula and isinstance(cell.value, bool):
                attributes["t"] = "b"
            elif has_formula and isinstance(cell.value, str):
                attributes["t"] = "str"
            element = ET.SubElement(row, _tag(SPREADSHEET_NS, "c"), attributes)
            if has_formula:
                formula_attributes: dict[str, str] = {}
                if cell.formula_type == "shared":
                    if cell.shared_index is None:
                        raise ValueError("A shared formula cell requires shared_index.")
                    formula_attributes.update(t="shared", si=str(cell.shared_index))
                    if cell.formula_ref is not None:
                        formula_attributes["ref"] = cell.formula_ref
                elif cell.formula_type in {"array", "dataTable"}:
                    if cell.shared_index is not None:
                        raise ValueError(
                            "shared_index is valid only for shared formulas."
                        )
                    formula_attributes["t"] = cell.formula_type
                    if cell.formula_ref is not None:
                        formula_attributes["ref"] = cell.formula_ref
                elif cell.shared_index is not None or cell.formula_ref is not None:
                    raise ValueError(
                        "shared_index and formula_ref are valid only for shared formulas."
                    )
                formula = ET.SubElement(
                    element,
                    _tag(SPREADSHEET_NS, "f"),
                    formula_attributes,
                )
                formula.text = cell.formula
            if (
                cell.data_type == "auto"
                and not has_formula
                and isinstance(cell.value, str)
            ):
                if shared_string_ids is not None:
                    value = ET.SubElement(element, _tag(SPREADSHEET_NS, "v"))
                    value.text = str(shared_string_ids[cell.value])
                else:
                    inline = ET.SubElement(element, _tag(SPREADSHEET_NS, "is"))
                    text = ET.SubElement(inline, _tag(SPREADSHEET_NS, "t"))
                    if cell.value != cell.value.strip():
                        text.set(_tag(XML_NS, "space"), "preserve")
                    text.text = cell.value
            elif cell.value is not None:
                value = ET.SubElement(element, _tag(SPREADSHEET_NS, "v"))
                if isinstance(cell.value, bool):
                    value.text = "1" if cell.value else "0"
                else:
                    value.text = str(cell.value)
    if empty_default_cell:
        row_digits = "".join(character for character in empty_default_cell if character.isdigit())
        empty_row = ET.SubElement(
            sheet_data,
            _tag(SPREADSHEET_NS, "row"),
            r=row_digits,
        )
        ET.SubElement(empty_row, _tag(SPREADSHEET_NS, "c"), r=empty_default_cell)
    return _xml_bytes(root)


def write_xlsx(
    path: Path,
    sheets: Mapping[str, Sequence[Sequence[CellLike]]] | Sequence[tuple[str, Sequence[Sequence[CellLike]]]],
    *,
    dimensions: Mapping[str, str] | None = None,
    use_shared_strings: bool = False,
    reverse_shared_strings: bool = False,
    style_order: Sequence[str] = STYLE_KEYS,
    empty_default_cell: str | None = None,
    full_column_format: bool = False,
) -> Path:
    """Write a minimal, valid XLSX package and return ``path``."""

    sheet_items = list(sheets.items()) if isinstance(sheets, Mapping) else list(sheets)
    normalized = [(name, _matrix(rows)) for name, rows in sheet_items]
    if not normalized:
        raise ValueError("A workbook must contain at least one sheet for these tests.")
    if len({name.casefold() for name, _rows in normalized}) != len(normalized):
        raise ValueError("Sheet names must be unique case-insensitively.")

    style_order = tuple(style_order)
    style_ids = {style: index for index, style in enumerate(style_order)}
    strings = _shared_string_values(normalized) if use_shared_strings else []
    if reverse_shared_strings:
        strings.reverse()
    string_ids = {value: index for index, value in enumerate(strings)} if use_shared_strings else None
    occurrence_count = sum(
        cell.data_type == "auto"
        and not _has_formula(cell)
        and isinstance(cell.value, str)
        for _name, rows in normalized
        for row in rows
        for cell in row
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(len(normalized), use_shared_strings))
        archive.writestr("_rels/.rels", _package_relationships())
        archive.writestr("xl/workbook.xml", _workbook([name for name, _rows in normalized]))
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            _workbook_relationships(len(normalized), use_shared_strings),
        )
        archive.writestr("xl/styles.xml", _styles(style_order))
        if use_shared_strings:
            archive.writestr(
                "xl/sharedStrings.xml", _shared_strings(strings, occurrence_count)
            )
        for index, (name, rows) in enumerate(normalized, start=1):
            width = len(rows[0]) if rows else 1
            actual_dimension = f"A1:{column_name(max(width, 1))}{max(len(rows), 1)}"
            declared_dimension = (dimensions or {}).get(name, actual_dimension)
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                _worksheet(
                    rows,
                    declared_dimension,
                    style_ids,
                    string_ids,
                    empty_default_cell=empty_default_cell,
                    full_column_format=full_column_format,
                ),
            )
    return path
