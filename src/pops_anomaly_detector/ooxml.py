"""Read structural worksheet evidence directly from the XLSX OOXML package.

The parser is intentionally read-only.  It treats worksheet ``dimension`` as
diagnostic metadata and builds its own active structure from stored cells,
row/column properties, ranges, tables, and drawing anchors.
"""

from __future__ import annotations

import hashlib
import json
import math
import posixpath
import re
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import BinaryIO, Iterable
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from .config import AnalysisConfig
from .coordinates import (
    MAX_EXCEL_COLUMN,
    MAX_EXCEL_ROW,
    CellRange,
    format_active_ref,
    index_to_column,
    parse_cell_reference,
    parse_range_reference,
)
from .models import CellSnapshot, KpiColumnSnapshot, KpiEntry, SheetMetrics


class WorkbookReadError(RuntimeError):
    """A workbook cannot be safely or faithfully analyzed."""


@dataclass(frozen=True, slots=True)
class AxisSignature:
    index: int
    weights: dict[str, float]
    strong_tokens: frozenset[str]
    digest: str
    information: float


@dataclass(slots=True)
class SheetStructure:
    name: str
    index: int
    state: str
    sheet_type: str
    part_name: str
    metrics: SheetMetrics
    rows: list[AxisSignature]
    columns: list[AxisSignature]
    cell_anchors: dict[str, list[tuple[int, int]]]
    material_cells: list[CellSnapshot]
    generated_output_ranges: list[CellRange]
    kpi_snapshot: KpiColumnSnapshot
    ref_error_coordinates: list[str] = field(default_factory=list)
    cached_ref_error_coordinates: list[str] = field(default_factory=list)
    formula_ref_error_coordinates: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class WorkbookStructure:
    path: Path
    sheets: list[SheetStructure]
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class _RangeFeature:
    kind: str
    bounds: CellRange
    strong: bool = False


@dataclass(frozen=True, slots=True)
class _SemanticCell:
    row: int
    col: int
    display_value: str
    comparison_key: str | None
    value_kind: str
    confidence: str
    literal_text: bool = False


@dataclass(slots=True)
class _SharedFormulaGroup:
    cell_indices: list[int] = field(default_factory=list)
    master_indices: list[int] = field(default_factory=list)


def _merged_intervals(
    intervals: Iterable[tuple[int, int]],
) -> tuple[tuple[int, int], ...]:
    merged: list[list[int]] = []
    for start, end in sorted(intervals):
        if not merged or start > merged[-1][1] + 1:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return tuple((start, end) for start, end in merged)


def _interval_contains(intervals: tuple[tuple[int, int], ...], value: int) -> bool:
    low = 0
    high = len(intervals)
    while low < high:
        middle = (low + high) // 2
        if intervals[middle][0] <= value:
            low = middle + 1
        else:
            high = middle
    return low > 0 and intervals[low - 1][1] >= value


@dataclass(slots=True)
class _GeneratedOutputIndex:
    """Fast point queries for generated rectangles in Excel's fixed grid."""

    row_intervals_by_column_node: dict[int, tuple[tuple[int, int], ...]]

    @classmethod
    def build(cls, ranges: list[CellRange]) -> _GeneratedOutputIndex:
        buckets: dict[int, list[tuple[int, int]]] = defaultdict(list)

        def register(
            node: int,
            left: int,
            right: int,
            query_left: int,
            query_right: int,
            row_interval: tuple[int, int],
        ) -> None:
            if query_left <= left and right <= query_right:
                buckets[node].append(row_interval)
                return
            middle = (left + right) // 2
            if query_left <= middle:
                register(
                    node * 2,
                    left,
                    middle,
                    query_left,
                    query_right,
                    row_interval,
                )
            if query_right > middle:
                register(
                    node * 2 + 1,
                    middle + 1,
                    right,
                    query_left,
                    query_right,
                    row_interval,
                )

        for bounds in ranges:
            register(
                1,
                1,
                MAX_EXCEL_COLUMN,
                bounds.min_col,
                bounds.max_col,
                (bounds.min_row, bounds.max_row),
            )
        return cls(
            row_intervals_by_column_node={
                node: _merged_intervals(intervals)
                for node, intervals in buckets.items()
            }
        )

    def contains(self, row: int, col: int) -> bool:
        node = 1
        left = 1
        right = MAX_EXCEL_COLUMN
        while True:
            intervals = self.row_intervals_by_column_node.get(node, ())
            if _interval_contains(intervals, row):
                return True
            if left == right:
                return False
            middle = (left + right) // 2
            if col <= middle:
                node = node * 2
                right = middle
            else:
                node = node * 2 + 1
                left = middle + 1

@dataclass(slots=True)
class _SheetAccumulator:
    row_weights: dict[int, dict[str, float]] = field(
        default_factory=lambda: defaultdict(dict)
    )
    column_weights: dict[int, dict[str, float]] = field(
        default_factory=lambda: defaultdict(dict)
    )
    row_strong: dict[int, set[str]] = field(
        default_factory=lambda: defaultdict(set)
    )
    column_strong: dict[int, set[str]] = field(
        default_factory=lambda: defaultdict(set)
    )
    cell_anchors: dict[str, list[tuple[int, int]]] = field(
        default_factory=lambda: defaultdict(list)
    )
    ranges: list[_RangeFeature] = field(default_factory=list)
    column_properties: list[tuple[int, int, str]] = field(default_factory=list)
    content_rows: set[int] = field(default_factory=set)
    content_columns: set[int] = field(default_factory=set)
    active_max_row: int = 0
    active_max_column: int = 0
    content_min_row: int = MAX_EXCEL_ROW + 1
    content_min_column: int = MAX_EXCEL_COLUMN + 1
    content_max_row: int = 0
    content_max_column: int = 0
    cell_count: int = 0
    formula_count: int = 0
    styled_blank_count: int = 0
    stored_cell_count: int = 0
    merged_count: int = 0
    table_count: int = 0
    declared_dimension: str | None = None
    capture_kpi: bool = False
    kpi_semantic_cell_limit: int = 0
    material_cell_limit: int = 0
    material_cells: list[CellSnapshot] = field(default_factory=list)
    kpi_semantic_cells: list[CellSnapshot] = field(default_factory=list)
    generated_output_ranges: list[CellRange] = field(default_factory=list)
    generated_output_range_set: set[CellRange] = field(default_factory=set)
    generated_output_index: _GeneratedOutputIndex | None = None
    shared_formula_groups: dict[str, _SharedFormulaGroup] = field(default_factory=dict)
    ref_error_count: int = 0
    cached_ref_error_count: int = 0
    formula_ref_error_count: int = 0
    ref_error_coordinates: list[str] = field(default_factory=list)
    cached_ref_error_coordinates: list[str] = field(default_factory=list)
    formula_ref_error_coordinates: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _attribute(element: ET.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _hash_payload(payload: object, length: int = 24) -> str:
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:length]


def _canonical_xml(element: ET.Element) -> object:
    false_default_attributes = {
        "applyAlignment",
        "applyBorder",
        "applyFill",
        "applyFont",
        "applyNumberFormat",
        "applyProtection",
        "bestFit",
        "collapsed",
        "customFormat",
        "customHeight",
        "diagonalDown",
        "diagonalUp",
        "hidden",
        "justifyLastLine",
        "outline",
        "pivotButton",
        "quotePrefix",
        "shadow",
        "shrinkToFit",
        "strike",
        "wrapText",
    }
    zero_default_attributes = {"indent", "relativeIndent", "textRotation"}
    attributes = sorted(
        (_local_name(key), value)
        for key, value in element.attrib.items()
        if not (
            value == "0"
            and _local_name(key) in false_default_attributes | zero_default_attributes
        )
    )
    children = [_canonical_xml(child) for child in list(element)]
    text = (element.text or "").strip()
    return (_local_name(element.tag), attributes, text, children)


def _xml_root(archive: zipfile.ZipFile, part_name: str) -> ET.Element:
    with archive.open(part_name) as stream:
        return ET.parse(stream).getroot()


_REF_ERROR_COORDINATE_LIMIT = 200


def _normalize_display_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    return " ".join(value.split())


def _normalize_text(value: str) -> str:
    return _normalize_display_text(value).casefold()


def _formula_contains_ref_error(formula: str) -> bool:
    """Return whether a formula contains an unquoted ``#REF!`` token.

    Excel uses double quotes for string literals and single quotes for sheet or
    external-workbook identifiers. Both quote forms escape themselves by
    doubling the quote character. A lexical scan avoids counting text such as
    ``IFERROR(A1, "#REF!")`` or a reference to a sheet named ``#REF!``.
    """

    index = 0
    in_double_string = False
    in_single_identifier = False
    while index < len(formula):
        char = formula[index]
        if in_double_string:
            if char == '"':
                if index + 1 < len(formula) and formula[index + 1] == '"':
                    index += 2
                    continue
                in_double_string = False
            index += 1
            continue
        if in_single_identifier:
            if char == "'":
                if index + 1 < len(formula) and formula[index + 1] == "'":
                    index += 2
                    continue
                in_single_identifier = False
            index += 1
            continue
        if char == '"':
            in_double_string = True
            index += 1
            continue
        if char == "'":
            in_single_identifier = True
            index += 1
            continue
        if formula[index : index + 5].casefold() == "#ref!":
            return True
        index += 1
    return False


def _rich_text(element: ET.Element) -> str:
    parts: list[str] = []

    def visit(node: ET.Element) -> None:
        name = _local_name(node.tag)
        if name in {"rPh", "phoneticPr"}:
            return
        if name == "t":
            parts.append(node.text or "")
            return
        for child in node:
            visit(child)

    visit(element)
    return "".join(parts)


_A1_REFERENCE_RE = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:(?P<sheet>'(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<first>\$?[A-Za-z]{1,3}\$?[0-9]{1,7})"
    r"(?::(?P<second>\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?"
    r"(?![A-Za-z0-9_.(])"
)


def _formula_segments(formula: str) -> Iterable[tuple[bool, str]]:
    """Yield (is_string_literal, text), respecting doubled Excel quotes."""

    start = 0
    index = 0
    in_string = False
    while index < len(formula):
        if formula[index] != '"':
            index += 1
            continue
        if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
            index += 2
            continue
        if index > start:
            yield in_string, formula[start:index]
        yield True, '"'
        in_string = not in_string
        index += 1
        start = index
    if start < len(formula):
        yield in_string, formula[start:]


def _reference_shape(reference: str, base_row: int, base_col: int) -> tuple[str, str]:
    row, col = parse_cell_reference(reference)
    col_absolute = reference.startswith("$")
    row_absolute = "$" in reference[1:]
    shape = f"C{'$' if col_absolute else '~'}R{'$' if row_absolute else '~'}"
    column_component = f"C${col}" if col_absolute else f"C[{col - base_col}]"
    row_component = f"R${row}" if row_absolute else f"R[{row - base_row}]"
    relative = f"{column_component}{row_component}"
    return shape, relative


def _normalize_formula(formula: str, row: int, col: int) -> tuple[str, str]:
    shape_parts: list[str] = []
    relative_parts: list[str] = []

    def replace(segment: str, relative: bool) -> str:
        def callback(match: re.Match[str]) -> str:
            sheet = (match.group("sheet") or "").casefold()
            first_shape, first_relative = _reference_shape(match.group("first"), row, col)
            second = match.group("second")
            if second:
                second_shape, second_relative = _reference_shape(second, row, col)
                body = (
                    f"{first_relative}:{second_relative}"
                    if relative
                    else f"{first_shape}:{second_shape}"
                )
            else:
                body = first_relative if relative else first_shape
            return f"REF[{sheet}!{body}]"

        return _A1_REFERENCE_RE.sub(callback, segment)

    for is_string, segment in _formula_segments(formula):
        if is_string:
            normalized = segment
            shape_parts.append(normalized)
            relative_parts.append(normalized)
        else:
            compact = re.sub(r"\s+", "", segment).upper()
            shape_parts.append(replace(compact, relative=False))
            relative_parts.append(replace(compact, relative=True))
    return "".join(shape_parts), "".join(relative_parts)


def _add_weight(target: dict[int, dict[str, float]], index: int, token: str, weight: float) -> None:
    bucket = target[index]
    bucket[token] = bucket.get(token, 0.0) + weight


def _mark_active(acc: _SheetAccumulator, row: int, col: int) -> None:
    acc.active_max_row = max(acc.active_max_row, row)
    acc.active_max_column = max(acc.active_max_column, col)


def _mark_content(acc: _SheetAccumulator, row: int, col: int) -> None:
    acc.content_rows.add(row)
    acc.content_columns.add(col)
    acc.content_min_row = min(acc.content_min_row, row)
    acc.content_min_column = min(acc.content_min_column, col)
    acc.content_max_row = max(acc.content_max_row, row)
    acc.content_max_column = max(acc.content_max_column, col)


class _StyleCatalog:
    def __init__(self, fingerprints: list[str]):
        self._fingerprints = fingerprints or [_hash_payload(("default",))]

    def get(self, index: int, warnings: list[str]) -> str:
        if 0 <= index < len(self._fingerprints):
            return self._fingerprints[index]
        raise WorkbookReadError(f"A cell references missing style index {index}.")


def _parse_styles(archive: zipfile.ZipFile, part_name: str | None) -> _StyleCatalog:
    if not part_name or part_name not in archive.namelist():
        return _StyleCatalog([])
    try:
        root = _xml_root(archive, part_name)
    except ET.ParseError as exc:
        raise WorkbookReadError(f"Invalid styles XML ({part_name}): {exc}") from exc

    sections: dict[str, list[ET.Element]] = defaultdict(list)
    num_formats: dict[str, str] = {}
    for child in root:
        name = _local_name(child.tag)
        sections[name] = list(child)
        if name == "numFmts":
            for item in child:
                num_id = _attribute(item, "numFmtId")
                code = _attribute(item, "formatCode")
                if num_id is not None and code is not None:
                    num_formats[num_id] = code

    xfs = sections.get("cellXfs", [])
    style_xfs = sections.get("cellStyleXfs", [])

    def referenced(section: str, raw_index: str | None) -> object | None:
        if raw_index is None:
            return None
        try:
            item = sections.get(section, [])[int(raw_index)]
        except (ValueError, IndexError):
            return ("invalid-reference", section, raw_index)
        return _canonical_xml(item)

    def semantic_xf(xf: ET.Element, include_base: bool) -> object:
        attributes = {_local_name(key): value for key, value in xf.attrib.items()}
        font = referenced("fonts", attributes.pop("fontId", None))
        fill = referenced("fills", attributes.pop("fillId", None))
        border = referenced("borders", attributes.pop("borderId", None))
        number_id = attributes.pop("numFmtId", None)
        number_format = num_formats.get(number_id or "", f"builtin:{number_id}")
        base_id = attributes.pop("xfId", None)
        for key in list(attributes):
            if attributes[key] == "0" and key in {
                "applyAlignment",
                "applyBorder",
                "applyFill",
                "applyFont",
                "applyNumberFormat",
                "applyProtection",
                "pivotButton",
                "quotePrefix",
            }:
                attributes.pop(key)
        base = None
        if include_base and base_id is not None:
            try:
                base = semantic_xf(style_xfs[int(base_id)], include_base=False)
            except (ValueError, IndexError):
                base = ("invalid-base", base_id)
        return {
            "attributes": sorted(attributes.items()),
            "font": font,
            "fill": fill,
            "border": border,
            "number_format": number_format,
            "base": base,
            "children": [_canonical_xml(child) for child in xf],
        }

    fingerprints: list[str] = []
    for xf in xfs:
        fingerprints.append(_hash_payload(semantic_xf(xf, include_base=True)))
    return _StyleCatalog(fingerprints)


def _parse_shared_strings(archive: zipfile.ZipFile, part_name: str | None) -> list[str]:
    if not part_name or part_name not in archive.namelist():
        return []
    strings: list[str] = []
    try:
        with archive.open(part_name) as stream:
            for event, element in ET.iterparse(stream, events=("end",)):
                if _local_name(element.tag) != "si":
                    continue
                strings.append(_rich_text(element))
                element.clear()
    except ET.ParseError as exc:
        raise WorkbookReadError(f"Invalid shared strings XML ({part_name}): {exc}") from exc
    return strings


def _relationships(archive: zipfile.ZipFile, rels_part: str) -> dict[str, tuple[str, str]]:
    if rels_part not in archive.namelist():
        return {}
    try:
        root = _xml_root(archive, rels_part)
    except ET.ParseError as exc:
        raise WorkbookReadError(f"Invalid relationships XML ({rels_part}): {exc}") from exc
    result: dict[str, tuple[str, str]] = {}
    for relation in root:
        if _local_name(relation.tag) != "Relationship":
            continue
        if (_attribute(relation, "TargetMode") or "").lower() == "external":
            continue
        rel_id = _attribute(relation, "Id")
        target = _attribute(relation, "Target")
        rel_type = _attribute(relation, "Type") or ""
        if rel_id and target:
            result[rel_id] = (target, rel_type)
    return result


def _rels_part(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _resolve_part(source_part: str, target: str) -> str:
    target = unquote(target).replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _worksheet_object_relationship_ids(
    archive: zipfile.ZipFile,
    part_name: str,
    sheet_name: str,
    acc: _SheetAccumulator,
) -> tuple[list[str], list[str]]:
    """Pre-read relationship ids and declared calculated-output ranges.

    Output ranges must be known before cells are tokenized. A bounded SAX
    pre-pass avoids trusting orphaned relationships while keeping the worksheet
    itself out of memory.
    """

    table_ids: list[str] = []
    pivot_ids: list[str] = []
    seen: set[tuple[str, str]] = set()
    current_row = 0
    next_row = 1
    next_column = 1
    in_row = False
    current_cell: tuple[int, int] | None = None
    stream: BinaryIO | None = None
    try:
        stream = archive.open(part_name)
        for event, element in ET.iterparse(stream, events=("start", "end")):
            if event == "start":
                tag = _local_name(element.tag)
                if tag == "row":
                    raw_row = _attribute(element, "r")
                    try:
                        current_row = int(raw_row) if raw_row else next_row
                    except ValueError as exc:
                        raise WorkbookReadError(
                            f"Worksheet {sheet_name!r} contains an invalid row index."
                        ) from exc
                    if not 1 <= current_row <= MAX_EXCEL_ROW:
                        raise WorkbookReadError(
                            f"Worksheet {sheet_name!r} contains out-of-range row {current_row}."
                        )
                    in_row = True
                    next_row = current_row + 1
                    next_column = 1
                    current_cell = None
                elif tag == "c":
                    if not in_row:
                        raise WorkbookReadError(
                            f"Worksheet {sheet_name!r} contains a cell outside a row element."
                        )
                    reference = _attribute(element, "r")
                    if reference:
                        try:
                            row, col = parse_cell_reference(reference)
                        except ValueError as exc:
                            raise WorkbookReadError(
                                f"Worksheet {sheet_name!r} contains invalid cell reference "
                                f"{reference!r}."
                            ) from exc
                    else:
                        row, col = current_row, next_column
                    if row != current_row:
                        raise WorkbookReadError(
                            f"Cell {reference!r} is inconsistent with enclosing row "
                            f"{current_row} in worksheet {sheet_name!r}."
                        )
                    current_cell = (row, col)
                    next_column = col + 1
                elif tag == "f":
                    formula_type = (
                        _attribute(element, "t") or "normal"
                    ).strip().casefold()
                    if formula_type in {"array", "datatable"}:
                        formula_range = _attribute(element, "ref")
                        if current_cell is None or not formula_range:
                            acc.warnings.append(
                                f"Ignored {formula_type} output mask without a bounded "
                                f"anchor/range in worksheet {sheet_name!r}."
                            )
                            continue
                        try:
                            bounds = parse_range_reference(formula_range)
                        except ValueError:
                            acc.warnings.append(
                                f"Ignored invalid {formula_type} output range "
                                f"{formula_range!r} in worksheet {sheet_name!r}."
                            )
                            continue
                        row, col = current_cell
                        if not (
                            bounds.min_row <= row <= bounds.max_row
                            and bounds.min_col <= col <= bounds.max_col
                        ):
                            acc.warnings.append(
                                f"Ignored {formula_type} output range {formula_range!r}: "
                                f"anchor {index_to_column(col)}{row} lies outside it."
                            )
                            continue
                        _add_generated_output_range(acc, bounds)
                elif tag in {"tablePart", "pivotTablePart"}:
                    relation_id = _attribute(element, "id")
                    if not relation_id:
                        raise WorkbookReadError(
                            f"Worksheet {sheet_name!r} contains a {tag} without a relationship id."
                        )
                    key = (tag, relation_id)
                    if key in seen:
                        raise WorkbookReadError(
                            f"Worksheet {sheet_name!r} repeats {tag} relationship {relation_id!r}."
                        )
                    seen.add(key)
                    (table_ids if tag == "tablePart" else pivot_ids).append(
                        relation_id
                    )
            elif _local_name(element.tag) == "c":
                current_cell = None
                element.clear()
            elif _local_name(element.tag) == "row":
                in_row = False
                current_cell = None
                element.clear()
            else:
                element.clear()
    except ET.ParseError as exc:
        raise WorkbookReadError(
            f"Invalid worksheet XML for {sheet_name!r}: {exc}"
        ) from exc
    finally:
        if stream is not None:
            stream.close()
    return table_ids, pivot_ids


def _office_document_part(archive: zipfile.ZipFile) -> str:
    for _rel_id, (target, rel_type) in _relationships(archive, "_rels/.rels").items():
        if rel_type.rstrip("/").endswith("/officeDocument"):
            return posixpath.normpath(target.lstrip("/"))
    if "xl/workbook.xml" in archive.namelist():
        return "xl/workbook.xml"
    raise WorkbookReadError("The package has no Office workbook relationship.")


def _validate_archive(archive: zipfile.ZipFile, config: AnalysisConfig) -> None:
    names: set[str] = set()
    total = 0
    for info in archive.infolist():
        normalized = posixpath.normpath(info.filename.replace("\\", "/"))
        if normalized.startswith("../") or normalized.startswith("/"):
            raise WorkbookReadError(f"Unsafe package member path: {info.filename!r}")
        if normalized in names:
            raise WorkbookReadError(f"Duplicate package member: {normalized!r}")
        names.add(normalized)
        if info.flag_bits & 0x1:
            raise WorkbookReadError("Encrypted XLSX package members are not supported.")
        total += info.file_size
        if total > config.max_uncompressed_bytes:
            raise WorkbookReadError(
                f"Workbook expands beyond the configured {config.max_uncompressed_bytes:,}-byte limit."
            )
        if normalized.lower().endswith(".xml") and info.file_size > config.max_xml_part_bytes:
            raise WorkbookReadError(
                f"XML part {normalized!r} exceeds the configured {config.max_xml_part_bytes:,}-byte limit."
            )


def _range_references(raw: str | None) -> Iterable[str]:
    if not raw:
        return []
    return [item for item in raw.replace(",", " ").split() if item]


def _register_range(
    acc: _SheetAccumulator,
    reference: str,
    kind: str,
    *,
    strong: bool = False,
    affects_extent: bool = True,
) -> None:
    try:
        bounds = parse_range_reference(reference)
    except ValueError:
        if strong:
            raise WorkbookReadError(f"Invalid {kind} range reference {reference!r}.")
        acc.warnings.append(f"Ignored invalid {kind} range reference {reference!r}.")
        return
    acc.ranges.append(_RangeFeature(kind, bounds, strong))
    if affects_extent and bounds.max_row < MAX_EXCEL_ROW:
        acc.active_max_row = max(acc.active_max_row, bounds.max_row)
    if affects_extent and bounds.max_col < MAX_EXCEL_COLUMN:
        acc.active_max_column = max(acc.active_max_column, bounds.max_col)


def _add_generated_output_range(
    acc: _SheetAccumulator,
    bounds: CellRange,
) -> None:
    if bounds not in acc.generated_output_range_set:
        acc.generated_output_range_set.add(bounds)
        acc.generated_output_ranges.append(bounds)


def _inside_generated_output(
    acc: _SheetAccumulator,
    row: int,
    col: int,
) -> bool:
    return (
        acc.generated_output_index is not None
        and acc.generated_output_index.contains(row, col)
    )


def _unsigned_attribute(
    element: ET.Element,
    name: str,
    *,
    default: int,
    maximum: int,
    context: str,
) -> int:
    raw = _attribute(element, name)
    if raw is None:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise WorkbookReadError(
            f"{context} has invalid {name}={raw!r}."
        ) from exc
    if not 0 <= value <= maximum:
        raise WorkbookReadError(
            f"{context} has out-of-range {name}={raw!r}."
        )
    return value


def _property_token(
    element: ET.Element,
    ignored: set[str],
    defaults: dict[str, str] | None = None,
    *,
    style_attribute: str | None = None,
    styles: _StyleCatalog | None = None,
    warnings: list[str] | None = None,
) -> str | None:
    defaults = defaults or {}
    attributes = [
        [_local_name(key), value]
        for key, value in element.attrib.items()
        if _local_name(key) not in ignored
        and value != defaults.get(_local_name(key))
    ]
    if style_attribute and styles:
        for item in attributes:
            if item[0] != style_attribute:
                continue
            try:
                item[1] = f"semantic:{styles.get(int(item[1]), warnings or [])}"
            except ValueError as exc:
                raise WorkbookReadError(
                    f"Invalid row/column style index {item[1]!r}."
                ) from exc
    attributes.sort()
    if not attributes:
        return None
    return _hash_payload(attributes)


def _canonical_number(raw_value: str) -> tuple[str, str] | None:
    try:
        value = Decimal(raw_value.strip())
    except (InvalidOperation, ValueError):
        return None
    if not value.is_finite():
        return None
    if value == 0:
        return "0", "NUMBER:0E0"
    sign, raw_digits, exponent = value.as_tuple()
    digits = list(raw_digits)
    while digits and digits[-1] == 0:
        digits.pop()
        exponent += 1
    coefficient = "".join(str(digit) for digit in digits)
    canonical = f"{'-' if sign else ''}{coefficient}E{exponent}"
    return str(value), f"NUMBER:{canonical}"


def _semantic_cell(
    *,
    row: int,
    col: int,
    cell_type: str,
    has_formula: bool,
    cache_present: bool,
    raw_value: str | None,
    inline_value: str | None,
    shared_strings: list[str],
) -> _SemanticCell | None:
    """Resolve the stored scalar without evaluating formulas."""

    kind_prefix = "FORMULA_" if has_formula else ""
    confidence = "MEDIUM" if has_formula else "HIGH"
    if has_formula and not cache_present:
        return _SemanticCell(
            row=row,
            col=col,
            display_value="<formula result unavailable>",
            comparison_key=None,
            value_kind="FORMULA_UNRESOLVED",
            confidence="LOW",
        )

    if cell_type == "s":
        if raw_value is None:
            return None
        try:
            value = shared_strings[int(raw_value)]
        except (ValueError, IndexError):
            raise WorkbookReadError(
                f"Cell {index_to_column(col)}{row} has an invalid shared-string index."
            )
        display = _normalize_display_text(value)
        if not display:
            return None
        return _SemanticCell(
            row=row,
            col=col,
            display_value=display,
            comparison_key=f"TEXT:{display}",
            value_kind=f"{kind_prefix}TEXT",
            confidence=confidence,
            literal_text=not has_formula,
        )

    if cell_type in {"inlineStr", "str"}:
        value = inline_value if inline_value is not None else raw_value
        # A present empty string cache is a known blank result, not an
        # unavailable formula result.
        display = _normalize_display_text(value or "")
        if not display:
            return None
        return _SemanticCell(
            row=row,
            col=col,
            display_value=display,
            comparison_key=f"TEXT:{display}",
            value_kind=f"{kind_prefix}TEXT",
            confidence=confidence,
            literal_text=not has_formula,
        )

    if cell_type == "b":
        normalized = (raw_value or "").strip().casefold()
        if normalized in {"1", "true"}:
            display, key = "TRUE", "BOOLEAN:true"
        elif normalized in {"0", "false"}:
            display, key = "FALSE", "BOOLEAN:false"
        else:
            return _SemanticCell(
                row=row,
                col=col,
                display_value=raw_value or "<invalid boolean>",
                comparison_key=None,
                value_kind=f"{kind_prefix}BOOLEAN_INVALID",
                confidence="LOW",
            )
        return _SemanticCell(
            row=row,
            col=col,
            display_value=display,
            comparison_key=key,
            value_kind=f"{kind_prefix}BOOLEAN",
            confidence=confidence,
        )

    if cell_type == "e":
        return _SemanticCell(
            row=row,
            col=col,
            display_value=_normalize_display_text(raw_value or "<error>"),
            comparison_key=None,
            value_kind=f"{kind_prefix}ERROR",
            confidence=confidence,
        )

    if cell_type == "d":
        display = _normalize_display_text(raw_value or "")
        if not display:
            return None
        return _SemanticCell(
            row=row,
            col=col,
            display_value=display,
            comparison_key=f"DATE:{display}",
            value_kind=f"{kind_prefix}DATE",
            confidence=confidence,
        )

    if raw_value is None:
        if has_formula:
            return _SemanticCell(
                row=row,
                col=col,
                display_value="<formula result unavailable>",
                comparison_key=None,
                value_kind="FORMULA_UNRESOLVED",
                confidence="LOW",
            )
        return None
    number = _canonical_number(raw_value)
    if number is None:
        return _SemanticCell(
            row=row,
            col=col,
            display_value=raw_value,
            comparison_key=None,
            value_kind=f"{kind_prefix}NUMBER_INVALID",
            confidence="LOW",
        )
    display, key = number
    return _SemanticCell(
        row=row,
        col=col,
        display_value=display,
        comparison_key=key,
        value_kind=f"{kind_prefix}NUMBER",
        confidence=confidence,
    )


def _append_capped(target: list[str], coordinate: str) -> None:
    if len(target) < _REF_ERROR_COORDINATE_LIMIT:
        target.append(coordinate)


def _record_ref_errors(
    acc: _SheetAccumulator,
    coordinate: str,
    *,
    cached_error: bool,
    formula_error: bool,
) -> None:
    if cached_error:
        acc.cached_ref_error_count += 1
        _append_capped(acc.cached_ref_error_coordinates, coordinate)
    if formula_error:
        acc.formula_ref_error_count += 1
        _append_capped(acc.formula_ref_error_coordinates, coordinate)
    if cached_error or formula_error:
        acc.ref_error_count += 1
        _append_capped(acc.ref_error_coordinates, coordinate)


def _capture_kpi_semantic_cell(
    acc: _SheetAccumulator,
    snapshot: CellSnapshot,
) -> None:
    if not acc.capture_kpi:
        return
    if len(acc.kpi_semantic_cells) >= acc.kpi_semantic_cell_limit:
        raise WorkbookReadError(
            "The KPI worksheet exceeds the configured "
            f"{acc.kpi_semantic_cell_limit:,}-semantic-cell limit."
        )
    acc.kpi_semantic_cells.append(snapshot)


def _capture_material_cell(acc: _SheetAccumulator, snapshot: CellSnapshot) -> int:
    if len(acc.material_cells) >= acc.material_cell_limit:
        raise WorkbookReadError(
            "Worksheet comparison evidence exceeds the configured "
            f"{acc.material_cell_limit:,}-material-cell limit."
        )
    acc.material_cells.append(snapshot)
    _capture_kpi_semantic_cell(acc, snapshot)
    return len(acc.material_cells) - 1


def _kpi_snapshot(acc: _SheetAccumulator, header_scan_rows: int) -> KpiColumnSnapshot:
    if not acc.capture_kpi:
        return KpiColumnSnapshot(status="NOT_APPLICABLE")

    candidates = sorted(
        (
            item
            for item in acc.kpi_semantic_cells
            if item.row <= header_scan_rows
            and not item.has_formula
            and item.value_kind == "TEXT"
            and item.display_value is not None
            and _normalize_text(item.display_value) == "kpi"
        ),
        key=lambda item: (item.row, item.col),
    )
    candidate_coordinates = [
        item.coordinate for item in candidates
    ]
    scan_note = f"Literal KPI headers were searched in the first {header_scan_rows:,} rows."
    if not candidates:
        return KpiColumnSnapshot(
            status="MISSING",
            header_candidates=[],
            notes=[scan_note, "No literal text cell normalized exactly to 'KPI'."],
        )
    if len(candidates) > 1:
        return KpiColumnSnapshot(
            status="AMBIGUOUS",
            header_candidates=candidate_coordinates,
            notes=[
                scan_note,
                f"{len(candidates)} possible KPI headers were found; no header was guessed.",
            ],
        )

    header = candidates[0]
    header_coordinate = candidate_coordinates[0]
    body = sorted(
        (
            item
            for item in acc.kpi_semantic_cells
            if item.col == header.col and item.row > header.row
        ),
        key=lambda item: item.row,
    )
    entries = [
        KpiEntry(
            display_value=item.display_value,
            comparison_key=item.comparison_key,
            coordinate=item.coordinate,
            row=item.row,
            value_kind=item.value_kind,
            confidence=item.confidence,
        )
        for item in body
        if item.display_value is not None
    ]
    coordinates_by_key: dict[str, list[str]] = defaultdict(list)
    for entry in entries:
        if entry.comparison_key is not None:
            coordinates_by_key[entry.comparison_key].append(entry.coordinate)
    duplicates = {
        key: coordinates
        for key, coordinates in coordinates_by_key.items()
        if len(coordinates) > 1
    }
    notes = [scan_note]
    if body:
        notes.append(
            f"KPI entries were read below {header_coordinate} through row {body[-1].row:,}."
        )
    else:
        notes.append(f"No nonblank semantic cells were found below {header_coordinate}.")
    formula_cache_count = sum(
        item.value_kind.startswith("FORMULA_") and item.confidence == "MEDIUM"
        for item in entries
    )
    unresolved_count = sum(item.comparison_key is None for item in entries)
    if formula_cache_count:
        notes.append(
            f"{formula_cache_count:,} KPI cell(s) use stored formula results; these are not recalculated."
        )
    if unresolved_count:
        notes.append(
            f"{unresolved_count:,} KPI cell(s) could not produce a comparable identifier."
        )
    return KpiColumnSnapshot(
        status="FOUND",
        header_coordinate=header_coordinate,
        header_candidates=candidate_coordinates,
        entries=entries,
        duplicate_keys=duplicates,
        notes=notes,
    )


def _parse_cell(
    cell: ET.Element,
    row: int,
    col: int,
    styles: _StyleCatalog,
    shared_strings: list[str],
    shared_formulas: dict[str, tuple[str, str]],
    acc: _SheetAccumulator,
) -> None:
    acc.stored_cell_count += 1
    cell_type = _attribute(cell, "t") or "n"
    style_index_raw = _attribute(cell, "s") or "0"
    try:
        style_index = int(style_index_raw)
    except ValueError:
        style_index = -1
    style = styles.get(style_index, acc.warnings)

    formula_element = next((node for node in cell if _local_name(node.tag) == "f"), None)
    value_element = next((node for node in cell if _local_name(node.tag) == "v"), None)
    inline_element = next((node for node in cell if _local_name(node.tag) == "is"), None)
    formula = formula_element.text if formula_element is not None else None
    raw_value = value_element.text if value_element is not None else None
    inline_value = None
    if inline_element is not None:
        inline_value = _rich_text(inline_element)

    has_formula = formula_element is not None
    coordinate = f"{index_to_column(col)}{row}"
    cached_ref_error = cell_type == "e" and _normalize_text(raw_value or "") == "#ref!"
    formula_ref_error = formula is not None and _formula_contains_ref_error(formula)
    _record_ref_errors(
        acc,
        coordinate,
        cached_error=cached_ref_error,
        formula_error=formula_ref_error,
    )
    cache_present = value_element is not None or inline_element is not None
    semantic = _semantic_cell(
        row=row,
        col=col,
        cell_type=cell_type,
        has_formula=has_formula,
        cache_present=cache_present,
        raw_value=raw_value,
        inline_value=inline_value,
        shared_strings=shared_strings,
    )
    generated_output = _inside_generated_output(acc, row, col)

    has_content = has_formula or raw_value is not None or inline_value is not None
    # A default-style empty <c> is serialization noise, not a materialized
    # row/column.  Styled blanks, on the other hand, are template evidence.
    if not has_content and style_index == 0:
        return
    acc.cell_count += int(has_content)
    acc.formula_count += int(has_formula)
    acc.styled_blank_count += int(not has_content and style_index != 0)
    if not generated_output:
        _mark_active(acc, row, col)
        if has_content:
            _mark_content(acc, row, col)

    style_token = f"S:{style}"
    if not generated_output:
        _add_weight(acc.row_weights, row, style_token, 1.0)
        _add_weight(acc.column_weights, col, style_token, 1.0)

    kind = (
        "FORMULA"
        if has_formula
        else "TEXT"
        if cell_type in {"s", "inlineStr", "str"}
        else "VALUE"
        if has_content
        else "BLANK_STYLED"
    )
    if not generated_output:
        _add_weight(acc.row_weights, row, f"K:{kind}", 0.35)
        _add_weight(acc.column_weights, col, f"K:{kind}", 0.35)

    if has_formula:
        formula_type = (_attribute(formula_element, "t") or "normal").strip()
        is_shared_formula = formula_type.casefold() == "shared"
        shared_index = _attribute(formula_element, "si")
        formula_range = _attribute(formula_element, "ref")
        explicit_formula = formula if formula is not None and formula.strip() else None
        comparison_shape: str | None = None
        comparison_relative: str | None = None
        formula_status = "UNRESOLVED"
        if explicit_formula is not None:
            try:
                comparison_shape, comparison_relative = _normalize_formula(
                    explicit_formula, row, col
                )
                formula_status = "RESOLVED"
            except ValueError:
                # Invalid A1 coordinates must not become a guessed comparable
                # representation. The raw formula remains available as evidence.
                comparison_shape = None
                comparison_relative = None

        material = CellSnapshot(
            coordinate=coordinate,
            row=row,
            col=col,
            has_formula=True,
            display_value=semantic.display_value if semantic is not None else None,
            comparison_key=semantic.comparison_key if semantic is not None else None,
            value_kind=(
                semantic.value_kind
                if semantic is not None
                else "FORMULA_BLANK" if cache_present else "FORMULA_UNRESOLVED"
            ),
            confidence=(
                semantic.confidence
                if semantic is not None
                else "MEDIUM" if cache_present else "LOW"
            ),
            formula_status=formula_status,
            formula_type=formula_type,
            formula_text=explicit_formula,
            formula_shape=comparison_shape,
            formula_relative=comparison_relative,
            formula_group=shared_index,
            formula_range=formula_range,
        )
        material_index = _capture_material_cell(acc, material)
        if is_shared_formula and shared_index:
            group = acc.shared_formula_groups.setdefault(
                shared_index, _SharedFormulaGroup()
            )
            group.cell_indices.append(material_index)
            if explicit_formula is not None:
                group.master_indices.append(material_index)

        # Structural signatures retain their established behavior. Formula
        # identity for semantic comparison is kept separately above and never
        # uses a cached result.
        cached_normalization = shared_formulas.get(shared_index or "")
        if is_shared_formula and formula is None and cached_normalization is not None:
            shape, relative = cached_normalization
        else:
            formula_text = formula or f"{formula_type.upper()}:{shared_index or '?'}"
            try:
                shape, relative = _normalize_formula(formula_text, row, col)
            except ValueError:
                shape = re.sub(r"\s+", "", formula_text).upper()
                relative = shape
            if is_shared_formula and shared_index and formula is not None:
                shared_formulas[shared_index] = (shape, relative)
        if generated_output:
            # Pivot/query/table-calculation contents are regenerated by Excel.
            # Formula evidence remains available for direct integrity checks,
            # while the expression/cache cannot become a structural anchor.
            return
        shape_token = f"F:{_hash_payload(shape)}"
        relative_token = f"R:{_hash_payload(relative)}"
        for target, axis in ((acc.row_weights, row), (acc.column_weights, col)):
            _add_weight(target, axis, shape_token, 1.0)
            _add_weight(target, axis, relative_token, 2.75)
        acc.row_strong[row].add(relative_token)
        acc.column_strong[col].add(relative_token)
        acc.cell_anchors[relative_token].append((row, col))
        return

    if semantic is not None:
        snapshot = CellSnapshot(
            coordinate=coordinate,
            row=row,
            col=col,
            has_formula=False,
            display_value=semantic.display_value,
            comparison_key=semantic.comparison_key,
            value_kind=semantic.value_kind,
            confidence=semantic.confidence,
        )
        if generated_output:
            # Generated literals are not generic prefills or structure, but a
            # KPI sheet still needs its dedicated semantic stream so an
            # unchanged query-backed KPI table does not look headerless.
            _capture_kpi_semantic_cell(acc, snapshot)
        else:
            _capture_material_cell(acc, snapshot)

    if generated_output:
        # The style, occupied coordinate, and explicit object range remain
        # structural evidence. The generated scalar itself is not protected
        # template content and does not identify a logical row or column.
        return

    text_value: str | None = None
    if cell_type == "s" and raw_value is not None:
        try:
            text_value = shared_strings[int(raw_value)]
        except (ValueError, IndexError):
            raise WorkbookReadError(
                f"Cell {index_to_column(col)}{row} has an invalid shared-string index."
            )
    elif cell_type in {"inlineStr", "str"}:
        text_value = inline_value if inline_value is not None else raw_value

    if text_value is not None:
        normalized = _normalize_text(text_value)
        if normalized:
            token = f"T:{_hash_payload(normalized)}"
            _add_weight(acc.row_weights, row, token, 5.0)
            _add_weight(acc.column_weights, col, token, 5.0)
            acc.row_strong[row].add(token)
            acc.column_strong[col].add(token)
            acc.cell_anchors[token].append((row, col))
    elif raw_value is not None:
        token = f"V:{_hash_payload(raw_value)}"
        _add_weight(acc.row_weights, row, token, 0.45)
        _add_weight(acc.column_weights, col, token, 0.45)


def _resolve_shared_formula_groups(acc: _SheetAccumulator) -> None:
    """Resolve shared followers only from a unique, bounded OOXML master."""

    missing_group = sum(
        item.has_formula
        and (item.formula_type or "").casefold() == "shared"
        and not item.formula_group
        and item.formula_text is None
        for item in acc.material_cells
    )
    if missing_group:
        acc.warnings.append(
            f"{missing_group:,} shared-formula follower(s) have no shared group id and remain unresolved."
        )

    for group_id, group in sorted(acc.shared_formula_groups.items()):
        master_indices = set(group.master_indices)
        followers = [
            index for index in group.cell_indices if index not in master_indices
        ]
        if not followers:
            continue
        if len(group.master_indices) != 1:
            acc.warnings.append(
                f"Shared formula group {group_id!r} has {len(group.master_indices)} explicit "
                f"master cells; {len(followers):,} follower(s) remain unresolved."
            )
            continue

        master = acc.material_cells[group.master_indices[0]]
        if (
            master.formula_status != "RESOLVED"
            or master.formula_shape is None
            or master.formula_relative is None
            or not master.formula_range
        ):
            acc.warnings.append(
                f"Shared formula group {group_id!r} has no fully resolved bounded master; "
                f"{len(followers):,} follower(s) remain unresolved."
            )
            continue
        try:
            bounds = parse_range_reference(master.formula_range)
        except ValueError:
            acc.warnings.append(
                f"Shared formula group {group_id!r} has invalid master range "
                f"{master.formula_range!r}; {len(followers):,} follower(s) remain unresolved."
            )
            continue
        if not (
            bounds.min_row <= master.row <= bounds.max_row
            and bounds.min_col <= master.col <= bounds.max_col
        ):
            acc.warnings.append(
                f"Shared formula group {group_id!r} master {master.coordinate} lies outside "
                f"its declared range {master.formula_range!r}; followers remain unresolved."
            )
            continue

        unresolved = 0
        for index in followers:
            follower = acc.material_cells[index]
            if not (
                bounds.min_row <= follower.row <= bounds.max_row
                and bounds.min_col <= follower.col <= bounds.max_col
            ):
                unresolved += 1
                continue
            # OOXML defines a shared follower as the master's formula translated
            # to this cell. Its coordinate-relative representation is therefore
            # the same as the master's. We intentionally do not fabricate an A1
            # formula string; the validated normalized pattern is sufficient for
            # identity comparison and the explicit text remains None.
            follower.formula_shape = master.formula_shape
            follower.formula_relative = master.formula_relative
            follower.formula_status = "RESOLVED_SHARED"
            follower.formula_range = master.formula_range
        if unresolved:
            acc.warnings.append(
                f"Shared formula group {group_id!r} has {unresolved:,} follower(s) outside "
                f"declared range {master.formula_range!r}; those followers remain unresolved."
            )


def _related_part(
    archive: zipfile.ZipFile,
    source_part: str,
    relation_id: str,
    relationships: dict[str, tuple[str, str]],
    *,
    expected_type: str,
    label: str,
) -> str:
    relation = relationships.get(relation_id)
    if not relation:
        raise WorkbookReadError(f"Missing {label} relationship {relation_id!r}.")
    relation_type = relation[1].rstrip("/").casefold()
    if not relation_type.endswith(f"/{expected_type.casefold()}"):
        raise WorkbookReadError(
            f"{label.capitalize()} relationship {relation_id!r} has unexpected type "
            f"{relation[1]!r}."
        )
    part = _resolve_part(source_part, relation[0])
    if part not in archive.namelist():
        raise WorkbookReadError(f"Missing {label} part {part!r}.")
    return part


def _required_part_range(
    root: ET.Element,
    *,
    part: str,
    label: str,
) -> tuple[str, CellRange]:
    reference = _attribute(root, "ref")
    if not reference:
        raise WorkbookReadError(f"{label.capitalize()} part {part!r} has no range reference.")
    if "!" in reference:
        raise WorkbookReadError(
            f"{label.capitalize()} part {part!r} has non-local range {reference!r}."
        )
    try:
        return reference, parse_range_reference(reference)
    except ValueError as exc:
        raise WorkbookReadError(
            f"{label.capitalize()} part {part!r} has invalid range {reference!r}."
        ) from exc


def _table_has_query_source(
    archive: zipfile.ZipFile,
    table_part: str,
    root: ET.Element,
) -> bool:
    relationships = _relationships(archive, _rels_part(table_part))
    query_relations = [
        (relation_id, relation)
        for relation_id, relation in relationships.items()
        if relation[1].rstrip("/").casefold().endswith("/querytable")
    ]
    for relation_id, _relation in query_relations:
        query_part = _related_part(
            archive,
            table_part,
            relation_id,
            relationships,
            expected_type="queryTable",
            label="query-table",
        )
        try:
            query_root = _xml_root(archive, query_part)
        except ET.ParseError as exc:
            raise WorkbookReadError(
                f"Invalid query-table part {query_part!r}: {exc}."
            ) from exc
        if _local_name(query_root.tag) != "queryTable":
            raise WorkbookReadError(
                f"Query-table part {query_part!r} has unexpected root "
                f"{_local_name(query_root.tag)!r}."
            )
    table_type = (_attribute(root, "tableType") or "").strip().casefold()
    return bool(query_relations) or table_type == "querytable"


def _register_table_generated_outputs(
    acc: _SheetAccumulator,
    root: ET.Element,
    bounds: CellRange,
    *,
    part: str,
    query_backed: bool,
) -> None:
    if query_backed:
        _add_generated_output_range(acc, bounds)
        return

    width = bounds.max_col - bounds.min_col + 1
    height = bounds.max_row - bounds.min_row + 1
    header_rows = _unsigned_attribute(
        root,
        "headerRowCount",
        default=1,
        maximum=height,
        context=f"Table part {part!r}",
    )
    totals_rows = _unsigned_attribute(
        root,
        "totalsRowCount",
        default=0,
        maximum=height,
        context=f"Table part {part!r}",
    )
    if header_rows + totals_rows > height:
        raise WorkbookReadError(
            f"Table part {part!r} has header/totals rows outside its range."
        )

    columns_parent = next(
        (child for child in root if _local_name(child.tag) == "tableColumns"),
        None,
    )
    if columns_parent is None:
        return
    columns = [
        child
        for child in columns_parent
        if _local_name(child.tag) == "tableColumn"
    ]
    if len(columns) != width:
        raise WorkbookReadError(
            f"Table part {part!r} declares {len(columns):,} column definition(s) "
            f"for a {width:,}-column range."
        )

    body_min_row = bounds.min_row + header_rows
    body_max_row = bounds.max_row - totals_rows
    totals_min_row = bounds.max_row - totals_rows + 1
    for offset, column in enumerate(columns):
        col = bounds.min_col + offset
        children = {_local_name(child.tag) for child in column}
        if (
            "calculatedColumnFormula" in children
            and body_min_row <= body_max_row
        ):
            _add_generated_output_range(
                acc,
                CellRange(body_min_row, col, body_max_row, col),
            )
        totals_function = (
            _attribute(column, "totalsRowFunction") or "none"
        ).strip().casefold()
        if totals_rows and (
            "totalsRowFormula" in children or totals_function != "none"
        ):
            _add_generated_output_range(
                acc,
                CellRange(totals_min_row, col, bounds.max_row, col),
            )


def _read_table_ranges(
    archive: zipfile.ZipFile,
    sheet_part: str,
    relation_ids: list[str],
    relationships: dict[str, tuple[str, str]],
    acc: _SheetAccumulator,
) -> None:
    for relation_id in relation_ids:
        part = _related_part(
            archive,
            sheet_part,
            relation_id,
            relationships,
            expected_type="table",
            label="table",
        )
        try:
            root = _xml_root(archive, part)
        except ET.ParseError as exc:
            raise WorkbookReadError(f"Invalid table part {part!r}: {exc}.") from exc
        if _local_name(root.tag) != "table":
            raise WorkbookReadError(
                f"Table part {part!r} has unexpected root {_local_name(root.tag)!r}."
            )
        reference, bounds = _required_part_range(
            root,
            part=part,
            label="table",
        )
        query_backed = _table_has_query_source(archive, part, root)
        # A query refresh can resize the table without a country inserting
        # worksheet rows or columns. Its bounds and generated cells therefore
        # must not become structural alignment anchors.
        if not query_backed:
            _register_range(acc, reference, "TABLE", strong=True)
        acc.table_count += 1
        _register_table_generated_outputs(
            acc,
            root,
            bounds,
            part=part,
            query_backed=query_backed,
        )


def _read_pivot_table_ranges(
    archive: zipfile.ZipFile,
    sheet_part: str,
    relation_ids: list[str],
    relationships: dict[str, tuple[str, str]],
    acc: _SheetAccumulator,
) -> None:
    for relation_id in relation_ids:
        part = _related_part(
            archive,
            sheet_part,
            relation_id,
            relationships,
            expected_type="pivotTable",
            label="PivotTable",
        )
        try:
            root = _xml_root(archive, part)
        except ET.ParseError as exc:
            raise WorkbookReadError(
                f"Invalid PivotTable part {part!r}: {exc}."
            ) from exc
        if _local_name(root.tag) != "pivotTableDefinition":
            raise WorkbookReadError(
                f"PivotTable part {part!r} has unexpected root "
                f"{_local_name(root.tag)!r}."
            )
        locations = [
            child for child in root if _local_name(child.tag) == "location"
        ]
        if len(locations) != 1:
            raise WorkbookReadError(
                f"PivotTable part {part!r} has {len(locations):,} location element(s)."
            )
        reference, bounds = _required_part_range(
            locations[0],
            part=part,
            label="PivotTable location",
        )
        # PivotTable locations are regenerated and may expand or contract on
        # refresh. Keep the range only as a generated-output mask, not as proof
        # of a manual worksheet-axis operation.
        _add_generated_output_range(acc, bounds)


def _read_drawing_ranges(
    archive: zipfile.ZipFile,
    sheet_part: str,
    relation_ids: list[str],
    relationships: dict[str, tuple[str, str]],
    acc: _SheetAccumulator,
) -> None:
    for relation_id in relation_ids:
        relation = relationships.get(relation_id)
        if not relation:
            continue
        part = _resolve_part(sheet_part, relation[0])
        if part not in archive.namelist():
            continue
        try:
            root = _xml_root(archive, part)
        except ET.ParseError:
            acc.warnings.append(f"Invalid drawing part {part!r}; its anchors were ignored.")
            continue
        for anchor in root:
            if _local_name(anchor.tag) not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
                continue
            points: list[tuple[int, int]] = []
            for marker in anchor:
                if _local_name(marker.tag) not in {"from", "to"}:
                    continue
                row_text = next((n.text for n in marker if _local_name(n.tag) == "row"), None)
                col_text = next((n.text for n in marker if _local_name(n.tag) == "col"), None)
                if row_text is not None and col_text is not None:
                    try:
                        points.append((int(row_text) + 1, int(col_text) + 1))
                    except ValueError:
                        pass
            if points:
                rows = [point[0] for point in points]
                cols = [point[1] for point in points]
                _register_range(
                    acc,
                    f"{index_to_column(min(cols))}{min(rows)}:{index_to_column(max(cols))}{max(rows)}",
                    "DRAWING",
                    affects_extent=False,
                )


def _apply_range_features(acc: _SheetAccumulator) -> None:
    for feature in acc.ranges:
        bounds = feature.bounds
        weight = 2.5 if feature.strong else 0.8
        row_end = min(bounds.max_row, acc.active_max_row)
        col_end = min(bounds.max_col, acc.active_max_column)

        if bounds.min_row <= row_end:
            rows: set[int] = {bounds.min_row, row_end}
            if row_end - bounds.min_row <= 5_000:
                rows.update(range(bounds.min_row, row_end + 1))
            for row in rows:
                role = "START" if row == bounds.min_row else "END" if row == row_end else "INSIDE"
                token = f"RG:{feature.kind}:ROW:{role}"
                _add_weight(acc.row_weights, row, token, weight)
                if feature.strong and role != "INSIDE":
                    acc.row_strong[row].add(token)

        if bounds.min_col <= col_end:
            columns: set[int] = {bounds.min_col, col_end}
            if col_end - bounds.min_col <= 2_000:
                columns.update(range(bounds.min_col, col_end + 1))
            for col in columns:
                role = "START" if col == bounds.min_col else "END" if col == col_end else "INSIDE"
                token = f"RG:{feature.kind}:COL:{role}"
                _add_weight(acc.column_weights, col, token, weight)
                if feature.strong and role != "INSIDE":
                    acc.column_strong[col].add(token)

    for start, end, token_hash in acc.column_properties:
        bounded_end = min(end, acc.active_max_column)
        for col in range(start, bounded_end + 1):
            token = f"CP:{token_hash}"
            _add_weight(acc.column_weights, col, token, 2.0)


def _axis_signatures(
    maximum: int,
    weights: dict[int, dict[str, float]],
    strong: dict[int, set[str]],
) -> list[AxisSignature]:
    signatures: list[AxisSignature] = []
    for index in range(1, maximum + 1):
        item_weights = weights.get(index)
        if not item_weights:
            item_weights = {"__BLANK__": 1.0}
            information = 0.0
        else:
            information = sum(item_weights.values())
        digest = _hash_payload(sorted(item_weights.items()), length=32)
        signatures.append(
            AxisSignature(
                index=index,
                weights=dict(item_weights),
                strong_tokens=frozenset(strong.get(index, set())),
                digest=digest,
                information=information,
            )
        )
    return signatures


def _format_bbox(min_row: int, min_col: int, max_row: int, max_col: int) -> str | None:
    if max_row <= 0 or max_col <= 0:
        return None
    return f"{index_to_column(min_col)}{min_row}:{index_to_column(max_col)}{max_row}"


def _parse_worksheet(
    archive: zipfile.ZipFile,
    part_name: str,
    name: str,
    index: int,
    state: str,
    styles: _StyleCatalog,
    shared_strings: list[str],
    config: AnalysisConfig,
) -> SheetStructure:
    if part_name not in archive.namelist():
        raise WorkbookReadError(f"Worksheet {name!r} points to missing part {part_name!r}.")

    acc = _SheetAccumulator(
        capture_kpi=_normalize_text(name) == "kpi",
        kpi_semantic_cell_limit=config.max_kpi_semantic_cells,
        material_cell_limit=config.max_comparison_cells_per_sheet,
    )
    current_row = 0
    in_row = False
    next_row = 1
    next_column = 1
    drawing_relation_ids: list[str] = []
    shared_formulas: dict[str, tuple[str, str]] = {}
    seen_rows: set[int] = set()
    seen_cells: set[tuple[int, int]] = set()
    relationships = _relationships(archive, _rels_part(part_name))
    table_relation_ids, pivot_relation_ids = _worksheet_object_relationship_ids(
        archive,
        part_name,
        name,
        acc,
    )
    _read_table_ranges(
        archive,
        part_name,
        table_relation_ids,
        relationships,
        acc,
    )
    _read_pivot_table_ranges(
        archive,
        part_name,
        pivot_relation_ids,
        relationships,
        acc,
    )
    acc.generated_output_ranges.sort(
        key=lambda bounds: (
            bounds.min_row,
            bounds.min_col,
            bounds.max_row,
            bounds.max_col,
        )
    )
    acc.generated_output_index = _GeneratedOutputIndex.build(
        acc.generated_output_ranges
    )

    stream: BinaryIO | None = None
    try:
        stream = archive.open(part_name)
        for event, element in ET.iterparse(stream, events=("start", "end")):
            tag = _local_name(element.tag)
            if event == "start":
                if tag == "dimension":
                    acc.declared_dimension = _attribute(element, "ref")
                elif tag == "row":
                    raw_row = _attribute(element, "r")
                    try:
                        current_row = int(raw_row) if raw_row else next_row
                    except ValueError:
                        raise WorkbookReadError(f"Worksheet {name!r} contains an invalid row index.")
                    if not 1 <= current_row <= MAX_EXCEL_ROW:
                        raise WorkbookReadError(
                            f"Worksheet {name!r} contains out-of-range row {current_row}."
                        )
                    if current_row in seen_rows:
                        raise WorkbookReadError(
                            f"Worksheet {name!r} contains duplicate row {current_row}."
                        )
                    seen_rows.add(current_row)
                    in_row = True
                    next_row = current_row + 1
                    next_column = 1
                    token = _property_token(
                        element,
                        {"r", "spans"},
                        {
                            "hidden": "0",
                            "customFormat": "0",
                            "customHeight": "0",
                            "outlineLevel": "0",
                            "collapsed": "0",
                            "thickTop": "0",
                            "thickBot": "0",
                            "s": "0",
                        },
                        style_attribute="s",
                        styles=styles,
                        warnings=acc.warnings,
                    )
                    if token:
                        _add_weight(acc.row_weights, current_row, f"RP:{token}", 2.0)
                        acc.active_max_row = max(acc.active_max_row, current_row)
                elif tag == "col":
                    try:
                        start = int(_attribute(element, "min") or "0")
                        end = int(_attribute(element, "max") or "0")
                    except ValueError as exc:
                        raise WorkbookReadError(f"Worksheet {name!r} contains invalid column metadata.") from exc
                    token = _property_token(
                        element,
                        {"min", "max"},
                        {
                            "hidden": "0",
                            "outlineLevel": "0",
                            "collapsed": "0",
                            "bestFit": "0",
                            "customWidth": "0",
                            "style": "0",
                            "phonetic": "0",
                        },
                        style_attribute="style",
                        styles=styles,
                        warnings=acc.warnings,
                    )
                    if token and 1 <= start <= end <= MAX_EXCEL_COLUMN:
                        acc.column_properties.append((start, end, token))
                        if end < MAX_EXCEL_COLUMN:
                            acc.active_max_column = max(acc.active_max_column, end)
                        else:
                            acc.warnings.append(
                                "A column property spans through XFD; it is retained within the "
                                "active area but does not define the active-column count."
                            )
                elif tag == "mergeCell":
                    reference = _attribute(element, "ref")
                    if reference:
                        _register_range(acc, reference, "MERGE", strong=True)
                        acc.merged_count += 1
                elif tag == "autoFilter":
                    for reference in _range_references(_attribute(element, "ref")):
                        _register_range(acc, reference, "AUTOFILTER", strong=True)
                elif tag == "conditionalFormatting":
                    for reference in _range_references(_attribute(element, "sqref")):
                        _register_range(acc, reference, "CONDITIONAL_FORMAT")
                elif tag == "dataValidation":
                    for reference in _range_references(_attribute(element, "sqref")):
                        _register_range(acc, reference, "DATA_VALIDATION")
                elif tag == "hyperlink":
                    for reference in _range_references(_attribute(element, "ref")):
                        _register_range(acc, reference, "HYPERLINK")
                elif tag == "pane":
                    reference = _attribute(element, "topLeftCell")
                    if reference:
                        _register_range(acc, reference, "FREEZE_PANE", affects_extent=False)
                elif tag == "drawing":
                    relation_id = _attribute(element, "id")
                    if relation_id:
                        drawing_relation_ids.append(relation_id)
                continue

            if tag == "c":
                if not in_row:
                    raise WorkbookReadError(
                        f"Worksheet {name!r} contains a cell outside a row element."
                    )
                reference = _attribute(element, "r")
                if reference:
                    try:
                        row, col = parse_cell_reference(reference)
                    except ValueError as exc:
                        raise WorkbookReadError(
                            f"Worksheet {name!r} contains invalid cell reference {reference!r}."
                        ) from exc
                else:
                    row, col = current_row, next_column
                if row != current_row:
                    raise WorkbookReadError(
                        f"Cell {reference!r} is inconsistent with enclosing row {current_row} "
                        f"in worksheet {name!r}."
                    )
                if (row, col) in seen_cells:
                    raise WorkbookReadError(
                        f"Worksheet {name!r} contains duplicate cell {index_to_column(col)}{row}."
                    )
                seen_cells.add((row, col))
                next_column = col + 1
                _parse_cell(
                    element,
                    row,
                    col,
                    styles,
                    shared_strings,
                    shared_formulas,
                    acc,
                )
                if acc.stored_cell_count > config.max_cells_per_sheet:
                    raise WorkbookReadError(
                        f"Worksheet {name!r} exceeds the configured {config.max_cells_per_sheet:,}-cell limit."
                    )
                element.clear()
            elif tag == "row":
                in_row = False
                element.clear()
    except ET.ParseError as exc:
        raise WorkbookReadError(f"Invalid worksheet XML for {name!r}: {exc}") from exc
    finally:
        if stream is not None:
            stream.close()

    _resolve_shared_formula_groups(acc)
    acc.material_cells.sort(key=lambda item: (item.row, item.col))
    _read_drawing_ranges(archive, part_name, drawing_relation_ids, relationships, acc)

    if acc.active_max_row > config.max_active_rows:
        raise WorkbookReadError(
            f"Worksheet {name!r} has active evidence at row {acc.active_max_row:,}, beyond the "
            f"configured {config.max_active_rows:,}-row analysis limit. Increase --max-active-rows "
            "only after checking for stray far-away formatting."
        )
    if acc.active_max_column > config.max_active_columns:
        raise WorkbookReadError(
            f"Worksheet {name!r} has active evidence at column {acc.active_max_column:,}, beyond "
            f"the configured {config.max_active_columns:,}-column analysis limit."
        )

    _apply_range_features(acc)
    rows = _axis_signatures(acc.active_max_row, acc.row_weights, acc.row_strong)
    columns = _axis_signatures(acc.active_max_column, acc.column_weights, acc.column_strong)
    content_ref = _format_bbox(
        acc.content_min_row,
        acc.content_min_column,
        acc.content_max_row,
        acc.content_max_column,
    )
    metrics = SheetMetrics(
        active_rows=acc.active_max_row,
        active_columns=acc.active_max_column,
        content_rows=len(acc.content_rows),
        content_columns=len(acc.content_columns),
        populated_cells=acc.cell_count,
        formula_cells=acc.formula_count,
        styled_blank_cells=acc.styled_blank_count,
        merged_ranges=acc.merged_count,
        table_ranges=acc.table_count,
        ref_error_count=acc.ref_error_count,
        cached_ref_error_count=acc.cached_ref_error_count,
        formula_ref_error_count=acc.formula_ref_error_count,
        active_ref=format_active_ref(acc.active_max_row, acc.active_max_column),
        content_ref=content_ref,
        declared_dimension=acc.declared_dimension,
    )
    for label, count, coordinates in (
        ("#REF!", acc.ref_error_count, acc.ref_error_coordinates),
        ("cached #REF!", acc.cached_ref_error_count, acc.cached_ref_error_coordinates),
        ("formula #REF!", acc.formula_ref_error_count, acc.formula_ref_error_coordinates),
    ):
        if count > len(coordinates):
            acc.warnings.append(
                f"Only the first {_REF_ERROR_COORDINATE_LIMIT:,} {label} cell coordinates "
                f"were retained out of {count:,}."
            )
    return SheetStructure(
        name=name,
        index=index,
        state=state,
        sheet_type="worksheet",
        part_name=part_name,
        metrics=metrics,
        rows=rows,
        columns=columns,
        cell_anchors=dict(acc.cell_anchors),
        material_cells=acc.material_cells,
        generated_output_ranges=list(acc.generated_output_ranges),
        kpi_snapshot=_kpi_snapshot(acc, config.kpi_header_scan_rows),
        ref_error_coordinates=list(acc.ref_error_coordinates),
        cached_ref_error_coordinates=list(acc.cached_ref_error_coordinates),
        formula_ref_error_coordinates=list(acc.formula_ref_error_coordinates),
        warnings=acc.warnings,
    )


def _empty_nonworksheet(
    name: str,
    index: int,
    state: str,
    sheet_type: str,
    part_name: str,
) -> SheetStructure:
    return SheetStructure(
        name=name,
        index=index,
        state=state,
        sheet_type=sheet_type,
        part_name=part_name,
        metrics=SheetMetrics(),
        rows=[],
        columns=[],
        cell_anchors={},
        material_cells=[],
        generated_output_ranges=[],
        kpi_snapshot=KpiColumnSnapshot(
            status="NOT_APPLICABLE",
            notes=(
                ["A sheet normalized as 'KPI' is not a worksheet."]
                if _normalize_text(name) == "kpi"
                else []
            ),
        ),
    )


def read_workbook(path: Path, config: AnalysisConfig) -> WorkbookStructure:
    """Open one XLSX read-only and return a normalized structural profile."""

    try:
        archive = zipfile.ZipFile(path, "r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise WorkbookReadError(f"Cannot open {path.name!r} as an XLSX package: {exc}") from exc

    with archive:
        _validate_archive(archive, config)
        workbook_part = _office_document_part(archive)
        if workbook_part not in archive.namelist():
            raise WorkbookReadError(f"Workbook part {workbook_part!r} is missing.")
        workbook_relationships = _relationships(archive, _rels_part(workbook_part))
        styles_part = None
        strings_part = None
        for target, rel_type in workbook_relationships.values():
            if rel_type.rstrip("/").endswith("/styles"):
                styles_part = _resolve_part(workbook_part, target)
            elif rel_type.rstrip("/").endswith("/sharedStrings"):
                strings_part = _resolve_part(workbook_part, target)
        styles = _parse_styles(archive, styles_part)
        shared_strings = _parse_shared_strings(archive, strings_part)

        try:
            workbook_root = _xml_root(archive, workbook_part)
        except ET.ParseError as exc:
            raise WorkbookReadError(f"Invalid workbook XML: {exc}") from exc

        sheet_elements = [node for node in workbook_root.iter() if _local_name(node.tag) == "sheet"]
        sheets: list[SheetStructure] = []
        warnings: list[str] = []
        seen_names: set[str] = set()
        for index, element in enumerate(sheet_elements, start=1):
            name = _attribute(element, "name") or f"Unnamed sheet {index}"
            folded = name.casefold()
            if folded in seen_names:
                raise WorkbookReadError(f"Workbook contains duplicate sheet name {name!r}.")
            seen_names.add(folded)
            state = _attribute(element, "state") or "visible"
            relation_id = _attribute(element, "id")
            relation = workbook_relationships.get(relation_id or "")
            if not relation:
                raise WorkbookReadError(f"Sheet {name!r} has no valid package relationship.")
            part_name = _resolve_part(workbook_part, relation[0])
            relation_type = relation[1].rstrip("/").rsplit("/", 1)[-1].lower()
            if relation_type == "worksheet":
                sheet = _parse_worksheet(
                    archive,
                    part_name,
                    name,
                    index,
                    state,
                    styles,
                    shared_strings,
                    config,
                )
            else:
                sheet = _empty_nonworksheet(name, index, state, relation_type, part_name)
            sheets.append(sheet)
            warnings.extend(f"{name}: {warning}" for warning in sheet.warnings)

        return WorkbookStructure(path=path, sheets=sheets, warnings=warnings)
