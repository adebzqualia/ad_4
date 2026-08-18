"""Directory pairing and workbook-level anomaly comparison."""

from __future__ import annotations

import hashlib
import re
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable

from . import __version__
from .alignment import AlignmentUnresolved, align_axis, validate_separable_grid
from .config import AnalysisConfig
from .coordinates import (
    MAX_EXCEL_COLUMN,
    MAX_EXCEL_ROW,
    CellRange,
    index_to_column,
    parse_range_reference,
)
from .formula_logic import (
    FormulaComparisonUnresolved,
    canonicalize_formula,
    formula_compatibility_signature,
    formula_has_ref_error,
    formula_uses_single_compatibility,
    neutralize_compatibility_identity,
    normalize_dynamic_array_identity,
    reference_agnostic_shape,
)
from .models import (
    CellSnapshot,
    CountryMetrics,
    CountryResult,
    FileEvidence,
    Finding,
    KpiColumnSnapshot,
    KpiComparison,
    KpiEntry,
    RunResult,
    RunSummary,
    SheetComparison,
)
from .ooxml import SheetStructure, WorkbookReadError, WorkbookStructure, read_workbook


class InputDiscoveryError(RuntimeError):
    """Input directories are absent or filenames are ambiguous."""


@dataclass(slots=True)
class _CellComparisonContext:
    """One exact-name worksheet pair and its validated logical coordinate map."""

    sent: SheetStructure
    received: SheetStructure
    comparison: SheetComparison
    row_mapping: dict[int, int] | None
    column_mapping: dict[int, int] | None
    reverse_row_mapping: dict[int, int] | None = field(init=False, default=None)
    reverse_column_mapping: dict[int, int] | None = field(init=False, default=None)

    def __post_init__(self) -> None:
        if self.row_mapping is not None:
            self.reverse_row_mapping = {
                actual: expected for expected, actual in self.row_mapping.items()
            }
        if self.column_mapping is not None:
            self.reverse_column_mapping = {
                actual: expected for expected, actual in self.column_mapping.items()
            }

    @property
    def is_mappable(self) -> bool:
        return self.row_mapping is not None and self.column_mapping is not None


def _discover(directory: Path, recursive: bool) -> dict[str, Path]:
    if not directory.is_dir():
        raise InputDiscoveryError(f"Input directory does not exist: {directory}")
    iterator: Iterable[Path] = directory.rglob("*") if recursive else directory.iterdir()
    result: dict[str, Path] = {}
    for path in iterator:
        if not path.is_file() or path.suffix.casefold() != ".xlsx" or path.name.startswith("~$"):
            continue
        key = path.name.casefold()
        if key in result:
            raise InputDiscoveryError(
                f"Multiple files resolve to the same case-insensitive workbook name: "
                f"{result[key]} and {path}. Pairing by filename would be ambiguous."
            )
        result[key] = path
    return result


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _file_evidence(path: Path, side: str, input_root: Path) -> FileEvidence:
    try:
        input_relative = path.resolve().relative_to(input_root.resolve()).as_posix()
    except ValueError:
        # Discovery normally guarantees containment. Keep evidence usable if a
        # caller supplies a path outside the configured root instead of
        # inventing a misleading parent path.
        input_relative = path.name
    stat = path.stat()
    return FileEvidence(
        name=path.name,
        relative_path=f"{side}/{input_relative}",
        size_bytes=stat.st_size,
        modified_at_utc=datetime.fromtimestamp(stat.st_mtime, UTC).isoformat().replace("+00:00", "Z"),
        sha256=_sha256(path),
    )


def _slug(value: str, discriminator: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_value = normalized.encode("ascii", "ignore").decode("ascii").lower()
    cleaned = re.sub(r"[^a-z0-9]+", "-", ascii_value).strip("-") or "country"
    suffix = hashlib.sha256(discriminator.encode("utf-8")).hexdigest()[:8]
    return f"{cleaned}-{suffix}"


def _finding_id(counter: list[int]) -> str:
    counter[0] += 1
    return f"ANOM-{counter[0]:04d}"


def _max_severity(findings: list[Finding]) -> str | None:
    if not findings:
        return None
    rank = {"LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}
    return max(findings, key=lambda item: rank.get(item.severity, 0)).severity


def _axis_position(axis: str, start: int, end: int) -> str:
    """Format positions the way Excel users identify each axis."""

    if axis == "ROW":
        return str(start) if start == end else f"{start}–{end}"
    start_letter = index_to_column(start)
    if start == end:
        return f"{start_letter} ({start})"
    end_letter = index_to_column(end)
    return f"{start_letter}–{end_letter} ({start}–{end})"


def _sheet_lookup(workbook: WorkbookStructure) -> dict[str, SheetStructure]:
    return {sheet.name: sheet for sheet in workbook.sheets}


def _is_kpi_sheet(name: str) -> bool:
    normalized = unicodedata.normalize("NFKC", name)
    return " ".join(normalized.split()).casefold() == "kpi"


def _axis_finding(
    counter: list[int],
    axis: str,
    operation,
    sheet: SheetStructure,
) -> Finding:
    plural = "rows" if axis == "ROW" else "columns"
    unit_label = ("row" if axis == "ROW" else "column") if operation.count == 1 else plural
    action_noun = "insertion" if operation.operation == "ADDED" else "deletion"
    code = f"{axis}S_{'INSERTED' if operation.operation == 'ADDED' else 'DELETED'}"
    coordinate_label = "received" if operation.coordinate_space == "RECEIVED" else "sent"
    position = _axis_position(axis, operation.start, operation.end)
    message = (
        f"{operation.count} {unit_label} {action_noun} inferred in {sheet.name!r} at {coordinate_label} "
        f"{plural} {position}."
    )
    return Finding(
        id=_finding_id(counter),
        category="STRUCTURAL",
        code=code,
        severity="HIGH",
        confidence=operation.confidence,
        scope="SHEET",
        message=message,
        sent_sheet_name=sheet.name,
        received_sheet_name=sheet.name,
        unit_count=operation.count,
        coordinate_space=operation.coordinate_space,
        start=operation.start,
        end=operation.end,
        evidence=list(operation.evidence),
    )


def _kpi_header_evidence(side: str, snapshot: KpiColumnSnapshot) -> list[str]:
    evidence = [f"{side} KPI-header status: {snapshot.status}."]
    if snapshot.header_candidates:
        evidence.append(
            f"{side} literal KPI-header candidates: "
            + ", ".join(snapshot.header_candidates)
            + "."
        )
    evidence.extend(f"{side}: {note}" for note in snapshot.notes)
    return evidence


def _entry_evidence(label: str, entries: list[KpiEntry], limit: int = 40) -> list[str]:
    evidence = [
        f"{label}: {entry.display_value!r} at {entry.coordinate} "
        f"({entry.value_kind}, {entry.confidence.lower()} confidence)."
        for entry in entries[:limit]
    ]
    if len(entries) > limit:
        evidence.append(f"{label}: {len(entries) - limit} additional occurrence(s) omitted.")
    return evidence


def _excess_entries(
    source: list[KpiEntry],
    target_counts: Counter[str],
) -> list[KpiEntry]:
    """Return occurrence-aware surplus entries, retaining useful coordinates."""

    seen: Counter[str] = Counter()
    result: list[KpiEntry] = []
    for entry in source:
        if entry.comparison_key is None:
            continue
        seen[entry.comparison_key] += 1
        if seen[entry.comparison_key] > target_counts[entry.comparison_key]:
            result.append(entry)
    return result


def _occurrence_tags(entries: list[KpiEntry]) -> list[tuple[str, int]]:
    seen: Counter[str] = Counter()
    result: list[tuple[str, int]] = []
    for entry in entries:
        if entry.comparison_key is None:
            continue
        seen[entry.comparison_key] += 1
        result.append((entry.comparison_key, seen[entry.comparison_key]))
    return result


def _compare_kpi_snapshots(
    counter: list[int],
    sent_sheet: SheetStructure,
    received_sheet: SheetStructure,
) -> tuple[KpiComparison, list[Finding]]:
    sent_snapshot = sent_sheet.kpi_snapshot
    received_snapshot = received_sheet.kpi_snapshot
    result = KpiComparison(status="OK", sent=sent_snapshot, received=received_snapshot)
    findings: list[Finding] = []

    for status, code, wording in (
        ("MISSING", "KPI_HEADER_MISSING", "could not be found"),
        ("AMBIGUOUS", "KPI_HEADER_AMBIGUOUS", "is ambiguous"),
    ):
        affected = [
            (side, snapshot)
            for side, snapshot in (
                ("Sent", sent_snapshot),
                ("Received", received_snapshot),
            )
            if snapshot.status == status
        ]
        if not affected:
            continue
        side_words = " and ".join(side.lower() for side, _snapshot in affected)
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code=code,
            severity="HIGH",
            confidence="HIGH",
            scope="SHEET",
            message=(
                f"The literal 'KPI' column header {wording} on the {side_words} side "
                f"of sheet {sent_sheet.name!r}; KPI identity cannot be certified."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(affected),
            evidence=[
                item
                for side, snapshot in affected
                for item in _kpi_header_evidence(side, snapshot)
            ],
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)

    unexpected_statuses = [
        (side, snapshot)
        for side, snapshot in (("Sent", sent_snapshot), ("Received", received_snapshot))
        if snapshot.status not in {"FOUND", "MISSING", "AMBIGUOUS"}
    ]
    if unexpected_statuses:
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code="KPI_HEADER_MISSING",
            severity="HIGH",
            confidence="LOW",
            scope="SHEET",
            message=(
                f"The KPI column in sheet {sent_sheet.name!r} could not be resolved; "
                "KPI identity cannot be certified."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(unexpected_statuses),
            evidence=[
                item
                for side, snapshot in unexpected_statuses
                for item in _kpi_header_evidence(side, snapshot)
            ],
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)

    if sent_snapshot.status != "FOUND" or received_snapshot.status != "FOUND":
        result.status = "UNRESOLVED"
        return result, findings

    unresolved = [
        (side, entry)
        for side, entries in (
            ("Sent", sent_snapshot.entries),
            ("Received", received_snapshot.entries),
        )
        for entry in entries
        if entry.comparison_key is None
    ]
    if unresolved:
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code="KPI_IDENTIFIER_UNRESOLVED",
            severity="HIGH",
            confidence="HIGH",
            scope="SHEET",
            message=(
                f"{len(unresolved)} KPI identifier occurrence(s) in sheet {sent_sheet.name!r} "
                "cannot be compared from stored workbook values."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(unresolved),
            evidence=[
                f"{side}: {entry.display_value!r} at {entry.coordinate} "
                f"({entry.value_kind}, {entry.confidence.lower()} confidence)."
                for side, entry in unresolved[:40]
            ]
            + ([f"{len(unresolved) - 40} additional occurrence(s) omitted."] if len(unresolved) > 40 else [])
            + ["Formulas are not recalculated; only a stored scalar cache can identify a formula result."],
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)
        result.status = "UNRESOLVED"
        return result, findings

    sent_counter: Counter[str] = Counter(
        entry.comparison_key for entry in sent_snapshot.entries if entry.comparison_key is not None
    )
    received_counter: Counter[str] = Counter(
        entry.comparison_key
        for entry in received_snapshot.entries
        if entry.comparison_key is not None
    )
    missing_entries = _excess_entries(sent_snapshot.entries, received_counter)
    unexpected_entries = _excess_entries(received_snapshot.entries, sent_counter)
    result.missing_count = len(missing_entries)
    result.unexpected_count = len(unexpected_entries)

    if missing_entries:
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code="KPI_IDENTIFIER_MISSING",
            severity="HIGH",
            confidence="HIGH",
            scope="SHEET",
            message=(
                f"{len(missing_entries)} KPI identifier occurrence(s) from the sent workbook "
                f"are missing in received sheet {sent_sheet.name!r}."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(missing_entries),
            evidence=_entry_evidence("Missing KPI", missing_entries),
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)

    if unexpected_entries:
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code="KPI_IDENTIFIER_UNEXPECTED",
            severity="HIGH",
            confidence="HIGH",
            scope="SHEET",
            message=(
                f"{len(unexpected_entries)} KPI identifier occurrence(s) occur only in the "
                f"received version of sheet {sent_sheet.name!r}."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(unexpected_entries),
            evidence=_entry_evidence("Unexpected KPI", unexpected_entries),
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)

    if missing_entries or unexpected_entries:
        result.status = "ERROR"
        return result, findings

    sent_tags = _occurrence_tags(sent_snapshot.entries)
    received_tags = _occurrence_tags(received_snapshot.entries)
    if sent_tags != received_tags:
        received_positions = {tag: index for index, tag in enumerate(received_tags)}
        moved: list[tuple[KpiEntry, KpiEntry]] = []
        for index, tag in enumerate(sent_tags):
            received_index = received_positions[tag]
            if received_index != index:
                moved.append(
                    (sent_snapshot.entries[index], received_snapshot.entries[received_index])
                )
        result.reordered_count = len(moved)
        evidence = [
            f"{sent_entry.display_value!r}: sent {sent_entry.coordinate}, "
            f"received {received_entry.coordinate}."
            for sent_entry, received_entry in moved[:40]
        ]
        if len(moved) > 40:
            evidence.append(f"{len(moved) - 40} additional moved occurrence(s) omitted.")
        finding = Finding(
            id=_finding_id(counter),
            category="KPI_INTEGRITY",
            code="KPI_ORDER_CHANGED",
            severity="MEDIUM",
            confidence="HIGH",
            scope="SHEET",
            message=(
                f"The KPI identifiers in sheet {sent_sheet.name!r} have identical membership "
                f"but a different order ({len(moved)} moved occurrence(s))."
            ),
            sent_sheet_name=sent_sheet.name,
            received_sheet_name=received_sheet.name,
            unit_count=len(moved),
            evidence=evidence,
        )
        findings.append(finding)
        result.anomaly_ids.append(finding.id)
        result.status = "ERROR"

    return result, findings


def _reference_error_finding(
    counter: list[int],
    sent_sheet: SheetStructure,
    received_sheet: SheetStructure,
) -> Finding | None:
    sent_count = sent_sheet.metrics.ref_error_count
    received_count = received_sheet.metrics.ref_error_count
    sent_explicit = sent_sheet.metrics.formula_ref_error_count
    received_explicit = received_sheet.metrics.formula_ref_error_count
    if received_explicit <= sent_explicit:
        return None
    delta = received_explicit - sent_explicit
    evidence = [
        f"Sent: {sent_count} unique affected cell(s) "
        f"({sent_sheet.metrics.cached_ref_error_count} cached error; "
        f"{sent_sheet.metrics.formula_ref_error_count} explicit broken-reference formula).",
        f"Received: {received_count} unique affected cell(s) "
        f"({received_sheet.metrics.cached_ref_error_count} cached error; "
        f"{received_sheet.metrics.formula_ref_error_count} explicit broken-reference formula).",
    ]
    if sent_sheet.ref_error_coordinates:
        evidence.append(
            "Sent affected-cell sample: " + ", ".join(sent_sheet.ref_error_coordinates[:25]) + "."
        )
    if received_sheet.ref_error_coordinates:
        evidence.append(
            "Received affected-cell sample: "
            + ", ".join(received_sheet.ref_error_coordinates[:25])
            + "."
        )
    evidence.append(
        "The parser reports stored error values and explicit #REF! formula tokens; it does not recalculate Excel formulas."
    )
    return Finding(
        id=_finding_id(counter),
        category="FORMULA_INTEGRITY",
        code="REFERENCE_ERRORS_INCREASED",
        severity="HIGH",
        confidence="HIGH",
        scope="SHEET",
        message=(
            f"Formula cells containing an explicit #REF! token increased by {delta} in sheet "
            f"{sent_sheet.name!r} ({sent_explicit} sent → {received_explicit} received)."
        ),
        sent_sheet_name=sent_sheet.name,
        received_sheet_name=received_sheet.name,
        unit_count=delta,
        evidence=evidence,
    )


_CELL_FINDING_EVIDENCE_LIMIT = 40


def _meaningful_prefilled_value(cell: CellSnapshot) -> bool:
    """Return whether a sent literal is protected template content.

    Empty cells are not retained as material cells. Numeric zero, textual zero,
    and a lone hyphen are deliberately treated as input placeholders that a
    country may complete without creating an anomaly.
    """

    if cell.has_formula or cell.comparison_key is None:
        return False
    if cell.comparison_key.startswith("NUMBER:"):
        try:
            return Decimal(cell.comparison_key.removeprefix("NUMBER:")) != 0
        except InvalidOperation:
            return False
    if cell.comparison_key.startswith("TEXT:"):
        normalized = unicodedata.normalize("NFKC", cell.display_value or "")
        normalized = " ".join(normalized.split())
        if normalized in {"", "-"}:
            return False
        try:
            if Decimal(normalized) == 0:
                return False
        except InvalidOperation:
            pass
    return True


def _formula_output_ranges(sheet: SheetStructure) -> list[CellRange]:
    """Ranges whose non-anchor cells are calculated outputs, not prefills."""

    ranges: list[CellRange] = list(sheet.generated_output_ranges)
    seen = set(ranges)
    for cell in sheet.material_cells:
        if (
            not cell.has_formula
            or (cell.formula_type or "").casefold() not in {"array", "datatable"}
            or not cell.formula_range
        ):
            continue
        try:
            bounds = parse_range_reference(cell.formula_range)
        except ValueError:
            # The formula itself remains comparable only as far as its parser
            # permits. An invalid output range must not be expanded or guessed.
            continue
        if not (
            bounds.min_row <= cell.row <= bounds.max_row
            and bounds.min_col <= cell.col <= bounds.max_col
        ):
            continue
        if bounds not in seen:
            seen.add(bounds)
            ranges.append(bounds)
    return ranges


def _cells_outside_ranges(
    cells: list[CellSnapshot],
    ranges: list[CellRange],
) -> list[CellSnapshot]:
    """Filter generated rectangles in O((n+r) log C) time.

    A workbook can legitimately contain hundreds of thousands of material
    cells and many calculated-output ranges. Scanning every range for every
    cell would be quadratic. A row sweep with a Fenwick difference tree keeps
    the check bounded by Excel's fixed 16,384-column axis without retaining a
    second potentially huge set of covered coordinates.
    """

    if not cells or not ranges:
        return list(cells)

    events: dict[int, list[tuple[int, int, int]]] = {}
    for bounds in ranges:
        if not (
            1 <= bounds.min_row <= bounds.max_row <= MAX_EXCEL_ROW
            and 1 <= bounds.min_col <= bounds.max_col <= MAX_EXCEL_COLUMN
        ):
            # An invalid internal mask must never hang the Fenwick update or
            # exempt unrelated template values.
            continue
        events.setdefault(bounds.min_row, []).append(
            (bounds.min_col, bounds.max_col, 1)
        )
        events.setdefault(bounds.max_row + 1, []).append(
            (bounds.min_col, bounds.max_col, -1)
        )

    tree = [0] * (MAX_EXCEL_COLUMN + 2)

    def add(index: int, delta: int) -> None:
        while index < len(tree):
            tree[index] += delta
            index += index & -index

    def range_add(left: int, right: int, delta: int) -> None:
        add(left, delta)
        add(right + 1, -delta)

    def point_value(index: int) -> int:
        total = 0
        while index:
            total += tree[index]
            index -= index & -index
        return total

    if not events:
        return list(cells)

    ordered_cells = cells
    if any(
        (left.row, left.col) > (right.row, right.col)
        for left, right in zip(cells, cells[1:])
    ):
        ordered_cells = sorted(cells, key=lambda item: (item.row, item.col))

    event_rows = sorted(events)
    event_index = 0
    outside: list[CellSnapshot] = []
    for cell in ordered_cells:
        while event_index < len(event_rows) and event_rows[event_index] <= cell.row:
            for left, right, delta in events[event_rows[event_index]]:
                range_add(left, right, delta)
            event_index += 1
        if point_value(cell.col) == 0:
            outside.append(cell)
    return outside


def _protected_prefilled_cells(
    sheet: SheetStructure,
    excluded_coordinates: set[str] | None = None,
) -> list[CellSnapshot]:
    """Select meaningful template literals that are safe to compare."""

    excluded = excluded_coordinates or set()
    candidates = [
        cell
        for cell in sheet.material_cells
        if _meaningful_prefilled_value(cell) and cell.coordinate not in excluded
    ]
    return _cells_outside_ranges(candidates, _formula_output_ranges(sheet))


def _kpi_identity_coordinates(sheet: SheetStructure) -> set[str]:
    """Cells already governed by the dedicated KPI-integrity comparison."""

    snapshot = sheet.kpi_snapshot
    coordinates = set(snapshot.header_candidates)
    if snapshot.header_coordinate is not None:
        coordinates.add(snapshot.header_coordinate)
    coordinates.update(entry.coordinate for entry in snapshot.entries)
    return coordinates


def _cell_coordinate(row: int, col: int) -> str:
    return f"{index_to_column(col)}{row}"


def _trim_evidence(value: str, limit: int = 220) -> str:
    compact = " ".join(value.split())
    if len(compact) <= limit:
        return compact
    return compact[: limit - 3] + "..."


def _cell_description(cell: CellSnapshot | None) -> str:
    if cell is None:
        return "<blank>"
    if cell.has_formula:
        if cell.formula_text is not None:
            return "=" + _trim_evidence(cell.formula_text)
        if cell.formula_relative is not None:
            return "<shared formula " + _trim_evidence(cell.formula_relative) + ">"
        return "<unresolved formula>"
    return f"{_trim_evidence(cell.display_value or '<blank>')!r} ({cell.value_kind})"


def _casefold_sheet_lookup(workbook: WorkbookStructure) -> dict[str, SheetStructure]:
    return {sheet.name.casefold(): sheet for sheet in workbook.sheets}


def _formula_reference_resolver(
    *,
    side: str,
    current: _CellComparisonContext,
    contexts: dict[str, _CellComparisonContext],
    sent_sheets: dict[str, SheetStructure],
    received_sheets: dict[str, SheetStructure],
):
    """Build a resolver into the sent workbook's logical coordinate space."""

    if side not in {"SENT", "RECEIVED"}:
        raise ValueError(f"Unsupported formula side: {side!r}")
    side_sheets = sent_sheets if side == "SENT" else received_sheets
    other_sheets = received_sheets if side == "SENT" else sent_sheets

    def resolve(
        sheet_name: str | None,
        row: int | None,
        col: int | None,
    ) -> tuple[str, int | None, int | None] | None:
        if sheet_name is None:
            target = current
        else:
            folded_name = sheet_name.casefold()
            sheet = side_sheets.get(folded_name)
            if sheet is None:
                # External workbook references are outside the local topology.
                # Their stored coordinates are stable for this comparison.
                if sheet_name.startswith("[") and "]" in sheet_name:
                    # OOXML external-book indexes can be renumbered between
                    # packages. Without resolving the externalLinks graph, the
                    # token is not a stable workbook identity.
                    return None
                # A name absent from both workbooks is retained conservatively;
                # a name present on only one side is a renamed/added/deleted
                # local sheet and therefore cannot be mapped safely.
                if folded_name in other_sheets:
                    return None
                return None
            target = contexts.get(sheet.name.casefold())
            if target is None:
                return None

        if not target.is_mappable:
            return None
        assert target.row_mapping is not None
        assert target.column_mapping is not None
        if side == "SENT":
            logical_row = row
            logical_col = col
            if row is not None and row not in target.row_mapping:
                return None
            if col is not None and col not in target.column_mapping:
                return None
        else:
            assert target.reverse_row_mapping is not None
            assert target.reverse_column_mapping is not None
            if row is not None and row not in target.reverse_row_mapping:
                return None
            if col is not None and col not in target.reverse_column_mapping:
                return None
            logical_row = (
                target.reverse_row_mapping.get(row) if row is not None else None
            )
            logical_col = (
                target.reverse_column_mapping.get(col) if col is not None else None
            )
        return f"SHEET:{target.sent.name.casefold()}", logical_row, logical_col

    return resolve


def _formula_is_fundamentally_changed(
    sent_cell: CellSnapshot,
    received_cell: CellSnapshot,
    sent_resolver,
    received_resolver,
) -> bool | None:
    """Return True/False, or None when identity cannot be certified safely."""

    if sent_cell.formula_status not in {"RESOLVED", "RESOLVED_SHARED"}:
        return None
    if received_cell.formula_status not in {"RESOLVED", "RESOLVED_SHARED"}:
        return None
    if formula_has_ref_error(sent_cell.formula_text or sent_cell.formula_shape):
        return None
    if formula_has_ref_error(received_cell.formula_text or received_cell.formula_shape):
        return None

    sent_text = sent_cell.formula_text
    received_text = received_cell.formula_text
    if sent_text is not None and received_text is not None:
        try:
            sent_identity = canonicalize_formula(sent_text, sent_resolver)
            received_identity = canonicalize_formula(received_text, received_resolver)
            if sent_identity == received_identity:
                return False
            sent_dynamic = normalize_dynamic_array_identity(sent_identity)
            received_dynamic = normalize_dynamic_array_identity(received_identity)
            if sent_dynamic == received_dynamic:
                # B4# and _xlfn.ANCHORARRAY(B4) are two exact OOXML
                # serializations of the same dynamic-array spill reference.
                return False
            sent_neutral = neutralize_compatibility_identity(sent_dynamic)
            received_neutral = neutralize_compatibility_identity(received_dynamic)
            if sent_neutral == received_neutral:
                # The syntax-only difference is not an anomaly, but implicit
                # intersection can affect spill semantics, so do not certify it
                # as identical without Excel's calculation engine.
                return None
            if (
                formula_compatibility_signature(sent_text)
                != formula_compatibility_signature(received_text)
                and (
                    formula_uses_single_compatibility(sent_text)
                    or formula_uses_single_compatibility(received_text)
                )
            ):
                return None
            return True
        except (FormulaComparisonUnresolved, ValueError):
            # Unsupported or only partially parsed reference syntax stays
            # unresolved. A target-agnostic fallback here could turn a lexical
            # parser limitation into a false manual-change allegation.
            return None

    # Shared followers have no fabricated A1 formula text and are compared via
    # their validated group masters by the caller. Other textless formula
    # records (notably data tables) cannot certify expression identity.
    return None


def _cell_change_evidence(
    sheet_name: str,
    sent_cell: CellSnapshot,
    received_cell: CellSnapshot | None,
    received_coordinate: str,
) -> str:
    return (
        f"{sheet_name}!{sent_cell.coordinate} -> {sheet_name}!{received_coordinate} | "
        f"sent {_cell_description(sent_cell)} | received {_cell_description(received_cell)}"
    )


def _cell_finding(
    counter: list[int],
    context: _CellComparisonContext,
    code: str,
    changes: list[tuple[CellSnapshot, CellSnapshot | None, str]],
) -> Finding:
    labels = {
        "FORMULA_REPLACED_WITH_VALUE": (
            "FORMULA_INTEGRITY",
            "formula cell(s) from the sent template were replaced with hardcoded values",
        ),
        "FORMULA_REMOVED": (
            "FORMULA_INTEGRITY",
            "formula cell(s) from the sent template are blank in the received workbook",
        ),
        "FORMULA_MODIFIED": (
            "FORMULA_INTEGRITY",
            "formula cell(s) were fundamentally modified",
        ),
        "PREFILLED_VALUE_CHANGED": (
            "VALUE_INTEGRITY",
            "meaningful prefilled value(s) from the sent template were changed",
        ),
    }
    category, wording = labels[code]
    evidence = [
        _cell_change_evidence(
            context.sent.name,
            sent_cell,
            received_cell,
            received_coordinate,
        )
        for sent_cell, received_cell, received_coordinate in changes[
            :_CELL_FINDING_EVIDENCE_LIMIT
        ]
    ]
    if len(changes) > _CELL_FINDING_EVIDENCE_LIMIT:
        evidence.append(
            f"{len(changes) - _CELL_FINDING_EVIDENCE_LIMIT} additional changed cell(s) omitted."
        )
    return Finding(
        id=_finding_id(counter),
        category=category,
        code=code,
        severity="MEDIUM",
        confidence="HIGH",
        scope="SHEET",
        message=f"{len(changes)} {wording} in sheet {context.sent.name!r}.",
        sent_sheet_name=context.sent.name,
        received_sheet_name=context.received.name,
        unit_count=len(changes),
        evidence=evidence,
    )


def _shared_formula_masters(sheet: SheetStructure) -> dict[str, CellSnapshot]:
    grouped: dict[str, list[CellSnapshot]] = {}
    for cell in sheet.material_cells:
        if (
            cell.has_formula
            and (cell.formula_type or "").casefold() == "shared"
            and cell.formula_group
            and cell.formula_text is not None
        ):
            grouped.setdefault(cell.formula_group, []).append(cell)
    return {
        group: cells[0]
        for group, cells in grouped.items()
        if len(cells) == 1
    }


def _shared_group_change(
    sent_cell: CellSnapshot,
    received_cell: CellSnapshot,
    context: _CellComparisonContext,
    sent_masters: dict[str, CellSnapshot],
    received_masters: dict[str, CellSnapshot],
    sent_resolver,
    received_resolver,
) -> bool | None:
    sent_group = sent_cell.formula_group
    received_group = received_cell.formula_group
    if not sent_group or not received_group:
        return None
    sent_master = sent_masters.get(sent_group)
    received_master = received_masters.get(received_group)
    if sent_master is None or received_master is None:
        return None
    assert context.row_mapping is not None
    assert context.column_mapping is not None
    # Comparing explicit masters is safe only when they represent the same
    # logical group coordinate. Otherwise Excel may simply have selected a
    # different master cell while preserving the shared expression pattern.
    if (
        context.row_mapping.get(sent_master.row) != received_master.row
        or context.column_mapping.get(sent_master.col) != received_master.col
    ):
        return None
    return _formula_is_fundamentally_changed(
        sent_master,
        received_master,
        sent_resolver,
        received_resolver,
    )


def _shared_follower_to_explicit_change(
    sent_cell: CellSnapshot,
    received_cell: CellSnapshot,
    sent_masters: dict[str, CellSnapshot],
    sent_resolver,
    received_resolver,
) -> bool | None:
    """Compare a split shared follower without inventing its A1 formula text."""

    if not sent_cell.formula_group or received_cell.formula_text is None:
        return None
    sent_master = sent_masters.get(sent_cell.formula_group)
    if sent_master is None or sent_master.formula_text is None:
        return None
    if formula_has_ref_error(sent_master.formula_text) or formula_has_ref_error(
        received_cell.formula_text
    ):
        return None
    if formula_compatibility_signature(
        sent_master.formula_text
    ) != formula_compatibility_signature(received_cell.formula_text):
        return None
    try:
        sent_identity = canonicalize_formula(sent_master.formula_text, sent_resolver)
        received_identity = canonicalize_formula(
            received_cell.formula_text,
            received_resolver,
        )
    except (FormulaComparisonUnresolved, ValueError):
        return None
    sent_skeleton = reference_agnostic_shape(
        neutralize_compatibility_identity(
            normalize_dynamic_array_identity(sent_identity)
        )
    )
    received_skeleton = reference_agnostic_shape(
        neutralize_compatibility_identity(
            normalize_dynamic_array_identity(received_identity)
        )
    )
    if sent_skeleton is None or received_skeleton is None:
        return None
    return sent_skeleton != received_skeleton


def _compare_cell_integrity(
    counter: list[int],
    context: _CellComparisonContext,
    contexts: dict[str, _CellComparisonContext],
    sent_sheets: dict[str, SheetStructure],
    received_sheets: dict[str, SheetStructure],
) -> list[Finding]:
    """Compare only validated logical sent cells, never raw physical positions."""

    sent_formula_cells = [cell for cell in context.sent.material_cells if cell.has_formula]
    dedicated_kpi_coordinates = (
        _kpi_identity_coordinates(context.sent)
        if _is_kpi_sheet(context.sent.name)
        else set()
    )
    sent_value_cells = _protected_prefilled_cells(
        context.sent,
        dedicated_kpi_coordinates,
    )
    if not context.is_mappable:
        context.comparison.formula_unresolved_count = len(sent_formula_cells)
        context.comparison.value_unresolved_count = len(sent_value_cells)
        return []

    assert context.row_mapping is not None
    assert context.column_mapping is not None
    received_cells = {
        (cell.row, cell.col): cell for cell in context.received.material_cells
    }
    sent_resolver = _formula_reference_resolver(
        side="SENT",
        current=context,
        contexts=contexts,
        sent_sheets=sent_sheets,
        received_sheets=received_sheets,
    )
    received_resolver = _formula_reference_resolver(
        side="RECEIVED",
        current=context,
        contexts=contexts,
        sent_sheets=sent_sheets,
        received_sheets=received_sheets,
    )
    sent_shared_masters = _shared_formula_masters(context.sent)
    received_shared_masters = _shared_formula_masters(context.received)
    shared_outcomes: dict[tuple[str, str], bool | None] = {}
    changes: dict[str, list[tuple[CellSnapshot, CellSnapshot | None, str]]] = {
        "FORMULA_REPLACED_WITH_VALUE": [],
        "FORMULA_REMOVED": [],
        "FORMULA_MODIFIED": [],
        "PREFILLED_VALUE_CHANGED": [],
    }

    for sent_cell in sent_formula_cells:
        received_row = context.row_mapping.get(sent_cell.row)
        received_col = context.column_mapping.get(sent_cell.col)
        if received_row is None or received_col is None:
            context.comparison.formula_unresolved_count += 1
            continue
        received_coordinate = _cell_coordinate(received_row, received_col)
        received_cell = received_cells.get((received_row, received_col))
        if received_cell is None:
            changes["FORMULA_REMOVED"].append(
                (sent_cell, None, received_coordinate)
            )
            continue
        if not received_cell.has_formula:
            changes["FORMULA_REPLACED_WITH_VALUE"].append(
                (sent_cell, received_cell, received_coordinate)
            )
            continue
        if sent_cell.formula_status not in {"RESOLVED", "RESOLVED_SHARED"}:
            context.comparison.formula_unresolved_count += 1
            continue
        if received_cell.formula_status not in {"RESOLVED", "RESOLVED_SHARED"}:
            context.comparison.formula_unresolved_count += 1
            continue
        sent_is_shared = (sent_cell.formula_type or "").casefold() == "shared"
        received_is_shared = (
            received_cell.formula_type or ""
        ).casefold() == "shared"
        if sent_is_shared and received_is_shared:
            group_key = (
                sent_cell.formula_group or "",
                received_cell.formula_group or "",
            )
            if group_key not in shared_outcomes:
                shared_outcomes[group_key] = _shared_group_change(
                    sent_cell,
                    received_cell,
                    context,
                    sent_shared_masters,
                    received_shared_masters,
                    sent_resolver,
                    received_resolver,
                )
            changed = shared_outcomes[group_key]
        elif (
            sent_is_shared
            and sent_cell.formula_text is None
            and received_cell.formula_text is not None
        ):
            changed = _shared_follower_to_explicit_change(
                sent_cell,
                received_cell,
                sent_shared_masters,
                sent_resolver,
                received_resolver,
            )
        else:
            changed = _formula_is_fundamentally_changed(
                sent_cell,
                received_cell,
                sent_resolver,
                received_resolver,
            )
        if changed is None:
            context.comparison.formula_unresolved_count += 1
        elif changed:
            changes["FORMULA_MODIFIED"].append(
                (sent_cell, received_cell, received_coordinate)
            )

    for sent_cell in sent_value_cells:
        received_row = context.row_mapping.get(sent_cell.row)
        received_col = context.column_mapping.get(sent_cell.col)
        if received_row is None or received_col is None:
            context.comparison.value_unresolved_count += 1
            continue
        received_coordinate = _cell_coordinate(received_row, received_col)
        received_cell = received_cells.get((received_row, received_col))
        unchanged = (
            received_cell is not None
            and not received_cell.has_formula
            and received_cell.comparison_key is not None
            and received_cell.comparison_key == sent_cell.comparison_key
        )
        if not unchanged:
            changes["PREFILLED_VALUE_CHANGED"].append(
                (sent_cell, received_cell, received_coordinate)
            )

    context.comparison.formula_changed_count = sum(
        len(changes[code])
        for code in (
            "FORMULA_REPLACED_WITH_VALUE",
            "FORMULA_REMOVED",
            "FORMULA_MODIFIED",
        )
    )
    context.comparison.value_changed_count = len(
        changes["PREFILLED_VALUE_CHANGED"]
    )
    return [
        _cell_finding(counter, context, code, code_changes)
        for code, code_changes in changes.items()
        if code_changes
    ]


def _compare_workbooks(
    sent: WorkbookStructure,
    received: WorkbookStructure,
    config: AnalysisConfig,
) -> tuple[list[SheetComparison], list[Finding], list[str]]:
    sent_by_name = _sheet_lookup(sent)
    received_by_name = _sheet_lookup(received)
    findings: list[Finding] = []
    comparisons: list[SheetComparison] = []
    warnings = [*sent.warnings, *received.warnings]
    counter = [0]
    cell_contexts: dict[str, _CellComparisonContext] = {}

    for sheet in sent.sheets:
        if sheet.name in received_by_name:
            continue
        finding = Finding(
            id=_finding_id(counter),
            category="STRUCTURAL",
            code="SHEET_DELETED",
            severity="HIGH",
            confidence="HIGH",
            scope="WORKBOOK",
            message=f"Sheet {sheet.name!r} exists only in the sent workbook.",
            sent_sheet_name=sheet.name,
            unit_count=1,
            evidence=[f"Sent sheet position: {sheet.index}.", "No exact sheet-name match exists in received."],
        )
        findings.append(finding)
        dedicated_kpi_coordinates = (
            _kpi_identity_coordinates(sheet) if _is_kpi_sheet(sheet.name) else set()
        )
        comparisons.append(
            SheetComparison(
                status="DELETED",
                sent_name=sheet.name,
                received_name=None,
                sent_index=sheet.index,
                received_index=None,
                sent_type=sheet.sheet_type,
                received_type=None,
                sent_metrics=sheet.metrics,
                received_metrics=None,
                formula_unresolved_count=sum(
                    cell.has_formula for cell in sheet.material_cells
                ),
                value_unresolved_count=len(
                    _protected_prefilled_cells(sheet, dedicated_kpi_coordinates)
                ),
                anomaly_ids=[finding.id],
            )
        )

    for sheet in received.sheets:
        if sheet.name in sent_by_name:
            continue
        finding = Finding(
            id=_finding_id(counter),
            category="STRUCTURAL",
            code="SHEET_ADDED",
            severity="HIGH",
            confidence="HIGH",
            scope="WORKBOOK",
            message=f"Sheet {sheet.name!r} exists only in the received workbook.",
            received_sheet_name=sheet.name,
            unit_count=1,
            evidence=[
                f"Received sheet position: {sheet.index}.",
                "No exact sheet-name match exists in sent.",
            ],
        )
        findings.append(finding)
        comparisons.append(
            SheetComparison(
                status="ADDED",
                sent_name=None,
                received_name=sheet.name,
                sent_index=None,
                received_index=sheet.index,
                sent_type=None,
                received_type=sheet.sheet_type,
                sent_metrics=None,
                received_metrics=sheet.metrics,
                anomaly_ids=[finding.id],
            )
        )

    for sent_sheet in sent.sheets:
        received_sheet = received_by_name.get(sent_sheet.name)
        if received_sheet is None:
            continue
        comparison = SheetComparison(
            status="UNCHANGED",
            sent_name=sent_sheet.name,
            received_name=received_sheet.name,
            sent_index=sent_sheet.index,
            received_index=received_sheet.index,
            sent_type=sent_sheet.sheet_type,
            received_type=received_sheet.sheet_type,
            sent_metrics=sent_sheet.metrics,
            received_metrics=received_sheet.metrics,
        )
        if sent_sheet.sheet_type != received_sheet.sheet_type:
            finding = Finding(
                id=_finding_id(counter),
                category="STRUCTURAL",
                code="SHEET_TYPE_CHANGED",
                severity="HIGH",
                confidence="HIGH",
                scope="SHEET",
                message=(
                    f"Sheet {sent_sheet.name!r} changed type from {sent_sheet.sheet_type!r} "
                    f"to {received_sheet.sheet_type!r}."
                ),
                sent_sheet_name=sent_sheet.name,
                received_sheet_name=received_sheet.name,
                evidence=["Rows and columns were not compared for a sheet-type change."],
            )
            findings.append(finding)
            comparison.status = "MODIFIED"
            comparison.anomaly_ids.append(finding.id)
            if sent_sheet.sheet_type == "worksheet":
                dedicated_kpi_coordinates = (
                    _kpi_identity_coordinates(sent_sheet)
                    if _is_kpi_sheet(sent_sheet.name)
                    else set()
                )
                comparison.formula_unresolved_count = sum(
                    cell.has_formula for cell in sent_sheet.material_cells
                )
                comparison.value_unresolved_count = len(
                    _protected_prefilled_cells(
                        sent_sheet,
                        dedicated_kpi_coordinates,
                    )
                )
            comparisons.append(comparison)
            continue
        if sent_sheet.sheet_type != "worksheet":
            comparisons.append(comparison)
            continue

        row_alignment = None
        column_alignment = None
        grid_validation_notes: list[str] = []
        unresolved_reasons: list[str] = []
        try:
            row_alignment = align_axis(sent_sheet.rows, received_sheet.rows, "ROW", config)
        except AlignmentUnresolved as exc:
            unresolved_reasons.append(f"Row structure: {exc}")
        try:
            column_alignment = align_axis(
                sent_sheet.columns, received_sheet.columns, "COLUMN", config
            )
        except AlignmentUnresolved as exc:
            unresolved_reasons.append(f"Column structure: {exc}")

        if row_alignment is not None and column_alignment is not None:
            try:
                grid_validation_notes = validate_separable_grid(
                    sent_sheet,
                    received_sheet,
                    row_alignment,
                    column_alignment,
                )
                comparison.alignment_notes.extend(grid_validation_notes)
            except AlignmentUnresolved as exc:
                unresolved_reasons.append(str(exc))
                row_alignment = None
                column_alignment = None

        if row_alignment is not None:
            comparison.row_operations = row_alignment.operations
            comparison.alignment_notes.extend(row_alignment.notes)
            comparison.alignment_notes.append(
                f"Row alignment similarity: {row_alignment.average_similarity:.0%}; "
                f"information coverage: {row_alignment.information_coverage:.0%}; "
                f"stable anchors: {row_alignment.stable_anchor_count}."
            )
            for operation in row_alignment.operations:
                finding = _axis_finding(counter, "ROW", operation, sent_sheet)
                findings.append(finding)
                comparison.anomaly_ids.append(finding.id)

        if column_alignment is not None:
            comparison.column_operations = column_alignment.operations
            comparison.alignment_notes.extend(column_alignment.notes)
            comparison.alignment_notes.append(
                f"Column alignment similarity: {column_alignment.average_similarity:.0%}; "
                f"information coverage: {column_alignment.information_coverage:.0%}; "
                f"stable anchors: {column_alignment.stable_anchor_count}."
            )
            for operation in column_alignment.operations:
                finding = _axis_finding(counter, "COLUMN", operation, sent_sheet)
                findings.append(finding)
                comparison.anomaly_ids.append(finding.id)

        if unresolved_reasons:
            finding = Finding(
                id=_finding_id(counter),
                category="STRUCTURAL",
                code="STRUCTURE_UNRESOLVED",
                severity="HIGH",
                confidence="LOW",
                scope="SHEET",
                message=(
                    f"The structure of sheet {sent_sheet.name!r} could not be mapped reliably; "
                    "manual review is required."
                ),
                sent_sheet_name=sent_sheet.name,
                received_sheet_name=received_sheet.name,
                unit_count=1,
                evidence=unresolved_reasons,
            )
            findings.append(finding)
            comparison.anomaly_ids.append(finding.id)
            comparison.alignment_notes.extend(unresolved_reasons)

        structural_operations = (
            list(row_alignment.operations) if row_alignment is not None else []
        ) + (
            list(column_alignment.operations) if column_alignment is not None else []
        )
        content_mapping_trusted = (
            row_alignment is not None
            and column_alignment is not None
            and not (
                structural_operations
                and (
                    any(
                        operation.confidence.upper() == "LOW"
                        for operation in structural_operations
                    )
                    or any(
                        note.startswith("Fewer than three unique cell anchors")
                        for note in grid_validation_notes
                    )
                )
            )
        )
        if row_alignment is not None and column_alignment is not None and not content_mapping_trusted:
            comparison.alignment_notes.append(
                "Formula and prefilled-value comparison was skipped because a structural edit lacks enough high-confidence positional evidence."
            )
        cell_contexts[sent_sheet.name.casefold()] = _CellComparisonContext(
            sent=sent_sheet,
            received=received_sheet,
            comparison=comparison,
            row_mapping=(row_alignment.mapping if content_mapping_trusted else None),
            column_mapping=(
                column_alignment.mapping if content_mapping_trusted else None
            ),
        )

        if _is_kpi_sheet(sent_sheet.name):
            kpi_comparison, kpi_findings = _compare_kpi_snapshots(
                counter,
                sent_sheet,
                received_sheet,
            )
            comparison.kpi_comparison = kpi_comparison
            for finding in kpi_findings:
                findings.append(finding)
                comparison.anomaly_ids.append(finding.id)

        reference_finding = _reference_error_finding(counter, sent_sheet, received_sheet)
        if reference_finding is not None:
            findings.append(reference_finding)
            comparison.anomaly_ids.append(reference_finding.id)

        if comparison.anomaly_ids:
            comparison.status = "MODIFIED"
        comparisons.append(comparison)

    # Formula references can cross sheet boundaries, so semantic comparisons
    # start only after every exact-name worksheet pair has a validated mapping.
    sent_casefold = _casefold_sheet_lookup(sent)
    received_casefold = _casefold_sheet_lookup(received)
    for context in cell_contexts.values():
        cell_findings = _compare_cell_integrity(
            counter,
            context,
            cell_contexts,
            sent_casefold,
            received_casefold,
        )
        for finding in cell_findings:
            findings.append(finding)
            context.comparison.anomaly_ids.append(finding.id)
        if cell_findings:
            context.comparison.status = "MODIFIED"

    comparisons.sort(
        key=lambda item: (
            item.sent_index if item.sent_index is not None else 10**9,
            item.received_index if item.received_index is not None else 10**9,
            item.sent_name or item.received_name or "",
        )
    )
    return comparisons, findings, warnings


def _metrics(
    sent_workbook: WorkbookStructure,
    received_workbook: WorkbookStructure,
    comparisons: list[SheetComparison],
    findings: list[Finding],
) -> CountryMetrics:
    affected = {
        finding.sent_sheet_name or finding.received_sheet_name
        for finding in findings
        if finding.sent_sheet_name or finding.received_sheet_name
    }
    kpi_comparisons = [
        comparison.kpi_comparison
        for comparison in comparisons
        if comparison.kpi_comparison is not None
    ]
    return CountryMetrics(
        sent_sheet_count=len(sent_workbook.sheets),
        received_sheet_count=len(received_workbook.sheets),
        affected_sheet_count=len(affected),
        sheets_added=sum(finding.code == "SHEET_ADDED" for finding in findings),
        sheets_deleted=sum(finding.code == "SHEET_DELETED" for finding in findings),
        rows_added=sum(
            operation.count
            for comparison in comparisons
            for operation in comparison.row_operations
            if operation.operation == "ADDED"
        ),
        rows_deleted=sum(
            operation.count
            for comparison in comparisons
            for operation in comparison.row_operations
            if operation.operation == "DELETED"
        ),
        columns_added=sum(
            operation.count
            for comparison in comparisons
            for operation in comparison.column_operations
            if operation.operation == "ADDED"
        ),
        columns_deleted=sum(
            operation.count
            for comparison in comparisons
            for operation in comparison.column_operations
            if operation.operation == "DELETED"
        ),
        sent_ref_errors=sum(sheet.metrics.ref_error_count for sheet in sent_workbook.sheets),
        received_ref_errors=sum(
            sheet.metrics.ref_error_count for sheet in received_workbook.sheets
        ),
        kpi_sent_count=sum(
            len(comparison.sent.entries)
            for comparison in kpi_comparisons
            if comparison.sent is not None and comparison.sent.status == "FOUND"
        ),
        kpi_received_count=sum(
            len(comparison.received.entries)
            for comparison in kpi_comparisons
            if comparison.received is not None and comparison.received.status == "FOUND"
        ),
        kpi_missing_count=sum(comparison.missing_count for comparison in kpi_comparisons),
        kpi_unexpected_count=sum(
            comparison.unexpected_count for comparison in kpi_comparisons
        ),
        formula_changed_count=sum(
            comparison.formula_changed_count for comparison in comparisons
        ),
        formula_unresolved_count=sum(
            comparison.formula_unresolved_count for comparison in comparisons
        ),
        value_changed_count=sum(
            comparison.value_changed_count for comparison in comparisons
        ),
        value_unresolved_count=sum(
            comparison.value_unresolved_count for comparison in comparisons
        ),
        sheet_names_match=(
            {sheet.name for sheet in sent_workbook.sheets}
            == {sheet.name for sheet in received_workbook.sheets}
        ),
        finding_count=len(findings),
    )


def _paired_result(
    key: str,
    sent_path: Path,
    received_path: Path,
    config: AnalysisConfig,
) -> CountryResult:
    display_name = sent_path.stem
    country_id = _slug(display_name, key)
    sent_evidence = _file_evidence(sent_path, "sent", config.sent_dir)
    received_evidence = _file_evidence(
        received_path,
        "received",
        config.received_dir,
    )
    try:
        sent_workbook = read_workbook(sent_path, config)
        received_workbook = read_workbook(received_path, config)
        sheets, findings, warnings = _compare_workbooks(sent_workbook, received_workbook, config)
        metrics = _metrics(sent_workbook, received_workbook, sheets, findings)
        status = "ERROR" if findings else "OK"
        return CountryResult(
            country_id=country_id,
            display_name=display_name,
            report_filename=f"{country_id}.html",
            overall_status=status,
            comparison_state="PAIRED",
            max_anomaly_severity=_max_severity(findings),
            sent_file=sent_evidence,
            received_file=received_evidence,
            sent_sheet_names=[sheet.name for sheet in sent_workbook.sheets],
            received_sheet_names=[sheet.name for sheet in received_workbook.sheets],
            metrics=metrics,
            sheets=sheets,
            findings=findings,
            warnings=warnings,
        )
    except (
        WorkbookReadError,
        OSError,
        ValueError,
        zipfile.BadZipFile,
        RuntimeError,
        EOFError,
    ) as exc:
        return CountryResult(
            country_id=country_id,
            display_name=display_name,
            report_filename=f"{country_id}.html",
            overall_status="ERROR",
            comparison_state="READ_ERROR",
            max_anomaly_severity=None,
            sent_file=sent_evidence,
            received_file=received_evidence,
            errors=[str(exc)],
        )


def _unpaired_result(
    key: str,
    path: Path,
    state: str,
    config: AnalysisConfig,
) -> CountryResult:
    display_name = path.stem
    country_id = _slug(display_name, f"{state}:{key}")
    missing_received = state == "MISSING_RECEIVED"
    error = (
        "No received workbook has the same filename as this sent workbook."
        if missing_received
        else "No sent workbook has the same filename as this received workbook."
    )
    return CountryResult(
        country_id=country_id,
        display_name=display_name,
        report_filename=f"{country_id}.html",
        overall_status="ERROR",
        comparison_state=state,
        max_anomaly_severity=None,
        sent_file=(
            _file_evidence(path, "sent", config.sent_dir)
            if missing_received
            else None
        ),
        received_file=(
            None
            if missing_received
            else _file_evidence(path, "received", config.received_dir)
        ),
        errors=[error],
    )


def analyze_directories(config: AnalysisConfig | None = None) -> RunResult:
    """Pair `.xlsx` files by case-insensitive filename and compare each pair."""

    source_config = config or AnalysisConfig()
    resolved = source_config.resolved()

    def display_directory(path: Path) -> str:
        if not path.is_absolute():
            return path.as_posix()
        try:
            return path.resolve().relative_to(Path.cwd().resolve()).as_posix()
        except ValueError:
            return path.name
    sent_files = _discover(resolved.sent_dir, resolved.recursive)
    received_files = _discover(resolved.received_dir, resolved.recursive)
    keys = sorted(sent_files.keys() | received_files.keys())
    countries: list[CountryResult] = []
    for key in keys:
        sent_path = sent_files.get(key)
        received_path = received_files.get(key)
        if sent_path and received_path:
            countries.append(_paired_result(key, sent_path, received_path, resolved))
        elif sent_path:
            countries.append(
                _unpaired_result(key, sent_path, "MISSING_RECEIVED", resolved)
            )
        elif received_path:
            countries.append(
                _unpaired_result(
                    key,
                    received_path,
                    "UNEXPECTED_RECEIVED",
                    resolved,
                )
            )

    countries.sort(key=lambda item: (item.overall_status != "ERROR", item.display_name.casefold()))
    summary = RunSummary(
        sent_files=len(sent_files),
        received_files=len(received_files),
        matched_pairs=sum(item.comparison_state in {"PAIRED", "READ_ERROR"} for item in countries),
        ok=sum(item.overall_status == "OK" for item in countries),
        error=sum(item.overall_status == "ERROR" for item in countries),
        missing_received=sum(item.comparison_state == "MISSING_RECEIVED" for item in countries),
        unexpected_received=sum(item.comparison_state == "UNEXPECTED_RECEIVED" for item in countries),
        comparison_failed=sum(
            item.comparison_state == "READ_ERROR"
            or any(finding.code == "STRUCTURE_UNRESOLVED" for finding in item.findings)
            for item in countries
        ),
        high_findings=sum(
            finding.severity == "HIGH" for item in countries for finding in item.findings
        ),
        medium_findings=sum(
            finding.severity == "MEDIUM" for item in countries for finding in item.findings
        ),
        affected_countries=sum(bool(item.findings) for item in countries),
    )
    now = datetime.now(UTC)
    run_id = now.strftime("%Y%m%dT%H%M%SZ")
    return RunResult(
        schema_version="1.2",
        run_id=run_id,
        generated_at_utc=now.isoformat().replace("+00:00", "Z"),
        comparator_version=__version__,
        scope=[
            "SHEETS",
            "ROWS",
            "COLUMNS",
            "KPI_IDENTIFIERS",
            "REFERENCE_ERRORS",
            "FORMULAS",
            "PREFILLED_VALUES",
        ],
        sent_directory=display_directory(source_config.sent_dir),
        received_directory=display_directory(source_config.received_dir),
        summary=summary,
        countries=countries,
        warnings=(
            ["No .xlsx files were found in either input directory."] if not countries else []
        ),
    )
