"""Directory pairing and workbook-level anomaly comparison."""

from __future__ import annotations

import hashlib
import re
import unicodedata
import zipfile
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable

from . import __version__
from .alignment import AlignmentUnresolved, align_axis, validate_separable_grid
from .config import AnalysisConfig
from .coordinates import index_to_column
from .models import (
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


def _file_evidence(path: Path, side: str) -> FileEvidence:
    stat = path.stat()
    return FileEvidence(
        name=path.name,
        relative_path=f"{side}/{path.name}",
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
    if received_count <= sent_count:
        return None
    delta = received_count - sent_count
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
            f"Stored or explicit #REF! cells increased by {delta} in sheet "
            f"{sent_sheet.name!r} ({sent_count} sent → {received_count} received)."
        ),
        sent_sheet_name=sent_sheet.name,
        received_sheet_name=received_sheet.name,
        unit_count=delta,
        evidence=evidence,
    )


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
            comparisons.append(comparison)
            continue
        if sent_sheet.sheet_type != "worksheet":
            comparisons.append(comparison)
            continue

        row_alignment = None
        column_alignment = None
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
                comparison.alignment_notes.extend(
                    validate_separable_grid(
                        sent_sheet,
                        received_sheet,
                        row_alignment,
                        column_alignment,
                    )
                )
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
    sent_evidence = _file_evidence(sent_path, "sent")
    received_evidence = _file_evidence(received_path, "received")
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


def _unpaired_result(key: str, path: Path, state: str) -> CountryResult:
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
        sent_file=_file_evidence(path, "sent") if missing_received else None,
        received_file=None if missing_received else _file_evidence(path, "received"),
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
            countries.append(_unpaired_result(key, sent_path, "MISSING_RECEIVED"))
        elif received_path:
            countries.append(_unpaired_result(key, received_path, "UNEXPECTED_RECEIVED"))

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
        schema_version="1.1",
        run_id=run_id,
        generated_at_utc=now.isoformat().replace("+00:00", "Z"),
        comparator_version=__version__,
        scope=["SHEETS", "ROWS", "COLUMNS", "KPI_IDENTIFIERS", "REFERENCE_ERRORS"],
        sent_directory=display_directory(source_config.sent_dir),
        received_directory=display_directory(source_config.received_dir),
        summary=summary,
        countries=countries,
        warnings=(
            ["No .xlsx files were found in either input directory."] if not countries else []
        ),
    )
