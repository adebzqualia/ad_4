"""Directory pairing and workbook-level structural comparison."""

from __future__ import annotations

import hashlib
import re
import unicodedata
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable

from . import __version__
from .alignment import AlignmentUnresolved, align_axis, validate_separable_grid
from .config import AnalysisConfig
from .models import (
    CountryMetrics,
    CountryResult,
    FileEvidence,
    Finding,
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
    return f"STRUCT-{counter[0]:04d}"


def _sheet_lookup(workbook: WorkbookStructure) -> dict[str, SheetStructure]:
    return {sheet.name: sheet for sheet in workbook.sheets}


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
    position = str(operation.start) if operation.start == operation.end else f"{operation.start}–{operation.end}"
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
            max_anomaly_severity="HIGH" if findings else None,
            sent_file=sent_evidence,
            received_file=received_evidence,
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
        affected_countries=sum(bool(item.findings) for item in countries),
    )
    now = datetime.now(UTC)
    run_id = now.strftime("%Y%m%dT%H%M%SZ")
    return RunResult(
        schema_version="1.0",
        run_id=run_id,
        generated_at_utc=now.isoformat().replace("+00:00", "Z"),
        comparator_version=__version__,
        scope=["SHEETS", "ROWS", "COLUMNS"],
        sent_directory=display_directory(source_config.sent_dir),
        received_directory=display_directory(source_config.received_dir),
        summary=summary,
        countries=countries,
        warnings=(
            ["No .xlsx files were found in either input directory."] if not countries else []
        ),
    )
