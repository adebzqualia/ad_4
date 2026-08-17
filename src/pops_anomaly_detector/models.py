"""Serializable result models used by the comparator and report renderer."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(slots=True)
class FileEvidence:
    name: str
    relative_path: str
    size_bytes: int
    modified_at_utc: str
    sha256: str


@dataclass(slots=True)
class AxisOperation:
    operation: str
    start: int
    end: int
    coordinate_space: str
    confidence: str
    before_expected: int | None = None
    before_actual: int | None = None
    after_expected: int | None = None
    after_actual: int | None = None
    evidence: list[str] = field(default_factory=list)

    @property
    def count(self) -> int:
        return self.end - self.start + 1


@dataclass(slots=True)
class KpiEntry:
    display_value: str
    comparison_key: str | None
    coordinate: str
    row: int
    value_kind: str
    confidence: str


@dataclass(slots=True)
class KpiColumnSnapshot:
    status: str
    header_coordinate: str | None = None
    header_candidates: list[str] = field(default_factory=list)
    entries: list[KpiEntry] = field(default_factory=list)
    duplicate_keys: dict[str, list[str]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)


@dataclass(slots=True)
class KpiComparison:
    status: str
    sent: KpiColumnSnapshot | None = None
    received: KpiColumnSnapshot | None = None
    missing_count: int = 0
    unexpected_count: int = 0
    reordered_count: int = 0
    anomaly_ids: list[str] = field(default_factory=list)


@dataclass(slots=True)
class SheetMetrics:
    active_rows: int = 0
    active_columns: int = 0
    content_rows: int = 0
    content_columns: int = 0
    populated_cells: int = 0
    formula_cells: int = 0
    styled_blank_cells: int = 0
    merged_ranges: int = 0
    table_ranges: int = 0
    ref_error_count: int = 0
    cached_ref_error_count: int = 0
    formula_ref_error_count: int = 0
    active_ref: str | None = None
    content_ref: str | None = None
    declared_dimension: str | None = None


@dataclass(slots=True)
class SheetComparison:
    status: str
    sent_name: str | None
    received_name: str | None
    sent_index: int | None
    received_index: int | None
    sent_type: str | None
    received_type: str | None
    sent_metrics: SheetMetrics | None
    received_metrics: SheetMetrics | None
    kpi_comparison: KpiComparison | None = None
    row_operations: list[AxisOperation] = field(default_factory=list)
    column_operations: list[AxisOperation] = field(default_factory=list)
    alignment_notes: list[str] = field(default_factory=list)
    anomaly_ids: list[str] = field(default_factory=list)


@dataclass(slots=True)
class Finding:
    id: str
    category: str
    code: str
    severity: str
    confidence: str
    scope: str
    message: str
    sent_sheet_name: str | None = None
    received_sheet_name: str | None = None
    unit_count: int = 1
    coordinate_space: str | None = None
    start: int | None = None
    end: int | None = None
    evidence: list[str] = field(default_factory=list)


@dataclass(slots=True)
class CountryMetrics:
    sent_sheet_count: int = 0
    received_sheet_count: int = 0
    affected_sheet_count: int = 0
    sheets_added: int = 0
    sheets_deleted: int = 0
    rows_added: int = 0
    rows_deleted: int = 0
    columns_added: int = 0
    columns_deleted: int = 0
    sent_ref_errors: int = 0
    received_ref_errors: int = 0
    kpi_sent_count: int = 0
    kpi_received_count: int = 0
    kpi_missing_count: int = 0
    kpi_unexpected_count: int = 0
    sheet_names_match: bool = True
    finding_count: int = 0

    @property
    def row_net_delta(self) -> int:
        return self.rows_added - self.rows_deleted

    @property
    def column_net_delta(self) -> int:
        return self.columns_added - self.columns_deleted

    @property
    def sheet_net_delta(self) -> int:
        return self.sheets_added - self.sheets_deleted


@dataclass(slots=True)
class CountryResult:
    country_id: str
    display_name: str
    report_filename: str
    overall_status: str
    comparison_state: str
    max_anomaly_severity: str | None
    sent_file: FileEvidence | None
    received_file: FileEvidence | None
    sent_sheet_names: list[str] = field(default_factory=list)
    received_sheet_names: list[str] = field(default_factory=list)
    metrics: CountryMetrics = field(default_factory=CountryMetrics)
    sheets: list[SheetComparison] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class RunSummary:
    sent_files: int = 0
    received_files: int = 0
    matched_pairs: int = 0
    ok: int = 0
    error: int = 0
    missing_received: int = 0
    unexpected_received: int = 0
    comparison_failed: int = 0
    high_findings: int = 0
    medium_findings: int = 0
    affected_countries: int = 0


@dataclass(slots=True)
class RunResult:
    schema_version: str
    run_id: str
    generated_at_utc: str
    comparator_version: str
    scope: list[str]
    sent_directory: str
    received_directory: str
    summary: RunSummary
    countries: list[CountryResult]
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        category_definitions = (
            ("STRUCTURAL", "Structural anomalies", "HIGH"),
            ("KPI_INTEGRITY", "KPI integrity anomalies", "HIGH"),
            ("FORMULA_INTEGRITY", "Formula integrity anomalies", "HIGH"),
        )
        severity_rank = {"LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}
        for serialized, country in zip(payload["countries"], self.countries, strict=True):
            serialized["report_href"] = f"countries/{country.report_filename}"
            serialized["metrics"].update(
                {
                    "sheet_net_delta": country.metrics.sheet_net_delta,
                    "row_net_delta": country.metrics.row_net_delta,
                    "column_net_delta": country.metrics.column_net_delta,
                }
            )
            category_results: list[dict[str, Any]] = []
            for code, label, default_severity in category_definitions:
                findings = [item for item in country.findings if item.category == code]
                severity = max(
                    (item.severity for item in findings),
                    key=lambda item: severity_rank.get(item, 0),
                    default=default_severity,
                )
                category_results.append(
                    {
                        "code": code,
                        "label": label,
                        "severity": severity,
                        "status": "ERROR" if findings else "OK",
                        "finding_count": len(findings),
                    }
                )
            serialized["category_results"] = category_results
        return payload
