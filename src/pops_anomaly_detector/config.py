"""Runtime configuration for workbook comparison."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class AnalysisConfig:
    """Limits are deliberately explicit so oversized files fail visibly."""

    sent_dir: Path = Path("data/sent")
    received_dir: Path = Path("data/received")
    output_dir: Path = Path("reports")
    recursive: bool = False
    max_uncompressed_bytes: int = 1_000_000_000
    max_xml_part_bytes: int = 300_000_000
    max_cells_per_sheet: int = 5_000_000
    max_comparison_cells_per_sheet: int = 500_000
    kpi_header_scan_rows: int = 200
    max_kpi_semantic_cells: int = 250_000
    max_active_rows: int = 100_000
    max_active_columns: int = 16_384
    full_alignment_cells: int = 2_000_000
    max_banded_alignment_cells: int = 20_000_000
    alignment_band: int = 240
    minimum_alignment_score: float = 0.12
    minimum_alignment_coverage: float = 0.10

    def resolved(self, base: Path | None = None) -> "AnalysisConfig":
        root = (base or Path.cwd()).resolve()

        def resolve(path: Path) -> Path:
            return path.resolve() if path.is_absolute() else (root / path).resolve()

        return AnalysisConfig(
            sent_dir=resolve(self.sent_dir),
            received_dir=resolve(self.received_dir),
            output_dir=resolve(self.output_dir),
            recursive=self.recursive,
            max_uncompressed_bytes=self.max_uncompressed_bytes,
            max_xml_part_bytes=self.max_xml_part_bytes,
            max_cells_per_sheet=self.max_cells_per_sheet,
            max_comparison_cells_per_sheet=self.max_comparison_cells_per_sheet,
            kpi_header_scan_rows=self.kpi_header_scan_rows,
            max_kpi_semantic_cells=self.max_kpi_semantic_cells,
            max_active_rows=self.max_active_rows,
            max_active_columns=self.max_active_columns,
            full_alignment_cells=self.full_alignment_cells,
            max_banded_alignment_cells=self.max_banded_alignment_cells,
            alignment_band=self.alignment_band,
            minimum_alignment_score=self.minimum_alignment_score,
            minimum_alignment_coverage=self.minimum_alignment_coverage,
        )
