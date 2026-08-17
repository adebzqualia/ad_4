"""Focused regression tests for the axis alignment safety properties."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import sys
import unittest


SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from pops_anomaly_detector.alignment import AlignmentUnresolved, align_axis  # noqa: E402
from pops_anomaly_detector.config import AnalysisConfig  # noqa: E402
from pops_anomaly_detector.ooxml import AxisSignature  # noqa: E402


def _signature(index: int, token: str) -> AxisSignature:
    return AxisSignature(
        index=index,
        weights={token: 5.0},
        strong_tokens=frozenset({token}),
        digest=token,
        information=5.0,
    )


def _sequence(tokens: str) -> list[AxisSignature]:
    return [_signature(index, token) for index, token in enumerate(tokens, start=1)]


class AlignmentTests(unittest.TestCase):
    def test_full_and_banded_alignment_return_same_operations(self) -> None:
        sent = _sequence("ABCDEFGHIJKLMNOPQRST")
        received = _sequence("ABXCDEFGHIJKLMPQRSTY")
        full_config = replace(AnalysisConfig(), full_alignment_cells=10_000)
        banded_config = replace(
            AnalysisConfig(),
            full_alignment_cells=1,
            alignment_band=8,
            max_banded_alignment_cells=10_000,
        )

        full = align_axis(sent, received, "ROW", full_config)
        banded = align_axis(sent, received, "ROW", banded_config)

        summarize = lambda result: [
            (operation.operation, operation.start, operation.end)
            for operation in result.operations
        ]
        self.assertEqual(summarize(full), summarize(banded))
        self.assertEqual(
            [(left, right) for left, right, _score in full.pairs],
            [(left, right) for left, right, _score in banded.pairs],
        )

    def test_large_anchor_free_segment_fails_before_unbounded_allocation(self) -> None:
        sent = [
            AxisSignature(i, {"same": 1.0}, frozenset(), "same", 1.0)
            for i in range(1, 501)
        ]
        received = [
            AxisSignature(i, {"other": 1.0}, frozenset(), "other", 1.0)
            for i in range(1, 1001)
        ]
        config = replace(
            AnalysisConfig(),
            full_alignment_cells=100,
            alignment_band=20,
            max_banded_alignment_cells=50_000,
        )

        with self.assertRaises(AlignmentUnresolved):
            align_axis(sent, received, "ROW", config)


if __name__ == "__main__":
    unittest.main()

