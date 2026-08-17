"""Monotone sequence alignment for inferred row and column operations."""

from __future__ import annotations

from bisect import bisect_left
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from math import floor, inf, isfinite

from .config import AnalysisConfig
from .models import AxisOperation
from .ooxml import AxisSignature, SheetStructure


class AlignmentUnresolved(RuntimeError):
    """There is not enough consistent evidence to infer a global axis map."""


@dataclass(slots=True)
class AxisAlignment:
    axis: str
    operations: list[AxisOperation]
    pairs: list[tuple[int, int, float]]
    average_similarity: float
    information_coverage: float
    stable_anchor_count: int
    notes: list[str] = field(default_factory=list)

    @property
    def mapping(self) -> dict[int, int]:
        return {expected: actual for expected, actual, _score in self.pairs}


def weighted_similarity(left: AxisSignature, right: AxisSignature) -> float:
    keys = left.weights.keys() | right.weights.keys()
    denominator = sum(max(left.weights.get(key, 0.0), right.weights.get(key, 0.0)) for key in keys)
    if denominator == 0:
        return 1.0
    numerator = sum(min(left.weights.get(key, 0.0), right.weights.get(key, 0.0)) for key in keys)
    return numerator / denominator


def _substitution_cost(left: AxisSignature, right: AxisSignature) -> float:
    if left.digest == right.digest:
        return 0.0
    return 1.05 - 0.85 * weighted_similarity(left, right)


def _best_state(values: tuple[float, float, float], preference: tuple[int, int, int]) -> int:
    rank = {state: position for position, state in enumerate(preference)}
    return min(range(3), key=lambda state: (values[state], rank[state]))


def _traceback(
    traces: list[tuple[int, bytearray]],
    expected_length: int,
    actual_length: int,
    state: int,
) -> list[tuple[str, int | None, int | None]]:
    i, j = expected_length, actual_length
    actions: list[tuple[str, int | None, int | None]] = []
    while i or j:
        row_min, row_trace = traces[i]
        offset = j - row_min
        if offset < 0 or offset * 3 + state >= len(row_trace):
            raise AlignmentUnresolved("The alignment path left its safe dynamic-programming band.")
        previous = row_trace[offset * 3 + state]
        if previous == 255:
            raise AlignmentUnresolved("The alignment traceback reached an unreachable state.")
        if state == 0:
            actions.append(("PAIR", i, j))
            i -= 1
            j -= 1
        elif state == 1:
            actions.append(("DELETE", i, None))
            i -= 1
        else:
            actions.append(("INSERT", None, j))
            j -= 1
        state = previous
    actions.reverse()
    return actions


def _full_affine_alignment(
    expected: list[AxisSignature],
    actual: list[AxisSignature],
) -> list[tuple[str, int | None, int | None]]:
    n, m = len(expected), len(actual)
    if not n:
        return [("INSERT", None, index) for index in range(1, m + 1)]
    if not m:
        return [("DELETE", index, None) for index in range(1, n + 1)]

    gap_open, gap_extend = 0.90, 0.65
    width = m + 1
    traces = [(0, bytearray([255]) * (width * 3)) for _ in range(n + 1)]
    previous_match = [inf] * width
    previous_delete = [inf] * width
    previous_insert = [inf] * width
    previous_match[0] = 0.0
    for j in range(1, width):
        previous_insert[j] = gap_open + (j - 1) * gap_extend
        traces[0][1][j * 3 + 2] = 0 if j == 1 else 2

    for i in range(1, n + 1):
        current_match = [inf] * width
        current_delete = [inf] * width
        current_insert = [inf] * width
        current_delete[0] = gap_open + (i - 1) * gap_extend
        traces[i][1][1] = 0 if i == 1 else 1
        for j in range(1, width):
            substitution = _substitution_cost(expected[i - 1], actual[j - 1])
            match_values = (
                previous_match[j - 1],
                previous_delete[j - 1],
                previous_insert[j - 1],
            )
            match_state = _best_state(match_values, (0, 1, 2))
            current_match[j] = match_values[match_state] + substitution
            traces[i][1][j * 3] = match_state

            delete_values = (
                previous_match[j] + gap_open,
                previous_delete[j] + gap_extend,
                previous_insert[j] + gap_open,
            )
            delete_state = _best_state(delete_values, (1, 0, 2))
            current_delete[j] = delete_values[delete_state]
            traces[i][1][j * 3 + 1] = delete_state

            insert_values = (
                current_match[j - 1] + gap_open,
                current_delete[j - 1] + gap_open,
                current_insert[j - 1] + gap_extend,
            )
            insert_state = _best_state(insert_values, (2, 0, 1))
            current_insert[j] = insert_values[insert_state]
            traces[i][1][j * 3 + 2] = insert_state
        previous_match, previous_delete, previous_insert = (
            current_match,
            current_delete,
            current_insert,
        )

    final_values = (previous_match[m], previous_delete[m], previous_insert[m])
    final_state = _best_state(final_values, (0, 1, 2))
    return _traceback(traces, n, m, final_state)


def _banded_affine_alignment(
    expected: list[AxisSignature],
    actual: list[AxisSignature],
    configured_band: int,
) -> list[tuple[str, int | None, int | None]]:
    n, m = len(expected), len(actual)
    if not n or not m:
        return _full_affine_alignment(expected, actual)
    gap_open, gap_extend = 0.90, 0.65
    band = configured_band
    traces: list[tuple[int, bytearray]] = []
    previous_match: dict[int, float] = {}
    previous_delete: dict[int, float] = {}
    previous_insert: dict[int, float] = {}

    for i in range(0, n + 1):
        center = round(i * m / n)
        start = max(0, center - band)
        end = min(m, center + band)
        if i == 0:
            start = 0
        if i == n:
            end = m
        trace = bytearray([255]) * ((end - start + 1) * 3)
        traces.append((start, trace))
        current_match: dict[int, float] = {}
        current_delete: dict[int, float] = {}
        current_insert: dict[int, float] = {}
        for j in range(start, end + 1):
            offset = (j - start) * 3
            if i == 0 and j == 0:
                current_match[j] = 0.0
                continue
            if i > 0 and j > 0:
                base_values = (
                    previous_match.get(j - 1, inf),
                    previous_delete.get(j - 1, inf),
                    previous_insert.get(j - 1, inf),
                )
                state = _best_state(base_values, (0, 1, 2))
                if isfinite(base_values[state]):
                    current_match[j] = base_values[state] + _substitution_cost(
                        expected[i - 1], actual[j - 1]
                    )
                    trace[offset] = state
            if i > 0:
                delete_values = (
                    previous_match.get(j, inf) + gap_open,
                    previous_delete.get(j, inf) + gap_extend,
                    previous_insert.get(j, inf) + gap_open,
                )
                state = _best_state(delete_values, (1, 0, 2))
                if isfinite(delete_values[state]):
                    current_delete[j] = delete_values[state]
                    trace[offset + 1] = state
            if j > 0:
                insert_values = (
                    current_match.get(j - 1, inf) + gap_open,
                    current_delete.get(j - 1, inf) + gap_open,
                    current_insert.get(j - 1, inf) + gap_extend,
                )
                state = _best_state(insert_values, (2, 0, 1))
                if isfinite(insert_values[state]):
                    current_insert[j] = insert_values[state]
                    trace[offset + 2] = state
        previous_match, previous_delete, previous_insert = (
            current_match,
            current_delete,
            current_insert,
        )

    final_values = (
        previous_match.get(m, inf),
        previous_delete.get(m, inf),
        previous_insert.get(m, inf),
    )
    final_state = _best_state(final_values, (0, 1, 2))
    if not isfinite(final_values[final_state]):
        raise AlignmentUnresolved(
            f"Required alignment displacement exceeded the configured band of {band} positions."
        )
    return _traceback(traces, n, m, final_state)


def _candidate_anchors(
    expected: list[AxisSignature],
    actual: list[AxisSignature],
) -> list[tuple[int, int]]:
    expected_digest: dict[str, list[int]] = defaultdict(list)
    actual_digest: dict[str, list[int]] = defaultdict(list)
    expected_tokens: dict[str, list[int]] = defaultdict(list)
    actual_tokens: dict[str, list[int]] = defaultdict(list)
    for item in expected:
        if item.information >= 3.0:
            expected_digest[item.digest].append(item.index)
        for token in item.strong_tokens:
            expected_tokens[token].append(item.index)
    for item in actual:
        if item.information >= 3.0:
            actual_digest[item.digest].append(item.index)
        for token in item.strong_tokens:
            actual_tokens[token].append(item.index)

    support: dict[tuple[int, int], int] = defaultdict(int)
    for digest, left_positions in expected_digest.items():
        right_positions = actual_digest.get(digest, [])
        if len(left_positions) == len(right_positions) == 1:
            support[(left_positions[0], right_positions[0])] += 3
    for token, left_positions in expected_tokens.items():
        right_positions = actual_tokens.get(token, [])
        if len(left_positions) == len(right_positions) == 1:
            support[(left_positions[0], right_positions[0])] += 1
    if not support:
        return []

    candidates = sorted(support, key=lambda pair: (pair[0], -pair[1]))
    tails: list[int] = []
    tail_candidate: list[int] = []
    previous: list[int] = [-1] * len(candidates)
    for candidate_index, (_left, right) in enumerate(candidates):
        position = bisect_left(tails, right)
        if position == len(tails):
            tails.append(right)
            tail_candidate.append(candidate_index)
        else:
            tails[position] = right
            tail_candidate[position] = candidate_index
        if position:
            previous[candidate_index] = tail_candidate[position - 1]
    if not tail_candidate:
        return []
    cursor = tail_candidate[-1]
    selected: list[tuple[int, int]] = []
    while cursor >= 0:
        selected.append(candidates[cursor])
        cursor = previous[cursor]
    selected.reverse()
    return selected


def _align_segment(
    expected: list[AxisSignature],
    actual: list[AxisSignature],
    config: AnalysisConfig,
) -> list[tuple[str, int | None, int | None]]:
    if len(expected) * len(actual) <= config.full_alignment_cells:
        return _full_affine_alignment(expected, actual)
    if abs(len(expected) - len(actual)) > config.alignment_band:
        raise AlignmentUnresolved(
            f"An anchor-free segment differs by {abs(len(expected) - len(actual)):,} positions, "
            f"beyond the configured alignment band of {config.alignment_band:,}."
        )
    estimated_cells = (len(expected) + 1) * min(
        len(actual) + 1, 2 * config.alignment_band + 1
    )
    if estimated_cells > config.max_banded_alignment_cells:
        raise AlignmentUnresolved(
            f"The anchor-free alignment segment would require approximately "
            f"{estimated_cells:,} banded states, above the configured "
            f"{config.max_banded_alignment_cells:,}-state safety limit."
        )
    return _banded_affine_alignment(expected, actual, config.alignment_band)


def _operation_from_group(
    axis: str,
    kind: str,
    action_start: int,
    action_end: int,
    actions: list[tuple[str, int | None, int | None]],
    expected: list[AxisSignature],
    actual: list[AxisSignature],
    hard_anchors: set[tuple[int, int]],
) -> AxisOperation:
    group = actions[action_start : action_end + 1]
    positions = [
        int(item[2] if kind == "INSERT" else item[1])
        for item in group
    ]
    previous_pair = next(
        (item for item in reversed(actions[:action_start]) if item[0] == "PAIR"),
        None,
    )
    next_pair = next((item for item in actions[action_end + 1 :] if item[0] == "PAIR"), None)

    previous_stable = False
    next_stable = False
    if previous_pair:
        previous_stable = (int(previous_pair[1]), int(previous_pair[2])) in hard_anchors
    if next_pair:
        next_stable = (int(next_pair[1]), int(next_pair[2])) in hard_anchors
    if previous_stable and next_stable:
        confidence = "HIGH"
    elif previous_stable or next_stable:
        confidence = "MEDIUM"
    else:
        confidence = "LOW"
    if next_pair is None:
        confidence = "LOW"

    coordinate_items = actual if kind == "INSERT" else expected
    digest_counts = Counter(item.digest for item in coordinate_items)
    ambiguous_repetition = any(
        digest_counts[coordinate_items[position - 1].digest] > 1
        and not coordinate_items[position - 1].strong_tokens
        for position in positions
    )
    if ambiguous_repetition:
        confidence = "MEDIUM" if confidence == "HIGH" else "LOW"

    axis_word = "row" if axis == "ROW" else "column"
    evidence = [
        f"The monotone {axis_word} alignment requires a contiguous {kind.lower()} gap of {len(positions)}."
    ]
    if previous_pair:
        evidence.append(
            f"Previous surviving anchor maps sent {axis_word} {previous_pair[1]} to received {axis_word} {previous_pair[2]}."
        )
    if next_pair:
        evidence.append(
            f"Next surviving anchor maps sent {axis_word} {next_pair[1]} to received {axis_word} {next_pair[2]}."
        )
    if confidence == "LOW":
        evidence.append(
            "The exact position is weakly supported because stable anchors do not bound the gap, "
            "or because it is a trailing extension/contraction that can resemble content entry/clearing."
        )
    if ambiguous_repetition:
        evidence.append(
            "Equivalent blank or repetitive axis fingerprints make the exact position within the run ambiguous."
        )

    return AxisOperation(
        operation="ADDED" if kind == "INSERT" else "DELETED",
        start=min(positions),
        end=max(positions),
        coordinate_space="RECEIVED" if kind == "INSERT" else "SENT",
        confidence=confidence,
        before_expected=int(previous_pair[1]) if previous_pair else None,
        before_actual=int(previous_pair[2]) if previous_pair else None,
        after_expected=int(next_pair[1]) if next_pair else None,
        after_actual=int(next_pair[2]) if next_pair else None,
        evidence=evidence,
    )


def align_axis(
    expected: list[AxisSignature],
    actual: list[AxisSignature],
    axis: str,
    config: AnalysisConfig,
) -> AxisAlignment:
    axis = axis.upper()
    if axis not in {"ROW", "COLUMN"}:
        raise ValueError(f"Unsupported axis: {axis}")
    if not expected and not actual:
        return AxisAlignment(axis, [], [], 1.0, 1.0, 0)
    if not expected or not actual:
        raise AlignmentUnresolved(
            f"The {axis.lower()} structure exists on only one side, so clearing content cannot be "
            "distinguished from deleting the entire active structure."
        )

    anchors = _candidate_anchors(expected, actual)
    actions: list[tuple[str, int | None, int | None]] = []
    previous_expected = 0
    previous_actual = 0
    for anchor_expected, anchor_actual in [*anchors, (len(expected) + 1, len(actual) + 1)]:
        expected_segment = expected[previous_expected : anchor_expected - 1]
        actual_segment = actual[previous_actual : anchor_actual - 1]
        segment_actions = _align_segment(expected_segment, actual_segment, config)
        for action, left, right in segment_actions:
            actions.append(
                (
                    action,
                    left + previous_expected if left is not None else None,
                    right + previous_actual if right is not None else None,
                )
            )
        if anchor_expected <= len(expected) and anchor_actual <= len(actual):
            actions.append(("PAIR", anchor_expected, anchor_actual))
        previous_expected = anchor_expected
        previous_actual = anchor_actual

    pairs: list[tuple[int, int, float]] = []
    stable_count = len(anchors)
    informative_scores: list[float] = []
    for action, left, right in actions:
        if action != "PAIR":
            continue
        left_item = expected[int(left) - 1]
        right_item = actual[int(right) - 1]
        score = weighted_similarity(left_item, right_item)
        pairs.append((int(left), int(right), score))
        if left_item.information > 0 or right_item.information > 0:
            informative_scores.append(score)
    has_information = any(item.information > 0 for item in expected) or any(
        item.information > 0 for item in actual
    )
    average = (
        sum(informative_scores) / len(informative_scores)
        if informative_scores
        else 0.0 if has_information else 1.0
    )
    matched_information = sum(
        min(expected[left - 1].information, actual[right - 1].information) * score
        for left, right, score in pairs
    )
    information_denominator = max(
        sum(item.information for item in expected),
        sum(item.information for item in actual),
    )
    coverage = (
        matched_information / information_denominator if information_denominator else 1.0
    )

    informative_expected = sum(item.information > 0 for item in expected)
    informative_actual = sum(item.information > 0 for item in actual)
    if (
        informative_expected
        and informative_actual
        and len(anchors) < 2
        and (
            average < config.minimum_alignment_score
            or coverage < config.minimum_alignment_coverage
        )
    ):
        raise AlignmentUnresolved(
            f"The best {axis.lower()} alignment has {average:.0%} average evidence similarity, "
            f"{coverage:.0%} matched-information coverage, and fewer than two stable anchors."
        )

    operations: list[AxisOperation] = []
    index = 0
    while index < len(actions):
        kind = actions[index][0]
        if kind == "PAIR":
            index += 1
            continue
        end = index
        previous_position = int(actions[index][2] if kind == "INSERT" else actions[index][1])
        while end + 1 < len(actions) and actions[end + 1][0] == kind:
            next_position = int(
                actions[end + 1][2] if kind == "INSERT" else actions[end + 1][1]
            )
            if next_position != previous_position + 1:
                break
            previous_position = next_position
            end += 1
        operations.append(
            _operation_from_group(
                axis,
                kind,
                index,
                end,
                actions,
                expected,
                actual,
                set(anchors),
            )
        )
        index = end + 1

    notes: list[str] = []
    low_confidence = sum(operation.confidence == "LOW" for operation in operations)
    if low_confidence:
        notes.append(
            f"{low_confidence} inferred {axis.lower()} operation(s) have low positional confidence."
        )
    gross_edits = sum(operation.count for operation in operations)
    if gross_edits and coverage < config.minimum_alignment_coverage and len(anchors) < 3:
        raise AlignmentUnresolved(
            f"The inferred {axis.lower()} edit path changes {gross_edits} position(s) but preserves "
            f"only {coverage:.0%} of structural information."
        )
    return AxisAlignment(axis, operations, pairs, average, coverage, stable_count, notes)


def validate_separable_grid(
    expected: SheetStructure,
    actual: SheetStructure,
    row_alignment: AxisAlignment,
    column_alignment: AxisAlignment,
) -> list[str]:
    """Reject a global row/column story contradicted by unique cell anchors.

    A full-row edit gives every surviving stable cell in an expected row the
    same actual row.  A local ``insert cells down`` operation does not.
    """

    shared: list[tuple[tuple[int, int], tuple[int, int]]] = []
    for token, expected_positions in expected.cell_anchors.items():
        actual_positions = actual.cell_anchors.get(token, [])
        if len(expected_positions) == len(actual_positions) == 1:
            shared.append((expected_positions[0], actual_positions[0]))
    if len(shared) < 3:
        return [
            "Fewer than three unique cell anchors survived; whole-axis attribution relies primarily on sequence evidence."
        ]

    row_map = row_alignment.mapping
    column_map = column_alignment.mapping
    row_residuals = 0
    column_residuals = 0
    row_checked = 0
    column_checked = 0
    actual_rows_by_expected: dict[int, set[int]] = defaultdict(set)
    actual_columns_by_expected: dict[int, set[int]] = defaultdict(set)
    for (expected_row, expected_col), (actual_row, actual_col) in shared:
        actual_rows_by_expected[expected_row].add(actual_row)
        actual_columns_by_expected[expected_col].add(actual_col)
        row_checked += 1
        row_residuals += int(
            expected_row not in row_map or row_map[expected_row] != actual_row
        )
        column_checked += 1
        column_residuals += int(
            expected_col not in column_map or column_map[expected_col] != actual_col
        )

    split_rows = sum(len(values) > 1 for values in actual_rows_by_expected.values())
    split_columns = sum(len(values) > 1 for values in actual_columns_by_expected.values())
    row_limit = 0 if row_checked < 10 else floor(row_checked * 0.10)
    column_limit = 0 if column_checked < 10 else floor(column_checked * 0.10)
    if split_rows or split_columns or row_residuals > row_limit or column_residuals > column_limit:
        details = (
            f"Unique cell anchors contradict a separable whole-row/whole-column map "
            f"(row residuals {row_residuals}/{row_checked}, column residuals "
            f"{column_residuals}/{column_checked}, split rows {split_rows}, split columns {split_columns})."
        )
        raise AlignmentUnresolved(
            details
            + " This can be caused by a local cell-range insertion/deletion, cut/paste, or a heavily rebuilt sheet."
        )
    return [
        f"The global grid map is consistent with {len(shared)} unique surviving cell anchor(s)."
    ]
