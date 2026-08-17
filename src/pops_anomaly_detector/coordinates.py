"""Excel coordinate helpers without relying on workbook libraries."""

from __future__ import annotations

import re
from dataclasses import dataclass


MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


@dataclass(frozen=True, slots=True)
class CellRange:
    min_row: int
    min_col: int
    max_row: int
    max_col: int


def column_to_index(label: str) -> int:
    value = 0
    for char in label.upper():
        if not "A" <= char <= "Z":
            raise ValueError(f"Invalid Excel column: {label!r}")
        value = value * 26 + ord(char) - 64
    if not 1 <= value <= MAX_EXCEL_COLUMN:
        raise ValueError(f"Excel column is out of range: {label!r}")
    return value


def index_to_column(index: int) -> str:
    if not 1 <= index <= MAX_EXCEL_COLUMN:
        raise ValueError(f"Excel column index is out of range: {index}")
    parts: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        parts.append(chr(65 + remainder))
    return "".join(reversed(parts))


def parse_cell_reference(reference: str) -> tuple[int, int]:
    match = _CELL_RE.match(reference.strip())
    if not match:
        raise ValueError(f"Invalid Excel cell reference: {reference!r}")
    row = int(match.group(2))
    if not 1 <= row <= MAX_EXCEL_ROW:
        raise ValueError(f"Excel row is out of range: {row}")
    return row, column_to_index(match.group(1))


def parse_range_reference(reference: str) -> CellRange:
    clean = reference.strip()
    if "!" in clean:
        clean = clean.rsplit("!", 1)[1]
    clean = clean.replace("$", "")
    if ":" not in clean:
        row, col = parse_cell_reference(clean)
        return CellRange(row, col, row, col)
    left, right = clean.split(":", 1)
    if left.isalpha() and right.isalpha():
        return CellRange(1, column_to_index(left), MAX_EXCEL_ROW, column_to_index(right))
    if left.isdigit() and right.isdigit():
        start_row, end_row = int(left), int(right)
        if not (1 <= start_row <= end_row <= MAX_EXCEL_ROW):
            raise ValueError(f"Invalid row range: {reference!r}")
        return CellRange(start_row, 1, end_row, MAX_EXCEL_COLUMN)
    start_row, start_col = parse_cell_reference(left)
    end_row, end_col = parse_cell_reference(right)
    if start_row > end_row or start_col > end_col:
        raise ValueError(f"Reversed Excel range: {reference!r}")
    return CellRange(start_row, start_col, end_row, end_col)


def format_active_ref(max_row: int, max_col: int) -> str | None:
    if max_row <= 0 or max_col <= 0:
        return None
    return f"A1:{index_to_column(max_col)}{max_row}"

