"""Conservative, alignment-aware Excel formula identity helpers."""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Callable

from .coordinates import MAX_EXCEL_ROW, column_to_index, parse_cell_reference


class FormulaComparisonUnresolved(RuntimeError):
    """A formula contains a reference that cannot be mapped safely."""


# The callback receives a decoded sheet name (or None for the current sheet)
# and one cell/axis coordinate. It returns a stable logical sheet key and
# logical coordinates, or None when the reference is not safely mappable.
ReferenceResolver = Callable[
    [str | None, int | None, int | None],
    tuple[str, int | None, int | None] | None,
]


_SHEET_TOKEN = r"'(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Za-z_\\][A-Za-z0-9_.\\]*"
_REFERENCE_RE = re.compile(
    rf"(?<![A-Za-z0-9_.])"
    rf"(?:(?P<sheet>{_SHEET_TOKEN})!)?"
    rf"(?:"
    rf"(?P<cell_first>\$?[A-Za-z]{{1,3}}\$?[0-9]{{1,7}})"
    rf"(?::(?P<cell_second>\$?[A-Za-z]{{1,3}}\$?[0-9]{{1,7}}))?"
    rf"|(?P<col_first>\$?[A-Za-z]{{1,3}}):(?P<col_second>\$?[A-Za-z]{{1,3}})"
    rf"|(?P<row_first>\$?[0-9]{{1,7}}):(?P<row_second>\$?[0-9]{{1,7}})"
    rf")"
    rf"(?![A-Za-z0-9_.(])",
    re.IGNORECASE,
)
_NUMBER_RE = re.compile(
    r"(?<![A-Z0-9_.])"
    r"(?P<number>(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:E[+-]?[0-9]+)?)"
    r"(?![A-Z0-9_.])",
    re.IGNORECASE,
)
_SHEET_BANG_RE = re.compile(rf"(?P<sheet>{_SHEET_TOKEN})!", re.IGNORECASE)
_EXTERNAL_SHEET_RE = re.compile(
    r"(?:'\[[^\]]+\](?:[^']|'')+'|\[[^\]]+\][A-Za-z_\\][A-Za-z0-9_.\\]*)!",
    re.IGNORECASE,
)


def _formula_segments(formula: str):
    """Yield code/string segments while respecting both Excel quote forms.

    Double quotes delimit string literals. Single quotes delimit sheet or
    external-workbook identifiers and remain inside code so A1 references can
    still be parsed, but double quotes inside them must not alter lexer state.
    """

    start = 0
    index = 0
    in_identifier = False
    while index < len(formula):
        char = formula[index]
        if char == "'":
            if in_identifier and index + 1 < len(formula) and formula[index + 1] == "'":
                index += 2
                continue
            in_identifier = not in_identifier
            index += 1
            continue
        if char != '"' or in_identifier:
            index += 1
            continue
        if index > start:
            yield False, formula[start:index]
        string_start = index
        index += 1
        while index < len(formula):
            if formula[index] != '"':
                index += 1
                continue
            if index + 1 < len(formula) and formula[index + 1] == '"':
                index += 2
                continue
            index += 1
            break
        yield True, formula[string_start:index]
        start = index
    if start < len(formula):
        yield False, formula[start:]


def _compact_code(segment: str) -> str:
    """Normalize code while preserving quoted identifier contents exactly."""

    result: list[str] = []
    index = 0
    in_identifier = False
    while index < len(segment):
        char = segment[index]
        if char == "'":
            result.append(char)
            if in_identifier and index + 1 < len(segment) and segment[index + 1] == "'":
                result.append("'")
                index += 2
                continue
            in_identifier = not in_identifier
        elif in_identifier:
            result.append(char)
        elif not char.isspace():
            result.append(char.upper())
        index += 1
    return "".join(result)


def _has_intersection_whitespace(segment: str) -> bool:
    """Detect Excel's semantic whitespace intersection operator."""

    in_identifier = False
    index = 0
    while index < len(segment):
        char = segment[index]
        if char == "'":
            if in_identifier and index + 1 < len(segment) and segment[index + 1] == "'":
                index += 2
                continue
            in_identifier = not in_identifier
            index += 1
            continue
        if in_identifier or not char.isspace():
            index += 1
            continue
        left = index - 1
        while left >= 0 and segment[left].isspace():
            left -= 1
        right = index + 1
        while right < len(segment) and segment[right].isspace():
            right += 1
        left_char = segment[left] if left >= 0 else ""
        right_char = segment[right] if right < len(segment) else ""
        if (
            bool(left_char)
            and bool(right_char)
            and (left_char.isalnum() or left_char in "$._])")
            and (right_char.isalnum() or right_char in "$_[('")
        ):
            return True
        index = right
    return False


def _validate_supported_code(segment: str, compact: str) -> None:
    if _has_intersection_whitespace(segment):
        raise FormulaComparisonUnresolved(
            "The Excel intersection operator is not normalized safely."
        )
    for match in _SHEET_BANG_RE.finditer(compact):
        decoded = _decoded_sheet(match.group("sheet")) or ""
        if ":" in decoded or (match.start() > 0 and compact[match.start() - 1] == ":"):
            raise FormulaComparisonUnresolved("A 3-D sheet reference is not supported.")
    without_external_identifiers = _EXTERNAL_SHEET_RE.sub("EXTERNAL!", compact)
    if "[" in without_external_identifiers or "]" in without_external_identifiers:
        raise FormulaComparisonUnresolved(
            "Structured references are not normalized safely."
        )


def compact_formula(formula: str) -> str:
    """Normalize case/whitespace outside string literals only."""

    parts: list[str] = []
    for is_string, segment in _formula_segments(formula.lstrip("=")):
        parts.append(segment if is_string else _compact_code(segment))
    return "".join(parts)


def formula_has_ref_error(formula: str | None) -> bool:
    """Return whether ``#REF!`` occurs outside strings/quoted identifiers."""

    if formula is None:
        return False
    for is_string, segment in _formula_segments(formula):
        if is_string:
            continue
        index = 0
        in_identifier = False
        outside: list[str] = []
        while index < len(segment):
            char = segment[index]
            if char == "'":
                if (
                    in_identifier
                    and index + 1 < len(segment)
                    and segment[index + 1] == "'"
                ):
                    index += 2
                    continue
                in_identifier = not in_identifier
            elif not in_identifier:
                outside.append(char)
            index += 1
        if "#REF!" in "".join(outside).upper():
            return True
    return False


def formula_compatibility_signature(formula: str | None) -> tuple[int, int, int]:
    """Count compatibility markers outside string literals."""

    if formula is None:
        return (0, 0, 0)
    xlfn = xlws = implicit = 0
    for is_string, segment in _formula_segments(formula):
        if is_string:
            continue
        outside: list[str] = []
        index = 0
        in_identifier = False
        while index < len(segment):
            char = segment[index]
            if char == "'":
                if (
                    in_identifier
                    and index + 1 < len(segment)
                    and segment[index + 1] == "'"
                ):
                    index += 2
                    continue
                in_identifier = not in_identifier
            elif not in_identifier:
                outside.append(char)
            index += 1
        compact = _compact_code("".join(outside))
        xlfn += compact.count("_XLFN.")
        xlws += compact.count("_XLWS.")
        implicit += compact.count("@")
    return (xlfn, xlws, implicit)


def formula_uses_single_compatibility(formula: str | None) -> bool:
    """Return whether formula code uses Excel's legacy implicit-intersection wrapper."""

    if formula is None:
        return False
    for is_string, segment in _formula_segments(formula):
        if is_string:
            continue
        outside: list[str] = []
        index = 0
        in_identifier = False
        while index < len(segment):
            char = segment[index]
            if char == "'":
                if (
                    in_identifier
                    and index + 1 < len(segment)
                    and segment[index + 1] == "'"
                ):
                    index += 2
                    continue
                in_identifier = not in_identifier
            elif not in_identifier:
                outside.append(char)
            index += 1
        if "_XLFN.SINGLE(" in _compact_code("".join(outside)):
            return True
    return False


def neutralize_compatibility_identity(identity: str) -> str:
    """Remove save-time markers only from code, never strings/reference keys."""

    result: list[str] = []
    index = 0
    while index < len(identity):
        if identity.startswith("REF[", index):
            start = index
            index += 4
            depth = 1
            while index < len(identity) and depth:
                if identity[index] == "[":
                    depth += 1
                elif identity[index] == "]":
                    depth -= 1
                index += 1
            result.append(identity[start:index])
            continue
        if identity[index] == '"':
            start = index
            index += 1
            while index < len(identity):
                if identity[index] != '"':
                    index += 1
                    continue
                if index + 1 < len(identity) and identity[index + 1] == '"':
                    index += 2
                    continue
                index += 1
                break
            result.append(identity[start:index])
            continue
        if identity.startswith("_XLFN.", index):
            index += len("_XLFN.")
            continue
        if identity.startswith("_XLWS.", index):
            index += len("_XLWS.")
            continue
        if identity[index] == "@":
            index += 1
            continue
        result.append(identity[index])
        index += 1
    return "".join(result)


def normalize_dynamic_array_identity(identity: str) -> str:
    """Normalize Excel's exact legacy ANCHORARRAY wrapper to spill syntax."""

    prefix = "_XLFN.ANCHORARRAY("
    result: list[str] = []
    index = 0
    while index < len(identity):
        if identity.startswith("REF[", index):
            start = index
            index += 4
            depth = 1
            while index < len(identity) and depth:
                if identity[index] == "[":
                    depth += 1
                elif identity[index] == "]":
                    depth -= 1
                index += 1
            result.append(identity[start:index])
            continue
        if identity[index] == '"':
            start = index
            index += 1
            while index < len(identity):
                if identity[index] != '"':
                    index += 1
                    continue
                if index + 1 < len(identity) and identity[index + 1] == '"':
                    index += 2
                    continue
                index += 1
                break
            result.append(identity[start:index])
            continue
        if identity.startswith(prefix, index):
            argument_start = index + len(prefix)
            if identity.startswith("REF[", argument_start):
                argument_end = argument_start + 4
                depth = 1
                while argument_end < len(identity) and depth:
                    if identity[argument_end] == "[":
                        depth += 1
                    elif identity[argument_end] == "]":
                        depth -= 1
                    argument_end += 1
                if (
                    depth == 0
                    and argument_end < len(identity)
                    and identity[argument_end] == ")"
                ):
                    result.append(identity[argument_start:argument_end] + "#")
                    index = argument_end + 1
                    continue
        result.append(identity[index])
        index += 1
    return "".join(result)


def _decoded_sheet(token: str | None) -> str | None:
    if token is None:
        return None
    if token.startswith("'") and token.endswith("'"):
        return token[1:-1].replace("''", "'")
    return token


def _cell_parts(reference: str) -> tuple[int, int, bool, bool]:
    row, col = parse_cell_reference(reference)
    col_absolute = reference.startswith("$")
    without_col_marker = reference[1:] if col_absolute else reference
    row_absolute = "$" in without_col_marker
    return row, col, row_absolute, col_absolute


def _normalize_number(match: re.Match[str]) -> str:
    raw = match.group("number")
    try:
        value = Decimal(raw)
    except InvalidOperation:
        return raw
    if not value.is_finite():
        return raw
    if value == 0:
        value = Decimal(0)
    return str(value.normalize())


def _normalize_identity_numbers(identity: str) -> str:
    """Normalize constants without touching generated reference payloads."""

    result: list[str] = []
    code: list[str] = []

    def flush() -> None:
        if code:
            result.append(_NUMBER_RE.sub(_normalize_number, "".join(code)))
            code.clear()

    index = 0
    while index < len(identity):
        if identity.startswith("REF[", index):
            flush()
            start = index
            index += 4
            depth = 1
            while index < len(identity) and depth:
                if identity[index] == "[":
                    depth += 1
                elif identity[index] == "]":
                    depth -= 1
                index += 1
            result.append(identity[start:index])
            continue
        if identity[index] == "'":
            flush()
            start = index
            index += 1
            while index < len(identity):
                if identity[index] != "'":
                    index += 1
                    continue
                if index + 1 < len(identity) and identity[index + 1] == "'":
                    index += 2
                    continue
                index += 1
                break
            result.append(identity[start:index])
            continue
        code.append(identity[index])
        index += 1
    flush()
    return "".join(result)


def canonicalize_formula(formula: str, resolver: ReferenceResolver) -> str:
    """Return formula identity in a stable logical-coordinate namespace.

    The caller supplies the sent/received coordinate mapping. Any reference to
    a deleted, inserted, renamed, or otherwise unresolved workbook coordinate
    causes a conservative unresolved result instead of a false anomaly.
    """

    def resolve(
        sheet: str | None,
        row: int | None,
        col: int | None,
    ) -> tuple[str, int | None, int | None]:
        result = resolver(sheet, row, col)
        if result is None:
            target = sheet or "the current sheet"
            raise FormulaComparisonUnresolved(
                f"Reference to {target!r} at row {row!r}, column {col!r} is not mappable."
            )
        return result

    def replace_reference(match: re.Match[str]) -> str:
        sheet = _decoded_sheet(match.group("sheet"))
        first_cell = match.group("cell_first")
        if first_cell:
            row, col, row_absolute, col_absolute = _cell_parts(first_cell)
            sheet_key, logical_row, logical_col = resolve(sheet, row, col)
            assert logical_row is not None and logical_col is not None
            first = (
                f"C{'$' if col_absolute else '~'}{logical_col}"
                f"R{'$' if row_absolute else '~'}{logical_row}"
            )
            second_cell = match.group("cell_second")
            if second_cell:
                row2, col2, row_abs2, col_abs2 = _cell_parts(second_cell)
                sheet_key2, logical_row2, logical_col2 = resolve(sheet, row2, col2)
                if sheet_key2 != sheet_key:
                    raise FormulaComparisonUnresolved(
                        "A range endpoint resolved to a different logical sheet."
                    )
                assert logical_row2 is not None and logical_col2 is not None
                second = (
                    f"C{'$' if col_abs2 else '~'}{logical_col2}"
                    f"R{'$' if row_abs2 else '~'}{logical_row2}"
                )
                return f"REF[{sheet_key}|{first}:{second}]"
            return f"REF[{sheet_key}|{first}]"

        first_col = match.group("col_first")
        if first_col:
            col1 = column_to_index(first_col.replace("$", ""))
            col2_token = match.group("col_second")
            assert col2_token is not None
            col2 = column_to_index(col2_token.replace("$", ""))
            sheet_key, _row, logical_col1 = resolve(sheet, None, col1)
            sheet_key2, _row2, logical_col2 = resolve(sheet, None, col2)
            if sheet_key2 != sheet_key or logical_col1 is None or logical_col2 is None:
                raise FormulaComparisonUnresolved(
                    "A whole-column range is not safely mappable."
                )
            return (
                f"REF[{sheet_key}|C{'$' if first_col.startswith('$') else '~'}{logical_col1}:"
                f"C{'$' if col2_token.startswith('$') else '~'}{logical_col2}]"
            )

        first_row = match.group("row_first")
        second_row = match.group("row_second")
        assert first_row is not None and second_row is not None
        row1 = int(first_row.replace("$", ""))
        row2 = int(second_row.replace("$", ""))
        if not 1 <= row1 <= MAX_EXCEL_ROW or not 1 <= row2 <= MAX_EXCEL_ROW:
            raise FormulaComparisonUnresolved("A whole-row reference is out of range.")
        sheet_key, logical_row1, _col = resolve(sheet, row1, None)
        sheet_key2, logical_row2, _col2 = resolve(sheet, row2, None)
        if sheet_key2 != sheet_key or logical_row1 is None or logical_row2 is None:
            raise FormulaComparisonUnresolved("A whole-row range is not safely mappable.")
        return (
            f"REF[{sheet_key}|R{'$' if first_row.startswith('$') else '~'}{logical_row1}:"
            f"R{'$' if second_row.startswith('$') else '~'}{logical_row2}]"
        )

    parts: list[str] = []
    for is_string, segment in _formula_segments(formula.lstrip("=")):
        if is_string:
            parts.append(segment)
            continue
        compact = _compact_code(segment)
        _validate_supported_code(segment, compact)
        if "#REF!" in compact:
            raise FormulaComparisonUnresolved(
                "A #REF! token prevents reliable fundamental-formula comparison."
            )
        try:
            replaced = _REFERENCE_RE.sub(replace_reference, compact)
        except ValueError as exc:
            raise FormulaComparisonUnresolved(
                "The formula contains an A1-looking token outside Excel's grid."
            ) from exc
        if "!" in replaced:
            raise FormulaComparisonUnresolved(
                "The formula contains an unsupported sheet or 3-D reference."
            )
        parts.append(_normalize_identity_numbers(replaced))
    return "".join(parts)


def reference_agnostic_shape(shape: str | None) -> str | None:
    """Remove reference targets while retaining functions/operators/constants."""

    if shape is None:
        return None
    result: list[str] = []
    index = 0
    while index < len(shape):
        if not shape.startswith("REF[", index):
            result.append(shape[index])
            index += 1
            continue
        result.append("REF")
        index += 4
        depth = 1
        while index < len(shape) and depth:
            if shape[index] == "[":
                depth += 1
            elif shape[index] == "]":
                depth -= 1
            index += 1
        if depth:
            return None
    skeleton = "".join(result)
    normalized: list[str] = []
    for is_string, segment in _formula_segments(skeleton):
        normalized.append(
            segment if is_string else _NUMBER_RE.sub(_normalize_number, segment)
        )
    return "".join(normalized)
