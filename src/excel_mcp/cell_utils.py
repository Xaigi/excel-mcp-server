import re

from openpyxl.utils import column_index_from_string, get_column_letter

from .exceptions import ValidationError

MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COL = 16_384  # XFD
_CELL_REF_RE = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")


def _column_letter_to_index(col_str: str) -> int:
    try:
        return column_index_from_string(col_str.upper())
    except ValueError as exc:
        raise ValidationError(f"Invalid column reference: {col_str}") from exc


def parse_cell_reference_strict(cell_ref: str) -> tuple[int, int]:
    """Parse a single Excel cell reference with full-string validation."""
    if not isinstance(cell_ref, str) or not cell_ref.strip():
        raise ValidationError("Cell reference must not be empty")

    match = _CELL_REF_RE.fullmatch(cell_ref.strip())
    if not match:
        raise ValidationError(f"Invalid cell reference: {cell_ref}")

    col_str, row_str = match.groups()
    row = int(row_str)
    col = _column_letter_to_index(col_str)

    if row < 1 or row > MAX_EXCEL_ROW:
        raise ValidationError(
            f"Row {row} is out of bounds (1-{MAX_EXCEL_ROW})"
        )
    if col < 1 or col > MAX_EXCEL_COL:
        raise ValidationError(
            f"Column {col_str.upper()} is out of bounds (A-{get_column_letter(MAX_EXCEL_COL)})"
        )
    return row, col


def parse_cell_range_strict(
    start_cell: str,
    end_cell: str | None = None,
) -> tuple[int, int, int, int]:
    """Parse a cell or range and ensure start is at or before end."""
    start_row, start_col = parse_cell_reference_strict(start_cell)
    if end_cell is None:
        return start_row, start_col, start_row, start_col

    end_row, end_col = parse_cell_reference_strict(end_cell)
    if end_row < start_row or end_col < start_col:
        raise ValidationError("End cell cannot be before start cell")
    return start_row, start_col, end_row, end_col


def parse_cell_range(
    cell_ref: str,
    end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into row and column indices."""
    if end_ref:
        start_cell = cell_ref
        end_cell = end_ref
    else:
        start_cell = cell_ref
        end_cell = None

    match = re.match(r"([A-Z]+)([0-9]+)", start_cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {start_cell}")
    col_str, row_str = match.groups()
    start_row = int(row_str)
    start_col = column_index_from_string(col_str)

    if end_cell:
        match = re.match(r"([A-Z]+)([0-9]+)", end_cell.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {end_cell}")
        col_str, row_str = match.groups()
        end_row = int(row_str)
        end_col = column_index_from_string(col_str)
    else:
        end_row = None
        end_col = None

    return start_row, start_col, end_row, end_col


def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format (e.g., 'A1', 'BC123')"""
    if not cell_ref:
        return False

    # Split into column and row parts
    col = row = ""
    for c in cell_ref:
        if c.isalpha():
            if row:  # Letters after numbers not allowed
                return False
            col += c
        elif c.isdigit():
            row += c
        else:
            return False

    return bool(col and row)
