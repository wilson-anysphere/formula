from __future__ import annotations

import re
from typing import Tuple

_CELL_RE = re.compile(r"^([A-Za-z]+)([0-9]+)$")
_RANGE_RE = re.compile(r"^([A-Za-z]+[0-9]+)(?::([A-Za-z]+[0-9]+))?$")


def _col_letters_to_index(letters: str) -> int:
    col = 0
    for ch in letters.upper():
        if ch < "A" or ch > "Z":
            raise ValueError(f"Invalid column letter: {ch!r}")
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1


def _parse_cell(cell: str) -> Tuple[int, int]:
    m = _CELL_RE.match(cell)
    if not m:
        raise ValueError(f"Invalid A1 cell reference: {cell!r}")
    col_letters, row_digits = m.groups()
    row = int(row_digits) - 1
    col = _col_letters_to_index(col_letters)
    if row < 0:
        raise ValueError(f"Invalid row in A1 reference: {cell!r}")
    return row, col


def parse_a1(ref: str) -> Tuple[int, int, int, int]:
    """
    Parse an A1-style range like "A1" or "A1:B10".

    Returns a tuple: (start_row, start_col, end_row, end_col) with 0-indexed
    coordinates.
    """

    m = _RANGE_RE.match(ref.strip())
    if not m:
        raise ValueError(f"Invalid A1 reference: {ref!r}")

    start_cell, end_cell = m.groups()
    start_row, start_col = _parse_cell(start_cell)
    if end_cell is None:
        return start_row, start_col, start_row, start_col

    end_row, end_col = _parse_cell(end_cell)
    if end_row < start_row or end_col < start_col:
        raise ValueError(f"Invalid A1 range (end before start): {ref!r}")

    return start_row, start_col, end_row, end_col

