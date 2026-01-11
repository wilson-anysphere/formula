export type CellAddress = { row: number; col: number };

export type RangeAddress = { start: CellAddress; end: CellAddress };

export function colToLetters(col: number): string {
  if (!Number.isInteger(col) || col < 0) {
    throw new Error(`colToLetters: invalid column index ${col}`);
  }

  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function toA1(addr: CellAddress): string {
  return `${colToLetters(addr.col)}${addr.row + 1}`;
}

export function normalizeRange(start: CellAddress, end: CellAddress): RangeAddress {
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return {
    start: { row: startRow, col: startCol },
    end: { row: endRow, col: endCol }
  };
}

export function rangeToA1(range: RangeAddress): string {
  const normalized = normalizeRange(range.start, range.end);
  const start = toA1(normalized.start);
  const end = toA1(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

