import { colToName, fromA1, toA1 as toA1Shared } from "@formula/spreadsheet-frontend/a1";

export type CellAddress = { row: number; col: number };

export type RangeAddress = { start: CellAddress; end: CellAddress };

export function colToLetters(col: number): string {
  if (!Number.isInteger(col) || col < 0) {
    throw new Error(`colToLetters: invalid column index ${col}`);
  }

  return colToName(col);
}

export function lettersToCol(letters: string): number | null {
  const normalized = letters.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return null;
  try {
    return fromA1(`${normalized}1`).col0;
  } catch {
    return null;
  }
}

export function toA1(addr: CellAddress): string {
  return toA1Shared(addr.row, addr.col);
}

export function parseA1(a1: string): CellAddress | null {
  const trimmed = a1.trim();
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(trimmed);
  if (!match) return null;

  try {
    const { row0, col0 } = fromA1(`${match[1]}${match[2]}`);
    return { row: row0, col: col0 };
  } catch {
    return null;
  }
}

export function normalizeRange(start: CellAddress, end: CellAddress): RangeAddress {
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return {
    start: { row: startRow, col: startCol },
    end: { row: endRow, col: endCol },
  };
}

export function rangeToA1(range: RangeAddress): string {
  const normalized = normalizeRange(range.start, range.end);
  const start = toA1(normalized.start);
  const end = toA1(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

export function parseA1Range(text: string): RangeAddress | null {
  const trimmed = text.trim();
  const [startText, endText] = trimmed.split(":");
  const start = parseA1(startText);
  if (!start) return null;

  if (endText === undefined) {
    return { start, end: start };
  }

  const end = parseA1(endText);
  if (!end) return null;

  return normalizeRange(start, end);
}
