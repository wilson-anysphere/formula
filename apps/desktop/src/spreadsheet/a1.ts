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

export function lettersToCol(letters: string): number | null {
  const normalized = letters.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return null;

  let col = 0;
  for (const ch of normalized) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return col - 1;
}

export function toA1(addr: CellAddress): string {
  return `${colToLetters(addr.col)}${addr.row + 1}`;
}

export function parseA1(a1: string): CellAddress | null {
  const trimmed = a1.trim();
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(trimmed);
  if (!match) return null;

  const col = lettersToCol(match[1]);
  if (col === null) return null;

  const rowNum = Number(match[2]);
  if (!Number.isFinite(rowNum) || rowNum < 1) return null;

  return { row: rowNum - 1, col };
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

