import type { CellCoord, GridLimits, Range } from "./types";

export function clampCell(cell: CellCoord, limits: GridLimits): CellCoord {
  return {
    row: clampInt(cell.row, 0, limits.maxRows - 1),
    col: clampInt(cell.col, 0, limits.maxCols - 1)
  };
}

export function clampRange(range: Range, limits: GridLimits): Range {
  const normalized = normalizeRange(range);
  return {
    startRow: clampInt(normalized.startRow, 0, limits.maxRows - 1),
    endRow: clampInt(normalized.endRow, 0, limits.maxRows - 1),
    startCol: clampInt(normalized.startCol, 0, limits.maxCols - 1),
    endCol: clampInt(normalized.endCol, 0, limits.maxCols - 1)
  };
}

export function normalizeRange(range: Range): Range {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

export function rangeFromCells(a: CellCoord, b: CellCoord): Range {
  return normalizeRange({
    startRow: a.row,
    endRow: b.row,
    startCol: a.col,
    endCol: b.col
  });
}

export function isSingleCellRange(range: Range): boolean {
  return range.startRow === range.endRow && range.startCol === range.endCol;
}

export function rangeArea(range: Range): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

export function cellInRange(cell: CellCoord, range: Range): boolean {
  return (
    cell.row >= range.startRow &&
    cell.row <= range.endRow &&
    cell.col >= range.startCol &&
    cell.col <= range.endCol
  );
}

export function equalsRange(a: Range, b: Range): boolean {
  return (
    a.startRow === b.startRow && a.endRow === b.endRow && a.startCol === b.startCol && a.endCol === b.endCol
  );
}

function clampInt(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, Math.trunc(n)));
}

