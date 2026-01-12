import type { GridLimits, Range } from "../selection/types";

export const DEFAULT_FORMATTING_APPLY_CELL_LIMIT = 100_000;

export type NormalizedSelectionRange = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export function normalizeSelectionRange(range: Range): NormalizedSelectionRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

export function selectionRangeCellCount(range: Range): number {
  const r = normalizeSelectionRange(range);
  const rows = r.endRow - r.startRow + 1;
  const cols = r.endCol - r.startCol + 1;
  if (rows <= 0 || cols <= 0) return 0;
  return rows * cols;
}

export function selectionCellCount(ranges: Range[]): number {
  let total = 0;
  for (const r of ranges) {
    total += selectionRangeCellCount(r);
  }
  return total;
}

export function isFullColumnBand(range: Range, limits: GridLimits): boolean {
  const r = normalizeSelectionRange(range);
  return r.startRow === 0 && r.endRow === limits.maxRows - 1;
}

export function isFullRowBand(range: Range, limits: GridLimits): boolean {
  const r = normalizeSelectionRange(range);
  return r.startCol === 0 && r.endCol === limits.maxCols - 1;
}

export function isBandSelectionRange(range: Range, limits: GridLimits): boolean {
  return isFullColumnBand(range, limits) || isFullRowBand(range, limits);
}

export type FormattingSelectionSizeDecision = {
  allowed: boolean;
  totalCells: number;
  /**
   * True when *every* selection range is a full-row/full-column/full-sheet band.
   */
  allRangesBand: boolean;
};

export function evaluateFormattingSelectionSize(
  selectionRanges: Range[],
  limits: GridLimits,
  options: { maxCells?: number } = {},
): FormattingSelectionSizeDecision {
  const maxCells = options.maxCells ?? DEFAULT_FORMATTING_APPLY_CELL_LIMIT;
  const totalCells = selectionCellCount(selectionRanges);
  const allRangesBand = selectionRanges.length > 0 && selectionRanges.every((r) => isBandSelectionRange(r, limits));

  const allowed = totalCells <= maxCells || allRangesBand;
  return { allowed, totalCells, allRangesBand };
}

