import type { GridLimits, Range } from "../selection/types";

export const DEFAULT_FORMATTING_APPLY_CELL_LIMIT = 100_000;
// `DocumentController.setRangeFormat` enumerates rows when applying full-width row formatting
// (rowStyleIds) and hard-caps this work at 50k rows by default. Keep the UI selection-size
// guard aligned so we block early (with a toast) rather than silently skipping.
export const DEFAULT_FORMATTING_BAND_ROW_LIMIT = 50_000;

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

function isFullSheetBand(range: Range, limits: GridLimits): boolean {
  return isFullColumnBand(range, limits) && isFullRowBand(range, limits);
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
  const maxBandRows = DEFAULT_FORMATTING_BAND_ROW_LIMIT;
  const totalCells = selectionCellCount(selectionRanges);

  const allRangesBand =
    selectionRanges.length > 0 &&
    selectionRanges.every((r) => {
      if (isFullSheetBand(r, limits)) return true;
      if (isFullColumnBand(r, limits)) return true;
      if (isFullRowBand(r, limits)) {
        const nr = normalizeSelectionRange(r);
        const rowCount = nr.endRow - nr.startRow + 1;
        return rowCount <= maxBandRows;
      }
      return false;
    });

  // Huge full-row selections (that are not the entire sheet) are not scalable with the current
  // layered formatting model. Block these early rather than relying on internal controller caps.
  let hasOversizedRowBand = false;
  for (const r of selectionRanges) {
    if (!isFullRowBand(r, limits)) continue;
    if (isFullSheetBand(r, limits)) continue;
    const nr = normalizeSelectionRange(r);
    const rowCount = nr.endRow - nr.startRow + 1;
    if (rowCount > maxBandRows) {
      hasOversizedRowBand = true;
      break;
    }
  }

  const allowed = !hasOversizedRowBand && (totalCells <= maxCells || allRangesBand);
  return { allowed, totalCells, allRangesBand };
}
