export type ExtensionCellRange = {
  /**
   * 0-based row index (inclusive).
   */
  startRow: number;
  /**
   * 0-based column index (inclusive).
   */
  startCol: number;
  /**
   * 0-based row index (inclusive).
   */
  endRow: number;
  /**
   * 0-based column index (inclusive).
   */
  endCol: number;
};

/**
 * Maximum number of cells extensions are allowed to materialize in memory when calling
 * `formula.cells.getSelection()` / `formula.cells.getRange()` / `formula.cells.setRange()`.
 *
 * Extensions can request A1 ranges that cover the full Excel grid. Since these APIs return or
 * accept 2D JavaScript arrays, unbounded ranges can easily OOM the renderer process even when
 * the underlying document is sparse.
 */
export const DEFAULT_EXTENSION_RANGE_CELL_LIMIT = 200_000;

export type ExtensionRangeSize = {
  rows: number;
  cols: number;
  cellCount: number;
};

export function normalizeExtensionRange(range: ExtensionCellRange): ExtensionCellRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, startCol, endRow, endCol };
}

export function getExtensionRangeSize(range: ExtensionCellRange): ExtensionRangeSize {
  const r = normalizeExtensionRange(range);
  const rows = Math.max(0, r.endRow - r.startRow + 1);
  const cols = Math.max(0, r.endCol - r.startCol + 1);
  return { rows, cols, cellCount: rows * cols };
}

export function assertExtensionRangeWithinLimits(
  range: ExtensionCellRange,
  options: { maxCells?: number; label?: string } = {},
): ExtensionRangeSize {
  const size = getExtensionRangeSize(range);
  const maxCells = options.maxCells ?? DEFAULT_EXTENSION_RANGE_CELL_LIMIT;
  if (size.cellCount > maxCells) {
    const label = options.label ?? "Range";
    throw new Error(
      `${label} is too large (${size.rows}x${size.cols}=${size.cellCount} cells). ` +
        `Limit is ${maxCells} cells.`,
    );
  }
  return size;
}

