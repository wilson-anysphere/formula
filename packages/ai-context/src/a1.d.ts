export const EXCEL_MAX_ROWS: number;
export const EXCEL_MAX_COLS: number;

export interface CellRef {
  row: number;
  col: number;
}

export interface CellRange {
  /**
   * Optional sheet name for a sheet-qualified A1 reference.
   *
   * `rangeToA1()` accepts either a raw sheet name (e.g. `"My Sheet"`) or an
   * already-quoted Excel-style sheet name (e.g. `"'My Sheet'"`).
   */
  sheetName?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/**
 * Alias for `CellRange` for callers that prefer an A1-oriented name.
 */
export type A1Range = CellRange;

export function isCellEmpty(value: unknown): boolean;

/**
 * Convert a 0-based column index to A1 letters.
 *
 * - `0 -> "A"`
 * - `25 -> "Z"`
 * - `26 -> "AA"`
 */
export function columnIndexToA1(columnIndex: number): string;

/**
 * Convert A1 column letters to a 0-based column index.
 */
export function a1ToColumnIndex(letters: string): number;

/**
 * Convert a 0-based cell reference to an A1 cell reference (e.g. `{ row: 0, col: 0 } -> "A1"`).
 */
export function cellRefToA1(cell: CellRef): string;

/**
 * Parse an A1 cell reference into a 0-based `{ row, col }` reference.
 *
 * Supports optional `$` absolute markers (e.g. `$A$1`, `$A1`, `A$1`).
 */
export function a1ToCellRef(a1Cell: string): CellRef;

/**
 * Format a 0-based range as an Excel-compatible A1 reference.
 *
 * Supports:
 * - Sheet-qualified references (`Sheet1!A1:B2`, `"'My Sheet'!A1:B2"`)
 * - Whole-column ranges (`A:C`)
 * - Whole-row ranges (`1:10`)
 */
export function rangeToA1(range: CellRange): string;

/**
 * Parse an Excel-style A1 reference into a 0-based range.
 *
 * Supports:
 * - Absolute markers (`$A$1`, `$A1`, `A$1`)
 * - Whole-column ranges (`A:C`)
 * - Whole-row ranges (`1:10`)
 */
export function parseA1Range(a1Range: string): CellRange;

/**
 * Normalize a range so `start* <= end*` for both rows and columns.
 */
export function normalizeRange(range: CellRange): CellRange;

