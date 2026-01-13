export interface CellRef {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  sheetName?: string;
}

export function isCellEmpty(value: unknown): boolean;

/**
 * 0 -> A, 25 -> Z, 26 -> AA
 */
export function columnIndexToA1(columnIndex: number): string;

export function a1ToColumnIndex(letters: string): number;

export function cellRefToA1(cell: CellRef): string;

export function a1ToCellRef(a1Cell: string): CellRef;

export function rangeToA1(range: CellRange): string;

/**
 * Parse "Sheet1!A1:B2" or "A1" into a 0-indexed range.
 */
export function parseA1Range(a1Range: string): CellRange;

export function normalizeRange(range: CellRange): CellRange;
