export const CLASSIFICATION_SCOPE: Readonly<{
  DOCUMENT: "document";
  SHEET: "sheet";
  RANGE: "range";
  COLUMN: "column";
  CELL: "cell";
}>;

export type CellCoord = { row: number; col: number };
export type CellRange = { start: CellCoord; end: CellCoord };

/**
 * Normalizes a range so that `start` is the top-left and `end` is the bottom-right.
 */
export function normalizeRange(range: CellRange): CellRange;

/**
 * Stable string key for selectors (used for indexing/caching).
 */
export function selectorKey(selector: unknown): string;

export function effectiveCellClassification(
  cellRef: { documentId: string; sheetId: string; row: number; col: number; tableId?: string; columnId?: string },
  records: Array<{ selector: any; classification: any }>
): any;

export function effectiveRangeClassification(
  rangeRef: { documentId: string; sheetId: string; range: { start: { row: number; col: number }; end: { row: number; col: number } } },
  records: Array<{ selector: any; classification: any }>
): any;

export function effectiveDocumentClassification(
  documentId: string,
  records: Array<{ selector: any; classification: any }>
): any;
