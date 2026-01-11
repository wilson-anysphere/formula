export const CLASSIFICATION_SCOPE: Readonly<{
  DOCUMENT: "document";
  SHEET: "sheet";
  RANGE: "range";
  COLUMN: "column";
  CELL: "cell";
}>;

export function effectiveCellClassification(
  cellRef: { documentId: string; sheetId: string; row: number; col: number; tableId?: string; columnId?: string },
  records: Array<{ selector: any; classification: any }>
): any;

export function effectiveRangeClassification(
  rangeRef: { documentId: string; sheetId: string; range: { start: { row: number; col: number }; end: { row: number; col: number } } },
  records: Array<{ selector: any; classification: any }>
): any;

