export class DocumentController {
  // The desktop DocumentController implementation is written in JavaScript with extensive
  // runtime behavior. For cross-package TypeScript tests we only need a permissive surface
  // area (the runtime implementation is authoritative).
  [key: string]: any;
  constructor(...args: any[]);

  on(event: string, listener: (payload: any) => void): () => void;

  markSaved(): void;
  markDirty(): void;
  readonly isDirty: boolean;
  readonly updateVersion: number;
  readonly contentVersion: number;

  getSheetIds(): string[];
  getVisibleSheetIds(): string[];
  getSheetContentVersion(sheetId: string): number;
  deleteSheet(sheetId: string): void;
  getSheetView(sheetId: string): {
    frozenRows: number;
    frozenCols: number;
    colWidths?: Record<string, number>;
    rowHeights?: Record<string, number>;
    mergedRanges?: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>;
  };
  setFrozen(sheetId: string, frozenRows: number, frozenCols: number, options?: unknown): void;
  setMergedRanges(
    sheetId: string,
    mergedRanges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> | null | undefined,
    options?: unknown,
  ): void;

  getSheetDrawings(sheetId: string): any[];
  setSheetDrawings(sheetId: string, drawings: any[] | null | undefined, options?: unknown): void;

  getCell(sheetId: string, coord: unknown): any;
  getCellFormatStyleIds(sheetId: string, coord: unknown): [number, number, number, number, number];
  getSheetDefaultStyleId(sheetId: string): number;
  getRowStyleId(sheetId: string, row: number): number;
  getColStyleId(sheetId: string, col: number): number;
  getCellFormat(sheetId: string, coord: unknown): Record<string, any>;

  setCellValue(sheetId: string, coord: unknown, value: unknown, options?: unknown): void;
  setCellFormula(sheetId: string, coord: unknown, formula: string | null, options?: unknown): void;
  setRangeFormat(sheetId: string, range: unknown, stylePatch: Record<string, any> | null, options?: unknown): boolean;
  setSheetFormat(sheetId: string, stylePatch: Record<string, any> | null, options?: unknown): void;
  setRowFormat(sheetId: string, row: number, stylePatch: Record<string, any> | null, options?: unknown): void;
  setColFormat(sheetId: string, col: number, stylePatch: Record<string, any> | null, options?: unknown): void;

  beginBatch(options?: unknown): void;
  endBatch(): void;
  cancelBatch(): void;

  applyExternalDeltas(deltas: any[], options?: unknown): void;
  applyExternalSheetViewDeltas(deltas: any[], options?: unknown): void;
  applyExternalDrawingDeltas(deltas: any[], options?: unknown): void;
  applyExternalFormatDeltas(deltas: any[], options?: unknown): void;
  applyExternalRangeRunDeltas(deltas: any[], options?: unknown): void;
}
