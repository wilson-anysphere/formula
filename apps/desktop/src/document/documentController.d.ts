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

  getSheetIds(): string[];
  getSheetView(sheetId: string): { frozenRows: number; frozenCols: number };
  setFrozen(sheetId: string, frozenRows: number, frozenCols: number, options?: unknown): void;

  getCell(sheetId: string, coord: unknown): any;

  setCellValue(sheetId: string, coord: unknown, value: unknown, options?: unknown): void;
  setCellFormula(sheetId: string, coord: unknown, formula: string | null, options?: unknown): void;
  setRangeFormat(sheetId: string, range: unknown, stylePatch: Record<string, any> | null, options?: unknown): void;

  beginBatch(options?: unknown): void;
  endBatch(): void;
  cancelBatch(): void;

  applyExternalDeltas(deltas: any[], options?: unknown): void;
}
