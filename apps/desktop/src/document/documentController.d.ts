export class DocumentController {
  // The desktop DocumentController implementation is written in JavaScript with extensive
  // runtime behavior. For cross-package TypeScript tests we only need a permissive surface
  // area (the runtime implementation is authoritative).
  [key: string]: any;
  constructor(...args: any[]);

  getSheetIds(): string[];

  setCellValue(sheetId: string, coord: unknown, value: unknown, options?: unknown): void;
  setCellFormula(sheetId: string, coord: unknown, formula: string | null, options?: unknown): void;
  setRangeFormat(sheetId: string, range: unknown, stylePatch: Record<string, any> | null, options?: unknown): void;
}
