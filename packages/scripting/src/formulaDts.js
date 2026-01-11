/**
 * String form of the `packages/scripting/formula.d.ts` file.
 *
 * The desktop Script Editor panel can register this with Monaco via
 * `addExtraLib(...)` to provide in-editor autocomplete for scripts.
 */
export const FORMULA_API_DTS = `// Auto-generated/hand-maintained Formula scripting API typings.
// This file is intended to be loaded into Monaco as an extra lib.

export type CellValue = string | number | boolean | null;

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string;
  backgroundColor?: string;
}

export interface Range {
  readonly address: string;
  getValues(): Promise<CellValue[][]>;
  setValues(values: CellValue[][]): Promise<void>;
  getValue(): Promise<CellValue>;
  setValue(value: CellValue): Promise<void>;
  getFormat(): Promise<CellFormat>;
  setFormat(format: Partial<CellFormat>): Promise<void>;
}

export interface Sheet {
  readonly name: string;
  getRange(address: string): Range;
}

export interface Workbook {
  getSheet(name: string): Sheet;
  setSelection(sheetName: string, address: string): Promise<void>;
  getSelection(): Promise<{ sheetName: string; address: string }>;
  getActiveSheetName(): Promise<string>;
}

export interface UIHelpers {
  log(...args: unknown[]): void;
}

export interface ScriptContext {
  workbook: Workbook;
  activeSheet: Sheet;
  selection: Range;
  ui: UIHelpers;

  // Network + logging helpers (subject to ScriptRuntime permissions).
  fetch: typeof fetch;
  console: Console;
}

declare const ctx: ScriptContext;
`; 
