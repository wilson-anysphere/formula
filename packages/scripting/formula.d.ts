// Formula scripting API typings.
//
// This file is intended for editor tooling (Monaco) and does not need to be
// imported by application code.

type CellValue = string | number | boolean | null;

interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string;
  backgroundColor?: string;
}

interface Range {
  readonly address: string;
  getValues(): Promise<CellValue[][]>;
  setValues(values: CellValue[][]): Promise<void>;
  getValue(): Promise<CellValue>;
  setValue(value: CellValue): Promise<void>;
  getFormat(): Promise<CellFormat>;
  setFormat(format: Partial<CellFormat>): Promise<void>;
}

interface Sheet {
  readonly name: string;
  getRange(address: string): Range;
}

interface Workbook {
  getSheet(name: string): Sheet;
  setSelection(sheetName: string, address: string): Promise<void>;
  getSelection(): Promise<{ sheetName: string; address: string }>;
  getActiveSheetName(): Promise<string>;
}

interface UIHelpers {
  log(...args: unknown[]): void;
}

interface ScriptContext {
  workbook: Workbook;
  activeSheet: Sheet;
  selection: Range;
  ui: UIHelpers;

  // Network + logging helpers (subject to ScriptRuntime permissions).
  fetch: typeof fetch;
  console: Console;
}

declare const ctx: ScriptContext;
