// Formula scripting API typings.
//
// This file is intended for editor tooling (Monaco) and does not need to be
// imported by application code.

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
}

declare const ctx: ScriptContext;

