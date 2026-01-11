// Formula scripting API typings.
//
// This file is intended for editor tooling (Monaco) and does not need to be
// imported by application code.
//
// NOTE: We intentionally avoid declaring a global `interface Range` because it
// would merge with the DOM `Range` type. Instead, spreadsheet types live under
// `Formula.*` and we provide a global `ScriptContext` alias for convenience.

declare namespace Formula {
  export type CellValue = string | number | boolean | null;

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string | null;
  backgroundColor?: string | null;
}

export interface Range {
  readonly address: string;
  getValues(): Promise<CellValue[][]>;
  setValues(values: CellValue[][]): Promise<void>;
  getValue(): Promise<CellValue>;
  setValue(value: CellValue): Promise<void>;
  getFormat(): Promise<CellFormat>;
  setFormat(format: Partial<CellFormat> | null): Promise<void>;
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
    alert(message: string): Promise<void>;
    confirm(message: string): Promise<boolean>;
    prompt(message: string, defaultValue?: string): Promise<string | null>;
  }

  export interface ScriptContext {
    workbook: Workbook;
    activeSheet: Sheet;
    selection: Range;
    ui: UIHelpers;

    // UI helpers (docs-style).
    alert(message: string): Promise<void>;
    confirm(message: string): Promise<boolean>;
    prompt(message: string, defaultValue?: string): Promise<string | null>;

    // Network + logging helpers (subject to ScriptRuntime permissions).
    fetch: typeof fetch;
    console: Console;
  }
}

type ScriptContext = Formula.ScriptContext;

declare const ctx: ScriptContext;
