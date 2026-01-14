/**
 * Workbook backend contract shared between desktop (Tauri) and web/WASM.
 *
 * These types are intentionally dependency-free so both environments can import
 * them without pulling in Tauri/DOM/React specifics.
 */
export type SheetVisibility = "visible" | "hidden" | "veryHidden";

// Keep in sync with:
// - `apps/desktop/src/sheets/workbookSheetStore.ts`
// - `apps/desktop/src/document/documentController.js`
// - `apps/desktop/src/workbook/workbook.ts`
export type TabColor = {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
};

export type SheetInfo = {
  id: string;
  name: string;
  /**
   * Optional sheet visibility metadata.
   *
   * Desktop backends may omit this (default to "visible").
   */
  visibility?: SheetVisibility;
  /**
   * Optional sheet tab color metadata.
   *
   * Desktop backends may omit this (default to none).
   */
  tabColor?: TabColor;
};

export type WorkbookInfo = {
  path: string | null;
  origin_path: string | null;
  sheets: SheetInfo[];
};

export type CellValue = {
  value: unknown | null;
  formula: string | null;
  display_value: string;
};

export type RangeData = {
  values: CellValue[][];
  start_row: number;
  start_col: number;
};

export type RangeCellEdit = {
  value: unknown | null;
  formula: string | null;
};

export type SheetUsedRange = {
  start_row: number;
  end_row: number;
  start_col: number;
  end_col: number;
};

export type WorkbookThemePalette = {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  followedHlink: string;
};

export type DefinedNameInfo = {
  name: string;
  refers_to: string;
  sheet_id: string | null;
};

export type TableInfo = {
  name: string;
  sheet_id: string;
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
  columns: string[];
};

export {
  EXCEL_MAX_SHEET_NAME_LEN,
  INVALID_SHEET_NAME_CHARACTERS,
  getSheetNameValidationError,
  getSheetNameValidationErrorMessage,
  type SheetNameValidationError,
  type SheetNameValidationOptions,
} from "./sheetNameValidation.js";

/**
 * WorkbookBackend v1
 *
 * Minimal cross-platform workbook surface used by the UI. Desktop implementations
 * call into Tauri commands, while web implementations route through a WASM engine
 * running in a Worker.
 */
export interface WorkbookBackend {
  newWorkbook(): Promise<WorkbookInfo>;

  /**
   * Open a workbook by filesystem path (desktop-only; optional for web backends).
   *
   * `options.password` is used for password-protected/encrypted workbooks ("Password to open").
   */
  openWorkbook?(path: string, options?: { password?: string }): Promise<WorkbookInfo>;
  openWorkbookFromBytes?(bytes: Uint8Array): Promise<WorkbookInfo>;

  /**
   * Desktop-only helpers (optional for cross-platform consumers).
   */
  getWorkbookThemePalette?(): Promise<WorkbookThemePalette | null>;
  listDefinedNames?(): Promise<DefinedNameInfo[]>;
  listTables?(): Promise<TableInfo[]>;

  getSheetUsedRange(sheetId: string): Promise<SheetUsedRange | null>;

  getRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  }): Promise<RangeData>;

  setCell(params: {
    sheetId: string;
    row: number;
    col: number;
    value: unknown | null;
    formula: string | null;
  }): Promise<void>;

  setRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    values: RangeCellEdit[][];
  }): Promise<void>;

  saveWorkbook?(path?: string, options?: { password?: string }): Promise<void>;
}

type RequiredKeys<T> = {
  [K in keyof T]-?: {} extends Pick<T, K> ? never : K;
}[keyof T];

type IsEqual<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2)
  ? (<T>() => T extends B ? 1 : 2) extends (<T>() => T extends A ? 1 : 2)
    ? true
    : false
  : false;

type AssertTrue<T extends true> = T;

export const WORKBOOK_BACKEND_REQUIRED_METHODS = [
  "newWorkbook",
  "getSheetUsedRange",
  "getRange",
  "setCell",
  "setRange",
] as const satisfies readonly RequiredKeys<WorkbookBackend>[];

// Compile-time guard: keep the runtime list of required methods in sync with the
// interface so shape tests catch real contract drift.
type _WorkbookBackendRequiredMethodsAreComplete = AssertTrue<
  IsEqual<RequiredKeys<WorkbookBackend>, (typeof WORKBOOK_BACKEND_REQUIRED_METHODS)[number]>
>;
