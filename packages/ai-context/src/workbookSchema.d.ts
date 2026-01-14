import type { InferredType } from "./schema.js";

export interface WorkbookSchemaRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export type WorkbookSchemaRectInput =
  | WorkbookSchemaRect
  | { startRow: number; startCol: number; endRow: number; endCol: number }
  | { start: { row: number; col: number }; end: { row: number; col: number } };

/**
 * Generic helper for workbook collection fields (sheets/tables/namedRanges).
 *
 * Note: Some fields are typed more precisely (e.g. `WorkbookSchemaSheets`) to model
 * differences between array vs map inputs, but this alias can be useful for callers
 * that want a single collection type.
 */
export type WorkbookSchemaCollection<T> = ReadonlyArray<T> | Set<T> | Map<unknown, T> | Record<string, T>;
export interface WorkbookSchemaSheet {
  name: string;
}

export interface WorkbookSchemaSheetObjectInput {
  name: string;
  cells?: unknown;
  values?: unknown[][];
  origin?: { row: number; col: number };
  getCell?: (row: number, col: number) => unknown;
  /**
   * Allow host-specific fields (ids, metadata, etc) without tripping TS excess-property checks.
   */
  [key: string]: unknown;
}

export interface WorkbookSchemaKeyedSheetObjectInput {
  /**
   * Sheet name. Optional when the sheet is provided via `Map`/`Record` keyed by name.
   */
  name?: string;
  cells?: unknown;
  values?: unknown[][];
  origin?: { row: number; col: number };
  getCell?: (row: number, col: number) => unknown;
  [key: string]: unknown;
}

export interface WorkbookSchemaCellMapLike {
  get(key: string): unknown;
  [key: string]: unknown;
}

/**
 * Sparse cell map as a plain object keyed by `"row,col"` or `"row:col"`.
 *
 * This is accepted at runtime via heuristics (see `looksLikeSparseCoordKeyedObject`).
 */
export type WorkbookSchemaSparseCellObjectMap = Record<string, unknown>;

export type WorkbookSchemaSheetInput =
  | WorkbookSchemaKeyedSheetObjectInput
  | unknown[][]
  | WorkbookSchemaCellMapLike
  | WorkbookSchemaSparseCellObjectMap
  | string;

/**
 * Supported shapes for `workbook.sheets`.
 *
 * Note: When sheets are provided as a `Map` or plain object keyed by name, the value can
 * be either a full sheet object, a matrix, or a sparse cell map. For array/set forms, a
 * sheet object must include `name` (or you can provide the sheet name directly as a string).
 */
export type WorkbookSchemaSheets =
  | ReadonlyArray<WorkbookSchemaSheetObjectInput | string>
  | Set<WorkbookSchemaSheetObjectInput | string>
  | Map<unknown, WorkbookSchemaSheetInput>
  | Record<string, WorkbookSchemaSheetInput>;

export interface WorkbookSchemaTable {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRect;
  rangeA1: string;
  headers: string[];
  inferredColumnTypes: InferredType[];
  /** Number of data rows (excludes an inferred header row, when present). */
  rowCount: number;
  columnCount: number;
}

export interface WorkbookSchemaNamedRange {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRect;
  rangeA1: string;
}

export interface WorkbookSchemaSummary {
  id: string;
  sheets: WorkbookSchemaSheet[];
  tables: WorkbookSchemaTable[];
  namedRanges: WorkbookSchemaNamedRange[];
}

export interface WorkbookSchemaTableObjectInput {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
  [key: string]: unknown;
}

export interface WorkbookSchemaKeyedTableObjectInput {
  name?: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
  [key: string]: unknown;
}

export type WorkbookSchemaTables =
  | ReadonlyArray<WorkbookSchemaTableObjectInput>
  | Set<WorkbookSchemaTableObjectInput>
  | Map<unknown, WorkbookSchemaKeyedTableObjectInput>
  | Record<string, WorkbookSchemaKeyedTableObjectInput>;

export interface WorkbookSchemaNamedRangeObjectInput {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
  [key: string]: unknown;
}

export interface WorkbookSchemaKeyedNamedRangeObjectInput {
  name?: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
  [key: string]: unknown;
}

export type WorkbookSchemaNamedRanges =
  | ReadonlyArray<WorkbookSchemaNamedRangeObjectInput>
  | Set<WorkbookSchemaNamedRangeObjectInput>
  | Map<unknown, WorkbookSchemaKeyedNamedRangeObjectInput>
  | Record<string, WorkbookSchemaKeyedNamedRangeObjectInput>;

export function extractWorkbookSchema(
  workbook: {
    id: string;
    sheets: WorkbookSchemaSheets;
    tables?: WorkbookSchemaTables;
    namedRanges?: WorkbookSchemaNamedRanges;
  },
  options?: { maxAnalyzeRows?: number; maxAnalyzeCols?: number; signal?: AbortSignal },
): WorkbookSchemaSummary;
