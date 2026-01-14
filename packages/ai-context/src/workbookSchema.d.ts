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

export type WorkbookSchemaCollection<T> =
  | ReadonlyArray<T>
  | Map<any, T>
  | Set<T>
  | Record<string, T>;

export interface WorkbookSchemaSheet {
  name: string;
}

export interface WorkbookSchemaSheetInput extends WorkbookSchemaSheet {
  cells?: unknown;
  values?: unknown[][];
  origin?: { row: number; col: number };
  getCell?: (row: number, col: number) => unknown;
}

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

export interface WorkbookSchemaTableInput {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
}

export interface WorkbookSchemaNamedRangeInput {
  name: string;
  sheetName: string;
  rect: WorkbookSchemaRectInput;
}

export function extractWorkbookSchema(
  workbook: {
    id: string;
    sheets: WorkbookSchemaCollection<WorkbookSchemaSheetInput>;
    tables?: WorkbookSchemaCollection<WorkbookSchemaTableInput>;
    namedRanges?: WorkbookSchemaCollection<WorkbookSchemaNamedRangeInput>;
  },
  options?: { maxAnalyzeRows?: number; maxAnalyzeCols?: number; signal?: AbortSignal },
): WorkbookSchemaSummary;
