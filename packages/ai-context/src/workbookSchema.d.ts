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

export interface WorkbookSchemaSheet {
  name: string;
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

export function extractWorkbookSchema(
  workbook: {
    id: string;
    sheets: Array<{
      name: string;
      cells?: unknown;
      values?: unknown[][];
      origin?: { row: number; col: number };
      getCell?: (row: number, col: number) => unknown;
    }>;
    tables?: Array<{ name: string; sheetName: string; rect: WorkbookSchemaRectInput }>;
    namedRanges?: Array<{ name: string; sheetName: string; rect: WorkbookSchemaRectInput }>;
  },
  options?: { maxAnalyzeRows?: number; maxAnalyzeCols?: number; signal?: AbortSignal },
): WorkbookSchemaSummary;
