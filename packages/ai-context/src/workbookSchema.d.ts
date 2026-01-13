import type { InferredType } from "./schema.js";

export interface WorkbookSchemaRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

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
      cells?: any;
      values?: any;
      getCell?: (row: number, col: number) => any;
    }>;
    tables?: Array<{ name: string; sheetName: string; rect: WorkbookSchemaRect }>;
    namedRanges?: Array<{ name: string; sheetName: string; rect: WorkbookSchemaRect }>;
  },
  options?: { maxAnalyzeRows?: number; signal?: AbortSignal },
): WorkbookSchemaSummary;

