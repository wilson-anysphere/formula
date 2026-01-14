export type InferredType = "empty" | "number" | "boolean" | "date" | "string" | "formula" | "mixed";

export interface ColumnSchema {
  name: string;
  type: InferredType;
  sampleValues: string[];
}

export interface TableSchema {
  name: string;
  range: string;
  columns: ColumnSchema[];
  rowCount: number;
}

export interface NamedRangeSchema {
  name: string;
  range: string;
}

export interface DataRegionSchema {
  range: string;
  hasHeader: boolean;
  headers: string[];
  inferredColumnTypes: InferredType[];
  rowCount: number;
  columnCount: number;
}

export interface SheetSchema {
  name: string;
  tables: TableSchema[];
  namedRanges: NamedRangeSchema[];
  dataRegions: DataRegionSchema[];
}

export function inferCellType(value: unknown): InferredType;
export function inferColumnType(values: unknown[], options?: { signal?: AbortSignal }): InferredType;
export function isLikelyHeaderRow(rowValues: unknown[], nextRowValues?: unknown[]): boolean;
export function detectDataRegions(
  values: unknown[][],
  options?: { maxCells?: number; signal?: AbortSignal },
): Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>;
export function extractSheetSchema(sheet: {
  name: string;
  values: unknown[][];
  /**
   * Optional coordinate origin (0-based) for the provided `values` matrix.
   *
   * When `values` is a cropped window of a larger sheet (e.g. a capped used-range
   * sample), `origin` lets schema extraction produce correct absolute A1 ranges.
   */
  origin?: { row: number; col: number };
  namedRanges?: Array<NamedRangeSchema & { [key: string]: unknown }>;
  tables?: Array<{ name: string; range: string; [key: string]: unknown }>;
}, options?: {
  signal?: AbortSignal;
  /**
   * Maximum number of data rows (excluding the header row) to scan when inferring column
   * types. Defaults to 500.
   */
  maxAnalyzeRows?: number;
  /**
   * Maximum number of unique sample values to capture per column. Defaults to 3.
   */
  maxSampleValuesPerColumn?: number;
}): SheetSchema;
