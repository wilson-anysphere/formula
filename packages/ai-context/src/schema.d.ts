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
export function inferColumnType(values: unknown[]): InferredType;
export function detectDataRegions(values: unknown[][]): Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>;
export function extractSheetSchema(sheet: {
  name: string;
  values: unknown[][];
  namedRanges?: NamedRangeSchema[];
  tables?: { name: string; range: string }[];
}): SheetSchema;

