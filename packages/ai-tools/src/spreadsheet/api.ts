import type { CellAddress, RangeAddress } from "./a1.ts";
import type { CellData, CellFormat } from "./types.ts";

export interface CellEntry {
  address: CellAddress;
  cell: CellData;
}

export type ChartType = "bar" | "line" | "pie" | "scatter" | "area";

export interface CreateChartSpec {
  chart_type: ChartType;
  data_range: string;
  title?: string;
  position?: string;
}

export interface CreateChartResult {
  chart_id: string;
}

export interface SpreadsheetApi {
  listSheets(): string[];
  listNonEmptyCells(sheet?: string): CellEntry[];

  getCell(address: CellAddress): CellData;
  setCell(address: CellAddress, cell: CellData): void;

  readRange(range: RangeAddress): CellData[][];
  writeRange(range: RangeAddress, cells: CellData[][]): void;

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number;

  /**
   * Optional chart support. If not provided, `create_chart` tool calls should
   * return a `not_implemented` error.
   */
  createChart?(spec: CreateChartSpec): CreateChartResult;

  getLastUsedRow(sheet: string): number;
  clone(): SpreadsheetApi;
}
