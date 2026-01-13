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

  /**
   * Return the cell's stored value and optional formula.
   *
   * Note: Real spreadsheet backends often store both `cell.formula` and a computed `cell.value`.
   * ToolExecutor defaults to treating formula cells as having no value (to match the in-memory
   * workbook model), but can optionally surface/use computed values when
   * `ToolExecutorOptions.include_formula_values` is enabled (with conservative DLP gating).
   */
  getCell(address: CellAddress): CellData;
  setCell(address: CellAddress, cell: CellData): void;

  readRange(range: RangeAddress): CellData[][];
  writeRange(range: RangeAddress, cells: CellData[][]): void;

  /**
   * Apply a formatting patch to a rectangular range.
   *
   * Returns the number of cells the caller *attempted* to format.
   *
   * Implementations should throw an Error if the formatting request cannot be
   * applied (e.g. host safety caps) rather than silently returning `0`.
   * ToolExecutor will surface the failure as `ok:false` with `runtime_error`.
   */
  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number;

  /**
   * Optional chart support. If not provided, `create_chart` tool calls should
   * return a `not_implemented` error.
   */
  createChart?(spec: CreateChartSpec): CreateChartResult;

  getLastUsedRow(sheet: string): number;
  clone(): SpreadsheetApi;
}
