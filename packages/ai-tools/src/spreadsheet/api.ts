import type { CellAddress, RangeAddress } from "./a1.js";
import type { CellData, CellFormat } from "./types.js";

export interface CellEntry {
  address: CellAddress;
  cell: CellData;
}

export interface SpreadsheetApi {
  listSheets(): string[];
  listNonEmptyCells(sheet?: string): CellEntry[];

  getCell(address: CellAddress): CellData;
  setCell(address: CellAddress, cell: CellData): void;

  readRange(range: RangeAddress): CellData[][];
  writeRange(range: RangeAddress, cells: CellData[][]): void;

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number;

  getLastUsedRow(sheet: string): number;
  clone(): SpreadsheetApi;
}
