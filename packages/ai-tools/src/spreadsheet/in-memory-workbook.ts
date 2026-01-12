import { DEFAULT_SHEET_NAME } from "./a1.ts";
import type { CellAddress, RangeAddress } from "./a1.ts";
import type { CellEntry, CreateChartResult, CreateChartSpec, SpreadsheetApi } from "./api.ts";
import { cloneCell, isCellEmpty, type CellData, type CellFormat } from "./types.ts";

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function parseCellKey(key: string): { row: number; col: number } {
  const [rowRaw, colRaw] = key.split(":");
  return { row: Number(rowRaw), col: Number(colRaw) };
}

class InMemorySheet {
  readonly name: string;
  private readonly cells = new Map<string, CellData>();

  constructor(name: string) {
    this.name = name;
  }

  clone(): InMemorySheet {
    const next = new InMemorySheet(this.name);
    for (const [key, cell] of this.cells.entries()) {
      next.cells.set(key, cloneCell(cell));
    }
    return next;
  }

  getCell(row: number, col: number): CellData {
    return this.cells.get(cellKey(row, col)) ?? { value: null };
  }

  setCell(row: number, col: number, cell: CellData): void {
    const normalized: CellData = {
      value: cell.value ?? null,
      formula: cell.formula || undefined,
      format: cell.format && Object.keys(cell.format).length > 0 ? { ...cell.format } : undefined
    };

    const key = cellKey(row, col);
    if (isCellEmpty(normalized)) {
      this.cells.delete(key);
      return;
    }

    this.cells.set(key, normalized);
  }

  listNonEmptyCells(): Array<{ row: number; col: number; cell: CellData }> {
    const entries: Array<{ row: number; col: number; cell: CellData }> = [];
    for (const [key, cell] of this.cells.entries()) {
      const { row, col } = parseCellKey(key);
      entries.push({ row, col, cell: cloneCell(cell) });
    }
    return entries;
  }

  getLastUsedRow(): number {
    let maxRow = 0;
    for (const key of this.cells.keys()) {
      const { row } = parseCellKey(key);
      if (row > maxRow) maxRow = row;
    }
    return maxRow;
  }
}

export class InMemoryWorkbook implements SpreadsheetApi {
  private readonly sheets = new Map<string, InMemorySheet>();
  private nextChartId = 1;
  private charts: Array<{ chart_id: string; spec: CreateChartSpec }> = [];

  constructor(sheetNames: string[] = [DEFAULT_SHEET_NAME]) {
    for (const name of sheetNames) {
      this.sheets.set(name, new InMemorySheet(name));
    }
    if (this.sheets.size === 0) {
      this.sheets.set(DEFAULT_SHEET_NAME, new InMemorySheet(DEFAULT_SHEET_NAME));
    }
  }

  clone(): InMemoryWorkbook {
    const next = new InMemoryWorkbook([]);
    next.sheets.clear();
    for (const [name, sheet] of this.sheets.entries()) {
      next.sheets.set(name, sheet.clone());
    }
    next.nextChartId = this.nextChartId;
    next.charts = this.charts.map((chart) => ({ chart_id: chart.chart_id, spec: { ...chart.spec } }));
    return next;
  }

  listSheets(): string[] {
    return [...this.sheets.keys()];
  }

  listNonEmptyCells(sheet?: string): CellEntry[] {
    const sheets = sheet ? [this.getOrCreateSheet(sheet)] : [...this.sheets.values()];
    const entries: CellEntry[] = [];
    for (const currentSheet of sheets) {
      for (const { row, col, cell } of currentSheet.listNonEmptyCells()) {
        entries.push({ address: { sheet: currentSheet.name, row, col }, cell });
      }
    }
    return entries;
  }

  getCell(address: CellAddress): CellData {
    return cloneCell(this.getOrCreateSheet(address.sheet).getCell(address.row, address.col));
  }

  setCell(address: CellAddress, cell: CellData): void {
    this.getOrCreateSheet(address.sheet).setCell(address.row, address.col, cell);
  }

  readRange(range: RangeAddress): CellData[][] {
    const sheet = this.getOrCreateSheet(range.sheet);
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellData[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(cloneCell(sheet.getCell(r, c)));
      }
      rows.push(row);
    }
    return rows;
  }

  writeRange(range: RangeAddress, cells: CellData[][]): void {
    const rowCount = range.endRow - range.startRow + 1;
    const colCount = range.endCol - range.startCol + 1;
    if (cells.length !== rowCount) {
      throw new Error(
        `writeRange expected ${rowCount} rows but got ${cells.length} rows for ${range.sheet}!R${range.startRow}C${range.startCol}:R${range.endRow}C${range.endCol}`
      );
    }

    for (const row of cells) {
      if (row.length !== colCount) {
        throw new Error(
          `writeRange expected ${colCount} columns but got ${row.length} columns for ${range.sheet}!R${range.startRow}C${range.startCol}:R${range.endRow}C${range.endCol}`
        );
      }
    }

    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        this.setCell(
          { sheet: range.sheet, row: range.startRow + r, col: range.startCol + c },
          cells[r][c] ?? { value: null }
        );
      }
    }
  }

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number {
    let count = 0;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const address = { sheet: range.sheet, row: r, col: c };
        const existing = this.getCell(address);
        const next: CellData = {
          value: existing.value,
          formula: existing.formula,
          format: { ...(existing.format ?? {}), ...format }
        };
        this.setCell(address, next);
        count++;
      }
    }
    return count;
  }

  /**
   * In-memory implementation for `create_chart` tool tests and previews.
   * The workbook does not attempt to render charts; it only records the specs.
   */
  createChart(spec: CreateChartSpec): CreateChartResult {
    const chart_id = `chart_${this.nextChartId++}`;
    this.charts.push({ chart_id, spec: { ...spec } });
    return { chart_id };
  }

  listCharts(): Array<{ chart_id: string; spec: CreateChartSpec }> {
    return this.charts.map((chart) => ({ chart_id: chart.chart_id, spec: { ...chart.spec } }));
  }

  getLastUsedRow(sheet: string): number {
    return this.getOrCreateSheet(sheet).getLastUsedRow();
  }

  private getOrCreateSheet(name: string): InMemorySheet {
    const existing = this.sheets.get(name);
    if (existing) return existing;
    const sheet = new InMemorySheet(name);
    this.sheets.set(name, sheet);
    return sheet;
  }
}
