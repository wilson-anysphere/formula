import { parseA1Range } from "./a1.js";

import type { ChartType, CreateChartResult, CreateChartSpec } from "../../../../packages/ai-tools/src/spreadsheet/api.js";

export type ChartSeriesDef = {
  name?: string | null;
  categories?: string | null;
  values?: string | null;
  xValues?: string | null;
  yValues?: string | null;
};

export type ChartAnchor =
  | {
      kind: "twoCell";
      fromCol: number;
      fromRow: number;
      fromColOffEmu: number;
      fromRowOffEmu: number;
      toCol: number;
      toRow: number;
      toColOffEmu: number;
      toRowOffEmu: number;
    }
  | {
      kind: "oneCell";
      fromCol: number;
      fromRow: number;
      fromColOffEmu: number;
      fromRowOffEmu: number;
      cxEmu: number;
      cyEmu: number;
    }
  | {
      kind: "absolute";
      xEmu: number;
      yEmu: number;
      cxEmu: number;
      cyEmu: number;
    };

export type ChartDef = {
  chartType: { kind: ChartType; name?: string };
  title?: string;
  series: ChartSeriesDef[];
  anchor: ChartAnchor;
};

export type ChartRecord = ChartDef & { id: string };

export interface ChartStoreOptions {
  defaultSheet: string;
  /**
   * Reads the raw value stored in a cell at 0-based row/col coordinates.
   */
  getCellValue: (sheetId: string, row: number, col: number) => unknown;
  onChange?: () => void;
}

const DEFAULT_ANCHOR_GAP_COLS = 2;
const DEFAULT_ANCHOR_WIDTH_COLS = 5;
const DEFAULT_ANCHOR_HEIGHT_ROWS = 11;

function columnIndexToLetters(col: number): string {
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function quoteSheetName(sheet: string): string {
  // Follow Excel conventions: quote when the name contains spaces/special chars.
  if (/^[A-Za-z0-9_]+$/.test(sheet)) return sheet;
  return `'${sheet.replace(/'/g, "''")}'`;
}

function formatAbsCellRef(row: number, col: number): string {
  return `$${columnIndexToLetters(col)}$${row + 1}`;
}

function formatAbsRange(sheet: string, startRow: number, startCol: number, endRow: number, endCol: number): string {
  const sheetPrefix = quoteSheetName(sheet);
  const start = formatAbsCellRef(startRow, startCol);
  const end = formatAbsCellRef(endRow, endCol);
  const body = start === end ? start : `${start}:${end}`;
  return `${sheetPrefix}!${body}`;
}

function isMostlyStrings(row: unknown[]): boolean {
  const nonEmpty = row.filter((value) => value != null && value !== "");
  if (nonEmpty.length === 0) return false;
  const stringCount = nonEmpty.filter((value) => typeof value === "string").length;
  return stringCount / nonEmpty.length >= 0.6;
}

export class ChartStore {
  private readonly options: ChartStoreOptions;
  private charts: ChartRecord[] = [];
  private nextId = 1;

  constructor(options: ChartStoreOptions) {
    this.options = options;
  }

  listCharts(): readonly ChartRecord[] {
    return this.charts;
  }

  setDefaultSheet(sheetId: string): void {
    this.options.defaultSheet = sheetId;
  }

  /**
   * Create a chart record compatible with `renderChartSvg` and store it in-memory.
   */
  createChart(spec: CreateChartSpec): CreateChartResult {
    const parsed = parseA1Range(spec.data_range);
    if (!parsed) {
      throw new Error(`Invalid data_range: ${spec.data_range}`);
    }

    const sheetId = parsed.sheetName ?? this.options.defaultSheet;
    const rowCount = parsed.endRow - parsed.startRow + 1;
    const colCount = parsed.endCol - parsed.startCol + 1;

    const headerRow: unknown[] = [];
    for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
      headerRow.push(this.options.getCellValue(sheetId, parsed.startRow, c));
    }

    const hasHeader = rowCount > 1 && isMostlyStrings(headerRow);
    const dataStartRow = hasHeader ? parsed.startRow + 1 : parsed.startRow;
    const seriesName =
      hasHeader && colCount >= 2 && headerRow[1] != null && headerRow[1] !== "" ? String(headerRow[1]) : undefined;

    const series: ChartSeriesDef[] = [];
    if (spec.chart_type === "scatter") {
      if (colCount < 2) {
        throw new Error(`scatter chart requires at least 2 columns (got ${colCount})`);
      }

      series.push({
        ...(seriesName ? { name: seriesName } : {}),
        xValues: formatAbsRange(sheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol),
        yValues: formatAbsRange(sheetId, dataStartRow, parsed.startCol + 1, parsed.endRow, parsed.startCol + 1)
      });
    } else {
      if (colCount >= 2) {
        series.push({
          ...(seriesName ? { name: seriesName } : {}),
          categories: formatAbsRange(sheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol),
          values: formatAbsRange(sheetId, dataStartRow, parsed.startCol + 1, parsed.endRow, parsed.startCol + 1)
        });
      } else {
        series.push({
          ...(seriesName ? { name: seriesName } : {}),
          values: formatAbsRange(sheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol)
        });
      }
    }

    const anchor = this.resolveAnchor(spec, parsed);
    const id = `chart_${this.nextId++}`;

    const chart: ChartRecord = {
      id,
      chartType: { kind: spec.chart_type },
      ...(spec.title ? { title: spec.title } : {}),
      series,
      anchor
    };

    this.charts = [...this.charts, chart];
    this.options.onChange?.();
    return { chart_id: id };
  }

  clone(options: Omit<ChartStoreOptions, "onChange">): ChartStore {
    const cloned = new ChartStore({ ...options });
    cloned.nextId = this.nextId;
    cloned.charts = this.charts.map((chart) => ({
      ...chart,
      chartType: { ...chart.chartType },
      series: chart.series.map((ser) => ({ ...ser })),
      anchor: { ...(chart.anchor as any) }
    }));
    return cloned;
  }

  private resolveAnchor(spec: CreateChartSpec, dataRange: NonNullable<ReturnType<typeof parseA1Range>>): ChartAnchor {
    if (spec.position && String(spec.position).trim() !== "") {
      const parsed = parseA1Range(spec.position);
      if (!parsed) {
        throw new Error(`Invalid position: ${spec.position}`);
      }

      const fromCol = parsed.startCol;
      const fromRow = parsed.startRow;

      // If a rectangular range is provided, use it as the chart bounds.
      const rangeCols = parsed.endCol - parsed.startCol + 1;
      const rangeRows = parsed.endRow - parsed.startRow + 1;
      const toCol = rangeCols > 1 ? parsed.endCol + 1 : fromCol + DEFAULT_ANCHOR_WIDTH_COLS;
      const toRow = rangeRows > 1 ? parsed.endRow + 1 : fromRow + DEFAULT_ANCHOR_HEIGHT_ROWS;

      return {
        kind: "twoCell",
        fromCol,
        fromRow,
        fromColOffEmu: 0,
        fromRowOffEmu: 0,
        toCol,
        toRow,
        toColOffEmu: 0,
        toRowOffEmu: 0
      };
    }

    const fromCol = dataRange.endCol + 1 + DEFAULT_ANCHOR_GAP_COLS;
    const fromRow = dataRange.startRow;
    return {
      kind: "twoCell",
      fromCol,
      fromRow,
      fromColOffEmu: 0,
      fromRowOffEmu: 0,
      toCol: fromCol + DEFAULT_ANCHOR_WIDTH_COLS,
      toRow: fromRow + DEFAULT_ANCHOR_HEIGHT_ROWS,
      toColOffEmu: 0,
      toRowOffEmu: 0
    };
  }
}
