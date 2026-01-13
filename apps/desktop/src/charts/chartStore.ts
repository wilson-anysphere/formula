import { parseA1Range } from "./a1.js";

import type { ChartType, CreateChartResult, CreateChartSpec } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import type { SheetNameResolver } from "../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";

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
  /**
   * Sheet the chart is anchored on (i.e. where it should be rendered).
   */
  sheetId: string;
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
  /**
   * Optional sheet display-name resolver.
   *
   * When provided, sheet-qualified range strings like `Budget!A1:B2` are resolved to
   * stable sheet ids before reading values.
   */
  sheetNameResolver?: SheetNameResolver | null;
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

function formatAbsCellRef(row: number, col: number): string {
  return `$${columnIndexToLetters(col)}$${row + 1}`;
}

function formatAbsRange(sheet: string, startRow: number, startCol: number, endRow: number, endCol: number): string {
  const sheetPrefix = formatSheetNameForA1(sheet);
  const start = formatAbsCellRef(startRow, startCol);
  const end = formatAbsCellRef(endRow, endCol);
  const body = start === end ? start : `${start}:${end}`;
  return `${sheetPrefix}!${body}`;
}

function getTextLike(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (value && typeof value === "object") {
    const maybe = value as { text?: unknown };
    if (typeof maybe.text === "string") return maybe.text;
  }
  return null;
}

function isMostlyStrings(row: unknown[]): boolean {
  const nonEmpty = row.filter((value) => value != null && value !== "");
  if (nonEmpty.length === 0) return false;
  const stringCount = nonEmpty.filter((value) => typeof getTextLike(value) === "string").length;
  return stringCount / nonEmpty.length >= 0.6;
}

export class ChartStore {
  private readonly options: ChartStoreOptions;
  private charts: ChartRecord[] = [];
  private nextId = 1;

  constructor(options: ChartStoreOptions) {
    this.options = options;
  }

  private resolveSheetIdFromToken(sheetToken: string): string | null {
    const trimmed = String(sheetToken ?? "").trim();
    if (!trimmed) return null;
    const resolver = this.options.sheetNameResolver ?? null;
    if (!resolver) return trimmed;

    const byName = resolver.getSheetIdByName(trimmed);
    if (byName) return byName;

    // Back-compat: allow stored chart ranges to continue using stable ids even after
    // the sheet has been renamed (id != name).
    const byId = resolver.getSheetNameById(trimmed);
    if (byId) return trimmed;

    return null;
  }

  listCharts(): readonly ChartRecord[] {
    return this.charts;
  }

  updateChartAnchor(chartId: string, anchor: ChartAnchor): void {
    const id = String(chartId ?? "");
    if (!id) return;

    let changed = false;
    const next = this.charts.map((chart) => {
      if (chart.id !== id) return chart;
      changed = true;
      return {
        ...chart,
        anchor: { ...(anchor as any) },
      };
    });
    if (!changed) return;
    this.charts = next;
    this.options.onChange?.();
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

    const dataSheetId = (() => {
      if (!parsed.sheetName) return this.options.defaultSheet;
      const resolved = this.resolveSheetIdFromToken(parsed.sheetName);
      if (!resolved) throw new Error(`Unknown sheet: ${parsed.sheetName}`);
      return resolved;
    })();
    const rowCount = parsed.endRow - parsed.startRow + 1;
    const colCount = parsed.endCol - parsed.startCol + 1;

    const headerRow: unknown[] = [];
    for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
      headerRow.push(this.options.getCellValue(dataSheetId, parsed.startRow, c));
    }

    const hasHeader = rowCount > 1 && isMostlyStrings(headerRow);
    const dataStartRow = hasHeader ? parsed.startRow + 1 : parsed.startRow;
    const seriesNameForOffset = (offset: number): string | undefined => {
      if (!hasHeader) return undefined;
      const raw = headerRow[offset];
      if (raw == null) return undefined;
      const text = getTextLike(raw);
      if (text == null) return undefined;
      const name = text.trim();
      return name ? name : undefined;
    };

    const series: ChartSeriesDef[] = [];
    if (spec.chart_type === "scatter") {
      if (colCount < 2) {
        throw new Error(`scatter chart requires at least 2 columns (got ${colCount})`);
      }

      series.push({
        ...(seriesNameForOffset(1) ? { name: seriesNameForOffset(1) } : {}),
        xValues: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol),
        yValues: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol + 1, parsed.endRow, parsed.startCol + 1)
      });
    } else if (spec.chart_type === "pie") {
      if (colCount >= 2) {
        series.push({
          ...(seriesNameForOffset(1) ? { name: seriesNameForOffset(1) } : {}),
          categories: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol),
          values: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol + 1, parsed.endRow, parsed.startCol + 1)
        });
      } else {
        series.push({
          ...(seriesNameForOffset(0) ? { name: seriesNameForOffset(0) } : {}),
          values: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol)
        });
      }
    } else {
      if (colCount >= 2) {
        const categories = formatAbsRange(dataSheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol);
        for (let offset = 1; offset < colCount; offset += 1) {
          const col = parsed.startCol + offset;
          series.push({
            ...(seriesNameForOffset(offset) ? { name: seriesNameForOffset(offset) } : {}),
            categories,
            values: formatAbsRange(dataSheetId, dataStartRow, col, parsed.endRow, col)
          });
        }
      } else {
        series.push({
          ...(seriesNameForOffset(0) ? { name: seriesNameForOffset(0) } : {}),
          values: formatAbsRange(dataSheetId, dataStartRow, parsed.startCol, parsed.endRow, parsed.startCol)
        });
      }
    }

    const positionParsed =
      spec.position && String(spec.position).trim() !== ""
        ? (() => {
            const parsedPosition = parseA1Range(spec.position);
            if (!parsedPosition) throw new Error(`Invalid position: ${spec.position}`);
            return parsedPosition;
          })()
        : null;

    const sheetId = (() => {
      if (!positionParsed?.sheetName) return dataSheetId;
      const resolved = this.resolveSheetIdFromToken(positionParsed.sheetName);
      if (!resolved) throw new Error(`Unknown sheet: ${positionParsed.sheetName}`);
      return resolved;
    })();
    const anchor = this.resolveAnchor(parsed, positionParsed);
    const id = `chart_${this.nextId++}`;

    const chart: ChartRecord = {
      id,
      sheetId,
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

  private resolveAnchor(
    dataRange: NonNullable<ReturnType<typeof parseA1Range>>,
    position: NonNullable<ReturnType<typeof parseA1Range>> | null
  ): ChartAnchor {
    if (position) {
      const fromCol = position.startCol;
      const fromRow = position.startRow;

      // If a rectangular range is provided, use it as the chart bounds.
      const rangeCols = position.endCol - position.startCol + 1;
      const rangeRows = position.endRow - position.startRow + 1;
      const toCol = rangeCols > 1 ? position.endCol + 1 : fromCol + DEFAULT_ANCHOR_WIDTH_COLS;
      const toRow = rangeRows > 1 ? position.endRow + 1 : fromRow + DEFAULT_ANCHOR_HEIGHT_ROWS;

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
