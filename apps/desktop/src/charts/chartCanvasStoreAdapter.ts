import { parseA1Range } from "./a1.js";

import type { ChartStore as ChartRendererStore } from "./chartRendererAdapter";
import type { ChartRecord } from "./chartStore";
import type { ChartModel, ChartTheme as CanvasChartTheme } from "./renderChart";

export interface ChartCanvasStoreAdapterOptions {
  getChart(chartId: string): ChartRecord | undefined;
  /**
   * Read the raw (unformatted) value stored in the given cell.
   */
  getCellValue(sheetId: string, row: number, col: number): unknown;
  /**
   * Resolve a sheet token from an A1 range (either a display name or a stable id) into a stable sheet id.
   */
  resolveSheetId(token: string): string | null;
  /**
   * Series palette (CSS colors).
   */
  getSeriesColors(): string[];
  maxDataCells: number;
}

type Entry = {
  /** Bumps whenever the chart data/model changes (anchor moves should not). */
  revision: number;
  dirty: boolean;
  dataKey: string;
  model: ChartModel | null;
};

function normalizeChartKind(input: unknown): ChartModel["chartType"]["kind"] {
  switch (input) {
    case "bar":
    case "line":
    case "pie":
    case "scatter":
    case "unknown":
      return input;
    // The canvas ChartModel renderer doesn't implement "area" yet; map it to "line"
    // to preserve a usable visualization during the migration.
    case "area":
      return "line";
    default:
      return "unknown";
  }
}

function chartDataKey(chart: ChartRecord): string {
  const seriesKey = (chart.series ?? [])
    .map((s) => [
      s.name ?? "",
      s.categories ?? "",
      s.values ?? "",
      s.xValues ?? "",
      s.yValues ?? "",
    ].join("|"))
    .join(";");
  // Exclude anchors so moving/resizing doesn't force data/model rebuilds.
  return [chart.chartType?.kind ?? "", chart.chartType?.name ?? "", chart.title ?? "", seriesKey].join("::");
}

function rangeCellCount(rangeRef: string): number | null {
  const parsed = parseA1Range(rangeRef);
  if (!parsed) return null;
  const rows = Math.max(0, parsed.endRow - parsed.startRow + 1);
  const cols = Math.max(0, parsed.endCol - parsed.startCol + 1);
  return rows * cols;
}

export class ChartCanvasStoreAdapter implements ChartRendererStore {
  private readonly entries = new Map<string, Entry>();
  private themeRevision = 0;
  private lastSeriesColorsKey = "";

  constructor(private readonly options: ChartCanvasStoreAdapterOptions) {}

  /**
   * Drop cached chart models for ids that are no longer needed.
   *
   * This is used by `SpreadsheetApp` to ensure the adapter does not retain models for charts
   * that were deleted (or otherwise removed from the backing `ChartStore`) after they have
   * fallen out of the render path.
   */
  pruneEntries(keep: ReadonlySet<string>): void {
    if (!keep || keep.size === 0) {
      this.entries.clear();
      return;
    }
    for (const id of this.entries.keys()) {
      if (keep.has(id)) continue;
      this.entries.delete(id);
    }
  }

  invalidate(chartId: string): void {
    const entry = this.entries.get(chartId);
    if (!entry) return;
    entry.dirty = true;
    entry.revision += 1;
  }

  invalidateAll(): void {
    for (const entry of this.entries.values()) {
      entry.dirty = true;
      entry.revision += 1;
    }
  }

  getChartRevision(chartId: string): number {
    this.maybeRefreshThemeRevision();
    let entry = this.entries.get(chartId);
    if (!entry) {
      const chart = this.options.getChart(chartId);
      if (chart) {
        entry = { revision: 1, dirty: true, dataKey: chartDataKey(chart), model: null };
        this.entries.set(chartId, entry);
      }
    }
    // Fold theme changes into the revision so ChartRendererAdapter can treat it as a single cache key.
    const base = entry ? entry.revision : 0;
    return this.themeRevision * 1_000_000 + base;
  }

  getChartTheme(_chartId: string): Partial<CanvasChartTheme> {
    this.maybeRefreshThemeRevision();
    const colors = this.options.getSeriesColors();
    return { seriesColors: colors };
  }

  getChartModel(chartId: string): ChartModel | undefined {
    const chart = this.options.getChart(chartId);
    if (!chart) {
      this.entries.delete(chartId);
      return undefined;
    }

    const key = chartDataKey(chart);
    let entry = this.entries.get(chartId);
    if (!entry) {
      entry = { revision: 1, dirty: true, dataKey: key, model: null };
      this.entries.set(chartId, entry);
    } else if (entry.dataKey !== key) {
      entry.dataKey = key;
      entry.dirty = true;
      entry.revision += 1;
    }

    if (!entry.dirty && entry.model) return entry.model;

    const model = this.buildModel(chart);
    entry.model = model;
    entry.dirty = false;
    return model;
  }

  getChartData(_chartId: string): undefined {
    // We bake live data into the model caches for now, so ChartRendererAdapter can
    // reuse its offscreen surfaces across scroll frames without recomputing ranges.
    return undefined;
  }

  private maybeRefreshThemeRevision(): void {
    const colors = this.options.getSeriesColors();
    const key = Array.isArray(colors) ? colors.join("|") : "";
    if (key !== this.lastSeriesColorsKey) {
      this.lastSeriesColorsKey = key;
      this.themeRevision += 1;
    }
  }

  private buildModel(chart: ChartRecord): ChartModel {
    const kind = normalizeChartKind(chart.chartType?.kind);
    const legend = { position: "right", overlay: false } as const;
    const toCategoryCache = (values: Array<unknown | null>): Array<string | number | null> =>
      values.map((value) => {
        if (value == null) return null;
        if (typeof value === "string" || typeof value === "number") return value;
        if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
        return String(value);
      });

    const toNumberCache = (values: Array<unknown | null>): Array<number | string | null> =>
      values.map((value) => {
        if (value == null) return null;
        if (typeof value === "number" || typeof value === "string") return value;
        if (typeof value === "boolean") return value ? 1 : 0;
        return String(value);
      });

    const axes: ChartModel["axes"] = (() => {
      if (kind === "pie") return null;
      if (kind === "scatter") {
        return [
          { kind: "value", position: "bottom", formatCode: "0" },
          { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
        ];
      }
      return [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ];
    })();

    const tooLarge = (chart.series ?? []).some((ser) => {
      const refs = [ser.categories, ser.values, ser.xValues, ser.yValues];
      for (const ref of refs) {
        if (typeof ref !== "string" || ref.trim() === "") continue;
        const count = rangeCellCount(ref);
        if (typeof count === "number" && count > this.options.maxDataCells) return true;
      }
      return false;
    });

    if (tooLarge) {
      return {
        chartType: { kind, ...(chart.chartType?.name ? { name: chart.chartType.name } : {}) },
        title: chart.title ?? null,
        legend,
        axes,
        series: [],
        options: {
          placeholder: `Chart range too large (>${this.options.maxDataCells.toLocaleString()} cells)`,
        },
      };
    }

    const series = (chart.series ?? []).map((ser) => ({
      ...(ser.name != null ? { name: ser.name } : {}),
      ...(ser.categories
        ? { categories: { cache: toCategoryCache(this.readRangeValues(chart.sheetId, ser.categories)) } }
        : {}),
      ...(ser.values ? { values: { cache: toNumberCache(this.readRangeValues(chart.sheetId, ser.values)) } } : {}),
      ...(ser.xValues ? { xValues: { cache: toNumberCache(this.readRangeValues(chart.sheetId, ser.xValues)) } } : {}),
      ...(ser.yValues ? { yValues: { cache: toNumberCache(this.readRangeValues(chart.sheetId, ser.yValues)) } } : {}),
    }));

    return {
      chartType: { kind, ...(chart.chartType?.name ? { name: chart.chartType.name } : {}) },
      title: chart.title ?? null,
      legend,
      axes,
      series,
    };
  }

  private readRangeValues(fallbackSheetId: string, rangeRef: string): Array<unknown | null> {
    const parsed = parseA1Range(rangeRef);
    if (!parsed) return [];

    const sheetId = parsed.sheetName ? this.options.resolveSheetId(parsed.sheetName) : fallbackSheetId;
    if (!sheetId) return [];

    const rows = Math.max(0, parsed.endRow - parsed.startRow + 1);
    const cols = Math.max(0, parsed.endCol - parsed.startCol + 1);
    if (rows * cols > this.options.maxDataCells) return [];

    const out: Array<unknown | null> = [];
    for (let r = parsed.startRow; r <= parsed.endRow; r += 1) {
      for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
        const value = this.options.getCellValue(sheetId, r, c);
        out.push(value == null ? null : value);
      }
    }
    return out;
  }
}
