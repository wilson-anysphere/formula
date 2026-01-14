import { describe, expect, it } from "vitest";

import { ChartCanvasStoreAdapter } from "../chartCanvasStoreAdapter";
import type { ChartRecord } from "../chartStore";

function chartRecord(id: string): ChartRecord {
  return {
    id,
    sheetId: "sheet_1",
    chartType: { kind: "bar" },
    title: "Chart",
    series: [],
    anchor: { kind: "absolute", xEmu: 0, yEmu: 0, cxEmu: 0, cyEmu: 0 },
  };
}

describe("ChartCanvasStoreAdapter.pruneEntries", () => {
  it("drops cached entries not present in the keep set", () => {
    const charts = new Map<string, ChartRecord>();
    const adapter = new ChartCanvasStoreAdapter({
      getChart: (chartId) => charts.get(chartId),
      getCellValue: () => null,
      resolveSheetId: () => "sheet_1",
      getSeriesColors: () => ["#ff0000"],
      maxDataCells: 1000,
    });

    charts.set("c1", chartRecord("c1"));
    charts.set("c2", chartRecord("c2"));

    // Seed the internal cache by building models.
    adapter.getChartModel("c1");
    adapter.getChartModel("c2");
    expect(((adapter as any).entries as Map<string, unknown>).size).toBe(2);

    adapter.pruneEntries(new Set(["c2"]));
    expect(((adapter as any).entries as Map<string, unknown>).size).toBe(1);
    expect(((adapter as any).entries as Map<string, unknown>).has("c2")).toBe(true);
  });

  it("clears all cached entries when keep set is empty", () => {
    const charts = new Map<string, ChartRecord>();
    const adapter = new ChartCanvasStoreAdapter({
      getChart: (chartId) => charts.get(chartId),
      getCellValue: () => null,
      resolveSheetId: () => "sheet_1",
      getSeriesColors: () => [],
      maxDataCells: 1000,
    });

    charts.set("c1", chartRecord("c1"));
    adapter.getChartModel("c1");
    expect(((adapter as any).entries as Map<string, unknown>).size).toBe(1);

    adapter.pruneEntries(new Set());
    expect(((adapter as any).entries as Map<string, unknown>).size).toBe(0);
  });
});

