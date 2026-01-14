import { describe, expect, it, vi } from "vitest";

import { ChartStore } from "./chartStore";

describe("ChartStore", () => {
  it("creates multi-series bar charts when the data range includes multiple value columns", () => {
    const sheet = [
      ["Category", "Sales", "Profit"],
      ["A", 10, 2],
      ["B", 20, 3],
    ];

    const store = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: (sheetId, row, col) => {
        if (sheetId !== "Sheet1") return null;
        return sheet[row]?.[col] ?? null;
      },
    });

    store.createChart({ chart_type: "bar", data_range: "Sheet1!A1:C3", title: "Totals" });

    const charts = store.listCharts();
    expect(charts).toHaveLength(1);
    const chart = charts[0]!;
    expect(chart.title).toBe("Totals");

    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]).toMatchObject({
      name: "Sales",
      categories: "Sheet1!$A$2:$A$3",
      values: "Sheet1!$B$2:$B$3",
    });
    expect(chart.series[1]).toMatchObject({
      name: "Profit",
      categories: "Sheet1!$A$2:$A$3",
      values: "Sheet1!$C$2:$C$3",
    });
  });

  it("avoids redundant onChange notifications when updating to an identical anchor", () => {
    const onChange = vi.fn();
    const store = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: () => null,
      onChange,
    });

    const { chart_id } = store.createChart({ chart_type: "bar", data_range: "Sheet1!A1:B2", title: "Test" });
    expect(onChange).toHaveBeenCalledTimes(1);

    const initialCharts = store.listCharts();
    const chart = initialCharts.find((c) => c.id === chart_id);
    expect(chart).toBeTruthy();

    store.updateChartAnchor(chart_id, { ...(chart!.anchor as any) });
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(store.listCharts()).toBe(initialCharts);

    store.updateChartAnchor(chart_id, { kind: "absolute", xEmu: 1, yEmu: 2, cxEmu: 3, cyEmu: 4 });
    expect(onChange).toHaveBeenCalledTimes(2);
    expect(store.listCharts()).not.toBe(initialCharts);
  });

  it("reorders charts within a sheet via arrangeChart (without moving other sheets' charts)", () => {
    const onChange = vi.fn();
    const store = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: () => null,
      onChange,
    });

    const { chart_id: chartA } = store.createChart({ chart_type: "bar", data_range: "Sheet1!A1:B2", title: "A" });
    const { chart_id: chartB } = store.createChart({ chart_type: "bar", data_range: "Sheet2!A1:B2", title: "B" });
    const { chart_id: chartC } = store.createChart({ chart_type: "bar", data_range: "Sheet1!A1:B2", title: "C" });

    expect(store.listCharts().map((c) => c.id)).toEqual([chartA, chartB, chartC]);

    // Bring A forward within Sheet1: swap with C while leaving the Sheet2 chart in place.
    expect(store.arrangeChart(chartA, "forward")).toBe(true);
    expect(store.listCharts().map((c) => c.id)).toEqual([chartC, chartB, chartA]);

    // Send A backward within Sheet1: swap back.
    expect(store.arrangeChart(chartA, "backward")).toBe(true);
    expect(store.listCharts().map((c) => c.id)).toEqual([chartA, chartB, chartC]);

    // Bring A to front within Sheet1: move to the end of the Sheet1 subset.
    expect(store.arrangeChart(chartA, "front")).toBe(true);
    expect(store.listCharts().map((c) => c.id)).toEqual([chartC, chartB, chartA]);

    // Sending A to back moves it to the start of the Sheet1 subset.
    expect(store.arrangeChart(chartA, "back")).toBe(true);
    expect(store.listCharts().map((c) => c.id)).toEqual([chartA, chartB, chartC]);

    // Sending A to back again is a no-op (already backmost within Sheet1 subset).
    expect(store.arrangeChart(chartA, "back")).toBe(false);

    // Ensure onChange fired at least once for each successful arrange operation.
    expect(onChange).toHaveBeenCalled();
  });

  it("normalizes position to a string before parsing (invalid non-string positions fail with a helpful error)", () => {
    const store = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: () => null,
    });

    expect(() => store.createChart({ chart_type: "bar", data_range: "Sheet1!A1:B2", position: 123 as any })).toThrow(
      /Invalid position/i,
    );
  });
});
