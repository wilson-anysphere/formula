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
});
