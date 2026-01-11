import { describe, expect, test } from "vitest";

import { computeChartLayout, DEFAULT_CHART_THEME, type ChartModel, type Rect } from "../layout";

function rectsOverlap(a: Rect, b: Rect): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

function expectNonOverlapping(a: Rect | null, b: Rect | null) {
  if (!a || !b) return;
  expect(rectsOverlap(a, b)).toBe(false);
}

describe("charts/layout", () => {
  test("bar chart layout is deterministic and produces non-overlapping rects", () => {
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Example Chart",
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Value",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
      ],
    };

    const viewport = { x: 0, y: 0, width: 480, height: 320 };

    const layout1 = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);
    const layout2 = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);
    expect(layout1).toEqual(layout2);

    expect(layout1.plotAreaRect.width).toBeGreaterThan(0);
    expect(layout1.plotAreaRect.height).toBeGreaterThan(0);

    expect(layout1.titleRect).not.toBeNull();
    expect(layout1.legendRect).toBeNull();

    expectNonOverlapping(layout1.titleRect, layout1.plotAreaRect);
    expect(layout1.axes.y.ticks.length).toBeGreaterThanOrEqual(5);
    expect(layout1.axes.y.ticks.length).toBeLessThanOrEqual(7);
    expect(layout1.axes.y.gridlines.length).toBeGreaterThan(0);
  });

  test("line chart layout reserves space for a right legend", () => {
    const model: ChartModel = {
      chartType: { kind: "line" },
      title: "Example Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0.0" },
      ],
      series: [
        {
          name: "Series A",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
        {
          name: "Series B",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [1, 3, 2, 4] },
        },
      ],
    };

    const viewport = { x: 0, y: 0, width: 520, height: 320 };
    const layout = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);

    expect(layout.legendRect).not.toBeNull();
    expect(layout.legend?.entries.length).toBe(2);

    expectNonOverlapping(layout.legendRect, layout.plotAreaRect);
    expectNonOverlapping(layout.titleRect, layout.legendRect);
    expect(layout.plotAreaRect.x + layout.plotAreaRect.width).toBeLessThan(layout.legendRect!.x);
  });

  test("scatter chart layout applies explicit axis bounds and reverse order", () => {
    const model: ChartModel = {
      chartType: { kind: "scatter" },
      title: "Example Chart",
      axes: [
        { kind: "value", position: "bottom", scaling: { reverseOrder: true } },
        { kind: "value", position: "left", scaling: { min: 0, max: 10 }, majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Points",
          xValues: { cache: [0, 1, 2, 3] },
          yValues: { cache: [0, 2, 3, 5] },
        },
      ],
    };

    const viewport = { x: 0, y: 0, width: 520, height: 320 };
    const layout = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);

    expect(layout.scales.y.type).toBe("linear");
    if (layout.scales.y.type === "linear") {
      expect(layout.scales.y.domain).toEqual([0, 10]);
    }

    expect(layout.scales.x.type).toBe("linear");
    if (layout.scales.x.type === "linear") {
      expect(layout.scales.x.range[0]).toBeGreaterThan(layout.scales.x.range[1]);
    }

    expect(layout.axes.y.ticks.length).toBeGreaterThanOrEqual(5);
    expect(layout.axes.y.ticks.length).toBeLessThanOrEqual(7);
    expect(layout.axes.y.gridlines.length).toBeGreaterThan(0);
  });
});
