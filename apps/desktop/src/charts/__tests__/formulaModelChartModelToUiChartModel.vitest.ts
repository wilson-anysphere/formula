import { describe, expect, it } from "vitest";

import { formulaModelChartModelToUiChartModel } from "../formulaModelChartModel";

describe("formulaModelChartModelToUiChartModel", () => {
  it("converts a bar chart model with title, legend, axes, and cached series", () => {
    const input = {
      chartKind: { kind: "bar" },
      title: { richText: { text: "My Bar Chart", runs: [] }, formula: null, style: null },
      legend: { position: "right", overlay: false },
      plotArea: { kind: "bar", barDirection: null, grouping: null, axIds: [1, 2] },
      axes: [
        {
          id: 1,
          kind: "category",
          position: "bottom",
          scaling: { min: null, max: null, logBase: null, reverse: false },
          numFmt: null,
          tickLabelPosition: null,
          majorGridlines: false,
        },
        {
          id: 2,
          kind: "value",
          position: "left",
          scaling: { min: 0, max: 10, logBase: null, reverse: true },
          numFmt: { formatCode: "0", sourceLinked: null },
          tickLabelPosition: null,
          majorGridlines: true,
        },
      ],
      series: [
        {
          name: { richText: { text: "Sales", runs: [] }, formula: null, style: null },
          categories: { formula: null, cache: ["A", "B"] },
          values: { formula: null, cache: [2, 4], formatCode: null },
          xValues: null,
          yValues: null,
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.chartType).toEqual({ kind: "bar" });
    expect(out.title).toBe("My Bar Chart");
    expect(out.legend).toEqual({ position: "right", overlay: false });
    expect(out.axes?.[0]).toMatchObject({ id: "1", kind: "category", position: "bottom" });
    expect(out.axes?.[1]).toMatchObject({
      id: "2",
      kind: "value",
      position: "left",
      majorGridlines: true,
      formatCode: "0",
      scaling: { min: 0, max: 10, reverseOrder: true },
    });
    expect(out.series).toEqual([
      {
        name: "Sales",
        categories: { cache: ["A", "B"] },
        values: { cache: [2, 4] },
      },
    ]);
  });

  it("converts a line chart model", () => {
    const input = {
      chartKind: { kind: "line" },
      title: { richText: { text: "My Line Chart", runs: [] }, formula: null, style: null },
      legend: { position: "topRight", overlay: true },
      plotArea: { kind: "line", grouping: null, axIds: [1, 2] },
      axes: [],
      series: [
        {
          name: { richText: { text: "Revenue", runs: [] }, formula: null, style: null },
          categories: { formula: null, cache: ["Q1", "Q2", "Q3"] },
          values: { formula: null, cache: [10, 20, 30], formatCode: null },
          xValues: null,
          yValues: null,
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.chartType).toEqual({ kind: "line" });
    expect(out.title).toBe("My Line Chart");
    // topRight is treated as right for now.
    expect(out.legend).toEqual({ position: "right", overlay: true });
    expect(out.series[0]).toMatchObject({
      name: "Revenue",
      categories: { cache: ["Q1", "Q2", "Q3"] },
      values: { cache: [10, 20, 30] },
    });
  });

  it("uses numeric categories when provided as categoriesNum", () => {
    const input = {
      chartKind: { kind: "line" },
      title: { richText: { text: "My Line Chart", runs: [] }, formula: null, style: null },
      legend: { position: "right", overlay: false },
      plotArea: { kind: "line", grouping: null, axIds: [1, 2] },
      axes: [],
      series: [
        {
          name: { richText: { text: "Revenue", runs: [] }, formula: null, style: null },
          categories: null,
          categoriesNum: { formula: null, cache: [45123, 45124], formatCode: null },
          values: { formula: null, cache: [10, 20], formatCode: null },
          xValues: null,
          yValues: null,
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.series[0]).toMatchObject({
      name: "Revenue",
      categories: { cache: [45123, 45124] },
      values: { cache: [10, 20] },
    });
  });

  it("converts a pie chart model (categories + values)", () => {
    const input = {
      chartKind: { kind: "pie" },
      title: { richText: { text: "My Pie Chart", runs: [] }, formula: null, style: null },
      legend: { position: "right", overlay: false },
      plotArea: { kind: "pie", varyColors: true, firstSliceAngle: null },
      axes: [],
      series: [
        {
          name: null,
          categories: { formula: null, cache: ["A", "B", "C"] },
          values: { formula: null, cache: [1, 2, 3], formatCode: null },
          xValues: null,
          yValues: null,
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.chartType).toEqual({ kind: "pie" });
    expect(out.title).toBe("My Pie Chart");
    expect(out.series).toEqual([
      {
        categories: { cache: ["A", "B", "C"] },
        values: { cache: [1, 2, 3] },
      },
    ]);
  });

  it("converts a scatter chart model with x/y caches", () => {
    const input = {
      chartKind: { kind: "scatter" },
      title: { richText: { text: "My Scatter Chart", runs: [] }, formula: null, style: null },
      legend: { position: "right", overlay: false },
      plotArea: { kind: "scatter", scatterStyle: null, axIds: [1, 2] },
      axes: [],
      series: [
        {
          name: { richText: { text: "Series 1", runs: [] }, formula: null, style: null },
          categories: null,
          values: null,
          xValues: { kind: "number", formula: null, cache: [0, 1, 2], formatCode: null },
          yValues: { kind: "number", formula: null, cache: [10, 11, 12], formatCode: null },
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.chartType).toEqual({ kind: "scatter" });
    expect(out.series).toEqual([
      {
        name: "Series 1",
        xValues: { cache: [0, 1, 2] },
        yValues: { cache: [10, 11, 12] },
      },
    ]);
  });

  it("trims whitespace around series formula refs", () => {
    const input = {
      chartKind: { kind: "line" },
      title: null,
      legend: null,
      plotArea: { kind: "line", grouping: null, axIds: [] },
      axes: [],
      series: [
        {
          name: null,
          categories: { formula: "  Sheet1!A1:A3  " },
          values: { formula: "  Sheet1!B1:B3  ", formatCode: null },
          xValues: null,
          yValues: null,
        },
      ],
      diagnostics: [],
    };

    const out = formulaModelChartModelToUiChartModel(input);
    expect(out.series).toEqual([
      {
        categories: { ref: "Sheet1!A1:A3" },
        values: { ref: "Sheet1!B1:B3" },
      },
    ]);
  });
});
