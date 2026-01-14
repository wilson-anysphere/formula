import { describe, expect, it } from "vitest";

import { defaultChartTheme, renderChartToSvg, resolveChartData, type ChartModel } from "../renderChart";

const size = { width: 320, height: 200 };

describe("charts/renderChart (ChartModel renderer)", () => {
  it("uses numeric categories from categoriesNum when categories are missing", () => {
    const model: any = {
      chartType: { kind: "line" },
      series: [
        {
          categories: null,
          categoriesNum: { cache: [45123, 45124] },
          values: { cache: [2, 4] },
        },
      ],
    };

    const data = resolveChartData(model as ChartModel);
    expect(data.series[0]?.categories).toEqual(["45123", "45124"]);
  });

  it("renders clustered bar chart with axes, gridlines, legend, and title", () => {
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Example Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Sales",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
        {
          name: "Budget",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [1, 3, 2, 4] },
        },
      ],
    };

    const data = resolveChartData(model);
    const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
    const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);

    expect(svg1).toEqual(svg2);
    expect(svg1).toContain("<svg");
    expect(svg1).toContain("Example Chart");
    expect(svg1).toContain("<rect");
    expect(svg1).toContain("<line");
    expect(svg1).toContain("stroke-dasharray"); // major gridlines
    expect(svg1).toContain("Sales");
    expect(svg1).toContain("Budget");
    expect(svg1).toMatchSnapshot();
  });

  it("renders line chart as polylines with optional circle markers", () => {
    const model: ChartModel = {
      chartType: { kind: "line" },
      title: "Example Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      options: { markers: true },
      series: [
        {
          name: "Revenue",
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
      ],
    };

    const data = resolveChartData(model);
    const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
    const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);

    expect(svg1).toEqual(svg2);
    expect(svg1).toContain("<polyline");
    expect(svg1).toContain("<circle");
    expect(svg1).toContain("Revenue");
    expect(svg1).toMatchSnapshot();
  });

  it("renders scatter chart with circles and axes", () => {
    const model: ChartModel = {
      chartType: { kind: "scatter" },
      title: "Example Chart",
      legend: { position: "right", overlay: false },
      axes: [
        { kind: "value", position: "bottom", formatCode: "0" },
        { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
      ],
      series: [
        {
          name: "Series 1",
          xValues: { cache: [0, 1, 2, 3] },
          yValues: { cache: [0, 2, 3, 5] },
        },
      ],
    };

    const data = resolveChartData(model);
    const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
    const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);

    expect(svg1).toEqual(svg2);
    expect(svg1).toContain("<circle");
    expect(svg1).toContain("<line");
    expect(svg1).toContain("Series 1");
    expect(svg1).toMatchSnapshot();
  });

  it("renders pie chart with slice paths and a category legend", () => {
    const model: ChartModel = {
      chartType: { kind: "pie" },
      title: "Example Chart",
      legend: { position: "right", overlay: false },
      series: [
        {
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
      ],
    };

    const data = resolveChartData(model);
    const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
    const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);

    expect(svg1).toEqual(svg2);
    expect(svg1).toContain("<path");
    expect(svg1).toContain("A");
    expect(svg1).toContain("D");
    expect(svg1).toMatchSnapshot();
  });

  it("trims placeholder overrides when rendering empty chart placeholders", () => {
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Empty Chart",
      series: [],
      options: { placeholder: "  Custom placeholder  " },
    };
    const data = resolveChartData(model);
    const svg = renderChartToSvg(model, data, defaultChartTheme, size);
    expect(svg).toContain(">Custom placeholder<");
  });

  it("includes chartType.name in the placeholder label for unsupported chart kinds", () => {
    const model: ChartModel = {
      chartType: { kind: "unknown", name: "radar" },
      title: "Imported Chart",
      series: [],
    };
    const data = resolveChartData(model);
    const svg = renderChartToSvg(model, data, defaultChartTheme, size);
    expect(svg).toContain("Unsupported chart: radar");
  });
});
