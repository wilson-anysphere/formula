import assert from "node:assert/strict";
import test from "node:test";

import { createMatrixRangeProvider, renderChartSvg } from "../renderSvg.js";

test("renders bar chart svg and updates when source data changes", () => {
  const sheets = {
    Sheet1: [
      ["Category", "Value"],
      ["A", 2],
      ["B", 4],
      ["C", 3],
      ["D", 5],
    ],
  };

  const provider = createMatrixRangeProvider(sheets);
  const chart = {
    chartType: { kind: "bar" },
    title: "Example Chart",
    series: [{ categories: "Sheet1!$A$2:$A$5", values: "Sheet1!$B$2:$B$5" }],
  };

  const svg1 = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.match(svg1, /<svg/);
  assert.match(svg1, /<rect/);

  sheets.Sheet1[2][1] = 10;
  const svg2 = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.notEqual(svg1, svg2);
});

test("renders line chart svg", () => {
  const provider = createMatrixRangeProvider({
    Sheet1: [
      ["Category", "Value"],
      ["A", 2],
      ["B", 4],
      ["C", 3],
      ["D", 5],
    ],
  });

  const chart = {
    chartType: { kind: "line" },
    title: "Example Chart",
    series: [{ categories: "Sheet1!$A$2:$A$5", values: "Sheet1!$B$2:$B$5" }],
  };

  const svg = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.match(svg, /<polyline/);
});

test("renders pie chart svg", () => {
  const provider = createMatrixRangeProvider({
    Sheet1: [
      ["Category", "Value"],
      ["A", 2],
      ["B", 4],
      ["C", 3],
      ["D", 5],
    ],
  });

  const chart = {
    chartType: { kind: "pie" },
    title: "Example Chart",
    series: [{ categories: "Sheet1!$A$2:$A$5", values: "Sheet1!$B$2:$B$5" }],
  };

  const svg = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.match(svg, /<path/);
});

test("renders pie chart svg with fallback labels when categories are missing", () => {
  const provider = createMatrixRangeProvider({
    Sheet1: [
      ["Value"],
      [2],
      [4],
      [3],
      [5],
    ],
  });

  const chart = {
    chartType: { kind: "pie" },
    title: "Example Chart",
    series: [{ values: "Sheet1!$A$2:$A$5" }],
  };

  const svg = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.match(svg, /<path/);
  assert.match(svg, /1, 2, 3, 4/);
});

test("renders scatter chart svg", () => {
  const provider = createMatrixRangeProvider({
    Sheet1: [
      ["X", "Y"],
      [0, 0],
      [1, 2],
      [2, 3],
      [3, 5],
    ],
  });

  const chart = {
    chartType: { kind: "scatter" },
    title: "Example Chart",
    series: [{ xValues: "Sheet1!$A$2:$A$5", yValues: "Sheet1!$B$2:$B$5" }],
  };

  const svg = renderChartSvg(chart, provider, { width: 320, height: 200 });
  assert.match(svg, /<circle/);
});

test("renders placeholder for unsupported chart types", () => {
  const chart = {
    chartType: { kind: "unknown", name: "radarChart" },
    title: "Unsupported",
    series: [],
  };

  const svg = renderChartSvg(chart, { getRange: () => [] }, { width: 320, height: 200 });
  assert.match(svg, /Unsupported chart/);
});

test("uses a provided theme series palette when rendering", () => {
  const provider = createMatrixRangeProvider({
    Sheet1: [
      ["Category", "Value"],
      ["A", 2],
      ["B", 4],
      ["C", 3],
      ["D", 5],
    ],
  });

  const chart = {
    chartType: { kind: "bar" },
    title: "Example Chart",
    series: [{ categories: "Sheet1!$A$2:$A$5", values: "Sheet1!$B$2:$B$5" }],
  };

  const svg = renderChartSvg(chart, provider, {
    width: 320,
    height: 200,
    theme: { seriesColors: ["#FF0000"] },
  });
  assert.match(svg, /fill="#FF0000"/);
});
