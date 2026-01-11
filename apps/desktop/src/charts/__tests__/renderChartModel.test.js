import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { defaultChartTheme, renderChartToSvg, resolveChartData } from "../renderChart.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readSnapshot(name) {
  return fs.readFileSync(path.join(__dirname, "__snapshots__", name), "utf8");
}

function expectSnapshot(name, received) {
  assert.equal(received, readSnapshot(name));
}

const size = { width: 320, height: 200 };

test("renders clustered bar chart with axes + legend + title", () => {
  const model = {
    chartType: { kind: "bar" },
    title: "Example Chart",
    legend: { show: true, position: "right" },
    axes: { value: { majorGridlines: true, tickCount: 5 } },
    series: [
      {
        name: "Sales",
        categories: { strCache: ["A", "B", "C", "D"] },
        values: { numCache: [2, 4, 3, 5] },
      },
      {
        name: "Budget",
        categories: { strCache: ["A", "B", "C", "D"] },
        values: { numCache: [1, 3, 2, 4] },
      },
    ],
  };

  const data = resolveChartData(model);
  const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
  const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);
  assert.equal(svg1, svg2);

  assert.match(svg1, /<svg/);
  assert.match(svg1, /Example Chart/);
  assert.match(svg1, /<rect /);
  assert.match(svg1, />A</);
  assert.match(svg1, /<line /);
  assert.match(svg1, />Sales</);
  assert.match(svg1, />Budget</);

  expectSnapshot("renderChartModel.bar.svg", svg1);
});

test("renders line chart with polyline series and optional markers", () => {
  const model = {
    chartType: { kind: "line" },
    title: "Example Chart",
    legend: { show: true, position: "right" },
    axes: { value: { tickCount: 5 } },
    options: { markers: true },
    series: [
      {
        name: "Revenue",
        categories: { strCache: ["A", "B", "C", "D"] },
        values: { numCache: [2, 4, 3, 5] },
      },
    ],
  };

  const data = resolveChartData(model);
  const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
  const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);
  assert.equal(svg1, svg2);

  assert.match(svg1, /<polyline /);
  assert.match(svg1, /<circle /);
  assert.match(svg1, />Revenue</);

  expectSnapshot("renderChartModel.line.svg", svg1);
});

test("renders scatter chart with point markers + axes + legend", () => {
  const model = {
    chartType: { kind: "scatter" },
    title: "Example Chart",
    legend: { show: true, position: "right" },
    axes: { x: { tickCount: 4 }, y: { tickCount: 4 } },
    series: [
      {
        name: "Series 1",
        xValues: { numCache: [0, 1, 2, 3] },
        yValues: { numCache: [0, 2, 3, 5] },
      },
    ],
  };

  const data = resolveChartData(model);
  const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
  const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);
  assert.equal(svg1, svg2);

  assert.match(svg1, /<circle /);
  assert.match(svg1, /<line /);
  assert.match(svg1, />Series 1</);

  expectSnapshot("renderChartModel.scatter.svg", svg1);
});

test("renders pie chart with slice paths + legend + title", () => {
  const model = {
    chartType: { kind: "pie" },
    title: "Example Chart",
    legend: { show: true, position: "right" },
    series: [
      {
        categories: { strCache: ["A", "B", "C", "D"] },
        values: { numCache: [2, 4, 3, 5] },
      },
    ],
  };

  const data = resolveChartData(model);
  const svg1 = renderChartToSvg(model, data, defaultChartTheme, size);
  const svg2 = renderChartToSvg(model, data, defaultChartTheme, size);
  assert.equal(svg1, svg2);

  assert.match(svg1, /<path /);
  assert.match(svg1, />A</);
  assert.match(svg1, />D</);

  expectSnapshot("renderChartModel.pie.svg", svg1);
});

