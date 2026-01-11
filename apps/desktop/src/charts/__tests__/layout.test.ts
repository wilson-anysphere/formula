import { describe, expect, test } from "vitest";

import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";

import { computeChartLayout, DEFAULT_CHART_THEME, type ChartModel, type Rect } from "../layout";

function rectsOverlap(a: Rect, b: Rect): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

function expectNonOverlapping(a: Rect | null, b: Rect | null) {
  if (!a || !b) return;
  expect(rectsOverlap(a, b)).toBe(false);
}

type ZipEntry = {
  name: string;
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
};

function findEocd(buf: Buffer): number {
  // EOCD record is at least 22 bytes; the comment can make it larger.
  for (let i = buf.length - 22; i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) return i;
  }
  throw new Error("zip: end of central directory not found");
}

function parseZipEntries(buf: Buffer): Map<string, ZipEntry> {
  const eocd = findEocd(buf);
  const totalEntries = buf.readUInt16LE(eocd + 10);
  const centralDirOffset = buf.readUInt32LE(eocd + 16);

  const out = new Map<string, ZipEntry>();
  let offset = centralDirOffset;

  for (let i = 0; i < totalEntries; i += 1) {
    if (buf.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error(`zip: invalid central directory signature at ${offset}`);
    }

    const compressionMethod = buf.readUInt16LE(offset + 10);
    const compressedSize = buf.readUInt32LE(offset + 20);
    const uncompressedSize = buf.readUInt32LE(offset + 24);
    const nameLen = buf.readUInt16LE(offset + 28);
    const extraLen = buf.readUInt16LE(offset + 30);
    const commentLen = buf.readUInt16LE(offset + 32);
    const localHeaderOffset = buf.readUInt32LE(offset + 42);
    const name = buf.toString("utf8", offset + 46, offset + 46 + nameLen);

    out.set(name, { name, compressionMethod, compressedSize, uncompressedSize, localHeaderOffset });
    offset += 46 + nameLen + extraLen + commentLen;
  }

  return out;
}

function readZipEntry(buf: Buffer, entry: ZipEntry): Buffer {
  const offset = entry.localHeaderOffset;
  if (buf.readUInt32LE(offset) !== 0x04034b50) {
    throw new Error(`zip: invalid local header signature at ${offset}`);
  }

  const nameLen = buf.readUInt16LE(offset + 26);
  const extraLen = buf.readUInt16LE(offset + 28);
  const dataOffset = offset + 30 + nameLen + extraLen;
  const compressed = buf.subarray(dataOffset, dataOffset + entry.compressedSize);

  if (entry.compressionMethod === 0) return Buffer.from(compressed);
  if (entry.compressionMethod === 8) return inflateRawSync(compressed);
  throw new Error(`zip: unsupported compression method ${entry.compressionMethod} for ${entry.name}`);
}

function extractXmlText(xml: string, tag: string): string[] {
  const out: string[] = [];
  const re = new RegExp(`<${tag}>([\\s\\S]*?)<\\/${tag}>`, "g");
  for (const match of xml.matchAll(re)) out.push(match[1] ?? "");
  return out;
}

function parseChartModelFromFixture(fixtureFile: string, chartKind: ChartModel["chartType"]["kind"]): ChartModel {
  const fixtureUrl = new URL(`../../../../../fixtures/xlsx/charts/${fixtureFile}`, import.meta.url);
  const bytes = readFileSync(fixtureUrl);
  const entries = parseZipEntries(bytes);

  const chartPath = [...entries.keys()]
    .filter((k) => /^xl\/charts\/chart\d+\.xml$/.test(k))
    .sort()[0];
  if (!chartPath) throw new Error(`fixture ${fixtureFile} missing xl/charts/chart*.xml`);

  const chartXml = readZipEntry(bytes, entries.get(chartPath)!).toString("utf8");

  const title = extractXmlText(chartXml, "a:t").join("").trim() || null;

  const legendPos = chartXml.match(/<c:legendPos val="([^"]+)"\/>/)?.[1] ?? null;
  const overlayVal = chartXml.match(/<c:overlay val="([^"]+)"\/>/)?.[1] ?? null;
  const legend =
    legendPos === "r"
      ? { position: "right" as const, overlay: overlayVal === "1" }
      : legendPos
        ? { position: "none" as const, overlay: overlayVal === "1" }
        : null;

  const axes: NonNullable<ChartModel["axes"]> = [];
  for (const match of chartXml.matchAll(/<c:(catAx|valAx)>([\s\S]*?)<\/c:\1>/g)) {
    const tag = match[1];
    const axisContent = match[2] ?? "";
    const id = axisContent.match(/<c:axId val="([^"]+)"\/>/)?.[1] ?? null;
    const axPos = axisContent.match(/<c:axPos val="([^"]+)"\/>/)?.[1] ?? null;
    const position =
      axPos === "b"
        ? "bottom"
        : axPos === "t"
          ? "top"
          : axPos === "r"
            ? "right"
            : "left";

    const orientation = axisContent.match(/<c:orientation val="([^"]+)"\/>/)?.[1] ?? null;
    const reverseOrder = orientation === "maxMin";
    const minStr = axisContent.match(/<c:min val="([^"]+)"\/>/)?.[1];
    const maxStr = axisContent.match(/<c:max val="([^"]+)"\/>/)?.[1];
    const min = minStr != null ? Number(minStr) : null;
    const max = maxStr != null ? Number(maxStr) : null;
    const scaling =
      reverseOrder || Number.isFinite(min) || Number.isFinite(max)
        ? { reverseOrder, min: Number.isFinite(min) ? min : null, max: Number.isFinite(max) ? max : null }
        : null;

    const majorGridlines = axisContent.includes("<c:majorGridlines");
    const formatCode = axisContent.match(/<c:numFmt[^>]*formatCode="([^"]+)"/)?.[1] ?? null;

    axes.push({
      id,
      kind: tag === "catAx" ? "category" : "value",
      position,
      scaling,
      majorGridlines: majorGridlines ? true : undefined,
      formatCode,
    });
  }

  const series: ChartModel["series"] = [];
  for (const match of chartXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const serXml = match[1] ?? "";
    const catXml = serXml.match(/<c:cat>([\s\S]*?)<\/c:cat>/)?.[1] ?? null;
    const valXml = serXml.match(/<c:val>([\s\S]*?)<\/c:val>/)?.[1] ?? null;
    const xXml = serXml.match(/<c:xVal>([\s\S]*?)<\/c:xVal>/)?.[1] ?? null;
    const yXml = serXml.match(/<c:yVal>([\s\S]*?)<\/c:yVal>/)?.[1] ?? null;

    const parseCache = (xmlPart: string | null): Array<string | number> => {
      if (!xmlPart) return [];
      const values = [...xmlPart.matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)].map((m) => (m[1] ?? "").trim());
      const numCache = xmlPart.includes("<c:numCache");
      if (numCache) {
        return values
          .map((v) => Number(v))
          .filter((n) => Number.isFinite(n));
      }
      return values;
    };

    const categories = parseCache(catXml);
    const values = parseCache(valXml);
    const xValues = parseCache(xXml);
    const yValues = parseCache(yXml);

    series.push({
      categories: categories.length ? { cache: categories } : null,
      values: values.length ? { cache: values } : null,
      xValues: xValues.length ? { cache: xValues } : null,
      yValues: yValues.length ? { cache: yValues } : null,
    });
  }

  return {
    chartType: { kind: chartKind },
    title,
    legend,
    axes,
    series,
  };
}

describe("charts/layout", () => {
  test("bar (fixture) layout is deterministic and produces non-overlapping rects", () => {
    const model = parseChartModelFromFixture("bar.xlsx", "bar");
    const viewport = { x: 0, y: 0, width: 480, height: 320 };

    const layout1 = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);
    const layout2 = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);
    expect(layout1).toEqual(layout2);

    expect(layout1.plotAreaRect.width).toBeGreaterThan(0);
    expect(layout1.plotAreaRect.height).toBeGreaterThan(0);

    expect(layout1.titleRect).not.toBeNull();
    expect(layout1.legendRect).not.toBeNull();

    expectNonOverlapping(layout1.titleRect, layout1.plotAreaRect);
    expectNonOverlapping(layout1.legendRect, layout1.plotAreaRect);
    expect(layout1.axes.y.ticks.length).toBeGreaterThanOrEqual(5);
    expect(layout1.axes.y.ticks.length).toBeLessThanOrEqual(7);
    expect(layout1.axes.y.gridlines.length).toBeGreaterThan(0);
  });

  test("line (fixture) layout reserves space for a right legend", () => {
    const model = parseChartModelFromFixture("line.xlsx", "line");
    const viewport = { x: 0, y: 0, width: 520, height: 320 };
    const layout = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);

    expect(layout.legendRect).not.toBeNull();
    expect(layout.legend?.entries.length).toBeGreaterThan(0);

    expectNonOverlapping(layout.legendRect, layout.plotAreaRect);
    expectNonOverlapping(layout.titleRect, layout.legendRect);
    expect(layout.plotAreaRect.x + layout.plotAreaRect.width).toBeLessThan(layout.legendRect!.x);
  });

  test("scatter (fixture) layout applies explicit axis bounds and reverse order", () => {
    const base = parseChartModelFromFixture("scatter.xlsx", "scatter");
    const model: ChartModel = {
      ...base,
      axes: (base.axes ?? []).map((axis) => {
        if (axis.position === "bottom" && axis.kind === "value") {
          return { ...axis, scaling: { ...(axis.scaling ?? {}), reverseOrder: true } };
        }
        if (axis.position === "left" && axis.kind === "value") {
          return { ...axis, scaling: { min: 0, max: 10 }, majorGridlines: true, formatCode: "0" };
        }
        return axis;
      }),
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

  test("pie (fixture) layout omits axes and keeps plot area separate from legend", () => {
    const model = parseChartModelFromFixture("pie.xlsx", "pie");
    const viewport = { x: 0, y: 0, width: 480, height: 320 };
    const layout = computeChartLayout(model, DEFAULT_CHART_THEME, viewport);

    expect(Object.keys(layout.axes)).toEqual([]);
    expect(Object.keys(layout.scales)).toEqual([]);

    expect(layout.legendRect).not.toBeNull();
    expectNonOverlapping(layout.titleRect, layout.plotAreaRect);
    expectNonOverlapping(layout.legendRect, layout.plotAreaRect);
  });

  test("accepts OOXML-style axis/legend position codes", () => {
    const model: ChartModel = {
      chartType: { kind: "bar" },
      title: "Example Chart",
      legend: { position: "r", overlay: false },
      axes: [
        { kind: "category", position: "b" },
        { kind: "value", position: "l", majorGridlines: true },
      ],
      series: [
        {
          categories: { cache: ["A", "B", "C", "D"] },
          values: { cache: [2, 4, 3, 5] },
        },
      ],
    };

    const layout = computeChartLayout(model, DEFAULT_CHART_THEME, { x: 0, y: 0, width: 480, height: 320 });
    expect(layout.legendRect).not.toBeNull();
    expect(layout.axes.y.ticks.length).toBeGreaterThanOrEqual(5);
  });
});
