import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { TextLayoutEngine, createHarfBuzzTextMeasurer } from "../src/index.js";

/**
 * @param {number} value
 * @param {number} decimals
 */
function round(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

const notoSansData = await readFile(new URL("./fixtures/fonts/NotoSans-Regular.ttf", import.meta.url));
const notoSansHebrewData = await readFile(
  new URL("./fixtures/fonts/NotoSansHebrew-Regular.ttf", import.meta.url),
);

const measurer = await createHarfBuzzTextMeasurer({
  fonts: [
    { family: "Noto Sans", weight: 400, style: "normal", data: notoSansData },
    { family: "Noto Sans Hebrew", weight: 400, style: "normal", data: notoSansHebrewData },
  ],
  // Bidirectional fallback so either script can render in a single run.
  fallbackFamilies: ["Noto Sans Hebrew", "Noto Sans"],
});

const engine = new TextLayoutEngine(measurer);

test("HarfBuzz measurer returns deterministic widths (kerning via shaping)", () => {
  const font = { family: "Noto Sans", sizePx: 13, weight: 400 };

  const a = engine.measure("A", font).width;
  const v = engine.measure("V", font).width;
  const av = engine.measure("AV", font).width;

  assert.equal(round(a, 3), 8.307);
  assert.equal(round(v, 3), 7.8);
  assert.equal(round(av, 3), 15.587);
  assert.ok(av < a + v, "expected kerning in AV to reduce total advance");
});

test("HarfBuzz layout produces stable word wrapping and line widths", () => {
  const font = { family: "Noto Sans", sizePx: 13, weight: 400 };
  const layout = engine.layout({
    text: "Hello world",
    font,
    maxWidth: 40,
    wrapMode: "word",
    align: "left",
    direction: "ltr",
  });

  assert.deepEqual(layout.lines.map((l) => l.text), ["Hello", "world"]);
  assert.equal(round(layout.lines[0].width, 3), 31.538);
  assert.equal(round(layout.lines[1].width, 3), 34.801);
});

test("RTL auto-direction keeps resolvedAlign stable and computes right-aligned x offsets", () => {
  const font = { family: "Noto Sans Hebrew", sizePx: 13, weight: 400 };
  const layout = engine.layout({
    text: "שלום world",
    font,
    maxWidth: 40,
    wrapMode: "word",
    align: "start",
    direction: "auto",
  });

  assert.equal(layout.direction, "rtl");
  assert.equal(layout.resolvedAlign, "right");
  assert.deepEqual(layout.lines.map((l) => l.text), ["שלום", "world"]);

  // maxWidth - lineWidth (right alignment).
  assert.equal(round(layout.lines[0].x, 3), 11.049);
  assert.equal(round(layout.lines[1].x, 3), 5.758);
});

test("font fallback is deterministic when the primary font lacks glyphs (mixed-script single run)", () => {
  const font = { family: "Noto Sans", sizePx: 13, weight: 400 };

  const mixed = engine.measure("Hello שלום", font).width;
  assert.equal(round(mixed, 3), 63.869);
});

test("per-run font changes are reflected in measurement (multi-run mixed sizes)", () => {
  const fontLatin = { family: "Noto Sans", sizePx: 13, weight: 400 };
  const fontHebrewBig = { family: "Noto Sans Hebrew", sizePx: 20, weight: 400 };

  const layout = engine.layout({
    runs: [
      { text: "Hello ", font: fontLatin },
      { text: "שלום", font: fontHebrewBig },
    ],
    font: fontLatin,
    maxWidth: Infinity,
    wrapMode: "none",
    align: "left",
    direction: "ltr",
  });

  assert.equal(layout.lines.length, 1);
  assert.equal(layout.lines[0].text, "Hello שלום");
  assert.equal(round(layout.lines[0].width, 3), 79.458);
});

