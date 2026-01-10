import test from "node:test";
import assert from "node:assert/strict";

import { TextLayoutEngine } from "../src/index.js";

function makeMonospaceMeasurer(clusterWidth = 1) {
  const segmenter = new Intl.Segmenter("und", { granularity: "grapheme" });

  return {
    calls: 0,
    measure(text, font) {
      this.calls++;
      let clusters = 0;
      for (const _ of segmenter.segment(text)) clusters++;
      return {
        width: clusters * clusterWidth,
        ascent: font.sizePx * 0.8,
        descent: font.sizePx * 0.2,
      };
    },
  };
}

test("word wrap breaks at spaces and trims leading whitespace on wrapped lines", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const layout = engine.layout({
    text: "Hello world",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 5,
    wrapMode: "word",
    align: "left",
    direction: "ltr",
  });

  assert.deepEqual(layout.lines.map((l) => l.text), ["Hello", "world"]);
});

test("char wrap uses grapheme clusters (combining marks stay with base)", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const layout = engine.layout({
    text: "a\u0301b", // a + combining acute + b
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 1,
    wrapMode: "char",
    align: "left",
    direction: "ltr",
  });

  assert.deepEqual(layout.lines.map((l) => l.text), ["a\u0301", "b"]);
});

test("auto direction chooses rtl when the first strong character is Hebrew", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const layout = engine.layout({
    text: "שלום עולם",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 4,
    wrapMode: "word",
    align: "start",
    direction: "auto",
  });

  assert.equal(layout.direction, "rtl");
  assert.equal(layout.resolvedAlign, "right");
  assert.deepEqual(layout.lines.map((l) => l.text), ["שלום", "עולם"]);
});

test("layout results are cached to avoid repeated measurement work", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const options = {
    text: "Hello world",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 5,
    wrapMode: "word",
    align: "left",
    direction: "ltr",
  };

  engine.layout(options);
  const callsAfterFirst = measurer.calls;
  engine.layout(options);

  assert.equal(measurer.calls, callsAfterFirst);
});

test("layout cache keys include non-metric run metadata so returned runs stay correct", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const base = {
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 100,
    wrapMode: "none",
    align: "left",
    direction: "ltr",
  };

  const layoutA = engine.layout({
    ...base,
    runs: [{ text: "Hello", color: "red" }],
  });

  const layoutB = engine.layout({
    ...base,
    runs: [{ text: "Hello", color: "blue" }],
  });

  assert.notStrictEqual(layoutA, layoutB);
  assert.equal(layoutA.lines[0].runs[0].color, "red");
  assert.equal(layoutB.lines[0].runs[0].color, "blue");
});

test("word wrap falls back to char wrapping when there are no word break opportunities", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const layout = engine.layout({
    text: "superlong",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 4,
    wrapMode: "word",
    align: "left",
    direction: "ltr",
  });

  assert.deepEqual(layout.lines.map((l) => l.text), ["supe", "rlon", "g"]);
});

test("maxLines truncates and applies ellipsis within maxWidth", () => {
  const measurer = makeMonospaceMeasurer();
  const engine = new TextLayoutEngine(measurer);

  const layout = engine.layout({
    text: "Hello world",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 5,
    wrapMode: "word",
    align: "left",
    direction: "ltr",
    maxLines: 1,
  });

  assert.deepEqual(layout.lines.map((l) => l.text), ["Hell…"]);
  assert.equal(layout.lines[0].width, 5);
});
