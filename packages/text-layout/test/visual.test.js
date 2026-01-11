import test from "node:test";
import assert from "node:assert/strict";

import { TextLayoutEngine } from "../src/index.js";
import { drawTextLayout } from "../src/draw.js";
import GraphemeSplitter from "grapheme-splitter";

function makeMonospaceMeasurer(clusterWidth = 1) {
  const splitter = new GraphemeSplitter();
  return {
    measure(text, font) {
      const clusters = splitter.countGraphemes(text);
      return {
        width: clusters * clusterWidth,
        ascent: font.sizePx * 0.8,
        descent: font.sizePx * 0.2,
      };
    },
  };
}

class FakeCanvasContext {
  constructor() {
    /** @type {Array<{text: string, x: number, y: number}>} */
    this.calls = [];
  }

  save() {}
  restore() {}
  translate() {}
  rotate() {}

  fillText(text, x, y) {
    this.calls.push({ text, x, y });
  }
}

test("drawTextLayout respects alignment offsets (simple snapshot)", () => {
  const engine = new TextLayoutEngine(makeMonospaceMeasurer());
  const layout = engine.layout({
    text: "abcd efgh",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 6,
    wrapMode: "word",
    align: "center",
    direction: "ltr",
    lineHeightPx: 12,
  });

  const ctx = new FakeCanvasContext();
  drawTextLayout(ctx, layout, 0, 0);

  assert.deepEqual(
    ctx.calls.map((c) => ({ text: c.text, x: c.x })),
    [
      { text: "abcd", x: 1 },
      { text: "efgh", x: 1 },
    ],
  );
});

test("drawTextLayout positions RTL text using start alignment (mixed-script snapshot)", () => {
  const engine = new TextLayoutEngine(makeMonospaceMeasurer());
  const layout = engine.layout({
    text: "שלום world",
    font: { family: "Inter", sizePx: 10, weight: 400 },
    maxWidth: 6,
    wrapMode: "word",
    align: "start",
    direction: "auto",
    lineHeightPx: 12,
  });

  assert.equal(layout.direction, "rtl");
  assert.equal(layout.resolvedAlign, "right");

  const ctx = new FakeCanvasContext();
  drawTextLayout(ctx, layout, 0, 0);

  assert.deepEqual(
    ctx.calls.map((c) => ({ text: c.text, x: c.x })),
    [
      { text: "שלום", x: 2 },
      { text: "world", x: 1 },
    ],
  );
});
