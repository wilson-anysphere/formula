import test from "node:test";
import assert from "node:assert/strict";

import { TextLayoutEngine } from "../src/index.js";
import { drawTextLayout } from "../src/draw.js";

function makeMonospaceMeasurer(clusterWidth = 1) {
  const segmenter = new Intl.Segmenter("und", { granularity: "grapheme" });
  return {
    measure(text, font) {
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

