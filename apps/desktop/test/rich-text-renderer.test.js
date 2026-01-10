import assert from "node:assert/strict";
import test from "node:test";

import { renderRichText } from "../src/grid/text/rich-text/render.js";

class RecordingContext {
  constructor() {
    this.calls = [];
    this.font = "";
    this.fillStyle = "";
    this.strokeStyle = "";
    this.lineWidth = 1;
    this.textBaseline = "alphabetic";
  }

  save() {
    this.calls.push({ op: "save" });
  }
  restore() {
    this.calls.push({ op: "restore" });
  }
  beginPath() {
    this.calls.push({ op: "beginPath" });
  }
  rect(x, y, w, h) {
    this.calls.push({ op: "rect", x, y, w, h });
  }
  clip() {
    this.calls.push({ op: "clip" });
  }
  moveTo(x, y) {
    this.calls.push({ op: "moveTo", x, y });
  }
  lineTo(x, y) {
    this.calls.push({ op: "lineTo", x, y });
  }
  stroke() {
    this.calls.push({ op: "stroke", strokeStyle: this.strokeStyle, lineWidth: this.lineWidth });
  }

  measureText(text) {
    // Deterministic width estimate based on current font size.
    const match = /([0-9]+)px/.exec(this.font);
    const px = match ? Number(match[1]) : 12;
    return { width: text.length * px * 0.6 };
  }

  fillText(text, x, y) {
    this.calls.push({ op: "fillText", text, x, y, font: this.font, fillStyle: this.fillStyle });
  }
}

test("renderRichText emits distinct font strings for bold segments", () => {
  const ctx = new RecordingContext();
  renderRichText(
    /** @type {any} */ (ctx),
    {
      text: "A😀Bold",
      runs: [
        { start: 0, end: 1, style: {} },
        { start: 1, end: 2, style: { bold: true } },
        { start: 2, end: 6, style: {} },
      ],
    },
    { x: 0, y: 0, width: 200, height: 20 },
    { fontFamily: "Calibri", fontSizePx: 12, align: "left", verticalAlign: "middle" }
  );

  const fillCalls = ctx.calls.filter((c) => c.op === "fillText");
  assert.equal(fillCalls.length, 3);
  assert.equal(fillCalls[0].text, "A");
  assert.equal(fillCalls[1].text, "😀");
  assert.equal(fillCalls[2].text, "Bold");
  assert.notEqual(fillCalls[0].font, fillCalls[1].font);
  assert.match(fillCalls[1].font, /bold/);
});
