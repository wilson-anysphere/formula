import assert from "node:assert/strict";
import test from "node:test";

import { drawConditionalFormattingLayer } from "./draw.js";
import { RecordingContext2D } from "./recording_context.js";

test("drawConditionalFormattingLayer emits deterministic draw calls", () => {
  const ctx = new RecordingContext2D();
  const cellRects = [
    { a1: "A1", x: 0, y: 0, width: 50, height: 20 },
    { a1: "B1", x: 50, y: 0, width: 50, height: 20 },
    { a1: "C1", x: 100, y: 0, width: 50, height: 20 }
  ];
  const byCell = {
    A1: { style: { fill: "FFFF0000" } },
    B1: { data_bar: { color: "FF638EC6", fill_ratio: 0.5 } },
    C1: { icon: { set: "ThreeArrows", index: 2 } }
  };

  drawConditionalFormattingLayer(ctx, cellRects, byCell);

  assert.deepStrictEqual(ctx.commands, [
    ["fillStyle", "rgba(255,0,0,1)"],
    ["fillRect", 0, 0, 50, 20],
    ["fillStyle", "rgba(99,142,198,1)"],
    ["fillRect", 51, 5, 24, 10],
    ["fillStyle", "rgba(0,255,0,1)"],
    ["beginPath"],
    ["moveTo", 142.5, 5],
    ["lineTo", 137.5, 15],
    ["lineTo", 147.5, 15],
    ["closePath"],
    ["fill"]
  ]);
});
