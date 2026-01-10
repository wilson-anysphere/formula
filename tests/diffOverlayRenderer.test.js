import test from "node:test";
import assert from "node:assert/strict";

import { DiffOverlayRenderer } from "../apps/desktop/src/grid/diff-renderer/DiffOverlayRenderer.js";

function makeFakeCtx() {
  /** @type {any[]} */
  const calls = [];

  /** @type {any} */
  const ctx = {
    canvas: { width: 100, height: 100 },
    _fillStyle: null,
    _strokeStyle: null,
    _globalAlpha: 1,
    _lineWidth: 1,
    save() {
      calls.push(["save"]);
    },
    restore() {
      calls.push(["restore"]);
    },
    clearRect(x, y, w, h) {
      calls.push(["clearRect", x, y, w, h]);
    },
    fillRect(x, y, w, h) {
      calls.push(["fillRect", x, y, w, h]);
    },
    strokeRect(x, y, w, h) {
      calls.push(["strokeRect", x, y, w, h]);
    },
    beginPath() {
      calls.push(["beginPath"]);
    },
    moveTo(x, y) {
      calls.push(["moveTo", x, y]);
    },
    lineTo(x, y) {
      calls.push(["lineTo", x, y]);
    },
    stroke() {
      calls.push(["stroke"]);
    },
    set fillStyle(v) {
      this._fillStyle = v;
      calls.push(["fillStyle", v]);
    },
    get fillStyle() {
      return this._fillStyle;
    },
    set strokeStyle(v) {
      this._strokeStyle = v;
      calls.push(["strokeStyle", v]);
    },
    get strokeStyle() {
      return this._strokeStyle;
    },
    set globalAlpha(v) {
      this._globalAlpha = v;
      calls.push(["globalAlpha", v]);
    },
    get globalAlpha() {
      return this._globalAlpha;
    },
    set lineWidth(v) {
      this._lineWidth = v;
      calls.push(["lineWidth", v]);
    },
    get lineWidth() {
      return this._lineWidth;
    },
  };

  return { ctx, calls };
}

test("DiffOverlayRenderer renders per-cell highlights for all diff buckets", () => {
  const renderer = new DiffOverlayRenderer();

  const diff = {
    added: [{ cell: { row: 0, col: 0 } }],
    removed: [{ cell: { row: 0, col: 1 } }],
    modified: [{ cell: { row: 1, col: 0 } }],
    formatOnly: [{ cell: { row: 1, col: 1 } }],
    moved: [
      {
        oldLocation: { row: 2, col: 2 },
        newLocation: { row: 3, col: 3 },
        value: "x",
      },
    ],
  };

  const { ctx, calls } = makeFakeCtx();
  const getCellRect = (row, col) => ({ x: col * 10, y: row * 10, width: 10, height: 10 });

  renderer.render(ctx, diff, { getCellRect });

  const fillRectCalls = calls.filter((c) => c[0] === "fillRect");
  const strokeRectCalls = calls.filter((c) => c[0] === "strokeRect");
  const strikethroughStrokes = calls.filter((c) => c[0] === "stroke");

  // One highlight per cell across all buckets, with moves highlighting both old and new locations.
  assert.equal(fillRectCalls.length, 6);
  assert.equal(strokeRectCalls.length, 6);

  // Removed cells draw a strikethrough line; we do this for removed + moved-from.
  assert.equal(strikethroughStrokes.length, 2);
});

