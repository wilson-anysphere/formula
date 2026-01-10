import test from "node:test";
import assert from "node:assert/strict";

import { cellKey, semanticDiff } from "../packages/versioning/src/diff/semanticDiff.js";

function sheetFromObject(obj) {
  const cells = new Map();
  for (const [k, v] of Object.entries(obj)) {
    cells.set(k, v);
  }
  return { cells };
}

test("semanticDiff: added cell", () => {
  const before = sheetFromObject({});
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 123 },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.added.length, 1);
  assert.deepEqual(diff.added[0].cell, { row: 0, col: 0 });
  assert.equal(diff.removed.length, 0);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.moved.length, 0);
});

test("semanticDiff: removed cell", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: "x" },
  });
  const after = sheetFromObject({});
  const diff = semanticDiff(before, after);
  assert.equal(diff.removed.length, 1);
  assert.deepEqual(diff.removed[0].cell, { row: 0, col: 0 });
  assert.equal(diff.added.length, 0);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.moved.length, 0);
});

test("semanticDiff: modified cell (value)", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: 1 },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 2 },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified[0].oldValue, 1);
  assert.equal(diff.modified[0].newValue, 2);
});

test("semanticDiff: format-only change", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: 1, format: { bold: false } },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 1, format: { bold: true } },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.formatOnly.length, 1);
  assert.deepEqual(diff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified.length, 0);
});

test("semanticDiff: moved cell detection", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: "move-me", formula: "=A1+B1" },
  });
  const after = sheetFromObject({
    [cellKey(2, 3)]: { value: "move-me", formula: "=B1 + A1" }, // commutative equiv
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 1);
  assert.deepEqual(diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(diff.moved[0].newLocation, { row: 2, col: 3 });
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
});

test("semanticDiff: semantic-equivalent formulas are not modified", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: null, formula: "=A1 + B1" },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: null, formula: "=B1+A1" },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
  assert.equal(diff.moved.length, 0);
  assert.equal(diff.formatOnly.length, 0);
});

