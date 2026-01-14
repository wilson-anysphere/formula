import assert from "node:assert/strict";
import test from "node:test";

import { cellToA1, rectToA1 } from "../src/workbook/rect.js";

test("cellToA1 converts 0-based row/col to A1 notation (including multi-letter columns)", () => {
  assert.equal(cellToA1(0, 0), "A1");
  assert.equal(cellToA1(0, 25), "Z1");
  assert.equal(cellToA1(0, 26), "AA1");
  assert.equal(cellToA1(0, 27), "AB1");
  assert.equal(cellToA1(9, 52), "BA10");
});

test("rectToA1 uses A1 notation for ranges", () => {
  assert.equal(rectToA1({ r0: 0, c0: 26, r1: 0, c1: 27 }), "AA1:AB1");
  assert.equal(rectToA1({ r0: 9, c0: 52, r1: 9, c1: 52 }), "BA10");
});
