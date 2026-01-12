import test from "node:test";
import assert from "node:assert/strict";
import { toggleBold } from "../src/formatting/toolbar.js";

const EXCEL_MAX_ROW = 1_048_576 - 1;
const EXCEL_MAX_COL = 16_384 - 1;

test("toolbar formatting guard blocks oversized non-band selections", () => {
  let called = false;
  const doc = {
    setRangeFormat() {
      called = true;
      return true;
    },
  };

  const ok = toggleBold(doc, 0, {
    start: { row: 0, col: 0 },
    end: { row: 500, col: 500 }, // 501 * 501 = 251,001 > 100,000
  });

  assert.equal(ok, false);
  assert.equal(called, false);
});

test("toolbar formatting guard allows full-sheet band selections", () => {
  const calls = [];
  const doc = {
    setRangeFormat(...args) {
      calls.push(args);
      return true;
    },
  };

  const ok = toggleBold(
    doc,
    0,
    {
      start: { row: 0, col: 0 },
      end: { row: EXCEL_MAX_ROW, col: EXCEL_MAX_COL },
    },
    { next: true },
  );

  assert.equal(ok, true);
  assert.equal(calls.length, 1);
});

test("toolbar formatting guard blocks oversized full-width row bands", () => {
  let called = false;
  const doc = {
    setRangeFormat() {
      called = true;
      return true;
    },
  };

  const ok = toggleBold(doc, 0, {
    start: { row: 0, col: 0 },
    end: { row: 50_000, col: EXCEL_MAX_COL }, // 50_001 rows > 50,000 row-band cap
  });

  assert.equal(ok, false);
  assert.equal(called, false);
});

test("toolbar formatting guard blocks multi-range selections over the cap", () => {
  let called = false;
  const doc = {
    setRangeFormat() {
      called = true;
      return true;
    },
  };

  const ranges = [
    // 300 * 300 = 90,000
    { start: { row: 0, col: 0 }, end: { row: 299, col: 299 } },
    { start: { row: 0, col: 300 }, end: { row: 299, col: 599 } },
  ];

  const ok = toggleBold(doc, 0, ranges);

  assert.equal(ok, false);
  assert.equal(called, false);
});

