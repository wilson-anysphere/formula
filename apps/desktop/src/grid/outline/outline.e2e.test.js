import test from "node:test";
import assert from "node:assert/strict";

import { Outline, isHidden } from "./outline.js";

test("collapse and expand group hides and shows rows (Excel-style)", () => {
  const outline = new Outline();

  // Mimic grouping rows 2-4.
  outline.groupRows(2, 4);
  outline.recomputeOutlineHiddenRows();

  for (let r = 2; r <= 4; r += 1) {
    assert.equal(outline.rows.entry(r).level, 1);
    assert.equal(isHidden(outline.rows.entry(r).hidden), false);
  }

  // Collapse via summary row 5 (summaryBelow=true).
  outline.toggleRowGroup(5);
  assert.equal(outline.rows.entry(5).collapsed, true);
  for (let r = 2; r <= 4; r += 1) {
    assert.equal(isHidden(outline.rows.entry(r).hidden), true);
  }

  // Expand again.
  outline.toggleRowGroup(5);
  assert.equal(outline.rows.entry(5).collapsed, false);
  for (let r = 2; r <= 4; r += 1) {
    assert.equal(isHidden(outline.rows.entry(r).hidden), false);
  }
});

test("collapse and expand group hides and shows columns (Excel-style)", () => {
  const outline = new Outline();

  outline.groupCols(2, 4);
  outline.recomputeOutlineHiddenCols();

  for (let c = 2; c <= 4; c += 1) {
    assert.equal(outline.cols.entry(c).level, 1);
    assert.equal(isHidden(outline.cols.entry(c).hidden), false);
  }

  // Collapse via summary column 5 (summaryRight=true).
  outline.toggleColGroup(5);
  assert.equal(outline.cols.entry(5).collapsed, true);
  for (let c = 2; c <= 4; c += 1) {
    assert.equal(isHidden(outline.cols.entry(c).hidden), true);
  }

  outline.toggleColGroup(5);
  assert.equal(outline.cols.entry(5).collapsed, false);
  for (let c = 2; c <= 4; c += 1) {
    assert.equal(isHidden(outline.cols.entry(c).hidden), false);
  }
});
