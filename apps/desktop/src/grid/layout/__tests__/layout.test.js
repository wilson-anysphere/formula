import assert from "node:assert/strict";
import test from "node:test";

import {
  MergedRegionIndex,
  isInteriorHorizontalGridline,
  isInteriorVerticalGridline,
  mergedRangeRect,
} from "../mergedCells.js";
import { layoutCellText } from "../textLayout.js";

const metrics = {
  getColWidth: (col) => (col === 0 ? 60 : 40),
  getRowHeight: (_row) => 20,
  getColLeft: (col) => (col === 0 ? 0 : 60),
  getRowTop: (row) => row * 20,
};

test("merged cells: computes merged rect and skips interior cells", () => {
  const regions = [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }];
  const index = new MergedRegionIndex(regions);

  assert.deepEqual(index.rangeAt({ row: 1, col: 1 }), regions[0]);
  assert.deepEqual(index.resolveCell({ row: 1, col: 1 }), { row: 0, col: 0 });
  assert.equal(index.shouldSkipCell({ row: 0, col: 1 }), true);
  assert.equal(index.shouldSkipCell({ row: 0, col: 0 }), false);

  // Internal gridlines are suppressed inside the merge region.
  assert.equal(isInteriorVerticalGridline(index, 0, 0), true); // between A and B
  assert.equal(isInteriorHorizontalGridline(index, 0, 0), true); // between row 1 and 2

  assert.deepEqual(mergedRangeRect(regions[0], metrics), { x: 0, y: 0, width: 100, height: 40 });
});

test("text layout: wraps text within cell width", () => {
  const result = layoutCellText({
    text: "hello world again",
    row: 0,
    col: 0,
    cellRect: { x: 0, y: 0, width: 60, height: 20 },
    style: {
      wrap: true,
      horizontalAlign: "left",
      verticalAlign: "top",
      rotationDeg: 0,
      lineHeight: 10,
    },
    measure: (t) => t.length * 6,
  });

  assert.deepEqual(
    result.lines.map((l) => l.text),
    ["hello ", "world ", "again"],
  );
});

test("text layout: overflows into adjacent empty cells when wrap is off", () => {
  const result = layoutCellText({
    text: "verylongtext",
    row: 0,
    col: 0,
    cellRect: { x: 0, y: 0, width: 60, height: 20 },
    style: {
      wrap: false,
      horizontalAlign: "left",
      verticalAlign: "center",
      rotationDeg: 0,
      lineHeight: 10,
    },
    measure: (t) => t.length * 10,
    overflow: {
      isCellEmpty: (_row, _col) => true,
      getColWidth: (c) => (c === 0 ? 60 : 40),
    },
  });

  assert.deepEqual(result.drawRect, { x: 0, y: 0, width: 140, height: 20 });
});
