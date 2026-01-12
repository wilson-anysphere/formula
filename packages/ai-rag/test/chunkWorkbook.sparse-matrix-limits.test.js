import assert from "node:assert/strict";
import test from "node:test";

import { chunkWorkbook } from "../src/workbook/chunkWorkbook.js";

test("chunkWorkbook: does not allocate dense visited grids for sparse large matrices", () => {
  // Model a sheet with Excel-scale row indices via a sparse array. This should not force
  // region detection to allocate a `rows x cols` visited matrix.
  /** @type {any[]} */
  const cells = [];
  cells.length = 1_000_000;
  cells[999_999] = [{ v: "A" }, { v: "B" }];

  const workbook = {
    id: "wb1",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: 999_999, c0: 0, r1: 999_999, c1: 1 });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[0][1].v, "B");
});

