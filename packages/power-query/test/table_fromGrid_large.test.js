import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../src/table.js";

test("DataTable.fromGrid handles large grids without stack overflow", () => {
  const rows = 150_000;
  const grid = Array.from({ length: rows }, (_, i) => [i, i * 2]);

  const table = DataTable.fromGrid(grid, { hasHeaders: false, inferTypes: false });

  assert.equal(table.rowCount, rows);
  assert.equal(table.columnCount, 2);
  assert.equal(table.getCell(0, 0), 0);
  assert.equal(table.getCell(rows - 1, 1), (rows - 1) * 2);
});

