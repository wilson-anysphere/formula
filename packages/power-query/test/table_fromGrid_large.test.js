import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../src/engine.js";
import { DataTable } from "../src/table.js";

test("Large datasets avoid stack overflows during table/range width calculations", async () => {
  const rows = 150_000;
  // Use a shared row to keep allocations low while still exercising the
  // "many rows" argument-limit edge case.
  const shared = [0, 0];
  const grid = Array.from({ length: rows }, () => shared);
  grid[0] = [1, 2];
  grid[rows - 1] = [3, 4];

  const engine = new QueryEngine();
  const query = {
    id: "q_range_large",
    name: "Range large",
    source: { type: "range", range: { values: grid, hasHeaders: false } },
    steps: [{ id: "s_take", name: "Take", operation: { type: "take", count: 2 } }],
  };

  /** @type {unknown[][]} */
  const streamed = [];
  const summary = await engine.executeQueryStreaming(query, {}, {
    materialize: false,
    includeHeader: false,
    batchSize: 2,
    onBatch: (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        streamed[batch.rowOffset + i] = batch.values[i];
      }
    },
  });
  assert.equal(summary.rowCount, 2);
  assert.equal(summary.columnCount, 2);
  assert.deepEqual(streamed, [
    [1, 2],
    [0, 0],
  ]);

  const table = DataTable.fromGrid(grid, { hasHeaders: false, inferTypes: false });

  assert.equal(table.rowCount, rows);
  assert.equal(table.columnCount, 2);
  assert.equal(table.getCell(0, 0), 1);
  assert.equal(table.getCell(0, 1), 2);
  assert.equal(table.getCell(rows - 1, 0), 3);
  assert.equal(table.getCell(rows - 1, 1), 4);
});
