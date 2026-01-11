import assert from "node:assert/strict";
import test from "node:test";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { QueryEngine } from "../src/engine.js";
import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";

let arrowAvailable = true;
try {
  await import("apache-arrow");
} catch {
  arrowAvailable = false;
}

function collectBatches(batches) {
  const grid = [];
  for (const batch of batches) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }
  return grid;
}

test("executeQueryStreaming streams DataTable results in grid batches", async () => {
  const engine = new QueryEngine();
  const table = DataTable.fromGrid(
    [
      ["Region", "Sales"],
      ["East", 100],
      ["West", 200],
      ["East", 150],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const query = {
    id: "q_stream",
    name: "Stream",
    source: { type: "table", table: "t" },
    steps: [{ id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } } }],
  };

  const batches = [];
  await engine.executeQueryStreaming(query, { tables: { t: table } }, {
    batchSize: 2,
    onBatch: (batch) => batches.push(batch),
  });

  const streamed = collectBatches(batches);
  const expected = (await engine.executeQuery(query, { tables: { t: table } }, {})).toGrid();
  assert.deepEqual(streamed, expected);
});

test("executeQueryStreaming streams Arrow results using arrowTableToGridBatches", { skip: !arrowAvailable }, async () => {
  const engine = new QueryEngine();
  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      Region: ["East", "West", "East"],
      Sales: [100, 200, 150],
    }),
  );

  const query = {
    id: "q_stream_arrow",
    name: "Stream Arrow",
    source: { type: "table", table: "t" },
    steps: [
      { id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } } },
      { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] } },
    ],
  };

  const batches = [];
  await engine.executeQueryStreaming(query, { tables: { t: arrowTable } }, {
    batchSize: 2,
    onBatch: (batch) => batches.push(batch),
  });

  const streamed = collectBatches(batches);
  const expected = (await engine.executeQuery(query, { tables: { t: arrowTable } }, {})).toGrid();
  assert.deepEqual(streamed, expected);
});

test("executeQueryStreaming uses adapter column names for Arrow headers (renameColumn)", { skip: !arrowAvailable }, async () => {
  const engine = new QueryEngine();
  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      Region: ["East", "West"],
      Sales: [100, 200],
    }),
  );

  const query = {
    id: "q_stream_arrow_rename",
    name: "Stream Arrow Rename",
    source: { type: "table", table: "t" },
    steps: [{ id: "s_rename", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" } }],
  };

  const grid = [];
  await engine.executeQueryStreaming(query, { tables: { t: arrowTable } }, {
    batchSize: 2,
    onBatch: (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        grid[batch.rowOffset + i] = batch.values[i];
      }
    },
  });

  assert.deepEqual(grid[0], ["Region", "Amount"]);
  assert.deepEqual(grid[1], ["East", 100]);
  assert.deepEqual(grid[2], ["West", 200]);
});

test("executeQueryStreaming emits Arrow date values as Date objects", { skip: !arrowAvailable }, async () => {
  const engine = new QueryEngine();
  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      When: [new Date("2020-01-01T00:00:00.000Z"), new Date("2020-01-02T00:00:00.000Z")],
      Value: [1, 2],
    }),
  );

  const query = {
    id: "q_stream_arrow_date",
    name: "Stream Arrow Date",
    source: { type: "table", table: "t" },
    steps: [],
  };

  const batches = [];
  await engine.executeQueryStreaming(query, { tables: { t: arrowTable } }, {
    batchSize: 1,
    onBatch: (batch) => batches.push(batch),
  });

  const streamed = collectBatches(batches);
  const expected = (await engine.executeQuery(query, { tables: { t: arrowTable } }, {})).toGrid();
  assert.deepEqual(streamed, expected);
  assert.ok(streamed[1][0] instanceof Date);
});
