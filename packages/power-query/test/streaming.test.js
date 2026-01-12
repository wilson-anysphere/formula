import assert from "node:assert/strict";
import test from "node:test";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { QueryEngine } from "../src/engine.js";
import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";

let arrowAvailable = true;
try {
  // Validate Arrow is actually usable via the data-io helpers (pnpm workspaces
  // don't necessarily hoist `apache-arrow` to the repo root).
  arrowTableFromColumns({ __probe: [1] });
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

/**
 * Build a minimal Arrow-like table object sufficient for `ArrowTableAdapter`.
 * This lets us exercise ArrowTableAdapter code paths even when `apache-arrow`
 * isn't installed.
 *
 * @param {Record<string, unknown[]>} columns
 * @param {Record<string, string>} typeHints
 */
function makeFakeArrowTable(columns, typeHints) {
  const names = Object.keys(columns);
  const rowCount = Math.max(0, ...names.map((name) => columns[name]?.length ?? 0));

  return {
    numRows: rowCount,
    schema: {
      fields: names.map((name) => ({
        name,
        type: { toString: () => typeHints[name] ?? "Utf8" },
      })),
    },
    getChildAt: (index) => {
      const name = names[index];
      const values = columns[name] ?? [];
      return {
        length: rowCount,
        get: (rowIndex) => values[rowIndex],
      };
    },
    slice: (start, end) => {
      /** @type {Record<string, unknown[]>} */
      const sliced = {};
      for (const name of names) {
        sliced[name] = (columns[name] ?? []).slice(start, end);
      }
      return makeFakeArrowTable(sliced, typeHints);
    },
  };
}

/**
 * @param {unknown[][]} grid
 */
function gridDatesToIso(grid) {
  return grid.map((row) => row.map((cell) => (cell instanceof Date ? cell.toISOString() : cell)));
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

test("executeQueryStreamingNonMaterializing supports distinctRows across batches (Dates)", async () => {
  const engine = new QueryEngine();
  const d1 = new Date("2020-01-01T00:00:00.000Z");
  const d1b = new Date("2020-01-01T00:00:00.000Z");
  const d2 = new Date("2020-01-02T00:00:00.000Z");
  const d2b = new Date("2020-01-02T00:00:00.000Z");

  const table = DataTable.fromGrid(
    [
      ["Id", "When"],
      [1, d1],
      [2, d2],
      [1, d1b],
      [2, d2b],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const query = {
    id: "q_stream_distinct",
    name: "Stream Distinct",
    source: { type: "table", table: "t" },
    steps: [{ id: "s_distinct", name: "Distinct", operation: { type: "distinctRows", columns: null } }],
  };

  const batches = [];
  await engine.executeQueryStreaming(query, { tables: { t: table } }, {
    batchSize: 2,
    materialize: false,
    onBatch: (batch) => batches.push(batch),
  });

  const streamed = collectBatches(batches);
  const expected = (await engine.executeQuery(query, { tables: { t: table } }, {})).toGrid();
  assert.deepEqual(gridDatesToIso(streamed), gridDatesToIso(expected));
});

test("executeQueryStreaming can stream ArrowTableAdapter values without apache-arrow (fallback batching)", { skip: arrowAvailable }, async () => {
  const engine = new QueryEngine();

  const fake = new ArrowTableAdapter(
    makeFakeArrowTable(
      { Region: ["East", "West"], Sales: [100, 200] },
      { Region: "Utf8", Sales: "Int32" },
    ),
    [
      { name: "Region", type: "string" },
      { name: "Sales", type: "number" },
    ],
  );

  const query = {
    id: "q_stream_fake_arrow",
    name: "Stream Fake Arrow",
    source: { type: "table", table: "t" },
    steps: [],
  };

  const batches = [];
  await engine.executeQueryStreaming(query, { tables: { t: fake } }, {
    batchSize: 1,
    onBatch: (batch) => batches.push(batch),
  });

  assert.deepEqual(collectBatches(batches), [
    ["Region", "Sales"],
    ["East", 100],
    ["West", 200],
  ]);
});
