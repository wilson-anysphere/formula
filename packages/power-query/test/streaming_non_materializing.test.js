import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../src/engine.js";

function collectBatches(batches) {
  const grid = [];
  for (const batch of batches) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }
  return grid;
}

test("executeQueryStreaming(..., materialize:false) matches materialized execution for a streamable CSV pipeline", async () => {
  const csvText = ["Region,Sales", "East,100", "West,200", "East,150"].join("\n") + "\n";

  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        throw new Error("readText should not be called in streaming mode");
      },
      readTextStream: async function* () {
        // Split mid-field to ensure the incremental parser handles chunk boundaries.
        yield "Region,Sales\nEast,1";
        yield "00\nWest,200\nEast,150\n";
      },
    },
  });

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readText: async () => csvText,
    },
  });

  const query = {
    id: "q_stream_csv_non_materialize",
    name: "Stream CSV non-materialize",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [
      {
        id: "s_filter",
        name: "Filter",
        operation: {
          type: "filterRows",
          predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" },
        },
      },
      { id: "s_add", name: "Add", operation: { type: "addColumn", name: "Double", formula: "=[Sales] * 2" } },
      { id: "s_rename", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" } },
      { id: "s_type", name: "Type", operation: { type: "changeType", column: "Double", newType: "number" } },
      {
        id: "s_transform",
        name: "Transform",
        operation: { type: "transformColumns", transforms: [{ column: "Amount", formula: "_ + 1", newType: null }] },
      },
      { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Amount", "Double"] } },
      { id: "s_take", name: "Take", operation: { type: "take", count: 2 } },
    ],
  };

  const batches = [];
  const summary = await engineStreaming.executeQueryStreaming(query, {}, { batchSize: 2, materialize: false, onBatch: (b) => batches.push(b) });
  assert.equal(summary.rowCount, 2);
  assert.equal(summary.columnCount, 3);

  const streamed = collectBatches(batches);
  const expected = (await engineMaterialized.executeQuery(query, {}, {})).toGrid();
  assert.deepEqual(streamed, expected);
});

test("executeQueryStreaming(..., materialize:false) honors take by stopping the CSV stream early", async () => {
  const engine = new QueryEngine({
    fileAdapter: {
      readTextStream: async function* () {
        yield "A\n";
        yield "1\n";
        yield "2\n";
        throw new Error("stream read past take limit");
      },
    },
  });

  const query = {
    id: "q_take_early_stop",
    name: "Take early stop",
    source: { type: "csv", path: "/tmp/take.csv", options: { hasHeaders: true } },
    steps: [{ id: "s_take", name: "Take", operation: { type: "take", count: 2 } }],
  };

  const batches = [];
  const summary = await engine.executeQueryStreaming(query, {}, { batchSize: 1, materialize: false, onBatch: (b) => batches.push(b) });
  assert.equal(summary.rowCount, 2);

  const streamed = collectBatches(batches);
  assert.deepEqual(streamed, [["A"], [1], [2]]);
});

test("executeQueryStreaming(..., materialize:false) succeeds with a stream-only file adapter", async () => {
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        throw new Error("full-file reads are not supported");
      },
      readBinary: async () => {
        throw new Error("binary reads are not supported");
      },
      readTextStream: async function* () {
        yield "X,Y\n";
        yield "1,2\n";
      },
    },
  });

  const query = {
    id: "q_stream_only_adapter",
    name: "Stream-only adapter",
    source: { type: "csv", path: "/tmp/stream.csv", options: { hasHeaders: true } },
    steps: [{ id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["Y"] } }],
  };

  const batches = [];
  const summary = await engine.executeQueryStreaming(query, {}, { batchSize: 10, materialize: false, onBatch: (b) => batches.push(b) });
  assert.equal(summary.rowCount, 1);
  assert.equal(summary.columnCount, 1);
  assert.deepEqual(collectBatches(batches), [["Y"], [2]]);
});

