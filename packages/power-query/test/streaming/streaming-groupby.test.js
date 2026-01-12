import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";

function collectBatches(batches) {
  const grid = [];
  for (const batch of batches) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }
  return grid;
}

const csvText = [
  "Region,Sales,Customer",
  "East,100,a",
  "West,200,b",
  "East,150,a",
  "West,250,c",
  "East,10,d",
].join("\n") + "\n";

test("streaming v2 groupBy matches materialized execution and can spill", async () => {
  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        throw new Error("readText should not be called in streaming mode");
      },
      readTextStream: async function* () {
        yield "Region,Sales,Customer\nEast,100,a\n";
        yield "West,200,b\nEast,150,a\nWest,250,c\nEast,10,d\n";
      },
    },
  });

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readText: async () => csvText,
    },
  });

  const query = {
    id: "q_stream_groupby_spill",
    name: "Streaming groupBy spill",
    source: { type: "csv", path: "/tmp/groupby.csv", options: { hasHeaders: true } },
    steps: [
      {
        id: "s_group",
        name: "Group",
        operation: {
          type: "groupBy",
          groupColumns: ["Region"],
          aggregations: [
            { column: "Sales", op: "sum", as: "Total" },
            { column: "Customer", op: "countDistinct", as: "Customers" },
            { column: "Sales", op: "average", as: "Avg" },
          ],
        },
      },
    ],
  };

  /** @type {any[]} */
  const progress = [];
  const batches = [];
  await engineStreaming.executeQueryStreaming(query, {}, {
    batchSize: 8,
    materialize: false,
    onProgress: (evt) => progress.push(evt),
    onBatch: (batch) => batches.push(batch),
    streaming: { spill: { kind: "memory" }, maxInMemoryRows: 2 },
  });

  const streamed = collectBatches(batches);
  const expected = (await engineMaterialized.executeQuery(query, {}, {})).toGrid();
  assert.deepEqual(streamed, expected);

  assert.ok(progress.some((evt) => evt?.type === "stream:spill" && evt.operator === "groupBy"), "expected spill progress event");
});

