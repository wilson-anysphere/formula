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

/**
 * @param {number} rowCount
 */
function makeSortCsv(rowCount) {
  const rows = ["Group,Value,Original"];
  for (let i = 0; i < rowCount; i++) {
    const value = (rowCount - i) % 25;
    const group = value % 2 === 0 ? "A" : "B";
    rows.push(`${group},${value},row-${i}`);
  }
  return rows.join("\n") + "\n";
}

test("streaming v2 sortRows matches materialized execution and spills when threshold is exceeded", async () => {
  const csvText = makeSortCsv(500);

  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        throw new Error("readText should not be called in streaming mode");
      },
      readTextStream: async function* () {
        // Chunk the text to exercise streaming parsing.
        yield csvText.slice(0, Math.floor(csvText.length / 2));
        yield csvText.slice(Math.floor(csvText.length / 2));
      },
    },
  });

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readText: async () => csvText,
    },
  });

  const query = {
    id: "q_stream_sort_spill",
    name: "Streaming sort spill",
    source: { type: "csv", path: "/tmp/sort.csv", options: { hasHeaders: true } },
    steps: [{ id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Value", direction: "ascending" }] } }],
  };

  /** @type {any[]} */
  const progress = [];
  const batches = [];
  await engineStreaming.executeQueryStreaming(query, {}, {
    batchSize: 64,
    materialize: false,
    onProgress: (evt) => progress.push(evt),
    onBatch: (batch) => batches.push(batch),
    streaming: { spill: { kind: "memory" }, maxInMemoryRows: 10 },
  });

  const streamed = collectBatches(batches);
  const expected = (await engineMaterialized.executeQuery(query, {}, {})).toGrid();
  assert.deepEqual(streamed, expected);

  assert.ok(progress.some((evt) => evt?.type === "stream:spill" && evt.operator === "sortRows"), "expected spill progress event");
});

