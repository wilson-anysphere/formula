import { performance } from "node:perf_hooks";

import { QueryEngine } from "../src/engine.js";

function fmtMs(ms) {
  return `${ms.toFixed(1)}ms`;
}

function mem() {
  const { heapUsed } = process.memoryUsage();
  return `${(heapUsed / 1024 / 1024).toFixed(1)}MB`;
}

/**
 * Generate CSV chunks without first materializing the entire file.
 *
 * @param {number} rowCount
 * @param {number} rowsPerChunk
 */
async function* makeCsvChunks(rowCount, rowsPerChunk) {
  yield "Id,Group,Value\n";
  let i = 0;
  while (i < rowCount) {
    const end = Math.min(rowCount, i + rowsPerChunk);
    let chunk = "";
    for (; i < end; i++) {
      // Reverse-ish ordering to keep sort/groupBy doing real work.
      const value = (rowCount - i) % 1000;
      chunk += `${i},${value % 25},${value}\n`;
    }
    yield chunk;
  }
}

const ROWS = 100_000;
const BATCH_SIZE = 2048;

const sortQuery = {
  id: "q_stream_v2_sort",
  name: "Streaming v2 sort",
  source: { type: "csv", path: "/tmp/v2_sort.csv", options: { hasHeaders: true } },
  steps: [{ id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Value", direction: "ascending" }] } }],
};

const groupQuery = {
  id: "q_stream_v2_group",
  name: "Streaming v2 groupBy",
  source: { type: "csv", path: "/tmp/v2_group.csv", options: { hasHeaders: true } },
  steps: [
    {
      id: "s_group",
      name: "Group",
      operation: { type: "groupBy", groupColumns: ["Group"], aggregations: [{ column: "Value", op: "sum", as: "Total" }] },
    },
  ],
};

async function benchQuery(query, label) {
  const engine = new QueryEngine({
    fileAdapter: {
      readTextStream: () => makeCsvChunks(ROWS, 10_000),
      readText: async () => {
        throw new Error("unexpected full read");
      },
    },
  });

  const before = mem();
  const start = performance.now();

  let rows = 0;
  await engine.executeQueryStreaming(query, {}, {
    batchSize: BATCH_SIZE,
    includeHeader: false,
    materialize: false,
    streaming: { maxInMemoryRows: 10_000 },
    onBatch: (batch) => {
      rows += batch.values.length;
    },
  });

  const end = performance.now();
  console.log(`${label}: ${fmtMs(end - start)} rowsOut=${rows.toLocaleString()} heap=${before} -> ${mem()}`);
}

console.log("Power Query streaming v2 benchmark (JS, single-threaded)");
console.log(`Node ${process.version}`);
console.log(`rows=${ROWS.toLocaleString()} batchSize=${BATCH_SIZE}`);
console.log("");

await benchQuery(sortQuery, "sortRows");
await benchQuery(groupQuery, "groupBy");

