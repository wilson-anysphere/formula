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
 * @param {number} rowCount
 */
function makeCsvText(rowCount) {
  const parts = ["A,B\n"];
  for (let i = 0; i < rowCount; i++) {
    parts.push(`${i},${i * 2}\n`);
  }
  return parts.join("");
}

/**
 * Generate CSV chunks without first materializing the entire file.
 *
 * @param {number} rowCount
 * @param {number} rowsPerChunk
 */
async function* makeCsvChunks(rowCount, rowsPerChunk) {
  yield "A,B\n";
  let i = 0;
  while (i < rowCount) {
    const end = Math.min(rowCount, i + rowsPerChunk);
    let chunk = "";
    for (; i < end; i++) {
      chunk += `${i},${i * 2}\n`;
    }
    yield chunk;
  }
}

const ROWS = 250_000;
const BATCH_SIZE = 2048;

const query = {
  id: "q_stream_bench",
  name: "Streaming bench",
  source: { type: "csv", path: "/tmp/bench.csv", options: { hasHeaders: true } },
  steps: [
    { id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "A", operator: "greaterThan", value: 100 } } },
    { id: "s_add", name: "Add", operation: { type: "addColumn", name: "C", formula: "=[B] * 2" } },
    { id: "s_take", name: "Take", operation: { type: "take", count: 100_000 } },
  ],
};

async function benchMaterialize() {
  const csvText = makeCsvText(ROWS);
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => csvText,
    },
  });

  const before = mem();
  const start = performance.now();
  const table = await engine.executeQuery(query, {}, {});
  const end = performance.now();

  // Touch the output so work isn't optimized away.
  void table.getCell(0, 0);

  console.log(`materialize: ${fmtMs(end - start)} rowsOut=${table.rowCount} heap=${before} -> ${mem()}`);
}

async function benchStreaming() {
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
    onBatch: (batch) => {
      rows += batch.values.length;
    },
  });

  const end = performance.now();
  console.log(`streaming:   ${fmtMs(end - start)} rowsOut=${rows} heap=${before} -> ${mem()}`);
}

console.log("Power Query streaming benchmark (JS, single-threaded)");
console.log(`Node ${process.version}`);
console.log(`rows=${ROWS.toLocaleString()} batchSize=${BATCH_SIZE}`);
console.log("");

await benchMaterialize();
await benchStreaming();

