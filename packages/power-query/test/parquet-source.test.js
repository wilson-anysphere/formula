import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { arrowTableFromColumns, arrowTableToParquet } from "../../data-io/src/index.js";

import { QueryEngine } from "../src/engine.js";
import { ArrowTableAdapter } from "../src/arrowTable.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

let parquetAvailable = true;
try {
  // Validate Parquet support is actually usable via the data-io helpers (pnpm workspaces
  // don't necessarily hoist `apache-arrow`/`parquet-wasm` to the repo root).
  await arrowTableToParquet(arrowTableFromColumns({ __probe: new Int32Array([1]) }));
} catch {
  parquetAvailable = false;
}

test("parquet query source loads into Arrow and runs transforms without materializing row arrays", { skip: !parquetAvailable }, async () => {
  const parquetPath = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  const engine = new QueryEngine({
    fileAdapter: {
      readBinary: async (p) => new Uint8Array(await readFile(p)),
    },
  });

  const query = {
    id: "q_parquet",
    name: "Parquet",
    source: { type: "parquet", path: parquetPath, options: { batchSize: 2 } },
    steps: [
      {
        id: "s_filter",
        name: "Active only",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "active", operator: "equals", value: true } },
      },
      { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name", "score"] } },
      { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "score", direction: "descending" }] } },
    ],
  };

  const result = await engine.executeQuery(query, {}, {});
  assert.ok(result instanceof ArrowTableAdapter);
  assert.deepEqual(result.toGrid(), [
    ["id", "name", "score"],
    [3, "Carla", 3.75],
    [1, "Alice", 1.5],
  ]);
});

test("parquet query source supports readBinaryStream for Arrow-backed execution", { skip: !parquetAvailable }, async () => {
  const parquetPath = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  const engine = new QueryEngine({
    fileAdapter: {
      readBinaryStream: async function* (p) {
        const bytes = new Uint8Array(await readFile(p));
        const chunkSize = 128;
        for (let offset = 0; offset < bytes.length; offset += chunkSize) {
          yield bytes.subarray(offset, Math.min(bytes.length, offset + chunkSize));
        }
      },
    },
  });

  const query = {
    id: "q_parquet_binary_stream",
    name: "Parquet (binary stream)",
    source: { type: "parquet", path: parquetPath, options: { batchSize: 2 } },
    steps: [
      {
        id: "s_filter",
        name: "Active only",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "active", operator: "equals", value: true } },
      },
      { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name", "score"] } },
      { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "score", direction: "descending" }] } },
    ],
  };

  const result = await engine.executeQuery(query, {}, {});
  assert.ok(result instanceof ArrowTableAdapter);
  assert.deepEqual(result.toGrid(), [
    ["id", "name", "score"],
    [3, "Carla", 3.75],
    [1, "Alice", 1.5],
  ]);
});

test("parquet query source supports executeQueryStreaming", { skip: !parquetAvailable }, async () => {
  const parquetPath = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  const engine = new QueryEngine({
    fileAdapter: {
      readBinary: async (p) => new Uint8Array(await readFile(p)),
    },
  });

  const query = {
    id: "q_parquet_stream",
    name: "Parquet stream",
    source: { type: "parquet", path: parquetPath, options: { batchSize: 2 } },
    steps: [{ id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name"] } }],
  };

  const grid = [];
  await engine.executeQueryStreaming(query, {}, {
    batchSize: 2,
    onBatch: (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        grid[batch.rowOffset + i] = batch.values[i];
      }
    },
  });

  assert.deepEqual(grid[0], ["id", "name"]);
  assert.deepEqual(grid[1], [1, "Alice"]);
  assert.deepEqual(grid[3], [3, "Carla"]);
});

test("parquet query source supports non-materializing executeQueryStreaming via readBinary", { skip: !parquetAvailable }, async () => {
  const parquetPath = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readBinary: async (p) => new Uint8Array(await readFile(p)),
    },
  });
  // Ensure we don't fall back to `executeQuery()` (which would materialize the parquet source).
  engineStreaming.executeQuery = async () => {
    throw new Error("executeQuery should not be called in non-materializing streaming mode");
  };

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readBinary: async (p) => new Uint8Array(await readFile(p)),
    },
  });

  const query = {
    id: "q_parquet_stream_non_materialize",
    name: "Parquet stream non-materialize",
    source: { type: "parquet", path: parquetPath, options: { batchSize: 2 } },
    steps: [{ id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name"] } }],
  };

  const grid = [];
  await engineStreaming.executeQueryStreaming(query, {}, {
    batchSize: 2,
    materialize: false,
    onBatch: (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        grid[batch.rowOffset + i] = batch.values[i];
      }
    },
  });

  const expected = (await engineMaterialized.executeQuery(query, {}, {})).toGrid();
  assert.deepEqual(grid, expected);
});

test("parquet query source supports non-materializing executeQueryStreaming via readBinaryStream", { skip: !parquetAvailable }, async () => {
  const parquetPath = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readBinaryStream: async function* (p) {
        const bytes = new Uint8Array(await readFile(p));
        // Yield small chunks to ensure the adapter path works with incremental reads.
        const chunkSize = 128;
        for (let offset = 0; offset < bytes.length; offset += chunkSize) {
          yield bytes.subarray(offset, Math.min(bytes.length, offset + chunkSize));
        }
      },
    },
  });
  engineStreaming.executeQuery = async () => {
    throw new Error("executeQuery should not be called in non-materializing streaming mode");
  };

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readBinary: async (p) => new Uint8Array(await readFile(p)),
    },
  });

  const query = {
    id: "q_parquet_stream_non_materialize_binary_stream",
    name: "Parquet stream non-materialize (binary stream)",
    source: { type: "parquet", path: parquetPath, options: { batchSize: 2 } },
    steps: [{ id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name"] } }],
  };

  const grid = [];
  await engineStreaming.executeQueryStreaming(query, {}, {
    batchSize: 2,
    materialize: false,
    onBatch: (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        grid[batch.rowOffset + i] = batch.values[i];
      }
    },
  });

  const expected = (await engineMaterialized.executeQuery(query, {}, {})).toGrid();
  assert.deepEqual(grid, expected);
});
