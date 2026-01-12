import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { arrowTableFromColumns, arrowTableToParquet } from "../../data-io/src/index.js";

import { ArrowTableAdapter } from "../src/arrowTable.js";
import { QueryEngine } from "../src/engine.js";
import { DataTable } from "../src/table.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const PARQUET_FIXTURE = path.join(__dirname, "..", "..", "data-io", "test", "fixtures", "simple.parquet");

const EXPECTED_GRID = [
  ["id", "name", "active", "score"],
  [1, "Alice", true, 1.5],
  [2, "Bob", false, 2.25],
  [3, "Carla", true, 3.75],
];

function collectBatches(batches) {
  const grid = [];
  for (const batch of batches) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }
  return grid;
}

test("Parquet sources fall back to readParquetTable when Arrow/Parquet support is unavailable", async () => {
  let parquetAvailable = true;
  try {
    await arrowTableToParquet(arrowTableFromColumns({ __probe: new Int32Array([1]) }));
  } catch {
    parquetAvailable = false;
  }

  let readParquetCalled = false;

  const engine = new QueryEngine({
    fileAdapter: {
      openFile: async (p) => new Blob([new Uint8Array(await readFile(p))]),
      readParquetTable: async () => {
        readParquetCalled = true;
        return DataTable.fromGrid(EXPECTED_GRID, { hasHeaders: true, inferTypes: true });
      },
    },
  });

  const result = await engine.executeQuery(
    {
      id: "q_parquet_fallback",
      name: "Parquet Fallback",
      source: { type: "parquet", path: PARQUET_FIXTURE, options: { batchSize: 2 } },
      steps: [],
    },
    {},
    {},
  );

  assert.deepEqual(result.toGrid(), EXPECTED_GRID);

  if (result instanceof ArrowTableAdapter) {
    assert.equal(readParquetCalled, false);
    assert.equal(parquetAvailable, true);
  } else {
    assert.ok(result instanceof DataTable);
    assert.equal(readParquetCalled, true);
    assert.equal(parquetAvailable, false);
  }
});

test("executeQueryStreamingNonMaterializing can stream Parquet via readParquetTable fallback", async () => {
  let parquetAvailable = true;
  try {
    await arrowTableToParquet(arrowTableFromColumns({ __probe: new Int32Array([1]) }));
  } catch {
    parquetAvailable = false;
  }

  let readParquetCalled = false;

  const engine = new QueryEngine({
    fileAdapter: {
      openFile: async (p) => new Blob([new Uint8Array(await readFile(p))]),
      readParquetTable: async () => {
        readParquetCalled = true;
        return DataTable.fromGrid(EXPECTED_GRID, { hasHeaders: true, inferTypes: true });
      },
    },
  });

  const batches = [];
  await engine.executeQueryStreaming(
    {
      id: "q_parquet_fallback_stream",
      name: "Parquet Fallback Stream",
      source: { type: "parquet", path: PARQUET_FIXTURE, options: { batchSize: 2 } },
      steps: [],
    },
    {},
    {
      batchSize: 2,
      materialize: false,
      onBatch: (batch) => batches.push(batch),
    },
  );

  assert.deepEqual(collectBatches(batches), EXPECTED_GRID);
  assert.equal(readParquetCalled, !parquetAvailable);
});

test("Parquet sources fall back to readParquetTable when only readBinaryStream is available", async () => {
  let parquetAvailable = true;
  try {
    await arrowTableToParquet(arrowTableFromColumns({ __probe: new Int32Array([1]) }));
  } catch {
    parquetAvailable = false;
  }

  let readParquetCalled = false;

  const engine = new QueryEngine({
    fileAdapter: {
      readBinaryStream: async function* (p) {
        const bytes = new Uint8Array(await readFile(p));
        // Emit at least two chunks so we cover the streaming code path.
        const mid = Math.floor(bytes.length / 2);
        yield bytes.slice(0, mid);
        yield bytes.slice(mid);
      },
      readParquetTable: async () => {
        readParquetCalled = true;
        return DataTable.fromGrid(EXPECTED_GRID, { hasHeaders: true, inferTypes: true });
      },
    },
  });

  const result = await engine.executeQuery(
    {
      id: "q_parquet_fallback_stream_bytes",
      name: "Parquet Fallback Stream Bytes",
      source: { type: "parquet", path: PARQUET_FIXTURE, options: { batchSize: 2 } },
      steps: [],
    },
    {},
    {},
  );

  assert.deepEqual(result.toGrid(), EXPECTED_GRID);

  if (result instanceof ArrowTableAdapter) {
    assert.equal(readParquetCalled, false);
    assert.equal(parquetAvailable, true);
  } else {
    assert.ok(result instanceof DataTable);
    assert.equal(readParquetCalled, true);
    assert.equal(parquetAvailable, false);
  }
});
