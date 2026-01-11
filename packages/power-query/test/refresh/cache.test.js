import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { arrowTableFromColumns, arrowTableToParquet } from "../../../data-io/src/index.js";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { ArrowTableAdapter } from "../../src/arrowTable.js";
import { QueryEngine } from "../../src/engine.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("CacheManager: hit/miss, TTL, manual invalidation", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  await cache.set("k1", { value: 1 }, { ttlMs: 10 });
  assert.deepEqual(await cache.get("k1"), { value: 1 });

  now = 11;
  assert.equal(await cache.get("k1"), null);

  await cache.set("k1", { value: 2 });
  assert.deepEqual(await cache.get("k1"), { value: 2 });
  await cache.delete("k1");
  assert.equal(await cache.get("k1"), null);
});

test("QueryEngine: caches by source + query + credentials hash and still checks permissions", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  let readCount = 0;
  let permissionCount = 0;
  let credentialCount = 0;

  const engine = new QueryEngine({
    cache,
    defaultCacheTtlMs: 10,
    fileAdapter: {
      readText: async () => {
        readCount += 1;
        return ["Region,Sales", "East,100", "West,200"].join("\n");
      },
    },
    onPermissionRequest: async () => {
      permissionCount += 1;
      return true;
    },
    onCredentialRequest: async () => {
      credentialCount += 1;
      return { token: "secret" };
    },
  });

  const query = {
    id: "q_sales",
    name: "Sales",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(readCount, 1);

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read the source");

  assert.equal(permissionCount, 2, "permissions should be checked even on cache hit");
  assert.equal(credentialCount, 2, "credentials are part of the cache key and should be requested per execution");

  now = 11;
  const third = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(third.meta.cache?.hit, false, "expired entry should force a refresh");
  assert.equal(readCount, 2, "expired entry should re-read the source");

  await engine.invalidateQueryCache(query, {}, {});
  const fourth = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(fourth.meta.cache?.hit, false);
  assert.equal(readCount, 3, "manual invalidation should force a refresh");
});

test("QueryEngine: caches Arrow-backed Parquet results and avoids re-reading the source", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store });

  const parquetPath = path.join(__dirname, "..", "..", "..", "data-io", "test", "fixtures", "simple.parquet");

  let readCount = 0;
  const engine = new QueryEngine({
    cache,
    fileAdapter: {
      readBinary: async (p) => {
        readCount += 1;
        return new Uint8Array(await readFile(p));
      },
    },
  });

  const query = { id: "q_parquet_cache", name: "Parquet cache", source: { type: "parquet", path: parquetPath }, steps: [] };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(readCount, 1);
  assert.ok(first.table instanceof ArrowTableAdapter);
  const firstGrid = first.table.toGrid();

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read the Parquet bytes");
  assert.ok(second.table instanceof ArrowTableAdapter);
  assert.deepEqual(second.table.toGrid(), firstGrid);
});

test("QueryEngine: Arrow cache roundtrip preserves date columns", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store });

  const inputTable = arrowTableFromColumns({
    id: new Int32Array([1, 2, 3]),
    occurredAt: [new Date("2024-01-01T00:00:00.000Z"), null, new Date("2024-01-03T12:34:56.000Z")],
    value: new Float64Array([1.25, 2.5, 3.75]),
  });
  const parquetBytes = await arrowTableToParquet(inputTable);

  let readCount = 0;
  const engine = new QueryEngine({
    cache,
    fileAdapter: {
      readBinary: async () => {
        readCount += 1;
        return parquetBytes;
      },
    },
  });

  const query = {
    id: "q_parquet_date_cache",
    name: "Parquet date cache",
    source: { type: "parquet", path: "/tmp/date-cache.parquet" },
    steps: [],
  };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(readCount, 1);
  assert.ok(first.table instanceof ArrowTableAdapter);
  const firstGrid = first.table.toGrid();
  assert.ok(firstGrid[1][1] instanceof Date, "date column should materialize as a Date");

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read the Parquet bytes");
  assert.ok(second.table instanceof ArrowTableAdapter);
  assert.deepEqual(second.table.toGrid(), firstGrid);
});
