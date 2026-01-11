import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { arrowTableFromColumns, arrowTableToParquet } from "../../../data-io/src/index.js";

import { CacheManager } from "../../src/cache/cache.js";
import { FileSystemCacheStore } from "../../src/cache/filesystem.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { ArrowTableAdapter } from "../../src/arrowTable.js";
import { QueryEngine } from "../../src/engine.js";
import { stableStringify } from "../../src/cache/key.js";

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

test("QueryEngine: caches by source + query + credentialId and still checks permissions", async () => {
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
      return { credentialId: "cred-1", getSecret: async () => ({ token: "secret" }) };
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

test("QueryEngine: cache key varies by credentialId and does not embed raw secrets", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  const secret = "super-secret-token";
  let currentId = "cred-a";

  const engine = new QueryEngine({
    cache,
    fileAdapter: {
      readText: async () => ["Region,Sales", "East,100"].join("\n"),
    },
    onPermissionRequest: async () => true,
    onCredentialRequest: async () => {
      return { credentialId: currentId, getSecret: async () => ({ token: secret }) };
    },
  });

  const query = {
    id: "q_sales",
    name: "Sales",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const keyA = await engine.getCacheKey(query, {}, {});
  currentId = "cred-b";
  const keyB = await engine.getCacheKey(query, {}, {});
  assert.notEqual(keyA, keyB);

  const state = { credentialCache: new Map(), permissionCache: new Map(), now: () => Date.now() };
  const signature = await engine.buildQuerySignature(query, {}, {}, state, new Set([query.id]));
  assert.equal(typeof signature, "object");
  assert.equal(stableStringify(signature).includes(secret), false);
});

test("QueryEngine: corrupted cache entry is treated as a miss and refreshed", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store });

  let readCount = 0;
  const engine = new QueryEngine({
    cache,
    fileAdapter: {
      readText: async () => {
        readCount += 1;
        return ["Region,Sales", "East,100", "West,200"].join("\n");
      },
    },
  });

  const query = {
    id: "q_sales_corrupt_cache",
    name: "Sales",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [],
  };

  const cacheKey = await engine.getCacheKey(query, {}, {});
  assert.ok(cacheKey, "cache key should be computed when cache is enabled");

  // Simulate a corrupted cache entry (e.g. partial write / old format).
  await cache.set(cacheKey, { version: 2, table: null, meta: null });

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false, "corrupted cache should not be treated as a hit");
  assert.equal(readCount, 1);

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true, "engine should refresh and then hit on subsequent executions");
  assert.equal(readCount, 1, "cache hit should not re-read the source");
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

  const query = {
    id: "q_parquet_cache",
    name: "Parquet cache",
    source: { type: "parquet", path: parquetPath },
    steps: [],
  };

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

test("FileSystemCacheStore: persists Arrow cache blobs and avoids re-reading Parquet on cache hit", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-arrow-"));

  try {
    const parquetPath = path.join(__dirname, "..", "..", "..", "data-io", "test", "fixtures", "simple.parquet");

    let readCount = 0;
    const firstEngine = new QueryEngine({
      cache: new CacheManager({ store: new FileSystemCacheStore({ directory: cacheDir }) }),
      fileAdapter: {
        readBinary: async (p) => {
          readCount += 1;
          return new Uint8Array(await readFile(p));
        },
      },
    });

    const query = { id: "q_parquet_fs_cache", name: "Parquet fs cache", source: { type: "parquet", path: parquetPath }, steps: [] };

    const first = await firstEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(first.meta.cache?.hit, false);
    assert.equal(readCount, 1);
    assert.ok(first.table instanceof ArrowTableAdapter);
    const grid = first.table.toGrid();

    const files = await readdir(cacheDir);
    assert.ok(files.some((name) => name.endsWith(".bin")), "filesystem cache should create a .bin blob for Arrow IPC bytes");

    let secondReadCount = 0;
    const secondEngine = new QueryEngine({
      cache: new CacheManager({ store: new FileSystemCacheStore({ directory: cacheDir }) }),
      fileAdapter: {
        readBinary: async (p) => {
          secondReadCount += 1;
          return new Uint8Array(await readFile(p));
        },
      },
    });

    const second = await secondEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(second.meta.cache?.hit, true);
    assert.equal(secondReadCount, 0, "cache hit should not re-read Parquet bytes");
    assert.ok(second.table instanceof ArrowTableAdapter);
    assert.deepEqual(second.table.toGrid(), grid);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
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
