import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import "fake-indexeddb/auto";
import { arrowTableFromColumns, arrowTableToParquet } from "../../../data-io/src/index.js";

import { CacheManager } from "../../src/cache/cache.js";
import { FileSystemCacheStore } from "../../src/cache/filesystem.js";
import { IndexedDBCacheStore } from "../../src/cache/indexeddb.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { ArrowTableAdapter } from "../../src/arrowTable.js";
import { HttpConnector } from "../../src/connectors/http.js";
import { QueryEngine } from "../../src/engine.js";
import { stableStringify } from "../../src/cache/key.js";
import { DataTable } from "../../src/table.js";

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
  assert.deepEqual(first.meta.outputSchema.columns, first.table.columns);

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read the source");
  assert.deepEqual(second.meta.outputSchema.columns, second.table.columns);

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

test("QueryEngine: reads legacy v1 DataTable cache entries", async () => {
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
    id: "q_sales_legacy_cache",
    name: "Sales",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [],
  };

  const cacheKey = await engine.getCacheKey(query, {}, {});
  assert.ok(cacheKey);

  const serializedTable = {
    columns: [
      { name: "Region", type: "string" },
      { name: "Sales", type: "number" },
    ],
    rows: [
      ["East", 100],
      ["West", 200],
    ],
  };

  await cache.set(cacheKey, {
    version: 1,
    table: serializedTable,
    meta: {
      queryId: query.id,
      refreshedAtMs: 0,
      sources: [],
      outputSchema: { columns: serializedTable.columns, inferred: true },
      outputRowCount: serializedTable.rows.length,
    },
  });

  // Legacy v1 cache entries predate source-state validation metadata. Disable validation so we can
  // prove the engine still understands the v1 payload shape.
  const result = await engine.executeQueryWithMeta(query, {}, { cache: { validation: "none" } });
  assert.equal(result.meta.cache?.hit, true);
  assert.equal(readCount, 0, "cache hit should not re-read the source");
  assert.deepEqual(result.meta.outputSchema.columns, result.table.columns);
  assert.deepEqual(result.table.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["West", 200],
  ]);
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
  assert.deepEqual(first.meta.outputSchema.columns, first.table.columns);
  const firstGrid = first.table.toGrid();

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read the Parquet bytes");
  assert.ok(second.table instanceof ArrowTableAdapter);
  assert.deepEqual(second.meta.outputSchema.columns, second.table.columns);
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
    assert.deepEqual(first.meta.outputSchema.columns, first.table.columns);
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
    assert.deepEqual(second.meta.outputSchema.columns, second.table.columns);
    assert.deepEqual(second.table.toGrid(), grid);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("FileSystemCacheStore: persists DataTable cache entries and preserves Date values", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-data-"));

  try {
    const query = {
      id: "q_range_date_fs_cache",
      name: "Range date fs cache",
      source: {
        type: "range",
        range: {
          hasHeaders: true,
          values: [
            ["id", "occurredAt"],
            [1, new Date("2024-01-01T00:00:00.000Z")],
            [2, null],
          ],
        },
      },
      steps: [],
    };

    const firstEngine = new QueryEngine({
      cache: new CacheManager({ store: new FileSystemCacheStore({ directory: cacheDir }) }),
    });

    const first = await firstEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(first.meta.cache?.hit, false);
    const firstGrid = first.table.toGrid();
    assert.ok(firstGrid[1][1] instanceof Date, "date cell should be materialized as a Date");

    const files = await readdir(cacheDir);
    assert.ok(files.some((name) => name.endsWith(".json")));
    assert.equal(
      files.some((name) => name.endsWith(".bin")),
      false,
      "DataTable cache should not create a .bin blob",
    );

    const secondEngine = new QueryEngine({
      cache: new CacheManager({ store: new FileSystemCacheStore({ directory: cacheDir }) }),
    });

    const second = await secondEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(second.meta.cache?.hit, true);
    const secondGrid = second.table.toGrid();
    assert.ok(secondGrid[1][1] instanceof Date, "date cell should survive cache roundtrip");
    assert.deepEqual(secondGrid, firstGrid);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("IndexedDBCacheStore: caches Arrow-backed Parquet results without re-reading the source", async () => {
  const dbName = `pq-cache-idb-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const store = new IndexedDBCacheStore({ dbName });
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

  const query = { id: "q_parquet_idb_cache", name: "Parquet idb cache", source: { type: "parquet", path: parquetPath }, steps: [] };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(readCount, 1);
  assert.ok(first.table instanceof ArrowTableAdapter);
  const grid = first.table.toGrid();

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(readCount, 1, "cache hit should not re-read Parquet bytes");
  assert.ok(second.table instanceof ArrowTableAdapter);
  assert.deepEqual(second.table.toGrid(), grid);

  const db = await store.open();
  db.close();

  await new Promise((resolve, reject) => {
    const req = indexedDB.deleteDatabase(dbName);
    req.onsuccess = () => resolve(undefined);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB deleteDatabase failed"));
    req.onblocked = () => resolve(undefined);
  });
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

test("QueryEngine: table sources incorporate host signatures into the cache key", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  const engine = new QueryEngine({ cache, defaultCacheTtlMs: 10 });

  const table = DataTable.fromGrid(
    [
      ["n"],
      [1],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const query = {
    id: "q_table",
    name: "TableSource",
    source: { type: "table", table: "T1" },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const contextV1 = { tables: { T1: table }, tableSignatures: { T1: 1 } };
  const first = await engine.executeQueryWithMeta(query, contextV1, {});
  assert.equal(first.meta.cache?.hit, false);

  const second = await engine.executeQueryWithMeta(query, contextV1, {});
  assert.equal(second.meta.cache?.hit, true);

  const contextV2 = { tables: { T1: table }, tableSignatures: { T1: 2 } };
  const third = await engine.executeQueryWithMeta(query, contextV2, {});
  assert.equal(third.meta.cache?.hit, false);
});

test("QueryEngine: invalidates file cache entries when mtime changes (within TTL)", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  let readCount = 0;
  let mtimeMs = 1;

  const engine = new QueryEngine({
    cache,
    defaultCacheTtlMs: 10_000,
    fileAdapter: {
      readText: async () => {
        readCount += 1;
        return ["Region,Sales", "East,100", "West,200"].join("\n");
      },
      stat: async () => ({ mtimeMs }),
    },
  });

  const query = {
    id: "q_file_mtime",
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
  assert.equal(readCount, 1);

  // Change the file state without expiring the entry.
  mtimeMs = 2;
  const third = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(third.meta.cache?.hit, false);
  assert.equal(readCount, 2);
});

test("QueryEngine: invalidates HTTP cache entries when ETag/Last-Modified changes (within TTL)", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  let etag = '"v1"';
  let lastModified = "Mon, 01 Jan 2024 00:00:00 GMT";
  let getCount = 0;

  /**
   * @param {{
   *   status?: number;
   *   headers?: Record<string, string>;
   *   body?: string;
   * }} init
   */
  function makeResponse(init = {}) {
    const status = init.status ?? 200;
    const headers = new Map(Object.entries(init.headers ?? {}).map(([k, v]) => [k.toLowerCase(), v]));
    const body = init.body ?? "";
    return {
      ok: status >= 200 && status < 300,
      status,
      headers: {
        get(name) {
          return headers.get(String(name).toLowerCase()) ?? null;
        },
      },
      async text() {
        return body;
      },
      async json() {
        return JSON.parse(body);
      },
    };
  }

  /** @type {typeof fetch} */
  const fetchMock = async (_url, init) => {
    const method = String(init?.method ?? "GET").toUpperCase();
    if (method === "HEAD") {
      return makeResponse({ headers: { etag, "last-modified": lastModified } });
    }
    getCount += 1;
    return makeResponse({
      headers: { "content-type": "application/json" },
      body: JSON.stringify([{ id: 1 }]),
    });
  };

  const engine = new QueryEngine({
    cache,
    defaultCacheTtlMs: 10_000,
    connectors: { http: new HttpConnector({ fetch: fetchMock }) },
  });

  const query = {
    id: "q_http",
    name: "HTTP",
    source: { type: "api", url: "https://example.com/data", method: "GET", headers: {} },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(getCount, 1);

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(getCount, 1, "cache hit should not refetch the resource");

  etag = '"v2"';
  lastModified = "Tue, 02 Jan 2024 00:00:00 GMT";
  const third = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(third.meta.cache?.hit, false);
  assert.equal(getCount, 2, "changed source state should invalidate the cache entry");
});
