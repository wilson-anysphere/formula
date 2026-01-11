import assert from "node:assert/strict";
import { mkdtemp, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { FileSystemCacheStore } from "../../src/cache/filesystem.js";

/**
 * @param {number} seed
 */
function makeArrowCacheValue(seed) {
  return {
    version: 2,
    table: {
      kind: "arrow",
      format: "ipc",
      columns: [],
      bytes: new Uint8Array([seed, seed + 1, seed + 2]),
    },
    meta: { seed },
  };
}

test("FileSystemCacheStore quotas: evicted entries remove both .json and .bin", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-quota-fs-"));
  let now = 0;

  try {
    const store = new FileSystemCacheStore({ directory: cacheDir, now: () => now });
    const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 1 } });

    now = 0;
    await cache.set("k1", makeArrowCacheValue(1));
    now = 1;
    await cache.set("k2", makeArrowCacheValue(2));

    const files = await readdir(cacheDir);
    const jsonFiles = files.filter((name) => name.endsWith(".json"));
    const binFiles = files.filter((name) => name.endsWith(".bin"));

    assert.equal(jsonFiles.length, 1, "should keep exactly one JSON entry file");
    assert.equal(binFiles.length, 1, "Arrow entries should keep exactly one binary blob");
    assert.equal(
      jsonFiles[0].replace(/\.json$/, ""),
      binFiles[0].replace(/\.bin$/, ""),
      "remaining .json/.bin pair should share the same base filename",
    );
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("FileSystemCacheStore quotas: eviction order respects lastAccessMs (LRU)", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-quota-fs-access-"));
  let now = 0;

  try {
    const store = new FileSystemCacheStore({ directory: cacheDir, now: () => now });
    const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

    now = 0;
    await cache.set("k1", makeArrowCacheValue(10));
    now = 1;
    await cache.set("k2", makeArrowCacheValue(20));

    now = 2;
    assert.deepEqual(await cache.get("k1"), makeArrowCacheValue(10));

    now = 3;
    await cache.set("k3", makeArrowCacheValue(30));

    assert.deepEqual(await cache.get("k1"), makeArrowCacheValue(10), "recently accessed entry should be retained");
    assert.equal(await cache.get("k2"), null, "least-recently-used entry should be evicted");
    assert.deepEqual(await cache.get("k3"), makeArrowCacheValue(30));
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("FileSystemCacheStore quotas: evicts least-recently-used when maxBytes is exceeded", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-quota-fs-bytes-"));
  let now = 0;

  try {
    const store = new FileSystemCacheStore({ directory: cacheDir, now: () => now });
    const cache = new CacheManager({ store, now: () => now, limits: { maxBytes: 6_000 } });

    const large = { bytes: new Uint8Array(4_096).fill(1) };
    const large2 = { bytes: new Uint8Array(4_096).fill(2) };

    now = 0;
    await cache.set("k1", large);
    now = 1;
    await cache.set("k2", large2);

    assert.equal(await cache.get("k1"), null, "oldest entry should be evicted to satisfy maxBytes");
    assert.deepEqual(await cache.get("k2"), large2);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("FileSystemCacheStore quotas: expired entries are removed before LRU eviction", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-quota-fs-expiry-"));
  let now = 0;

  try {
    const store = new FileSystemCacheStore({ directory: cacheDir, now: () => now });
    const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

    now = 0;
    await cache.set("k1", makeArrowCacheValue(1), { ttlMs: 5 }); // expires at t=5
    now = 1;
    await cache.set("k2", makeArrowCacheValue(2));

    // Touch k1 so it would not be the LRU entry, then let it expire.
    now = 4;
    assert.deepEqual(await cache.get("k1"), makeArrowCacheValue(1));

    now = 6;
    await cache.set("k3", makeArrowCacheValue(3));

    assert.equal(await cache.get("k1"), null, "expired entries should be removed preferentially");
    assert.deepEqual(await cache.get("k2"), makeArrowCacheValue(2));
    assert.deepEqual(await cache.get("k3"), makeArrowCacheValue(3));
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});
