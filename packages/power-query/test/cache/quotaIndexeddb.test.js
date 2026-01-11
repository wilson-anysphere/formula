import assert from "node:assert/strict";
import test from "node:test";

let indexedDbAvailable = true;
try {
  await import("fake-indexeddb/auto");
} catch {
  indexedDbAvailable = false;
}

import { CacheManager } from "../../src/cache/cache.js";
import { IndexedDBCacheStore } from "../../src/cache/indexeddb.js";

/**
 * @param {string} dbName
 */
async function deleteDatabase(dbName) {
  await new Promise((resolve, reject) => {
    const req = indexedDB.deleteDatabase(dbName);
    req.onsuccess = () => resolve(undefined);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB deleteDatabase failed"));
    req.onblocked = () => resolve(undefined);
  });
}

/**
 * @param {IndexedDBCacheStore} store
 * @param {string} dbName
 */
async function closeAndDelete(store, dbName) {
  try {
    const db = await store.open();
    db.close();
  } catch {
    // ignore
  }
  await deleteDatabase(dbName);
}

test(
  "IndexedDBCacheStore quotas: evicts least-recently-used when maxEntries is exceeded",
  { skip: !indexedDbAvailable },
  async () => {
  let now = 0;
  const dbName = `pq-cache-idb-quota-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const store = new IndexedDBCacheStore({ dbName, now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

  try {
    now = 0;
    await cache.set("k1", { v: 1 });
    now = 1;
    await cache.set("k2", { v: 2 });
    now = 2;
    await cache.set("k3", { v: 3 });

    assert.equal(await cache.get("k1"), null, "oldest entry should be evicted");
    assert.deepEqual(await cache.get("k2"), { v: 2 });
    assert.deepEqual(await cache.get("k3"), { v: 3 });
  } finally {
    await closeAndDelete(store, dbName);
  }
  },
);

test(
  "IndexedDBCacheStore quotas: get updates access time and affects eviction order",
  { skip: !indexedDbAvailable },
  async () => {
  let now = 0;
  const dbName = `pq-cache-idb-quota-access-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const store = new IndexedDBCacheStore({ dbName, now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

  try {
    now = 0;
    await cache.set("k1", { v: 1 });
    now = 1;
    await cache.set("k2", { v: 2 });

    now = 2;
    assert.deepEqual(await cache.get("k1"), { v: 1 });

    now = 3;
    await cache.set("k3", { v: 3 });

    assert.deepEqual(await cache.get("k1"), { v: 1 }, "recently accessed key should be retained");
    assert.equal(await cache.get("k2"), null, "least-recently-used key should be evicted");
    assert.deepEqual(await cache.get("k3"), { v: 3 });
  } finally {
    await closeAndDelete(store, dbName);
  }
  },
);

test(
  "IndexedDBCacheStore quotas: expired entries are removed before LRU eviction",
  { skip: !indexedDbAvailable },
  async () => {
  let now = 0;
  const dbName = `pq-cache-idb-quota-expiry-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const store = new IndexedDBCacheStore({ dbName, now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

  try {
    now = 0;
    await cache.set("k1", { v: 1 }, { ttlMs: 5 }); // expires at t=5
    now = 1;
    await cache.set("k2", { v: 2 });

    // Touch k1 so it is not the LRU entry, then let it expire.
    now = 4;
    assert.deepEqual(await cache.get("k1"), { v: 1 });

    // Setting k3 forces pruning; k1 is expired and should be deleted first, leaving k2 + k3.
    now = 6;
    await cache.set("k3", { v: 3 });

    assert.equal(await cache.get("k1"), null, "expired entries should be removed preferentially");
    assert.deepEqual(await cache.get("k2"), { v: 2 }, "non-expired entry should be retained");
    assert.deepEqual(await cache.get("k3"), { v: 3 });
  } finally {
    await closeAndDelete(store, dbName);
  }
  },
);
