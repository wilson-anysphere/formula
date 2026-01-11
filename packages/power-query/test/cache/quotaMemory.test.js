import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";

test("MemoryCacheStore quotas: evicts least-recently-used when maxEntries is exceeded", async () => {
  let now = 0;
  const store = new MemoryCacheStore({ now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

  now = 0;
  await cache.set("k1", { v: 1 });
  now = 1;
  await cache.set("k2", { v: 2 });
  now = 2;
  await cache.set("k3", { v: 3 });

  assert.equal(await cache.get("k1"), null, "oldest entry should be evicted");
  assert.deepEqual(await cache.get("k2"), { v: 2 });
  assert.deepEqual(await cache.get("k3"), { v: 3 });
});

test("MemoryCacheStore quotas: get updates access time and affects eviction order", async () => {
  let now = 0;
  const store = new MemoryCacheStore({ now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

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
});

test("MemoryCacheStore quotas: expired entries are removed before LRU eviction", async () => {
  let now = 0;
  const store = new MemoryCacheStore({ now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxEntries: 2 } });

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
});

test("MemoryCacheStore quotas: evicts least-recently-used when maxBytes is exceeded", async () => {
  let now = 0;
  const store = new MemoryCacheStore({ now: () => now });
  const cache = new CacheManager({ store, now: () => now, limits: { maxBytes: 6_000 } });

  const large = { bytes: new Uint8Array(4_096).fill(1) };
  const large2 = { bytes: new Uint8Array(4_096).fill(2) };

  now = 0;
  await cache.set("k1", large);
  now = 1;
  await cache.set("k2", large2);

  assert.equal(await cache.get("k1"), null, "oldest entry should be evicted to satisfy maxBytes");
  assert.deepEqual(await cache.get("k2"), large2);
});

