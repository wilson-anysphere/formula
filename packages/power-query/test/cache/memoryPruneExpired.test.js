import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";

test("MemoryCacheStore.pruneExpired removes expired entries", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  await cache.set("expired", { value: 1 }, { ttlMs: 10 });
  await cache.set("alive", { value: 2 }, { ttlMs: 100 });

  now = 20;
  await cache.pruneExpired();

  assert.equal(await cache.get("expired"), null);
  assert.deepEqual(await cache.get("alive"), { value: 2 });
});

