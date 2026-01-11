import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { QueryEngine } from "../../src/engine.js";

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

