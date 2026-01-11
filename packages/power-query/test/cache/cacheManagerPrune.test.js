import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";

test("CacheManager.prune calls pruneExpired before store.prune (with configured limits)", async () => {
  /** @type {Array<any>} */
  const calls = [];

  /** @type {import("../../src/cache/cache.js").CacheStore} */
  const store = {
    get: async () => null,
    set: async () => {},
    delete: async () => {},
    pruneExpired: async (nowMs) => {
      calls.push(["expired", nowMs]);
    },
    prune: async ({ nowMs, maxEntries, maxBytes }) => {
      calls.push(["prune", nowMs, maxEntries, maxBytes]);
    },
  };

  const cache = new CacheManager({ store, now: () => 123, limits: { maxEntries: 10, maxBytes: 20 } });
  await cache.prune();

  assert.deepEqual(calls, [
    ["expired", 123],
    ["prune", 123, 10, 20],
  ]);
});

test("CacheManager.prune forwards explicit limit overrides", async () => {
  /** @type {Array<any>} */
  const calls = [];

  /** @type {import("../../src/cache/cache.js").CacheStore} */
  const store = {
    get: async () => null,
    set: async () => {},
    delete: async () => {},
    pruneExpired: async (nowMs) => {
      calls.push(["expired", nowMs]);
    },
    prune: async ({ nowMs, maxEntries, maxBytes }) => {
      calls.push(["prune", nowMs, maxEntries, maxBytes]);
    },
  };

  const cache = new CacheManager({ store, now: () => 456, limits: { maxEntries: 99, maxBytes: 99 } });
  await cache.prune({ maxEntries: 1 });

  assert.deepEqual(calls, [
    ["expired", 456],
    ["prune", 456, 1, undefined],
  ]);
});

test("CacheManager.prune without limits only calls pruneExpired", async () => {
  /** @type {Array<any>} */
  const calls = [];

  /** @type {import("../../src/cache/cache.js").CacheStore} */
  const store = {
    get: async () => null,
    set: async () => {},
    delete: async () => {},
    pruneExpired: async (nowMs) => {
      calls.push(["expired", nowMs]);
    },
    prune: async () => {
      calls.push(["prune"]);
    },
  };

  const cache = new CacheManager({ store, now: () => 789 });
  await cache.prune();

  assert.deepEqual(calls, [["expired", 789]]);
});

