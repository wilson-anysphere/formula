import assert from "node:assert/strict";
import test from "node:test";

let indexedDbAvailable = true;
try {
  await import("fake-indexeddb/auto");
} catch {
  indexedDbAvailable = false;
}

import { IndexedDBCacheStore } from "../../src/cache/indexeddb.js";

test("IndexedDBCacheStore.pruneExpired removes expired entries", { skip: !indexedDbAvailable }, async () => {
  const dbName = `pq-cache-prune-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const store = new IndexedDBCacheStore({ dbName });

  await store.set("expired", { value: "x", createdAtMs: 0, expiresAtMs: 5 });
  await store.set("alive", { value: "y", createdAtMs: 0, expiresAtMs: 50 });

  await store.pruneExpired(10);

  assert.equal(await store.get("expired"), null);
  assert.ok(await store.get("alive"));

  // Close DB handle before deleting to keep fake-indexeddb happy.
  const db = await store.open();
  db.close();

  await new Promise((resolve, reject) => {
    const req = indexedDB.deleteDatabase(dbName);
    req.onsuccess = () => resolve(undefined);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB deleteDatabase failed"));
    req.onblocked = () => resolve(undefined);
  });
});
