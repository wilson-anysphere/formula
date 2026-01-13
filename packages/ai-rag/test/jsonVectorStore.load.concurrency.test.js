import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

function defer() {
  /** @type {(value: any) => void} */
  let resolve;
  /** @type {(reason?: any) => void} */
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  if (!resolve || !reject) {
    throw new Error("Deferred promise failed to initialize");
  }
  return { promise, resolve, reject };
}

test("JsonVectorStore.load waits for an in-flight load when called concurrently", async () => {
  // Seed persisted bytes containing record "x".
  const seededStorage = new InMemoryBinaryStorage();
  const seed = new JsonVectorStore({ storage: seededStorage, dimension: 2, autoSave: true });
  await seed.upsert([{ id: "x", vector: [1, 0], metadata: { label: "X" } }]);
  const seededBytes = await seededStorage.load();
  assert.ok(seededBytes, "Expected seed JsonVectorStore to persist bytes");

  // Create a storage implementation whose load() resolves only when manually released.
  const loadDeferred = defer();
  const storage = {
    async load() {
      return await loadDeferred.promise;
    },
    async save() {},
  };

  const store = new JsonVectorStore({ storage, dimension: 2, autoSave: false });

  // Start an upsert that triggers the initial load but do not await it. Immediately start a delete.
  // Without an internal "load mutex", the delete could run before load finishes and miss deleting "x",
  // allowing the later load to re-introduce the record.
  const upsertPromise = store.upsert([{ id: "a", vector: [0, 1], metadata: {} }]);
  const deletePromise = store.delete(["x"]);

  // Release the load payload after both operations have begun.
  loadDeferred.resolve(seededBytes);

  await Promise.all([upsertPromise, deletePromise]);

  assert.equal(await store.get("x"), null);
  assert.ok(await store.get("a"));
});

