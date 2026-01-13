import assert from "node:assert/strict";
import test from "node:test";

import { JsonVectorStore } from "../src/store/jsonVectorStore.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

function createDeferred() {
  /** @type {(value?: void) => void} */
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

class ControlledBinaryStorage {
  constructor() {
    /** @type {Uint8Array | null} */
    this._data = null;
    /** @type {{ data: Uint8Array, deferred: ReturnType<typeof createDeferred> }[]} */
    this._saves = [];
    /** @type {Set<number>} */
    this._preReleased = new Set();
    this._releaseAll = false;
    /** @type {{ count: number, resolve: () => void }[]} */
    this._waiters = [];
  }

  async load() {
    return this._data ? new Uint8Array(this._data) : null;
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    const idx = this._saves.length + 1; // 1-based
    const deferred = createDeferred();
    const copy = new Uint8Array(data);
    this._saves.push({ data: copy, deferred });

    // Commit the write only when the async save is released.
    deferred.promise.then(
      () => {
        this._data = new Uint8Array(copy);
      },
      () => {},
    );

    for (let i = this._waiters.length - 1; i >= 0; i -= 1) {
      if (this._saves.length >= this._waiters[i].count) {
        this._waiters[i].resolve();
        this._waiters.splice(i, 1);
      }
    }

    if (this._releaseAll || this._preReleased.has(idx)) {
      deferred.resolve();
    }

    return deferred.promise;
  }

  /**
   * @param {number} count
   */
  async waitForSaves(count) {
    if (this._saves.length >= count) return;
    await new Promise((resolve) => {
      this._waiters.push({ count, resolve });
    });
  }

  /**
   * Release a specific `save()` call (1-based).
   * @param {number} index
   */
  release(index) {
    if (this._releaseAll) return;
    const entry = this._saves[index - 1];
    if (entry) {
      entry.deferred.resolve();
      return;
    }
    this._preReleased.add(index);
  }

  releaseAll() {
    this._releaseAll = true;
    for (const entry of this._saves) entry.deferred.resolve();
  }

  get data() {
    return this._data ? new Uint8Array(this._data) : null;
  }

  get saveCount() {
    return this._saves.length;
  }
}

test("JsonVectorStore serializes persistence writes to prevent lost updates", async () => {
  const storage = new ControlledBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 2, autoSave: true });
  await store.load();

  const p1 = store.upsert([{ id: "a", vector: [1, 0], metadata: { label: "A" } }]);
  const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { label: "B" } }]);

  // Ensure the first save is in-flight, then release the second save before the first.
  await storage.waitForSaves(1);
  storage.release(2);
  await Promise.resolve(); // allow the second upsert to enqueue its persist in buggy implementations
  storage.release(1);

  await Promise.all([p1, p2]);

  assert.equal(storage.saveCount, 2);
  const persisted = storage.data;
  assert.ok(persisted, "Expected JsonVectorStore to persist bytes");
  const parsed = JSON.parse(new TextDecoder().decode(persisted));
  const ids = parsed.records.map((r) => r.id).sort();
  assert.deepEqual(ids, ["a", "b"]);
});

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

test(
  "SqliteVectorStore serializes persistence writes to prevent lost updates",
  { skip: !sqlJsAvailable },
  async () => {
    const storage = new ControlledBinaryStorage();
    const store = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: true });

    const p1 = store.upsert([{ id: "a", vector: [1, 0], metadata: { label: "A" } }]);
    const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { label: "B" } }]);

    await storage.waitForSaves(1);
    storage.release(2);
    await Promise.resolve();
    storage.release(1);

    await Promise.all([p1, p2]);

    // We only care about controlling the first two persists for this test. Any
    // subsequent persists (e.g. close, reload) should complete immediately.
    storage.releaseAll();

    const persisted = storage.data;
    assert.ok(persisted, "Expected SqliteVectorStore to persist bytes");

    // Reload the database from the persisted bytes and verify it includes both
    // records, proving the last write wasn't lost.
    const reloaded = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: false });
    assert.ok(await reloaded.get("a"));
    assert.ok(await reloaded.get("b"));

    await store.close();
    await reloaded.close();
  },
);
