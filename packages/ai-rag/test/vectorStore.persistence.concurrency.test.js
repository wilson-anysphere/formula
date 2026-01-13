import assert from "node:assert/strict";
import test from "node:test";

import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

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

  // Start the first upsert but do not await it. Wait until its persist has
  // actually reached `storage.save()` so we know the first snapshot is taken.
  const p1 = store.upsert([{ id: "a", vector: [1, 0], metadata: { label: "A" } }]);
  await storage.waitForSaves(1);

  // Start a second upsert while the first persist is still in-flight.
  const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { label: "B" } }]);

  // Simulate out-of-order completion: allow "save #2" to complete before "save #1".
  // In fixed implementations, save #2 won't even be invoked until save #1 completes,
  // but pre-releasing it keeps the test deterministic.
  storage.release(2);

  // Give the second upsert a chance to reach `storage.save()` in buggy
  // implementations where saves are not serialized.
  for (let i = 0; i < 25 && storage.saveCount < 2; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await Promise.resolve();
  }

  storage.release(1);

  await Promise.all([p1, p2]);

  assert.equal(storage.saveCount, 2);
  const persisted = storage.data;
  assert.ok(persisted, "Expected JsonVectorStore to persist bytes");
  const parsed = JSON.parse(new TextDecoder().decode(persisted));
  const ids = parsed.records.map((r) => r.id).sort();
  assert.deepEqual(ids, ["a", "b"]);
});

test("JsonVectorStore serializes deleteWorkbook persistence writes to prevent lost updates", async () => {
  const storage = new ControlledBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 2, autoSave: true });
  await store.load();

  // Pre-release the initial seed persist so we can focus on the overlapping
  // persists triggered by deleteWorkbook + upsert below.
  storage.release(1);
  await store.upsert([
    { id: "a", vector: [1, 0], metadata: { workbookId: "wb1" } },
    { id: "b", vector: [0, 1], metadata: { workbookId: "wb2" } },
  ]);

  const p1 = store.deleteWorkbook("wb1");
  // Wait until deleteWorkbook has actually invoked `storage.save()` so its snapshot
  // is taken before we start the concurrent upsert.
  await storage.waitForSaves(2);

  const p2 = store.upsert([{ id: "c", vector: [1, 1], metadata: { workbookId: "wb2" } }]);

  // Simulate out-of-order completion: allow "save #3" to complete before "save #2".
  // In fixed implementations, save #3 won't even be invoked until save #2 completes,
  // but pre-releasing it keeps the test deterministic.
  storage.release(3);
  for (let i = 0; i < 25 && storage.saveCount < 3; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await Promise.resolve();
  }
  storage.release(2);

  await Promise.all([p1, p2]);

  assert.equal(storage.saveCount, 3);
  const persisted = storage.data;
  assert.ok(persisted, "Expected JsonVectorStore to persist bytes");
  const parsed = JSON.parse(new TextDecoder().decode(persisted));
  const ids = parsed.records.map((r) => r.id).sort();
  assert.deepEqual(ids, ["b", "c"]);
});

test("JsonVectorStore serializes clear persistence writes to prevent lost updates", async () => {
  const storage = new ControlledBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 2, autoSave: true });
  await store.load();

  storage.release(1);
  await store.upsert([
    { id: "a", vector: [1, 0], metadata: { workbookId: "wb1" } },
    { id: "b", vector: [0, 1], metadata: { workbookId: "wb2" } },
  ]);

  const p1 = store.clear();
  // Wait until clear has invoked `storage.save()` so its snapshot is taken before
  // starting the concurrent upsert.
  await storage.waitForSaves(2);

  const p2 = store.upsert([{ id: "c", vector: [1, 1], metadata: { workbookId: "wb2" } }]);

  storage.release(3);
  for (let i = 0; i < 25 && storage.saveCount < 3; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await Promise.resolve();
  }
  storage.release(2);

  await Promise.all([p1, p2]);

  assert.equal(storage.saveCount, 3);
  const persisted = storage.data;
  assert.ok(persisted, "Expected JsonVectorStore to persist bytes");
  const parsed = JSON.parse(new TextDecoder().decode(persisted));
  const ids = parsed.records.map((r) => r.id).sort();
  assert.deepEqual(ids, ["c"]);
});

test("JsonVectorStore serializes updateMetadata persistence writes to prevent lost updates", async () => {
  const storage = new ControlledBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 2, autoSave: true });
  await store.load();

  // Seed a record, allowing the initial persist to complete immediately.
  storage.release(1);
  await store.upsert([{ id: "a", vector: [1, 0], metadata: { workbookId: "wb", tag: "v1" } }]);

  const p1 = store.updateMetadata([{ id: "a", metadata: { workbookId: "wb", tag: "v2" } }]);
  await storage.waitForSaves(2);
  const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { workbookId: "wb" } }]);

  // Allow the upsert save to complete first in buggy implementations where saves overlap.
  storage.release(3);
  for (let i = 0; i < 25 && storage.saveCount < 3; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await Promise.resolve();
  }
  storage.release(2);

  await Promise.all([p1, p2]);

  assert.equal(storage.saveCount, 3);
  const persisted = storage.data;
  assert.ok(persisted, "Expected JsonVectorStore to persist bytes");
  const parsed = JSON.parse(new TextDecoder().decode(persisted));
  const byId = new Map(parsed.records.map((r) => [r.id, r]));
  assert.deepEqual(Array.from(byId.keys()).sort(), ["a", "b"]);
  assert.equal(byId.get("a").metadata.tag, "v2");
});

let sqlJsAvailable = true;
try {
  // Keep this as a computed dynamic import (no literal bare specifier) so
  // `scripts/run-node-tests.mjs` can still execute this file when `node_modules/`
  // is missing.
  const sqlJsModuleName = "sql" + ".js";
  await import(sqlJsModuleName);
} catch {
  sqlJsAvailable = false;
}

/** @type {Promise<any> | null} */
let sqliteVectorStorePromise = null;

async function getSqliteVectorStore() {
  if (!sqliteVectorStorePromise) {
    // Avoid a literal specifier in a dynamic import for sqliteVectorStore so
    // `scripts/run-node-tests.mjs` (regex-based import detection) doesn't treat this
    // file as requiring external deps when `node_modules/` is missing.
    const modulePath = "../src/store/" + "sqliteVectorStore.js";
    sqliteVectorStorePromise = import(modulePath).then((mod) => mod.SqliteVectorStore);
  }
  return sqliteVectorStorePromise;
}

test(
  "SqliteVectorStore serializes persistence writes to prevent lost updates",
  { skip: !sqlJsAvailable },
  async () => {
    const storage = new ControlledBinaryStorage();
    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: true });

    const p1 = store.upsert([{ id: "a", vector: [1, 0], metadata: { label: "A" } }]);
    await storage.waitForSaves(1);
    const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { label: "B" } }]);

    storage.release(2);
    for (let i = 0; i < 25 && storage.saveCount < 2; i += 1) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.resolve();
    }
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

test(
  "SqliteVectorStore serializes updateMetadata persistence writes to prevent lost updates",
  { skip: !sqlJsAvailable },
  async () => {
    const storage = new ControlledBinaryStorage();
    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: true });

    // Seed one record. Pre-release the initial save so it completes immediately.
    storage.release(1);
    await store.upsert([{ id: "a", vector: [1, 0], metadata: { workbookId: "wb", tag: "v1" } }]);

    const p1 = store.updateMetadata([{ id: "a", metadata: { workbookId: "wb", tag: "v2" } }]);
    await storage.waitForSaves(2);
    const p2 = store.upsert([{ id: "b", vector: [0, 1], metadata: { workbookId: "wb" } }]);

    storage.release(3);
    for (let i = 0; i < 25 && storage.saveCount < 3; i += 1) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.resolve();
    }
    storage.release(2);

    await Promise.all([p1, p2]);
    storage.releaseAll();

    const reloaded = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: false });
    const recA = await reloaded.get("a");
    assert.ok(recA);
    assert.equal(recA.metadata.tag, "v2");
    assert.ok(await reloaded.get("b"));

    await store.close();
    await reloaded.close();
  }
);

test(
  "SqliteVectorStore serializes deleteWorkbook persistence writes to prevent lost updates",
  { skip: !sqlJsAvailable },
  async () => {
    const storage = new ControlledBinaryStorage();
    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: true });

    // Seed two workbooks.
    const seed = store.upsert([
      { id: "a", vector: [1, 0], metadata: { workbookId: "wb1" } },
      { id: "b", vector: [0, 1], metadata: { workbookId: "wb2" } },
    ]);
    await storage.waitForSaves(1);
    storage.release(1);
    await seed;

    const p1 = store.deleteWorkbook("wb1");
    const p2 = store.upsert([{ id: "c", vector: [1, 1], metadata: { workbookId: "wb2" } }]);

    // Ensure the deleteWorkbook save is in-flight, then release the upsert save before it.
    await storage.waitForSaves(2);
    storage.release(3);
    await Promise.resolve();
    storage.release(2);

    await Promise.all([p1, p2]);

    storage.releaseAll();
    const reloaded = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: false });
    const ids = (await reloaded.list({ includeVector: false })).map((r) => r.id).sort();
    assert.deepEqual(ids, ["b", "c"]);

    await store.close();
    await reloaded.close();
  }
);

test(
  "SqliteVectorStore serializes clear persistence writes to prevent lost updates",
  { skip: !sqlJsAvailable },
  async () => {
    const storage = new ControlledBinaryStorage();
    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: true });

    const seed = store.upsert([
      { id: "a", vector: [1, 0], metadata: { workbookId: "wb1" } },
      { id: "b", vector: [0, 1], metadata: { workbookId: "wb2" } },
    ]);
    await storage.waitForSaves(1);
    storage.release(1);
    await seed;

    const p1 = store.clear();
    const p2 = store.upsert([{ id: "c", vector: [1, 1], metadata: { workbookId: "wb2" } }]);

    await storage.waitForSaves(2);
    storage.release(3);
    await Promise.resolve();
    storage.release(2);

    await Promise.all([p1, p2]);

    storage.releaseAll();
    const reloaded = await SqliteVectorStore.create({ storage, dimension: 2, autoSave: false });
    const ids = (await reloaded.list({ includeVector: false })).map((r) => r.id).sort();
    assert.deepEqual(ids, ["c"]);

    await store.close();
    await reloaded.close();
  }
);
