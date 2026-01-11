import test from "node:test";
import assert from "node:assert/strict";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { IndexedDBVersionStore } from "../packages/versioning/src/store/indexeddbVersionStore.js";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("IndexedDBVersionStore saves, lists, updates, and retrieves versions", async () => {
  const dbName = `formula-versioning-${crypto.randomUUID()}`;
  const store = new IndexedDBVersionStore({ dbName });

  const v1 = {
    id: crypto.randomUUID(),
    kind: "checkpoint",
    timestampMs: Date.now(),
    userId: "u1",
    userName: "User",
    description: "Approved",
    checkpointName: "Approved",
    checkpointLocked: false,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([1, 2, 3]),
  };

  await store.saveVersion(v1);

  const fetched = await store.getVersion(v1.id);
  assert.ok(fetched);
  assert.equal(fetched.id, v1.id);
  assert.equal(fetched.kind, "checkpoint");
  assert.deepEqual(Array.from(fetched.snapshot), [1, 2, 3]);

  await store.updateVersion(v1.id, { checkpointLocked: true });
  const updated = await store.getVersion(v1.id);
  assert.equal(updated?.checkpointLocked, true);

  // Ensure ordering by timestampMs desc.
  await sleep(2);
  const v2 = {
    id: crypto.randomUUID(),
    kind: "snapshot",
    timestampMs: Date.now(),
    userId: null,
    userName: null,
    description: "Auto-save",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([9]),
  };
  await store.saveVersion(v2);

  const list = await store.listVersions();
  assert.equal(list.length, 2);
  assert.equal(list[0].id, v2.id);
  assert.equal(list[1].id, v1.id);

  store.close();
});

test("IndexedDBVersionStore deleteVersion removes records", async () => {
  const dbName = `formula-versioning-${crypto.randomUUID()}`;
  const store = new IndexedDBVersionStore({ dbName });

  const v1 = {
    id: crypto.randomUUID(),
    kind: "snapshot",
    timestampMs: Date.now(),
    userId: null,
    userName: null,
    description: "temp",
    checkpointName: null,
    checkpointLocked: null,
    checkpointAnnotations: null,
    snapshot: new Uint8Array([7, 8, 9]),
  };

  await store.saveVersion(v1);
  assert.ok(await store.getVersion(v1.id));

  await store.deleteVersion(v1.id);
  assert.equal(await store.getVersion(v1.id), null);

  const list = await store.listVersions();
  assert.equal(list.length, 0);

  store.close();
});
