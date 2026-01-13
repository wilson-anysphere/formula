import test from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";

import * as Y from "yjs";
import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { IndexedDbCollabPersistence } from "../src/indexeddb.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function requestResult(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
  });
}

function transactionDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
  });
}

async function openDb(name) {
  const req = indexedDB.open(name);
  return await requestResult(req);
}

async function countUpdateRecords(dbName) {
  const db = await openDb(dbName);
  try {
    const tx = db.transaction("updates", "readonly");
    const store = tx.objectStore("updates");
    const count = await requestResult(store.count());
    await transactionDone(tx);
    return count;
  } finally {
    db.close();
  }
}

async function getLastUpdateKey(dbName) {
  const db = await openDb(dbName);
  try {
    const tx = db.transaction("updates", "readonly");
    const store = tx.objectStore("updates");
    const key = await new Promise((resolve, reject) => {
      const req = store.openCursor(null, "prev");
      req.onsuccess = () => resolve(req.result?.key ?? null);
      req.onerror = () => reject(req.error ?? new Error("IndexedDB cursor failed"));
    });
    await transactionDone(tx);
    return key;
  } finally {
    db.close();
  }
}

test("IndexedDbCollabPersistence flush compacts the updates store (bounded size + reload)", async () => {
  const docId = `doc-${randomUUID()}`;

  const doc = new Y.Doc({ guid: docId });
  const persistence = new IndexedDbCollabPersistence();
  const binding = persistence.bind(docId, doc);
  await persistence.load(docId, doc);

  const root = doc.getMap("root");
  for (let i = 0; i < 25; i += 1) {
    root.set(`k${i}`, `v${i}`);
    // Give y-indexeddb a microtask to enqueue its own update write before we compact.
    await new Promise((resolve) => queueMicrotask(resolve));
    await persistence.flush(docId);
  }

  // Allow any trailing y-indexeddb writes (from the last mutation) to commit.
  await sleep(25);

  const count = await countUpdateRecords(docId);
  assert.ok(count <= 3, `expected <=3 update records after compaction, got ${count}`);

  await binding.destroy();
  doc.destroy();

  const restartedDoc = new Y.Doc({ guid: docId });
  const restarted = new IndexedDbCollabPersistence();
  restarted.bind(docId, restartedDoc);
  await restarted.load(docId, restartedDoc);

  const restartedRoot = restartedDoc.getMap("root");
  for (let i = 0; i < 25; i += 1) {
    assert.equal(restartedRoot.get(`k${i}`), `v${i}`);
  }

  await restarted.clear(docId);
  restartedDoc.destroy();
});

test("IndexedDbCollabPersistence compaction preserves updates written by other persistence instances", async () => {
  const docId = `doc-${randomUUID()}`;

  // Instance A writes an update after instance B has already loaded state. When B compacts,
  // it must not drop A's persisted update even though B's in-memory doc doesn't include it.
  const docA = new Y.Doc({ guid: docId });
  const persistenceA = new IndexedDbCollabPersistence();
  persistenceA.bind(docId, docA);
  await persistenceA.load(docId, docA);

  docA.getMap("root").set("a", "one");
  await sleep(25);

  const docB = new Y.Doc({ guid: docId });
  const persistenceB = new IndexedDbCollabPersistence();
  persistenceB.bind(docId, docB);
  await persistenceB.load(docId, docB);

  // Write another update from A after B is synced/loaded. B won't see this update in-memory.
  docA.getMap("root").set("b", "two");
  await sleep(25);

  // Compaction should merge existing persisted updates (including A's) with B's snapshot.
  await persistenceB.flush(docId);
  await sleep(25);

  const docC = new Y.Doc({ guid: docId });
  const persistenceC = new IndexedDbCollabPersistence();
  persistenceC.bind(docId, docC);
  await persistenceC.load(docId, docC);

  const rootC = docC.getMap("root");
  assert.equal(rootC.get("a"), "one");
  assert.equal(rootC.get("b"), "two");

  docA.destroy();
  docB.destroy();
  await sleep(25);

  await persistenceC.clear(docId);
  docC.destroy();
});

test("IndexedDbCollabPersistence flush skips compaction rewrite when already compact + unchanged", async () => {
  const docId = `doc-${randomUUID()}`;

  const doc = new Y.Doc({ guid: docId });
  const persistence = new IndexedDbCollabPersistence();
  persistence.bind(docId, doc);
  await persistence.load(docId, doc);

  doc.getMap("root").set("k", "v");
  await sleep(25);
  await persistence.flush(docId);
  await sleep(25);

  assert.equal(await countUpdateRecords(docId), 1);
  const lastKeyBefore = await getLastUpdateKey(docId);

  // No further edits; flush should be a no-op (no new snapshot rewrite).
  await persistence.flush(docId);
  await sleep(25);

  assert.equal(await countUpdateRecords(docId), 1);
  const lastKeyAfter = await getLastUpdateKey(docId);
  assert.equal(lastKeyAfter, lastKeyBefore);

  await persistence.clear(docId);
  doc.destroy();
});

test("IndexedDbCollabPersistence flush resolves if doc is destroyed during compaction", async () => {
  const docId = `doc-${randomUUID()}`;

  const doc = new Y.Doc({ guid: docId });
  const persistence = new IndexedDbCollabPersistence();
  persistence.bind(docId, doc);
  await persistence.load(docId, doc);

  doc.getMap("root").set("k", "v");
  await sleep(10);

  const flushPromise = persistence.flush(docId);
  doc.destroy();

  await Promise.race([
    flushPromise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("Timed out waiting for flush to settle")), 2_000)),
  ]);

  await persistence.clear(docId);
});

test("IndexedDbCollabPersistence compaction does not hang when load() is in-flight", async () => {
  const docId = `doc-${randomUUID()}`;
  const doc = new Y.Doc({ guid: docId });
  const persistence = new IndexedDbCollabPersistence();

  persistence.bind(docId, doc);
  const loadPromise = persistence.load(docId, doc);
  const flushPromise = persistence.flush(docId);

  await Promise.race([
    Promise.all([loadPromise, flushPromise]),
    new Promise((_, reject) => setTimeout(() => reject(new Error("Timed out waiting for load+flush to settle")), 2_000)),
  ]);

  await persistence.clear(docId);
  doc.destroy();
});
