import test from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";

import * as Y from "yjs";
import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";
import { createCollabSession } from "../src/index.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function openDb(name) {
  return await new Promise((resolve, reject) => {
    const req = indexedDB.open(name);
    req.onerror = () => reject(req.error ?? new Error("Failed to open IndexedDB"));
    req.onsuccess = () => resolve(req.result);
  });
}

async function countUpdatesInDb(db) {
  return await new Promise((resolve, reject) => {
    try {
      const tx = db.transaction(["updates"], "readonly");
      const store = tx.objectStore("updates");
      const req = store.count();
      req.onerror = () => reject(req.error ?? new Error("Failed to count updates"));
      req.onsuccess = () => resolve(Number(req.result) || 0);
    } catch (err) {
      reject(err);
    }
  });
}

async function countIndexedDbUpdates(docId) {
  const db = await openDb(docId);
  try {
    return await countUpdatesInDb(db);
  } finally {
    db.close();
  }
}

test("CollabSession local IndexedDB persistence round-trip (restart)", async () => {
  const docId = `doc-${randomUUID()}`;

  {
    const persistence = new IndexedDbCollabPersistence();
    const session = createCollabSession({ docId, persistence });
    await session.whenLocalPersistenceLoaded();

    session.setCellValue("Sheet1:0:0", "hello");
    session.setCellFormula("Sheet1:0:1", "=2+2");

    // Allow the IndexedDB transaction to commit.
    await sleep(10);

    session.destroy();
    session.doc.destroy();
  }

  {
    const persistence = new IndexedDbCollabPersistence();
    const session = createCollabSession({ docId, persistence });
    await session.whenLocalPersistenceLoaded();

    assert.equal((await session.getCell("Sheet1:0:0"))?.value, "hello");
    assert.equal((await session.getCell("Sheet1:0:1"))?.formula, "=2+2");

    await persistence.clear(docId);
    session.destroy();
    session.doc.destroy();
  }
});

test("IndexedDbCollabPersistence compact rewrites update log to a snapshot", async () => {
  const docId = `doc-${randomUUID()}`;

  const persistence = new IndexedDbCollabPersistence({ maxUpdates: 0 });
  const session = createCollabSession({ docId, persistence });
  await session.whenLocalPersistenceLoaded();

  // Generate multiple incremental updates.
  for (let i = 0; i < 10; i += 1) {
    session.setCellValue(`Sheet1:${i}:0`, `v${i}`);
  }

  // Give y-indexeddb a tick to commit its transactions.
  await sleep(25);
  const before = await countIndexedDbUpdates(docId);
  assert.ok(before >= 1);

  await persistence.compact(docId);
  await sleep(10);

  const after = await countIndexedDbUpdates(docId);
  // Best-effort: depending on timing, y-indexeddb may still append a small number of
  // updates that race with compaction, but the log should be bounded.
  assert.ok(after <= 3);

  session.destroy();
  session.doc.destroy();

  // Restart: ensure state is still recoverable after compaction.
  {
    const restarted = createCollabSession({ docId, persistence: new IndexedDbCollabPersistence() });
    await restarted.whenLocalPersistenceLoaded();
    assert.equal((await restarted.getCell("Sheet1:0:0"))?.value, "v0");
    assert.equal((await restarted.getCell("Sheet1:9:0"))?.value, "v9");
    await restarted.flushLocalPersistence?.().catch(() => {});
    restarted.destroy();
    restarted.doc.destroy();
  }

  await persistence.clear(docId);
});

test("IndexedDbCollabPersistence load settles if destroyed before initial sync completes", async () => {
  const docId = `doc-${randomUUID()}`;
  const doc = new Y.Doc({ guid: docId });
  const persistence = new IndexedDbCollabPersistence();

  const binding = persistence.bind(docId, doc);
  const loadPromise = persistence.load(docId, doc);

  // Destroy the binding immediately to simulate app teardown while the
  // underlying y-indexeddb `whenSynced` promise is still pending.
  await binding.destroy();

  await Promise.race([
    loadPromise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error("Timed out waiting for IndexedDB load to settle")), 2_000)
    ),
  ]);

  await persistence.clear(docId);
  doc.destroy();
});
