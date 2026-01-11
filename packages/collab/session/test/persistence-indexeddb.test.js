import test from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";
import { createCollabSession } from "../src/index.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
