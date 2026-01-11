import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { attachOfflinePersistence } from "../src/index.node.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("attachOfflinePersistence restores Yjs state from IndexedDB across restarts", async () => {
  const key = `formula-collab-offline-${crypto.randomUUID()}`;

  {
    const doc = new Y.Doc({ guid: key });
    const persistence = attachOfflinePersistence(doc, { mode: "indexeddb", key });
    await persistence.whenLoaded();

    doc.getMap("cells").set("Sheet1:0:0", 123);

    // Give y-indexeddb a tick to commit its transaction before simulating a restart.
    await sleep(25);

    persistence.destroy();
    doc.destroy();
  }

  {
    const doc = new Y.Doc({ guid: key });
    const persistence = attachOfflinePersistence(doc, { mode: "indexeddb", key });
    await persistence.whenLoaded();

    assert.equal(doc.getMap("cells").get("Sheet1:0:0"), 123);

    await persistence.clear();
    persistence.destroy();
    doc.destroy();
  }
});

