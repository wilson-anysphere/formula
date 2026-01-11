import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import {
  ensureWorkbookSchema,
  MetadataManager,
  NamedRangeManager,
  SheetManager,
} from "@formula/collab-workbook";

import { attachOfflinePersistence } from "../src/index.node.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function snapshotSheets(doc) {
  return doc.getArray("sheets").toArray().map((sheet) => ({
    id: String(sheet.get("id") ?? ""),
    name: sheet.get("name") == null ? null : String(sheet.get("name")),
  }));
}

test("attachOfflinePersistence restores workbook metadata (sheets + namedRanges + metadata) across restarts", async () => {
  const key = `formula-collab-offline-workbook-${crypto.randomUUID()}`;

  {
    const doc = new Y.Doc({ guid: key });
    const persistence = attachOfflinePersistence(doc, { mode: "indexeddb", key });
    await persistence.whenLoaded();

    ensureWorkbookSchema(doc, { defaultSheetId: "Sheet1", defaultSheetName: "Sheet1" });
    const sheets = new SheetManager({ doc });
    const namedRanges = new NamedRangeManager({ doc });
    const metadata = new MetadataManager({ doc });

    sheets.addSheet({ id: "Sheet2", name: "Budget" });
    sheets.moveSheet("Sheet1", 1);
    namedRanges.set("MyRange", { sheetId: "Sheet2", range: "A1:B2" });
    metadata.set("title", "Quarterly Budget");

    // Give y-indexeddb a tick to commit its transaction before simulating a restart.
    await sleep(25);

    persistence.destroy();
    doc.destroy();
  }

  {
    const doc = new Y.Doc({ guid: key });
    const persistence = attachOfflinePersistence(doc, { mode: "indexeddb", key });
    await persistence.whenLoaded();

    // Restored workbook metadata.
    assert.deepEqual(snapshotSheets(doc).map((s) => s.id), ["Sheet2", "Sheet1"]);
    assert.equal(snapshotSheets(doc).find((s) => s.id === "Sheet2")?.name, "Budget");
    assert.deepEqual(doc.getMap("namedRanges").get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });
    assert.equal(doc.getMap("metadata").get("title"), "Quarterly Budget");

    await persistence.clear();
    persistence.destroy();
    doc.destroy();
  }

  {
    const doc = new Y.Doc({ guid: key });
    const persistence = attachOfflinePersistence(doc, { mode: "indexeddb", key });
    await persistence.whenLoaded();

    // State should have been cleared by the previous session.
    assert.equal(doc.getArray("sheets").length, 0);
    assert.equal(doc.getMap("namedRanges").size, 0);
    assert.equal(doc.getMap("metadata").size, 0);

    persistence.destroy();
    doc.destroy();
  }
});
