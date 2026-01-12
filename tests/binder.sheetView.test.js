import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

async function waitForCondition(predicate, timeoutMs = 2_000, intervalMs = 5) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

function findSheetEntry(ydoc, sheetId) {
  const sheets = ydoc.getArray("sheets");
  for (const entry of sheets.toArray()) {
    if (!entry || typeof entry.get !== "function") continue;
    if (entry.get("id") === sheetId) return entry;
  }
  return null;
}

test("binder: DocumentController→Yjs syncs sheet view state (freeze panes + row/col sizes)", async () => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentController();

  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    documentController.setFrozen("Sheet1", 2, 3);
    documentController.setColWidth("Sheet1", 0, 120);
    documentController.setRowHeight("Sheet1", 5, 40);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get("view");
      return (
        view?.frozenRows === 2 &&
        view?.frozenCols === 3 &&
        view?.colWidths?.["0"] === 120 &&
        view?.rowHeights?.["5"] === 40
      );
    });

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected a Yjs sheets entry for Sheet1");
    assert.deepEqual(entry.get("view"), {
      frozenRows: 2,
      frozenCols: 3,
      colWidths: { "0": 120 },
      rowHeights: { "5": 40 },
    });

    // Binder should not write legacy top-level frozen rows/cols.
    assert.equal(entry.get("frozenRows"), undefined);
    assert.equal(entry.get("frozenCols"), undefined);
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: Yjs→DocumentController syncs sheet view state (initial hydration + remote changes)", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    entry.set("view", { frozenRows: 1, frozenCols: 1, colWidths: { "0": 111 } });
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 1 && view.colWidths?.["0"] === 111;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 1,
      colWidths: { "0": 111 },
    });

    // Simulate a remote collaborator updating the view object.
    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(
      () => {
        const entry = findSheetEntry(ydoc, "Sheet1");
        assert.ok(entry, "expected Sheet1 entry");
        entry.set("view", { frozenRows: 0, frozenCols: 2, rowHeights: { "10": 55 } });
      },
      remoteOrigin,
    );

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 0 && view.frozenCols === 2 && view.rowHeights?.["10"] === 55 && !view.colWidths;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 0,
      frozenCols: 2,
      rowHeights: { "10": 55 },
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: local sheet view changes do not echo back as external changes", async () => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentController();

  /** @type {any[]} */
  const changeEvents = [];
  const unsubscribe = documentController.on("change", (payload) => changeEvents.push(payload));

  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    // Give hydration a chance to run (it should be a no-op for empty docs).
    await new Promise((r) => setTimeout(r, 25));
    changeEvents.length = 0;

    documentController.setFrozen("Sheet1", 1, 1);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get("view");
      return view?.frozenRows === 1 && view?.frozenCols === 1;
    });

    // Allow any potential echo to schedule.
    await new Promise((r) => setTimeout(r, 25));

    const collabEvents = changeEvents.filter((evt) => evt?.source === "collab");
    assert.equal(collabEvents.length, 0);
  } finally {
    unsubscribe();
    binder.destroy();
    ydoc.destroy();
  }
});

