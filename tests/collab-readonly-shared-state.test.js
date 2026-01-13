import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindCollabSessionToDocumentController, createCollabSession } from "../packages/collab/session/src/index.ts";

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
    const id = entry?.get?.("id") ?? entry?.id;
    if (id === sheetId) return entry;
  }
  return null;
}

test("read-only collab roles do not write sheet-level view/format state into Yjs (local-only)", async () => {
  const ydoc = new Y.Doc();
  const session = createCollabSession({ doc: ydoc });
  session.setPermissions({ role: "viewer", rangeRestrictions: [], userId: "viewer-1" });

  const documentController = new DocumentController();

  const binder = await bindCollabSessionToDocumentController({
    session,
    documentController,
    defaultSheetId: "Sheet1",
    userId: "viewer-1",
  });

  /** @type {number} */
  let updates = 0;
  const onUpdate = () => {
    updates += 1;
  };

  try {
    // Let initial schema + hydration settle before counting updates.
    await new Promise((r) => setTimeout(r, 25));
    updates = 0;
    ydoc.on("update", onUpdate);

    // Local changes should be applied to the DocumentController but not persisted into Yjs.
    documentController.setFrozen("Sheet1", 2, 1);
    documentController.setColWidth("Sheet1", 0, 120);
    documentController.setSheetFormat("Sheet1", { font: { bold: true } });

    // Give the binder a chance to enqueue any writes (should be none).
    await new Promise((r) => setTimeout(r, 25));

    assert.equal(updates, 0, "expected no Yjs updates from local view/format changes in read-only role");

    // Local UI state should reflect the changes.
    assert.equal(documentController.getSheetView("Sheet1").frozenRows, 2);
    assert.equal(documentController.getSheetView("Sheet1").frozenCols, 1);
    assert.equal(documentController.getSheetView("Sheet1").colWidths?.["0"], 120);

    await waitForCondition(() => documentController.getCellFormat("Sheet1", "A1")?.font?.bold === true);
    assert.equal(documentController.getCellFormat("Sheet1", "A1")?.font?.bold, true);

    // Shared Yjs doc should not contain the sheet view state or formatting defaults.
    const sheetEntry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(sheetEntry && typeof sheetEntry.get === "function", "expected Sheet1 entry in Yjs");
    assert.equal(sheetEntry.get("view"), undefined);
    assert.equal(sheetEntry.get("defaultFormat"), undefined);
  } finally {
    ydoc.off("update", onUpdate);
    binder.destroy();
    session.destroy();
    ydoc.destroy();
  }
});

