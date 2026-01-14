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

    // In read-only roles, sheet-level view/format state is considered *shared state*.
    // Local UI actions that attempt to mutate that state should be reverted so the
    // viewer doesn't diverge from the shared document.
    const initialView = documentController.getSheetView("Sheet1");
    const initialFrozenRows = initialView.frozenRows;
    const initialFrozenCols = initialView.frozenCols;
    const initialColWidth = initialView.colWidths?.["0"];
    const initialBold = documentController.getCellFormat("Sheet1", "A1")?.font?.bold;

    // Local changes should neither be persisted into Yjs nor remain applied locally.
    // The binder reverts them in read-only roles to prevent local-only divergence.
    documentController.setFrozen("Sheet1", 2, 1);
    documentController.setColWidth("Sheet1", 0, 120);
    documentController.setSheetFormat("Sheet1", { font: { bold: true } });

    // Give the binder a chance to enqueue any writes (should be none).
    await new Promise((r) => setTimeout(r, 25));

    assert.equal(updates, 0, "expected no Yjs updates from local view/format changes in read-only role");

    // Local UI state should have been reverted (no local-only divergence).
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return (
        view.frozenRows === initialFrozenRows &&
        view.frozenCols === initialFrozenCols &&
        (view.colWidths?.["0"] ?? null) === (initialColWidth ?? null)
      );
    });
    await waitForCondition(() => {
      const bold = documentController.getCellFormat("Sheet1", "A1")?.font?.bold;
      return (bold ?? null) === (initialBold ?? null);
    });

    // Shared Yjs doc should not contain the sheet view state or formatting defaults.
    const sheetEntry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(sheetEntry && typeof sheetEntry.get === "function", "expected Sheet1 entry in Yjs");
    assert.equal(sheetEntry.get("view"), undefined);
    assert.equal(sheetEntry.get("defaultFormat"), undefined);

    // Remote Yjs updates should still apply to the DocumentController, even in a read-only role.
    const REMOTE_ORIGIN = { type: "remote-test" };
    ydoc.transact(
      () => {
        sheetEntry.set("view", { frozenRows: 0, frozenCols: 3, colWidths: { "0": 200 } });
        sheetEntry.set("defaultFormat", { font: { italic: true } });
      },
      REMOTE_ORIGIN,
    );

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 0 && view.frozenCols === 3 && view.colWidths?.["0"] === 200;
    });
    await waitForCondition(() => documentController.getCellFormat("Sheet1", "A1")?.font?.italic === true);
    assert.ok(updates > 0, "expected remote Yjs mutations to produce Yjs updates");
  } finally {
    ydoc.off("update", onUpdate);
    binder.destroy();
    session.destroy();
    ydoc.destroy();
  }
});
