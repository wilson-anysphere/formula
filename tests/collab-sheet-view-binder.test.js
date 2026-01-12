import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindSheetViewToCollabSession } from "../apps/desktop/src/collab/sheetViewBinder.ts";

const REMOTE_ORIGIN = Symbol("remote");

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

async function waitForCondition(fn, timeoutMs = 2000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const ok = await fn();
      if (ok) return;
    } catch {
      // Ignore transient errors while waiting for async state to settle.
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

test("CollabSession sheet view binder syncs frozen panes + axis overrides without polluting undo history", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindSheetViewToCollabSession({ session: sessionA, documentController: dcA });
  const binderB = bindSheetViewToCollabSession({ session: sessionB, documentController: dcB });

  dcA.setFrozen("Sheet1", 2, 1, { label: "Freeze" });
  dcA.setColWidth("Sheet1", 0, 120, { label: "Resize Column" });
  dcA.setRowHeight("Sheet1", 1, 40, { label: "Resize Row" });

  await waitForCondition(() => {
    const view = dcB.getSheetView("Sheet1");
    return (
      view.frozenRows === 2 &&
      view.frozenCols === 1 &&
      view.colWidths?.["0"] === 120 &&
      view.rowHeights?.["1"] === 40
    );
  });

  // Remote changes should not create local undo history entries.
  assert.deepEqual(dcB.getStackDepths(), { undo: 0, redo: 0 });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

