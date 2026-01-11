import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";
import { bindCollabSessionToDocumentController, createCollabSession } from "../src/index.ts";

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
    if (fn()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

test("CollabSession↔DocumentController binder masks unreadable remote values/formulas", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 1,
        readAllowlist: ["u-a"],
        editAllowlist: [],
      },
    ],
  });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  const binderB = bindCollabSessionToDocumentController({ session: sessionB, documentController: dcB });

  // Perform edits via DocumentController (typical UI path) so we exercise
  // DocumentController→Yjs propagation as well.
  dcA.setCellValue("Sheet1", "A1", "super secret");
  dcA.setCellFormula("Sheet1", "B1", "=TOP_SECRET()");

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "A1");
    return cellA.value === "super secret" && cellA.formula == null;
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "A1");
    return cellB.value === "###" && cellB.formula == null;
  });

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "B1");
    return cellA.formula === "=TOP_SECRET()" && cellA.value == null;
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "B1");
    return cellB.value === "###" && cellB.formula == null;
  });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder blocks edits to non-editable cells", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only u-a can edit.
        readAllowlist: [],
        editAllowlist: ["u-a"],
      },
    ],
  });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  const binderB = bindCollabSessionToDocumentController({ session: sessionB, documentController: dcB });

  // Seed a value as the editable user.
  dcA.setCellValue("Sheet1", "A1", "original");
  await waitForCondition(() => dcB.getCell("Sheet1", "A1").value === "original");
  await waitForCondition(() => sessionA.getCell("Sheet1:0:0")?.value === "original");

  // Attempt an edit as the restricted user.
  dcB.setCellValue("Sheet1", "A1", "hacked");

  // Local UI and shared Yjs document should remain unchanged.
  await new Promise((r) => setTimeout(r, 25));
  assert.equal(dcB.getCell("Sheet1", "A1").value, "original");
  assert.equal(dcA.getCell("Sheet1", "A1").value, "original");
  assert.equal(sessionA.getCell("Sheet1:0:0")?.value, "original");

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
