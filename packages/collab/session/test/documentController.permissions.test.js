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

test("CollabSession↔DocumentController binder masks unreadable remote values/formulas", () => {
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
        readAllowlist: ["u-a"],
        editAllowlist: [],
      },
    ],
  });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  const binderB = bindCollabSessionToDocumentController({ session: sessionB, documentController: dcB });

  sessionA.setCellFormula("Sheet1:0:0", "=TOP_SECRET()");

  const cellA = dcA.getCell("Sheet1", "A1");
  assert.equal(cellA.formula, "=TOP_SECRET()");
  assert.equal(cellA.value, null);

  const cellB = dcB.getCell("Sheet1", "A1");
  assert.equal(cellB.value, "###");
  assert.equal(cellB.formula, null);

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder blocks edits to non-editable cells", () => {
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
  assert.equal(dcB.getCell("Sheet1", "A1").value, "original");

  // Attempt an edit as the restricted user.
  dcB.setCellValue("Sheet1", "A1", "hacked");

  // Local UI and shared Yjs document should remain unchanged.
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

