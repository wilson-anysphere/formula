import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { createCollabSession } from "../src/index.ts";

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

test("CollabSession undo only reverts local edits (in-memory sync)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  sessionA.setCellValue("Sheet1:0:0", "from-a");
  sessionB.setCellValue("Sheet1:0:1", "from-b");

  assert.equal(sessionA.getCell("Sheet1:0:0")?.value, "from-a");
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "from-a");
  assert.equal(sessionA.getCell("Sheet1:0:1")?.value, "from-b");
  assert.equal(sessionB.getCell("Sheet1:0:1")?.value, "from-b");

  sessionA.undo?.undo();

  assert.equal(sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(sessionB.getCell("Sheet1:0:0"), null);
  assert.equal(sessionA.getCell("Sheet1:0:1")?.value, "from-b");
  assert.equal(sessionB.getCell("Sheet1:0:1")?.value, "from-b");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

