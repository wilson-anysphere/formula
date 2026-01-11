import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

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

test("CollabSession E2E cell encryption: encrypted in Yjs, decrypted with key, masked without key (in-memory sync)", async () => {
  const docId = "collab-session-encryption-test-doc";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({
    doc: docA,
    encryption: { keyForCell: keyForA1 },
  });
  const sessionB = createCollabSession({ doc: docB });

  await sessionA.setCellValue("Sheet1:0:0", "top-secret");

  // Raw Yjs should not contain plaintext.
  const cellMap = sessionA.cells.get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(typeof cellMap.get, "function");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, "top-secret");

  const masked = await sessionB.getCell("Sheet1:0:0");
  assert.equal(masked?.value, "###");
  assert.equal(masked?.formula, null);
  assert.equal(masked?.encrypted, true);

  // Permission-like helpers should reflect that encrypted cells are unreadable/uneditable
  // without the relevant encryption key.
  assert.equal(sessionB.canReadCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(sessionB.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);

  // safeSet* APIs should fail gracefully (return false) rather than throwing when
  // the cell is encrypted but the key is unavailable.
  assert.equal(await sessionB.safeSetCellValue("Sheet1:0:0", "hacked"), false);
  assert.equal(await sessionB.safeSetCellFormula("Sheet1:0:0", "=HACK()"), false);

  // Now "grant" the key by recreating the session with a resolver.
  sessionB.destroy();
  const sessionBWithKey = createCollabSession({ doc: docB, encryption: { keyForCell: keyForA1 } });
  assert.equal((await sessionBWithKey.getCell("Sheet1:0:0"))?.value, "top-secret");

  sessionA.destroy();
  sessionBWithKey.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
