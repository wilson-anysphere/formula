import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { createCollabSession } from "../src/index.ts";
import { createEncryptedRangeManagerForSession } from "../../encrypted-ranges/src/index.ts";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

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

test("EncryptedRangeManager for session: add/remove sync and local undo (in-memory)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  const mgrA = createEncryptedRangeManagerForSession(sessionA);
  const mgrB = createEncryptedRangeManagerForSession(sessionB);

  const id = mgrA.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });
  sessionA.undo?.stopCapturing();

  assert.equal(mgrB.list().length, 1);
  assert.equal(mgrB.list()[0].id, id);

  // Undo should revert only the local add (and propagate to peers).
  sessionA.undo?.undo();
  assert.equal(mgrA.list().length, 0);
  assert.equal(mgrB.list().length, 0);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("EncryptedRangeManager normalizes foreign (CJS) encryptedRanges arrays before undo-tracked edits", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const metadata = remote.getMap("metadata");
  const ranges = new Ycjs.Array();
  const entry = new Ycjs.Map();
  entry.set("id", "r1");
  entry.set("sheetId", "Sheet1");
  entry.set("startRow", 0);
  entry.set("startCol", 0);
  entry.set("endRow", 0);
  entry.set("endCol", 0);
  entry.set("keyId", "k1");
  ranges.push([entry]);

  remote.transact(() => {
    metadata.set("encryptedRanges", ranges);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);
  const doc = new Y.Doc();

  // Apply update via the CJS build to simulate providers that use a different Yjs instance.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  const session = createCollabSession({ doc, undo: {} });
  const mgr = createEncryptedRangeManagerForSession(session);

  // Existing foreign entries should be visible.
  assert.deepEqual(mgr.list().map((r) => r.id), ["r1"]);

  const id2 = mgr.add({ sheetId: "Sheet1", startRow: 1, startCol: 0, endRow: 1, endCol: 0, keyId: "k1" });
  session.undo?.stopCapturing();

  assert.deepEqual(mgr.list().map((r) => r.id).sort(), ["r1", id2].sort());

  // Undo should remove only the newly-added range and preserve the foreign range.
  session.undo?.undo();
  assert.deepEqual(mgr.list().map((r) => r.id), ["r1"]);

  session.destroy();
  doc.destroy();
  remote.destroy();
});

