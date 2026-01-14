import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { requireYjsCjs } from "../../yjs-utils/test/require-yjs-cjs.js";

import { REMOTE_ORIGIN, createCollabUndoService } from "../index.js";

test("collab undo: overwriting a foreign item added after UndoManager construction is undoable (CJS applyUpdate)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.getMap("foo").set("a", 1);
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const map = doc.getMap("foo");

  // Construct UndoManager *before* the foreign item is introduced.
  const undo = createCollabUndoService({ doc, scope: map });

  // Apply remote update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);
  assert.equal(map.get("a"), 1);

  undo.transact(() => {
    map.set("a", 2);
  });
  undo.stopCapturing();

  assert.equal(map.get("a"), 2);
  assert.equal(undo.canUndo(), true);

  undo.undo();
  assert.equal(map.get("a"), 1);

  undo.undoManager.destroy();
  doc.destroy();
  remote.destroy();
});
