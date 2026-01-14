import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { requireYjsCjs } from "../../yjs-utils/test/require-yjs-cjs.js";

import { REMOTE_ORIGIN, createCollabUndoService } from "../index.js";

test("collab undo: patches foreign Item constructors even when the constructor is renamed", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.getMap("foo").set("a", 1);
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const map = doc.getMap("foo"); // ensure local root constructor
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  assert.equal(map.get("a"), 1);

  const foreignItem = map._map?.get("a");
  assert.ok(foreignItem, "expected Y.Map to contain an internal Item for key 'a'");
  assert.equal(foreignItem instanceof Y.Item, false, "expected foreign Item struct (not instanceof the ESM Y.Item)");

  // Simulate a bundler-renamed foreign `Item` constructor.
  class RenamedItem extends foreignItem.constructor {}
  Object.setPrototypeOf(foreignItem, RenamedItem.prototype);
  assert.equal(foreignItem.constructor?.name, "RenamedItem");

  const undo = createCollabUndoService({ doc, scope: map });
  // The constructor patch should make the foreign item pass `instanceof` checks.
  assert.equal(foreignItem instanceof Y.Item, true);

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
