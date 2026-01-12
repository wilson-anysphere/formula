import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN, createCollabUndoService } from "../index.js";

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

test("collab undo: insert into a foreign nested Y.Map added after UndoManager construction is undoable", () => {
  const Ycjs = requireYjsCjs();

  // Remote update introduces a nested Y.Map instance created by the CJS Yjs build.
  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const root = remote.getMap("foo");
    root.set("nested", new Ycjs.Map());
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const root = doc.getMap("foo");

  // Construct UndoManager *before* the foreign nested type is introduced.
  const undo = createCollabUndoService({ doc, scope: root });

  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  const nested = root.get("nested");
  assert.ok(nested, "expected nested map to exist");
  assert.equal(
    nested instanceof Y.Map,
    false,
    "expected nested map to be a foreign Y.Map instance (not instanceof the ESM Y.Map)"
  );

  undo.transact(() => {
    nested.set("y", 1);
  });
  undo.stopCapturing();

  assert.equal(nested.get("y"), 1);
  assert.equal(undo.canUndo(), true);

  undo.undo();
  assert.equal(nested.get("y"), undefined);

  undo.undoManager.destroy();
  doc.destroy();
  remote.destroy();
});

