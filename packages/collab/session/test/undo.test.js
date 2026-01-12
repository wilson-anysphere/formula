import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { createCollabSession } from "../src/index.ts";

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

test("CollabSession undo only reverts local edits (in-memory sync)", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  await sessionA.setCellValue("Sheet1:0:0", "from-a");
  await sessionB.setCellValue("Sheet1:0:1", "from-b");

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, "from-a");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "from-a");
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value, "from-b");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "from-b");

  sessionA.undo?.undo();

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value, "from-b");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "from-b");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession undo captures cell edits when cell maps were created by a different Yjs instance (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const cells = remote.getMap("cells");
  remote.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", "from-cjs");
    cell.set("formula", null);
    cell.set("modified", 1);
    cells.set("Sheet1:0:0", cell);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Ensure the root exists in this module so the update only introduces foreign
  // nested cell maps (not a foreign `cells` root, which would prevent session
  // construction).
  doc.getMap("cells");
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  const session = createCollabSession({ doc, undo: {} });

  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  await session.setCellValue("Sheet1:0:0", "edited");
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo captures cell edits when foreign Yjs maps have renamed constructors", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const cells = remote.getMap("cells");
  remote.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", "from-cjs");
    cell.set("formula", null);
    cell.set("modified", 1);
    cells.set("Sheet1:0:0", cell);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  doc.getMap("cells");
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  // Simulate a bundler-renamed constructor without mutating global `yjs` state
  // (which can cause cross-test interference under concurrency).
  const foreignCellMap = doc.getMap("cells").get("Sheet1:0:0");
  assert.ok(foreignCellMap);
  class RenamedMap extends foreignCellMap.constructor {}
  Object.setPrototypeOf(foreignCellMap, RenamedMap.prototype);

  const session = createCollabSession({ doc, undo: {} });

  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  await session.setCellValue("Sheet1:0:0", "edited");
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo works when the cells root was created by a different Yjs instance (CJS Doc.getMap)", async () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate a mixed module loader environment where another Yjs instance eagerly
  // instantiates the `cells` root before CollabSession is constructed.
  const foreignCells = Ycjs.Doc.prototype.getMap.call(doc, "cells");

  const foreignCell = new Ycjs.Map();
  foreignCell.set("value", "from-cjs");
  foreignCell.set("formula", null);
  foreignCell.set("modified", 1);
  foreignCells.set("Sheet1:0:0", foreignCell);

  const session = createCollabSession({ doc, undo: {} });

  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  await session.setCellValue("Sheet1:0:0", "edited");
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo scopeNames works when an additional root was created by a different Yjs instance (CJS Doc.getMap)", async () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate an eager foreign instantiation of a root map that we want to include
  // in the undo scope.
  const foreignExtra = Ycjs.Doc.prototype.getMap.call(doc, "extraUndoScope");
  foreignExtra.set("k", 0);

  const session = createCollabSession({ doc, undo: { scopeNames: ["extraUndoScope"] } });

  // Regression: session construction should not throw "different constructor" and
  // the root should be accessible via the local `doc.getMap`.
  const extra = doc.getMap("extraUndoScope");
  assert.ok(extra instanceof Y.Map);
  assert.equal(extra.get("k"), 0);

  session.undo?.transact?.(() => {
    extra.set("k", 1);
  });
  session.undo?.stopCapturing();

  assert.equal(extra.get("k"), 1);
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal(extra.get("k"), 0);

  session.destroy();
  doc.destroy();
});
