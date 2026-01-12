import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "./yjsSpreadsheetDocAdapter.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    // yjs prints a warning when imported multiple times via different module systems.
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

function isYMapLike(value) {
  return (
    value &&
    typeof value === "object" &&
    typeof value.get === "function" &&
    typeof value.set === "function" &&
    typeof value.delete === "function" &&
    typeof value.observeDeep === "function" &&
    typeof value.unobserveDeep === "function"
  );
}

test('createYjsSpreadsheetDocAdapter.applyState uses origin "versioning-restore"', (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const adapter = createYjsSpreadsheetDocAdapter(doc);
  const cells = doc.getMap("cells");

  // Seed a simple workbook cell using the canonical collab cell schema (Y.Map).
  const cellA = new Y.Map();
  cellA.set("value", "alpha");
  cellA.set("formula", null);
  cells.set("Sheet1:0:0", cellA);

  const snapshot = adapter.encodeState();

  // Mutate the doc so applyState has real work to do.
  const cellB = new Y.Map();
  cellB.set("value", "beta");
  cellB.set("formula", null);
  cells.set("Sheet1:0:0", cellB);

  /** @type {any[]} */
  const origins = [];
  const onAfterTx = (tx) => origins.push(tx?.origin);
  doc.on("afterTransaction", onAfterTx);
  t.after(() => {
    doc.off("afterTransaction", onAfterTx);
  });

  adapter.applyState(snapshot);

  assert.deepEqual(origins, ["versioning-restore"]);

  const restored = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.equal(restored?.get?.("value") ?? null, "alpha");
  assert.equal(restored?.get?.("formula") ?? null, null);
});

test("createYjsSpreadsheetDocAdapter.applyState works when roots were created by a different Yjs instance (CJS applyUpdate)", (t) => {
  const Ycjs = requireYjsCjs();
  const remote = new Ycjs.Doc();
  t.after(() => remote.destroy());

  remote.transact(() => {
    const cells = remote.getMap("cells");
    const cell = new Ycjs.Map();
    cell.set("value", "alpha");
    cell.set("formula", null);
    cells.set("Sheet1:0:0", cell);

    const comments = remote.getMap("comments");
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("content", "hello");
    comments.set("c1", comment);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update);

  // Verify the roots exist but are not necessarily from this module's Yjs instance.
  assert.equal(doc.share.has("cells"), true);
  assert.equal(doc.share.has("comments"), true);
  assert.equal(doc.share.get("cells") instanceof Y.Map, false);
  assert.equal(doc.share.get("comments") instanceof Y.Map, false);

  const adapter = createYjsSpreadsheetDocAdapter(doc);
  const snapshot = adapter.encodeState();

  // Mutate the doc using the foreign Yjs instance so we don't depend on `doc.getMap`
  // working with foreign roots.
  const foreignCells = Ycjs.Doc.prototype.getMap.call(doc, "cells");
  const cellB = new Ycjs.Map();
  cellB.set("value", "beta");
  cellB.set("formula", null);
  foreignCells.set("Sheet1:0:0", cellB);

  const foreignComments = Ycjs.Doc.prototype.getMap.call(doc, "comments");
  const commentB = new Ycjs.Map();
  commentB.set("id", "c1");
  commentB.set("content", "bye");
  foreignComments.set("c1", commentB);

  adapter.applyState(snapshot);

  const cellsRoot = doc.share.get("cells");
  const commentsRoot = doc.share.get("comments");
  assert.equal(isYMapLike(cellsRoot), true);
  assert.equal(isYMapLike(commentsRoot), true);

  const restoredCell = /** @type {any} */ (cellsRoot).get("Sheet1:0:0");
  assert.equal(isYMapLike(restoredCell), true);
  assert.equal(restoredCell.get("value"), "alpha");
  assert.equal(restoredCell.get("formula"), null);

  const restoredComment = /** @type {any} */ (commentsRoot).get("c1");
  assert.equal(isYMapLike(restoredComment), true);
  assert.equal(restoredComment.get("content"), "hello");

  // Ensure the resulting doc can still be encoded by the local Yjs module instance.
  assert.doesNotThrow(() => {
    Y.encodeStateAsUpdate(doc);
  });
});
