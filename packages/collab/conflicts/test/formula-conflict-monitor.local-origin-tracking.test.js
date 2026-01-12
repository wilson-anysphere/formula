import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { FormulaConflictMonitor } from "../src/formula-conflict-monitor.js";

const REMOTE_ORIGIN = { type: "remote" };

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

  // Initial sync.
  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

test("FormulaConflictMonitor tracks local-origin edits for causal conflict detection (without setLocalFormula)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    onConflict: (c) => conflicts.push(c)
  });

  const cellKey = "Sheet1:0:0";

  // Establish a shared base cell map so concurrent edits race on the formula key
  // (not on the `cells[cellKey] = new Y.Map()` insertion).
  docA.transact(() => {
    docA.getMap("cells").set(cellKey, new Y.Map());
  }, localOrigin);
  assert.ok(docB.getMap("cells").get(cellKey), "expected base cell map to sync to docB");

  // Simulate offline concurrent edits (same cell, different formulas).
  disconnect();
  docA.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docA.getMap("cells").get(cellKey));
    cell.set("formula", "=1");
    cell.set("value", null);
    // Intentionally omit `modifiedBy` so conflict detection must rely on causality.
  }, localOrigin);

  docB.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docB.getMap("cells").get(cellKey));
    cell.set("formula", "=2");
    cell.set("value", null);
    // Intentionally omit `modifiedBy`.
  });

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  assert.ok(conflicts.length >= 1, "expected at least one conflict to be detected");
  assert.equal(conflicts[0].kind, "formula");

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

