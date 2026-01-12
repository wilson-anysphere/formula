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

test("FormulaConflictMonitor tracks local-origin value writes in formula-only mode via formula=null marker (for formula conflict detection)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };
  const cellKey = "Sheet1:0:0";

  // Establish base cell with a formula so the value write clears it via a marker.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("formula", "=0");
    cell.set("value", null);
    docA.getMap("cells").set(cellKey, cell);
  }, localOrigin);
  assert.equal(/** @type {any} */ (docB.getMap("cells").get(cellKey))?.get?.("formula"), "=0");

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    // Default mode is "formula" (formula-only).
    onConflict: (c) => conflicts.push(c)
  });

  // Offline concurrent edits:
  // - A writes a value using binder-style encoding (formula=null marker + value="ours"),
  //   but omits modifiedBy.
  // - B overwrites with a formula (and wins tie by clientID).
  disconnect();
  docA.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docA.getMap("cells").get(cellKey));
    cell.set("formula", null);
    cell.set("value", "ours");
    // Intentionally omit `modifiedBy` so conflict detection must rely on local-origin tracking.
  }, localOrigin);

  docB.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docB.getMap("cells").get(cellKey));
    cell.set("formula", "=1");
    cell.set("value", null);
    // Intentionally omit `modifiedBy`.
  });

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  assert.ok(conflicts.length >= 1, "expected at least one conflict to be detected");
  assert.equal(conflicts[0].kind, "formula");
  assert.equal(conflicts[0].localFormula.trim(), "");
  assert.equal(conflicts[0].remoteFormula.trim(), "=1");

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

