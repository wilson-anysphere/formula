import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { yjsValueToJson } from "@formula/collab-yjs-utils";
import { FormulaConflictMonitor } from "../src/formula-conflict-monitor.js";

/**
 * @param {{ localClientId: number, remoteClientId: number }} ids
 */
function runScenario(ids) {
  const { localClientId, remoteClientId } = ids;

  const cellKey = "Sheet1:0:0";

  // Seed a baseline document using a third client id so local/remote edits both start at clock=0.
  const baseDoc = new Y.Doc();
  baseDoc.clientID = 1000;
  const baseCells = baseDoc.getMap("cells");
  const baseCell = new Y.Map();
  baseCells.set(cellKey, baseCell);
  baseCell.set("formula", null);
  baseCell.set("value", null);

  const baseUpdate = Y.encodeStateAsUpdate(baseDoc);
  const baseStateVector = Y.encodeStateVector(baseDoc);

  const localDoc = new Y.Doc();
  localDoc.clientID = localClientId;
  Y.applyUpdate(localDoc, baseUpdate);

  const cells = localDoc.getMap("cells");

  /** @type {any[]} */
  const conflicts = [];
  const localOrigin = { type: "local" };
  const monitor = new FormulaConflictMonitor({
    doc: localDoc,
    cells,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    ignoredOrigins: new Set(["versioning-restore"]),
    onConflict: (conflict) => conflicts.push(conflict),
  });

  monitor.setLocalFormula(cellKey, "=1");
  const trackedBefore = monitor._lastLocalFormulaEditByCellKey.get(cellKey);
  const trackedContentBefore = monitor._lastLocalContentEditByCellKey.get(cellKey);

  const remoteDoc = new Y.Doc();
  remoteDoc.clientID = remoteClientId;
  Y.applyUpdate(remoteDoc, baseUpdate);

  const remoteCells = remoteDoc.getMap("cells");
  const remoteCell = /** @type {any} */ (remoteCells.get(cellKey));

  remoteDoc.transact(() => {
    remoteCell.set("formula", "=2");
    remoteCell.set("value", 2);
    // Intentionally do not touch `modifiedBy` to simulate snapshot/restore operations.
  });

  const overwriteUpdate = Y.encodeStateAsUpdate(remoteDoc, baseStateVector);

  // Apply the overwrite as a "time travel" operation.
  Y.applyUpdate(localDoc, overwriteUpdate, "versioning-restore");

  const finalCell = /** @type {any} */ (cells.get(cellKey));
  const finalFormula = String(yjsValueToJson(finalCell?.get?.("formula") ?? ""));

  const trackedAfter = monitor._lastLocalFormulaEditByCellKey.get(cellKey);
  const trackedContentAfter = monitor._lastLocalContentEditByCellKey.get(cellKey);

  monitor.dispose();

  return { conflicts, finalFormula, trackedBefore, trackedAfter, trackedContentBefore, trackedContentAfter };
}

test("FormulaConflictMonitor ignores version restore origins", () => {
  // Yjs chooses a deterministic winner for concurrent map updates, but the ordering isn't
  // something we want to depend on here. Try both client-id orderings and pick the one
  // where the remote overwrite actually wins.
  const attemptA = runScenario({ localClientId: 1, remoteClientId: 2 });
  const attemptB = attemptA.finalFormula === "=2" ? null : runScenario({ localClientId: 2, remoteClientId: 1 });
  const result = attemptB ?? attemptA;

  // Ensure the overwrite applied so the test isn't vacuously passing.
  assert.equal(result.finalFormula, "=2");

  // The restore transaction should be ignored entirely (no conflict emission).
  assert.equal(result.conflicts.length, 0);

  // The restore transaction should also not pollute local-edit tracking (Task 38).
  assert.equal(result.trackedBefore?.formula, "=1");
  assert.equal(result.trackedAfter?.formula, "=1");
  assert.equal(result.trackedContentBefore?.kind, "formula");
  assert.equal(result.trackedContentBefore?.formula, "=1");
  assert.equal(result.trackedContentAfter?.kind, "formula");
  assert.equal(result.trackedContentAfter?.formula, "=1");
});
