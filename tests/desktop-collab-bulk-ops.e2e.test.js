import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import {
  VERSIONING_RESTORE_ORIGIN,
  createDesktopCellStructuralConflictMonitor,
  createDesktopFormulaConflictMonitor,
} from "../apps/desktop/src/collab/conflict-monitors.js";

test("Desktop wiring: version restore origin does not surface formula conflict UI", () => {
  // Deterministic map-overwrite tie-break: restored snapshot should "win".
  const doc = new Y.Doc();
  doc.clientID = 1;
  const cells = doc.getMap("cells");

  const sessionOrigin = { type: "session" };
  const binderOrigin = { type: "binder" };

  /** @type {any[]} */
  const conflicts = [];

  const monitor = createDesktopFormulaConflictMonitor({
    doc,
    cells,
    localUserId: "alice",
    sessionOrigin,
    binderOrigin,
    onConflict: (c) => conflicts.push(c),
    mode: "formula",
  });

  try {
    // Establish a baseline local formula.
    monitor.setLocalFormula("s:0:0", "=1");

    // Simulate a restore by applying a snapshot update that overwrites the cell,
    // tagging the transaction with the versioning restore origin.
    const restored = new Y.Doc();
    restored.clientID = 2;
    const restoredCells = restored.getMap("cells");
    const restoredCell = new Y.Map();
    restoredCell.set("formula", "=2");
    restoredCell.set("value", null);
    restoredCell.set("modified", Date.now());
    restoredCell.set("modifiedBy", "restorer");
    restoredCells.set("s:0:0", restoredCell);

    const restoreUpdate = Y.encodeStateAsUpdate(restored);
    Y.applyUpdate(doc, restoreUpdate, VERSIONING_RESTORE_ORIGIN);

    // Restore should not emit conflict events.
    assert.equal(conflicts.length, 0);
    assert.equal(monitor.listConflicts().length, 0);
  } finally {
    monitor.dispose();
  }
});

test("Desktop wiring: structural op log does not grow during version restore or branch checkout", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const sessionOrigin = { type: "session" };
  const binderOrigin = { type: "binder" };

  /** @type {any[]} */
  const conflicts = [];

  const monitor = createDesktopCellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "alice",
    sessionOrigin,
    binderOrigin,
    onConflict: (c) => conflicts.push(c),
  });

  try {
    const ops = doc.getMap("cellStructuralOps");
    assert.equal(ops.size, 0);

    // Simulate a branch checkout/merge (bulk apply) using session origin.
    doc.transact(() => {
      const cell = new Y.Map();
      cell.set("value", "bulk");
      cell.set("modifiedBy", "alice");
      cell.set("modified", Date.now());
      cells.set("Sheet1:0:0", cell);
    }, sessionOrigin);

    // Simulate a version restore using the restore origin.
    doc.transact(() => {
      cells.delete("Sheet1:0:0");
    }, VERSIONING_RESTORE_ORIGIN);

    // Neither bulk operation should be logged as a local structural op.
    assert.equal(ops.size, 0);
    assert.equal(conflicts.length, 0);

    // A normal DocumentController-driven edit (binder origin) *should* be logged.
    doc.transact(() => {
      const cell = new Y.Map();
      cell.set("value", "user-edit");
      cell.set("modifiedBy", "alice");
      cell.set("modified", Date.now());
      cells.set("Sheet1:0:1", cell);
    }, binderOrigin);

    assert.ok(ops.size > 0, "expected DocumentController edits to create structural op records");
  } finally {
    monitor.dispose();
  }
});
