import test from "node:test";
import assert from "node:assert/strict";

import { JSDOM } from "jsdom";
import * as Y from "yjs";

import { ConflictUiController, StructuralConflictUiController } from "../apps/desktop/src/collab/conflicts-ui/index.js";
import {
  BRANCHING_APPLY_ORIGIN,
  VERSIONING_RESTORE_ORIGIN,
  createDesktopCellStructuralConflictMonitor,
  createDesktopFormulaConflictMonitor,
} from "../apps/desktop/src/collab/conflict-monitors.js";

test("Desktop wiring: version restore origin does not surface formula conflict UI", () => {
  const dom = new JSDOM('<div id="root"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  // Deterministic map-overwrite tie-break: restored snapshot should "win".
  const doc = new Y.Doc();
  doc.clientID = 1;
  const cells = doc.getMap("cells");

  const sessionOrigin = { type: "session" };
  const binderOrigin = { type: "binder" };

  /** @type {any[]} */
  const conflicts = [];

  /** @type {ConflictUiController | null} */
  let ui = null;

  const monitor = createDesktopFormulaConflictMonitor({
    doc,
    cells,
    localUserId: "alice",
    sessionOrigin,
    binderOrigin,
    onConflict: (c) => {
      conflicts.push(c);
      ui?.addConflict(c);
    },
    mode: "formula",
  });

  try {
    const container = dom.window.document.getElementById("root");
    assert.ok(container);
    ui = new ConflictUiController({ container, monitor });

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
    assert.equal(container.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    monitor.dispose();
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});

test("Desktop wiring: branch checkout/merge origin does not surface formula conflict UI", () => {
  const dom = new JSDOM('<div id="root"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  // Deterministic map-overwrite tie-break: branch snapshot should "win".
  const doc = new Y.Doc();
  doc.clientID = 1;
  const cells = doc.getMap("cells");

  const sessionOrigin = { type: "session" };
  const binderOrigin = { type: "binder" };

  /** @type {any[]} */
  const conflicts = [];

  /** @type {ConflictUiController | null} */
  let ui = null;

  const monitor = createDesktopFormulaConflictMonitor({
    doc,
    cells,
    localUserId: "alice",
    sessionOrigin,
    binderOrigin,
    onConflict: (c) => {
      conflicts.push(c);
      ui?.addConflict(c);
    },
    mode: "formula",
  });

  try {
    const container = dom.window.document.getElementById("root");
    assert.ok(container);
    ui = new ConflictUiController({ container, monitor });

    // Establish a baseline local formula.
    monitor.setLocalFormula("s:0:0", "=1");

    // Simulate a branch checkout/merge by applying a snapshot update tagged with the
    // branching-apply origin (bulk rewrite).
    const branchDoc = new Y.Doc();
    branchDoc.clientID = 2;
    const branchCells = branchDoc.getMap("cells");
    const branchCell = new Y.Map();
    branchCell.set("formula", "=2");
    branchCell.set("value", null);
    branchCell.set("modified", Date.now());
    branchCell.set("modifiedBy", "branch");
    branchCells.set("s:0:0", branchCell);

    const branchUpdate = Y.encodeStateAsUpdate(branchDoc);
    Y.applyUpdate(doc, branchUpdate, BRANCHING_APPLY_ORIGIN);

    assert.equal(conflicts.length, 0);
    assert.equal(monitor.listConflicts().length, 0);
    assert.equal(container.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    monitor.dispose();
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});

test("Desktop wiring: structural op log does not grow during version restore or branch checkout", () => {
  const dom = new JSDOM('<div id="root"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const sessionOrigin = { type: "session" };
  const binderOrigin = { type: "binder" };

  /** @type {any[]} */
  const conflicts = [];

  /** @type {StructuralConflictUiController | null} */
  let ui = null;

  const monitor = createDesktopCellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "alice",
    sessionOrigin,
    binderOrigin,
    onConflict: (c) => {
      conflicts.push(c);
      ui?.addConflict(c);
    },
  });

  try {
    const container = dom.window.document.getElementById("root");
    assert.ok(container);
    ui = new StructuralConflictUiController({ container, monitor });

    const ops = doc.getMap("cellStructuralOps");
    assert.equal(ops.size, 0);

    // Simulate a branch checkout/merge (bulk apply) using the branching-apply origin.
    doc.transact(() => {
      const cell = new Y.Map();
      cell.set("value", "bulk");
      cell.set("modifiedBy", "alice");
      cell.set("modified", Date.now());
      cells.set("Sheet1:0:0", cell);
    }, BRANCHING_APPLY_ORIGIN);

    // Simulate a version restore using the restore origin.
    doc.transact(() => {
      cells.delete("Sheet1:0:0");
    }, VERSIONING_RESTORE_ORIGIN);

    // Neither bulk operation should be logged as a local structural op.
    assert.equal(ops.size, 0);
    assert.equal(conflicts.length, 0);
    assert.equal(container.querySelector('[data-testid="structural-conflict-toast"]'), null);

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
    ui?.destroy();
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});
