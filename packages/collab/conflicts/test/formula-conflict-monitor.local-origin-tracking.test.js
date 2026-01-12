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

test("FormulaConflictMonitor tracks local-origin formula clears for causal conflict detection (without setLocalFormula)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };
  const cellKey = "Sheet1:0:0";

  // Establish base cell map + formula so the conflict is an overwrite on the
  // formula key (not the cell map insertion).
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("formula", "=0");
    cell.set("value", null);
    docA.getMap("cells").set(cellKey, cell);
  }, localOrigin);
  assert.equal(
    /** @type {any} */ (docB.getMap("cells").get(cellKey))?.get?.("formula"),
    "=0",
    "expected base cell to sync to docB"
  );

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    onConflict: (c) => conflicts.push(c)
  });

  // Offline concurrent edits: A clears; B overwrites with a new formula (B wins).
  disconnect();
  docA.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docA.getMap("cells").get(cellKey));
    cell.set("formula", null);
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
  assert.equal(conflicts[0].localFormula.trim(), "");
  assert.equal(conflicts[0].remoteFormula.trim(), "=2");

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("FormulaConflictMonitor tracks local-origin value clears for value conflicts (binder-style order, formula+value mode)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };
  const cellKey = "Sheet1:0:0";

  // Establish base state with a literal value and no formula key present.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "base");
    docA.getMap("cells").set(cellKey, cell);
  }, localOrigin);

  const cellB = /** @type {any} */ (docB.getMap("cells").get(cellKey));
  assert.ok(cellB, "expected base cell map to sync to docB");
  assert.equal(cellB.get("value"), "base");
  assert.equal(cellB.get("formula"), undefined);

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    mode: "formula+value",
    onConflict: (c) => conflicts.push(c)
  });

  // Simulate offline concurrent edits:
  // - A clears the cell using binder-style ordering: formula=null marker, then value=null.
  // - B overwrites with a literal value (and omits modifiedBy).
  disconnect();
  docA.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docA.getMap("cells").get(cellKey));
    cell.set("formula", null);
    cell.set("value", null);
    // Intentionally omit `modifiedBy` so conflict detection must rely on causality.
  }, localOrigin);

  docB.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docB.getMap("cells").get(cellKey));
    cell.set("value", "theirs");
    // Intentionally omit `modifiedBy`.
  });

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const valueConflict = conflicts.find((c) => c.kind === "value") ?? null;
  assert.ok(valueConflict, "expected a value conflict to be detected");
  assert.equal(valueConflict.localValue, null);
  assert.equal(valueConflict.remoteValue, "theirs");

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("FormulaConflictMonitor tracks local-origin value edits for content conflicts (binder-style write, formula+value mode)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };
  const cellKey = "Sheet1:0:0";

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    mode: "formula+value",
    onConflict: (c) => conflicts.push(c)
  });

  // Establish a shared base cell map so concurrent edits race on map-entry overwrites,
  // not on the `cells[cellKey] = new Y.Map()` insertion.
  docA.transact(() => {
    docA.getMap("cells").set(cellKey, new Y.Map());
  }, localOrigin);
  assert.ok(docB.getMap("cells").get(cellKey), "expected base cell map to sync to docB");

  // Offline concurrent edits:
  // - A writes a literal value using binder-style encoding (formula=null marker + value=...).
  // - B writes a formula. B wins tie by clientID, so A should surface a content conflict.
  disconnect();
  docA.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docA.getMap("cells").get(cellKey));
    cell.set("formula", null);
    cell.set("value", "ours");
    // Intentionally omit `modifiedBy` so conflict detection must rely on causality.
  }, localOrigin);

  docB.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docB.getMap("cells").get(cellKey));
    cell.set("formula", "=1");
    cell.set("value", null);
    // Intentionally omit `modifiedBy`.
  });

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const contentConflict = conflicts.find((c) => c.kind === "content") ?? null;
  assert.ok(contentConflict, "expected a content conflict to be detected");
  assert.equal(contentConflict.local.type, "value");
  assert.equal(contentConflict.local.value, "ours");
  assert.equal(contentConflict.remote.type, "formula");
  assert.equal(contentConflict.remote.formula, "=1");

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("FormulaConflictMonitor local-origin tracking does not misclassify setLocalValue(null) clears on formula cells (formula+value mode)", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  const localOrigin = { type: "local-a" };
  const cellKey = "Sheet1:0:0";

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    mode: "formula+value",
    onConflict: (c) => conflicts.push(c)
  });

  // Establish base as a formula cell.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("formula", "=0");
    cell.set("value", null);
    docA.getMap("cells").set(cellKey, cell);
  }, localOrigin);

  const cellB = /** @type {any} */ (docB.getMap("cells").get(cellKey));
  assert.ok(cellB, "expected base cell map to sync to docB");
  assert.equal(cellB.get("formula"), "=0");

  // Offline concurrent edits:
  // - A clears via the monitor API `setLocalValue(null)` (this writes value first, then formula).
  // - B overwrites with a literal value (and omits modifiedBy).
  disconnect();
  monitor.setLocalValue(cellKey, null);

  docB.transact(() => {
    const cell = /** @type {Y.Map<any>} */ (docB.getMap("cells").get(cellKey));
    cell.set("value", "theirs");
    cell.set("formula", null);
    // Intentionally omit `modifiedBy`.
  });

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const valueConflict = conflicts.find((c) => c.kind === "value") ?? null;
  assert.ok(valueConflict, "expected a value conflict to be detected");
  assert.equal(valueConflict.localValue, null);
  assert.equal(valueConflict.remoteValue, "theirs");
  assert.equal(
    conflicts.some((c) => c.kind === "content"),
    false,
    "expected clear-vs-value overwrite to be treated as a value conflict, not a content conflict"
  );

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});
