import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";
import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor prunes old op records opportunistically on op log writes", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const ops = doc.getMap("cellStructuralOps");

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "local",
    onConflict: () => {},
    maxOpRecordAgeMs: 1_000,
  });

  // Constructor performs an initial forced prune and sets an internal throttle
  // timestamp; reset it so this test can exercise the on-write pruning hook
  // deterministically without waiting for real time.
  monitor._lastAgePruneAt = 0;

  const now = Date.now();
  const oldId = "op-old-write";
  doc.transact(() => {
    ops.set(oldId, {
      id: oldId,
      kind: "edit",
      userId: "remote",
      createdAt: now - 60_000,
      beforeState: [],
      afterState: [],
    });
  });

  assert.equal(ops.has(oldId), false);

  monitor.dispose();
  doc.destroy();
});

test("CellStructuralConflictMonitor prunes old shared op records by age when enabled", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const ops = doc.getMap("cellStructuralOps");

  const now = Date.now();
  const oldId = "op-old";
  const freshId = "op-fresh";

  doc.transact(() => {
    ops.set(oldId, {
      id: oldId,
      kind: "edit",
      userId: "user-old",
      createdAt: now - 60_000,
      beforeState: [],
      afterState: [],
    });
    ops.set(freshId, {
      id: freshId,
      kind: "edit",
      userId: "user-fresh",
      createdAt: now,
      beforeState: [],
      afterState: [],
    });
  });

  assert.equal(ops.size, 2);

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "user-local",
    onConflict: () => {},
    maxOpRecordAgeMs: 1_000,
  });

  assert.equal(ops.has(oldId), false);
  assert.equal(ops.has(freshId), true);

  monitor.dispose();
  doc.destroy();
});

test("CellStructuralConflictMonitor still detects recent conflicts when age pruning is enabled", () => {
  const cellKey = "Sheet1:0:0";

  // Seed a shared starting state.
  const docA = new Y.Doc();
  const cellsA = docA.getMap("cells");
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "seed");
    cell.set("formula", null);
    cellsA.set(cellKey, cell);
  });

  const docB = new Y.Doc();
  const cellsB = docB.getMap("cells");
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA));

  const originA = { type: "local-a" };
  const originB = { type: "local-b" };

  /** @type {Array<any>} */
  const conflictsA = [];

  const monitorA = new CellStructuralConflictMonitor({
    doc: docA,
    cells: cellsA,
    localUserId: "user-a",
    origin: originA,
    localOrigins: new Set([originA]),
    onConflict: (c) => conflictsA.push(c),
    maxOpRecordAgeMs: 60_000,
  });

  const monitorB = new CellStructuralConflictMonitor({
    doc: docB,
    cells: cellsB,
    localUserId: "user-b",
    origin: originB,
    localOrigins: new Set([originB]),
    onConflict: () => {},
    maxOpRecordAgeMs: 60_000,
  });

  // Make concurrent changes:
  // - user A deletes A1
  // - user B edits A1
  docA.transact(() => {
    cellsA.delete(cellKey);
  }, originA);

  docB.transact(() => {
    const cell = cellsB.get(cellKey);
    assert.ok(cell instanceof Y.Map);
    cell.set("value", "edited");
    cell.set("formula", null);
  }, originB);

  // Exchange updates to merge the concurrent edits.
  const updateA = Y.encodeStateAsUpdate(docA);
  const updateB = Y.encodeStateAsUpdate(docB);
  Y.applyUpdate(docA, updateB);
  Y.applyUpdate(docB, updateA);

  assert.equal(conflictsA.length > 0, true);
  assert.equal(
    conflictsA.some((c) => c.reason === "delete-vs-edit" && c.cellKey === cellKey),
    true,
  );

  monitorA.dispose();
  monitorB.dispose();
  docA.destroy();
  docB.destroy();
});

test("CellStructuralConflictMonitor age pruning is conservative relative to local queued ops", () => {
  const cellKey = "Sheet1:0:0";

  // Shared starting state.
  const docA = new Y.Doc();
  const cellsA = docA.getMap("cells");
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "seed");
    cell.set("formula", null);
    cellsA.set(cellKey, cell);
  });

  const docB = new Y.Doc();
  const cellsB = docB.getMap("cells");
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA));

  const originA = { type: "local-a" };
  const originB = { type: "local-b" };

  /** @type {Array<any>} */
  const conflictsA = [];

  const monitorA = new CellStructuralConflictMonitor({
    doc: docA,
    cells: cellsA,
    localUserId: "user-a",
    origin: originA,
    localOrigins: new Set([originA]),
    onConflict: (c) => conflictsA.push(c),
    // Extremely small age window to ensure age-based pruning would normally apply.
    maxOpRecordAgeMs: 1_000,
  });

  const monitorB = new CellStructuralConflictMonitor({
    doc: docB,
    cells: cellsB,
    localUserId: "user-b",
    origin: originB,
    localOrigins: new Set([originB]),
    onConflict: () => {},
    maxOpRecordAgeMs: 1_000,
  });

  // Create a local delete op (user A) and then "age" it by updating its createdAt
  // far into the past. This simulates clock skew / long-offline edits without
  // relying on a real time delay.
  docA.transact(() => {
    cellsA.delete(cellKey);
  }, originA);

  const opsA = docA.getMap("cellStructuralOps");
  assert.equal(opsA.size, 1);
  const [opId] = Array.from(opsA.keys());
  const record = opsA.get(opId);
  assert.ok(record);

  const agedCreatedAt = Date.now() - 60_000;
  docA.transact(() => {
    opsA.set(opId, { ...record, createdAt: agedCreatedAt });
  });

  // Force pruning. The aged record is older than the age cutoff, but should be
  // retained so we can still compare it against remote ops that arrive later.
  monitorA._pruneOpLogByAge({ force: true });
  assert.equal(opsA.has(opId), true);

  // Concurrent remote edit (user B) made without seeing A's delete.
  docB.transact(() => {
    const cell = cellsB.get(cellKey);
    assert.ok(cell instanceof Y.Map);
    cell.set("value", "edited");
    cell.set("formula", null);
  }, originB);

  // Merge the concurrent edits.
  const updateA = Y.encodeStateAsUpdate(docA);
  const updateB = Y.encodeStateAsUpdate(docB);
  Y.applyUpdate(docA, updateB);
  Y.applyUpdate(docB, updateA);

  assert.equal(
    conflictsA.some((c) => c.reason === "delete-vs-edit" && c.cellKey === cellKey),
    true,
  );

  monitorA.dispose();
  monitorB.dispose();
  docA.destroy();
  docB.destroy();
});
