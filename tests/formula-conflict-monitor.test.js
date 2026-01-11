import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createUndoService, REMOTE_ORIGIN } from "../packages/collab/undo/index.js";
import { FormulaConflictMonitor } from "../packages/collab/conflicts/src/formula-conflict-monitor.js";

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function syncDocs(docA, docB) {
  // Exchange incremental updates until both docs converge. This matters for
  // auto-resolvers that may write new local changes in response to a remote
  // update (e.g. extension/subset preference).
  for (let i = 0; i < 10; i += 1) {
    let changed = false;
    const updateA = Y.encodeStateAsUpdate(docA, Y.encodeStateVector(docB));
    if (updateA.length > 0) {
      Y.applyUpdate(docB, updateA, REMOTE_ORIGIN);
      changed = true;
    }
    const updateB = Y.encodeStateAsUpdate(docB, Y.encodeStateVector(docA));
    if (updateB.length > 0) {
      Y.applyUpdate(docA, updateB, REMOTE_ORIGIN);
      changed = true;
    }
    if (!changed) break;
  }
}

/**
 * @param {string} userId
 * @param {{ clientID?: number, mode?: "formula" | "formula+value" }} [opts]
 */
function createClient(userId, opts = {}) {
  const doc = new Y.Doc();
  if (typeof opts.clientID === "number") {
    doc.clientID = opts.clientID;
  }
  const cells = doc.getMap("cells");
  const origin = { type: "local", userId };
  const undo = createUndoService({ mode: "collab", doc, scope: cells, origin, captureTimeoutMs: 10_000 });

  /** @type {Array<any>} */
  const conflicts = [];
  const monitor = new FormulaConflictMonitor({
    doc,
    cells,
    localUserId: userId,
    origin,
    localOrigins: undo.localOrigins,
    mode: opts.mode,
    onConflict: (c) => conflicts.push(c)
  });

  return { doc, cells, undo, monitor, conflicts };
}

/**
 * @param {Y.Map<any>} cells
 * @param {string} key
 */
function getFormula(cells, key) {
  const cell = /** @type {Y.Map<any>|undefined} */ (cells.get(key));
  return (cell?.get("formula") ?? "").toString();
}

/**
 * @param {Y.Map<any>} cells
 * @param {string} key
 */
function getValue(cells, key) {
  const cell = /** @type {Y.Map<any>|undefined} */ (cells.get(key));
  return cell?.get("value") ?? null;
}

test("concurrent same-cell edits surface a conflict and converge after resolution", () => {
  const a = createClient("alice");
  const b = createClient("bob");

  // Establish a common base formula.
  a.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(a.doc, b.doc);

  // Simulate offline concurrent edits.
  a.monitor.setLocalFormula("s:0:0", "=1+1");
  b.monitor.setLocalFormula("s:0:0", "=1*2");

  // Reconnect.
  syncDocs(a.doc, b.doc);

  const allConflicts = [...a.conflicts, ...b.conflicts];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = a.conflicts.length > 0 ? a : b;
  const conflict = conflictSide.conflicts[0];

  // Resolve by keeping the local version for the side that detected conflict.
  assert.ok(conflictSide.monitor.resolveConflict(conflict.id, conflict.localFormula));
  syncDocs(a.doc, b.doc);

  assert.equal(getFormula(a.cells, "s:0:0"), conflict.localFormula.trim());
  assert.equal(getFormula(b.cells, "s:0:0"), conflict.localFormula.trim());
});

test("AST-equivalent concurrent edits auto-resolve without surfacing UI", () => {
  const a = createClient("alice");
  const b = createClient("bob");

  a.monitor.setLocalFormula("s:0:0", "");
  syncDocs(a.doc, b.doc);

  a.monitor.setLocalFormula("s:0:0", "=SUM(A1:A2)");
  b.monitor.setLocalFormula("s:0:0", "=sum(a1:a2)");
  syncDocs(a.doc, b.doc);

  assert.equal(a.conflicts.length, 0);
  assert.equal(b.conflicts.length, 0);
});

test("extension/subset concurrent edits auto-resolve to the extension", () => {
  const a = createClient("alice");
  const b = createClient("bob");

  a.monitor.setLocalFormula("s:0:0", "=A1");
  syncDocs(a.doc, b.doc);

  // Concurrent: alice writes subset, bob writes extension.
  a.monitor.setLocalFormula("s:0:0", "=A1+1");
  b.monitor.setLocalFormula("s:0:0", "=A1+1+1");
  syncDocs(a.doc, b.doc);

  assert.equal(a.conflicts.length, 0);
  assert.equal(b.conflicts.length, 0);
  assert.equal(getFormula(a.cells, "s:0:0"), "=A1+1+1");
  assert.equal(getFormula(b.cells, "s:0:0"), "=A1+1+1");
});

test("sequential deletes do not resurrect formulas or surface conflicts", () => {
  const a = createClient("alice");
  const b = createClient("bob");

  // Alice writes a formula and Bob syncs it.
  a.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(a.doc, b.doc);

  // Bob deletes after seeing Alice's write (sequential overwrite, not a conflict).
  b.monitor.setLocalFormula("s:0:0", "");
  syncDocs(a.doc, b.doc);

  assert.equal(a.conflicts.length, 0);
  assert.equal(b.conflicts.length, 0);
  assert.equal(getFormula(a.cells, "s:0:0"), "");
  assert.equal(getFormula(b.cells, "s:0:0"), "");
});

test("concurrent delete vs overwrite surfaces a conflict", () => {
  const a = createClient("alice");
  const b = createClient("bob");

  // Establish a common base formula.
  a.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(a.doc, b.doc);

  // Offline concurrent edits: alice deletes, bob overwrites.
  a.monitor.setLocalFormula("s:0:0", "");
  b.monitor.setLocalFormula("s:0:0", "=2");

  syncDocs(a.doc, b.doc);

  const allConflicts = [...a.conflicts, ...b.conflicts];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = a.conflicts.length > 0 ? a : b;
  const conflict = conflictSide.conflicts[0];

  assert.equal(conflict.kind, "formula");
  assert.ok([conflict.localFormula.trim(), conflict.remoteFormula.trim()].includes(""), "expected one side of conflict to be empty");
  assert.ok(
    [conflict.localFormula.trim(), conflict.remoteFormula.trim()].some((f) => f.startsWith("=2") || f === "=2"),
    "expected one side of conflict to be the overwrite"
  );
});

test("sequential remote deletes that use key deletion do not surface conflicts", () => {
  const alice = createClient("alice");

  // Simulate a legacy client that clears formulas by deleting the key entirely.
  const docB = new Y.Doc();
  const cellsB = docB.getMap("cells");

  alice.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(alice.doc, docB);

  const cellB = /** @type {Y.Map<any>} */ (cellsB.get("s:0:0"));
  assert.ok(cellB);

  docB.transact(() => {
    cellB.delete("formula");
    cellB.set("modifiedBy", "bob");
    cellB.set("modified", Date.now());
  });

  syncDocs(alice.doc, docB);

  assert.equal(alice.conflicts.length, 0);
  assert.equal(getFormula(alice.cells, "s:0:0"), "");
});

test("sequential remote deletes that use key deletion do not surface value conflicts", () => {
  const alice = createClient("alice", { mode: "formula+value" });

  // Simulate a legacy client that clears values by deleting the key entirely.
  const docB = new Y.Doc();
  const cellsB = docB.getMap("cells");

  alice.monitor.setLocalValue("s:0:0", "x");
  syncDocs(alice.doc, docB);

  const cellB = /** @type {Y.Map<any>} */ (cellsB.get("s:0:0"));
  assert.ok(cellB);

  docB.transact(() => {
    cellB.delete("value");
    cellB.set("modifiedBy", "bob");
    cellB.set("modified", Date.now());
  });

  syncDocs(alice.doc, docB);

  assert.equal(alice.conflicts.length, 0);
  assert.equal(getValue(alice.cells, "s:0:0"), null);
});

test("concurrent value vs formula surfaces a content conflict on the value writer (and choosing remote does not clobber the formula)", () => {
  // Ensure deterministic tie-breaking for concurrent map-entry overwrites:
  // higher clientID wins in Yjs.
  const alice = createClient("alice", { clientID: 2, mode: "formula+value" });
  const bob = createClient("bob", { clientID: 1, mode: "formula+value" });

  // Establish a shared base cell map so concurrent edits race on keys (not on cell insertion).
  alice.monitor.setLocalValue("s:0:0", "base");
  syncDocs(alice.doc, bob.doc);

  // Offline concurrent edits: alice sets a formula; bob sets a literal value.
  // Alice wins both formula + value keys (clientID 2 > 1), so bob is the "losing"
  // side and should see a content conflict (value vs formula).
  alice.monitor.setLocalFormula("s:0:0", "=1");
  bob.monitor.setLocalValue("s:0:0", "bob");

  syncDocs(alice.doc, bob.doc);

  const conflict = bob.conflicts.find((c) => c.kind === "content") ?? null;
  assert.ok(conflict, "expected a content conflict on bob");
  assert.equal(conflict.remote.type, "formula");
  assert.equal(conflict.remote.formula, "=1");
  assert.equal(conflict.local.type, "value");
  assert.equal(conflict.local.value, "bob");

  // The concurrently-written formula should be present in the doc at conflict time.
  assert.equal(getFormula(bob.cells, "s:0:0"), "=1");

  // Choosing the remote formula is a no-op (remote state is already applied) and must not
  // clear the formula via setLocalValue().
  assert.ok(bob.monitor.resolveConflict(conflict.id, conflict.remote));
  syncDocs(alice.doc, bob.doc);

  assert.equal(getFormula(alice.cells, "s:0:0"), "=1");
  assert.equal(getFormula(bob.cells, "s:0:0"), "=1");
});

test("concurrent value vs formula surfaces a content conflict on the formula writer when the value writer wins", () => {
  // Value writer has higher clientID, so their formula=null write wins and clears
  // the concurrent formula.
  const alice = createClient("alice", { clientID: 1, mode: "formula+value" });
  const bob = createClient("bob", { clientID: 2, mode: "formula+value" });

  alice.monitor.setLocalValue("s:0:0", "base");
  syncDocs(alice.doc, bob.doc);

  alice.monitor.setLocalFormula("s:0:0", "=1");
  bob.monitor.setLocalValue("s:0:0", "bob");
  syncDocs(alice.doc, bob.doc);

  const conflict = alice.conflicts.find((c) => c.kind === "content") ?? null;
  assert.ok(conflict, "expected a content conflict on alice");
  assert.equal(conflict.local.type, "formula");
  assert.equal(conflict.local.formula.trim(), "=1");
  assert.equal(conflict.remote.type, "value");
  assert.equal(conflict.remote.value, "bob");

  // Choosing the remote value is a no-op and must preserve the remote value.
  assert.ok(alice.monitor.resolveConflict(conflict.id, conflict.remote));
  syncDocs(alice.doc, bob.doc);
  assert.equal(getFormula(alice.cells, "s:0:0"), "");
  assert.equal(getFormula(bob.cells, "s:0:0"), "");
  assert.equal(getValue(alice.cells, "s:0:0"), "bob");
  assert.equal(getValue(bob.cells, "s:0:0"), "bob");
});

test("legacy value writes that do not create a formula marker still surface a content conflict (and resolving remote value clears the formula)", () => {
  // Alice uses the monitor in formula+value mode, but Bob is a legacy client that
  // sets `value` without creating a `formula=null` marker. If Bob doesn't have the
  // formula key locally, `delete(\"formula\")` is a no-op in Yjs, so the concurrent
  // formula insert can survive and the cell may contain both `formula` and `value`.
  const alice = createClient("alice", { clientID: 1, mode: "formula+value" });
  const bobDoc = new Y.Doc();
  bobDoc.clientID = 2;
  const bobCells = bobDoc.getMap("cells");

  // Establish a shared base cell map with only a literal value (no formula key).
  alice.doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "base");
    cell.set("modifiedBy", "alice");
    cell.set("modified", Date.now());
    alice.cells.set("s:0:0", cell);
  }, alice.monitor.origin);
  syncDocs(alice.doc, bobDoc);

  const bobCell = /** @type {Y.Map<any>} */ (bobCells.get("s:0:0"));
  assert.ok(bobCell, "expected base cell map on bob");
  assert.equal(bobCell.get("formula"), undefined);

  // Offline concurrent edits: alice writes a formula (also sets value=null); bob writes a value
  // and attempts to delete formula (no-op because it doesn't exist on bob).
  alice.monitor.setLocalFormula("s:0:0", "=1");
  bobDoc.transact(() => {
    bobCell.set("value", "bob");
    bobCell.delete("formula");
    bobCell.set("modifiedBy", "bob");
    bobCell.set("modified", Date.now());
  });
  syncDocs(alice.doc, bobDoc);

  const conflict = alice.conflicts.find((c) => c.kind === "content") ?? null;
  assert.ok(conflict, "expected a content conflict on alice");
  assert.equal(conflict.local.type, "formula");
  assert.equal(conflict.local.formula, "=1");
  assert.equal(conflict.remote.type, "value");
  assert.equal(conflict.remote.value, "bob");

  // Choosing the remote value must clear the formula even if the value is already applied.
  assert.ok(alice.monitor.resolveConflict(conflict.id, conflict.remote));
  syncDocs(alice.doc, bobDoc);

  assert.equal(getFormula(alice.cells, "s:0:0"), "");
  assert.equal(getValue(alice.cells, "s:0:0"), "bob");
  assert.equal(getFormula(bobCells, "s:0:0"), "");
  assert.equal(getValue(bobCells, "s:0:0"), "bob");
});

test("concurrent formula clear vs value edit surfaces a content conflict", () => {
  // Value writer has higher clientID and wins both formula/value null markers.
  const alice = createClient("alice", { clientID: 1, mode: "formula+value" });
  const bob = createClient("bob", { clientID: 2, mode: "formula+value" });

  // Establish a shared base formula.
  alice.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(alice.doc, bob.doc);

  // Offline concurrent edits: alice clears the formula, bob writes a value.
  alice.monitor.setLocalFormula("s:0:0", "");
  bob.monitor.setLocalValue("s:0:0", "bob");
  syncDocs(alice.doc, bob.doc);

  const conflict = alice.conflicts.find((c) => c.kind === "content") ?? null;
  assert.ok(conflict, "expected a content conflict on alice");
  assert.equal(conflict.local.type, "formula");
  assert.equal(conflict.local.formula.trim(), "");
  assert.equal(conflict.remote.type, "value");
  assert.equal(conflict.remote.value, "bob");

  // Resolve by keeping the local clear (should clear the value).
  assert.ok(alice.monitor.resolveConflict(conflict.id, conflict.local));
  syncDocs(alice.doc, bob.doc);

  assert.equal(getFormula(alice.cells, "s:0:0"), "");
  assert.equal(getFormula(bob.cells, "s:0:0"), "");
  assert.equal(getValue(alice.cells, "s:0:0"), null);
  assert.equal(getValue(bob.cells, "s:0:0"), null);
});

test("sequential value -> formula edits do not surface a content conflict", () => {
  const alice = createClient("alice", { mode: "formula+value" });
  const bob = createClient("bob", { mode: "formula+value" });

  alice.monitor.setLocalValue("s:0:0", "x");
  syncDocs(alice.doc, bob.doc);

  // Bob sees Alice's value write (including the formula=null marker) before applying the formula.
  bob.monitor.setLocalFormula("s:0:0", "=1");
  syncDocs(alice.doc, bob.doc);

  assert.equal(alice.conflicts.length, 0);
  assert.equal(bob.conflicts.length, 0);
});
