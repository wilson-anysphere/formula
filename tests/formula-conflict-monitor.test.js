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
 */
function createClient(userId) {
  const doc = new Y.Doc();
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
