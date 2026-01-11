import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { createCollabSession } from "../src/index.ts";

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

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

test("CollabSession formula conflict monitor detects offline concurrent edits and converges after resolution", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const concurrencyWindowMs = 50;

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: {
      localUserId: "user-a",
      concurrencyWindowMs,
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: {
      localUserId: "user-b",
      concurrencyWindowMs,
      onConflict: (c) => conflictsB.push(c),
    },
  });

  // Establish a shared base cell map so concurrent edits race on the formula key
  // (not on the `cells[cellKey] = new Y.Map()` insertion).
  await sessionA.setCellFormula("Sheet1:0:0", "=0");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=0");

  // Simulate offline concurrent edits (same cell, different formulas).
  disconnect();
  await sessionA.setCellFormula("Sheet1:0:0", "=1");
  await sessionB.setCellFormula("Sheet1:0:0", "=2");

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  // Wait for the other side's concurrency window to elapse so the resolution
  // doesn't get misclassified as a concurrent overwrite.
  await new Promise((r) => setTimeout(r, concurrencyWindowMs + 25));

  assert.ok(conflictSide.formulaConflictMonitor?.resolveConflict(conflict.id, conflict.localFormula));

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, conflict.localFormula.trim());
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, conflict.localFormula.trim());

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
