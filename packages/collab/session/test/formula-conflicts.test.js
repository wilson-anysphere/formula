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

/**
 * @param {number} ms
 */
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

test("CollabSession formula conflict monitor detects true offline concurrent edits via causality", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  // Deliberately tiny: causal conflict detection should still fire even if the
  // offline window far exceeds this old heuristic parameter.
  const concurrencyWindowMs = 5;

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
  // Wait longer than the old heuristic window to ensure we still detect the conflict.
  await new Promise((r) => setTimeout(r, concurrencyWindowMs + 150));
  await sessionB.setCellFormula("Sheet1:0:0", "=2");

  // Simulate a long offline period relative to the old heuristic.
  await sleep(concurrencyWindowMs + 25);

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.kind, "formula");
  assert.ok(conflictSide.formulaConflictMonitor?.resolveConflict(conflict.id, conflict.localFormula));

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, conflict.localFormula.trim());
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, conflict.localFormula.trim());

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession formula conflict monitor does not emit conflicts for sequential overwrites", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: {
      localUserId: "user-a",
      concurrencyWindowMs: 1,
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: {
      localUserId: "user-b",
      concurrencyWindowMs: 1,
      onConflict: (c) => conflictsB.push(c),
    },
  });

  await sessionA.setCellFormula("Sheet1:0:0", "=0");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=0");

  // Sequential: B sees A's edit before overwriting.
  await sessionA.setCellFormula("Sheet1:0:0", "=1");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=1");
  await sessionB.setCellFormula("Sheet1:0:0", "=2");
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, "=2");

  assert.equal(conflictsA.length, 0);
  assert.equal(conflictsB.length, 0);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession formula conflict monitor does not resurrect formulas on sequential deletes", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: {
      localUserId: "user-a",
      concurrencyWindowMs: 1,
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: {
      localUserId: "user-b",
      concurrencyWindowMs: 1,
      onConflict: (c) => conflictsB.push(c),
    },
  });

  // Establish the cell in both docs.
  await sessionA.setCellFormula("Sheet1:0:0", "=0");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=0");

  // Sequential: B sees A's formula before deleting it.
  await sessionA.setCellFormula("Sheet1:0:0", "=1");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=1");
  await sessionB.setCellFormula("Sheet1:0:0", null);

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, null);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, null);

  assert.equal(conflictsA.length, 0);
  assert.equal(conflictsB.length, 0);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession value conflicts surface when enabled (formula+value mode)", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const concurrencyWindowMs = 5;

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: {
      localUserId: "user-a",
      concurrencyWindowMs,
      mode: "formula+value",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: {
      localUserId: "user-b",
      concurrencyWindowMs,
      mode: "formula+value",
      onConflict: (c) => conflictsB.push(c),
    },
  });

  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  disconnect();
  await sessionA.setCellValue("Sheet1:0:0", "a");
  await sessionB.setCellValue("Sheet1:0:0", "b");
  await sleep(concurrencyWindowMs + 25);

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one value conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
  assert.equal(conflict.kind, "value");

  assert.ok(conflictSide.formulaConflictMonitor?.resolveConflict(conflict.id, conflict.localValue));

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, conflict.localValue);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, conflict.localValue);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
