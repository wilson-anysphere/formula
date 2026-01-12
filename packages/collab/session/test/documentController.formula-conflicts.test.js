import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";
import { bindCollabSessionToDocumentController, createCollabSession } from "../src/index.ts";

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

async function waitForCondition(fn, timeoutMs = 2000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const ok = await fn();
      if (ok) return;
    } catch {
      // ignore while polling
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

test("FormulaConflictMonitor detects formula-vs-formula conflicts for DocumentController+binder edits (formula mode)", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: { localUserId: "user-a", onConflict: (c) => conflictsA.push(c), mode: "formula" },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: { localUserId: "user-b", onConflict: (c) => conflictsB.push(c), mode: "formula" },
  });
  sessionA.setPermissions({ role: "editor", userId: "user-a", rangeRestrictions: [] });
  sessionB.setPermissions({ role: "editor", userId: "user-b", rangeRestrictions: [] });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({
    session: sessionA,
    documentController: dcA,
    userId: "user-a",
    formulaConflictsMode: "formula",
  });
  const binderB = await bindCollabSessionToDocumentController({
    session: sessionB,
    documentController: dcB,
    userId: "user-b",
    formulaConflictsMode: "formula",
  });

  // Establish a shared base cell map before simulating offline concurrent edits.
  dcA.setCellFormula("Sheet1", "A1", "=0");
  await waitForCondition(() => dcB.getCell("Sheet1", "A1").formula === "=0");

  // Offline concurrent edits (same cell, different formulas).
  disconnect();
  dcA.setCellFormula("Sheet1", "A1", "=1");
  dcB.setCellFormula("Sheet1", "A1", "=2");
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.formula === "=1");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.formula === "=2");

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  await waitForCondition(() => conflictsA.length + conflictsB.length >= 1);

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.kind, "formula");
  assert.ok(conflictSide.formulaConflictMonitor?.resolveConflict(conflict.id, conflict.localFormula));

  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.formula === conflict.localFormula.trim());
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, conflict.localFormula.trim());

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("FormulaConflictMonitor detects concurrent delete-vs-overwrite for DocumentController+binder edits (formula mode)", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: { localUserId: "user-a", onConflict: (c) => conflictsA.push(c), mode: "formula" },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: { localUserId: "user-b", onConflict: (c) => conflictsB.push(c), mode: "formula" },
  });
  sessionA.setPermissions({ role: "editor", userId: "user-a", rangeRestrictions: [] });
  sessionB.setPermissions({ role: "editor", userId: "user-b", rangeRestrictions: [] });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({
    session: sessionA,
    documentController: dcA,
    userId: "user-a",
    formulaConflictsMode: "formula",
  });
  const binderB = await bindCollabSessionToDocumentController({
    session: sessionB,
    documentController: dcB,
    userId: "user-b",
    formulaConflictsMode: "formula",
  });

  // Establish base.
  dcA.setCellFormula("Sheet1", "A1", "=1");
  await waitForCondition(() => dcB.getCell("Sheet1", "A1").formula === "=1");

  // Offline concurrent edits: A clears, B overwrites.
  disconnect();
  dcA.setCellFormula("Sheet1", "A1", null);
  dcB.setCellFormula("Sheet1", "A1", "=2");
  // `CollabSession.getCell()` intentionally treats marker-only cells (value/formula both null)
  // as empty state, returning null even though a `formula=null` marker is preserved in Yjs.
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0")) === null);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.formula === "=2");

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  await waitForCondition(() => conflictsA.length + conflictsB.length >= 1);
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
  assert.equal(conflict.kind, "formula");
  assert.ok(
    [conflict.localFormula.trim(), conflict.remoteFormula.trim()].includes(""),
    "expected one side of conflict to be empty",
  );

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("FormulaConflictMonitor detects formula-vs-value content conflicts for DocumentController+binder edits (formula+value mode)", async () => {
  // Ensure deterministic concurrent tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: { localUserId: "user-a", onConflict: (c) => conflictsA.push(c), mode: "formula+value" },
  });
  const sessionB = createCollabSession({
    doc: docB,
    formulaConflicts: { localUserId: "user-b", onConflict: (c) => conflictsB.push(c), mode: "formula+value" },
  });
  sessionA.setPermissions({ role: "editor", userId: "user-a", rangeRestrictions: [] });
  sessionB.setPermissions({ role: "editor", userId: "user-b", rangeRestrictions: [] });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({
    session: sessionA,
    documentController: dcA,
    userId: "user-a",
    formulaConflictsMode: "formula+value",
  });
  const binderB = await bindCollabSessionToDocumentController({
    session: sessionB,
    documentController: dcB,
    userId: "user-b",
    formulaConflictsMode: "formula+value",
  });

  // Establish base.
  dcA.setCellValue("Sheet1", "A1", "base");
  await waitForCondition(() => dcB.getCell("Sheet1", "A1").value === "base");

  // Offline concurrent edits: A writes a value, B writes a formula.
  disconnect();
  dcA.setCellValue("Sheet1", "A1", "ours");
  dcB.setCellFormula("Sheet1", "A1", "=1");
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.value === "ours");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.formula === "=1");

  disconnect = connectDocs(docA, docB);

  await waitForCondition(() => conflictsA.length + conflictsB.length >= 1);

  const conflict = [...conflictsA, ...conflictsB].find((c) => c.kind === "content") ?? null;
  assert.ok(conflict, "expected a content conflict");
  assert.ok(
    (conflict.local.type === "value" && conflict.remote.type === "formula") ||
      (conflict.local.type === "formula" && conflict.remote.type === "value"),
    "expected a value-vs-formula content conflict",
  );

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
