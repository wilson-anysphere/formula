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
 * Cut/paste a single cell in a single transaction (delete source key and set
 * destination key to identical content).
 *
 * @param {import("../src/index.ts").CollabSession} session
 * @param {string} fromKey
 * @param {string} toKey
 */
function cutPaste(session, fromKey, toKey) {
  session.doc.transact(() => {
    const from = session.cells.get(fromKey);
    const fromMap = from instanceof Y.Map ? from : null;
    const value = fromMap?.get("value") ?? null;
    const formula = fromMap?.get("formula") ?? null;
    const format = fromMap?.get("format") ?? fromMap?.get("style") ?? null;
 
    const next = new Y.Map();
    next.set("value", value);
    if (formula) next.set("formula", formula);
    if (format) next.set("format", format);
 
    session.cells.set(toKey, next);
    session.cells.delete(fromKey);
  }, session.origin);
}
 
test("CellStructuralConflictMonitor detects concurrent move-destination conflicts and converges after resolution", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);
 
  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];
 
  const sessionA = createCollabSession({
    doc: docA,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });
 
  // Base cell at A1.
  sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "hello");
 
  // Offline concurrent moves of the same source to different destinations.
  disconnect();
  cutPaste(sessionA, "Sheet1:0:0", "Sheet1:0:1"); // A1 -> B1
  cutPaste(sessionB, "Sheet1:0:0", "Sheet1:0:2"); // A1 -> C1
 
  // Reconnect and sync.
  disconnect = connectDocs(docA, docB);
 
  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");
 
  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
 
  assert.equal(conflict.type, "move");
  assert.equal(conflict.reason, "move-destination");
 
  // Resolve by choosing B1.
  assert.ok(
    conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, {
      choice: "manual",
      to: "Sheet1:0:1",
    }),
  );
 
  assert.equal(sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(sessionB.getCell("Sheet1:0:0"), null);
  assert.equal(sessionA.getCell("Sheet1:0:1")?.value, "hello");
  assert.equal(sessionB.getCell("Sheet1:0:1")?.value, "hello");
  assert.equal(sessionA.getCell("Sheet1:0:2"), null);
  assert.equal(sessionB.getCell("Sheet1:0:2"), null);
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
test("CellStructuralConflictMonitor detects delete-vs-edit conflicts and converges after resolution", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);
 
  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];
 
  const sessionA = createCollabSession({
    doc: docA,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });
 
  sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "hello");
 
  disconnect();
  sessionA.doc.transact(() => {
    sessionA.cells.delete("Sheet1:0:0");
  }, sessionA.origin);
  sessionB.setCellValue("Sheet1:0:0", "world");
 
  disconnect = connectDocs(docA, docB);
 
  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");
 
  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
 
  assert.equal(conflict.type, "cell");
  assert.equal(conflict.reason, "delete-vs-edit");
 
  // Resolve by choosing deletion.
  assert.ok(conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, { choice: "manual", cell: null }));
 
  assert.equal(sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(sessionB.getCell("Sheet1:0:0"), null);
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
test("CellStructuralConflictMonitor auto-merges move vs edit by rewriting the edit to the destination", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);
 
  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];
 
  const sessionA = createCollabSession({
    doc: docA,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });
 
  sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "hello");
 
  disconnect();
  cutPaste(sessionA, "Sheet1:0:0", "Sheet1:0:1"); // A1 -> B1
  sessionB.setCellValue("Sheet1:0:0", "world"); // edit A1, does not touch B1
 
  disconnect = connectDocs(docA, docB);
 
  assert.equal(conflictsA.length, 0);
  assert.equal(conflictsB.length, 0);
 
  assert.equal(sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(sessionB.getCell("Sheet1:0:0"), null);
  assert.equal(sessionA.getCell("Sheet1:0:1")?.value, "world");
  assert.equal(sessionB.getCell("Sheet1:0:1")?.value, "world");
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
