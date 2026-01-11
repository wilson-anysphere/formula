import test from "node:test";
import assert from "node:assert/strict";
 
import * as Y from "yjs";
 
import { REMOTE_ORIGIN } from "@formula/collab-undo";
 
import { createCollabSession } from "../src/index.ts";
 
/**
 * Duck-type Y.Map detection to avoid `instanceof` pitfalls when multiple Yjs
 * module instances are present (pnpm workspaces can produce this in Node).
 *
 * @param {any} value
 */
function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  if (value.constructor?.name !== "YMap") return false;
  if (typeof value.get !== "function") return false;
  if (typeof value.set !== "function") return false;
  if (typeof value.delete !== "function") return false;
  return true;
}
 
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
    const fromMap = isYMap(from) ? from : null;
    const value = fromMap?.get("value") ?? null;
    const formula = fromMap?.get("formula") ?? null;
    const enc = fromMap?.get("enc") ?? null;
    const format = fromMap?.get("format") ?? fromMap?.get("style") ?? null;
 
    const next = new Y.Map();
    if (enc) {
      next.set("enc", enc);
    } else {
      next.set("value", value);
      if (formula) next.set("formula", formula);
    }
    if (format) next.set("format", format);
 
    session.cells.set(toKey, next);
    session.cells.delete(fromKey);
  }, session.origin);
}
 
test("CellStructuralConflictMonitor detects concurrent move-destination conflicts and converges after resolution", async () => {
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
  await sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "hello");
 
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
 
  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal(await sessionA.getCell("Sheet1:0:2"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:2"), null);
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
test("CellStructuralConflictMonitor detects delete-vs-edit conflicts and converges after resolution", async () => {
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
 
  await sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "hello");
 
  disconnect();
  sessionA.doc.transact(() => {
    sessionA.cells.delete("Sheet1:0:0");
  }, sessionA.origin);
  await sessionB.setCellValue("Sheet1:0:0", "world");
 
  disconnect = connectDocs(docA, docB);
 
  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");
 
  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
 
  assert.equal(conflict.type, "cell");
  assert.equal(conflict.reason, "delete-vs-edit");
 
  // Resolve by choosing deletion.
  assert.ok(conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, { choice: "manual", cell: null }));
 
  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
test("CellStructuralConflictMonitor auto-merges move vs edit by rewriting the edit to the destination", async () => {
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
 
  await sessionA.setCellValue("Sheet1:0:0", "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "hello");
 
  disconnect();
  cutPaste(sessionA, "Sheet1:0:0", "Sheet1:0:1"); // A1 -> B1
  await sessionB.setCellValue("Sheet1:0:0", "world"); // edit A1, does not touch B1
 
  disconnect = connectDocs(docA, docB);
 
  assert.equal(conflictsA.length, 0);
  assert.equal(conflictsB.length, 0);
 
  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value, "world");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "world");
 
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CellStructuralConflictMonitor move conflict resolution applies the chosen side's moved cell content", async () => {
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
  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  disconnect();

  // A: edit A1 then move A1 -> B1 (single transaction).
  sessionA.doc.transact(() => {
    const a1 = sessionA.cells.get("Sheet1:0:0");
    assert.ok(isYMap(a1));
    a1.set("value", "from-a");
    const b1 = new Y.Map();
    b1.set("value", "from-a");
    sessionA.cells.set("Sheet1:0:1", b1);
    sessionA.cells.delete("Sheet1:0:0");
  }, sessionA.origin);

  // B: edit A1 then move A1 -> C1 (single transaction).
  sessionB.doc.transact(() => {
    const a1 = sessionB.cells.get("Sheet1:0:0");
    assert.ok(isYMap(a1));
    a1.set("value", "from-b");
    const c1 = new Y.Map();
    c1.set("value", "from-b");
    sessionB.cells.set("Sheet1:0:2", c1);
    sessionB.cells.delete("Sheet1:0:0");
  }, sessionB.origin);

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.type, "move");
  assert.equal(conflict.reason, "move-destination");

  const expectedKey = conflict.remote.toCellKey;
  const expectedValue = conflict.remote.cell?.value ?? null;
  const otherKey = conflict.local.toCellKey;
  assert.ok(typeof expectedKey === "string" && expectedKey.length > 0);
  assert.ok(typeof otherKey === "string" && otherKey.length > 0);
  assert.notEqual(expectedKey, otherKey);

  // Resolve by choosing "theirs" (remote side) and ensure that side's moved
  // content is the one that lands at the chosen destination.
  assert.ok(conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, { choice: "theirs" }));

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell(expectedKey))?.value ?? null, expectedValue);
  assert.equal((await sessionB.getCell(expectedKey))?.value ?? null, expectedValue);
  assert.equal(await sessionA.getCell(otherKey), null);
  assert.equal(await sessionB.getCell(otherKey), null);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
