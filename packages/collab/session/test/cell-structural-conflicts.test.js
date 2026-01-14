import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { getYMap } from "@formula/collab-yjs-utils";
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
    const fromMap = getYMap(from);
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

/**
 * Cut/paste a single cell in a single transaction, overriding the moved cell's
 * format. The source cell is updated to match the format before deleting so the
 * structural monitor still infers a move (delete at X + add at Y with identical
 * fingerprint).
 *
 * @param {import("../src/index.ts").CollabSession} session
 * @param {string} fromKey
 * @param {string} toKey
 * @param {any} format
 */
function cutPasteWithFormat(session, fromKey, toKey, format) {
  session.doc.transact(() => {
    const from = session.cells.get(fromKey);
    const fromMap = getYMap(from);
    const value = fromMap?.get("value") ?? null;
    const formula = fromMap?.get("formula") ?? null;
    const enc = fromMap?.get("enc") ?? null;

    if (fromMap) {
      fromMap.delete("style");
      if (format != null) {
        fromMap.set("format", format);
      } else {
        fromMap.delete("format");
      }
    }

    const next = new Y.Map();
    if (enc) {
      next.set("enc", enc);
    } else {
      next.set("value", value);
      if (formula) next.set("formula", formula);
    }
    if (format != null) next.set("format", format);

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

test("CellStructuralConflictMonitor supports manual move conflict resolution with overridden cell content", async () => {
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
  cutPaste(sessionB, "Sheet1:0:0", "Sheet1:0:2"); // A1 -> C1

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];
  assert.equal(conflict.type, "move");
  assert.equal(conflict.reason, "move-destination");

  assert.ok(
    conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, {
      choice: "manual",
      to: "Sheet1:0:1",
      cell: { value: "manual" },
    }),
  );

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value, "manual");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "manual");
  assert.equal(await sessionA.getCell("Sheet1:0:2"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:2"), null);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession cellConflicts.maxOpRecordAgeMs prunes old persisted op log records on startup", () => {
  const doc = new Y.Doc();
  const ops = doc.getMap("cellStructuralOps");
  const now = Date.now();

  doc.transact(() => {
    ops.set("old-op", {
      id: "old-op",
      kind: "edit",
      userId: "someone",
      createdAt: now - 60_000,
      beforeState: [],
      afterState: [],
    });
  });

  const session = createCollabSession({
    doc,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: () => {},
      maxOpRecordAgeMs: 1_000,
    },
  });

  assert.equal(doc.getMap("cellStructuralOps").has("old-op"), false);

  session.destroy();
  doc.destroy();
});

test("CellStructuralConflictMonitor detects conflicts after re-instantiating a session with pre-existing local op log entries", async () => {
  const docA1 = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA1, docB);

  /** @type {Array<any>} */
  const conflictsA2 = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA1 = createCollabSession({
    doc: docA1,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: () => {},
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });

  await sessionA1.setCellValue("Sheet1:0:0", "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "hello");

  // Go offline.
  disconnect();

  // Local offline move.
  cutPaste(sessionA1, "Sheet1:0:0", "Sheet1:0:1"); // A1 -> B1
  // Concurrent remote move to another destination.
  cutPaste(sessionB, "Sheet1:0:0", "Sheet1:0:2"); // A1 -> C1

  // Simulate closing/reopening the app by copying doc state into a new Y.Doc
  // and creating a new session/monitor instance.
  const docA2 = new Y.Doc();
  Y.applyUpdate(docA2, Y.encodeStateAsUpdate(docA1));

  sessionA1.destroy();
  docA1.destroy();

  const sessionA2 = createCollabSession({
    doc: docA2,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA2.push(c),
    },
  });

  // Reconnect and sync docA2 with docB.
  disconnect = connectDocs(docA2, docB);

  const allConflicts = [...conflictsA2, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA2.length > 0 ? sessionA2 : sessionB;
  const conflict = conflictsA2.length > 0 ? conflictsA2[0] : conflictsB[0];

  assert.equal(conflict.type, "move");
  assert.equal(conflict.reason, "move-destination");

  assert.ok(
    conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, {
      choice: "manual",
      to: "Sheet1:0:1",
    }),
  );

  assert.equal(await sessionA2.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA2.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal(await sessionA2.getCell("Sheet1:0:2"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:2"), null);

  sessionA2.destroy();
  sessionB.destroy();
  disconnect();
  docA2.destroy();
  docB.destroy();
});

test("CellStructuralConflictMonitor bootstraps conflicts from an existing shared op log", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  const sessionA1 = createCollabSession({
    doc: docA,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: () => {},
    },
  });
  const sessionB1 = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: () => {},
    },
  });

  await sessionA1.setCellValue("Sheet1:0:0", "hello");
  assert.equal((await sessionB1.getCell("Sheet1:0:0"))?.value, "hello");

  disconnect();
  cutPaste(sessionA1, "Sheet1:0:0", "Sheet1:0:1"); // A1 -> B1
  cutPaste(sessionB1, "Sheet1:0:0", "Sheet1:0:2"); // A1 -> C1

  // Destroy sessions before syncing so no conflict monitor runs during the initial merge.
  sessionA1.destroy();
  sessionB1.destroy();

  // Reconnect and sync docs (op log entries for both sides now exist in both docs).
  disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA2 = [];
  /** @type {Array<any>} */
  const conflictsB2 = [];

  const sessionA2 = createCollabSession({
    doc: docA,
    cellConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA2.push(c),
    },
  });
  const sessionB2 = createCollabSession({
    doc: docB,
    cellConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB2.push(c),
    },
  });

  const allConflicts = [...conflictsA2, ...conflictsB2];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA2.length > 0 ? sessionA2 : sessionB2;
  const conflict = conflictsA2.length > 0 ? conflictsA2[0] : conflictsB2[0];

  assert.equal(conflict.type, "move");
  assert.equal(conflict.reason, "move-destination");

  assert.ok(
    conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, {
      choice: "manual",
      to: "Sheet1:0:1",
    }),
  );

  assert.equal(await sessionA2.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB2.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA2.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal((await sessionB2.getCell("Sheet1:0:1"))?.value, "hello");
  assert.equal(await sessionA2.getCell("Sheet1:0:2"), null);
  assert.equal(await sessionB2.getCell("Sheet1:0:2"), null);

  sessionA2.destroy();
  sessionB2.destroy();
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

test("CellStructuralConflictMonitor emits a content conflict when a move destination is concurrently edited", async () => {
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

  // A moves A1 -> B1 while offline.
  cutPaste(sessionA, "Sheet1:0:0", "Sheet1:0:1");
  // B concurrently types into B1 (was empty in the shared base).
  await sessionB.setCellValue("Sheet1:0:1", "world");

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.type, "cell");
  assert.equal(conflict.reason, "content");
  assert.equal(conflict.cellKey, "Sheet1:0:1");

  const expected = conflict.local?.after?.value ?? null;
  assert.ok(conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, { choice: "ours" }));

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value ?? null, expected);
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value ?? null, expected);

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
    assert.ok(getYMap(a1));
    a1.set("value", "from-a");
    const b1 = new Y.Map();
    b1.set("value", "from-a");
    sessionA.cells.set("Sheet1:0:1", b1);
    sessionA.cells.delete("Sheet1:0:0");
  }, sessionA.origin);

  // B: edit A1 then move A1 -> C1 (single transaction).
  sessionB.doc.transact(() => {
    const a1 = sessionB.cells.get("Sheet1:0:0");
    assert.ok(getYMap(a1));
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

test("CellStructuralConflictMonitor surfaces content conflicts when two users move the same cell to the same destination with different content", async () => {
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

  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  disconnect();

  // Diverge content offline, then both move A1 -> B1.
  await sessionA.setCellValue("Sheet1:0:0", "from-a");
  await sessionB.setCellValue("Sheet1:0:0", "from-b");
  cutPaste(sessionA, "Sheet1:0:0", "Sheet1:0:1");
  cutPaste(sessionB, "Sheet1:0:0", "Sheet1:0:1");

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.type, "cell");
  assert.equal(conflict.reason, "content");
  assert.equal(conflict.cellKey, "Sheet1:0:1");

  const expected = conflict.local?.after?.value ?? null;
  assert.ok(
    conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, {
      choice: "ours",
    }),
  );

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value ?? null, expected);
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value ?? null, expected);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CellStructuralConflictMonitor surfaces format conflicts when two users move the same cell to the same destination with different formats", async () => {
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

  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  disconnect();

  cutPasteWithFormat(sessionA, "Sheet1:0:0", "Sheet1:0:1", { a: 1 });
  cutPasteWithFormat(sessionB, "Sheet1:0:0", "Sheet1:0:1", { a: 2 });

  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.equal(conflict.type, "cell");
  assert.equal(conflict.reason, "format");
  assert.equal(conflict.cellKey, "Sheet1:0:1");

  const expectedValue = conflict.local?.after?.value ?? null;
  const expectedFormat = conflict.local?.after?.format ?? null;

  assert.ok(conflictSide.cellConflictMonitor?.resolveConflict(conflict.id, { choice: "ours" }));

  assert.equal(await sessionA.getCell("Sheet1:0:0"), null);
  assert.equal(await sessionB.getCell("Sheet1:0:0"), null);
  assert.equal((await sessionA.getCell("Sheet1:0:1"))?.value ?? null, expectedValue);
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value ?? null, expectedValue);

  const yCellA = sessionA.cells.get("Sheet1:0:1");
  const yCellB = sessionB.cells.get("Sheet1:0:1");
  assert.ok(getYMap(yCellA));
  assert.ok(getYMap(yCellB));
  assert.deepEqual(yCellA.get("format") ?? null, expectedFormat);
  assert.deepEqual(yCellB.get("format") ?? null, expectedFormat);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
 
