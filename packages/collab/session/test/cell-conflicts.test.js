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

test("CollabSession cell value conflict monitor detects long-offline concurrent value edits and converges after resolution", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellValueConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });

  // Establish a shared base cell map so concurrent edits race on the value key
  // (not on the `cells[cellKey] = new Y.Map()` insertion).
  await sessionA.setCellValue("Sheet1:0:0", 0);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, 0);

  // Simulate offline concurrent edits (same cell, different values).
  disconnect();
  await sessionA.setCellValue("Sheet1:0:0", "a");
  await new Promise((r) => setTimeout(r, 200));
  await sessionB.setCellValue("Sheet1:0:0", "b");

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  const allConflicts = [...conflictsA, ...conflictsB];
  assert.ok(allConflicts.length >= 1, "expected at least one conflict to be detected");

  const conflictSide = conflictsA.length > 0 ? sessionA : sessionB;
  const conflict = conflictsA.length > 0 ? conflictsA[0] : conflictsB[0];

  assert.ok(conflictSide.cellValueConflictMonitor?.resolveConflict(conflict.id, conflict.localValue));

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, conflict.localValue);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, conflict.localValue);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession cell value conflict monitor still detects conflicts after recreating the session (restart)", async () => {
  // For concurrent overwrites, Yjs deterministically breaks ties using clientID (higher wins).
  // Ensure the remote value writer wins the value key.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;
  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];

  let sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellValueConflicts: {
      localUserId: "user-b",
      onConflict: () => {},
    },
  });

  // Establish base.
  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  // Offline concurrent edits.
  disconnect();
  await sessionA.setCellValue("Sheet1:0:0", "ours");

  // Simulate app reload: recreate the session (and its conflict monitor) before reconnecting.
  sessionA.destroy();
  conflictsA.length = 0;
  sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });

  await sessionB.setCellValue("Sheet1:0:0", "theirs");

  // Reconnect.
  disconnect = connectDocs(docA, docB);

  assert.ok(conflictsA.length >= 1, "expected at least one conflict after restart");
  const conflict = conflictsA[0];
  assert.equal(conflict.localValue, "ours");
  assert.equal(conflict.remoteValue, "theirs");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession cell value conflict monitor does not flag sequential value edits", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];
  /** @type {Array<any>} */
  const conflictsB = [];

  const sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({
    doc: docB,
    cellValueConflicts: {
      localUserId: "user-b",
      onConflict: (c) => conflictsB.push(c),
    },
  });

  await sessionA.setCellValue("Sheet1:0:0", "x");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "x");

  await sessionB.setCellValue("Sheet1:0:0", "y");
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, "y");

  assert.equal(conflictsA.length, 0);
  assert.equal(conflictsB.length, 0);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession cell value conflict monitor ignores sequential deletes that remove the value key", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];

  const sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });

  // Remote doc (legacy) without a monitor.
  const sessionB = createCollabSession({ doc: docB });

  await sessionA.setCellValue("Sheet1:0:0", "x");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "x");

  const cellMap = sessionB.cells.get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(typeof cellMap.get, "function");

  // Simulate a legacy client that clears cells by deleting the `value` key.
  docB.transact(() => {
    cellMap.delete("value");
    cellMap.set("modifiedBy", "user-b");
    cellMap.set("modified", Date.now());
  });

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, null);
  assert.equal(conflictsA.length, 0);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession resolving a value conflict by choosing remote preserves concurrent formulas", async () => {
  // For concurrent map-entry overwrites, Yjs deterministically breaks ties using
  // clientID (higher wins). Ensure the remote formula writer wins the value key.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];

  const sessionA = createCollabSession({
    doc: docA,
    cellValueConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });

  const sessionB = createCollabSession({ doc: docB });

  // Establish base.
  await sessionA.setCellValue("Sheet1:0:0", "base");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "base");

  // Offline concurrent edits: A writes a value, B writes a formula (which sets value=null).
  disconnect();
  await sessionA.setCellValue("Sheet1:0:0", "ours");
  await sessionB.setCellFormula("Sheet1:0:0", "=1");

  // Reconnect.
  disconnect = connectDocs(docA, docB);

  assert.ok(conflictsA.length >= 1, "expected at least one conflict to be detected");
  const conflict = conflictsA[0];

  // The concurrent formula should be present in the doc at conflict time.
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, "=1");

  // Choosing the remote value is a no-op (remote state is already applied) and must not
  // clear the formula via setLocalValue().
  assert.ok(sessionA.cellValueConflictMonitor?.resolveConflict(conflict.id, conflict.remoteValue));

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.formula, "=1");
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.formula, "=1");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
