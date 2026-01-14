import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { FormulaConflictMonitor } from "../src/formula-conflict-monitor.js";

const REMOTE_ORIGIN = { type: "remote" };

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

  // Initial sync.
  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

test("FormulaConflictMonitor ignores invalid cell keys instead of throwing", () => {
  // Ensure deterministic map-entry overwrite tie-breaking: higher clientID wins.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  const localOrigin = { type: "local-a" };

  /** @type {Array<any>} */
  const conflicts = [];

  const monitor = new FormulaConflictMonitor({
    doc: docA,
    localUserId: "user-a",
    origin: localOrigin,
    localOrigins: new Set([localOrigin]),
    onConflict: (c) => conflicts.push(c),
  });

  let disconnect = connectDocs(docA, docB);

  const cellKey = "bad-key";

  // Ensure both docs have a shared base cell map so concurrent edits race on the
  // formula key (not on the `cells[cellKey] = new Y.Map()` insertion).
  docA.transact(() => {
    docA.getMap("cells").set(cellKey, new Y.Map());
  }, localOrigin);
  assert.ok(docB.getMap("cells").get(cellKey), "expected base cell map to sync to docB");

  // Simulate offline concurrent formula edits.
  disconnect();

  monitor.setLocalFormula(cellKey, "=1");

  docB.transact(() => {
    const cells = docB.getMap("cells");
    let cell = /** @type {Y.Map<any>} */ (cells.get(cellKey));
    if (!cell) {
      cell = new Y.Map();
      cells.set(cellKey, cell);
    }
    cell.set("formula", "=2");
    cell.set("value", null);
    cell.set("modified", Date.now());
  });

  // Reconnect and sync state. Previously, invalid keys could throw when the
  // conflict monitor attempted to parse the cell key for the emitted conflict.
  assert.doesNotThrow(() => {
    disconnect = connectDocs(docA, docB);
  });

  // Invalid keys are ignored for conflict emission.
  assert.equal(conflicts.length, 0);

  monitor.dispose();
  disconnect();
  docA.destroy();
  docB.destroy();
});

