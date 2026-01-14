import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor ignores invalid cell keys instead of throwing", () => {
  const doc = new Y.Doc();

  const monitor = new CellStructuralConflictMonitor({
    doc,
    localUserId: "user-a",
    onConflict: () => {},
  });

  assert.doesNotThrow(() => {
    // Intentionally call the internal helper with an invalid key; conflict monitors
    // should ignore malformed keys rather than crashing observers.
    monitor._emitConflict({
      type: "cell",
      reason: "content",
      sourceCellKey: "bad-key",
      local: {},
      remote: {},
    });
    monitor._emitConflict({
      type: "cell",
      reason: "content",
      sourceCellKey: "Sheet1:x:y",
      local: {},
      remote: {},
    });
  });

  assert.equal(monitor.listConflicts().length, 0);

  monitor.dispose();
  doc.destroy();
});
