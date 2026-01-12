import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor clears plaintext cells via value/formula null markers (no root delete)", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const cellKey = "Sheet1:0:0";

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "local",
    onConflict: () => {},
  });

  // Seed with a non-empty plaintext cell so clearing has an effect.
  doc.transact(() => {
    const cellMap = new Y.Map();
    cellMap.set("value", 123);
    cellMap.set("formula", null);
    cells.set(cellKey, cellMap);
  });

  const conflictId = "conflict-clear-1";
  monitor._conflicts.set(conflictId, {
    id: conflictId,
    type: "cell",
    reason: "content",
    sheetId: "Sheet1",
    cell: "A1",
    cellKey,
    // Local side keeps the value; remote side clears the cell.
    local: { after: { value: 123 } },
    remote: { after: null },
    remoteUserId: "",
    detectedAt: Date.now(),
  });

  // Choose the remote resolution (clear).
  assert.equal(monitor.resolveConflict(conflictId, { choice: "theirs" }), true);

  // The cell entry should still exist, but be marker-only empty.
  const written = cells.get(cellKey);
  assert.ok(written instanceof Y.Map);
  assert.equal(written.has("value"), true);
  assert.equal(written.get("value"), null);
  assert.equal(written.has("formula"), true);
  assert.equal(written.get("formula"), null);

  monitor.dispose();
});

