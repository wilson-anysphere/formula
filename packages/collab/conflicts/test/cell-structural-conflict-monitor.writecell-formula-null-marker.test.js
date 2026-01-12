import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor writes value cells with formula=null marker (not key deletion)", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const cellKey = "Sheet1:0:0";

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "local",
    onConflict: () => {}
  });

  // Seed with a formula cell so we can ensure conflict resolution "clears" the
  // formula using an explicit marker rather than deleting the key.
  doc.transact(() => {
    const cellMap = new Y.Map();
    cellMap.set("formula", "=1+1");
    cellMap.set("value", null);
    cells.set(cellKey, cellMap);
  });

  const conflictId = "conflict-1";
  monitor._conflicts.set(conflictId, {
    id: conflictId,
    type: "cell",
    reason: "content",
    sheetId: "Sheet1",
    cell: "A1",
    cellKey,
    local: { after: { value: 123 } },
    remote: { after: { formula: "=9+9" } },
    remoteUserId: "",
    detectedAt: Date.now()
  });

  assert.equal(monitor.resolveConflict(conflictId, { choice: "ours" }), true);

  const written = cells.get(cellKey);
  assert.ok(written instanceof Y.Map);
  assert.equal(written.get("value"), 123);
  assert.equal(written.get("formula"), null);

  monitor.dispose();
});

