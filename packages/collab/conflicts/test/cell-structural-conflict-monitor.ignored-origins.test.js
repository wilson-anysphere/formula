import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor ignores bulk time-travel origins (no op log growth)", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const ops = doc.getMap("cellStructuralOps");
  assert.equal(ops.size, 0);

  const localOrigin = { type: "local" };

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "user-a",
    origin: localOrigin,
    // Simulate a misconfiguration where bulk origins are accidentally treated as local.
    localOrigins: new Set([localOrigin, "versioning-restore", "branching-apply"]),
    ignoredOrigins: new Set(["versioning-restore", "branching-apply"]),
    onConflict: () => {},
  });

  assert.ok(monitor.localOrigins.has("versioning-restore"));
  assert.ok(monitor.localOrigins.has("branching-apply"));
  assert.ok(monitor.ignoredOrigins.has("versioning-restore"));
  assert.ok(monitor.ignoredOrigins.has("branching-apply"));

  // Even though the transaction origins are considered "local", they are ignored entirely,
  // so they must not be logged into the shared `cellStructuralOps` log.
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "bulk-restore");
    cell.set("modifiedBy", "restorer");
    cell.set("modified", Date.now());
    cells.set("Sheet1:0:0", cell);
  }, "versioning-restore");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "bulk-branch-apply");
    cell.set("modifiedBy", "brancher");
    cell.set("modified", Date.now());
    cells.set("Sheet1:0:1", cell);
  }, "branching-apply");

  assert.equal(ops.size, 0);

  monitor.dispose();
});

