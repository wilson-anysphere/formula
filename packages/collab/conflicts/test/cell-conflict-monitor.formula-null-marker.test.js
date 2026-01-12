import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { CellConflictMonitor } from "../src/cell-conflict-monitor.js";

test("CellConflictMonitor.setLocalValue clears formulas via a null marker", () => {
  const doc = new Y.Doc();
  const monitor = new CellConflictMonitor({
    doc,
    localUserId: "user1",
    onConflict: () => {}
  });

  monitor.setLocalValue("Sheet1:0:0", "x");

  const cells = doc.getMap("cells");
  const cell = /** @type {any} */ (cells.get("Sheet1:0:0"));

  assert.equal(cell.get("value"), "x");
  assert.equal(cell.get("formula"), null);
});

