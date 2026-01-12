import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { FormulaConflictMonitor } from "../src/formula-conflict-monitor.js";

test("FormulaConflictMonitor.setLocalValue writes a formula=null marker in formula-only mode", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const monitor = new FormulaConflictMonitor({
    doc,
    cells,
    localUserId: "local",
    onConflict: () => {
      // Not expected in this test.
    }
  });

  monitor.setLocalValue("Sheet1:0:0", "x");

  const cell = cells.get("Sheet1:0:0");
  assert.ok(cell);
  assert.equal(cell.get("value"), "x");
  assert.equal(cell.get("formula"), null);

  monitor.dispose();
});

