import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";
import { FormulaConflictMonitor } from "../src/formula-conflict-monitor.js";

test("FormulaConflictMonitor.setLocalValue clears formulas via a null marker even in formula-only mode", () => {
  const doc = new Y.Doc();
  const monitor = new FormulaConflictMonitor({
    doc,
    localUserId: "user1",
    onConflict: () => {},
  });

  monitor.setLocalValue("Sheet1:0:0", "x");

  const cell = /** @type {any} */ (doc.getMap("cells").get("Sheet1:0:0"));
  assert.equal(cell.get("value"), "x");
  assert.equal(cell.get("formula"), null);

  monitor.dispose();
  doc.destroy();
});

test("CellStructuralConflictMonitor writes value cells with a formula=null marker", () => {
  const doc = new Y.Doc();
  const monitor = new CellStructuralConflictMonitor({
    doc,
    localUserId: "user1",
    onConflict: () => {},
  });

  // Use the internal writer helper directly to validate the Yjs-level encoding
  // used for structural conflict resolution.
  monitor._writeCell("Sheet1:0:0", { value: 123 });

  const cell = /** @type {any} */ (doc.getMap("cells").get("Sheet1:0:0"));
  assert.equal(cell.get("value"), 123);
  assert.equal(cell.get("formula"), null);

  monitor.dispose();
  doc.destroy();
});

