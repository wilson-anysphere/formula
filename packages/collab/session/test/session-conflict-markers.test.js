import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

test("CollabSession preserves formula null markers for clears and value writes", async (t) => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc });

  t.after(() => {
    session.destroy();
    doc.destroy();
  });

  const cells = doc.getMap("cells");

  // 1) Formula write.
  await session.setCellFormula("Sheet1:0:0", "=1");
  const yFormulaCell = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(yFormulaCell, "expected cell map to exist after setting a formula");
  assert.equal(yFormulaCell.get("formula"), "=1");

  // 2) Formula clear should keep the cell map and write an explicit `formula=null` marker.
  await session.setCellFormula("Sheet1:0:0", null);
  const yFormulaCleared = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(yFormulaCleared, "expected cell map to still exist after clearing a formula");
  assert.equal(yFormulaCleared.has("formula"), true);
  assert.equal(yFormulaCleared.get("formula"), null);

  // 3) Literal value writes/clears should also preserve a `formula=null` marker.
  await session.setCellValue("Sheet1:0:1", 123);
  const yValueCell = /** @type {any} */ (cells.get("Sheet1:0:1"));
  assert.ok(yValueCell, "expected cell map to exist after setting a value");
  assert.equal(yValueCell.get("value"), 123);
  assert.equal(yValueCell.has("formula"), true);
  assert.equal(yValueCell.get("formula"), null);

  await session.setCellValue("Sheet1:0:1", null);
  const yValueCleared = /** @type {any} */ (cells.get("Sheet1:0:1"));
  assert.ok(yValueCleared, "expected cell map to still exist after clearing a value");
  assert.equal(yValueCleared.has("formula"), true);
  assert.equal(yValueCleared.get("formula"), null);
});

