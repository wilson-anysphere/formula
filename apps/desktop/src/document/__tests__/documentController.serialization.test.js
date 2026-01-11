import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("encodeState/applyState roundtrip restores cell inputs and clears history", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellFormula("Sheet1", "B1", "SUM(A1:A3)");
  doc.setRangeFormat("Sheet1", "A1", { bold: true });
  assert.equal(doc.canUndo, true);

  const snapshot = doc.encodeState();
  assert.ok(snapshot instanceof Uint8Array);

  const restored = new DocumentController();
  restored.applyState(snapshot);

  // applyState clears history and marks dirty until the host explicitly marks saved.
  assert.equal(restored.canUndo, false);
  assert.equal(restored.canRedo, false);
  assert.equal(restored.isDirty, true);

  assert.equal(restored.getCell("Sheet1", "A1").value, 1);
  const a1 = restored.getCell("Sheet1", "A1");
  assert.deepEqual(restored.styleTable.get(a1.styleId), { bold: true });
  assert.equal(restored.getCell("Sheet1", "B1").formula, "=SUM(A1:A3)");
});

test("update event fires on edits and undo/redo", () => {
  const doc = new DocumentController();
  let updates = 0;
  doc.on("update", () => {
    updates += 1;
  });

  doc.setCellValue("Sheet1", "A1", "x");
  doc.undo();
  doc.redo();

  assert.equal(updates, 3);
});
