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
  let lastChange = null;
  restored.on("change", (payload) => {
    lastChange = payload;
  });
  restored.applyState(snapshot);

  // applyState clears history and marks dirty until the host explicitly marks saved.
  assert.equal(restored.canUndo, false);
  assert.equal(restored.canRedo, false);
  assert.equal(restored.isDirty, true);

  assert.equal(lastChange?.source, "applyState");
  assert.equal(restored.getCell("Sheet1", "A1").value, 1);
  const a1 = restored.getCell("Sheet1", "A1");
  assert.deepEqual(restored.styleTable.get(a1.styleId), { bold: true });
  assert.equal(restored.getCell("Sheet1", "B1").formula, "=SUM(A1:A3)");
});

test("applyState materializes empty sheets from snapshots", () => {
  const doc = new DocumentController();
  // DocumentController lazily creates sheets on first access. This creates an empty sheet
  // that should still survive encode/apply roundtrips.
  doc.getCell("EmptySheet", "A1");

  const snapshot = doc.encodeState();
  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.deepEqual(restored.getSheetIds(), ["EmptySheet"]);
});

test("applyState removes sheets that are not present in the snapshot", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.getCell("ExtraSheet", "A1"); // empty sheet

  const next = new DocumentController();
  next.setCellValue("OnlySheet", "A1", 2);

  doc.applyState(next.encodeState());

  assert.deepEqual(doc.getSheetIds(), ["OnlySheet"]);
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
