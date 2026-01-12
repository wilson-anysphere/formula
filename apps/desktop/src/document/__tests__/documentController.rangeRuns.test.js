import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("setRangeFormat uses compressed range runs for huge rectangles without materializing cells", () => {
  const doc = new DocumentController();

  // 26 columns * 1,000,000 rows = 26,000,000 cells. This must not enumerate per cell.
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  // No cells should be created for format-only range runs.
  assert.equal(sheet.cells.size, 0);

  // Range runs should be stored per-column.
  assert.equal(sheet.formatRunsByCol.size, 26);

  // Effective formatting should apply to empty cells inside the rectangle.
  const inside = doc.getCellFormat("Sheet1", "A1");
  assert.equal(inside.font?.bold, true);

  // Cells outside the rectangle should not have the format.
  const outside = doc.getCellFormat("Sheet1", "AA1"); // column 27 (0-based 26)
  assert.equal(outside.font?.bold, undefined);

  // Styles should be interned per segment, not per cell.
  assert.equal(doc.styleTable.size, 2); // default + bold

  // includeFormat used range should incorporate range-run formatting (without cell materialization).
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 999_999,
    startCol: 0,
    endCol: 25,
  });
});

test("range-run formatting is undoable + redoable", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
  doc.undo();
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, undefined);
  doc.redo();
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
});

test("encodeState/applyState roundtrip preserves range-run formatting", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  const snapshot = doc.encodeState();
  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.equal(restored.getCellFormat("Sheet1", "A1").font?.bold, true);
  const sheet = restored.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.formatRunsByCol.size, 26);
});
