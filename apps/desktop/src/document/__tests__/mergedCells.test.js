import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";
import { mergeAcross, mergeCells, mergeCenter, unmergeCells } from "../mergedCells.js";

test("mergeAcross creates a merged region per row", () => {
  const doc = new DocumentController();

  // A1:C3 => A1:C1, A2:C2, A3:C3
  doc.beginBatch({ label: "Merge Across" });
  mergeAcross(doc, "Sheet1", { startRow: 0, endRow: 2, startCol: 0, endCol: 2 }, { label: "Merge Across" });
  doc.endBatch();

  assert.deepEqual(doc.getMergedRanges("Sheet1"), [
    { startRow: 0, endRow: 0, startCol: 0, endCol: 2 },
    { startRow: 1, endRow: 1, startCol: 0, endCol: 2 },
    { startRow: 2, endRow: 2, startCol: 0, endCol: 2 },
  ]);
});

test("unmerge removes merges and preserves anchor value", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "keep");
  doc.setCellValue("Sheet1", "B1", "discard");

  doc.beginBatch({ label: "Merge Cells" });
  mergeCells(doc, "Sheet1", { startRow: 0, endRow: 0, startCol: 0, endCol: 1 }, { label: "Merge Cells" });
  doc.endBatch();

  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]);
  assert.equal(doc.getCell("Sheet1", "A1").value, "keep");
  assert.equal(doc.getCell("Sheet1", "B1").value, null);

  unmergeCells(doc, "Sheet1", { startRow: 0, endRow: 0, startCol: 1, endCol: 1 }, { label: "Unmerge Cells" });

  assert.deepEqual(doc.getMergedRanges("Sheet1"), []);
  assert.equal(doc.getCell("Sheet1", "A1").value, "keep");
});

test("mergeCenter sets horizontal center alignment on the merged cell", () => {
  const doc = new DocumentController();

  doc.beginBatch({ label: "Merge & Center" });
  mergeCenter(doc, "Sheet1", { startRow: 0, endRow: 0, startCol: 0, endCol: 2 }, { label: "Merge & Center" });
  doc.endBatch();

  const format = doc.getCellFormat("Sheet1", "A1");
  assert.equal(format?.alignment?.horizontal, "center");
});

test("merged ranges shift/expand with structural row edits", () => {
  const doc = new DocumentController();

  doc.setMergedRanges("Sheet1", [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }], { label: "Merge Cells" }); // A1:B1

  // Insert a row above: merge should shift down.
  doc.insertRows("Sheet1", 0, 1, { label: "Insert Rows" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 1, endRow: 1, startCol: 0, endCol: 1 }]);

  // Insert a row inside the merge (between the merged row and below): merge should expand.
  doc.setMergedRanges("Sheet1", [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }], { label: "Merge Cells" }); // A1:B2
  doc.insertRows("Sheet1", 1, 1, { label: "Insert Rows" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 2, startCol: 0, endCol: 1 }]);

  // Delete a row inside the merge: merge should shrink.
  doc.deleteRows("Sheet1", 1, 1, { label: "Delete Rows" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
});

test("merged ranges shift/expand with structural column edits", () => {
  const doc = new DocumentController();

  doc.setMergedRanges("Sheet1", [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }], { label: "Merge Cells" }); // A1:B1

  // Insert a col to the left: merge should shift right.
  doc.insertCols("Sheet1", 0, 1, { label: "Insert Columns" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 0, startCol: 1, endCol: 2 }]);

  // Insert a col inside the merge: merge should expand.
  doc.setMergedRanges("Sheet1", [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }], { label: "Merge Cells" }); // A1:B1
  doc.insertCols("Sheet1", 1, 1, { label: "Insert Columns" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 0, startCol: 0, endCol: 2 }]);

  // Delete a col inside the merge: merge should shrink.
  doc.deleteCols("Sheet1", 1, 1, { label: "Delete Columns" });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]);
});
