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

