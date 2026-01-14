import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("insertCellsShiftRight shifts stored cells and range-run formatting within the selected rows", () => {
  const doc = new DocumentController();

  // Create range-run formatting in column A (large rectangle).
  doc.setRangeFormat("Sheet1", "A1:A50001", { font: { bold: true } });
  doc.setCellValue("Sheet1", "A1", "moved");
  doc.setCellValue("Sheet1", "A2", "stay");

  doc.insertCellsShiftRight("Sheet1", "A1:B1", { label: "Insert Cells" });

  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "C1").value, "moved");
  assert.equal(doc.getCell("Sheet1", "A2").value, "stay");

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  const runsCol0 = sheet.formatRunsByCol.get(0) ?? [];
  const runsCol2 = sheet.formatRunsByCol.get(2) ?? [];

  assert.deepEqual(
    runsCol0.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 1, endRowExclusive: 50001 }],
  );
  assert.deepEqual(
    runsCol2.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 0, endRowExclusive: 1 }],
  );
});

test("deleteCellsShiftLeft deletes cells and shifts remaining cells + range-run formatting", () => {
  const doc = new DocumentController();

  // Formatting in column C (index 2).
  doc.setRangeFormat("Sheet1", "C1:C50001", { font: { bold: true } });
  doc.setCellValue("Sheet1", "C1", "moved");

  doc.deleteCellsShiftLeft("Sheet1", "A1:B1", { label: "Delete Cells" });

  assert.equal(doc.getCell("Sheet1", "C1").value, null);
  assert.equal(doc.getCell("Sheet1", "A1").value, "moved");

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  const runsCol0 = sheet.formatRunsByCol.get(0) ?? [];
  const runsCol2 = sheet.formatRunsByCol.get(2) ?? [];

  assert.deepEqual(
    runsCol0.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 0, endRowExclusive: 1 }],
  );
  assert.deepEqual(
    runsCol2.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 1, endRowExclusive: 50001 }],
  );
});

test("insertCellsShiftDown shifts range-run formatting down within selected columns", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A50001", { font: { bold: true } });

  doc.insertCellsShiftDown("Sheet1", "A1:A2", { label: "Insert Cells" });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  const runs = sheet.formatRunsByCol.get(0) ?? [];
  assert.deepEqual(
    runs.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 2, endRowExclusive: 50003 }],
  );
});

test("deleteCellsShiftUp shifts range-run formatting up within selected columns", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A50001", { font: { bold: true } });

  doc.deleteCellsShiftUp("Sheet1", "A1:A2", { label: "Delete Cells" });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  const runs = sheet.formatRunsByCol.get(0) ?? [];
  assert.deepEqual(
    runs.map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [{ startRow: 0, endRowExclusive: 49999 }],
  );
});
