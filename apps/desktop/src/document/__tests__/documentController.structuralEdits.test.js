import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("insertRows shifts sparse cell map entries and preserves styleId (undoable)", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A2", "moved");
  doc.setRangeFormat("Sheet1", "A2", { font: { bold: true } });
  const before = doc.getCell("Sheet1", "A2");
  assert.equal(before.value, "moved");
  assert.ok(before.styleId !== 0, "expected styleId to be non-zero after formatting");

  doc.insertRows("Sheet1", 1, 1, { label: "Insert Rows" });

  assert.equal(doc.getCell("Sheet1", "A2").value, null);
  const moved = doc.getCell("Sheet1", "A3");
  assert.equal(moved.value, "moved");
  assert.equal(moved.styleId, before.styleId);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A2").value, "moved");
  assert.equal(doc.getCell("Sheet1", "A3").value, null);
});

test("structural row/col edits shift layered formatting, range runs, and sheet view overrides", () => {
  const doc = new DocumentController();

  // Row formatting + row height override at row 20.
  doc.setRowFormat("Sheet1", 20, { font: { italic: true } });
  const rowStyleBefore = doc.getRowStyleId("Sheet1", 20);
  assert.ok(rowStyleBefore !== 0);
  doc.setRowHeight("Sheet1", 20, 40);

  // Create a large range-run format in column A so we exercise run-shifting.
  doc.setRangeFormat("Sheet1", "A1:A50001", { fill: { color: "#FFFF0000" } });
  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.ok(sheet.formatRunsByCol.get(0)?.length > 0);

  // Insert 2 rows at row 10.
  doc.insertRows("Sheet1", 10, 2, { label: "Insert Rows" });

  assert.equal(doc.getRowStyleId("Sheet1", 20), 0);
  assert.equal(doc.getRowStyleId("Sheet1", 22), rowStyleBefore);

  const viewAfterRows = doc.getSheetView("Sheet1");
  assert.equal(viewAfterRows.rowHeights?.["22"], 40);
  assert.equal(viewAfterRows.rowHeights?.["20"], undefined);

  const runsAfterInsert = sheet.formatRunsByCol.get(0) ?? [];
  // Insertion splits the run and shifts the suffix down.
  assert.deepEqual(
    runsAfterInsert.slice(0, 2).map((r) => ({ startRow: r.startRow, endRowExclusive: r.endRowExclusive })),
    [
      { startRow: 0, endRowExclusive: 10 },
      { startRow: 12, endRowExclusive: 50003 },
    ],
  );

  // Column formatting + width override at col 5 (F).
  doc.setColFormat("Sheet1", 5, { font: { bold: true } });
  const colStyleBefore = doc.getColStyleId("Sheet1", 5);
  assert.ok(colStyleBefore !== 0);
  doc.setColWidth("Sheet1", 5, 123);

  // Create range-run formatting in column C (index 2).
  doc.setRangeFormat("Sheet1", "C1:C50001", { font: { underline: true } });
  assert.ok(sheet.formatRunsByCol.get(2)?.length > 0);

  // Insert a column at B (index 1): C->D, F->G.
  doc.insertCols("Sheet1", 1, 1, { label: "Insert Columns" });

  assert.equal(doc.getColStyleId("Sheet1", 5), 0);
  assert.equal(doc.getColStyleId("Sheet1", 6), colStyleBefore);

  const viewAfterCols = doc.getSheetView("Sheet1");
  assert.equal(viewAfterCols.colWidths?.["6"], 123);
  assert.equal(viewAfterCols.colWidths?.["5"], undefined);

  assert.equal(sheet.formatRunsByCol.has(2), false);
  assert.ok(sheet.formatRunsByCol.get(3)?.length > 0);
});

test("structural edits rewrite formulas (best-effort A1 semantics)", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 123);
  doc.setCellFormula("Sheet1", "B2", "=A1");

  // Insert a row at the top. A1 -> A2, B2 -> B3; formula should follow A1's new address.
  doc.insertRows("Sheet1", 0, 1, { label: "Insert Rows" });

  assert.equal(doc.getCell("Sheet1", "B3").formula, "=A2");

  // Delete the referenced row (row 1 / A2). The formula should become #REF!.
  doc.deleteRows("Sheet1", 1, 1, { label: "Delete Rows" });
  assert.equal(doc.getCell("Sheet1", "B2").formula, "=#REF!");
});

test("structural edits apply engine-provided formulaRewrites when present", () => {
  const doc = new DocumentController();

  // B1 will move to B2 when we insert row 0.
  doc.setCellFormula("Sheet1", "B1", "=A1");

  // Formula on a different sheet should also rewrite when provided by the engine.
  doc.setCellFormula("Sheet2", "A1", "='Sheet1'!A1");

  doc.insertRows("Sheet1", 0, 1, {
    label: "Insert Rows",
    formulaRewrites: [
      { sheet: "Sheet1", address: "B2", before: "=A1", after: "=Z99" },
      { sheet: "Sheet2", address: "A1", before: "='Sheet1'!A1", after: "='Sheet1'!A2" },
    ],
  });

  assert.equal(doc.getCell("Sheet1", "B2").formula, "=Z99");
  assert.equal(doc.getCell("Sheet2", "A1").formula, "='Sheet1'!A2");
});
