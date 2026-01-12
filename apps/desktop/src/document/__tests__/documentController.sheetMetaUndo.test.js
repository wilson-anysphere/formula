import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("renameSheet -> undo -> redo restores sheet name", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  doc.renameSheet("Sheet1", "Budget");
  assert.equal(doc.getSheetMeta("Sheet1")?.name, "Budget");

  doc.undo();
  assert.equal(doc.getSheetMeta("Sheet1")?.name, "Sheet1");

  doc.redo();
  assert.equal(doc.getSheetMeta("Sheet1")?.name, "Budget");
});

test("reorderSheets -> undo -> redo restores sheet order", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);
  doc.setCellValue("Sheet3", "A1", 3);

  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2", "Sheet3"]);

  doc.reorderSheets(["Sheet3", "Sheet1", "Sheet2"]);
  assert.deepEqual(doc.getSheetIds(), ["Sheet3", "Sheet1", "Sheet2"]);

  doc.undo();
  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2", "Sheet3"]);

  doc.redo();
  assert.deepEqual(doc.getSheetIds(), ["Sheet3", "Sheet1", "Sheet2"]);
});

test("hideSheet/unhide undo/redo", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  doc.hideSheet("Sheet1");
  assert.equal(doc.getSheetMeta("Sheet1")?.visibility, "hidden");
  assert.deepEqual(doc.getVisibleSheetIds(), []);

  doc.undo();
  assert.equal(doc.getSheetMeta("Sheet1")?.visibility, "visible");
  assert.deepEqual(doc.getVisibleSheetIds(), ["Sheet1"]);

  doc.redo();
  assert.equal(doc.getSheetMeta("Sheet1")?.visibility, "hidden");
});

test("deleteSheet undo restores the sheet and its cell contents", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "keep");
  doc.setCellValue("Sheet2", "A1", "hello");

  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2"]);

  doc.deleteSheet("Sheet2");
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);

  doc.undo();
  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2"]);
  assert.equal(doc.getCell("Sheet2", "A1").value, "hello");
});

