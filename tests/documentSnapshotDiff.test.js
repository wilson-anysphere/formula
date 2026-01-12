import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { cellKey } from "../packages/versioning/src/diff/semanticDiff.js";
import { diffDocumentSnapshots } from "../packages/versioning/src/document/diffSnapshots.js";
import { sheetStateFromDocumentSnapshot } from "../packages/versioning/src/document/sheetState.js";

test("sheetStateFromDocumentSnapshot extracts a sheet into SheetState", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellFormula("Sheet1", "B1", "SUM(A1:A3)");
  doc.setRangeFormat("Sheet1", "A1", { bold: true });

  const snapshot = doc.encodeState();
  const state = sheetStateFromDocumentSnapshot(snapshot, { sheetId: "Sheet1" });
  assert.equal(state.cells.size, 2);
  assert.deepEqual(state.cells.get(cellKey(0, 0)), { value: 1, formula: null, format: { bold: true } });
  assert.deepEqual(state.cells.get(cellKey(0, 1)), { value: null, formula: "=SUM(A1:A3)", format: null });
});

test("diffDocumentSnapshots computes semantic diffs between two snapshots", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  const beforeSnapshot = doc.encodeState();

  doc.setCellValue("Sheet1", "A1", 2);
  doc.setCellValue("Sheet1", "B1", "new");
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentSnapshots({ beforeSnapshot, afterSnapshot, sheetId: "Sheet1" });
  assert.equal(diff.modified.length, 1);
  assert.equal(diff.added.length, 1);
  assert.equal(diff.removed.length, 0);
});

test("diffDocumentSnapshots detects formatOnly edits from column-default formatting", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  const beforeSnapshot = doc.encodeState();

  // Task 44 layered formatting: formatting a full column should update the column-default
  // formatting layer, without enumerating the full column.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentSnapshots({ beforeSnapshot, afterSnapshot, sheetId: "Sheet1" });
  assert.equal(diff.formatOnly.length, 1);
  assert.deepEqual(diff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
});

test("diffDocumentSnapshots detects formatOnly edits from sheet-default formatting", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  const beforeSnapshot = doc.encodeState();

  // Task 44 layered formatting: formatting the entire sheet should update the sheet-default
  // formatting layer, without expanding the full sheet into per-cell styles.
  doc.setRangeFormat("Sheet1", "A1:XFD1048576", { font: { bold: true } });
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentSnapshots({ beforeSnapshot, afterSnapshot, sheetId: "Sheet1" });
  assert.equal(diff.formatOnly.length, 1);
  assert.deepEqual(diff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
});

test("exportSheetForSemanticDiff exports effective formats from column defaults", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const exported = doc.exportSheetForSemanticDiff("Sheet1");
  assert.deepEqual(exported.cells.get(cellKey(0, 0))?.format, { font: { bold: true } });
});

test("exportSheetForSemanticDiff exports effective formats from sheet defaults", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setRangeFormat("Sheet1", "A1:XFD1048576", { font: { bold: true } });

  const exported = doc.exportSheetForSemanticDiff("Sheet1");
  assert.deepEqual(exported.cells.get(cellKey(0, 0))?.format, { font: { bold: true } });
});
