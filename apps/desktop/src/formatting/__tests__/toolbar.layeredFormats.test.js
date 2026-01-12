import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toggleBold reads full-column formatting from the column layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Apply bold to the entire column A via the layered column format path.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "Sheet1 should exist after formatting");
  assert.equal(sheet.cells.size, 0, "Full-column formatting should not materialize per-cell overrides");

  // Toggling again should flip bold OFF (because all cells are currently bold).
  toggleBold(doc, "Sheet1", "A1:A1048576");
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(sheet.cells.size, 0, "Toggling full-column formatting should not materialize per-cell overrides");
});

test("toggleBold on a full-column selection treats a single conflicting cell override as mixed state", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  // Introduce a single conflicting cell-level override inside the column.
  doc.setRangeFormat("Sheet1", "A2", { font: { bold: false } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "Sheet1 should exist after formatting");
  assert.equal(sheet.cells.size, 1, "A single cell override should not expand to the full column");

  // Selection is now mixed, so toggleBold should choose next=true and make everything bold.
  toggleBold(doc, "Sheet1", "A1:A1048576");

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), true);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A2").font?.bold), true);
  assert.ok(sheet.cells.size <= 1, "Toggling should not materialize per-cell overrides across the selection");
});

test("toggleBold toggles OFF for large rectangles formatted via range runs (no per-cell scan)", () => {
  const doc = new DocumentController();

  const hugeRect = "A1:C100000"; // 300k cells -> range-run layer
  doc.setRangeFormat("Sheet1", hugeRect, { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Range-run formatting should not materialize cells");
  assert.equal(sheet.formatRunsByCol.size, 3, "Expected per-column range runs for A:C");

  // Guardrail: ensure the toggle read-path does not enumerate every cell.
  const originalGetCellFormat = doc.getCellFormat.bind(doc);
  let getCellFormatCalls = 0;
  doc.getCellFormat = (...args) => {
    getCellFormatCalls += 1;
    if (getCellFormatCalls > 10_000) {
      throw new Error(`toggleBold performed O(area) getCellFormat calls (${getCellFormatCalls})`);
    }
    return originalGetCellFormat(...args);
  };

  toggleBold(doc, "Sheet1", hugeRect);
  doc.getCellFormat = originalGetCellFormat;

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "C99999").font?.bold), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "D1").font?.bold), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleBold considers range-run overrides when computing full-column mixed state", () => {
  const doc = new DocumentController();

  // Ensure the sheet exists.
  doc.getCell("Sheet1", "A1");
  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  const boldTrueStyleId = doc.styleTable.intern({ font: { bold: true } });
  const boldFalseStyleId = doc.styleTable.intern({ font: { bold: false } });

  // Column is bold by default, but a large run overrides part of it to bold=false.
  sheet.colStyleIds.set(0, boldTrueStyleId);
  sheet.formatRunsByCol.set(0, [{ startRow: 0, endRowExclusive: 100_000, styleId: boldFalseStyleId }]);

  let lastPatch = null;
  doc.setRangeFormat = (_sheetId, _range, patch) => {
    lastPatch = patch;
  };

  toggleBold(doc, "Sheet1", "A1:A1048576");
  assert.deepEqual(lastPatch, { font: { bold: true } });
});

test("toggleBold reads full-row formatting from the row layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Entire row 1 via row layer.
  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Full-row formatting should not materialize per-cell overrides");

  // Second toggle should flip bold OFF.
  toggleBold(doc, "Sheet1", "A1:XFD1");
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleBold treats a single conflicting cell override in a full-row selection as mixed", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "B1", { font: { bold: false } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 1);

  toggleBold(doc, "Sheet1", "A1:XFD1");

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), true);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "B1").font?.bold), true);
  assert.ok(sheet.cells.size <= 1, "Should not materialize per-cell overrides across the full row");
});

test("toggleBold reads full-sheet formatting from the sheet layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Full sheet in Excel address space.
  const fullSheet = "A1:XFD1048576";
  doc.setRangeFormat("Sheet1", fullSheet, { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Sheet formatting should not materialize per-cell overrides");

  toggleBold(doc, "Sheet1", fullSheet);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "XFD1048576").font?.bold), false);
  assert.equal(sheet.cells.size, 0);
});
