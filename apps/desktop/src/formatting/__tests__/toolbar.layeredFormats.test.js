import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold, toggleItalic, toggleSubscript, toggleSuperscript, toggleUnderline, toggleWrap } from "../toolbar.js";

function withGetCellFormatCallLimit(doc, limit, label, fn) {
  const originalGetCellFormat = doc.getCellFormat.bind(doc);
  let calls = 0;
  doc.getCellFormat = (...args) => {
    calls += 1;
    if (calls > limit) {
      throw new Error(`${label} performed too many getCellFormat calls (${calls})`);
    }
    return originalGetCellFormat(...args);
  };
  try {
    return fn();
  } finally {
    doc.getCellFormat = originalGetCellFormat;
  }
}

test("toggleBold reads full-column formatting from the column layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Apply bold to the entire column A via the layered column format path.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "Sheet1 should exist after formatting");
  assert.equal(sheet.cells.size, 0, "Full-column formatting should not materialize per-cell overrides");

  // Toggling again should flip bold OFF (because all cells are currently bold).
  withGetCellFormatCallLimit(doc, 10_000, "toggleBold", () => toggleBold(doc, "Sheet1", "A1:A1048576"));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(sheet.cells.size, 0, "Toggling full-column formatting should not materialize per-cell overrides");
});

test("toggleBold does not scan sheet.cells when sheet has many content-only cells (styleId=0)", () => {
  const doc = new DocumentController();

  // Create a bunch of stored cells with content but no explicit cell-level formatting.
  const values = Array.from({ length: 5000 }, () => ["x"]);
  doc.setRangeValues("Sheet1", "A1", values);

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 5000);
  assert.equal(sheet.styledCells.size, 0);

  // If `toggleBold` iterates `sheet.cells.entries()` (O(#stored cells)) this will throw.
  const originalEntries = sheet.cells.entries;
  sheet.cells.entries = () => {
    throw new Error("sheet.cells.entries() was called");
  };

  toggleBold(doc, "Sheet1", "A1:A1048576");

  // Restore for test hygiene.
  sheet.cells.entries = originalEntries;

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), true);
  // Content cells remain; we should not have materialized extra per-cell style entries.
  assert.equal(sheet.cells.size, 5000);
});

test("toggleWrap reads full-column formatting from the column layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { alignment: { wrapText: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Full-column formatting should not materialize per-cell overrides");

  // Toggling again should flip wrap OFF (because all cells are currently wrapped).
  withGetCellFormatCallLimit(doc, 10_000, "toggleWrap", () => toggleWrap(doc, "Sheet1", "A1:A1048576"));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").alignment?.wrapText), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleItalic reads full-column formatting from the column layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { italic: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0);

  withGetCellFormatCallLimit(doc, 10_000, "toggleItalic", () => toggleItalic(doc, "Sheet1", "A1:A1048576"));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.italic), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleSubscript toggles font.vertAlign between subscript and null", () => {
  const doc = new DocumentController();

  toggleSubscript(doc, "Sheet1", "A1");
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.vertAlign, "subscript");

  toggleSubscript(doc, "Sheet1", "A1");
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.vertAlign, null);
});

test("toggleSuperscript toggles font.vertAlign between superscript and null", () => {
  const doc = new DocumentController();

  toggleSuperscript(doc, "Sheet1", "A1");
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.vertAlign, "superscript");

  toggleSuperscript(doc, "Sheet1", "A1");
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.vertAlign, null);
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
  withGetCellFormatCallLimit(doc, 10_000, "toggleBold", () => toggleBold(doc, "Sheet1", "A1:A1048576"));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), true);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A2").font?.bold), true);
  assert.ok(sheet.cells.size <= 1, "Toggling should not materialize per-cell overrides across the selection");
});

test("toggleBold toggles OFF for large rectangles formatted via range runs (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Keep this below the UI formatting apply guard (100k cells) while still exceeding
  // the range-run threshold (50k cells) so the formatting is stored in runs, not cells.
  const hugeRect = "A1:C20000"; // 60k cells -> range-run layer
  doc.setRangeFormat("Sheet1", hugeRect, { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Range-run formatting should not materialize cells");
  assert.equal(sheet.formatRunsByCol.size, 3, "Expected per-column range runs for A:C");

  withGetCellFormatCallLimit(doc, 10_000, "toggleBold", () => toggleBold(doc, "Sheet1", hugeRect));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "C20000").font?.bold), false);
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

test("toggleBold respects row > col precedence when computing full-row toggle state", () => {
  const doc = new DocumentController();

  // Row 1 is bold by default (row formatting layer).
  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { bold: true } });
  // Column A explicitly sets bold=false, but row formatting overrides column formatting.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: false } });

  let lastPatch = null;
  doc.setRangeFormat = (_sheetId, _range, patch) => {
    lastPatch = patch;
  };

  toggleBold(doc, "Sheet1", "A1:XFD1");
  assert.deepEqual(lastPatch, { font: { bold: false } });
});

test("toggleBold reads full-row formatting from the row layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Entire row 1 via row layer.
  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Full-row formatting should not materialize per-cell overrides");

  // Second toggle should flip bold OFF.
  withGetCellFormatCallLimit(doc, 10_000, "toggleBold", () => toggleBold(doc, "Sheet1", "A1:XFD1"));

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

test("toggleUnderline reads full-row formatting from the row layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { underline: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0);

  withGetCellFormatCallLimit(doc, 10_000, "toggleUnderline", () => toggleUnderline(doc, "Sheet1", "A1:XFD1"));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.underline), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleBold reads full-sheet formatting from the sheet layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Full sheet in Excel address space.
  const fullSheet = "A1:XFD1048576";
  doc.setRangeFormat("Sheet1", fullSheet, { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Sheet formatting should not materialize per-cell overrides");

  withGetCellFormatCallLimit(doc, 10_000, "toggleBold", () => toggleBold(doc, "Sheet1", fullSheet));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "XFD1048576").font?.bold), false);
  assert.equal(sheet.cells.size, 0);
});

test("toggleWrap toggles OFF for large rectangles formatted via range runs (no per-cell scan)", () => {
  const doc = new DocumentController();

  // 60k cells -> range-run layer (see comment above).
  const hugeRect = "A1:C20000";
  doc.setRangeFormat("Sheet1", hugeRect, { alignment: { wrapText: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0, "Range-run formatting should not materialize cells");
  assert.equal(sheet.formatRunsByCol.size, 3, "Expected per-column range runs for A:C");

  withGetCellFormatCallLimit(doc, 10_000, "toggleWrap", () => toggleWrap(doc, "Sheet1", hugeRect));

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").alignment?.wrapText), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "C20000").alignment?.wrapText), false);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "D1").alignment?.wrapText), false);
  assert.equal(sheet.cells.size, 0);
});
