import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";
import { MockEngine } from "../engine.js";

test("setCellValue -> undo -> redo stays in sync with engine", () => {
  const engine = new MockEngine();
  const doc = new DocumentController({ engine });

  doc.setCellValue("Sheet1", "A1", "hello", { label: "Edit A1" });
  assert.equal(doc.getCell("Sheet1", "A1").value, "hello");
  assert.equal(engine.getCell("Sheet1", 0, 0).value, "hello");
  assert.equal(doc.isDirty, true);
  assert.deepEqual(doc.getStackDepths(), { undo: 1, redo: 0 });

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(engine.getCell("Sheet1", 0, 0).value, null);
  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 1 });

  doc.redo();
  assert.equal(doc.getCell("Sheet1", "A1").value, "hello");
  assert.equal(engine.getCell("Sheet1", 0, 0).value, "hello");
  assert.deepEqual(doc.getStackDepths(), { undo: 1, redo: 0 });
});

test("batching collapses multiple edits into a single undo step", () => {
  const doc = new DocumentController();

  doc.beginBatch({ label: "Typing" });
  doc.setCellValue("Sheet1", "A1", "a");
  doc.setCellValue("Sheet1", "A1", "ab");
  doc.setCellValue("Sheet1", "A1", "abc");
  doc.endBatch();

  assert.equal(doc.getCell("Sheet1", "A1").value, "abc");
  assert.deepEqual(doc.getStackDepths(), { undo: 1, redo: 0 });

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 1 });
});

test("dirty tracking toggles when undoing back to saved state", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  doc.setCellValue("Sheet1", "A1", 2);
  assert.equal(doc.isDirty, true);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
  assert.equal(doc.isDirty, false);
});

test("setRangeValues + clearRange invert correctly via undo", () => {
  const doc = new DocumentController();

  doc.setRangeValues("Sheet1", "A1", [
    [1, 2],
    [3, 4],
  ]);
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
  assert.equal(doc.getCell("Sheet1", "B2").value, 4);

  doc.clearRange("Sheet1", "A1:B2");
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "B2").value, null);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
  assert.equal(doc.getCell("Sheet1", "B2").value, 4);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "B2").value, null);
});

test("clearRange iterates sparse stored cells (does not scale with range area)", () => {
  const doc = new DocumentController();

  // Populate a couple of cells (and give them formatting) so there are only a few stored cells.
  doc.setCellValue("Sheet1", "A1", "hello");
  doc.setCellFormula("Sheet1", "A5000", "1+2");
  doc.setCellValue("Sheet1", "B1", "keep");

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "A5000", { font: { bold: true } });

  const a1Before = doc.getCell("Sheet1", "A1");
  const a5000Before = doc.getCell("Sheet1", "A5000");

  let getCellCalls = 0;
  const originalGetCell = doc.model.getCell;
  doc.model.getCell = (sheetId, row, col) => {
    getCellCalls += 1;
    return originalGetCell.call(doc.model, sheetId, row, col);
  };

  doc.clearRange("Sheet1", "A1:A10000");

  // Restore before assertions (doc.getCell below calls through the model).
  doc.model.getCell = originalGetCell;

  // If clearRange scanned the full rectangle, this would be ~10,000.
  assert.ok(getCellCalls < 100, `expected < 100 getCell calls, got ${getCellCalls}`);

  const a1After = doc.getCell("Sheet1", "A1");
  assert.equal(a1After.value, null);
  assert.equal(a1After.formula, null);
  assert.equal(a1After.styleId, a1Before.styleId);

  const a5000After = doc.getCell("Sheet1", "A5000");
  assert.equal(a5000After.value, null);
  assert.equal(a5000After.formula, null);
  assert.equal(a5000After.styleId, a5000Before.styleId);

  // Cells outside the cleared range should be untouched.
  assert.equal(doc.getCell("Sheet1", "B1").value, "keep");
});

test("setCellFormula is undoable", () => {
  const doc = new DocumentController();

  doc.setCellFormula("Sheet1", "A1", " SUM(B1:B3)");
  assert.equal(doc.getCell("Sheet1", "A1").formula, "=SUM(B1:B3)");
  assert.equal(doc.getCell("Sheet1", "A1").value, null);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").formula, null);
});

test("formula normalization trims and treats bare '=' as empty", () => {
  const doc = new DocumentController();

  doc.setCellFormula("Sheet1", "A1", "  =  SUM(A1:A3)  ");
  assert.equal(doc.getCell("Sheet1", "A1").formula, "=SUM(A1:A3)");

  doc.setCellFormula("Sheet1", "A2", "==1+1");
  assert.equal(doc.getCell("Sheet1", "A2").formula, "==1+1");

  // Empty formulas (including a bare "=") clear the cell's formula.
  doc.setCellFormula("Sheet1", "A3", "=");
  assert.equal(doc.getCell("Sheet1", "A3").formula, null);
  assert.equal(doc.getCell("Sheet1", "A3").value, null);

  doc.setCellInput("Sheet1", "A4", "   =   ");
  assert.equal(doc.getCell("Sheet1", "A4").formula, null);
  assert.equal(doc.getCell("Sheet1", "A4").value, null);

  doc.setCellInput("Sheet1", "A5", "=");
  assert.equal(doc.getCell("Sheet1", "A5").formula, null);
  assert.equal(doc.getCell("Sheet1", "A5").value, null);
});

test("setRangeFormat is undoable (formatting changes)", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  const styled = doc.getCell("Sheet1", "A1");
  assert.equal(styled.styleId, 1);
  assert.deepEqual(doc.styleTable.get(styled.styleId), { font: { bold: true } });

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
});

test("getUsedRange can include format-only cells", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "B2", { font: { bold: true } });

  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 1,
    endRow: 1,
    startCol: 1,
    endCol: 1,
  });
});

test("setRangeFormat for full-height columns does not materialize 1M cell entries", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0);
  assert.equal(sheet.colStyleIds.size, 1);

  // Deep row formatting applies even when the cell is empty.
  assert.equal(doc.getCell("Sheet1", "A1048576").styleId, 0);
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1048576"), { font: { bold: true } });

  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 1_048_575,
    startCol: 0,
    endCol: 0,
  });

  // Undo/redo works for column-level formatting.
  doc.undo();
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1048576"), {});

  doc.redo();
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1048576"), { font: { bold: true } });
});

test("getUsedRange maintains separate content and format bounds", () => {
  const doc = new DocumentController();

  // Style-only cell should only affect includeFormat bounds.
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

  // Content cells establish the default used range.
  doc.setCellFormula("Sheet1", "B2", "1+1");
  doc.setCellValue("Sheet1", "C3", "x");

  assert.deepEqual(doc.getUsedRange("Sheet1"), {
    startRow: 1,
    endRow: 2,
    startCol: 1,
    endCol: 2,
  });

  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 2,
  });
});

test("getUsedRange recomputes bounds only when clearing a boundary content cell", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet1", "B2", 2);
  doc.setCellValue("Sheet1", "C3", 3);

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  // No rescans needed for inserts or repeated reads.
  assert.deepEqual(doc.getUsedRange("Sheet1"), {
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 2,
  });
  assert.equal(sheet.__contentBoundsRecomputeCount, 0);
  doc.getUsedRange("Sheet1");
  assert.equal(sheet.__contentBoundsRecomputeCount, 0);

  // Clearing an interior cell should not invalidate bounds.
  doc.clearCell("Sheet1", "B2");
  assert.deepEqual(doc.getUsedRange("Sheet1"), {
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 2,
  });
  assert.equal(sheet.__contentBoundsRecomputeCount, 0);

  // Clearing a boundary cell requires a rescan to shrink.
  doc.clearCell("Sheet1", "C3");
  assert.deepEqual(doc.getUsedRange("Sheet1"), {
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
  });
  assert.equal(sheet.__contentBoundsRecomputeCount, 1);

  // Subsequent reads reuse the cached bounds.
  doc.getUsedRange("Sheet1");
  assert.equal(sheet.__contentBoundsRecomputeCount, 1);
});

test("clearRange preserves style-only cells for includeFormat used range", () => {
  const doc = new DocumentController();

  // Establish a format-only region.
  doc.setRangeFormat("Sheet1", "A1:C3", { font: { italic: true } });

  // Add some content inside the formatted region.
  doc.setCellValue("Sheet1", "B2", 1);
  doc.setCellValue("Sheet1", "C3", 2);

  assert.deepEqual(doc.getUsedRange("Sheet1"), {
    startRow: 1,
    endRow: 2,
    startCol: 1,
    endCol: 2,
  });
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 2,
  });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  // Clearing content should not clear formatting, so includeFormat bounds remain.
  doc.clearRange("Sheet1", "B2:C3");
  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 2,
  });
  assert.equal(sheet.__formatBoundsRecomputeCount, 0);
});

test("getCellFormatStyleIds exposes layered style id tuple (sheet/row/col/cell)", () => {
  const doc = new DocumentController();

  // Column default: bold.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const [sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId] = doc.getCellFormatStyleIds("Sheet1", "A1");
  assert.equal(sheetDefaultStyleId, 0);
  assert.equal(rowStyleId, 0);
  assert.equal(cellStyleId, 0);
  assert.equal(Boolean(doc.styleTable.get(colStyleId).font?.bold), true);

  // Convenience accessors match.
  assert.equal(doc.getSheetDefaultStyleId("Sheet1"), sheetDefaultStyleId);
  assert.equal(doc.getRowStyleId("Sheet1", 0), rowStyleId);
  assert.equal(doc.getColStyleId("Sheet1", 0), colStyleId);
});

test("setFrozen is undoable", () => {
  const doc = new DocumentController();

  doc.setFrozen("Sheet1", 2, 1);
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });

  doc.undo();
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });

  doc.redo();
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });
});

test("setFrozen preserves row/col size overrides", () => {
  const doc = new DocumentController();

  doc.setColWidth("Sheet1", 0, 120);
  doc.setRowHeight("Sheet1", 1, 40);

  doc.setFrozen("Sheet1", 2, 1);
  assert.deepEqual(doc.getSheetView("Sheet1"), {
    frozenRows: 2,
    frozenCols: 1,
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
});

test("undo/redo of sheet view changes does not trigger engine recalc", () => {
  const engine = new MockEngine();
  const doc = new DocumentController({ engine });

  doc.setColWidth("Sheet1", 0, 120);
  assert.equal(engine.recalcCount, 0);

  doc.undo();
  assert.equal(engine.recalcCount, 0);

  doc.redo();
  assert.equal(engine.recalcCount, 0);
});

test("sheet view updates increment updateVersion but not contentVersion", () => {
  const doc = new DocumentController();
  assert.equal(doc.updateVersion, 0);
  assert.equal(doc.contentVersion, 0);

  doc.setFrozen("Sheet1", 2, 1);
  assert.equal(doc.updateVersion, 1);
  assert.equal(doc.contentVersion, 0);

  doc.setColWidth("Sheet1", 0, 120);
  assert.equal(doc.updateVersion, 2);
  assert.equal(doc.contentVersion, 0);

  doc.setCellValue("Sheet1", "A1", "hello");
  assert.equal(doc.updateVersion, 3);
  assert.equal(doc.contentVersion, 1);
});

test("setColWidth/setRowHeight are undoable", () => {
  const doc = new DocumentController();

  doc.setColWidth("Sheet1", 0, 120, { label: "Resize Column" });
  doc.setRowHeight("Sheet1", 1, 40, { label: "Resize Row" });

  assert.deepEqual(doc.getSheetView("Sheet1"), {
    frozenRows: 0,
    frozenCols: 0,
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });

  doc.undo();
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, colWidths: { "0": 120 } });

  doc.undo();
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });

  doc.redo();
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, colWidths: { "0": 120 } });

  doc.redo();
  assert.deepEqual(doc.getSheetView("Sheet1"), {
    frozenRows: 0,
    frozenCols: 0,
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
});

test("mergeKey collapses consecutive edits into one history entry, but saving stops merging", () => {
  const doc = new DocumentController({ mergeWindowMs: 10_000 });

  doc.setCellValue("Sheet1", "A1", "a", { mergeKey: "cell:A1" });
  doc.setCellValue("Sheet1", "A1", "ab", { mergeKey: "cell:A1" });
  doc.setCellValue("Sheet1", "A1", "abc", { mergeKey: "cell:A1" });
  assert.deepEqual(doc.getStackDepths(), { undo: 1, redo: 0 });

  doc.markSaved();
  assert.equal(doc.isDirty, false);

  doc.setCellValue("Sheet1", "A1", "abcd", { mergeKey: "cell:A1" });
  assert.deepEqual(doc.getStackDepths(), { undo: 2, redo: 0 });
  assert.equal(doc.isDirty, true);
});

test("engine recalc is deferred to endBatch", () => {
  const engine = new MockEngine();
  const doc = new DocumentController({ engine });

  doc.beginBatch({ label: "Paste" });
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet1", "B1", 2);
  assert.equal(engine.recalcCount, 0);
  doc.endBatch();
  assert.equal(engine.recalcCount, 1);
});

test("dirty tracking considers in-progress batches (before endBatch)", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  doc.beginBatch({ label: "Typing" });
  doc.setCellValue("Sheet1", "A1", 2);
  assert.equal(doc.isDirty, true);

  doc.endBatch();
  assert.equal(doc.isDirty, true);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
  assert.equal(doc.isDirty, false);
});

test("history enablement updates when entering/leaving an empty batch", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");

  /** @type {{ canUndo: boolean, canRedo: boolean }[]} */
  const states = [];
  doc.on("history", (payload) => {
    states.push(payload);
  });

  assert.equal(doc.canUndo, true);
  doc.beginBatch({ label: "No-op" });
  assert.equal(doc.canUndo, false);
  doc.endBatch();
  assert.equal(doc.canUndo, true);

  // We only subscribe after the first edit, so the state sequence should reflect
  // entering/leaving the empty batch.
  assert.deepEqual(states, [
    { canUndo: false, canRedo: false },
    { canUndo: true, canRedo: false },
  ]);
});

test("setCellInput interprets '=' as formula and apostrophe as literal text", () => {
  const doc = new DocumentController();

  doc.setCellInput("Sheet1", "A1", "=1+2");
  assert.equal(doc.getCell("Sheet1", "A1").formula, "=1+2");
  assert.equal(doc.getCell("Sheet1", "A1").value, null);

  doc.setCellInput("Sheet1", "A3", "   =1+2");
  assert.equal(doc.getCell("Sheet1", "A3").formula, "=1+2");

  doc.setCellInput("Sheet1", "A2", "'=1+2");
  assert.equal(doc.getCell("Sheet1", "A2").formula, null);
  assert.equal(doc.getCell("Sheet1", "A2").value, "=1+2");
});

test("setCellInput coerces numeric/boolean literals (Excel-style)", () => {
  const doc = new DocumentController();

  doc.setCellInput("Sheet1", "A1", "123");
  assert.equal(doc.getCell("Sheet1", "A1").value, 123);

  doc.setCellInput("Sheet1", "A2", "TRUE");
  assert.equal(doc.getCell("Sheet1", "A2").value, true);

  doc.setCellInput("Sheet1", "A3", "false");
  assert.equal(doc.getCell("Sheet1", "A3").value, false);

  // Apostrophe forces text (no coercion).
  doc.setCellInput("Sheet1", "A4", "'123");
  assert.equal(doc.getCell("Sheet1", "A4").value, "123");
  doc.setCellInput("Sheet1", "A5", "'TRUE");
  assert.equal(doc.getCell("Sheet1", "A5").value, "TRUE");
});

test("setRangeValues treats strings starting with '=' as formulas", () => {
  const doc = new DocumentController();

  doc.setRangeValues("Sheet1", "A1", [["=A2+1", "'=literal", { formula: "A1+1" }]]);
  assert.equal(doc.getCell("Sheet1", "A1").formula, "=A2+1");
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "B1").formula, null);
  assert.equal(doc.getCell("Sheet1", "B1").value, "=literal");
  assert.equal(doc.getCell("Sheet1", "C1").formula, "=A1+1");
});

test("cancelBatch reverts uncommitted batch changes without affecting history", () => {
  const engine = new MockEngine();
  const doc = new DocumentController({ engine });

  doc.markSaved();
  assert.equal(doc.isDirty, false);
  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });

  doc.beginBatch({ label: "Typing" });
  doc.setCellInput("Sheet1", "A1", "=1+2");
  assert.equal(doc.canUndo, false);
  assert.equal(engine.recalcCount, 0);
  assert.equal(doc.getCell("Sheet1", "A1").formula, "=1+2");

  const reverted = doc.cancelBatch();
  assert.equal(reverted, true);
  assert.equal(doc.getCell("Sheet1", "A1").formula, null);
  assert.equal(engine.getCell("Sheet1", 0, 0).formula, null);
  assert.equal(engine.recalcCount, 1);
  assert.equal(doc.isDirty, false);
  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });
});
