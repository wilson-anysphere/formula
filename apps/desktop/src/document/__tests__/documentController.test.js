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
