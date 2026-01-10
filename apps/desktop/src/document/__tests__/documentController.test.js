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

  doc.setCellFormula("Sheet1", "A1", "SUM(B1:B3)");
  assert.equal(doc.getCell("Sheet1", "A1").formula, "SUM(B1:B3)");
  assert.equal(doc.getCell("Sheet1", "A1").value, null);

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").formula, null);
});

test("setRangeFormat is undoable (formatting changes)", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setRangeFormat("Sheet1", "A1", { bold: true });
  assert.deepEqual(doc.getCell("Sheet1", "A1").format, { bold: true });

  doc.undo();
  assert.equal(doc.getCell("Sheet1", "A1").format, null);
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
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
