import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";

test("DocumentController.applyExternalFormatDeltas applies layered column formatting without creating an undo step", () => {
  const doc = new DocumentController();
  const boldStyleId = doc.styleTable.intern({ font: { bold: true } });

  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });

  doc.applyExternalFormatDeltas([
    {
      sheetId: "Sheet1",
      layer: "col",
      index: 0,
      beforeStyleId: 0,
      afterStyleId: boldStyleId,
    },
  ]);

  // External format deltas must not create undo/redo history.
  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "expected Sheet1 to exist");
  assert.equal(sheet.colStyleIds.get(0), boldStyleId);

  // Layered formatting affects the effective cell format, but should not materialize per-cell style ids.
  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);
  assert.equal(doc.getCellFormat("Sheet1", "A1")?.font?.bold, true);
});

test("DocumentController.applyExternalRangeRunDeltas applies range-run formatting without creating an undo step", () => {
  const doc = new DocumentController();
  const italicStyleId = doc.styleTable.intern({ font: { italic: true } });

  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });

  doc.applyExternalRangeRunDeltas([
    {
      sheetId: "Sheet1",
      col: 0,
      startRow: 0,
      endRowExclusive: 50_001,
      beforeRuns: [],
      afterRuns: [{ startRow: 0, endRowExclusive: 50_001, styleId: italicStyleId }],
    },
  ]);

  assert.deepEqual(doc.getStackDepths(), { undo: 0, redo: 0 });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "expected Sheet1 to exist");
  assert.deepEqual(sheet.formatRunsByCol.get(0), [{ startRow: 0, endRowExclusive: 50_001, styleId: italicStyleId }]);

  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);
  assert.equal(doc.getCellFormat("Sheet1", "A1")?.font?.italic, true);

  // The run is a half-open interval, so the row immediately after should not be styled.
  assert.equal(doc.getCellFormat("Sheet1", "A50002")?.font?.italic, undefined);
});
