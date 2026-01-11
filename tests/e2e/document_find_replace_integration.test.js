import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../apps/desktop/src/document/documentController.js";
import { DocumentWorkbookAdapter } from "../../apps/desktop/src/search/documentWorkbookAdapter.js";
import { FindReplaceController } from "../../apps/desktop/src/panels/find-replace/findReplaceController.js";

test("Find/Replace controller batches replaceAll into a single undo step for DocumentController", async () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "foo");
  doc.setCellValue("Sheet1", "A2", "foo");
  doc.setCellFormula("Sheet1", "B1", "SUM(1,2,3)");

  const workbook = new DocumentWorkbookAdapter({ document: doc });

  let activeCell = { sheetName: "Sheet1", row: 0, col: 0 };

  const controller = new FindReplaceController({
    workbook,
    getCurrentSheetName: () => activeCell.sheetName,
    getActiveCell: () => activeCell,
    setActiveCell: (next) => {
      activeCell = next;
    },
    getSelectionRanges: () => [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }],
    beginBatch: (opts) => doc.beginBatch(opts),
    endBatch: () => doc.endBatch(),
  });

  controller.scope = "sheet";
  controller.lookIn = "values";
  controller.query = "foo";
  controller.replacement = "bar";

  const before = doc.getStackDepths().undo;
  const res = await controller.replaceAll();
  const after = doc.getStackDepths().undo;

  assert.deepEqual(res, { replacedCells: 2, replacedOccurrences: 2 });
  assert.equal(after, before + 1);
  assert.equal(doc.getCell("Sheet1", "A1").value, "bar");
  assert.equal(doc.getCell("Sheet1", "A2").value, "bar");
  assert.equal(doc.getCell("Sheet1", "B1").formula, "=SUM(1,2,3)");
});

test("ReplaceAll (look in formulas) rewrites formula text through DocumentWorkbookAdapter", async () => {
  const doc = new DocumentController();
  doc.setCellFormula("Sheet1", "A1", "SUM(1,2,3)");

  const workbook = new DocumentWorkbookAdapter({ document: doc });
  let activeCell = { sheetName: "Sheet1", row: 0, col: 0 };

  const controller = new FindReplaceController({
    workbook,
    getCurrentSheetName: () => activeCell.sheetName,
    getActiveCell: () => activeCell,
    setActiveCell: (next) => {
      activeCell = next;
    },
    beginBatch: (opts) => doc.beginBatch(opts),
    endBatch: () => doc.endBatch(),
  });

  controller.scope = "sheet";
  controller.lookIn = "formulas";
  controller.query = "SUM";
  controller.replacement = "AVERAGE";

  await controller.replaceAll();

  const stored = doc.getCell("Sheet1", "A1").formula;
  assert.ok(stored?.includes("AVERAGE"), `expected formula to include AVERAGE, got: ${stored}`);
  assert.ok(!stored?.includes("SUM("), `expected SUM to be replaced, got: ${stored}`);
});
