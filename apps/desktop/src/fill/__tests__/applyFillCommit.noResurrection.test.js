import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import {
  applyFillCommitToDocumentController,
  applyFillCommitToDocumentControllerWithFormulaRewrite,
} from "../applyFillCommit.ts";

test("fill commits do not resurrect deleted sheets when called with a stale sheet id (no phantom creation)", async () => {
  const doc = new DocumentController();

  // Ensure Sheet1 exists so deleting Sheet2 doesn't trip the last-sheet guard.
  doc.getCell("Sheet1", { row: 0, col: 0 });
  doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2"]);

  doc.deleteSheet("Sheet2");
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);

  const result = applyFillCommitToDocumentController({
    document: doc,
    sheetId: "Sheet2",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "copy",
  });
  assert.equal(result.editsApplied, 0);
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);

  let rewriteCalls = 0;
  const result2 = await applyFillCommitToDocumentControllerWithFormulaRewrite({
    document: doc,
    sheetId: "Sheet2",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
    rewriteFormulasForCopyDelta: async (requests) => {
      rewriteCalls += 1;
      return requests.map((r) => r.formula);
    },
  });
  assert.equal(result2.editsApplied, 0);
  assert.equal(rewriteCalls, 0);
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);
});

