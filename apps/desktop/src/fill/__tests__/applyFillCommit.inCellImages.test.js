import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import {
  applyFillCommitToDocumentController,
  applyFillCommitToDocumentControllerWithFormulaRewrite,
} from "../applyFillCommit.ts";

test("fill commits coerce in-cell image values to alt text (not [object Object])", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { type: "image", value: { imageId: "img-1", altText: " Logo " } });

  applyFillCommitToDocumentController({
    document: doc,
    sheetId: "Sheet1",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
  });

  assert.equal(doc.getCell("Sheet1", "A2").value, "Logo");
});

test("fill commits coerce image values without alt text to a stable placeholder", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { type: "image", value: { imageId: "img-1" } });

  applyFillCommitToDocumentController({
    document: doc,
    sheetId: "Sheet1",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
  });

  assert.equal(doc.getCell("Sheet1", "A2").value, "[Image]");
});

test("fill (with formula rewrite path) also coerces in-cell images", async () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { type: "image", value: { imageId: "img-1", altText: " Logo " } });

  await applyFillCommitToDocumentControllerWithFormulaRewrite({
    document: doc,
    sheetId: "Sheet1",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
    rewriteFormulasForCopyDelta: () => {
      throw new Error("rewriteFormulasForCopyDelta should not be called for constant fills");
    },
  });

  assert.equal(doc.getCell("Sheet1", "A2").value, "Logo");
});
