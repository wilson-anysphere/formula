import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import {
  applyFillCommitToDocumentController,
  applyFillCommitToDocumentControllerWithFormulaRewrite,
} from "../applyFillCommit.ts";

test("fill commits normalize formula strings missing leading '=' (defensive)", () => {
  const doc = new DocumentController();

  // Bypass DocumentController's `normalizeFormula` invariant so this test covers
  // defensive behavior for imported/external/corrupt states.
  doc.model.setCell("Sheet1", 0, 0, { value: null, formula: "A1+1", styleId: 0 });

  applyFillCommitToDocumentController({
    document: doc,
    sheetId: "Sheet1",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
  });

  const filled = doc.getCell("Sheet1", "A2");
  assert.equal(filled.formula, "=A2+1");
});

test("fill commits (formula rewrite path) also normalize formulas missing leading '='", async () => {
  const doc = new DocumentController();

  // Use an absolute reference so the expected rewritten formula is stable even when our
  // test rewrite hook is an identity function.
  doc.model.setCell("Sheet1", 0, 0, { value: null, formula: "$A$1+1", styleId: 0 });

  let rewriteCalls = 0;
  await applyFillCommitToDocumentControllerWithFormulaRewrite({
    document: doc,
    sheetId: "Sheet1",
    sourceRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    targetRange: { startRow: 1, endRow: 2, startCol: 0, endCol: 1 },
    mode: "formulas",
    rewriteFormulasForCopyDelta: async (requests) => {
      rewriteCalls += 1;
      return requests.map((r) => r.formula);
    },
  });

  assert.equal(rewriteCalls, 1);
  const filled = doc.getCell("Sheet1", "A2");
  assert.equal(filled.formula, "=$A$1+1");
});

