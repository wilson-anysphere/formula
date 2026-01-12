import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../../document/documentController.js";
import { applyBranchStateToDocumentController, documentControllerToBranchState } from "../branchStateAdapter.js";

test("branchStateAdapter round-trips sheet metadata (name/visibility/tabColor) with real DocumentController", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  doc.renameSheet("Sheet1", "Budget");
  doc.hideSheet("Sheet1");
  doc.setSheetTabColor("Sheet1", { rgb: "FF00FF00" });

  const state = documentControllerToBranchState(doc);
  assert.equal(state.sheets.metaById.Sheet1?.name, "Budget");
  assert.equal(state.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(state.sheets.metaById.Sheet1?.tabColor, "FF00FF00");

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  const meta = restored.getSheetMeta("Sheet1");
  assert.ok(meta);
  assert.equal(meta.name, "Budget");
  assert.equal(meta.visibility, "hidden");
  assert.equal(meta.tabColor?.rgb, "FF00FF00");
});

