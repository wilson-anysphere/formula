import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../../document/documentController.js";
import { applyBranchStateToDocumentController, documentControllerToBranchState } from "../branchStateAdapter.js";

test("branchStateAdapter round-trips range-run formatting without materializing per-cell styles", () => {
  const doc = new DocumentController();

  // 3 columns * 20,000 rows = 60,000 cells. This should exceed the range-run threshold and avoid per-cell formatting.
  doc.setRangeFormat("Sheet1", "A1:C20000", { font: { bold: true } });

  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.ok(sheet.formatRunsByCol.size > 0);

  const state = documentControllerToBranchState(doc);

  const exportedRuns = state.sheets.metaById.Sheet1?.view?.formatRunsByCol;
  assert.ok(Array.isArray(exportedRuns));
  assert.ok(exportedRuns.length > 0);
  assert.equal(exportedRuns[0].runs[0].format.font?.bold, true);

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  // Effective formatting should survive checkout.
  assert.equal(restored.getCellFormat("Sheet1", "B100").font?.bold, true);
  // And should still be stored as range runs (no per-cell style materialization).
  assert.equal(restored.getCell("Sheet1", "B100").styleId, 0);

  const stateRoundtrip = documentControllerToBranchState(restored);
  assert.deepEqual(
    stateRoundtrip.sheets.metaById.Sheet1?.view?.formatRunsByCol,
    state.sheets.metaById.Sheet1?.view?.formatRunsByCol,
  );
});

