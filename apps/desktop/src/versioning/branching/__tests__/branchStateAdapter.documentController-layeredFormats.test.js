import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../../document/documentController.js";
import { applyBranchStateToDocumentController, documentControllerToBranchState } from "../branchStateAdapter.js";

test("branchStateAdapter round-trips layered sheet/column formats with real DocumentController", () => {
  const doc = new DocumentController();

  // Apply layered defaults without touching any individual cell formatting.
  doc.setSheetFormat("Sheet1", { fill: { fgColor: "red" } });
  doc.setColFormat("Sheet1", 0, { numberFormat: "yyyy-mm-dd" });

  const state = documentControllerToBranchState(doc);

  assert.deepEqual(state.sheets.metaById.Sheet1.view.defaultFormat, { fill: { fgColor: "red" } });
  assert.deepEqual(state.sheets.metaById.Sheet1.view.colFormats, { "0": { numberFormat: "yyyy-mm-dd" } });

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  // Effective formatting should be present even though the cell itself has styleId=0.
  const a1 = restored.getCell("Sheet1", "A1");
  assert.equal(a1.styleId, 0);

  const format = restored.getCellFormat("Sheet1", "A1");
  assert.equal(format.fill?.fgColor, "red");
  assert.equal(format.numberFormat, "yyyy-mm-dd");
});

