import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import { applyBranchStateToDocumentController, documentControllerToBranchState } from "./branchStateAdapter.js";

test("documentControllerToBranchState/applyBranchStateToDocumentController: preserves sheet order", () => {
  const doc = new DocumentController();

  // Create sheets in a non-lexicographic order.
  doc.setCellValue("SheetB", "A1", 1);
  doc.setCellValue("SheetA", "A1", 1);
  doc.setCellValue("SheetC", "A1", 1);

  const state = documentControllerToBranchState(doc);
  assert.deepEqual(state.sheets.order, ["SheetB", "SheetA", "SheetC"]);

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);
  assert.deepEqual(restored.getSheetIds(), ["SheetB", "SheetA", "SheetC"]);

  const roundTrip = documentControllerToBranchState(restored);
  assert.deepEqual(roundTrip.sheets.order, ["SheetB", "SheetA", "SheetC"]);
});

test("documentControllerToBranchState/applyBranchStateToDocumentController: round-trips sheet display names", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);
  doc.renameSheet("Sheet1", "Summary");
  doc.renameSheet("Sheet2", "Data");

  const state = documentControllerToBranchState(doc);
  assert.equal(state.sheets.metaById.Sheet1?.name, "Summary");
  assert.equal(state.sheets.metaById.Sheet2?.name, "Data");

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  assert.equal(restored.getSheetMeta("Sheet1")?.name, "Summary");
  assert.equal(restored.getSheetMeta("Sheet2")?.name, "Data");

  const roundTrip = documentControllerToBranchState(restored);
  assert.equal(roundTrip.sheets.metaById.Sheet1?.name, "Summary");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.name, "Data");
});

test("documentControllerToBranchState/applyBranchStateToDocumentController: includes sheet visibility + tabColor when present", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.renameSheet("Sheet1", "Summary");
  doc.hideSheet("Sheet1");
  doc.setSheetTabColor("Sheet1", "FF00FF00");

  doc.setCellValue("Sheet2", "A1", 2);
  doc.renameSheet("Sheet2", "Data");
  // Explicitly clear any tab color (DocumentController stores this as `undefined`).
  doc.setSheetTabColor("Sheet2", null);

  const state = documentControllerToBranchState(doc);
  assert.equal(state.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(state.sheets.metaById.Sheet1?.tabColor, "FF00FF00");
  assert.equal(state.sheets.metaById.Sheet2?.visibility, "visible");
  assert.equal(state.sheets.metaById.Sheet2?.tabColor, null);

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  const meta1 = restored.getSheetMeta("Sheet1");
  assert.ok(meta1);
  assert.equal(meta1.name, "Summary");
  assert.equal(meta1.visibility, "hidden");
  assert.equal(meta1.tabColor?.rgb, "FF00FF00");

  const meta2 = restored.getSheetMeta("Sheet2");
  assert.ok(meta2);
  assert.equal(meta2.name, "Data");
  assert.equal(meta2.visibility, "visible");
  assert.equal(meta2.tabColor, undefined);

  const roundTrip = documentControllerToBranchState(restored);
  assert.equal(roundTrip.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(roundTrip.sheets.metaById.Sheet1?.tabColor, "FF00FF00");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.visibility, "visible");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.tabColor, null);
});

test("applyBranchStateToDocumentController: masks cells with enc=null markers", () => {
  const doc = new DocumentController();

  applyBranchStateToDocumentController(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } },
    },
    cells: {
      Sheet1: {
        A1: { enc: null },
      },
    },
    metadata: {},
    namedRanges: {},
    comments: {},
  });

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.value, "###");
  assert.equal(cell.formula, null);
});
