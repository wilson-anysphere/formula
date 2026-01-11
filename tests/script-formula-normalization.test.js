import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";

test("DocumentControllerWorkbookAdapter normalizes formula text consistently", () => {
  const controller = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });
  const sheet = workbook.getActiveSheet();

  sheet.getRange("A1").setFormulas([["  =  SUM(A1:A3)  "]]);
  assert.equal(controller.getCell("Sheet1", "A1").formula, "=SUM(A1:A3)");

  sheet.getRange("A2").setFormulas([["==1+1"]]);
  assert.equal(controller.getCell("Sheet1", "A2").formula, "==1+1");

  // Bare "=" (and whitespace around it) is treated as an empty formula.
  sheet.getRange("A3").setFormulas([["="]]);
  assert.equal(controller.getCell("Sheet1", "A3").formula, null);
  assert.equal(controller.getCell("Sheet1", "A3").value, null);

  sheet.getRange("A4").setFormulas([["   =   "]]);
  assert.equal(controller.getCell("Sheet1", "A4").formula, null);
  assert.equal(controller.getCell("Sheet1", "A4").value, null);

  sheet.getRange("A5").setFormulas([[""]]);
  assert.equal(controller.getCell("Sheet1", "A5").formula, null);
  assert.equal(controller.getCell("Sheet1", "A5").value, null);

  workbook.dispose();
});

