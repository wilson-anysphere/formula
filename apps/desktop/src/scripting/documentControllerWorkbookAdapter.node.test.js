import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../document/documentController.js";
import { DocumentControllerWorkbookAdapter } from "./documentControllerWorkbookAdapter.js";

test("DocumentControllerWorkbookAdapter uses stable sheet ids for usedRange + formats (no phantom display-name sheets)", () => {
  const controller = new DocumentController();
  // Use a stable id that differs from the display name.
  controller.setCellValue("sheet-1", "C3", 1);

  const sheetNameResolver = {
    getSheetNameById: (id) => (String(id) === "sheet-1" ? "Budget" : null),
    getSheetIdByName: (name) => (String(name).trim().toLowerCase() === "budget" ? "sheet-1" : null),
  };

  const workbook = new DocumentControllerWorkbookAdapter(controller, { sheetNameResolver });
  const sheet = workbook.getSheet("Budget");

  assert.equal(sheet.name, "Budget");
  assert.equal(sheet.sheetId, "sheet-1");

  // getUsedRange should use the stable id so it reflects C3 rather than falling back to A1.
  const used = sheet.getUsedRange();
  assert.equal(used.address, "C3");
  assert.deepEqual(controller.getSheetIds(), ["sheet-1"]);

  // getFormats/setFormats must not materialize a phantom sheet keyed by the display name.
  sheet.getRange("C3").getFormats();
  assert.deepEqual(controller.getSheetIds(), ["sheet-1"]);

  sheet.getRange("C3").setFormats([[{ bold: true }]]);
  assert.deepEqual(controller.getSheetIds(), ["sheet-1"]);

  const effective = controller.getCellFormat("sheet-1", "C3");
  assert.equal(effective?.font?.bold, true);

  workbook.dispose();
});

test("DocumentControllerWorkbookAdapter guards getValues against huge ranges", () => {
  const controller = new DocumentController();

  let scanned = 0;
  const origGetCell = controller.getCell.bind(controller);
  controller.getCell = (...args) => {
    scanned += 1;
    return origGetCell(...args);
  };

  const workbook = new DocumentControllerWorkbookAdapter(controller, { activeSheetName: "Sheet1" });
  const sheet = workbook.getSheet("Sheet1");

  scanned = 0;
  assert.throws(() => sheet.getRange("A1:Z8000").getValues(), /getValues skipped/i);
  assert.equal(scanned, 0);

  workbook.dispose();
});

test("DocumentControllerWorkbookAdapter surfaces snake_case number_format and respects cleared numberFormat", () => {
  const controller = new DocumentController();

  // Use a stable id that differs from the display name.
  controller.setCellValue("sheet-1", "A1", 1);

  const importedStyleId = controller.styleTable.intern({ number_format: "0.00" });
  controller.setRangeValues("sheet-1", "A1", [[{ value: 1.23, styleId: importedStyleId }]]);

  // If the UI clears a number format (numberFormat: null), that should override any imported snake_case value.
  const clearedStyleId = controller.styleTable.intern({ number_format: "0.00", numberFormat: null });
  controller.setRangeValues("sheet-1", "A2", [[{ value: 1.23, styleId: clearedStyleId }]]);

  const sheetNameResolver = {
    getSheetNameById: (id) => (String(id) === "sheet-1" ? "Budget" : null),
    getSheetIdByName: (name) => (String(name).trim().toLowerCase() === "budget" ? "sheet-1" : null),
  };

  const workbook = new DocumentControllerWorkbookAdapter(controller, { sheetNameResolver });
  const sheet = workbook.getSheet("Budget");

  assert.equal(sheet.getRange("A1").getFormat().numberFormat, "0.00");
  assert.equal(sheet.getRange("A2").getFormat().numberFormat, undefined);

  workbook.dispose();
});

test("DocumentControllerWorkbookAdapter treats numberFormat='General' as clearing (stores numberFormat: null)", () => {
  const controller = new DocumentController();

  controller.setCellValue("sheet-1", "A1", 1);
  const importedStyleId = controller.styleTable.intern({ number_format: "0.00" });
  controller.setRangeValues("sheet-1", "A1", [[{ value: 1.23, styleId: importedStyleId }]]);

  const sheetNameResolver = {
    getSheetNameById: (id) => (String(id) === "sheet-1" ? "Budget" : null),
    getSheetIdByName: (name) => (String(name).trim().toLowerCase() === "budget" ? "sheet-1" : null),
  };

  const workbook = new DocumentControllerWorkbookAdapter(controller, { sheetNameResolver });
  const sheet = workbook.getSheet("Budget");

  sheet.getRange("A1").setFormat({ numberFormat: "General" });

  const cell = controller.getCell("sheet-1", "A1");
  const style = controller.styleTable.get(cell.styleId);
  assert.equal(style.numberFormat, null);

  // Scripts should observe "General" as a cleared number format.
  assert.equal(sheet.getRange("A1").getFormat().numberFormat, undefined);

  workbook.dispose();
});
