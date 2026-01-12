import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("renameSheet emits sheetMetaDeltas in change payload", () => {
  const doc = new DocumentController();
  doc.getCell("Sheet1", "A1"); // materialize sheet without emitting changes

  /** @type {any} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.renameSheet("Sheet1", "Budget");

  assert.equal(lastChange?.source, undefined);
  assert.equal(lastChange?.recalc, false);
  assert.ok(Array.isArray(lastChange?.sheetMetaDeltas));
  assert.equal(lastChange.sheetMetaDeltas.length, 1);
  assert.deepEqual(lastChange.sheetMetaDeltas[0], {
    sheetId: "Sheet1",
    before: { name: "Sheet1", visibility: "visible" },
    after: { name: "Budget", visibility: "visible" },
  });
});

test("reorderSheets emits sheetOrderDelta in change payload", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);
  doc.setCellValue("Sheet3", "A1", 3);

  /** @type {any} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.reorderSheets(["Sheet3", "Sheet1", "Sheet2"]);

  assert.equal(lastChange?.recalc, false);
  assert.deepEqual(lastChange?.sheetOrderDelta, {
    before: ["Sheet1", "Sheet2", "Sheet3"],
    after: ["Sheet3", "Sheet1", "Sheet2"],
  });
  assert.deepEqual(doc.getSheetIds(), ["Sheet3", "Sheet1", "Sheet2"]);
});

test("setSheetTabColor accepts ARGB string and emits a meta delta", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  /** @type {any} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.setSheetTabColor("Sheet1", "FF00FF00");

  assert.ok(Array.isArray(lastChange?.sheetMetaDeltas));
  assert.equal(lastChange.sheetMetaDeltas.length, 1);
  assert.equal(lastChange.sheetMetaDeltas[0].sheetId, "Sheet1");
  assert.equal(lastChange.sheetMetaDeltas[0].before?.tabColor, undefined);
  assert.deepEqual(lastChange.sheetMetaDeltas[0].after?.tabColor, { rgb: "FF00FF00" });
  assert.deepEqual(doc.getSheetMeta("Sheet1")?.tabColor, { rgb: "FF00FF00" });
});

