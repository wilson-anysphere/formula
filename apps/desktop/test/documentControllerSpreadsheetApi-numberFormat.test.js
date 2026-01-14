import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../src/document/documentController.js";
import { DocumentControllerSpreadsheetApi } from "../src/ai/tools/documentControllerSpreadsheetApi.ts";

test("DocumentControllerSpreadsheetApi exports number_format from snake_case styles", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 123);
  doc.setRangeFormat("Sheet1", "A1", { number_format: "yyyy-mm-dd" });

  const api = new DocumentControllerSpreadsheetApi(doc);
  const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });

  assert.equal(cell.value, 123);
  assert.equal(cell.format?.number_format, "yyyy-mm-dd");
});

test("DocumentControllerSpreadsheetApi does not fall back to snake_case number_format when numberFormat is cleared", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 123);
  // Simulate imported formula-model formatting (snake_case).
  doc.setRangeFormat("Sheet1", "A1", { number_format: "yyyy-mm-dd" });
  // User clears back to General via UI (camelCase null override).
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: null });

  const api = new DocumentControllerSpreadsheetApi(doc);
  const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });

  assert.equal(cell.value, 123);
  assert.equal(cell.format, undefined);
});

test("DocumentControllerSpreadsheetApi applyFormatting treats 'General' number_format as clearing", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 123);
  doc.setRangeFormat("Sheet1", "A1", { number_format: "yyyy-mm-dd" });

  const api = new DocumentControllerSpreadsheetApi(doc);
  const formatted = api.applyFormatting(
    { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 1 },
    { number_format: "General" },
  );
  assert.equal(formatted, 1);

  const cell = api.getCell({ sheet: "Sheet1", row: 1, col: 1 });
  assert.equal(cell.format, undefined);
});

