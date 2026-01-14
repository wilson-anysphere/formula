import assert from "node:assert/strict";
import test from "node:test";

import { chunkWorkbook } from "../src/workbook/chunkWorkbook.js";
import { workbookFromSpreadsheetApi } from "../src/workbook/fromSpreadsheetApi.js";

test("workbookFromSpreadsheetApi: default (one-based) input converts to 0-based internal coords", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [
        { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: "Region" } },
        { address: { sheet: "Sheet1", row: 1, col: 2 }, cell: { value: "Revenue" } },
        { address: { sheet: "Sheet1", row: 2, col: 1 }, cell: { value: "North" } },
        { address: { sheet: "Sheet1", row: 2, col: 2 }, cell: { value: 1000 } },
        { address: { sheet: "Sheet1", row: 3, col: 1 }, cell: { value: "South" } },
        { address: { sheet: "Sheet1", row: 3, col: 2 }, cell: { value: 2000 } },
      ];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1" });
  const sheet = workbook.sheets[0];

  assert.equal(sheet.cells.get("0,0")?.value, "Region");
  assert.equal(sheet.cells.get("0,1")?.value, "Revenue");
  assert.equal(sheet.cells.get("1,0")?.value, "North");
  assert.equal(sheet.cells.get("1,1")?.value, 1000);

  const chunks = chunkWorkbook(workbook);
  const dataRegion = chunks.find((c) => c.kind === "dataRegion");
  assert.ok(dataRegion);
  assert.equal(dataRegion.title, "Data region A1:B3");
});

test("workbookFromSpreadsheetApi: coordinateBase='zero' keeps 0-based input", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells() {
      return [{ address: { sheet: "Sheet1", row: 0, col: 0 }, cell: { value: "A1" } }];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1", coordinateBase: "zero" });
  assert.equal(workbook.sheets[0].cells.get("0,0")?.value, "A1");
});

test("workbookFromSpreadsheetApi: coordinateBase='auto' prefers zero-based if any entry contains 0", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells() {
      return [
        { address: { sheet: "Sheet1", row: 0, col: 0 }, cell: { value: "Z" } },
        { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: "A" } },
      ];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1", coordinateBase: "auto" });
  const cells = workbook.sheets[0].cells;
  assert.equal(cells.get("0,0")?.value, "Z");
  assert.equal(cells.get("1,1")?.value, "A");
});

test("workbookFromSpreadsheetApi: drops formatting-only / empty entries", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [
        {
          address: { sheet: "Sheet1", row: 1, col: 1 },
          cell: { value: null, format: { bold: true } },
        },
      ];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1" });
  assert.equal(workbook.sheets[0].cells.size, 0);

  const chunks = chunkWorkbook(workbook);
  assert.ok(!chunks.some((c) => c.kind === "dataRegion" && c.sheetName === "Sheet1"));
});

test("workbookFromSpreadsheetApi: drops cached formula values by default (includeFormulaValues=false)", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [{ address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: 2, formula: "=1+1" } }];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1" });
  const cell = workbook.sheets[0].cells.get("0,0");
  assert.equal(cell?.formula, "=1+1");
  assert.equal(cell?.value, null);
});

test("workbookFromSpreadsheetApi: includeFormulaValues=true preserves cached formula values", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [{ address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: 2, formula: "=1+1" } }];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1", includeFormulaValues: true });
  const cell = workbook.sheets[0].cells.get("0,0");
  assert.equal(cell?.formula, "=1+1");
  assert.equal(cell?.value, 2);
});

test("workbookFromSpreadsheetApi: include_formula_values=true preserves cached formula values", () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [{ address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: 2, formula: "=1+1" } }];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1", include_formula_values: true });
  const cell = workbook.sheets[0].cells.get("0,0");
  assert.equal(cell?.formula, "=1+1");
  assert.equal(cell?.value, 2);
});

test("workbookFromSpreadsheetApi: does not call toString on non-string formula values", () => {
  let calls = 0;
  const formula = {
    toString() {
      calls += 1;
      return "=1+1";
    },
  };

  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [{ address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: 2, formula } }];
    },
  };

  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId: "wb1" });
  assert.equal(calls, 0);
  const cell = workbook.sheets[0].cells.get("0,0");
  // Since the formula is not a string, treat it as absent.
  assert.equal(cell?.formula, null);
  assert.equal(cell?.value, 2);
});
