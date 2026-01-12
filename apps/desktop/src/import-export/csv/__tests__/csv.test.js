import test from "node:test";
import assert from "node:assert/strict";

import { exportCellGridToCsv, exportDocumentRangeToCsv } from "../export.js";
import { parseCsv } from "../csv.js";
import { importCsvToCellGrid } from "../import.js";
import { DocumentController } from "../../../document/documentController.js";
import { dateToExcelSerial } from "../../../shared/valueParsing.js";

test("CSV parses quoted fields, quotes, and embedded newlines", () => {
  const rows = parseCsv('name,notes\nAlice,"Line1\nLine2"\nBob,"He said ""hi"""', {
    delimiter: ",",
  });

  assert.deepEqual(rows, [
    ["name", "notes"],
    ["Alice", "Line1\nLine2"],
    ["Bob", 'He said "hi"'],
  ]);
});

test("CSV import infers column types and preserves header strings", () => {
  const csv = "id,amount,active,date\n001,10,true,2024-01-31\n002,20,false,2024-02-01\n";
  const { grid } = importCsvToCellGrid(csv, { delimiter: "," });

  // Header row stays strings.
  assert.equal(grid[0][0].value, "id");
  assert.equal(grid[0][1].value, "amount");

  // Inferred typing for subsequent rows.
  assert.equal(grid[1][0].value, "001"); // leading zeros preserved
  assert.equal(grid[1][1].value, 10);
  assert.equal(grid[1][2].value, true);
  assert.equal(typeof grid[1][3].value, "number");
  assert.equal(grid[1][3].format.numberFormat, "yyyy-mm-dd");
});

test("CSV import treats leading whitespace before '=' as a formula indicator", () => {
  const csv = "col\n  =SUM(A1:A2)\n=\n";
  const { grid } = importCsvToCellGrid(csv, { delimiter: "," });

  assert.equal(grid[0][0].value, "col");

  assert.equal(grid[1][0].formula, "=SUM(A1:A2)");
  assert.equal(grid[1][0].value, null);

  assert.equal(grid[2][0].formula, "=");
  assert.equal(grid[2][0].value, null);
});

test("CSV export quotes fields when needed", () => {
  const csv = exportCellGridToCsv(
    [
      [{ value: "a" }, { value: "b,c" }],
      [{ value: 1 }, { value: true }],
    ],
    { delimiter: "," }
  );

  assert.equal(csv, 'a,"b,c"\r\n1,TRUE');
});

test("CSV export respects DocumentController cell numberFormat when serializing date numbers", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  doc.setCellValue("Sheet1", "A1", serial);
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "yyyy-mm-dd" });

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1");
  assert.equal(csv, "2024-01-31");
});

test("CSV export respects layered column formats (styleId may be 0)", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  // Apply a full-column format without enumerating every cell.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { numberFormat: "yyyy-mm-dd" });
  doc.setCellValue("Sheet1", "A1000", serial);

  assert.equal(doc.getCell("Sheet1", "A1000").styleId, 0);

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1000");
  assert.equal(csv, "2024-01-31");
});
