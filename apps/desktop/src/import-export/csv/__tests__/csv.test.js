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

test("CSV strips a UTF-8 BOM from the first field", () => {
  const rows = parseCsv("\uFEFFcol1,col2\n1,2\n", { delimiter: "," });
  assert.deepEqual(rows, [
    ["col1", "col2"],
    ["1", "2"],
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

test("CSV import treats leading apostrophe as a text indicator (Excel convention)", () => {
  const csv = "col\n'001\n";
  const { grid } = importCsvToCellGrid(csv, { delimiter: "," });

  assert.equal(grid[0][0].value, "col");
  assert.equal(grid[1][0].value, "001");
  assert.equal(grid[1][0].formula, null);
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

test("CSV export serializes rich text values as plain text", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { text: "Hello", runs: [] });

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1");
  assert.equal(csv, "Hello");
});

test("CSV export falls back to formula text when no value is present", () => {
  const doc = new DocumentController();
  doc.setCellFormula("Sheet1", "A1", "SUM(1,2)");

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1");
  // Commas require quoting in CSV.
  assert.equal(csv, '"=SUM(1,2)"');
});

test("CSV export escapes literal strings that would otherwise be parsed as formulas", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "=literal");

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1");
  assert.equal(csv, "'=literal");

  const { grid } = importCsvToCellGrid(csv, { delimiter: "," });
  assert.equal(grid[0][0].value, "=literal");
  assert.equal(grid[0][0].formula, null);
});

test("CSV export respects DocumentController cell numberFormat when serializing date numbers", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  doc.setCellValue("Sheet1", "A1", serial);
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "yyyy-mm-dd" });

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1");
  assert.equal(csv, "2024-01-31");
});

test("CSV export treats m/d/yyyy numberFormat as date-like and serializes to an ISO date string", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  doc.setCellValue("Sheet1", "A1", serial);
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "m/d/yyyy" });

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

test("CSV export respects layered row formats (styleId may be 0)", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  // Apply a full-row format without enumerating every cell.
  doc.setRangeFormat("Sheet1", "A10:XFD10", { numberFormat: "yyyy-mm-dd" });
  doc.setCellValue("Sheet1", "C10", serial);

  assert.equal(doc.getCell("Sheet1", "C10").styleId, 0);

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "C10");
  assert.equal(csv, "2024-01-31");
});

test("CSV export respects sheet default formats (styleId may be 0)", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  // Apply a full-sheet default format.
  doc.setRangeFormat("Sheet1", "A1:XFD1048576", { numberFormat: "yyyy-mm-dd" });
  doc.setCellValue("Sheet1", "B2", serial);

  assert.equal(doc.getCell("Sheet1", "B2").styleId, 0);

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "B2");
  assert.equal(csv, "2024-01-31");
});

test("CSV export uses row formats to override column formats (layer precedence)", () => {
  const doc = new DocumentController();
  const dateTime = new Date(Date.UTC(2024, 0, 31, 12, 34, 56));
  const serial = dateToExcelSerial(dateTime);

  // Column sets date-only.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { numberFormat: "yyyy-mm-dd" });
  // Row overrides with date+time.
  doc.setRangeFormat("Sheet1", "A1000:XFD1000", { numberFormat: "yyyy-mm-dd hh:mm:ss" });
  doc.setCellValue("Sheet1", "A1000", serial);

  assert.equal(doc.getCell("Sheet1", "A1000").styleId, 0);

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A1000");
  assert.equal(csv, dateTime.toISOString());
});

test("CSV export respects range-run formats for large rectangles (styleId may be 0)", () => {
  const doc = new DocumentController();
  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));

  // This range is large enough to trigger DocumentController's compressed range-run formatting layer
  // (instead of enumerating every cell).
  doc.setRangeFormat("Sheet1", "A1:A50001", { numberFormat: "yyyy-mm-dd" });
  doc.setCellValue("Sheet1", "A50001", serial);

  assert.equal(doc.getCell("Sheet1", "A50001").styleId, 0);
  assert.notEqual(doc.getCellFormatStyleIds("Sheet1", "A50001")[4], 0);

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A50001");
  assert.equal(csv, "2024-01-31");
});

test("CSV export prefers range-run formats over row formats (layer precedence)", () => {
  const doc = new DocumentController();
  const dateTime = new Date(Date.UTC(2024, 0, 31, 12, 34, 56));
  const serial = dateToExcelSerial(dateTime);

  // Create a range-run style (higher precedence than row/col defaults).
  doc.setRangeFormat("Sheet1", "A1:A50001", { numberFormat: "yyyy-mm-dd" });
  // Set a conflicting row format (lower precedence than range runs).
  // A50001 is 1-based row 50001 -> 0-based row index 50000.
  doc.setRowFormat("Sheet1", 50000, { numberFormat: "yyyy-mm-dd hh:mm:ss" });

  doc.setCellValue("Sheet1", "A50001", serial);
  assert.equal(doc.getCell("Sheet1", "A50001").styleId, 0);

  const tuple = doc.getCellFormatStyleIds("Sheet1", "A50001");
  assert.notEqual(tuple[1], 0); // row style id
  assert.notEqual(tuple[4], 0); // range-run style id

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A50001");
  assert.equal(csv, "2024-01-31");
});

test("CSV export prefers cell formats over range-run formats (layer precedence)", () => {
  const doc = new DocumentController();
  const dateTime = new Date(Date.UTC(2024, 0, 31, 12, 34, 56));
  const serial = dateToExcelSerial(dateTime);

  // Create a range-run style.
  doc.setRangeFormat("Sheet1", "A1:A50001", { numberFormat: "yyyy-mm-dd" });
  // Then apply an explicit per-cell format with higher precedence than range runs.
  doc.setCellInput("Sheet1", "A50001", { value: serial, format: { numberFormat: "yyyy-mm-dd hh:mm:ss" } });

  const tuple = doc.getCellFormatStyleIds("Sheet1", "A50001");
  assert.notEqual(tuple[3], 0); // cell style id
  assert.notEqual(tuple[4], 0); // range-run style id

  const csv = exportDocumentRangeToCsv(doc, "Sheet1", "A50001");
  assert.equal(csv, dateTime.toISOString());
});
