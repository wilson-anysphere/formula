import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, parseGoTo } from "../index.js";

test("parseGoTo canonicalizes sheet names via workbook.getSheet when available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  // InMemoryWorkbook resolves sheets case-insensitively; parseGoTo should return the
  // canonical sheet name (as stored on the sheet object).
  const parsed = parseGoTo("sheet1!A1", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 });
});

test("parseGoTo canonicalizes currentSheetName for unqualified A1 references", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  const parsed = parseGoTo("B3", { workbook: wb, currentSheetName: "sheet1" });
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 2, endRow: 2, startCol: 1, endCol: 1 });
});

test("parseGoTo supports Excel-style full column references (A:A)", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  const parsed = parseGoTo("A:A", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.source, "a1");
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 1_048_575, startCol: 0, endCol: 0 });
});

test("parseGoTo supports Excel-style full row references (1:1)", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  const parsed = parseGoTo("1:1", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.source, "a1");
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 0, startCol: 0, endCol: 16_383 });
});

test("parseGoTo canonicalizes named range sheet names via workbook.getSheet when available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.defineName("MyRange", { sheetName: "sheet1", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } });

  const parsed = parseGoTo("MyRange", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 1, startCol: 0, endCol: 0 });
});

test("parseGoTo throws for named ranges referring to an unknown sheet when workbook.getSheet is available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.defineName("Bad", { sheetName: "NoSuchSheet", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
  assert.throws(() => parseGoTo("Bad", { workbook: wb, currentSheetName: "Sheet1" }), /Unknown sheet/i);
});

test("parseGoTo supports structured table references (Table1[#All])", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.addTable({
    name: "Table1",
    sheetName: "sheet1", // intentionally wrong case to verify canonicalization
    startRow: 0,
    endRow: 9,
    startCol: 0,
    endCol: 1,
    columns: ["Col1", "Col2"],
  });

  const parsed = parseGoTo("Table1[#All]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.source, "table");
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 9, startCol: 0, endCol: 1 });
});

test("parseGoTo supports structured table specifiers (#Headers/#Data/#Totals)", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.addTable({
    name: "Table1",
    sheetName: "Sheet1",
    startRow: 0,
    endRow: 9,
    startCol: 0,
    endCol: 1,
    columns: ["Col1", "Col2"],
  });

  const headers = parseGoTo("Table1[#Headers]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(headers.source, "table");
  assert.deepEqual(headers.range, { startRow: 0, endRow: 0, startCol: 0, endCol: 1 });

  const data = parseGoTo("Table1[#Data]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(data.source, "table");
  assert.deepEqual(data.range, { startRow: 1, endRow: 9, startCol: 0, endCol: 1 });

  const totals = parseGoTo("Table1[#Totals]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(totals.source, "table");
  assert.deepEqual(totals.range, { startRow: 9, endRow: 9, startCol: 0, endCol: 1 });
});

test("parseGoTo supports structured table column references (Table1[Col2])", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.addTable({
    name: "Table1",
    sheetName: "Sheet1",
    startRow: 0,
    endRow: 9,
    startCol: 0,
    endCol: 1,
    columns: ["Col1", "Col2"],
  });

  const parsed = parseGoTo("Table1[Col2]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.source, "table");
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 9, startCol: 1, endCol: 1 });
});

test("parseGoTo supports selector-qualified structured table column references", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  wb.addTable({
    name: "Table1",
    sheetName: "Sheet1",
    startRow: 0,
    endRow: 9,
    startCol: 0,
    endCol: 1,
    columns: ["Col1", "Col2"],
  });

  const headers = parseGoTo("Table1[[#Headers],[Col2]]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(headers.source, "table");
  assert.deepEqual(headers.range, { startRow: 0, endRow: 0, startCol: 1, endCol: 1 });

  const data = parseGoTo("Table1[[#Data],[Col2]]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(data.source, "table");
  assert.deepEqual(data.range, { startRow: 1, endRow: 9, startCol: 1, endCol: 1 });

  const totals = parseGoTo("Table1[[#Totals],[Col2]]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(totals.source, "table");
  assert.deepEqual(totals.range, { startRow: 9, endRow: 9, startCol: 1, endCol: 1 });

  const all = parseGoTo("Table1[[#All],[Col2]]", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(all.source, "table");
  assert.deepEqual(all.range, { startRow: 0, endRow: 9, startCol: 1, endCol: 1 });
});

test("parseGoTo throws for unknown sheet-qualified references when workbook.getSheet is available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  assert.throws(() => parseGoTo("NoSuchSheet!A1", { workbook: wb, currentSheetName: "Sheet1" }), /Unknown sheet/i);
});
