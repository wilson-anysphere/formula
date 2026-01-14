import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, findAll, formatCellValue } from "../index.js";

test("merged cells: selection scope treats merged region as the top-left cell", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Merge A1:B2 and put the value in A1 (Excel semantics).
  sheet.mergeCells({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
  sheet.setValue(0, 0, "foo");

  // Select B2 only – Excel still searches the merged cell (A1).
  const matches = await findAll(wb, "foo", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [{ startRow: 1, endRow: 1, startCol: 1, endCol: 1 }],
  });

  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!A1"]);
});

test("valueMode: display vs raw is respected for typed values", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setCell(0, 0, { value: { t: "n", v: 1.234 }, display: "1.23" });

  const displayMatches = await findAll(wb, "1.23", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "display",
    matchEntireCell: true,
  });
  assert.deepEqual(displayMatches.map((m) => m.address), ["Sheet1!A1"]);

  const rawMatches = await findAll(wb, "1.23", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "raw",
    matchEntireCell: true,
  });
  assert.deepEqual(rawMatches.map((m) => m.address), []);
});

test("error cells: display/raw formatting matches #DIV/0! string", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setCell(0, 1, { value: { t: "e", v: "#DIV/0!" } });

  const matches = await findAll(wb, "#DIV/0!", { scope: "sheet", currentSheetName: "Sheet1" });
  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!B1"]);
  assert.equal(matches[0].text, "#DIV/0!");
});

test("blank typed values stringify as empty string (no [object Object])", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Formula cells can legitimately have a blank value.
  sheet.setFormula(0, 0, "=IF(1=2,\"x\",\"\")", { value: { t: "blank" } });

  const matches = await findAll(wb, "*", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
    matchEntireCell: true,
  });

  assert.equal(matches.length, 1);
  assert.equal(matches[0].address, "Sheet1!A1");
  assert.equal(matches[0].text, "");
});

test("rich values (images, rich text) stringify as plain text (no [object Object])", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setValue(0, 0, { type: "image", value: { imageId: "img_1", altText: " Kitten " } });
  sheet.setValue(0, 1, { type: "image", value: { imageId: "img_2" } });
  sheet.setValue(0, 2, { text: "Hello", runs: [{ start: 0, end: 5, style: {} }] });

  // Default: search in values (display semantics). Rich values should match by their
  // stable text representation (alt text / placeholder / rich text).
  const kittenMatches = await findAll(wb, "Kitten", { scope: "sheet", currentSheetName: "Sheet1", matchEntireCell: true });
  assert.deepEqual(kittenMatches.map((m) => m.address), ["Sheet1!A1"]);
  assert.equal(kittenMatches[0].text, "Kitten");

  const placeholderMatches = await findAll(wb, "[Image]", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    matchEntireCell: true,
  });
  assert.deepEqual(placeholderMatches.map((m) => m.address), ["Sheet1!B1"]);

  const richTextMatches = await findAll(wb, "Hello", { scope: "sheet", currentSheetName: "Sheet1", matchEntireCell: true });
  assert.deepEqual(richTextMatches.map((m) => m.address), ["Sheet1!C1"]);

  // Look in formulas: constants should still stringify (Excel semantics). Ensure we don't
  // regress to `[object Object]` for rich values.
  const formulaLookMatches = await findAll(wb, "Kitten", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "formulas",
    matchEntireCell: true,
  });
  assert.deepEqual(formulaLookMatches.map((m) => m.address), ["Sheet1!A1"]);

  const objectMatches = await findAll(wb, "[object Object]", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    matchEntireCell: true,
  });
  assert.deepEqual(objectMatches.map((m) => m.address), []);
});

test("unknown typed values stringify stably (no [object Object])", () => {
  const text = formatCellValue({ t: "image", v: { imageId: "img_1" } });
  assert.notEqual(text, "[object Object]");
  assert.match(text, /\{\"t\":\"image\"/);
});

test("valueMode: display formats non-scalar display values (no [object Object])", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Some backends may store rich display values; ensure we still search using their text.
  sheet.setCell(0, 0, { value: 1, display: { text: "One", runs: [{ start: 0, end: 3, style: { bold: true } }] } });

  const matches = await findAll(wb, "One", { scope: "sheet", currentSheetName: "Sheet1", matchEntireCell: true });
  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!A1"]);
});
