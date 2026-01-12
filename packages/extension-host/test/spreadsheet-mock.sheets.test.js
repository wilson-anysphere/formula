const test = require("node:test");
const assert = require("node:assert/strict");

const { InMemorySpreadsheet } = require("../src/spreadsheet-mock");

test("InMemorySpreadsheet: createSheet inserts after the active sheet by default", () => {
  const sheet = new InMemorySpreadsheet();

  // Create a second sheet (active=Sheet1). With only one sheet this matches append.
  sheet.createSheet("Data");
  assert.deepEqual(sheet.listSheets().map((s) => s.name), ["Sheet1", "Data"]);

  // Activate the first sheet again and create another sheet. It should be inserted
  // immediately after the active sheet (before the previously-created sheet).
  sheet.activateSheet("Sheet1");
  sheet.createSheet("Second");
  assert.deepEqual(sheet.listSheets().map((s) => s.name), ["Sheet1", "Second", "Data"]);
  assert.equal(sheet.getActiveSheet().name, "Second");
});

test("InMemorySpreadsheet: sheet names are case-insensitive + validated", () => {
  const sheet = new InMemorySpreadsheet();

  assert.ok(sheet.getSheet("sheet1"));

  assert.throws(() => sheet.createSheet("sheet1"), /sheet name already exists/i);
  assert.throws(() => sheet.createSheet("Bad:Name"), /sheet name contains invalid character `:`/i);
  assert.throws(() => sheet.createSheet("'Budget"), /sheet name cannot begin or end with an apostrophe/i);
  assert.throws(() => sheet.createSheet("A".repeat(32)), /sheet name cannot exceed 31 characters/i);
});

