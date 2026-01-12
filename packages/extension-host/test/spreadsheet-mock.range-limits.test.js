const test = require("node:test");
const assert = require("node:assert/strict");

const { InMemorySpreadsheet } = require("../src/spreadsheet-mock");

test("InMemorySpreadsheet: guards against huge range materialization", async () => {
  const sheet = new InMemorySpreadsheet();

  // Small selections should still materialize values normally.
  sheet.setCell(0, 0, 1);
  sheet.setCell(0, 1, 2);
  sheet.setCell(1, 0, 3);
  sheet.setCell(1, 1, 4);

  /** @type {any} */
  let lastSelectionEvent = null;
  sheet.onSelectionChanged((e) => {
    lastSelectionEvent = e;
  });

  sheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  assert.deepEqual(lastSelectionEvent?.selection?.values, [
    [1, 2],
    [3, 4]
  ]);

  // Huge selections should not allocate a full 2D values matrix.
  // 10,000 rows x 26 cols = 260,000 cells (> 200,000 cap).
  sheet.setSelection({ startRow: 0, startCol: 0, endRow: 9999, endCol: 25 });
  assert.deepEqual(lastSelectionEvent?.selection?.values, []);
  assert.equal(lastSelectionEvent?.selection?.truncated, true);

  // API reads/writes should fail fast with an explicit size error.
  assert.throws(() => sheet.getSelection(), /too large/i);
  assert.throws(() => sheet.getRange("A1:Z10000"), /too large/i);
  assert.throws(() => sheet.setRange("A1:Z10000", []), /too large/i);
});

