import test from "node:test";
import assert from "node:assert/strict";

import { SearchSession, WorkbookSearchIndex } from "../index.js";

function immediate() {
  return new Promise((resolve) => setImmediate(resolve));
}

class SyntheticSheet {
  constructor(name, { rows, cols, needlePositions }) {
    this.name = name;
    this._rows = rows;
    this._cols = cols;
    this._needle = needlePositions;

    this._defaultCell = { value: "x" };
    this._needleCell = { value: "needle" };
  }

  getUsedRange() {
    return { startRow: 0, endRow: this._rows - 1, startCol: 0, endCol: this._cols - 1 };
  }

  getCell(row, col) {
    const key = `${row},${col}`;
    return this._needle.has(key) ? this._needleCell : this._defaultCell;
  }

  *iterateCells(range, { order = "byRows" } = {}) {
    // Avoid allocating 1M objects – yield a mutable record.
    const entry = { row: 0, col: 0, cell: this._defaultCell };

    if (order === "byColumns") {
      for (let col = range.startCol; col <= range.endCol; col++) {
        for (let row = range.startRow; row <= range.endRow; row++) {
          entry.row = row;
          entry.col = col;
          entry.cell = this.getCell(row, col);
          yield entry;
        }
      }
      return;
    }

    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        entry.row = row;
        entry.col = col;
        entry.cell = this.getCell(row, col);
        yield entry;
      }
    }
  }
}

class SyntheticWorkbook {
  constructor(sheet) {
    this.sheets = [sheet];
  }

  getSheet(name) {
    const sheet = this.sheets.find((s) => s.name === name);
    if (!sheet) throw new Error(`Unknown sheet: ${name}`);
    return sheet;
  }
}

test("perf: repeated findNext becomes sub-linear once the workbook index is built", async () => {
  const needlePositions = new Set(["999,999"]);
  const sheet = new SyntheticSheet("Sheet1", { rows: 1000, cols: 1000, needlePositions });
  const wb = new SyntheticWorkbook(sheet);

  const index = new WorkbookSearchIndex(wb, { autoThresholdCells: 0 });

  const s1 = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
    // Give the index builder a bit of time per slice so the test doesn't
    // spend too much time context-switching.
    timeBudgetMs: 5,
    checkEvery: 2048,
  });
  const m1 = await s1.findNext();
  assert.equal(m1.address, "Sheet1!ALL1000"); // (999,999)

  const s2 = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  const m2 = await s2.findNext();
  assert.equal(m2.address, "Sheet1!ALL1000");

  // Index build work should have been amortized into the first session.
  assert.equal(s2.stats.indexCellsVisited, 0);
  // Candidates should be tiny (O(matches) not O(sheet size)).
  assert.ok(s2.stats.cellsScanned <= 5);
});

test("perf: AbortSignal cancels long-running index builds quickly", async () => {
  const needlePositions = new Set(["999,999"]);
  const sheet = new SyntheticSheet("Sheet1", { rows: 1000, cols: 1000, needlePositions });
  const wb = new SyntheticWorkbook(sheet);

  const index = new WorkbookSearchIndex(wb, { autoThresholdCells: 0 });
  const controller = new AbortController();

  let yields = 0;
  const scheduler = async () => {
    yields++;
    // Under heavy CI load, the first timeslice checkpoint can occur before the first cell
    // is processed. Abort on the *second* yield so we deterministically observe partial
    // progress before cancellation.
    if (yields === 2) controller.abort();
    await immediate();
  };

  const session = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
    timeBudgetMs: 0.5,
    checkEvery: 1,
    scheduler,
  });

  await assert.rejects(session.findNext({ signal: controller.signal }), (err) => err?.name === "AbortError");

  const idx = index.getSheetModeIndex("Sheet1", "values:display");
  assert.ok(idx);
  assert.ok(idx.stats.cellsVisited > 0);
  assert.ok(idx.stats.cellsVisited < 1_000_000);
});

test("perf: AbortSignal cancels long-running scans quickly", async () => {
  const needlePositions = new Set(["999,999"]);
  const sheet = new SyntheticSheet("Sheet1", { rows: 1000, cols: 1000, needlePositions });
  const wb = new SyntheticWorkbook(sheet);

  const controller = new AbortController();
  let yields = 0;
  const scheduler = async () => {
    yields++;
    if (yields === 1) controller.abort();
    await immediate();
  };

  const session = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    // Force scan path (no index).
    timeBudgetMs: 0.5,
    checkEvery: 1,
    scheduler,
  });

  await assert.rejects(session.findNext({ signal: controller.signal }), (err) => err?.name === "AbortError");

  assert.ok(session.stats.cellsScanned > 0);
  assert.ok(session.stats.cellsScanned < 1_000_000);
});
