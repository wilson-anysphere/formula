import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, SearchSession, WorkbookSearchIndex } from "../index.js";

test("SearchSession: findNext/findPrev are stateful and wrap like Excel", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setValue(0, 0, "foo"); // A1
  sheet.setValue(0, 1, "foo"); // B1
  sheet.setValue(1, 0, "foo"); // A2

  const session = new SearchSession(wb, "foo", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    searchOrder: "byRows",
    wrap: true,
  });

  const m1 = await session.findNext();
  assert.equal(m1.address, "Sheet1!A1");
  assert.equal(m1.wrapped, false);

  const m2 = await session.findNext();
  assert.equal(m2.address, "Sheet1!B1");

  const m3 = await session.findNext();
  assert.equal(m3.address, "Sheet1!A2");

  const m4 = await session.findNext();
  assert.equal(m4.address, "Sheet1!A1");
  assert.equal(m4.wrapped, true);

  const p1 = await session.findPrev();
  assert.equal(p1.address, "Sheet1!A2");
  assert.equal(p1.wrapped, true);

  const p2 = await session.findPrev();
  assert.equal(p2.address, "Sheet1!B1");
});

test("SearchSession: replaceNext updates cached matches and advances the cursor", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setValue(0, 0, "foo"); // A1
  sheet.setValue(0, 1, "foo"); // B1

  const session = new SearchSession(wb, "foo", { scope: "sheet", currentSheetName: "Sheet1" });

  const r1 = await session.replaceNext("bar");
  assert.equal(r1.replaced, true);
  assert.equal(r1.match.address, "Sheet1!A1");
  assert.equal(sheet.getCell(0, 0).value, "bar");

  const next = await session.findNext();
  assert.equal(next.address, "Sheet1!B1");
});

test("WorkbookSearchIndex: can be shared across sessions for sub-linear repeated queries", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Sparse-but-large: only a few cells are interesting, but used range is big.
  // This is representative of real workbooks where formatting/structured data
  // can expand the used range.
  sheet.setValue(999, 999, "needle");
  sheet.setValue(100, 100, "needle");

  const index = new WorkbookSearchIndex(wb, { autoThresholdCells: 0 });

  const s1 = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  const m1 = await s1.findNext();
  assert.equal(m1.address, "Sheet1!CW101"); // (100,100) => CW101

  const s2 = new SearchSession(wb, "needle", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  const m2 = await s2.findNext();
  assert.equal(m2.address, "Sheet1!CW101");
  assert.equal(s2.stats.indexCellsVisited, 0);
  assert.ok(s2.stats.cellsScanned <= 5);
});

test("WorkbookSearchIndex: incremental update after replaceNext keeps the index fresh", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setValue(0, 0, "foo");

  const index = new WorkbookSearchIndex(wb, { autoThresholdCells: 0 });

  const s1 = new SearchSession(wb, "foo", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  await s1.replaceNext("baz");
  assert.equal(sheet.getCell(0, 0).value, "baz");

  const s2 = new SearchSession(wb, "baz", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  const m = await s2.findNext();
  assert.equal(m.address, "Sheet1!A1");
  assert.equal(s2.stats.indexCellsVisited, 0);
});

test("SearchSession: uses options.signal when no signal is passed to methods", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setValue(0, 0, "foo");

  const controller = new AbortController();
  controller.abort();

  const session = new SearchSession(wb, "foo", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    signal: controller.signal,
  });

  await assert.rejects(session.findNext(), (err) => err?.name === "AbortError");
});
