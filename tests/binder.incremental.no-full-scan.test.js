import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

async function waitForCondition(condition, timeoutMs, intervalMs = 10) {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

test("binder: Yjs→DocumentController incremental updates avoid full cells-map scans", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");

  // Seed a large sheet (~10k cells) to ensure any per-update full scan will be noticeable.
  ydoc.transact(() => {
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        // Use the `${sheetId}:${row},${col}` encoding to ensure the binder
        // supports legacy key formats without extra scans.
        const key = `Sheet1:${row},${col}`;
        const cell = new Y.Map();
        // Intentionally leave cells "empty" (no value/formula) so hydration is fast,
        // while still creating a large `cells` map that would be expensive to rescan.
        cells.set(key, cell);
      }
    }
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  // The binder applies the initial hydration through an async apply chain (encryption/decryption can
  // be async). Wait for that initial apply to settle so the single-cell update below isn't queued
  // behind the initial ~10k-cell hydration under CPU contention.
  await binder.whenIdle();

  // After initial hydration, the binder must never iterate the full cells map on single-cell updates.
  const originalForEach = cells.forEach.bind(cells);
  cells.forEach = () => {
    throw new Error("cells.forEach should not be called after initial hydration");
  };

  try {
    // Apply a remote update to a single cell and ensure it propagates without scanning.
    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(() => {
      const cell = cells.get("Sheet1:0,0");
      assert.ok(cell instanceof Y.Map);
      cell.set("value", "updated");
    }, remoteOrigin);

    await waitForCondition(() => documentController.getCell("Sheet1", "A1").value === "updated", 10_000);
    assert.equal(documentController.getCell("Sheet1", "A1").formula, null);

    // Also support `r{row}c{col}` keys (resolved against defaultSheetId).
    ydoc.transact(() => {
      const cell = new Y.Map();
      cell.set("value", "rc-updated");
      cells.set("r100c0", cell);
    }, remoteOrigin);

    await waitForCondition(
      () => documentController.getCell("Sheet1", { row: 100, col: 0 }).value === "rc-updated",
      10_000
    );
  } finally {
    cells.forEach = originalForEach;
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: Yjs sheet formatRunsByCol array edits avoid full sheets-array scans", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  // Store format runs using a nested Y.Array encoding (some snapshots/clients may do this).
  ydoc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const runsByCol = new Y.Array();
    runsByCol.push([
      {
        col: 0,
        runs: [{ startRow: 0, endRowExclusive: 1, format: { font: { bold: true } } }],
      },
    ]);
    sheet.set("formatRunsByCol", runsByCol);

    sheets.push([sheet]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  // Wait for initial hydration to settle.
  await waitForCondition(() => Boolean(documentController.getCellFormat("Sheet1", "A1")?.font?.bold), 10_000);

  // After initial hydration, the binder should not scan all sheets when applying a single-sheet
  // nested array edit (Y.Array change where `changes.keys` is empty).
  const originalToArray = sheets.toArray.bind(sheets);
  sheets.toArray = () => {
    throw new Error("sheets.toArray should not be called after initial hydration");
  };

  try {
    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(() => {
      const sheet = sheets.get(0);
      assert.ok(sheet instanceof Y.Map);
      const runsByCol = sheet.get("formatRunsByCol");
      assert.ok(runsByCol instanceof Y.Array);
      runsByCol.push([
        {
          col: 1,
          runs: [{ startRow: 0, endRowExclusive: 1, format: { font: { italic: true } } }],
        },
      ]);
    }, remoteOrigin);

    await waitForCondition(() => Boolean(documentController.getCellFormat("Sheet1", "B1")?.font?.italic), 10_000);
  } finally {
    sheets.toArray = originalToArray;
    binder.destroy();
    ydoc.destroy();
  }
});
