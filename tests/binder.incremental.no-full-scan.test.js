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
