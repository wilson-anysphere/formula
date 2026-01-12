import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import { createUndoService } from "../packages/collab/undo/index.js";

async function waitForCondition(condition, timeoutMs = 10_000, intervalMs = 10) {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

test("binder: large-rectangle format syncs via range runs without per-cell materialization", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");
  const sheets = ydoc.getArray("sheets");

  // Seed a default Sheet1 entry so per-sheet formatting metadata has a place to live.
  ydoc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const docA = new DocumentController();
  const docB = new DocumentController();

  const binderA = bindYjsToDocumentController({ ydoc, documentController: docA, defaultSheetId: "Sheet1", userId: "u-a" });
  const binderB = bindYjsToDocumentController({ ydoc, documentController: docB, defaultSheetId: "Sheet1", userId: "u-b" });

  try {
    // Apply a very large rectangle that is NOT a full row/col/sheet selection so it uses
    // the compressed per-column range-run formatting layer.
    //
    // 26 cols * 1,000,000 rows = 26,000,000 cells -> should exceed the run threshold.
    docA.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

    await waitForCondition(() => Boolean(docB.getCellFormat("Sheet1", "A5000")?.font?.bold));

    assert.equal(docB.getCellFormat("Sheet1", "A5000")?.font?.bold, true);

    // Ensure the binder stored the compressed formatting in `sheets[*].formatRunsByCol`
    // (not legacy `view.formatRunsByCol`) and that it stores *formats* (style objects),
    // not local style ids.
    const sheetEntry = sheets.get(0);
    assert.ok(sheetEntry instanceof Y.Map);
    const formatRunsByCol = sheetEntry.get("formatRunsByCol");
    assert.ok(formatRunsByCol instanceof Y.Map, "expected Sheet1.formatRunsByCol to be a Y.Map");

    const col0Runs = formatRunsByCol.get("0");
    assert.ok(Array.isArray(col0Runs) && col0Runs.length > 0, "expected serialized runs for col 0");
    assert.equal(col0Runs[0]?.styleId, undefined);
    assert.equal(col0Runs[0]?.format?.font?.bold, true);

    const col25Runs = formatRunsByCol.get("25");
    assert.ok(Array.isArray(col25Runs) && col25Runs.length > 0, "expected serialized runs for col 25");
    assert.equal(col25Runs[0]?.styleId, undefined);
    assert.equal(col25Runs[0]?.format?.font?.bold, true);

    // No per-cell materialization should occur for the range formatting (it lives on sheet metadata).
    assert.equal(cells.size, 0);
    assert.equal(docB.model?.sheets?.get?.("Sheet1")?.cells?.size ?? 0, 0);
  } finally {
    binderA.destroy();
    binderB.destroy();
    ydoc.destroy();
  }
});

test("binder: top-level formatRunsByCol overrides legacy view.formatRunsByCol (prevents stale fallback)", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  // Seed a sheet entry with legacy `view.formatRunsByCol` only.
  ydoc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      formatRunsByCol: {
        "0": [{ startRow: 0, endRowExclusive: 2, format: { font: { bold: true } } }],
      },
    });
    sheets.push([sheet]);
  });

  const docA = new DocumentController();
  const binderA = bindYjsToDocumentController({ ydoc, documentController: docA, defaultSheetId: "Sheet1", userId: "u-a" });

  try {
    await waitForCondition(() => Boolean(docA.getCellFormat("Sheet1", "A1")?.font?.bold));
    assert.equal(docA.getCellFormat("Sheet1", "A1")?.font?.bold, true);

    // Clear the formatting via range-run deltas. This triggers DocumentController→Yjs writes
    // that should create an *empty* top-level `formatRunsByCol` map, preventing future clients
    // from falling back to the stale legacy `view.formatRunsByCol` formatting.
    const beforeRuns = docA.model.sheets.get("Sheet1")?.formatRunsByCol?.get?.(0) ?? [];
    assert.ok(beforeRuns.length > 0, "expected Sheet1 col 0 to have hydrated legacy runs");

    docA.applyExternalRangeRunDeltas([
      { sheetId: "Sheet1", col: 0, startRow: 0, endRowExclusive: 2, beforeRuns, afterRuns: [] },
    ]);

    await waitForCondition(() => {
      const sheetEntry = sheets.get(0);
      if (!(sheetEntry instanceof Y.Map)) return false;
      const map = sheetEntry.get("formatRunsByCol");
      return map instanceof Y.Map && map.size === 0;
    });

    const sheetEntry = sheets.get(0);
    assert.ok(sheetEntry instanceof Y.Map);
    const formatRunsByCol = sheetEntry.get("formatRunsByCol");
    assert.ok(formatRunsByCol instanceof Y.Map);
    assert.equal(formatRunsByCol.size, 0);

    const view = sheetEntry.get("view");
    assert.equal(view?.formatRunsByCol?.["0"]?.[0]?.format?.font?.bold, true);

    // A freshly-bound client should ignore legacy `view.formatRunsByCol` once the top-level
    // key exists (even if empty).
    const docB = new DocumentController();
    const binderB = bindYjsToDocumentController({ ydoc, documentController: docB, defaultSheetId: "Sheet1", userId: "u-b" });
    try {
      assert.equal(docB.getCellFormat("Sheet1", "A1")?.font?.bold, undefined);
    } finally {
      binderB.destroy();
    }
  } finally {
    binderA.destroy();
    ydoc.destroy();
  }
});

test("binder: collab undo/redo reverts range-run formatting changes (formatRunsByCol in undo scope)", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");
  const sheets = ydoc.getArray("sheets");

  ydoc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
    userId: "u-a",
  });

  try {
    // A1:Z2000 is just over the range-run threshold (26 * 2000 = 52,000).
    documentController.setRangeFormat("Sheet1", "A1:Z2000", { font: { bold: true } });

    await waitForCondition(() => Boolean(documentController.getCellFormat("Sheet1", "A1")?.font?.bold));
    assert.equal(documentController.getCellFormat("Sheet1", "A1")?.font?.bold, true);
    assert.equal(documentController.getCellFormat("Sheet1", "Z2000")?.font?.bold, true);

    // Ensure the write landed in Yjs so the UndoManager has something to undo.
    await waitForCondition(() => {
      const sheetEntry = sheets.get(0);
      if (!(sheetEntry instanceof Y.Map)) return false;
      const runs = sheetEntry.get("formatRunsByCol");
      if (!(runs instanceof Y.Map)) return false;
      const col0 = runs.get("0");
      return Array.isArray(col0) && col0.length > 0 && col0[0]?.format?.font?.bold === true;
    });

    undo.stopCapturing();
    assert.equal(undo.canUndo(), true);

    undo.undo();
    await waitForCondition(() => !documentController.getCellFormat("Sheet1", "A1")?.font?.bold);
    assert.equal(documentController.getCellFormat("Sheet1", "A1")?.font?.bold, undefined);
    assert.equal(documentController.getCellFormat("Sheet1", "Z2000")?.font?.bold, undefined);

    assert.equal(undo.canRedo(), true);
    undo.redo();
    await waitForCondition(() => Boolean(documentController.getCellFormat("Sheet1", "A1")?.font?.bold));
    assert.equal(documentController.getCellFormat("Sheet1", "A1")?.font?.bold, true);
    assert.equal(documentController.getCellFormat("Sheet1", "Z2000")?.font?.bold, true);

    // No per-cell materialization should occur; range formatting lives on sheet metadata.
    assert.equal(cells.size, 0);
    assert.equal(documentController.model?.sheets?.get?.("Sheet1")?.cells?.size ?? 0, 0);
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});
