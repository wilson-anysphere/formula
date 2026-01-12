import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

async function waitForCondition(predicate, timeoutMs = 2_000, intervalMs = 5) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

test("binder: Yjs→DocumentController hydrates layered format defaults without creating undo history", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  // BranchService-style sheet entry: view contains formatting defaults.
  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    entry.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      defaultFormat: { font: { bold: true } },
      rowFormats: { "0": { font: { italic: true } } },
      colFormats: { "0": { fill: { fgColor: "red" } } },
    });
    sheets.push([entry]);
  });

  const documentController = new DocumentController();

  // Pre-intern the expected styles so we can assert style ids deterministically.
  const boldId = documentController.styleTable.intern({ font: { bold: true } });
  const italicId = documentController.styleTable.intern({ font: { italic: true } });
  const fillId = documentController.styleTable.intern({ fill: { fgColor: "red" } });

  const beforeDepth = documentController.getStackDepths();

  /** @type {any[]} */
  const changeEvents = [];
  const unsubscribe = documentController.on("change", (payload) => changeEvents.push(payload));

  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => {
      const format = documentController.getCellFormat("Sheet1", "A1");
      return (
        format.font?.bold === true &&
        format.font?.italic === true &&
        format.fill?.fgColor === "red"
      );
    });

    // Applying remote defaults should not create a local undo step.
    assert.deepEqual(documentController.getStackDepths(), beforeDepth);

    const formatEvent = changeEvents.find((evt) => Array.isArray(evt?.formatDeltas) && evt.formatDeltas.length > 0);
    assert.ok(formatEvent, "expected a change event that includes formatDeltas");
    assert.equal(formatEvent.source, "collab");
    assert.deepEqual(formatEvent.formatDeltas, [
      { sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId },
      { sheetId: "Sheet1", layer: "row", index: 0, beforeStyleId: 0, afterStyleId: italicId },
      { sheetId: "Sheet1", layer: "col", index: 0, beforeStyleId: 0, afterStyleId: fillId },
    ]);
  } finally {
    unsubscribe();
    binder.destroy();
    ydoc.destroy();
  }
});

