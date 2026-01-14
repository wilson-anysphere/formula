import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";
import { getWorkbookRoots } from "../../workbook/src/index.ts";

async function flushAsync(times = 3) {
  for (let i = 0; i < times; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
  }
}

class TestDocumentController {
  constructor() {
    this._emitter = new EventEmitter();
    /** @type {Map<string, any>} */
    this._sheetViews = new Map();
    this.styleTable = {
      intern: () => 0,
      get: () => null,
    };
  }

  /**
   * @param {"change"} event
   * @param {(payload: any) => void} cb
   */
  on(event, cb) {
    this._emitter.on(event, cb);
    return () => this._emitter.off(event, cb);
  }

  /**
   * @param {string} sheetId
   */
  getSheetView(sheetId) {
    return this._sheetViews.get(sheetId) ?? { frozenRows: 0, frozenCols: 0 };
  }

  /**
   * External sheet-view updates (e.g. from collaboration) should not be written back into Yjs.
   *
   * @param {any[]} deltas
   * @param {{ source?: string }} [options]
   */
  applyExternalSheetViewDeltas(deltas, options = {}) {
    for (const delta of deltas) {
      this._sheetViews.set(delta.sheetId, delta.after);
    }
    this._emitter.emit("change", { sheetViewDeltas: deltas, source: options.source });
  }
}

test("binder ignores DocumentController change events with source=collab (prevents echo from other binders)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    let updateCount = 0;
    doc.on("update", () => {
      updateCount += 1;
    });

    await flushAsync(5);
    updateCount = 0;

    const before = documentController.getSheetView("Sheet1");
    const after = { ...before, frozenRows: 2 };
    documentController.applyExternalSheetViewDeltas([{ sheetId: "Sheet1", before, after }], { source: "collab" });

    await flushAsync(5);

    assert.equal(updateCount, 0);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

test("binder ignores DocumentController change events with source=applyState (prevents snapshot restore from overwriting Yjs)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    let updateCount = 0;
    doc.on("update", () => {
      updateCount += 1;
    });

    await flushAsync(5);
    updateCount = 0;

    const before = documentController.getSheetView("Sheet1");
    const after = { ...before, frozenRows: 2 };
    documentController.applyExternalSheetViewDeltas([{ sheetId: "Sheet1", before, after }], { source: "applyState" });

    await flushAsync(5);

    assert.equal(updateCount, 0);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});
