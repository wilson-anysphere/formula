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
   * @param {any[]} deltas
   */
  applyExternalSheetViewDeltas(deltas) {
    for (const delta of deltas) {
      this._sheetViews.set(delta.sheetId, delta.after);
    }
  }

  /**
   * Simulate a user sheet view edit (e.g. drawings metadata update).
   *
   * @param {string} sheetId
   * @param {any} after
   */
  setSheetView(sheetId, after) {
    const before = this.getSheetView(sheetId);
    this._sheetViews.set(sheetId, after);
    this._emitter.emit("change", { sheetViewDeltas: [{ sheetId, before, after }] });
  }
}

test("binder ignores sheet view deltas that only touch unknown keys", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("view", { frozenRows: 0, frozenCols: 0, drawings: [{ id: "drawing-1" }] });
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  // Seed controller with drawings in view state (these are not synced by the binder).
  documentController._sheetViews.set("Sheet1", { frozenRows: 0, frozenCols: 0, drawings: [{ id: "drawing-1" }] });

  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    // Count Yjs updates produced by DocumentController -> Yjs writes.
    let updateCount = 0;
    doc.on("update", () => {
      updateCount += 1;
    });

    // Allow initial hydration to settle (it should not produce writes).
    await flushAsync(5);
    updateCount = 0;

    // Update drawings only (unknown key for the binder). This should not trigger any
    // Yjs writes from the binder.
    documentController.setSheetView("Sheet1", { frozenRows: 0, frozenCols: 0, drawings: [{ id: "drawing-2" }] });
    await flushAsync(5);

    assert.equal(updateCount, 0);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

