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
   * Simulate a user sheet view edit (e.g. freezing panes).
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

test("binder preserves unknown sheet.view keys when the existing view is a renamed Y.Map", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    // Store sheet.view as a Y.Map (BranchService-style snapshot encoding), with an
    // extra custom key that should be preserved when applying view deltas.
    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    view.set("customKey", { foo: "bar" });

    // Simulate a bundler-renamed constructor without mutating global module state.
    class RenamedMap extends view.constructor {}
    Object.setPrototypeOf(view, RenamedMap.prototype);

    sheet.set("view", view);
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  await flushAsync();

  // Apply a local view delta that should preserve `customKey` from the existing view.
  documentController.setSheetView("Sheet1", { frozenRows: 1, frozenCols: 0 });
  await flushAsync(5);

  const sheet = sheets.get(0);
  assert.ok(sheet);
  const view = sheet.get("view");
  assert.ok(view && typeof view === "object");
  assert.equal(typeof view.get, "function", "expected sheet.view to remain a Y.Map");
  assert.equal(view.get("frozenRows"), 1);
  assert.equal(view.get("frozenCols"), 0);
  assert.deepEqual(view.get("customKey"), { foo: "bar" });

  binder.destroy();
  doc.destroy();
});
