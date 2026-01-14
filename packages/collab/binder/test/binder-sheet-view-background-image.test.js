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
   * Simulate a user sheet view edit.
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

test("binder syncs sheet view backgroundImageId between Yjs and DocumentController", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    view.set("backgroundImageId", "bg1.png");
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  // Initial hydration should apply backgroundImageId.
  assert.equal(documentController.getSheetView("Sheet1").backgroundImageId, "bg1.png");

  // Remote update (Yjs -> DocumentController).
  doc.transact(() => {
    const sheet = sheets.get(0);
    assert.ok(sheet);
    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    view.set("backgroundImageId", "bg2.png");
  });
  await flushAsync(5);
  assert.equal(documentController.getSheetView("Sheet1").backgroundImageId, "bg2.png");

  // Remote clear.
  doc.transact(() => {
    const sheet = sheets.get(0);
    assert.ok(sheet);
    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    view.delete("backgroundImageId");
  });
  await flushAsync(5);
  assert.equal(documentController.getSheetView("Sheet1").backgroundImageId, undefined);

  // Local update (DocumentController -> Yjs).
  documentController.setSheetView("Sheet1", { frozenRows: 0, frozenCols: 0, backgroundImageId: "bg3.png" });
  await flushAsync(5);
  {
    const sheet = sheets.get(0);
    assert.ok(sheet);
    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    assert.equal(view.backgroundImageId, "bg3.png");
  }

  // Local clear should remove the key from the Yjs view object.
  documentController.setSheetView("Sheet1", { frozenRows: 0, frozenCols: 0 });
  await flushAsync(5);
  {
    const sheet = sheets.get(0);
    assert.ok(sheet);
    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    assert.equal(view.backgroundImageId, undefined);
  }

  binder.destroy();
  doc.destroy();
});

