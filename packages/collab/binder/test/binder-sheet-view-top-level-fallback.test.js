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

test("binder hydrates legacy top-level sheet view fields even when sheet.view exists (unknown keys only)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    // A `view` map exists but contains only metadata this binder does not sync.
    const view = new Y.Map();
    view.set("drawings", [{ id: "drawing-1" }]);
    sheet.set("view", view);

    // Legacy top-level view fields (should still hydrate).
    sheet.set("frozenRows", 2);
    sheet.set("frozenCols", 1);
    sheet.set("background_image_id", "bg.png");
    sheet.set("colWidths", { "0": 120 });
    sheet.set("rowHeights", { "1": 40 });

    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });
  await flushAsync(5);

  assert.deepEqual(documentController.getSheetView("Sheet1"), {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });

  binder.destroy();
  doc.destroy();
});

test("binder hydrates legacy top-level view fields when sheet.view is partially migrated", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    // Canonical view object exists but is missing some keys (partial migration).
    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    sheet.set("view", view);

    // Legacy top-level fields for the missing keys.
    sheet.set("background_image", "bg2.png");
    sheet.set("colWidths", { "0": 123 });
    sheet.set("rowHeights", { "2": 33 });

    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });
  await flushAsync(5);

  assert.deepEqual(documentController.getSheetView("Sheet1"), {
    frozenRows: 0,
    frozenCols: 0,
    backgroundImageId: "bg2.png",
    colWidths: { "0": 123 },
    rowHeights: { "2": 33 },
  });

  binder.destroy();
  doc.destroy();
});

test("binder deletes legacy top-level sheet view keys when writing view updates (prevents stale fallback)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    // Canonical view map.
    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    sheet.set("view", view);

    // Seed legacy top-level view fields that should be removed once the binder writes.
    sheet.set("frozenRows", 2);
    sheet.set("frozenCols", 1);
    sheet.set("background_image_id", "stale-bg.png");
    sheet.set("backgroundImage", "stale-bg-2.png");
    sheet.set("background_image", "stale-bg-3.png");
    sheet.set("colWidths", { "0": 120 });
    sheet.set("rowHeights", { "1": 40 });

    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });
  await flushAsync(5);

  // Local edit triggers DocumentController -> Yjs write.
  documentController.setSheetView("Sheet1", {
    frozenRows: 3,
    frozenCols: 4,
    backgroundImageId: "bg3.png",
    colWidths: { "0": 100 },
    rowHeights: { "1": 20 },
  });
  await flushAsync(5);

  const sheet = sheets.get(0);
  assert.ok(sheet);
  const view = sheet.get("view");
  assert.ok(view && typeof view === "object");
  assert.equal(view.get("frozenRows"), 3);
  assert.equal(view.get("frozenCols"), 4);
  assert.equal(view.get("backgroundImageId"), "bg3.png");
  assert.deepEqual(view.get("colWidths"), { "0": 100 });
  assert.deepEqual(view.get("rowHeights"), { "1": 20 });

  // Top-level legacy keys should be removed.
  assert.equal(sheet.get("frozenRows"), undefined);
  assert.equal(sheet.get("frozenCols"), undefined);
  assert.equal(sheet.get("backgroundImageId"), undefined);
  assert.equal(sheet.get("background_image_id"), undefined);
  assert.equal(sheet.get("backgroundImage"), undefined);
  assert.equal(sheet.get("background_image"), undefined);
  assert.equal(sheet.get("colWidths"), undefined);
  assert.equal(sheet.get("rowHeights"), undefined);

  binder.destroy();
  doc.destroy();
});

