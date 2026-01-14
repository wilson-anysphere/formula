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
   * @param {string} sheetId
   * @param {any} after
   */
  setSheetView(sheetId, after) {
    const before = this.getSheetView(sheetId);
    this._sheetViews.set(sheetId, after);
    this._emitter.emit("change", { sheetViewDeltas: [{ sheetId, before, after }] });
  }
}

test("binder drops oversized drawing ids when preserving unknown sheet view keys during view writes", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversized = "x".repeat(5000);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    // Store view as a plain object so the binder's view-write path has to materialize it
    // (and preserve unknown keys like drawings).
    sheet.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      drawings: [{ id: oversized }, { id: "ok" }],
    });
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);

    // Trigger a view update that the binder will persist into Yjs.
    documentController.setSheetView("Sheet1", { frozenRows: 1, frozenCols: 0 });
    await flushAsync(5);

    const sheet = sheets.get(0);
    assert.ok(sheet instanceof Y.Map);

    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    assert.equal(view.frozenRows, 1);
    assert.equal(view.frozenCols, 0);
    assert.ok(Array.isArray(view.drawings));
    assert.deepEqual(
      view.drawings.map((d) => d.id),
      ["ok"],
    );
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

test("binder does not materialize oversized Y.Text drawing ids when preserving unknown sheet view keys during view writes", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedText = new Y.Text();
  oversizedText.insert(0, "x".repeat(5000));
  oversizedText.toString = () => {
    throw new Error("unexpected Y.Text.toString() on oversized drawing id");
  };

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const bad = new Y.Map();
    bad.set("id", oversizedText);

    // Store view as a plain object so the binder's view-write path has to materialize it
    // (and preserve unknown keys like drawings).
    sheet.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      drawings: [bad, { id: "ok" }],
    });
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);

    // Trigger a view update that the binder will persist into Yjs.
    documentController.setSheetView("Sheet1", { frozenRows: 1, frozenCols: 0 });
    await flushAsync(5);

    const sheet = sheets.get(0);
    assert.ok(sheet instanceof Y.Map);

    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    assert.equal(view.frozenRows, 1);
    assert.equal(view.frozenCols, 0);
    assert.ok(Array.isArray(view.drawings));
    assert.deepEqual(
      view.drawings.map((d) => d.id),
      ["ok"],
    );
  } finally {
    binder.destroy();
    doc.destroy();
  }
});
