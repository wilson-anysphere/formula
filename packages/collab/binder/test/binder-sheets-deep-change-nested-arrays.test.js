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
    /** @type {string[]} */
    this.getSheetViewCalls = [];
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
    this.getSheetViewCalls.push(sheetId);
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
}

test("binder does not full-scan all sheets when nested sheet arrays change (e.g. drawings)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    for (let i = 0; i < 3; i += 1) {
      const sheet = new Y.Map();
      sheet.set("id", `Sheet${i + 1}`);
      sheet.set("name", `Sheet${i + 1}`);

      const view = new Y.Map();
      view.set("frozenRows", 0);
      view.set("frozenCols", 0);
      sheet.set("view", view);

      if (i === 0) {
        // Store drawings metadata at the sheet root as a Y.Array (this is not synced by the binder,
        // but should also not cause a full scan of all sheet ids on every mutation).
        const drawings = new Y.Array();
        drawings.push([{ id: "drawing-0" }]);
        sheet.set("drawings", drawings);
      }

      sheets.push([sheet]);
    }
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);
    documentController.getSheetViewCalls.length = 0;

    // Mutate nested drawings array under Sheet1. This should only trigger sheet view
    // hydration for Sheet1 (not a full scan of all sheets).
    doc.transact(() => {
      const sheet = sheets.get(0);
      assert.ok(sheet instanceof Y.Map);
      const drawings = sheet.get("drawings");
      assert.ok(drawings instanceof Y.Array);
      drawings.push([{ id: "drawing-1" }]);
    });

    await flushAsync(5);

    const unique = new Set(documentController.getSheetViewCalls);
    assert.ok(unique.has("Sheet1"), "expected binder to process Sheet1");
    assert.deepEqual([...unique].sort(), ["Sheet1"]);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

