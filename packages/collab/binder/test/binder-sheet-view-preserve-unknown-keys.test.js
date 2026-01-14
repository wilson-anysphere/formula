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
}

test("binder preserves unknown sheet view keys when applying Yjs -> DocumentController updates", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const mergedRanges = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }];

  // Seed the controller with extra view metadata that the binder does not sync.
  documentController._sheetViews.set("Sheet1", { frozenRows: 1, frozenCols: 0, mergedRanges });

  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  // Initial hydration should update frozen rows while preserving `mergedRanges`.
  assert.equal(documentController.getSheetView("Sheet1").frozenRows, 0);
  assert.deepEqual(documentController.getSheetView("Sheet1").mergedRanges, mergedRanges);

  // Subsequent remote updates should also preserve `mergedRanges`.
  doc.transact(() => {
    const sheet = sheets.get(0);
    assert.ok(sheet);
    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    view.set("frozenRows", 2);
  });

  await flushAsync(5);

  assert.equal(documentController.getSheetView("Sheet1").frozenRows, 2);
  assert.deepEqual(documentController.getSheetView("Sheet1").mergedRanges, mergedRanges);

  binder.destroy();
  doc.destroy();
});

