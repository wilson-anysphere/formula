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
      get: (id) => (id === 1 ? { style: "ok" } : null),
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

  /**
   * @param {string} sheetId
   * @param {number} afterStyleId
   */
  setDefaultFormat(sheetId, afterStyleId) {
    this._emitter.emit("change", {
      formatDeltas: [{ sheetId, layer: "sheet", beforeStyleId: 0, afterStyleId }],
    });
  }

  /**
   * @param {string} sheetId
   */
  setRangeRuns(sheetId) {
    this._emitter.emit("change", {
      rangeRunDeltas: [
        {
          sheetId,
          col: 0,
          afterRuns: [{ startRow: 0, endRowExclusive: 1, styleId: 1 }],
        },
      ],
    });
  }
}

test("binder filters oversized drawing ids when upgrading plain-object sheets during sheet view writes", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedViewId = "x".repeat(5000);
  const oversizedTopId = "y".repeat(5000);

  doc.transact(() => {
    sheets.push([
      {
        id: "Sheet1",
        name: "Sheet1",
        view: {
          frozenRows: 0,
          frozenCols: 0,
          drawings: [{ id: oversizedViewId }, { id: "ok-view" }],
        },
        drawings: [{ id: oversizedTopId }, { id: "ok-top" }],
      },
    ]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);

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
      ["ok-view"],
    );

    const topDrawings = sheet.get("drawings");
    assert.ok(Array.isArray(topDrawings));
    assert.deepEqual(
      topDrawings.map((d) => d.id),
      ["ok-top"],
    );
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

test("binder filters oversized drawing ids when upgrading plain-object sheets during format writes", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedViewId = "x".repeat(5000);
  const oversizedTopId = "y".repeat(5000);

  doc.transact(() => {
    sheets.push([
      {
        id: "Sheet1",
        name: "Sheet1",
        view: {
          frozenRows: 0,
          frozenCols: 0,
          drawings: [{ id: oversizedViewId }, { id: "ok-view" }],
        },
        drawings: [{ id: oversizedTopId }, { id: "ok-top" }],
      },
    ]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);

    documentController.setDefaultFormat("Sheet1", 1);
    await flushAsync(5);

    const sheet = sheets.get(0);
    assert.ok(sheet instanceof Y.Map);
    assert.deepEqual(sheet.get("defaultFormat"), { style: "ok" });

    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");

    assert.ok(Array.isArray(view.drawings));
    assert.deepEqual(
      view.drawings.map((d) => d.id),
      ["ok-view"],
    );

    const topDrawings = sheet.get("drawings");
    assert.ok(Array.isArray(topDrawings));
    assert.deepEqual(
      topDrawings.map((d) => d.id),
      ["ok-top"],
    );
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

test("binder filters oversized drawing ids when upgrading plain-object sheets during range-run writes", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedViewId = "x".repeat(5000);
  const oversizedTopId = "y".repeat(5000);

  doc.transact(() => {
    sheets.push([
      {
        id: "Sheet1",
        name: "Sheet1",
        view: {
          frozenRows: 0,
          frozenCols: 0,
          drawings: [{ id: oversizedViewId }, { id: "ok-view" }],
        },
        drawings: [{ id: oversizedTopId }, { id: "ok-top" }],
      },
    ]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);

    documentController.setRangeRuns("Sheet1");
    await flushAsync(5);

    const sheet = sheets.get(0);
    assert.ok(sheet instanceof Y.Map);

    const view = sheet.get("view");
    assert.ok(view && typeof view === "object");
    assert.ok(Array.isArray(view.drawings));
    assert.deepEqual(
      view.drawings.map((d) => d.id),
      ["ok-view"],
    );

    const topDrawings = sheet.get("drawings");
    assert.ok(Array.isArray(topDrawings));
    assert.deepEqual(
      topDrawings.map((d) => d.id),
      ["ok-top"],
    );

    const formatRunsByCol = sheet.get("formatRunsByCol");
    assert.ok(formatRunsByCol instanceof Y.Map);
    assert.deepEqual(formatRunsByCol.get("0"), [{ startRow: 0, endRowExclusive: 1, format: { style: "ok" } }]);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});

test("binder does not materialize oversized drawing ids when reading sheet.view from Y.Maps", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Y.Map();
    view.set("frozenRows", 2);
    view.set("frozenCols", 1);

    const drawings = new Y.Array();
    const drawing = new Y.Map();
    const idText = new Y.Text();
    idText.insert(0, "x".repeat(5000));
    // If the binder calls `toString()` on this id (via yjsValueToJson(rawView)), this test should fail.
    idText.toString = () => {
      throw new Error("unexpected Y.Text.toString() on oversized drawing id");
    };
    drawing.set("id", idText);
    drawings.push([drawing]);
    view.set("drawings", drawings);

    sheet.set("view", view);
    sheets.push([sheet]);
  });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({ ydoc: doc, documentController, defaultSheetId: "Sheet1" });

  try {
    await flushAsync(5);
    const view = documentController.getSheetView("Sheet1");
    assert.equal(view.frozenRows, 2);
    assert.equal(view.frozenCols, 1);
  } finally {
    binder.destroy();
    doc.destroy();
  }
});
