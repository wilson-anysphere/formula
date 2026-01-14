import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import * as Y from "yjs";

import { createCollabUndoService, REMOTE_ORIGIN } from "../packages/collab/undo/index.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import { getWorkbookRoots } from "../packages/collab/workbook/src/index.ts";
import { bindSheetViewToCollabSession } from "../apps/desktop/src/collab/sheetViewBinder.ts";

async function flushAsync(times = 3) {
  for (let i = 0; i < times; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
  }
}

class DocumentControllerStub {
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
   * @param {{ source?: string } | undefined} options
   */
  applyExternalSheetViewDeltas(deltas, options = {}) {
    for (const delta of deltas) {
      this._sheetViews.set(delta.sheetId, delta.after);
    }
    // Mirror DocumentController's change payload shape.
    this._emitter.emit("change", { sheetViewDeltas: deltas, source: options.source });
  }
}

test("remote sheet view updates applied via sheetViewBinder are not echoed back into Yjs by the full binder (prevents undo pollution)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const binderOrigin = { type: "test:binder-origin" };
  const undoService = createCollabUndoService({ doc, scope: [sheets], origin: binderOrigin });

  const dc = new DocumentControllerStub();
  const sheetViewBinder = bindSheetViewToCollabSession({
    session: /** @type {any} */ ({ doc, sheets, localOrigins: new Set(), isReadOnly: () => false }),
    documentController: /** @type {any} */ (dc),
    origin: binderOrigin,
  });
  const fullBinder = bindYjsToDocumentController({
    ydoc: doc,
    documentController: dc,
    undoService,
    defaultSheetId: "Sheet1",
  });

  let localUpdates = 0;
  const onUpdate = (_update, origin) => {
    if (origin === binderOrigin) localUpdates += 1;
  };

  try {
    // Give any binder hydration work a chance to settle.
    await flushAsync(5);

    // Start with a clean undo stack and zero local-origin update count.
    undoService.undoManager.clear();
    doc.on("update", onUpdate);

    // Simulate a remote collaborator overwriting the sheet view state.
    doc.transact(() => {
      const sheet = sheets.get(0);
      assert.ok(sheet, "expected Sheet1 entry in Yjs");
      sheet.set("view", { frozenRows: 2, frozenCols: 1 });
    }, REMOTE_ORIGIN);

    // Allow any (incorrect) echo writes to run.
    await flushAsync(10);

    assert.deepEqual(dc.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });
    assert.equal(localUpdates, 0, "expected no binder-origin Yjs updates from remote sheet view changes");
    assert.equal(undoService.canUndo(), false, "expected remote sheet view changes to not create undoable local steps");
  } finally {
    doc.off("update", onUpdate);
    sheetViewBinder.destroy();
    fullBinder.destroy();
    undoService.undoManager.destroy();
    doc.destroy();
  }
});

test("remote drawings updates applied via sheetViewBinder are not echoed back into Yjs by the full binder (prevents undo pollution)", async () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);
  });

  const binderOrigin = { type: "test:binder-origin" };
  const undoService = createCollabUndoService({ doc, scope: [sheets], origin: binderOrigin });

  const dc = new DocumentControllerStub();
  const sheetViewBinder = bindSheetViewToCollabSession({
    session: /** @type {any} */ ({ doc, sheets, localOrigins: new Set(), isReadOnly: () => false }),
    documentController: /** @type {any} */ (dc),
    origin: binderOrigin,
  });
  const fullBinder = bindYjsToDocumentController({
    ydoc: doc,
    documentController: dc,
    undoService,
    defaultSheetId: "Sheet1",
  });

  let localUpdates = 0;
  const onUpdate = (_update, origin) => {
    if (origin === binderOrigin) localUpdates += 1;
  };

  try {
    await flushAsync(5);

    undoService.undoManager.clear();
    doc.on("update", onUpdate);

    const drawings = [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ];

    doc.transact(() => {
      const sheet = sheets.get(0);
      assert.ok(sheet, "expected Sheet1 entry in Yjs");
      sheet.set("view", { frozenRows: 0, frozenCols: 0, drawings });
    }, REMOTE_ORIGIN);

    await flushAsync(10);

    assert.deepEqual(dc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, drawings });
    assert.equal(localUpdates, 0, "expected no binder-origin Yjs updates from remote drawings changes");
    assert.equal(undoService.canUndo(), false, "expected remote drawings changes to not create undoable local steps");
  } finally {
    doc.off("update", onUpdate);
    sheetViewBinder.destroy();
    fullBinder.destroy();
    undoService.undoManager.destroy();
    doc.destroy();
  }
});
