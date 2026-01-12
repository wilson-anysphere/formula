import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { EventEmitter } from "node:events";

import * as Y from "yjs";

import { createUndoService } from "@formula/collab-undo";

import { bindYjsToDocumentController } from "../index.js";
import { getWorkbookRoots } from "../../workbook/src/index.ts";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

async function flushAsync(times = 3) {
  for (let i = 0; i < times; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
  }
}

class TestDocumentController {
  constructor() {
    this._emitter = new EventEmitter();
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();
    this.styleTable = {
      intern: (_format) => 1,
      get: (_id) => null,
    };
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} address
   */
  _key(sheetId, address) {
    return `${sheetId}:${address.row}:${address.col}`;
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
   * @param {{ row: number, col: number }} address
   */
  getCell(sheetId, address) {
    const key = this._key(sheetId, address);
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }

  /**
   * @param {any[]} deltas
   */
  applyExternalDeltas(deltas) {
    for (const delta of deltas) {
      const key = this._key(delta.sheetId, { row: delta.row, col: delta.col });
      this._cells.set(key, {
        value: delta.after?.value ?? null,
        formula: delta.after?.formula ?? null,
        styleId: Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0,
      });
    }
  }

  /**
   * Simulate a user edit. Updates local state and emits a DocumentController change
   * event so the binder writes into Yjs.
   *
   * @param {string} sheetId
   * @param {{ row: number, col: number }} address
   * @param {any} value
   */
  setCellValue(sheetId, address, value) {
    const before = this.getCell(sheetId, address);
    const after = { value, formula: null, styleId: before.styleId };
    const key = this._key(sheetId, address);
    this._cells.set(key, after);
    this._emitter.emit("change", {
      deltas: [
        {
          sheetId,
          row: address.row,
          col: address.col,
          before,
          after,
        },
      ],
    });
  }
}

test("binder normalizes foreign nested cell maps before mutating so collab undo works", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const canonical = new Ycjs.Map();
    canonical.set("value", "from-cjs");
    canonical.set("formula", null);
    canonical.set("modified", 1);
    remoteCells.set("Sheet1:0:0", canonical);

    // Also store the same cell under a legacy encoding so the binder writes to
    // multiple raw keys for a single canonical coordinate.
    const legacy = new Ycjs.Map();
    legacy.set("value", "from-cjs");
    legacy.set("formula", null);
    legacy.set("modified", 1);
    remoteCells.set("Sheet1:0,0", legacy);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Ensure the root exists in the ESM Yjs instance so the update only introduces
  // foreign nested cell maps (not a foreign `cells` root).
  const cellsRoot = getWorkbookRoots(doc).cells;
  Ycjs.applyUpdate(doc, update);

  const beforeWrite = cellsRoot.get("Sheet1:0:0");
  assert.ok(beforeWrite);
  assert.equal(beforeWrite instanceof Y.Map, false);

  const localOrigin = { type: "local:test" };
  const undoService = createUndoService({ mode: "collab", doc, scope: [cellsRoot], origin: localOrigin });

  const documentController = new TestDocumentController();
  const binder = bindYjsToDocumentController({
    ydoc: doc,
    documentController,
    undoService,
    defaultSheetId: "Sheet1",
  });

  await flushAsync();

  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).value, "from-cjs");

  // Clear the cell (null write). This exercises the case where the binder is about to
  // mutate a foreign nested Y.Map even though the resulting value is empty.
  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, null);
  await flushAsync();

  const afterWrite = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterWrite);
  assert.ok(afterWrite instanceof Y.Map, "cell map should be normalized to local Y.Map");
  assert.equal(afterWrite.get("value"), null);

  const afterWriteLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(afterWriteLegacy);
  assert.ok(afterWriteLegacy instanceof Y.Map, "legacy key cell map should also be normalized to local Y.Map");
  assert.equal(afterWriteLegacy.get("value"), null);

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await flushAsync();

  const afterUndo = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterUndo);
  assert.ok(afterUndo instanceof Y.Map, "undo should not revert normalization to a foreign Y.Map");
  assert.equal(afterUndo.get("value"), "from-cjs");

  const afterUndoLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(afterUndoLegacy);
  assert.ok(afterUndoLegacy instanceof Y.Map, "undo should not revert legacy normalization to a foreign Y.Map");
  assert.equal(afterUndoLegacy.get("value"), "from-cjs");

  assert.equal(undoService.canUndo(), false, "normalization should not create an extra undo step");
  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).value, "from-cjs");

  binder.destroy();
  doc.destroy();
});
