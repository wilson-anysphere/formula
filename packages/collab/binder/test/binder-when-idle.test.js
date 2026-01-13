import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";

class StyleTableStub {
  get(_id) {
    return null;
  }
}

class DocumentControllerStub {
  constructor() {
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();
    /** @type {Set<(payload: any) => void>} */
    this._changeListeners = new Set();
    this.styleTable = new StyleTableStub();
  }

  /**
   * @param {"change"} event
   * @param {(payload: any) => void} cb
   */
  on(event, cb) {
    if (event !== "change") throw new Error(`Unsupported event: ${event}`);
    this._changeListeners.add(cb);
    return () => {
      this._changeListeners.delete(cb);
    };
  }

  /**
   * @param {{ deltas: any[] }} payload
   */
  _emitChange(payload) {
    for (const cb of this._changeListeners) {
      cb(payload);
    }
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} pos
   */
  _key(sheetId, pos) {
    return `${sheetId}:${pos.row}:${pos.col}`;
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} pos
   */
  getCell(sheetId, pos) {
    const key = this._key(sheetId, pos);
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} pos
   * @param {any} value
   */
  setCellValue(sheetId, pos, value) {
    const before = this.getCell(sheetId, pos);
    const after = { value: value ?? null, formula: null, styleId: before.styleId ?? 0 };
    this._cells.set(this._key(sheetId, pos), after);
    this._emitChange({
      deltas: [
        {
          sheetId,
          row: pos.row,
          col: pos.col,
          before,
          after,
        },
      ],
    });
  }
}

test("binder whenIdle waits for async encrypted DocumentController->Yjs writes", async (t) => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentControllerStub();

  const keyBytes = new Uint8Array(32);
  for (let i = 0; i < keyBytes.length; i += 1) keyBytes[i] = i & 0xff;

  const binding = bindYjsToDocumentController({
    ydoc,
    documentController,
    userId: "u1",
    encryption: {
      keyForCell: () => ({ keyId: "k1", keyBytes }),
    },
  });

  t.after(() => {
    binding.destroy();
    ydoc.destroy();
  });

  assert.equal(typeof binding.whenIdle, "function");

  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, 123);

  // `encryptCellPlaintext` is async; ensure `whenIdle` waits for it before resolving.
  await binding.whenIdle();

  const cells = ydoc.getMap("cells");
  const cell = cells.get("Sheet1:0:0");
  assert.ok(cell, "expected Yjs cell entry to exist after binder write");
  assert.equal(cell.has("enc"), true, "expected encrypted payload written to Yjs");
  // Encrypted cells must not leak plaintext value/formula keys.
  assert.equal(cell.has("value"), false);
  assert.equal(cell.has("formula"), false);
});

