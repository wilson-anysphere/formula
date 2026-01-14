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

test("binder includes encryptionKeyId when rejecting edits to a nested Y.Map enc payload (unsupported schema)", async (t) => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");
  const cellMap = new Y.Map();
  const encMap = new Y.Map();
  encMap.set("v", 2);
  encMap.set("alg", "AES-256-GCM");
  encMap.set("keyId", "k-range-1");
  encMap.set("ivBase64", "AA==");
  encMap.set("tagBase64", "AA==");
  encMap.set("ciphertextBase64", "AA==");
  cellMap.set("enc", encMap);
  cells.set("Sheet1:0:0", cellMap);

  const documentController = new DocumentControllerStub();
  /** @type {any[] | null} */
  let rejected = null;
  const binding = bindYjsToDocumentController({
    ydoc,
    documentController,
    userId: "u1",
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });

  t.after(() => {
    binding.destroy();
    ydoc.destroy();
  });

  // Wait for any initial hydration to settle so our edit is treated as local.
  await binding.whenIdle?.();

  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");

  await binding.whenIdle?.();

  assert.ok(Array.isArray(rejected) && rejected.length > 0, "expected edit to be rejected");
  assert.equal(rejected[0].rejectionReason, "encryption");
  assert.equal(rejected[0].encryptionPayloadUnsupported, true);
  assert.equal(rejected[0].encryptionKeyId, "k-range-1");
});

