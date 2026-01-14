import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";
import { encryptCellPlaintext } from "../../encryption/src/index.node.js";

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

test("binder allows edits when canonical key has enc=null marker but ciphertext exists under a legacy key", async (t) => {
  const keyBytes = new Uint8Array(32).fill(7);
  const keyId = "k1";

  const ydoc = new Y.Doc({ guid: "binder-encryption-null-marker-precedence-write-doc" });
  const cells = ydoc.getMap("cells");

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null },
    key: { keyId, keyBytes },
    context: { docId: ydoc.guid, sheetId: "Sheet1", row: 0, col: 0 },
  });

  ydoc.transact(() => {
    const marker = new Y.Map();
    marker.set("enc", null);
    cells.set("Sheet1:0:0", marker);

    const payload = new Y.Map();
    payload.set("enc", enc);
    cells.set("Sheet1:0,0", payload);
  });

  const documentController = new DocumentControllerStub();
  /** @type {any[] | null} */
  let rejected = null;
  const binding = bindYjsToDocumentController({
    ydoc,
    documentController,
    userId: "u1",
    encryption: {
      keyForCell: (cell) => {
        if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
          return { keyId, keyBytes };
        }
        return null;
      },
    },
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });

  t.after(() => {
    binding.destroy();
    ydoc.destroy();
  });

  await binding.whenIdle?.();

  // Local edit should be allowed (encrypted write), not rejected due to the marker.
  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "updated");
  await binding.whenIdle?.();

  assert.ok(!rejected || rejected.length === 0, "expected edit not to be rejected");

  const canonical = cells.get("Sheet1:0:0");
  assert.ok(canonical instanceof Y.Map);
  const canonicalEnc = canonical.get("enc");
  assert.ok(canonicalEnc && typeof canonicalEnc === "object", "expected ciphertext to be written to canonical key");
  assert.equal(canonicalEnc.keyId, keyId);
});

