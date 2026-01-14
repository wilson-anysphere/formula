import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";
import { encryptCellPlaintext } from "../../encryption/src/index.node.js";

async function waitForCondition(fn, timeoutMs = 2000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const ok = await fn();
      if (ok) return;
    } catch {
      // Ignore transient errors while waiting for async state to settle.
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

class StyleTableStub {
  constructor() {
    /** @type {Map<string, number>} */
    this._idsByKey = new Map();
    /** @type {Map<number, any>} */
    this._formatsById = new Map();
    this._nextId = 1;
  }

  /**
   * @param {any} format
   */
  intern(format) {
    const key = JSON.stringify(format);
    const existing = this._idsByKey.get(key);
    if (existing) return existing;
    const id = this._nextId++;
    this._idsByKey.set(key, id);
    this._formatsById.set(id, format);
    return id;
  }

  /**
   * @param {number} id
   */
  get(id) {
    return this._formatsById.get(id) ?? null;
  }
}

class DocumentControllerStub {
  constructor() {
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this._listeners = new Map();
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();
    this.styleTable = new StyleTableStub();
    this.canEditCell = null;
  }

  on(event, listener) {
    let set = this._listeners.get(event);
    if (!set) {
      set = new Set();
      this._listeners.set(event, set);
    }
    set.add(listener);
    return () => set.delete(listener);
  }

  /**
   * @param {any[]} deltas
   */
  applyExternalDeltas(deltas) {
    for (const delta of deltas ?? []) {
      const key = `${delta.sheetId}:${delta.row}:${delta.col}`;
      const after = delta.after ?? {};
      this._cells.set(key, {
        value: after.value ?? null,
        formula: after.formula ?? null,
        styleId: Number.isInteger(after.styleId) ? after.styleId : 0,
      });
    }
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} coord
   */
  getCell(sheetId, coord) {
    const key = `${sheetId}:${coord.row}:${coord.col}`;
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }
}

test("binder reads legacy per-cell `style` key as a `format` alias", async (t) => {
  const ydoc = new Y.Doc();
  const dc = new DocumentControllerStub();
  const binder = bindYjsToDocumentController({ ydoc, documentController: dc });

  t.after(() => {
    binder.destroy();
    ydoc.destroy();
  });

  const cells = ydoc.getMap("cells");
  ydoc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    // Legacy key: some clients wrote per-cell formatting under `style`.
    cell.set("style", { font: { bold: true } });
    cells.set("Sheet1:0:0", cell);
  });

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    const format = dc.styleTable.get(cell.styleId);
    return cell.value === "x" && cell.styleId !== 0 && format?.font?.bold === true;
  });

  // Ensure updates to `style` are observed too.
  ydoc.transact(() => {
    const cell = /** @type {any} */ (cells.get("Sheet1:0:0"));
    cell.set("style", { font: { italic: true } });
  });

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    const format = dc.styleTable.get(cell.styleId);
    return cell.value === "x" && cell.styleId !== 0 && format?.font?.italic === true;
  });
});

test("binder preserves legacy `style` across duplicate key encodings even when the canonical cell is encrypted", async (t) => {
  const ydoc = new Y.Doc({ guid: "binder-style-alias-encrypted-doc" });
  const dc = new DocumentControllerStub();
  const binder = bindYjsToDocumentController({ ydoc, documentController: dc });

  t.after(() => {
    binder.destroy();
    ydoc.destroy();
  });

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null },
    key: { keyId: "k1", keyBytes: new Uint8Array(32).fill(7) },
    context: { docId: ydoc.guid, sheetId: "Sheet1", row: 0, col: 0 },
  });

  const cells = ydoc.getMap("cells");
  ydoc.transact(() => {
    const canonical = new Y.Map();
    canonical.set("enc", enc);
    // No plaintext format on the canonical cell key.
    cells.set("Sheet1:0:0", canonical);

    const legacy = new Y.Map();
    legacy.set("style", { font: { bold: true } });
    // Legacy key encoding for the same cell coordinate.
    cells.set("Sheet1:0,0", legacy);
  });

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    const format = dc.styleTable.get(cell.styleId);
    return cell.value === "###" && cell.styleId !== 0 && format?.font?.bold === true;
  });
});

test("binder prefers ciphertext over an enc=null marker across duplicate key encodings", async (t) => {
  const ydoc = new Y.Doc({ guid: "binder-encryption-null-marker-precedence-doc" });
  const dc = new DocumentControllerStub();

  const keyBytes = new Uint8Array(32).fill(7);
  const keyId = "k1";

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: dc,
    encryption: {
      keyForCell: (cell) => {
        if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
          return { keyId, keyBytes };
        }
        return null;
      },
    },
  });

  t.after(() => {
    binder.destroy();
    ydoc.destroy();
  });

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null },
    key: { keyId, keyBytes },
    context: { docId: ydoc.guid, sheetId: "Sheet1", row: 0, col: 0 },
  });

  const cells = ydoc.getMap("cells");
  ydoc.transact(() => {
    const canonical = new Y.Map();
    canonical.set("enc", null);
    cells.set("Sheet1:0:0", canonical);

    const legacy = new Y.Map();
    legacy.set("enc", enc);
    cells.set("Sheet1:0,0", legacy);
  });

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    return cell.value === "top-secret" && cell.formula == null;
  });
});

test("binder rehydrate applies decrypted value + decrypted format when an encryption key becomes available (encryptFormat)", async (t) => {
  const ydoc = new Y.Doc({ guid: "binder-rehydrate-encryptFormat-doc" });
  const dc = new DocumentControllerStub();

  const keyBytes = new Uint8Array(32).fill(7);
  const keyId = "k1";
  let hasKey = false;

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: dc,
    encryption: {
      encryptFormat: true,
      keyForCell: (cell) => {
        if (!hasKey) return null;
        if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
          return { keyId, keyBytes };
        }
        return null;
      },
    },
  });

  t.after(() => {
    binder.destroy();
    ydoc.destroy();
  });

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null, format: { font: { bold: true } } },
    key: { keyId, keyBytes },
    context: { docId: ydoc.guid, sheetId: "Sheet1", row: 0, col: 0 },
  });

  const cells = ydoc.getMap("cells");
  ydoc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    cells.set("Sheet1:0:0", cell);
  });

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    return cell.value === "###" && cell.formula == null && cell.styleId === 0;
  });

  // Simulate importing an encryption key locally: keyForCell starts returning the key,
  // but there is no Yjs mutation. Binder must rehydrate to reveal the cell.
  hasKey = true;
  binder.rehydrate();

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    const format = dc.styleTable.get(cell.styleId);
    return cell.value === "top-secret" && cell.formula == null && cell.styleId !== 0 && format?.font?.bold === true;
  });
});
