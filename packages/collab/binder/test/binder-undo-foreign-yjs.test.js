import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { EventEmitter } from "node:events";

import * as Y from "yjs";

import { createUndoService } from "../../undo/index.js";

import { bindYjsToDocumentController } from "../index.js";
import { decryptCellPlaintext, encryptCellPlaintext, isEncryptedCellPayload } from "../../encryption/src/index.node.js";
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

async function waitFor(predicate, { timeoutMs = 2000 } = {}) {
  const start = Date.now();
  while (true) {
    let ok = false;
    try {
      ok = await predicate();
    } catch {
      ok = false;
    }
    if (ok) return;
    if (Date.now() - start > timeoutMs) {
      throw new Error("Timed out waiting for condition");
    }
    await new Promise((resolve) => setImmediate(resolve));
  }
}

class TestDocumentController {
  constructor() {
    this._emitter = new EventEmitter();
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();
    this.externalDeltaCount = 0;
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
    this.externalDeltaCount += 1;
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

  /**
   * Simulate a user formula edit.
   *
   * @param {string} sheetId
   * @param {{ row: number, col: number }} address
   * @param {string | null} formula
   */
  setCellFormula(sheetId, address, formula) {
    const before = this.getCell(sheetId, address);
    const after = { value: null, formula, styleId: before.styleId };
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

test("binder normalizes foreign nested cell maps before delete+undo so collab undo works", async () => {
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
  assert.equal(documentController.externalDeltaCount, 1, "initial hydration should apply exactly one external delta batch");

  // Clear the cell (null write). With default binder semantics (conflict mode off),
  // this deletes the cell entry. We still expect normalization to occur *before*
  // the delete so undo restores a local Y.Map (not a foreign one).
  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, null);
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 1, "binder should not echo local edits back into DocumentController");

  assert.equal(cellsRoot.get("Sheet1:0:0"), undefined, "empty cell should be deleted when conflict semantics are off");
  assert.equal(cellsRoot.get("Sheet1:0,0"), undefined, "legacy key cell should be deleted when conflict semantics are off");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 2, "undo should apply exactly one external delta batch");

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

test("binder normalizes foreign nested cell maps when only legacy keys exist (delete+undo)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const legacy = new Ycjs.Map();
    legacy.set("value", "from-cjs");
    legacy.set("formula", null);
    legacy.set("modified", 1);
    remoteCells.set("Sheet1:0,0", legacy);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const cellsRoot = getWorkbookRoots(doc).cells;
  Ycjs.applyUpdate(doc, update);

  const beforeWriteLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(beforeWriteLegacy);
  assert.equal(beforeWriteLegacy instanceof Y.Map, false);
  assert.equal(cellsRoot.get("Sheet1:0:0"), undefined);

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
  assert.equal(documentController.externalDeltaCount, 1);

  // Clear via canonical coordinates. Binder targets the legacy raw key. With
  // conflict semantics off, this deletes the legacy entry; undo should restore
  // it as a local Y.Map.
  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, null);
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 1);

  assert.equal(cellsRoot.get("Sheet1:0,0"), undefined);

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 2);

  const afterUndoLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(afterUndoLegacy);
  assert.ok(afterUndoLegacy instanceof Y.Map);
  assert.equal(afterUndoLegacy.get("value"), "from-cjs");
  assert.equal(undoService.canUndo(), false);
  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).value, "from-cjs");

  binder.destroy();
  doc.destroy();
});

test("binder normalizes foreign nested cell maps when only r{row}c{col} keys exist (delete+undo)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const legacy = new Ycjs.Map();
    legacy.set("value", "from-cjs");
    legacy.set("formula", null);
    legacy.set("modified", 1);
    remoteCells.set("r0c0", legacy);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const cellsRoot = getWorkbookRoots(doc).cells;
  Ycjs.applyUpdate(doc, update);

  const beforeWriteLegacy = cellsRoot.get("r0c0");
  assert.ok(beforeWriteLegacy);
  assert.equal(beforeWriteLegacy instanceof Y.Map, false);
  assert.equal(cellsRoot.get("Sheet1:0:0"), undefined);

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
  assert.equal(documentController.externalDeltaCount, 1);

  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, null);
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 1);

  assert.equal(cellsRoot.get("r0c0"), undefined);

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 2);

  const afterUndoLegacy = cellsRoot.get("r0c0");
  assert.ok(afterUndoLegacy);
  assert.ok(afterUndoLegacy instanceof Y.Map);
  assert.equal(afterUndoLegacy.get("value"), "from-cjs");
  assert.equal(undoService.canUndo(), false);
  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).value, "from-cjs");

  binder.destroy();
  doc.destroy();
});

test("binder normalizes foreign nested cell maps for formula edits so collab undo works", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const canonical = new Ycjs.Map();
    canonical.set("value", "from-cjs");
    canonical.set("formula", null);
    canonical.set("modified", 1);
    remoteCells.set("Sheet1:0:0", canonical);

    const legacy = new Ycjs.Map();
    legacy.set("value", "from-cjs");
    legacy.set("formula", null);
    legacy.set("modified", 1);
    remoteCells.set("Sheet1:0,0", legacy);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
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
  assert.equal(documentController.externalDeltaCount, 1);

  documentController.setCellFormula("Sheet1", { row: 0, col: 0 }, "=1+1");
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 1);

  const afterWrite = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterWrite);
  assert.ok(afterWrite instanceof Y.Map);
  assert.equal(afterWrite.get("formula"), "=1+1");
  assert.equal(afterWrite.get("value"), null);

  const afterWriteLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(afterWriteLegacy);
  assert.ok(afterWriteLegacy instanceof Y.Map);
  assert.equal(afterWriteLegacy.get("formula"), "=1+1");
  assert.equal(afterWriteLegacy.get("value"), null);

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await flushAsync();
  assert.equal(documentController.externalDeltaCount, 2);

  const afterUndo = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterUndo);
  assert.ok(afterUndo instanceof Y.Map);
  assert.equal(afterUndo.get("value"), "from-cjs");
  assert.equal(afterUndo.get("formula"), null);

  const afterUndoLegacy = cellsRoot.get("Sheet1:0,0");
  assert.ok(afterUndoLegacy);
  assert.ok(afterUndoLegacy instanceof Y.Map);
  assert.equal(afterUndoLegacy.get("value"), "from-cjs");
  assert.equal(afterUndoLegacy.get("formula"), null);

  assert.equal(undoService.canUndo(), false);
  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).value, "from-cjs");
  assert.equal(documentController.getCell("Sheet1", { row: 0, col: 0 }).formula, null);

  binder.destroy();
  doc.destroy();
});

test("binder normalizes foreign nested cell maps for encrypted edits so collab undo works", async () => {
  const Ycjs = requireYjsCjs();

  const guid = `doc-${Date.now()}`;
  const key = { keyId: "key-1", keyBytes: new Uint8Array(32).fill(7) };

  const encryptedPayload = await encryptCellPlaintext({
    plaintext: { value: "from-cjs", formula: null },
    key,
    context: { docId: guid, sheetId: "Sheet1", row: 0, col: 0 },
  });

  const remote = new Ycjs.Doc({ guid });
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const canonical = new Ycjs.Map();
    canonical.set("enc", encryptedPayload);
    canonical.set("modified", 1);
    remoteCells.set("Sheet1:0:0", canonical);

    const legacy = new Ycjs.Map();
    legacy.set("enc", encryptedPayload);
    legacy.set("modified", 1);
    remoteCells.set("Sheet1:0,0", legacy);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc({ guid });
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
    encryption: {
      keyForCell: () => key,
      shouldEncryptCell: () => true,
    },
  });

  await waitFor(() => documentController.getCell("Sheet1", { row: 0, col: 0 }).value === "from-cjs");
  assert.equal(documentController.externalDeltaCount, 1);

  documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "edited");
  await waitFor(async () => {
    const cell = cellsRoot.get("Sheet1:0:0");
    if (!(cell instanceof Y.Map)) return false;
    const enc = cell.get("enc");
    if (!isEncryptedCellPayload(enc)) return false;
    const plaintext = await decryptCellPlaintext({
      encrypted: enc,
      key,
      context: { docId: guid, sheetId: "Sheet1", row: 0, col: 0 },
    });
    return plaintext?.value === "edited";
  });

  const afterWrite = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterWrite);
  assert.ok(afterWrite instanceof Y.Map);
  const afterEnc = afterWrite.get("enc");
  assert.ok(isEncryptedCellPayload(afterEnc));
  const decryptedAfter = await decryptCellPlaintext({
    encrypted: afterEnc,
    key,
    context: { docId: guid, sheetId: "Sheet1", row: 0, col: 0 },
  });
  assert.equal(decryptedAfter?.value, "edited");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  await waitFor(() => documentController.getCell("Sheet1", { row: 0, col: 0 }).value === "from-cjs");
  assert.equal(documentController.externalDeltaCount, 2);

  const afterUndo = cellsRoot.get("Sheet1:0:0");
  assert.ok(afterUndo);
  assert.ok(afterUndo instanceof Y.Map);
  const afterUndoEnc = afterUndo.get("enc");
  assert.ok(isEncryptedCellPayload(afterUndoEnc));
  const decryptedAfterUndo = await decryptCellPlaintext({
    encrypted: afterUndoEnc,
    key,
    context: { docId: guid, sheetId: "Sheet1", row: 0, col: 0 },
  });
  assert.equal(decryptedAfterUndo?.value, "from-cjs");

  assert.equal(undoService.canUndo(), false);

  binder.destroy();
  doc.destroy();
});
