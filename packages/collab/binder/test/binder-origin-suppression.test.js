import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createUndoService } from "../../undo/index.js";
import { bindYjsToDocumentController } from "../index.js";

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

  /**
   * @param {string} event
   * @param {(payload: any) => void} listener
   */
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
   * @param {string} sheetId
   * @param {{ row: number, col: number } | string} coord
   */
  getCell(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const key = `${sheetId}:${c.row}:${c.col}`;
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }

  /**
   * @param {any[]} deltas
   * @param {{ source?: string, recalc?: boolean }} [options]
   */
  applyExternalDeltas(deltas, options = {}) {
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      const key = `${delta.sheetId}:${delta.row}:${delta.col}`;
      const next = {
        value: delta.after?.value ?? null,
        formula: delta.after?.formula ?? null,
        styleId: Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0,
      };
      this._cells.set(key, next);
    }
    this.#emit("change", { deltas, source: options.source, recalc: options.recalc });
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number } | string} coord
   * @param {any} value
   */
  setCellValue(sheetId, coord, value) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const before = this.getCell(sheetId, c);
    const after = { value: value ?? null, formula: null, styleId: before.styleId };
    const key = `${sheetId}:${c.row}:${c.col}`;
    this._cells.set(key, after);
    this.#emit("change", { deltas: [{ sheetId, row: c.row, col: c.col, before, after }] });
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number } | string} coord
   * @param {string | null} formula
   */
  setCellFormula(sheetId, coord, formula) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const before = this.getCell(sheetId, c);
    const after = { value: null, formula: formula ?? null, styleId: before.styleId };
    const key = `${sheetId}:${c.row}:${c.col}`;
    this._cells.set(key, after);
    this.#emit("change", { deltas: [{ sheetId, row: c.row, col: c.col, before, after }] });
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  #emit(event, payload) {
    const set = this._listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }
}

/**
 * Very small A1 parser (enough for binder tests).
 * @param {string} a1
 */
function parseA1(a1) {
  const match = /^([A-Z]+)(\d+)$/.exec(String(a1).toUpperCase());
  if (!match) throw new Error(`Invalid A1: ${a1}`);
  const [, colLetters, rowDigits] = match;
  let col = 0;
  for (const ch of colLetters) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  col -= 1;
  const row = Number(rowDigits) - 1;
  return { row, col };
}

test("binder applies session-origin local Yjs writes to DocumentController", async () => {
  const ydoc = new Y.Doc();
  const dc = new DocumentControllerStub();

  const sessionOrigin = { type: "test-session-origin" };
  const undoService = {
    transact: (fn) => ydoc.transact(fn, sessionOrigin),
    origin: sessionOrigin,
    localOrigins: new Set([sessionOrigin]),
  };

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: dc,
    undoService,
    defaultSheetId: "Sheet1",
  });

  const cells = ydoc.getMap("cells");
  ydoc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cells.set("Sheet1:0:0", cell);
  }, sessionOrigin);

  await waitForCondition(() => dc.getCell("Sheet1", { row: 0, col: 0 }).value === "x");
  assert.equal(dc.getCell("Sheet1", { row: 0, col: 0 }).value, "x");

  binder.destroy();
  ydoc.destroy();
});

test("binder propagates collaborative undo into DocumentController", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");
  const dc = new DocumentControllerStub();
  const sessionOrigin = { type: "test-session-origin" };

  const undoService = createUndoService({
    mode: "collab",
    doc: ydoc,
    scope: [cells],
    origin: sessionOrigin,
  });

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: dc,
    undoService,
    defaultSheetId: "Sheet1",
  });

  dc.setCellValue("Sheet1", { row: 0, col: 0 }, "x");

  await waitForCondition(() => {
    const cell = cells.get("Sheet1:0:0");
    return cell && typeof cell.get === "function" && cell.get("value") === "x";
  });

  undoService.undo();

  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", { row: 0, col: 0 });
    return cell.value == null && cell.formula == null;
  });

  const finalCell = dc.getCell("Sheet1", { row: 0, col: 0 });
  assert.equal(finalCell.value, null);
  assert.equal(finalCell.formula, null);

  binder.destroy();
  ydoc.destroy();
});
