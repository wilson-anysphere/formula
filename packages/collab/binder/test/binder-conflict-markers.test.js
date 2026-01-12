import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";
import { getWorkbookRoots } from "../../workbook/src/index.ts";

class StyleTableStub {
  constructor() {
    /** @type {Map<number, any>} */
    this._formatsById = new Map();
    /** @type {Map<string, number>} */
    this._idsByKey = new Map();
    this._nextId = 1;
  }

  /**
   * @param {number} id
   */
  get(id) {
    return this._formatsById.get(id) ?? null;
  }

  /**
   * @param {any} format
   */
  intern(format) {
    const key = JSON.stringify(format);
    const existing = this._idsByKey.get(key);
    if (existing !== undefined) return existing;
    const id = this._nextId++;
    this._idsByKey.set(key, id);
    this._formatsById.set(id, format);
    return id;
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
   * @param {any[]} deltas
   */
  applyExternalDeltas(deltas) {
    for (const delta of deltas ?? []) {
      const key = this._key(delta.sheetId, { row: delta.row, col: delta.col });
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
   * @param {{ row: number, col: number }} pos
   * @param {any} value
   */
  setCellValue(sheetId, pos, value) {
    const key = this._key(sheetId, pos);
    const before = this.getCell(sheetId, pos);
    const after = { value: value ?? null, formula: null, styleId: before.styleId ?? 0 };
    this._cells.set(key, after);
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

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} pos
   * @param {string | null} formula
   */
  setCellFormula(sheetId, pos, formula) {
    const key = this._key(sheetId, pos);
    const before = this.getCell(sheetId, pos);
    const normalizedFormula = formula ?? null;
    const after = {
      value: null,
      formula: normalizedFormula,
      styleId: before.styleId ?? 0,
    };
    this._cells.set(key, after);
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

function nextTick() {
  return new Promise((resolve) => setImmediate(resolve));
}

test("binder preserves formula null markers and retains cell maps on clears", async (t) => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentControllerStub();
  const binding = bindYjsToDocumentController({
    ydoc,
    documentController,
    userId: "u",
    // Enable explicit `formula=null` markers + marker-preserving empty-cell behavior
    // so downstream causal conflict detection (FormulaConflictMonitor) can detect
    // delete-vs-overwrite and formula-vs-value conflicts reliably.
    formulaConflictsMode: "formula+value",
  });

  t.after(() => {
    binding.destroy();
    ydoc.destroy();
  });

  const cells = getWorkbookRoots(ydoc).cells;

  // 1) Set a formula and ensure it is written into Yjs.
  documentController.setCellFormula("Sheet1", { row: 0, col: 0 }, "=1");
  await nextTick();

  const yCell = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(yCell, "expected Yjs cell map to exist after setting a formula");
  assert.equal(yCell.get("formula"), "=1");

  // 2) Clear the formula. Binder should keep the cell entry and write `formula=null`
  //    rather than deleting the key.
  documentController.setCellFormula("Sheet1", { row: 0, col: 0 }, null);
  await nextTick();

  const yCellCleared = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(yCellCleared, "expected Yjs cell map to still exist after clearing a formula");
  assert.equal(yCellCleared.has("formula"), true);
  assert.equal(yCellCleared.get("formula"), null);

  // 3) Literal value writes and clears should also preserve a `formula=null` marker.
  documentController.setCellValue("Sheet1", { row: 0, col: 1 }, 123);
  await nextTick();

  const yValueCell = /** @type {any} */ (cells.get("Sheet1:0:1"));
  assert.ok(yValueCell, "expected Yjs cell map to exist after setting a value");
  assert.equal(yValueCell.get("value"), 123);
  assert.equal(yValueCell.has("formula"), true);
  assert.equal(yValueCell.get("formula"), null);

  documentController.setCellValue("Sheet1", { row: 0, col: 1 }, null);
  await nextTick();

  const yValueCellCleared = /** @type {any} */ (cells.get("Sheet1:0:1"));
  assert.ok(yValueCellCleared, "expected Yjs cell map to still exist after clearing a value");
  assert.equal(yValueCellCleared.has("formula"), true);
  assert.equal(yValueCellCleared.get("formula"), null);
});
