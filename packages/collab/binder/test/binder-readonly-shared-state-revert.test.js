import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { bindYjsToDocumentController } from "../index.js";

class DocumentControllerStub {
  constructor() {
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this._listeners = new Map();

    /** @type {Map<string, { frozenRows: number, frozenCols: number }>} */
    this._sheetViews = new Map();

    /** @type {Map<string, number>} */
    this._sheetStyleIds = new Map();

    /** @type {Map<string, Map<number, any[]>>} */
    this._rangeRunsBySheet = new Map();

    // The binder patches this when permission guards are provided; keep it present.
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
   * @param {string} event
   * @param {any} payload
   */
  _emit(event, payload) {
    const set = this._listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }

  /**
   * @param {string} sheetId
   */
  getSheetView(sheetId) {
    return this._sheetViews.get(sheetId) ?? { frozenRows: 0, frozenCols: 0 };
  }

  /**
   * @param {string} sheetId
   * @param {number} frozenRows
   * @param {number} frozenCols
   */
  setFrozen(sheetId, frozenRows, frozenCols) {
    const before = this.getSheetView(sheetId);
    const after = { frozenRows, frozenCols };
    this._sheetViews.set(sheetId, after);
    this._emit("change", { deltas: [], sheetViewDeltas: [{ sheetId, before, after }] });
  }

  /**
   * @param {any[]} deltas
   * @param {{ source?: string }} [options]
   */
  applyExternalSheetViewDeltas(deltas, options = {}) {
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      this._sheetViews.set(delta.sheetId, delta.after);
    }
    this._emit("change", { deltas: [], sheetViewDeltas: deltas, source: options.source });
  }

  /**
   * @param {string} sheetId
   */
  getSheetDefaultStyleId(sheetId) {
    return this._sheetStyleIds.get(sheetId) ?? 0;
  }

  /**
   * @param {string} sheetId
   * @param {number} styleId
   */
  setSheetDefaultStyleId(sheetId, styleId) {
    const beforeStyleId = this.getSheetDefaultStyleId(sheetId);
    const afterStyleId = styleId;
    this._sheetStyleIds.set(sheetId, afterStyleId);
    this._emit("change", {
      deltas: [],
      formatDeltas: [{ sheetId, layer: "sheet", beforeStyleId, afterStyleId }],
    });
  }

  /**
   * @param {any[]} deltas
   * @param {{ source?: string }} [options]
   */
  applyExternalFormatDeltas(deltas, options = {}) {
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      if (delta.layer !== "sheet") continue;
      this._sheetStyleIds.set(delta.sheetId, delta.afterStyleId);
    }
    this._emit("change", { deltas: [], formatDeltas: deltas, source: options.source });
  }

  /**
   * @param {string} sheetId
   * @param {number} col
   */
  getRangeRuns(sheetId, col) {
    const sheet = this._rangeRunsBySheet.get(sheetId);
    if (!sheet) return [];
    return sheet.get(col) ?? [];
  }

  /**
   * @param {string} sheetId
   * @param {number} col
   * @param {any[]} runs
   */
  setRangeRuns(sheetId, col, runs) {
    let sheet = this._rangeRunsBySheet.get(sheetId);
    if (!sheet) {
      sheet = new Map();
      this._rangeRunsBySheet.set(sheetId, sheet);
    }
    const beforeRuns = sheet.get(col) ?? [];
    const afterRuns = Array.isArray(runs) ? runs : [];
    sheet.set(col, afterRuns);
    this._emit("change", {
      deltas: [],
      rangeRunDeltas: [{ sheetId, col, startRow: 0, endRowExclusive: 10, beforeRuns, afterRuns }],
    });
  }

  /**
   * @param {any[]} deltas
   * @param {{ source?: string }} [options]
   */
  applyExternalRangeRunDeltas(deltas, options = {}) {
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      let sheet = this._rangeRunsBySheet.get(delta.sheetId);
      if (!sheet) {
        sheet = new Map();
        this._rangeRunsBySheet.set(delta.sheetId, sheet);
      }
      sheet.set(delta.col, Array.isArray(delta.afterRuns) ? delta.afterRuns : []);
    }
    this._emit("change", { deltas: [], rangeRunDeltas: deltas, source: options.source });
  }
}

test("binder reverts local sheet view / format / range-run mutations when canWriteSharedState is false", async () => {
  const ydoc = new Y.Doc();
  const dc = new DocumentControllerStub();

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController: dc,
    defaultSheetId: "Sheet1",
    canWriteSharedState: () => false,
  });

  // Sheet view (freeze panes) should snap back immediately.
  assert.deepEqual(dc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });
  dc.setFrozen("Sheet1", 1, 0);
  assert.deepEqual(dc.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });

  // Layered formatting should also be reverted (prevent local-only format changes).
  assert.equal(dc.getSheetDefaultStyleId("Sheet1"), 0);
  dc.setSheetDefaultStyleId("Sheet1", 123);
  assert.equal(dc.getSheetDefaultStyleId("Sheet1"), 0);

  // Range-run formatting deltas should be reverted.
  assert.deepEqual(dc.getRangeRuns("Sheet1", 0), []);
  dc.setRangeRuns("Sheet1", 0, [{ startRow: 0, endRowExclusive: 10, styleId: 1 }]);
  assert.deepEqual(dc.getRangeRuns("Sheet1", 0), []);

  binder.destroy();
  ydoc.destroy();
});

