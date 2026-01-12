import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { bindCollabSessionToDocumentController, createCollabSession } from "../src/index.ts";

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

async function waitForCondition(fn, timeoutMs = 2000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      if (await fn()) return;
    } catch {
      // ignore while waiting
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

function makeCellKey(sheetId, row, col) {
  return `${sheetId}:${row}:${col}`;
}

class DocumentControllerStub {
  constructor() {
    this._events = new EventEmitter();
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();
    /** @type {Map<string, any>} */
    this._sheetViews = new Map();
    this.styleTable = {
      intern: (_format) => 0,
      get: (_styleId) => null,
    };
  }

  /**
   * @param {"change"} event
   * @param {(payload: any) => void} cb
   */
  on(event, cb) {
    this._events.on(event, cb);
    return () => this._events.off(event, cb);
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} coord
   */
  getCell(sheetId, coord) {
    const key = makeCellKey(sheetId, coord.row, coord.col);
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }

  /**
   * @param {any[]} deltas
   */
  applyExternalDeltas(deltas) {
    for (const delta of deltas ?? []) {
      const key = makeCellKey(delta.sheetId, delta.row, delta.col);
      this._cells.set(key, {
        value: delta.after?.value ?? null,
        formula: delta.after?.formula ?? null,
        styleId: Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0,
      });
    }
  }

  /**
   * @param {string} sheetId
   */
  getSheetView(sheetId) {
    return this._sheetViews.get(sheetId) ?? { frozenRows: 0, frozenCols: 0 };
  }

  /**
   * @param {any[]} deltas
   */
  applyExternalSheetViewDeltas(deltas) {
    for (const delta of deltas ?? []) {
      if (!delta?.sheetId) continue;
      this._sheetViews.set(delta.sheetId, delta.after ?? { frozenRows: 0, frozenCols: 0 });
    }
  }

  /**
   * Older binder fallback methods (should not be used in this test, but implemented
   * to keep the stub compatible if binder behavior changes).
   */
  setFrozen(sheetId, frozenRows, frozenCols) {
    const prev = this.getSheetView(sheetId);
    this._sheetViews.set(sheetId, { ...prev, frozenRows, frozenCols });
  }
  setColWidth(sheetId, col, width) {
    const prev = this.getSheetView(sheetId);
    const colWidths = { ...(prev.colWidths ?? {}) };
    colWidths[String(col)] = width;
    this._sheetViews.set(sheetId, { ...prev, colWidths });
  }
  setRowHeight(sheetId, row, height) {
    const prev = this.getSheetView(sheetId);
    const rowHeights = { ...(prev.rowHeights ?? {}) };
    rowHeights[String(row)] = height;
    this._sheetViews.set(sheetId, { ...prev, rowHeights });
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} coord
   * @param {string | null} formula
   */
  setCellFormula(sheetId, coord, formula) {
    const key = makeCellKey(sheetId, coord.row, coord.col);
    const before = this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
    const after = { value: null, formula, styleId: before.styleId };
    this._cells.set(key, after);

    this._events.emit("change", {
      deltas: [
        {
          sheetId,
          row: coord.row,
          col: coord.col,
          before,
          after,
        },
      ],
    });
  }
}

test("CollabSession↔DocumentController binder edits are treated as local by FormulaConflictMonitor even when modifiedBy is missing", async () => {
  // Deterministic tie-break: higher clientID wins map entry overwrites.
  const docA = new Y.Doc();
  docA.clientID = 1;
  const docB = new Y.Doc();
  docB.clientID = 2;

  let disconnect = connectDocs(docA, docB);

  /** @type {Array<any>} */
  const conflictsA = [];

  const sessionA = createCollabSession({
    doc: docA,
    formulaConflicts: {
      localUserId: "user-a",
      onConflict: (c) => conflictsA.push(c),
    },
  });
  const sessionB = createCollabSession({ doc: docB });

  const documentController = new DocumentControllerStub();
  const binder = await bindCollabSessionToDocumentController({
    session: sessionA,
    documentController,
    userId: null,
  });

  // Establish a shared base cell map without `modifiedBy` so conflict detection
  // cannot fall back to attribution metadata.
  await sessionB.setCellFormula("Sheet1:0:0", "=0");
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.formula === "=0");
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.modifiedBy, null);

  // Ensure DocumentController has hydrated before we simulate user edits.
  await waitForCondition(() => documentController.getCell("Sheet1", { row: 0, col: 0 }).formula === "=0");

  // Offline concurrent edits:
  // - Local (via DocumentController binder): =1
  // - Remote (docB): =2
  disconnect();

  documentController.setCellFormula("Sheet1", { row: 0, col: 0 }, "=1");

  // Binder writes are async (queued). Wait for the Yjs doc to reflect the local edit.
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.formula === "=1");
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.modifiedBy, null);

  await sessionB.setCellFormula("Sheet1:0:0", "=2");

  // Reconnect and sync state.
  disconnect = connectDocs(docA, docB);

  await waitForCondition(() => conflictsA.length >= 1);
  const conflict = conflictsA[0];
  assert.equal(conflict.kind, "formula");
  assert.equal(conflict.remoteUserId, "", "expected remoteUserId to be unknown when modifiedBy is missing");

  // Optional: resolve the conflict and assert convergence.
  assert.ok(sessionA.formulaConflictMonitor?.resolveConflict(conflict.id, conflict.localFormula));
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.formula === conflict.localFormula.trim());
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.formula === conflict.localFormula.trim());

  binder.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
