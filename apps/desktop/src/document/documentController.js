import { cellStateEquals, cloneCellState, emptyCellState, isCellStateEmpty } from "./cell.js";
import { normalizeRange, parseA1, parseRangeA1 } from "./coords.js";
import { applyStylePatch, StyleTable } from "../formatting/styleTable.js";

/**
 * @typedef {import("./cell.js").CellState} CellState
 * @typedef {import("./cell.js").CellValue} CellValue
 * @typedef {import("./coords.js").CellCoord} CellCoord
 * @typedef {import("./coords.js").CellRange} CellRange
 * @typedef {import("./engine.js").Engine} Engine
 * @typedef {import("./engine.js").CellChange} CellChange
 */

function mapKey(sheetId, row, col) {
  return `${sheetId}:${row},${col}`;
}

function sortKey(sheetId, row, col) {
  return `${sheetId}\u0000${row.toString().padStart(10, "0")}\u0000${col
    .toString()
    .padStart(10, "0")}`;
}

function parseRowColKey(key) {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid cell key: ${key}`);
  }
  return { row, col };
}

function semanticDiffCellKey(row, col) {
  return `r${row}c${col}`;
}

function decodeUtf8(bytes) {
  if (typeof TextDecoder !== "undefined") {
    return new TextDecoder().decode(bytes);
  }
  // Node fallback (Buffer is a Uint8Array).
  // eslint-disable-next-line no-undef
  return Buffer.from(bytes).toString("utf8");
}

function encodeUtf8(text) {
  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(text);
  }
  // eslint-disable-next-line no-undef
  return Buffer.from(text, "utf8");
}

/**
 * Canonicalize formula text for storage.
 *
 * Invariant: `CellState.formula` is either `null` or a string starting with "=".
 *
 * @param {string | null | undefined} formula
 * @returns {string | null}
 */
function normalizeFormula(formula) {
  if (formula == null) return null;
  const trimmed = String(formula).trimStart();
  if (trimmed === "") return null;
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

/**
 * @typedef {{
 *   sheetId: string,
 *   row: number,
 *   col: number,
 *   before: CellState,
 *   after: CellState,
 * }} CellDelta
 */

/**
 * @typedef {{
 *   label?: string,
 *   mergeKey?: string,
 *   timestamp: number,
 *   deltasByCell: Map<string, CellDelta>,
 * }} HistoryEntry
 */

function cloneDelta(delta) {
  return {
    sheetId: delta.sheetId,
    row: delta.row,
    col: delta.col,
    before: cloneCellState(delta.before),
    after: cloneCellState(delta.after),
  };
}

/**
 * @param {HistoryEntry} entry
 * @returns {CellDelta[]}
 */
function entryDeltas(entry) {
  const deltas = Array.from(entry.deltasByCell.values()).map(cloneDelta);
  deltas.sort((a, b) => {
    const ak = sortKey(a.sheetId, a.row, a.col);
    const bk = sortKey(b.sheetId, b.row, b.col);
    return ak < bk ? -1 : ak > bk ? 1 : 0;
  });
  return deltas;
}

/**
 * @param {CellDelta[]} deltas
 * @returns {CellDelta[]}
 */
function invertDeltas(deltas) {
  return deltas.map((d) => ({
    sheetId: d.sheetId,
    row: d.row,
    col: d.col,
    before: cloneCellState(d.after),
    after: cloneCellState(d.before),
  }));
}

class SheetModel {
  constructor() {
    /** @type {Map<string, CellState>} */
    this.cells = new Map();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @returns {CellState}
   */
  getCell(row, col) {
    return cloneCellState(this.cells.get(`${row},${col}`) ?? emptyCellState());
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {CellState} cell
   */
  setCell(row, col, cell) {
    if (isCellStateEmpty(cell)) {
      this.cells.delete(`${row},${col}`);
      return;
    }
    this.cells.set(`${row},${col}`, cloneCellState(cell));
  }
}

class WorkbookModel {
  constructor() {
    /** @type {Map<string, SheetModel>} */
    this.sheets = new Map();
  }

  /**
   * @param {string} sheetId
   * @returns {SheetModel}
   */
  #sheet(sheetId) {
    let sheet = this.sheets.get(sheetId);
    if (!sheet) {
      sheet = new SheetModel();
      this.sheets.set(sheetId, sheet);
    }
    return sheet;
  }

  /**
   * @param {string} sheetId
   * @param {number} row
   * @param {number} col
   * @returns {CellState}
   */
  getCell(sheetId, row, col) {
    return this.#sheet(sheetId).getCell(row, col);
  }

  /**
   * @param {string} sheetId
   * @param {number} row
   * @param {number} col
   * @param {CellState} cell
   */
  setCell(sheetId, row, col, cell) {
    this.#sheet(sheetId).setCell(row, col, cell);
  }
}

/**
 * DocumentController is the authoritative state machine for a workbook.
 *
 * It owns:
 * - The canonical cell inputs (value/formula/styleId)
 * - Undo/redo stacks (with inversion)
 * - Dirty tracking since last save
 * - Optional integration hooks for an external calc engine and UI layers
 */
export class DocumentController {
  /**
   * @param {{
   *   engine?: Engine,
   *   mergeWindowMs?: number,
   *   canEditCell?: (cell: { sheetId: string, row: number, col: number }) => boolean
   * }} [options]
   */
  constructor(options = {}) {
    /** @type {Engine | null} */
    this.engine = options.engine ?? null;

    this.mergeWindowMs = options.mergeWindowMs ?? 1000;

    this.canEditCell = typeof options.canEditCell === "function" ? options.canEditCell : null;

    this.model = new WorkbookModel();
    this.styleTable = new StyleTable();

    /** @type {HistoryEntry[]} */
    this.history = [];
    this.cursor = 0;
    /** @type {number | null} */
    this.savedCursor = 0;

    this.batchDepth = 0;
    /** @type {HistoryEntry | null} */
    this.activeBatch = null;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
  }

  /**
   * Subscribe to controller events.
   *
   * Events:
   * - `change`: { deltas: CellDelta[], source?: string, recalc?: boolean }
   * - `history`: { canUndo: boolean, canRedo: boolean }
   * - `dirty`: { isDirty: boolean }
   * - `update`: emitted after any applied change (including undo/redo) for versioning adapters
   *
   * @template {string} T
   * @param {T} event
   * @param {(payload: any) => void} listener
   * @returns {() => void}
   */
  on(event, listener) {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(listener);
    return () => set.delete(listener);
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  #emit(event, payload) {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }

  #emitHistory() {
    this.#emit("history", { canUndo: this.canUndo, canRedo: this.canRedo });
  }

  #emitDirty() {
    this.#emit("dirty", { isDirty: this.isDirty });
  }

  /**
   * @returns {boolean}
   */
  get canUndo() {
    return this.batchDepth === 0 && this.cursor > 0;
  }

  /**
   * @returns {boolean}
   */
  get canRedo() {
    return this.batchDepth === 0 && this.cursor < this.history.length;
  }

  /**
   * @returns {boolean}
   */
  get isDirty() {
    if (this.savedCursor == null) return true;
    if (this.cursor !== this.savedCursor) return true;
    // While a batch is active we may have applied uncommitted changes to the
    // model/engine. Those should still be treated as "dirty" for close prompts.
    if (this.batchDepth > 0 && this.activeBatch && this.activeBatch.deltasByCell.size > 0) {
      return true;
    }
    return false;
  }

  /**
   * Mark the current state as saved (not dirty).
   */
  markSaved() {
    this.savedCursor = this.cursor;
    // Avoid merging future edits into what is now the saved state.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;
    this.#emitDirty();
  }

  /**
   * @returns {{ undo: number, redo: number }}
   */
  getStackDepths() {
    return { undo: this.cursor, redo: this.history.length - this.cursor };
  }

  /**
   * Convenience labels for menu items ("Undo Paste", etc).
   *
   * @returns {string | null}
   */
  get undoLabel() {
    if (!this.canUndo) return null;
    return this.history[this.cursor - 1]?.label ?? null;
  }

  /**
   * @returns {string | null}
   */
  get redoLabel() {
    if (!this.canRedo) return null;
    return this.history[this.cursor]?.label ?? null;
  }

  /**
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {CellState}
   */
  getCell(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    return this.model.getCell(sheetId, c.row, c.col);
  }

  /**
   * Return the set of sheet ids that exist in the underlying model.
   *
   * Note: the DocumentController currently creates sheets lazily when a sheet id is first
   * referenced by an edit/read. Empty workbooks will return an empty array until at least
   * one cell is accessed.
   *
   * @returns {string[]}
   */
  getSheetIds() {
    return Array.from(this.model.sheets.keys());
  }

  /**
   * Compute the bounding box of non-empty cells in a sheet.
   *
   * By default this ignores format-only cells (value/formula must be present).
   *
   * @param {string} sheetId
   * @param {{ includeFormat?: boolean }} [options]
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  getUsedRange(sheetId, options = {}) {
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet || !sheet.cells || sheet.cells.size === 0) return null;

    const includeFormat = Boolean(options.includeFormat);

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;
    let hasData = false;

    for (const [key, cell] of sheet.cells.entries()) {
      if (!cell) continue;
      const hasContent = includeFormat
        ? cell.value != null || cell.formula != null || cell.styleId !== 0
        : cell.value != null || cell.formula != null;
      if (!hasContent) continue;

      const { row, col } = parseRowColKey(key);
      hasData = true;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }

    if (!hasData) return null;
    return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
  }

  /**
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @param {CellValue} value
   * @param {{ mergeKey?: string, label?: string }} [options]
   */
  setCellValue(sheetId, coord, value, options = {}) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const before = this.model.getCell(sheetId, c.row, c.col);
    const after = { value: value ?? null, formula: null, styleId: before.styleId };
    this.#applyUserDeltas([
      { sheetId, row: c.row, col: c.col, before, after: cloneCellState(after) },
    ], options);
  }

  /**
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @param {string | null} formula
   * @param {{ mergeKey?: string, label?: string }} [options]
   */
  setCellFormula(sheetId, coord, formula, options = {}) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const before = this.model.getCell(sheetId, c.row, c.col);
    const after = { value: null, formula: normalizeFormula(formula), styleId: before.styleId };
    this.#applyUserDeltas([
      { sheetId, row: c.row, col: c.col, before, after: cloneCellState(after) },
    ], options);
  }

  /**
   * Set a cell from raw user input (e.g. formula bar / cell editor contents).
   *
   * - Strings starting with "=" are treated as formulas.
   * - Strings starting with "'" have the apostrophe stripped and are treated as literal text.
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @param {any} input
   * @param {{ mergeKey?: string, label?: string }} [options]
   */
  setCellInput(sheetId, coord, input, options = {}) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const before = this.model.getCell(sheetId, c.row, c.col);
    const after = this.#normalizeCellInput(before, input);
    if (cellStateEquals(before, after)) return;
    this.#applyUserDeltas(
      [{ sheetId, row: c.row, col: c.col, before, after: cloneCellState(after) }],
      options
    );
  }

  /**
   * Clear a single cell's contents (preserving formatting).
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @param {{ label?: string }} [options]
   */
  clearCell(sheetId, coord, options = {}) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    this.clearRange(sheetId, { start: c, end: c }, options);
  }

  /**
   * Clear values/formulas (preserving formats).
   *
   * @param {string} sheetId
   * @param {CellRange | string} range
   * @param {{ label?: string }} [options]
   */
  clearRange(sheetId, range, options = {}) {
    const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
    /** @type {CellDelta[]} */
    const deltas = [];
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const before = this.model.getCell(sheetId, row, col);
        const after = { value: null, formula: null, styleId: before.styleId };
        if (cellStateEquals(before, after)) continue;
        deltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }
    }
    this.#applyUserDeltas(deltas, { label: options.label });
  }

  /**
   * Set values/formulas in a rectangular region.
   *
   * The range can either be inferred from `values` dimensions (when `range` is a single
   * start cell) or explicitly provided as an A1 range (e.g. "A1:B2").
   *
   * @param {string} sheetId
   * @param {CellCoord | string | CellRange} rangeOrStart
   * @param {ReadonlyArray<ReadonlyArray<any>>} values
   * @param {{ label?: string }} [options]
   */
  setRangeValues(sheetId, rangeOrStart, values, options = {}) {
    if (!Array.isArray(values) || values.length === 0) return;
    const rowCount = values.length;
    const colCount = Math.max(...values.map((row) => (Array.isArray(row) ? row.length : 0)));
    if (colCount === 0) return;

    /** @type {CellRange} */
    let range;
    if (typeof rangeOrStart === "string") {
      if (rangeOrStart.includes(":")) {
        range = parseRangeA1(rangeOrStart);
      } else {
        const start = parseA1(rangeOrStart);
        range = { start, end: { row: start.row + rowCount - 1, col: start.col + colCount - 1 } };
      }
    } else if (rangeOrStart && "start" in rangeOrStart && "end" in rangeOrStart) {
      range = normalizeRange(rangeOrStart);
    } else {
      const start = /** @type {CellCoord} */ (rangeOrStart);
      range = { start, end: { row: start.row + rowCount - 1, col: start.col + colCount - 1 } };
    }

    /** @type {CellDelta[]} */
    const deltas = [];
    for (let r = 0; r < rowCount; r++) {
      const rowValues = values[r] ?? [];
      for (let c = 0; c < colCount; c++) {
        const input = rowValues[c] ?? null;
        const row = range.start.row + r;
        const col = range.start.col + c;
        const before = this.model.getCell(sheetId, row, col);
        const next = this.#normalizeCellInput(before, input);
        if (cellStateEquals(before, next)) continue;
        deltas.push({ sheetId, row, col, before, after: cloneCellState(next) });
      }
    }

    this.#applyUserDeltas(deltas, { label: options.label });
  }

  /**
   * Apply a formatting patch to a range.
   *
   * @param {string} sheetId
   * @param {CellRange | string} range
   * @param {Record<string, any> | null} stylePatch
   * @param {{ label?: string }} [options]
   */
  setRangeFormat(sheetId, range, stylePatch, options = {}) {
    const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
    /** @type {CellDelta[]} */
    const deltas = [];
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const before = this.model.getCell(sheetId, row, col);
        const baseStyle = this.styleTable.get(before.styleId);
        const merged = applyStylePatch(baseStyle, stylePatch);
        const afterStyleId = this.styleTable.intern(merged);
        const after = { value: before.value, formula: before.formula, styleId: afterStyleId };
        if (cellStateEquals(before, after)) continue;
        deltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }
    }
    this.#applyUserDeltas(deltas, { label: options.label });
  }

  /**
   * Export a single sheet in the `SheetState` shape expected by the versioning `semanticDiff`.
   *
   * @param {string} sheetId
   * @returns {{ cells: Map<string, { value?: any, formula?: string | null, format?: any }> }}
   */
  exportSheetForSemanticDiff(sheetId) {
    const sheet = this.model.sheets.get(sheetId);
    /** @type {Map<string, any>} */
    const cells = new Map();
    if (!sheet) return { cells };

    for (const [key, cell] of sheet.cells.entries()) {
      const { row, col } = parseRowColKey(key);
      cells.set(semanticDiffCellKey(row, col), {
        value: cell.value ?? null,
        formula: cell.formula ?? null,
        format: cell.styleId === 0 ? null : this.styleTable.get(cell.styleId),
      });
    }
    return { cells };
  }

  /**
   * Encode the document's current cell inputs as a snapshot suitable for the VersionManager.
   *
   * Undo/redo history is intentionally *not* included; snapshots represent workbook contents.
   *
   * @returns {Uint8Array}
   */
  encodeState() {
    const sheetIds = Array.from(this.model.sheets.keys()).sort();
    const sheets = sheetIds.map((id) => {
      const sheet = this.model.sheets.get(id);
      const cells = Array.from(sheet?.cells.entries() ?? []).map(([key, cell]) => {
        const { row, col } = parseRowColKey(key);
        return {
          row,
          col,
          value: cell.value ?? null,
          formula: cell.formula ?? null,
          format: cell.styleId === 0 ? null : this.styleTable.get(cell.styleId),
        };
      });
      cells.sort((a, b) => (a.row - b.row === 0 ? a.col - b.col : a.row - b.row));
      return { id, cells };
    });

    return encodeUtf8(JSON.stringify({ schemaVersion: 1, sheets }));
  }

  /**
   * Replace the workbook state from a snapshot produced by `encodeState`.
   *
   * This method clears undo/redo history (restoring a version is not itself undoable) and
   * marks the document dirty until the host calls `markSaved()`.
   *
   * @param {Uint8Array} snapshot
   */
  applyState(snapshot) {
    const parsed = JSON.parse(decodeUtf8(snapshot));
    const sheets = Array.isArray(parsed?.sheets) ? parsed.sheets : [];

    /** @type {Map<string, Map<string, CellState>>} */
    const nextSheets = new Map();
    for (const sheet of sheets) {
      if (!sheet?.id) continue;
      const cellList = Array.isArray(sheet.cells) ? sheet.cells : [];
      /** @type {Map<string, CellState>} */
      const cellMap = new Map();
      for (const entry of cellList) {
        const row = Number(entry?.row);
        const col = Number(entry?.col);
        if (!Number.isInteger(row) || row < 0) continue;
        if (!Number.isInteger(col) || col < 0) continue;
        const format = entry?.format ?? null;
        const styleId = format == null ? 0 : this.styleTable.intern(format);
        const cell = { value: entry?.value ?? null, formula: normalizeFormula(entry?.formula), styleId };
        cellMap.set(`${row},${col}`, cloneCellState(cell));
      }
      nextSheets.set(sheet.id, cellMap);
    }

    const existingSheetIds = new Set(this.model.sheets.keys());
    const nextSheetIds = new Set(nextSheets.keys());
    const allSheetIds = new Set([...existingSheetIds, ...nextSheetIds]);
    const removedSheetIds = Array.from(existingSheetIds).filter((id) => !nextSheetIds.has(id));

    /** @type {CellDelta[]} */
    const deltas = [];
    for (const sheetId of allSheetIds) {
      const nextCellMap = nextSheets.get(sheetId) ?? new Map();
      const existingSheet = this.model.sheets.get(sheetId);
      const existingKeys = existingSheet ? Array.from(existingSheet.cells.keys()) : [];
      const nextKeys = Array.from(nextCellMap.keys());
      const allKeys = new Set([...existingKeys, ...nextKeys]);

      for (const key of allKeys) {
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const after = nextCellMap.get(key) ?? emptyCellState();
        if (cellStateEquals(before, after)) continue;
        deltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }
    }

    // Ensure all snapshot sheet ids exist even when they contain no cells (the model is otherwise
    // lazily materialized via reads/writes).
    for (const sheetId of nextSheetIds) {
      this.model.getCell(sheetId, 0, 0);
    }

    // Clear history first: restoring content is not itself undoable.
    this.history = [];
    this.cursor = 0;
    this.savedCursor = null;
    this.batchDepth = 0;
    this.activeBatch = null;
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    // Apply changes as a single engine batch.
    this.engine?.beginBatch?.();
    this.#applyDeltas(deltas, { recalc: false, emitChange: true, source: "applyState" });
    this.engine?.endBatch?.();
    this.engine?.recalculate();

    for (const sheetId of removedSheetIds) {
      this.model.sheets.delete(sheetId);
    }

    this.#emitHistory();
    this.#emitDirty();
  }

  /**
   * Apply a set of deltas that originated externally (e.g. collaboration sync).
   *
   * Unlike user edits, these changes:
   * - bypass `canEditCell` (permissions should be enforced at the collaboration layer)
   * - do NOT create a new undo/redo history entry
   *
   * They still emit `change` + `update` events so UI layers can react, and they
   * mark the document dirty.
   *
   * @param {CellDelta[]} deltas
   * @param {{ recalc?: boolean }} [options]
   */
  applyExternalDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const recalc = options.recalc ?? true;
    this.#applyDeltas(deltas, { recalc, emitChange: true });

    // Mark dirty even though we didn't advance the undo cursor.
    this.savedCursor = null;
    this.#emitDirty();
  }

  /**
   * @param {CellState} before
   * @param {any} input
   * @returns {CellState}
   */
  #normalizeCellInput(before, input) {
    // Object form: { value?, formula?, styleId?, format? }.
    if (
      input &&
      typeof input === "object" &&
      ("formula" in input || "value" in input || "styleId" in input || "format" in input)
    ) {
      /** @type {any} */
      const obj = input;

      let value = before.value;
      let formula = before.formula;
      let styleId = before.styleId;

      if ("styleId" in obj) {
        const next = Number(obj.styleId);
        styleId = Number.isInteger(next) && next >= 0 ? next : 0;
      } else if ("format" in obj) {
        const format = obj.format ?? null;
        styleId = format == null ? 0 : this.styleTable.intern(format);
      }

      if ("formula" in obj) {
        const nextFormula = typeof obj.formula === "string" ? normalizeFormula(obj.formula) : null;
        formula = nextFormula;
        if (nextFormula != null) {
          value = null;
        } else if ("value" in obj) {
          value = obj.value ?? null;
          formula = null;
        }
      } else if ("value" in obj) {
        value = obj.value ?? null;
        formula = null;
      }

      return { value, formula, styleId };
    }

    // String primitives: interpret leading "=" as a formula, and leading apostrophe as a literal.
    if (typeof input === "string") {
      if (input.startsWith("'")) {
        return { value: input.slice(1), formula: null, styleId: before.styleId };
      }

      const trimmed = input.trimStart();
      if (trimmed.startsWith("=") && trimmed.length > 1) {
        return { value: null, formula: normalizeFormula(trimmed), styleId: before.styleId };
      }
    }

    // Primitive value (or null => clear); styles preserved.
    return { value: input ?? null, formula: null, styleId: before.styleId };
  }

  /**
   * Start an explicit batch. All subsequent edits are merged into one undo step until `endBatch`.
   *
   * @param {{ label?: string }} [options]
   */
  beginBatch(options = {}) {
    this.batchDepth += 1;
    if (this.batchDepth === 1) {
      this.activeBatch = {
        label: options.label,
        timestamp: Date.now(),
        deltasByCell: new Map(),
      };
      this.engine?.beginBatch?.();
      this.#emitHistory();
    }
  }

  /**
   * Commit the current batch into the undo stack.
   */
  endBatch() {
    if (this.batchDepth === 0) return;
    this.batchDepth -= 1;
    if (this.batchDepth > 0) return;

    const batch = this.activeBatch;
    this.activeBatch = null;
    this.engine?.endBatch?.();

    if (!batch || batch.deltasByCell.size === 0) {
      this.#emitHistory();
      this.#emitDirty();
      return;
    }

    this.#commitHistoryEntry(batch);
    this.engine?.recalculate();
    this.#emit("change", { deltas: [], source: "endBatch", recalc: true });
  }

  /**
   * Cancel the current batch by reverting all changes applied since `beginBatch()`.
   *
   * This is useful for editor cancellation (Esc) when the UI updates the document
   * incrementally while the user types.
   *
   * @returns {boolean} Whether any changes were reverted.
   */
  cancelBatch() {
    if (this.batchDepth === 0) return false;

    const batch = this.activeBatch;
    const hadDeltas = Boolean(batch && batch.deltasByCell.size > 0);

    // Reset batching state first so observers see consistent canUndo/canRedo.
    this.batchDepth = 0;
    this.activeBatch = null;
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    if (hadDeltas && batch) {
      const inverse = invertDeltas(entryDeltas(batch));
      this.#applyDeltas(inverse, { recalc: false, emitChange: true });
    }

    this.engine?.endBatch?.();
    if (hadDeltas) this.engine?.recalculate();
    if (hadDeltas) {
      this.#emit("change", { deltas: [], source: "cancelBatch", recalc: true });
    }

    this.#emitHistory();
    this.#emitDirty();
    return hadDeltas;
  }

  /**
   * Undo the most recent committed history entry.
   * @returns {boolean} Whether an undo occurred
   */
  undo() {
    if (!this.canUndo) return false;
    const entry = this.history[this.cursor - 1];
    const deltas = entryDeltas(entry);
    const inverse = invertDeltas(deltas);
    this.cursor -= 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    this.#applyDeltas(inverse, { recalc: true, emitChange: true });
    this.#emitHistory();
    this.#emitDirty();
    return true;
  }

  /**
   * Redo the next history entry.
   * @returns {boolean} Whether a redo occurred
   */
  redo() {
    if (!this.canRedo) return false;
    const entry = this.history[this.cursor];
    const deltas = entryDeltas(entry);
    this.cursor += 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    this.#applyDeltas(deltas, { recalc: true, emitChange: true });
    this.#emitHistory();
    this.#emitDirty();
    return true;
  }

  /**
   * @param {CellDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string }} options
   */
  #applyUserDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    if (this.canEditCell) {
      deltas = deltas.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
      if (deltas.length === 0) return;
    }

    this.#applyDeltas(deltas, { recalc: this.batchDepth === 0, emitChange: true });

    if (this.batchDepth > 0) {
      this.#mergeIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(deltas, options);
  }

  /**
   * @param {CellDelta[]} deltas
   */
  #mergeIntoBatch(deltas) {
    if (!this.activeBatch) {
      // Should be unreachable, but avoid dropping history silently.
      this.activeBatch = { timestamp: Date.now(), deltasByCell: new Map() };
    }
    for (const delta of deltas) {
      const key = mapKey(delta.sheetId, delta.row, delta.col);
      const existing = this.activeBatch.deltasByCell.get(key);
      if (!existing) {
        this.activeBatch.deltasByCell.set(key, cloneDelta(delta));
      } else {
        existing.after = cloneCellState(delta.after);
      }
    }
  }

  /**
   * @param {CellDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string }} options
   */
  #commitOrMergeHistoryEntry(deltas, options) {
    // If we have redo history, truncate it before pushing a new edit.
    if (this.cursor < this.history.length) {
      if (this.savedCursor != null && this.savedCursor > this.cursor) {
        // The saved state is no longer reachable once we branch.
        this.savedCursor = null;
      }
      this.history.splice(this.cursor);
      this.lastMergeKey = null;
      this.lastMergeTime = 0;
    }

    const now = Date.now();
    const mergeKey = options.mergeKey;
    const canMerge =
      mergeKey &&
      this.cursor > 0 &&
      this.cursor === this.history.length &&
      this.lastMergeKey === mergeKey &&
      now - this.lastMergeTime < this.mergeWindowMs &&
      // Never mutate what has been marked as saved.
      (this.savedCursor == null || this.cursor > this.savedCursor);

    if (canMerge) {
      const entry = this.history[this.cursor - 1];
      for (const delta of deltas) {
        const key = mapKey(delta.sheetId, delta.row, delta.col);
        const existing = entry.deltasByCell.get(key);
        if (!existing) {
          entry.deltasByCell.set(key, cloneDelta(delta));
        } else {
          existing.after = cloneCellState(delta.after);
        }
      }
      entry.timestamp = now;
      entry.mergeKey = mergeKey;
      entry.label = options.label ?? entry.label;

      this.lastMergeKey = mergeKey;
      this.lastMergeTime = now;

      this.#emitHistory();
      this.#emitDirty();
      return;
    }

    const entry = {
      label: options.label,
      mergeKey,
      timestamp: now,
      deltasByCell: new Map(),
    };

    for (const delta of deltas) {
      entry.deltasByCell.set(mapKey(delta.sheetId, delta.row, delta.col), cloneDelta(delta));
    }

    this.#commitHistoryEntry(entry);

    if (mergeKey) {
      this.lastMergeKey = mergeKey;
      this.lastMergeTime = now;
    } else {
      this.lastMergeKey = null;
      this.lastMergeTime = 0;
    }
  }

  /**
   * @param {HistoryEntry} entry
   */
  #commitHistoryEntry(entry) {
    if (entry.deltasByCell.size === 0) return;
    this.history.push(entry);
    this.cursor += 1;
    this.#emitHistory();
    this.#emitDirty();
  }

  /**
   * Apply deltas to the model and engine. This is the single authoritative mutation path.
   *
   * @param {CellDelta[]} deltas
   * @param {{ recalc: boolean, emitChange: boolean, source?: string }} options
   */
  #applyDeltas(deltas, options) {
    // Apply to the canonical model first.
    for (const delta of deltas) {
      this.model.setCell(delta.sheetId, delta.row, delta.col, delta.after);
    }

    /** @type {CellChange[] | null} */
    let engineChanges = null;
    if (this.engine) {
      engineChanges = deltas.map((d) => ({
        sheetId: d.sheetId,
        row: d.row,
        col: d.col,
        cell: cloneCellState(d.after),
      }));
    }

    try {
      if (engineChanges) this.engine.applyChanges(engineChanges);
      if (options.recalc) this.engine?.recalculate();
    } catch (err) {
      // Roll back the canonical model if the engine rejects a change.
      for (const delta of deltas) {
        this.model.setCell(delta.sheetId, delta.row, delta.col, delta.before);
      }
      throw err;
    }

    if (options.emitChange) {
      const payload = { deltas: deltas.map(cloneDelta), recalc: options.recalc };
      if (options.source) payload.source = options.source;
      this.#emit("change", payload);
    }

    this.#emit("update", {});
  }
}
