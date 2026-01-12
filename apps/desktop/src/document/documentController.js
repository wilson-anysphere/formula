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

function formatKey(sheetId, layer, index) {
  return `${sheetId}:${layer}:${index == null ? "" : index}`;
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

const NUMERIC_LITERAL_RE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

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
  const trimmed = String(formula).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

/**
 * @typedef {{
 *   frozenRows: number,
 *   frozenCols: number,
 *   /**
 *    * Sparse column width overrides (base units, zoom=1), keyed by 0-based column index.
 *    * Values are interpreted by the UI layer (e.g. shared grid) and are not validated against
 *    * a default width here.
 *    *\/
 *   colWidths?: Record<string, number>,
 *   /**
 *    * Sparse row height overrides (base units, zoom=1), keyed by 0-based row index.
 *    *\/
 *   rowHeights?: Record<string, number>,
 * }} SheetViewState
 */

/**
 * @param {any} value
 * @returns {number}
 */
function normalizeFrozenCount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

/**
 * @param {any} view
 * @returns {SheetViewState}
 */
function normalizeSheetViewState(view) {
  const normalizeAxisSize = (value) => {
    const num = Number(value);
    if (!Number.isFinite(num)) return null;
    if (num <= 0) return null;
    return num;
  };

  const normalizeAxisOverrides = (raw) => {
    if (!raw) return null;

    /** @type {Record<string, number>} */
    const out = {};

    if (Array.isArray(raw)) {
      for (const entry of raw) {
        const index = Array.isArray(entry) ? entry[0] : entry?.index;
        const size = Array.isArray(entry) ? entry[1] : entry?.size;
        const idx = Number(index);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeAxisSize(size);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    } else if (typeof raw === "object") {
      for (const [key, value] of Object.entries(raw)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeAxisSize(value);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    }

    return Object.keys(out).length === 0 ? null : out;
  };

  const colWidths = normalizeAxisOverrides(view?.colWidths);
  const rowHeights = normalizeAxisOverrides(view?.rowHeights);

  return {
    frozenRows: normalizeFrozenCount(view?.frozenRows),
    frozenCols: normalizeFrozenCount(view?.frozenCols),
    ...(colWidths ? { colWidths } : {}),
    ...(rowHeights ? { rowHeights } : {}),
  };
}

/**
 * @returns {SheetViewState}
 */
function emptySheetViewState() {
  return { frozenRows: 0, frozenCols: 0 };
}

/**
 * @param {SheetViewState} view
 * @returns {SheetViewState}
 */
function cloneSheetViewState(view) {
  /** @type {SheetViewState} */
  const next = { frozenRows: view.frozenRows, frozenCols: view.frozenCols };
  if (view.colWidths) next.colWidths = { ...view.colWidths };
  if (view.rowHeights) next.rowHeights = { ...view.rowHeights };
  return next;
}

/**
 * @param {SheetViewState} a
 * @param {SheetViewState} b
 * @returns {boolean}
 */
function sheetViewStateEquals(a, b) {
  if (a === b) return true;

  const axisEquals = (left, right) => {
    if (left === right) return true;
    const leftKeys = left ? Object.keys(left) : [];
    const rightKeys = right ? Object.keys(right) : [];
    if (leftKeys.length !== rightKeys.length) return false;
    leftKeys.sort((x, y) => Number(x) - Number(y));
    rightKeys.sort((x, y) => Number(x) - Number(y));
    for (let i = 0; i < leftKeys.length; i++) {
      const key = leftKeys[i];
      if (key !== rightKeys[i]) return false;
      const lv = left[key];
      const rv = right[key];
      if (Math.abs(lv - rv) > 1e-6) return false;
    }
    return true;
  };

  return (
    a.frozenRows === b.frozenRows &&
    a.frozenCols === b.frozenCols &&
    axisEquals(a.colWidths, b.colWidths) &&
    axisEquals(a.rowHeights, b.rowHeights)
  );
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
 *   sheetId: string,
 *   before: SheetViewState,
 *   after: SheetViewState,
 * }} SheetViewDelta
 */

/**
 * Style id deltas for layered formatting.
 *
 * Layer precedence (for conflicts) is defined in `getCellFormat()`:
 * `sheet < col < row < cell`.
 *
 * @typedef {{
 *   sheetId: string,
 *   layer: "sheet" | "row" | "col",
 *   /**
 *    * Row/col index for `layer: "row"`/`"col"`.
 *    * Omitted for `layer: "sheet"`.
 *    *\/
 *   index?: number,
 *   beforeStyleId: number,
 *   afterStyleId: number,
 * }} FormatDelta
 */

/**
 * @typedef {{
 *   label?: string,
 *   mergeKey?: string,
 *   timestamp: number,
 *   deltasByCell: Map<string, CellDelta>,
 *   deltasBySheetView: Map<string, SheetViewDelta>,
 *   deltasByFormat: Map<string, FormatDelta>,
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
 * @param {SheetViewDelta} delta
 * @returns {SheetViewDelta}
 */
function cloneSheetViewDelta(delta) {
  return {
    sheetId: delta.sheetId,
    before: cloneSheetViewState(delta.before),
    after: cloneSheetViewState(delta.after),
  };
}

/**
 * @param {FormatDelta} delta
 * @returns {FormatDelta}
 */
function cloneFormatDelta(delta) {
  const out = {
    sheetId: delta.sheetId,
    layer: delta.layer,
    beforeStyleId: delta.beforeStyleId,
    afterStyleId: delta.afterStyleId,
  };
  if (delta.index != null) out.index = delta.index;
  return out;
}

/**
 * @param {HistoryEntry} entry
 * @returns {CellDelta[]}
 */
function entryCellDeltas(entry) {
  const deltas = Array.from(entry.deltasByCell.values()).map(cloneDelta);
  deltas.sort((a, b) => {
    const ak = sortKey(a.sheetId, a.row, a.col);
    const bk = sortKey(b.sheetId, b.row, b.col);
    return ak < bk ? -1 : ak > bk ? 1 : 0;
  });
  return deltas;
}

/**
 * @param {HistoryEntry} entry
 * @returns {SheetViewDelta[]}
 */
function entrySheetViewDeltas(entry) {
  const deltas = Array.from(entry.deltasBySheetView.values()).map(cloneSheetViewDelta);
  deltas.sort((a, b) => (a.sheetId < b.sheetId ? -1 : a.sheetId > b.sheetId ? 1 : 0));
  return deltas;
}

/**
 * @param {HistoryEntry} entry
 * @returns {FormatDelta[]}
 */
function entryFormatDeltas(entry) {
  const deltas = Array.from(entry.deltasByFormat.values()).map(cloneFormatDelta);
  const layerOrder = (layer) => (layer === "sheet" ? 0 : layer === "col" ? 1 : 2);
  deltas.sort((a, b) => {
    if (a.sheetId !== b.sheetId) return a.sheetId < b.sheetId ? -1 : 1;
    if (a.layer !== b.layer) return layerOrder(a.layer) - layerOrder(b.layer);
    const ai = a.index ?? -1;
    const bi = b.index ?? -1;
    return ai - bi;
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

/**
 * @param {SheetViewDelta[]} deltas
 * @returns {SheetViewDelta[]}
 */
function invertSheetViewDeltas(deltas) {
  return deltas.map((d) => ({
    sheetId: d.sheetId,
    before: cloneSheetViewState(d.after),
    after: cloneSheetViewState(d.before),
  }));
}

/**
 * @param {FormatDelta[]} deltas
 * @returns {FormatDelta[]}
 */
function invertFormatDeltas(deltas) {
  return deltas.map((d) => {
    const out = {
      sheetId: d.sheetId,
      layer: d.layer,
      beforeStyleId: d.afterStyleId,
      afterStyleId: d.beforeStyleId,
    };
    if (d.index != null) out.index = d.index;
    return out;
  });
}

class SheetModel {
  constructor() {
    /** @type {Map<string, CellState>} */
    this.cells = new Map();
    /** @type {SheetViewState} */
    this.view = emptySheetViewState();

    /**
     * Layered formatting.
     *
     * We store formatting at multiple granularities (sheet/col/row/cell) so the UI can apply
     * formatting to whole rows/columns without eagerly materializing every cell.
     *
     * The per-cell layer continues to live on `CellState.styleId`. The remaining layers live here.
     */
    this.sheetStyleId = 0;
    /** @type {Map<number, number>} */
    this.rowStyles = new Map();
    /** @type {Map<number, number>} */
    this.colStyles = new Map();
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

  /**
   * @returns {SheetViewState}
   */
  getView() {
    return cloneSheetViewState(this.view);
  }

  /**
   * @param {SheetViewState} view
   */
  setView(view) {
    this.view = cloneSheetViewState(view);
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

  /**
   * @param {string} sheetId
   * @returns {SheetViewState}
   */
  getSheetView(sheetId) {
    return this.#sheet(sheetId).getView();
  }

  /**
   * @param {string} sheetId
   * @param {SheetViewState} view
   */
  setSheetView(sheetId, view) {
    this.#sheet(sheetId).setView(view);
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
    * - `change`: {
    *     deltas: CellDelta[],
    *     sheetViewDeltas?: SheetViewDelta[],
    *     formatDeltas?: FormatDelta[],
    *     source?: string,
    *     recalc?: boolean,
    *   }
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
    if (
      this.batchDepth > 0 &&
      this.activeBatch &&
      (this.activeBatch.deltasByCell.size > 0 ||
        this.activeBatch.deltasBySheetView.size > 0 ||
        this.activeBatch.deltasByFormat.size > 0)
    ) {
      return true;
    }
    return false;
  }

  /**
   * Mark the current document state as dirty (without creating an undo step).
   *
   * This is useful for metadata changes that are persisted outside the cell grid
   * (e.g. workbook-embedded Power Query definitions).
   */
  markDirty() {
    // Mark dirty even though we didn't advance the undo cursor.
    this.savedCursor = null;
    // Avoid merging future edits into what is now considered an unsaved state.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;
    this.#emitDirty();
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
   * Return the effective formatting for a cell, taking layered styles into account.
   *
   * Merge semantics:
   * - Non-conflicting keys compose via deep merge (e.g. `{ font: { bold:true } }` + `{ font: { italic:true } }`).
   * - Conflicts resolve deterministically by layer precedence:
   *   `sheet < col < row < cell` (later layers override earlier layers for the same property).
   *
   * This mirrors the common spreadsheet model where cell-level formatting always wins, and
   * row formatting overrides column formatting when both specify the same property.
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {Record<string, any>}
   */
  getCellFormat(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;

    // Ensure the sheet is materialized (DocumentController is lazily sheet-creating).
    const cell = this.model.getCell(sheetId, c.row, c.col);
    const sheet = this.model.sheets.get(sheetId);

    const sheetStyle = this.styleTable.get(sheet?.sheetStyleId ?? 0);
    const colStyle = this.styleTable.get(sheet?.colStyles.get(c.col) ?? 0);
    const rowStyle = this.styleTable.get(sheet?.rowStyles.get(c.row) ?? 0);
    const cellStyle = this.styleTable.get(cell.styleId ?? 0);

    // Precedence: sheet < col < row < cell.
    const sheetCol = applyStylePatch(sheetStyle, colStyle);
    const sheetColRow = applyStylePatch(sheetCol, rowStyle);
    return applyStylePatch(sheetColRow, cellStyle);
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
   * Iterate over all *stored* cells in a sheet.
   *
   * This visits only entries present in the underlying sparse cell map:
   * - value cells
   * - formula cells
   * - format-only cells (styleId != 0)
   *
   * It intentionally does NOT scan the full grid area.
   *
   * NOTE: The `cell` argument is a reference to the internal model state; callers MUST
   * treat it as read-only.
   *
   * @param {string} sheetId
   * @param {(cell: { sheetId: string, row: number, col: number, cell: CellState }) => void} visitor
   */
  forEachCellInSheet(sheetId, visitor) {
    if (typeof visitor !== "function") return;
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet || !sheet.cells || sheet.cells.size === 0) return;
    for (const [key, cell] of sheet.cells.entries()) {
      if (!cell) continue;
      const { row, col } = parseRowColKey(key);
      visitor({ sheetId, row, col, cell });
    }
  }

  /**
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @param {CellValue} value
   * @param {{ mergeKey?: string, label?: string, source?: string }} [options]
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
   * @param {{ mergeKey?: string, label?: string, source?: string }} [options]
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
   * @param {{ mergeKey?: string, label?: string, source?: string }} [options]
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
   * Apply a sparse list of cell input updates in a single change event / history entry.
   *
   * This is more efficient than calling `setCellInput()` in a loop because it:
   * - emits one `change` event (instead of one per cell)
   * - batches backend sync bridges (desktop Tauri workbookSync, etc)
   *
   * @param {ReadonlyArray<{ sheetId: string, row: number, col: number, value: any, formula: string | null }>} inputs
   * @param {{ mergeKey?: string, label?: string, source?: string }} [options]
   */
  setCellInputs(inputs, options = {}) {
    if (!Array.isArray(inputs) || inputs.length === 0) return;

    /** @type {Map<string, { sheetId: string, row: number, col: number, value: any, formula: string | null }>} */
    const deduped = new Map();
    for (const input of inputs) {
      const sheetId = String(input?.sheetId ?? "").trim();
      if (!sheetId) continue;
      const row = Number(input?.row);
      const col = Number(input?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      deduped.set(`${sheetId}:${row},${col}`, {
        sheetId,
        row,
        col,
        value: input?.value ?? null,
        formula: typeof input?.formula === "string" ? input.formula : null,
      });
    }
    if (deduped.size === 0) return;

    /** @type {CellDelta[]} */
    const deltas = [];
    for (const input of deduped.values()) {
      const before = this.model.getCell(input.sheetId, input.row, input.col);
      const after = this.#normalizeCellInput(before, { value: input.value, formula: input.formula });
      if (cellStateEquals(before, after)) continue;
      deltas.push({
        sheetId: input.sheetId,
        row: input.row,
        col: input.col,
        before,
        after: cloneCellState(after),
      });
    }

    this.#applyUserDeltas(deltas, options);
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
   * Apply a formatting patch to the sheet-level formatting layer.
   *
   * @param {string} sheetId
   * @param {Record<string, any> | null} stylePatch
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  setSheetFormat(sheetId, stylePatch, options = {}) {
    // Ensure sheet exists.
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return;

    const beforeStyleId = sheet.sheetStyleId ?? 0;
    const baseStyle = this.styleTable.get(beforeStyleId);
    const merged = applyStylePatch(baseStyle, stylePatch);
    const afterStyleId = this.styleTable.intern(merged);

    if (beforeStyleId === afterStyleId) return;
    this.#applyUserFormatDeltas(
      [{ sheetId, layer: "sheet", beforeStyleId, afterStyleId }],
      options,
    );
  }

  /**
   * Apply a formatting patch to a single row formatting layer (0-based row index).
   *
   * @param {string} sheetId
   * @param {number} row
   * @param {Record<string, any> | null} stylePatch
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  setRowFormat(sheetId, row, stylePatch, options = {}) {
    const rowIdx = Number(row);
    if (!Number.isInteger(rowIdx) || rowIdx < 0) return;

    // Ensure sheet exists.
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return;

    const beforeStyleId = sheet.rowStyles.get(rowIdx) ?? 0;
    const baseStyle = this.styleTable.get(beforeStyleId);
    const merged = applyStylePatch(baseStyle, stylePatch);
    const afterStyleId = this.styleTable.intern(merged);

    if (beforeStyleId === afterStyleId) return;
    this.#applyUserFormatDeltas(
      [{ sheetId, layer: "row", index: rowIdx, beforeStyleId, afterStyleId }],
      options,
    );
  }

  /**
   * Apply a formatting patch to a single column formatting layer (0-based column index).
   *
   * @param {string} sheetId
   * @param {number} col
   * @param {Record<string, any> | null} stylePatch
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  setColFormat(sheetId, col, stylePatch, options = {}) {
    const colIdx = Number(col);
    if (!Number.isInteger(colIdx) || colIdx < 0) return;

    // Ensure sheet exists.
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return;

    const beforeStyleId = sheet.colStyles.get(colIdx) ?? 0;
    const baseStyle = this.styleTable.get(beforeStyleId);
    const merged = applyStylePatch(baseStyle, stylePatch);
    const afterStyleId = this.styleTable.intern(merged);

    if (beforeStyleId === afterStyleId) return;
    this.#applyUserFormatDeltas(
      [{ sheetId, layer: "col", index: colIdx, beforeStyleId, afterStyleId }],
      options,
    );
  }

  /**
   * Return the currently frozen pane counts for a sheet.
   *
   * @param {string} sheetId
   * @returns {SheetViewState}
   */
  getSheetView(sheetId) {
    return this.model.getSheetView(sheetId);
  }

  /**
   * Set frozen pane counts for a sheet.
   *
   * This is undoable and persisted in `encodeState()` snapshots.
   *
   * @param {string} sheetId
   * @param {number} frozenRows
   * @param {number} frozenCols
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setFrozen(sheetId, frozenRows, frozenCols, options = {}) {
    const before = this.model.getSheetView(sheetId);
    const after = cloneSheetViewState(before);
    after.frozenRows = frozenRows;
    after.frozenCols = frozenCols;
    const normalized = normalizeSheetViewState(after);
    if (sheetViewStateEquals(before, normalized)) return;
    this.#applyUserSheetViewDeltas([{ sheetId, before, after: normalized }], options);
  }

  /**
   * Set a single column width override for a sheet (base units, zoom=1).
   *
   * @param {string} sheetId
   * @param {number} col
   * @param {number | null} width
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setColWidth(sheetId, col, width, options = {}) {
    const colIdx = Number(col);
    if (!Number.isInteger(colIdx) || colIdx < 0) return;

    const before = this.model.getSheetView(sheetId);
    const after = cloneSheetViewState(before);

    const nextWidth = width == null ? null : Number(width);
    const validWidth = nextWidth != null && Number.isFinite(nextWidth) && nextWidth > 0 ? nextWidth : null;

    if (validWidth == null) {
      if (after.colWidths) {
        delete after.colWidths[String(colIdx)];
        if (Object.keys(after.colWidths).length === 0) delete after.colWidths;
      }
    } else {
      if (!after.colWidths) after.colWidths = {};
      after.colWidths[String(colIdx)] = validWidth;
    }

    if (sheetViewStateEquals(before, after)) return;
    this.#applyUserSheetViewDeltas([{ sheetId, before, after: normalizeSheetViewState(after) }], options);
  }

  /**
   * Reset a column width to the default by removing any override.
   *
   * @param {string} sheetId
   * @param {number} col
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  resetColWidth(sheetId, col, options = {}) {
    this.setColWidth(sheetId, col, null, options);
  }

  /**
   * Set a single row height override for a sheet (base units, zoom=1).
   *
   * @param {string} sheetId
   * @param {number} row
   * @param {number | null} height
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setRowHeight(sheetId, row, height, options = {}) {
    const rowIdx = Number(row);
    if (!Number.isInteger(rowIdx) || rowIdx < 0) return;

    const before = this.model.getSheetView(sheetId);
    const after = cloneSheetViewState(before);

    const nextHeight = height == null ? null : Number(height);
    const validHeight = nextHeight != null && Number.isFinite(nextHeight) && nextHeight > 0 ? nextHeight : null;

    if (validHeight == null) {
      if (after.rowHeights) {
        delete after.rowHeights[String(rowIdx)];
        if (Object.keys(after.rowHeights).length === 0) delete after.rowHeights;
      }
    } else {
      if (!after.rowHeights) after.rowHeights = {};
      after.rowHeights[String(rowIdx)] = validHeight;
    }

    if (sheetViewStateEquals(before, after)) return;
    this.#applyUserSheetViewDeltas([{ sheetId, before, after: normalizeSheetViewState(after) }], options);
  }

  /**
   * Reset a row height to the default by removing any override.
   *
   * @param {string} sheetId
   * @param {number} row
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  resetRowHeight(sheetId, row, options = {}) {
    this.setRowHeight(sheetId, row, null, options);
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
      const view = sheet?.view ? cloneSheetViewState(sheet.view) : emptySheetViewState();
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
      /** @type {any} */
      const out = { id, frozenRows: view.frozenRows, frozenCols: view.frozenCols, cells };
      if (view.colWidths && Object.keys(view.colWidths).length > 0) out.colWidths = view.colWidths;
      if (view.rowHeights && Object.keys(view.rowHeights).length > 0) out.rowHeights = view.rowHeights;

      // Layered formatting (sheet/col/row).
      if (sheet && sheet.sheetStyleId && sheet.sheetStyleId !== 0) {
        out.sheetFormat = this.styleTable.get(sheet.sheetStyleId);
      }
      if (sheet && sheet.colStyles && sheet.colStyles.size > 0) {
        /** @type {Record<string, any>} */
        const colFormats = {};
        for (const [col, styleId] of sheet.colStyles.entries()) {
          if (!styleId || styleId === 0) continue;
          colFormats[String(col)] = this.styleTable.get(styleId);
        }
        if (Object.keys(colFormats).length > 0) out.colFormats = colFormats;
      }
      if (sheet && sheet.rowStyles && sheet.rowStyles.size > 0) {
        /** @type {Record<string, any>} */
        const rowFormats = {};
        for (const [row, styleId] of sheet.rowStyles.entries()) {
          if (!styleId || styleId === 0) continue;
          rowFormats[String(row)] = this.styleTable.get(styleId);
        }
        if (Object.keys(rowFormats).length > 0) out.rowFormats = rowFormats;
      }
      return out;
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
    /** @type {Map<string, SheetViewState>} */
    const nextViews = new Map();
    /** @type {Map<string, { sheetStyleId: number, rowStyles: Map<number, number>, colStyles: Map<number, number> }>} */
    const nextFormats = new Map();

    const normalizeFormatOverrides = (raw) => {
      /** @type {Map<number, number>} */
      const out = new Map();
      if (!raw) return out;

      if (Array.isArray(raw)) {
        for (const entry of raw) {
          const index = Array.isArray(entry) ? entry[0] : entry?.index;
          const format = Array.isArray(entry) ? entry[1] : entry?.format;
          const idx = Number(index);
          if (!Number.isInteger(idx) || idx < 0) continue;
          const styleId = format == null ? 0 : this.styleTable.intern(format);
          if (styleId !== 0) out.set(idx, styleId);
        }
        return out;
      }

      if (typeof raw === "object") {
        for (const [key, value] of Object.entries(raw)) {
          const idx = Number(key);
          if (!Number.isInteger(idx) || idx < 0) continue;
          const styleId = value == null ? 0 : this.styleTable.intern(value);
          if (styleId !== 0) out.set(idx, styleId);
        }
      }

      return out;
    };
    for (const sheet of sheets) {
      if (!sheet?.id) continue;
      const cellList = Array.isArray(sheet.cells) ? sheet.cells : [];
      const view = normalizeSheetViewState({
        frozenRows: sheet?.frozenRows,
        frozenCols: sheet?.frozenCols,
        colWidths: sheet?.colWidths,
        rowHeights: sheet?.rowHeights,
      });
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
      nextViews.set(sheet.id, view);

      const sheetStyleId = sheet?.sheetFormat == null ? 0 : this.styleTable.intern(sheet.sheetFormat);
      nextFormats.set(sheet.id, {
        sheetStyleId,
        rowStyles: normalizeFormatOverrides(sheet?.rowFormats),
        colStyles: normalizeFormatOverrides(sheet?.colFormats),
      });
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

    /** @type {SheetViewDelta[]} */
    const sheetViewDeltas = [];
    for (const sheetId of allSheetIds) {
      const before = this.model.getSheetView(sheetId);
      const after = nextViews.get(sheetId) ?? emptySheetViewState();
      if (sheetViewStateEquals(before, after)) continue;
      sheetViewDeltas.push({ sheetId, before, after });
    }

    /** @type {FormatDelta[]} */
    const formatDeltas = [];
    for (const sheetId of allSheetIds) {
      const existingSheet = this.model.sheets.get(sheetId);
      const beforeSheetStyleId = existingSheet?.sheetStyleId ?? 0;
      const beforeRowStyles = existingSheet?.rowStyles ?? new Map();
      const beforeColStyles = existingSheet?.colStyles ?? new Map();

      const next = nextFormats.get(sheetId);
      const afterSheetStyleId = next?.sheetStyleId ?? 0;
      const afterRowStyles = next?.rowStyles ?? new Map();
      const afterColStyles = next?.colStyles ?? new Map();

      if (beforeSheetStyleId !== afterSheetStyleId) {
        formatDeltas.push({
          sheetId,
          layer: "sheet",
          beforeStyleId: beforeSheetStyleId,
          afterStyleId: afterSheetStyleId,
        });
      }

      const rowKeys = new Set([...beforeRowStyles.keys(), ...afterRowStyles.keys()]);
      for (const row of rowKeys) {
        const beforeStyleId = beforeRowStyles.get(row) ?? 0;
        const afterStyleId = afterRowStyles.get(row) ?? 0;
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "row", index: row, beforeStyleId, afterStyleId });
      }

      const colKeys = new Set([...beforeColStyles.keys(), ...afterColStyles.keys()]);
      for (const col of colKeys) {
        const beforeStyleId = beforeColStyles.get(col) ?? 0;
        const afterStyleId = afterColStyles.get(col) ?? 0;
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "col", index: col, beforeStyleId, afterStyleId });
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
    this.#applyEdits(deltas, sheetViewDeltas, formatDeltas, { recalc: false, emitChange: true, source: "applyState" });
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
   * @param {{ recalc?: boolean, source?: string, markDirty?: boolean }} [options]
   */
  applyExternalDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const recalc = options.recalc ?? true;
    const source = typeof options.source === "string" ? options.source : undefined;
    this.#applyEdits(deltas, [], [], { recalc, emitChange: true, source });

    // Mark dirty even though we didn't advance the undo cursor.
    //
    // Some integrations apply derived/computed updates (e.g. backend pivot auto-refresh output)
    // that should not affect dirty tracking (the user edit that triggered them already did).
    // Allow callers to suppress clearing `savedCursor` for those cases.
    if (options.markDirty !== false) {
      this.savedCursor = null;
    }
    this.#emitDirty();
  }

  /**
   * Apply a set of sheet view deltas that originated externally (e.g. collaboration sync).
   *
   * Like {@link applyExternalDeltas}, these updates:
   * - bypass undo/redo history (not user-editable)
   * - emit `change` + `update` events so UI + versioning layers can react
   * - mark the document dirty by default
   *
   * @param {SheetViewDelta[]} deltas
   * @param {{ source?: string, markDirty?: boolean }} [options]
   */
  applyExternalSheetViewDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const source = typeof options.source === "string" ? options.source : undefined;
    this.#applyEdits([], deltas, { recalc: false, emitChange: true, source });

    // Mark dirty even though we didn't advance the undo cursor.
    if (options.markDirty !== false) {
      this.savedCursor = null;
    }
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
      if (trimmed.startsWith("=")) {
        return { value: null, formula: normalizeFormula(trimmed), styleId: before.styleId };
      }

      // Excel-style scalar coercion: numeric literals and TRUE/FALSE become typed values.
      // (More complex conversions like dates are handled by number formats / future parsing layers.)
      const scalar = input.trim();
      if (scalar) {
        const upper = scalar.toUpperCase();
        if (upper === "TRUE") return { value: true, formula: null, styleId: before.styleId };
        if (upper === "FALSE") return { value: false, formula: null, styleId: before.styleId };
        if (NUMERIC_LITERAL_RE.test(scalar)) {
          const num = Number(scalar);
          if (Number.isFinite(num)) return { value: num, formula: null, styleId: before.styleId };
        }
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
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
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

    if (
      !batch ||
      (batch.deltasByCell.size === 0 && batch.deltasBySheetView.size === 0 && batch.deltasByFormat.size === 0)
    ) {
      this.#emitHistory();
      this.#emitDirty();
      return;
    }

    this.#commitHistoryEntry(batch);

    // Only recalculate for batches that included cell input edits. Sheet view changes
    // (frozen panes, row/col sizes, etc.) do not affect formula results.
    const shouldRecalc = batch.deltasByCell.size > 0;
    if (shouldRecalc) {
      this.engine?.recalculate();
      // Emit a follow-up change so observers know formula results may have changed.
      this.#emit("change", { deltas: [], sheetViewDeltas: [], formatDeltas: [], source: "endBatch", recalc: true });
    }
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
    const hadDeltas = Boolean(
      batch && (batch.deltasByCell.size > 0 || batch.deltasBySheetView.size > 0 || batch.deltasByFormat.size > 0),
    );

    // Reset batching state first so observers see consistent canUndo/canRedo.
    this.batchDepth = 0;
    this.activeBatch = null;
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    if (hadDeltas && batch) {
      const inverseCells = invertDeltas(entryCellDeltas(batch));
      const inverseViews = invertSheetViewDeltas(entrySheetViewDeltas(batch));
      const inverseFormats = invertFormatDeltas(entryFormatDeltas(batch));
      this.#applyEdits(inverseCells, inverseViews, inverseFormats, {
        recalc: false,
        emitChange: true,
        source: "cancelBatch",
      });
    }

    this.engine?.endBatch?.();

    // Only recalculate when canceling a batch that mutated cell inputs.
    const shouldRecalc = Boolean(batch && batch.deltasByCell.size > 0);
    if (shouldRecalc) {
      this.engine?.recalculate();
      this.#emit("change", { deltas: [], sheetViewDeltas: [], formatDeltas: [], source: "cancelBatch", recalc: true });
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
    const cellDeltas = entryCellDeltas(entry);
    const viewDeltas = entrySheetViewDeltas(entry);
    const formatDeltas = entryFormatDeltas(entry);
    const inverseCells = invertDeltas(cellDeltas);
    const inverseViews = invertSheetViewDeltas(viewDeltas);
    const inverseFormats = invertFormatDeltas(formatDeltas);
    this.cursor -= 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const shouldRecalc = cellDeltas.length > 0;
    this.#applyEdits(inverseCells, inverseViews, inverseFormats, {
      recalc: shouldRecalc,
      emitChange: true,
      source: "undo",
    });
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
    const cellDeltas = entryCellDeltas(entry);
    const viewDeltas = entrySheetViewDeltas(entry);
    const formatDeltas = entryFormatDeltas(entry);
    this.cursor += 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const shouldRecalc = cellDeltas.length > 0;
    this.#applyEdits(cellDeltas, viewDeltas, formatDeltas, {
      recalc: shouldRecalc,
      emitChange: true,
      source: "redo",
    });
    this.#emitHistory();
    this.#emitDirty();
    return true;
  }

  /**
   * @param {CellDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    if (this.canEditCell) {
      deltas = deltas.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
      if (deltas.length === 0) return;
    }

    const source = typeof options?.source === "string" ? options.source : undefined;
    this.#applyEdits(deltas, [], [], { recalc: this.batchDepth === 0, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(deltas, [], [], options);
  }

  /**
   * @param {CellDelta[]} deltas
   */
  #mergeIntoBatch(deltas) {
    if (!this.activeBatch) {
      // Should be unreachable, but avoid dropping history silently.
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
      };
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
   * @param {SheetViewDelta[]} deltas
   */
  #mergeSheetViewIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
      };
    }
    for (const delta of deltas) {
      const existing = this.activeBatch.deltasBySheetView.get(delta.sheetId);
      if (!existing) {
        this.activeBatch.deltasBySheetView.set(delta.sheetId, cloneSheetViewDelta(delta));
      } else {
        existing.after = cloneSheetViewState(delta.after);
      }
    }
  }

  /**
   * @param {FormatDelta[]} deltas
   */
  #mergeFormatIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
      };
    }
    for (const delta of deltas) {
      const key = formatKey(delta.sheetId, delta.layer, delta.index);
      const existing = this.activeBatch.deltasByFormat.get(key);
      if (!existing) {
        this.activeBatch.deltasByFormat.set(key, cloneFormatDelta(delta));
      } else {
        existing.afterStyleId = delta.afterStyleId;
      }
    }
  }

  /**
   * @param {SheetViewDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserSheetViewDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    this.#applyEdits([], deltas, [], { recalc: false, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeSheetViewIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry([], deltas, [], options);
  }

  /**
   * @param {FormatDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserFormatDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    this.#applyEdits([], [], deltas, { recalc: false, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeFormatIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry([], [], deltas, options);
  }

  /**
   * @param {CellDelta[]} cellDeltas
   * @param {SheetViewDelta[]} sheetViewDeltas
   * @param {FormatDelta[]} formatDeltas
   * @param {{ label?: string, mergeKey?: string }} options
   */
  #commitOrMergeHistoryEntry(cellDeltas, sheetViewDeltas, formatDeltas, options) {
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
      for (const delta of cellDeltas) {
        const key = mapKey(delta.sheetId, delta.row, delta.col);
        const existing = entry.deltasByCell.get(key);
        if (!existing) {
          entry.deltasByCell.set(key, cloneDelta(delta));
        } else {
          existing.after = cloneCellState(delta.after);
        }
      }

      for (const delta of sheetViewDeltas) {
        const existing = entry.deltasBySheetView.get(delta.sheetId);
        if (!existing) {
          entry.deltasBySheetView.set(delta.sheetId, cloneSheetViewDelta(delta));
        } else {
          existing.after = cloneSheetViewState(delta.after);
        }
      }

      for (const delta of formatDeltas) {
        const key = formatKey(delta.sheetId, delta.layer, delta.index);
        const existing = entry.deltasByFormat.get(key);
        if (!existing) {
          entry.deltasByFormat.set(key, cloneFormatDelta(delta));
        } else {
          existing.afterStyleId = delta.afterStyleId;
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
      deltasBySheetView: new Map(),
      deltasByFormat: new Map(),
    };

    for (const delta of cellDeltas) {
      entry.deltasByCell.set(mapKey(delta.sheetId, delta.row, delta.col), cloneDelta(delta));
    }

    for (const delta of sheetViewDeltas) {
      entry.deltasBySheetView.set(delta.sheetId, cloneSheetViewDelta(delta));
    }

    for (const delta of formatDeltas) {
      entry.deltasByFormat.set(formatKey(delta.sheetId, delta.layer, delta.index), cloneFormatDelta(delta));
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
    if (entry.deltasByCell.size === 0 && entry.deltasBySheetView.size === 0 && entry.deltasByFormat.size === 0) return;
    this.history.push(entry);
    this.cursor += 1;
    this.#emitHistory();
    this.#emitDirty();
  }

  /**
   * Apply deltas to the model and engine. This is the single authoritative mutation path.
   *
   * @param {CellDelta[]} cellDeltas
   * @param {SheetViewDelta[]} sheetViewDeltas
   * @param {FormatDelta[]} formatDeltas
   * @param {{ recalc: boolean, emitChange: boolean, source?: string }} options
   */
  #applyEdits(cellDeltas, sheetViewDeltas, formatDeltas, options) {
    // Apply to the canonical model first.
    for (const delta of formatDeltas) {
      // Ensure sheet exists for format-only changes.
      this.model.getCell(delta.sheetId, 0, 0);
      const sheet = this.model.sheets.get(delta.sheetId);
      if (!sheet) continue;
      if (delta.layer === "sheet") {
        sheet.sheetStyleId = delta.afterStyleId;
        continue;
      }
      const index = delta.index;
      if (index == null) continue;
      if (delta.layer === "row") {
        if (delta.afterStyleId === 0) sheet.rowStyles.delete(index);
        else sheet.rowStyles.set(index, delta.afterStyleId);
        continue;
      }
      if (delta.layer === "col") {
        if (delta.afterStyleId === 0) sheet.colStyles.delete(index);
        else sheet.colStyles.set(index, delta.afterStyleId);
      }
    }
    for (const delta of sheetViewDeltas) {
      this.model.setSheetView(delta.sheetId, delta.after);
    }
    for (const delta of cellDeltas) {
      this.model.setCell(delta.sheetId, delta.row, delta.col, delta.after);
    }

    /** @type {CellChange[] | null} */
    let engineChanges = null;
    if (this.engine && cellDeltas.length > 0) {
      engineChanges = cellDeltas.map((d) => ({
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
      for (const delta of formatDeltas) {
        this.model.getCell(delta.sheetId, 0, 0);
        const sheet = this.model.sheets.get(delta.sheetId);
        if (!sheet) continue;
        if (delta.layer === "sheet") {
          sheet.sheetStyleId = delta.beforeStyleId;
          continue;
        }
        const index = delta.index;
        if (index == null) continue;
        if (delta.layer === "row") {
          if (delta.beforeStyleId === 0) sheet.rowStyles.delete(index);
          else sheet.rowStyles.set(index, delta.beforeStyleId);
          continue;
        }
        if (delta.layer === "col") {
          if (delta.beforeStyleId === 0) sheet.colStyles.delete(index);
          else sheet.colStyles.set(index, delta.beforeStyleId);
        }
      }
      for (const delta of sheetViewDeltas) {
        this.model.setSheetView(delta.sheetId, delta.before);
      }
      for (const delta of cellDeltas) {
        this.model.setCell(delta.sheetId, delta.row, delta.col, delta.before);
      }
      throw err;
    }

    if (options.emitChange) {
      const payload = {
        deltas: cellDeltas.map(cloneDelta),
        sheetViewDeltas: sheetViewDeltas.map(cloneSheetViewDelta),
        formatDeltas: formatDeltas.map(cloneFormatDelta),
        recalc: options.recalc,
      };
      if (options.source) payload.source = options.source;
      this.#emit("change", payload);
    }

    this.#emit("update", {});
  }
}
