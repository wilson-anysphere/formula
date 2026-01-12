import {
  cellStateEquals,
  cloneCellState,
  emptyCellState,
  normalizeCellState,
} from "./cell.js";
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

// Excel grid limits (used by the UI selection model and for scalable formatting ops).
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLS = 16_384;
const EXCEL_MAX_ROW = EXCEL_MAX_ROWS - 1;
const EXCEL_MAX_COL = EXCEL_MAX_COLS - 1;

/**
 * @param {any} a
 * @param {any} b
 * @returns {boolean}
 */
function cellValueEquals(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return a === b;
  if (typeof a !== typeof b) return false;
  if (typeof a === "object") {
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch {
      return false;
    }
  }
  return false;
}

/**
 * Compare only the *content* portion of a cell state (value/formula), ignoring styleId.
 *
 * This is intentionally aligned with how we construct AI workbook context: formatting-only
 * changes should not invalidate caches.
 *
 * @param {CellState} a
 * @param {CellState} b
 * @returns {boolean}
 */
function cellContentEquals(a, b) {
  return cellValueEquals(a?.value ?? null, b?.value ?? null) && (a?.formula ?? null) === (b?.formula ?? null);
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

/**
 * @param {CellDelta[]} deltas
 * @returns {boolean}
 */
function cellDeltasAffectRecalc(deltas) {
  for (const d of deltas) {
    if (!d) continue;
    if ((d.before?.formula ?? null) !== (d.after?.formula ?? null)) return true;
    if ((d.before?.value ?? null) !== (d.after?.value ?? null)) return true;
  }
  return false;
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
    this.defaultStyleId = 0;
    /** @type {Map<number, number>} */
    this.rowStyleIds = new Map();
    /** @type {Map<number, number>} */
    this.colStyleIds = new Map();

    /**
     * Bounding box of row-level formatting overrides (rowStyleIds).
     *
     * This represents the used-range impact of row formatting alone (full width columns).
     *
     * @type {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
     */
    this.rowStyleBounds = null;

    /**
     * Bounding box of column-level formatting overrides (colStyleIds).
     *
     * This represents the used-range impact of column formatting alone (full height rows).
     *
     * @type {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
     */
    this.colStyleBounds = null;

    /**
     * Bounding box of cells with user-visible contents (value/formula).
     *
     * This intentionally ignores format-only cells so default `getUsedRange()` preserves its
     * historical semantics.
     *
     * @type {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
     */
    this.contentBounds = null;

    /**
     * Bounding box of non-empty *stored* cell states (value/formula/cell-level styleId).
     *
     * Note: This does NOT include row/col/sheet formatting layers (those are tracked separately
     * on the sheet model and are incorporated by `DocumentController.getUsedRange({ includeFormat:true })`).
     *
     * @type {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
     */
    this.formatBounds = null;

    // Bounds invalidation flags. We avoid eager rescans during large edits (e.g. clearRange)
    // by lazily recomputing on demand when a boundary cell is cleared.
    this.contentBoundsDirty = false;
    this.formatBoundsDirty = false;
    this.rowStyleBoundsDirty = false;
    this.colStyleBoundsDirty = false;

    // Track the number of cells that contribute to `contentBounds` so we can fast-path the
    // empty case (common when clearing contents but preserving styles).
    this.contentCellCount = 0;

    // Debug counters for unit tests to verify recomputation only occurs when required.
    this.__contentBoundsRecomputeCount = 0;
    this.__formatBoundsRecomputeCount = 0;
    this.__rowStyleBoundsRecomputeCount = 0;
    this.__colStyleBoundsRecomputeCount = 0;
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
    const key = `${row},${col}`;
    const before = this.cells.get(key) ?? null;
    const beforeHasContent = Boolean(before && (before.value != null || before.formula != null));
    const beforeHasFormat = Boolean(before);

    const normalized = normalizeCellState(cell);
    const afterIsEmpty = normalized.value == null && normalized.formula == null && normalized.styleId === 0;
    const afterHasContent = Boolean(normalized.value != null || normalized.formula != null);
    const afterHasFormat = !afterIsEmpty;

    // Update the canonical cell map first.
    if (afterIsEmpty) {
      this.cells.delete(key);
    } else {
      this.cells.set(key, cloneCellState(normalized));
    }

    // Maintain content-cell count.
    if (beforeHasContent !== afterHasContent) {
      this.contentCellCount += afterHasContent ? 1 : -1;
      if (this.contentCellCount < 0) this.contentCellCount = 0;
    }

    const expandBounds = (bounds) => {
      bounds.startRow = Math.min(bounds.startRow, row);
      bounds.endRow = Math.max(bounds.endRow, row);
      bounds.startCol = Math.min(bounds.startCol, col);
      bounds.endCol = Math.max(bounds.endCol, col);
    };

    const isOnEdge = (bounds) =>
      row === bounds.startRow || row === bounds.endRow || col === bounds.startCol || col === bounds.endCol;

    // Update content bounds (value/formula only).
    if (afterHasContent) {
      if (!this.contentBounds) {
        this.contentBounds = { startRow: row, endRow: row, startCol: col, endCol: col };
        this.contentBoundsDirty = false;
      } else {
        expandBounds(this.contentBounds);
      }
    } else if (beforeHasContent) {
      // Content removed (or converted to style-only).
      if (this.contentCellCount === 0) {
        this.contentBounds = null;
        this.contentBoundsDirty = false;
      } else if (this.contentBounds && !this.contentBoundsDirty && isOnEdge(this.contentBounds)) {
        this.contentBoundsDirty = true;
      }
    }

    // Update format bounds (any non-empty cell state).
    if (afterHasFormat) {
      if (!this.formatBounds) {
        this.formatBounds = { startRow: row, endRow: row, startCol: col, endCol: col };
        this.formatBoundsDirty = false;
      } else {
        expandBounds(this.formatBounds);
      }
    } else if (beforeHasFormat) {
      // Entire cell state cleared (including style).
      if (this.cells.size === 0) {
        this.formatBounds = null;
        this.formatBoundsDirty = false;
      } else if (this.formatBounds && !this.formatBoundsDirty && isOnEdge(this.formatBounds)) {
        this.formatBoundsDirty = true;
      }
    }
  }

  /**
   * @param {number} row
   * @param {number} styleId
   */
  setRowStyleId(row, styleId) {
    const rowIdx = Number(row);
    if (!Number.isInteger(rowIdx) || rowIdx < 0) return;
    const nextStyle = Number(styleId);
    const afterStyleId = Number.isInteger(nextStyle) && nextStyle >= 0 ? nextStyle : 0;

    const beforeStyleId = this.rowStyleIds.get(rowIdx) ?? 0;
    if (beforeStyleId === afterStyleId) return;

    const beforeHad = beforeStyleId !== 0;
    const afterHas = afterStyleId !== 0;

    if (afterHas) {
      this.rowStyleIds.set(rowIdx, afterStyleId);
    } else {
      this.rowStyleIds.delete(rowIdx);
    }

    if (afterHas) {
      if (!this.rowStyleBounds) {
        this.rowStyleBounds = { startRow: rowIdx, endRow: rowIdx, startCol: 0, endCol: EXCEL_MAX_COL };
        this.rowStyleBoundsDirty = false;
      } else {
        this.rowStyleBounds.startRow = Math.min(this.rowStyleBounds.startRow, rowIdx);
        this.rowStyleBounds.endRow = Math.max(this.rowStyleBounds.endRow, rowIdx);
      }
      return;
    }

    if (beforeHad) {
      if (this.rowStyleIds.size === 0) {
        this.rowStyleBounds = null;
        this.rowStyleBoundsDirty = false;
      } else if (
        this.rowStyleBounds &&
        !this.rowStyleBoundsDirty &&
        (rowIdx === this.rowStyleBounds.startRow || rowIdx === this.rowStyleBounds.endRow)
      ) {
        this.rowStyleBoundsDirty = true;
      }
    }
  }

  /**
   * @param {number} col
   * @param {number} styleId
   */
  setColStyleId(col, styleId) {
    const colIdx = Number(col);
    if (!Number.isInteger(colIdx) || colIdx < 0) return;
    const nextStyle = Number(styleId);
    const afterStyleId = Number.isInteger(nextStyle) && nextStyle >= 0 ? nextStyle : 0;

    const beforeStyleId = this.colStyleIds.get(colIdx) ?? 0;
    if (beforeStyleId === afterStyleId) return;

    const beforeHad = beforeStyleId !== 0;
    const afterHas = afterStyleId !== 0;

    if (afterHas) {
      this.colStyleIds.set(colIdx, afterStyleId);
    } else {
      this.colStyleIds.delete(colIdx);
    }

    if (afterHas) {
      if (!this.colStyleBounds) {
        this.colStyleBounds = { startRow: 0, endRow: EXCEL_MAX_ROW, startCol: colIdx, endCol: colIdx };
        this.colStyleBoundsDirty = false;
      } else {
        this.colStyleBounds.startCol = Math.min(this.colStyleBounds.startCol, colIdx);
        this.colStyleBounds.endCol = Math.max(this.colStyleBounds.endCol, colIdx);
      }
      return;
    }

    if (beforeHad) {
      if (this.colStyleIds.size === 0) {
        this.colStyleBounds = null;
        this.colStyleBoundsDirty = false;
      } else if (
        this.colStyleBounds &&
        !this.colStyleBoundsDirty &&
        (colIdx === this.colStyleBounds.startCol || colIdx === this.colStyleBounds.endCol)
      ) {
        this.colStyleBoundsDirty = true;
      }
    }
  }

  /**
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  getRowStyleBounds() {
    if (this.rowStyleIds.size === 0) {
      this.rowStyleBounds = null;
      this.rowStyleBoundsDirty = false;
      return null;
    }
    if (this.rowStyleBoundsDirty || !this.rowStyleBounds) {
      this.__rowStyleBoundsRecomputeCount += 1;
      this.rowStyleBounds = this.#recomputeRowStyleBounds();
      this.rowStyleBoundsDirty = false;
    }
    return this.rowStyleBounds ? { ...this.rowStyleBounds } : null;
  }

  /**
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  getColStyleBounds() {
    if (this.colStyleIds.size === 0) {
      this.colStyleBounds = null;
      this.colStyleBoundsDirty = false;
      return null;
    }
    if (this.colStyleBoundsDirty || !this.colStyleBounds) {
      this.__colStyleBoundsRecomputeCount += 1;
      this.colStyleBounds = this.#recomputeColStyleBounds();
      this.colStyleBoundsDirty = false;
    }
    return this.colStyleBounds ? { ...this.colStyleBounds } : null;
  }

  /**
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  getContentBounds() {
    if (this.contentCellCount === 0) {
      this.contentBounds = null;
      this.contentBoundsDirty = false;
      return null;
    }
    if (!this.contentBounds) return null;
    if (this.contentBoundsDirty) {
      this.__contentBoundsRecomputeCount += 1;
      this.contentBounds = this.#recomputeBounds({ includeFormat: false });
      this.contentBoundsDirty = false;
    }
    return this.contentBounds ? { ...this.contentBounds } : null;
  }

  /**
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  getFormatBounds() {
    if (this.cells.size === 0) {
      this.formatBounds = null;
      this.formatBoundsDirty = false;
      return null;
    }
    if (!this.formatBounds) return null;
    if (this.formatBoundsDirty) {
      this.__formatBoundsRecomputeCount += 1;
      this.formatBounds = this.#recomputeBounds({ includeFormat: true });
      this.formatBoundsDirty = false;
    }
    return this.formatBounds ? { ...this.formatBounds } : null;
  }

  /**
   * Recompute bounds by scanning the sparse cell map.
   *
   * This is intentionally only used when a boundary cell is cleared (shrinking requires
   * discovering the next extreme), keeping `getUsedRange` amortized O(1).
   *
   * @param {{ includeFormat: boolean }} options
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  #recomputeBounds(options) {
    const includeFormat = options.includeFormat;

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;
    let hasData = false;

    for (const [key, cell] of this.cells.entries()) {
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
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  #recomputeRowStyleBounds() {
    let minRow = Infinity;
    let maxRow = -Infinity;
    for (const row of this.rowStyleIds.keys()) {
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
    }
    if (minRow === Infinity) return null;
    return { startRow: minRow, endRow: maxRow, startCol: 0, endCol: EXCEL_MAX_COL };
  }

  /**
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  #recomputeColStyleBounds() {
    let minCol = Infinity;
    let maxCol = -Infinity;
    for (const col of this.colStyleIds.keys()) {
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    }
    if (minCol === Infinity) return null;
    return { startRow: 0, endRow: EXCEL_MAX_ROW, startCol: minCol, endCol: maxCol };
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

    /**
     * Monotonic counters for downstream caching adapters.
     *
     * `updateVersion` increments after every successful `#applyEdits` (cell deltas or sheet-view deltas).
     * `contentVersion` increments only when workbook *content* changes (value/formula) or when the set
     * of sheets changes via `applyState`.
     *
     * These are kept separate so view-only interactions (frozen panes, row/col sizing) do not churn
     * AI workbook context caches.
     *
     * @type {number}
     */
    this._updateVersion = 0;
    /** @type {number} */
    this._contentVersion = 0;

    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
  }

  /**
   * Subscribe to controller events.
   *
   * Events:
   * - `change`: {
   *     deltas: CellDelta[],
   *     sheetViewDeltas: SheetViewDelta[],
   *     formatDeltas: FormatDelta[],
   *     // Preferred explicit delta streams for layered formatting.
   *     rowStyleDeltas: Array<{ sheetId: string, row: number, beforeStyleId: number, afterStyleId: number }>,
   *     colStyleDeltas: Array<{ sheetId: string, col: number, beforeStyleId: number, afterStyleId: number }>,
   *     sheetStyleDeltas: Array<{ sheetId: string, beforeStyleId: number, afterStyleId: number }>,
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
   * Monotonic version that increments after every successful workbook mutation (cell or sheet-view).
   *
   * Useful for coarse invalidation of UI layers that care about any update.
   *
   * @returns {number}
   */
  get updateVersion() {
    return this._updateVersion;
  }

  /**
   * Monotonic version that increments only when workbook content changes:
   * - at least one cell delta changes `value` or `formula` (format-only changes ignored)
   * - sheet ids are added/removed via `applyState`
   *
   * This is intended for AI context caching (schema + sampled data blocks).
   *
   * @returns {number}
   */
  get contentVersion() {
    return this._contentVersion;
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
   * Return the contributing style ids for a cell's effective formatting (layered formats).
   *
   * Tuple order matches the clipboard helper expectations:
   * `[sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId]`.
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {[number, number, number, number]}
   */
  getCellFormatStyleIds(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;

    // Ensure the sheet is materialized (DocumentController is lazily sheet-creating).
    const cell = this.model.getCell(sheetId, c.row, c.col);
    const sheet = this.model.sheets.get(sheetId);

    const sheetDefaultStyleId = sheet?.defaultStyleId ?? 0;
    const rowStyleId = sheet?.rowStyleIds.get(c.row) ?? 0;
    const colStyleId = sheet?.colStyleIds.get(c.col) ?? 0;
    const cellStyleId = cell.styleId ?? 0;

    return [sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId];
  }

  /**
   * @param {string} sheetId
   * @returns {number}
   */
  getSheetDefaultStyleId(sheetId) {
    // Ensure the sheet is materialized (DocumentController is lazily sheet-creating).
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    return sheet?.defaultStyleId ?? 0;
  }

  /**
   * @param {string} sheetId
   * @param {number} row
   * @returns {number}
   */
  getRowStyleId(sheetId, row) {
    const idx = Number(row);
    if (!Number.isInteger(idx) || idx < 0) return 0;
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    return sheet?.rowStyleIds.get(idx) ?? 0;
  }

  /**
   * @param {string} sheetId
   * @param {number} col
   * @returns {number}
   */
  getColStyleId(sheetId, col) {
    const idx = Number(col);
    if (!Number.isInteger(idx) || idx < 0) return 0;
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    return sheet?.colStyleIds.get(idx) ?? 0;
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

    const sheetStyle = this.styleTable.get(sheet?.defaultStyleId ?? 0);
    const colStyle = this.styleTable.get(sheet?.colStyleIds.get(c.col) ?? 0);
    const rowStyle = this.styleTable.get(sheet?.rowStyleIds.get(c.row) ?? 0);
    const cellStyle = this.styleTable.get(cell.styleId ?? 0);

    // Precedence: sheet < col < row < cell.
    const sheetCol = applyStylePatch(sheetStyle, colStyle);
    const sheetColRow = applyStylePatch(sheetCol, rowStyle);
    return applyStylePatch(sheetColRow, cellStyle);
  }

  /**
   * Return the set of style ids contributing to a cell's effective formatting.
   *
   * This is useful for callers that want to cache derived formatting (clipboard export,
   * render caches, etc) without needing to stringify full style objects.
   *
   * Tuple order is:
   * `[sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId]`.
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {[number, number, number, number]}
   */
  getCellFormatStyleIds(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const cell = this.model.getCell(sheetId, c.row, c.col);
    const sheet = this.model.sheets.get(sheetId);
    return [
      sheet?.defaultStyleId ?? 0,
      sheet?.rowStyleIds.get(c.row) ?? 0,
      sheet?.colStyleIds.get(c.col) ?? 0,
      cell.styleId ?? 0,
    ];
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
    const includeFormat = Boolean(options.includeFormat);
    if (!sheet) return null;

    // Default behavior: content bounds (value/formula only).
    if (!includeFormat) {
      return sheet.getContentBounds();
    }

    // includeFormat=true: formatting layers can apply to otherwise-empty cells without
    // materializing them in the sparse cell map. Incorporate all layers.

    // Sheet default formatting applies to every cell.
    if (sheet.defaultStyleId !== 0) {
      return { startRow: 0, endRow: EXCEL_MAX_ROW, startCol: 0, endCol: EXCEL_MAX_COL };
    }

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;
    let hasData = false;

    const mergeBounds = (bounds) => {
      if (!bounds) return;
      hasData = true;
      minRow = Math.min(minRow, bounds.startRow);
      minCol = Math.min(minCol, bounds.startCol);
      maxRow = Math.max(maxRow, bounds.endRow);
      maxCol = Math.max(maxCol, bounds.endCol);
    };

    mergeBounds(sheet.getColStyleBounds());
    mergeBounds(sheet.getRowStyleBounds());

    // Merge in bounds from stored cell states (values/formulas/cell-level format-only entries).
    mergeBounds(sheet.getFormatBounds());

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

    // Iterate only stored cells in the sheet (sparse map) instead of scanning the full rectangle.
    // This keeps clearRange O(#stored cells) rather than O(range area) for huge ranges.
    this.forEachCellInSheet(sheetId, ({ row, col, cell }) => {
      if (row < r.start.row || row > r.end.row) return;
      if (col < r.start.col || col > r.end.col) return;

      // Skip format-only cells (styleId-only) since clearing content would be a no-op.
      if (cell.value == null && cell.formula == null) return;

      const before = cloneCellState(cell);
      const after = { value: null, formula: null, styleId: before.styleId };
      if (cellStateEquals(before, after)) return;
      deltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
    });

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
    const isFullSheet = r.start.row === 0 && r.end.row === EXCEL_MAX_ROW && r.start.col === 0 && r.end.col === EXCEL_MAX_COL;
    const isFullHeightCols = r.start.row === 0 && r.end.row === EXCEL_MAX_ROW;
    const isFullWidthRows = r.start.col === 0 && r.end.col === EXCEL_MAX_COL;

    // Ensure sheet exists so we can mutate format layers without materializing cells.
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return;

    /** @type {CellDelta[]} */
    const cellDeltas = [];
    /** @type {FormatDelta[]} */
    const formatDeltas = [];

    /** @type {Map<number, number>} */
    const patchedStyleIdCache = new Map();
    const patchStyleId = (beforeStyleId) => {
      const cached = patchedStyleIdCache.get(beforeStyleId);
      if (cached != null) return cached;
      const baseStyle = this.styleTable.get(beforeStyleId);
      const merged = applyStylePatch(baseStyle, stylePatch);
      const afterStyleId = this.styleTable.intern(merged);
      patchedStyleIdCache.set(beforeStyleId, afterStyleId);
      return afterStyleId;
    };

    if (isFullSheet) {
      if (this.canEditCell && !this.canEditCell({ sheetId, row: 0, col: 0 })) return;

      // Patch sheet default.
      const beforeStyleId = sheet.defaultStyleId ?? 0;
      const afterStyleId = patchStyleId(beforeStyleId);
      if (beforeStyleId !== afterStyleId) {
        formatDeltas.push({ sheetId, layer: "sheet", beforeStyleId, afterStyleId });
      }

      // Patch existing row/col overrides so the new formatting wins over explicit values.
      for (const [row, rowBeforeStyleId] of sheet.rowStyleIds.entries()) {
        const rowAfterStyleId = patchStyleId(rowBeforeStyleId);
        if (rowBeforeStyleId === rowAfterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "row", index: row, beforeStyleId: rowBeforeStyleId, afterStyleId: rowAfterStyleId });
      }
      for (const [col, colBeforeStyleId] of sheet.colStyleIds.entries()) {
        const colAfterStyleId = patchStyleId(colBeforeStyleId);
        if (colBeforeStyleId === colAfterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "col", index: col, beforeStyleId: colBeforeStyleId, afterStyleId: colAfterStyleId });
      }

      // Patch existing cell overrides so the new formatting wins over explicit values.
      for (const [key, cell] of sheet.cells.entries()) {
        if (!cell || cell.styleId === 0) continue;
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) continue;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, { label: options.label });
      return;
    }

    if (isFullHeightCols) {
      if (this.canEditCell && !this.canEditCell({ sheetId, row: 0, col: r.start.col })) return;

      for (let col = r.start.col; col <= r.end.col; col++) {
        const beforeStyleId = sheet.colStyleIds.get(col) ?? 0;
        const afterStyleId = patchStyleId(beforeStyleId);
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "col", index: col, beforeStyleId, afterStyleId });
      }

      // Ensure the patch overrides explicit cell formatting (sparse overrides only).
      for (const [key, cell] of sheet.cells.entries()) {
        if (!cell || cell.styleId === 0) continue;
        const { row, col } = parseRowColKey(key);
        if (col < r.start.col || col > r.end.col) continue;
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) continue;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, { label: options.label });
      return;
    }

    if (isFullWidthRows) {
      if (this.canEditCell && !this.canEditCell({ sheetId, row: r.start.row, col: 0 })) return;

      for (let row = r.start.row; row <= r.end.row; row++) {
        const beforeStyleId = sheet.rowStyleIds.get(row) ?? 0;
        const afterStyleId = patchStyleId(beforeStyleId);
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "row", index: row, beforeStyleId, afterStyleId });
      }

      // Ensure the patch overrides explicit cell formatting (sparse overrides only).
      for (const [key, cell] of sheet.cells.entries()) {
        if (!cell || cell.styleId === 0) continue;
        const { row, col } = parseRowColKey(key);
        if (row < r.start.row || row > r.end.row) continue;
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) continue;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, { label: options.label });
      return;
    }

    // Fallback: sparse per-cell overrides.
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const before = this.model.getCell(sheetId, row, col);
        const afterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: afterStyleId };
        if (cellStateEquals(before, after)) continue;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }
    }

    this.#applyUserCellAndFormatDeltas(cellDeltas, [], { label: options.label });
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

    const beforeStyleId = sheet.defaultStyleId ?? 0;
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

    const beforeStyleId = sheet.rowStyleIds.get(rowIdx) ?? 0;
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

    const beforeStyleId = sheet.colStyleIds.get(colIdx) ?? 0;
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
      const effectiveStyle = this.#resolveEffectiveCellStyle(sheetId, row, col, cell.styleId);
      cells.set(semanticDiffCellKey(row, col), {
        value: cell.value ?? null,
        formula: cell.formula ?? null,
        // Semantic diff consumers expect the *effective* format for stored cells so inherited
        // row/col/sheet formatting is visible even when `styleId === 0`.
        format: effectiveStyle && Object.keys(effectiveStyle).length > 0 ? effectiveStyle : null,
      });
    }
    return { cells };
  }

  /**
   * Resolve a cell's effective formatting, taking layered formatting into account:
   * sheet -> column -> row -> range-runs -> cell.
   *
   * This helper is intentionally resilient to schema evolution: it looks for common
   * field names used by layered-formatting implementations (Task 44 / Task 118) and
   * falls back to the legacy per-cell `styleId` only model when those structures are
   * absent.
   *
   * @param {string} sheetId
   * @param {number} row
   * @param {number} col
   * @param {number} cellStyleId
   * @returns {Record<string, any>}
   */
  #resolveEffectiveCellStyle(sheetId, row, col, cellStyleId) {
    const sheet = this.model.sheets.get(sheetId);

    const layerToStyle = (layer) => {
      if (layer == null) return null;
      if (typeof layer === "number") {
        if (!Number.isFinite(layer) || layer === 0) return null;
        return this.styleTable.get(layer);
      }
      if (typeof layer === "object") return layer;
      return null;
    };

    const merge = (base, layer) => {
      const patch = layerToStyle(layer);
      if (!patch) return base;
      if (typeof patch !== "object") return base;
      if (Object.keys(patch).length === 0) return base;
      return applyStylePatch(base, patch);
    };

    const axisLookup = (container, index) => {
      if (!container) return null;
      // Support either Map-like or plain object encodings.
      if (typeof container.get === "function") {
        return container.get(index) ?? container.get(String(index)) ?? null;
      }
      if (typeof container === "object") {
        return container[String(index)] ?? null;
      }
      return null;
    };

    /** @type {Record<string, any>} */
    let style = {};

    if (sheet) {
      // Sheet default.
      style = merge(
        style,
        sheet.sheetStyleId ??
          sheet.sheetFormatId ??
          sheet.defaultStyleId ??
          sheet.sheetDefaultStyleId ??
          sheet.format ??
          sheet.sheetFormat ??
          null,
      );

      // Column default.
      const colLayer =
        axisLookup(sheet.colStyleIds, col) ??
        axisLookup(sheet.colFormatIds, col) ??
        axisLookup(sheet.colFormats, col) ??
        axisLookup(sheet.columnStyleIds, col) ??
        axisLookup(sheet.columnFormats, col) ??
        null;
      style = merge(style, colLayer);

      // Row default.
      const rowLayer =
        axisLookup(sheet.rowStyleIds, row) ??
        axisLookup(sheet.rowFormatIds, row) ??
        axisLookup(sheet.rowFormats, row) ??
        axisLookup(sheet.rowsStyleIds, row) ??
        axisLookup(sheet.rowsFormats, row) ??
        null;
      style = merge(style, rowLayer);

      // Range runs (Task 118). Expected to be sparse and reasonably small.
      const runs =
        sheet.formatRuns ??
        sheet.rangeFormatRuns ??
        sheet.rangeRuns ??
        sheet.formattingRuns ??
        sheet.formatRanges ??
        null;
      if (Array.isArray(runs)) {
        for (const run of runs) {
          if (!run || typeof run !== "object") continue;
          const startRow = Number(run.startRow ?? run.start?.row ?? run.sr);
          const startCol = Number(run.startCol ?? run.start?.col ?? run.sc);
          const endRow = Number(run.endRow ?? run.end?.row ?? run.er);
          const endCol = Number(run.endCol ?? run.end?.col ?? run.ec);
          if (!Number.isInteger(startRow) || !Number.isInteger(startCol)) continue;
          if (!Number.isInteger(endRow) || !Number.isInteger(endCol)) continue;
          if (row < startRow || row > endRow || col < startCol || col > endCol) continue;
          style = merge(style, run.styleId ?? run.format ?? run.style ?? null);
        }
      }
    }

    // Cell override (legacy cell-level formatting).
    style = merge(style, cellStyleId);

    return style;
  }

  /**
   * Encode the document's current cell inputs as a snapshot suitable for the VersionManager.
   *
   * Undo/redo history is intentionally *not* included; snapshots represent workbook contents.
   *
   * @returns {Uint8Array}
   */
  encodeState() {
    // Preserve sheet insertion order so sheet tab reordering can survive snapshot roundtrips.
    // (Sorting here would destroy workbook navigation order.)
    const sheetIds = Array.from(this.model.sheets.keys());
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
      out.defaultFormat = sheet && sheet.defaultStyleId !== 0 ? this.styleTable.get(sheet.defaultStyleId) : null;

      const rowFormats = Array.from(sheet?.rowStyleIds?.entries?.() ?? []).map(([row, styleId]) => ({
        row,
        format: styleId === 0 ? null : this.styleTable.get(styleId),
      }));
      rowFormats.sort((a, b) => a.row - b.row);
      out.rowFormats = rowFormats;

      const colFormats = Array.from(sheet?.colStyleIds?.entries?.() ?? []).map(([col, styleId]) => ({
        col,
        format: styleId === 0 ? null : this.styleTable.get(styleId),
      }));
      colFormats.sort((a, b) => a.col - b.col);
      out.colFormats = colFormats;
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
    /** @type {Map<string, { defaultStyleId: number, rowStyleIds: Map<number, number>, colStyleIds: Map<number, number> }>} */
    const nextFormats = new Map();

    const normalizeFormatOverrides = (raw, axisKey) => {
      /** @type {Map<number, number>} */
      const out = new Map();
      if (!raw) return out;

      if (Array.isArray(raw)) {
        for (const entry of raw) {
          const index = Array.isArray(entry) ? entry[0] : entry?.index ?? entry?.[axisKey] ?? entry?.row ?? entry?.col;
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

      const defaultFormat = sheet?.defaultFormat ?? sheet?.sheetFormat ?? null;
      const defaultStyleId = defaultFormat == null ? 0 : this.styleTable.intern(defaultFormat);
      nextFormats.set(sheet.id, {
        defaultStyleId,
        rowStyleIds: normalizeFormatOverrides(sheet?.rowFormats, "row"),
        colStyleIds: normalizeFormatOverrides(sheet?.colFormats, "col"),
      });
    }

    const existingSheetIds = new Set(this.model.sheets.keys());
    const nextSheetIds = new Set(nextSheets.keys());
    const allSheetIds = new Set([...existingSheetIds, ...nextSheetIds]);
    const removedSheetIds = Array.from(existingSheetIds).filter((id) => !nextSheetIds.has(id));
    const addedSheetIds = Array.from(nextSheetIds).filter((id) => !existingSheetIds.has(id));
    const sheetStructureChanged = removedSheetIds.length > 0 || addedSheetIds.length > 0;

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
      const beforeSheetStyleId = existingSheet?.defaultStyleId ?? 0;
      const beforeRowStyles = existingSheet?.rowStyleIds ?? new Map();
      const beforeColStyles = existingSheet?.colStyleIds ?? new Map();

      const next = nextFormats.get(sheetId);
      const afterSheetStyleId = next?.defaultStyleId ?? 0;
      const afterRowStyles = next?.rowStyleIds ?? new Map();
      const afterColStyles = next?.colStyleIds ?? new Map();

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
    this.#applyEdits(deltas, sheetViewDeltas, formatDeltas, {
      recalc: false,
      emitChange: true,
      source: "applyState",
      sheetStructureChanged,
    });
    this.engine?.endBatch?.();
    this.engine?.recalculate();

    for (const sheetId of removedSheetIds) {
      this.model.sheets.delete(sheetId);
    }

    // Match the sheet Map iteration order to the snapshot ordering so sheet tab order
    // roundtrips through encodeState/applyState (including when restoring onto an
    // existing DocumentController instance).
    const orderedSheetIds = Array.from(nextSheets.keys());
    if (orderedSheetIds.length > 0) {
      for (const sheetId of orderedSheetIds) {
        const sheet = this.model.sheets.get(sheetId);
        if (!sheet) continue;
        // Re-insert to update insertion order without changing sheet identity.
        this.model.sheets.delete(sheetId);
        this.model.sheets.set(sheetId, sheet);
      }
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
    this.#applyEdits([], deltas, [], { recalc: false, emitChange: true, source });

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
    const shouldRecalc = cellDeltasAffectRecalc(Array.from(batch.deltasByCell.values()));
    if (shouldRecalc) {
      this.engine?.recalculate();
      // Emit a follow-up change so observers know formula results may have changed.
      this.#emit("change", {
        deltas: [],
        sheetViewDeltas: [],
        formatDeltas: [],
        rowStyleDeltas: [],
        colStyleDeltas: [],
        sheetStyleDeltas: [],
        source: "endBatch",
        recalc: true,
      });
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
    const shouldRecalc = Boolean(batch && cellDeltasAffectRecalc(Array.from(batch.deltasByCell.values())));
    if (shouldRecalc) {
      this.engine?.recalculate();
      this.#emit("change", {
        deltas: [],
        sheetViewDeltas: [],
        formatDeltas: [],
        rowStyleDeltas: [],
        colStyleDeltas: [],
        sheetStyleDeltas: [],
        source: "cancelBatch",
        recalc: true,
      });
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

    const shouldRecalc = cellDeltasAffectRecalc(cellDeltas);
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

    const shouldRecalc = cellDeltasAffectRecalc(cellDeltas);
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
    const shouldRecalc = this.batchDepth === 0 && cellDeltasAffectRecalc(deltas);
    this.#applyEdits(deltas, [], [], { recalc: shouldRecalc, emitChange: true, source });

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
   * Apply a set of cell deltas and format deltas as a single user edit (one change event / one
   * undo step), merging into an active batch if present.
   *
   * This is primarily used for range formatting operations that need to update both the
   * sheet/row/col layers and a sparse set of cell overrides.
   *
   * @param {CellDelta[]} cellDeltas
   * @param {FormatDelta[]} formatDeltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, options) {
    cellDeltas = Array.isArray(cellDeltas) ? cellDeltas : [];
    formatDeltas = Array.isArray(formatDeltas) ? formatDeltas : [];

    if (cellDeltas.length > 0 && this.canEditCell) {
      cellDeltas = cellDeltas.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
    }

    if (cellDeltas.length === 0 && formatDeltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    const shouldRecalc = this.batchDepth === 0 && cellDeltasAffectRecalc(cellDeltas);
    this.#applyEdits(cellDeltas, [], formatDeltas, { recalc: shouldRecalc, emitChange: true, source });

    if (this.batchDepth > 0) {
      if (cellDeltas.length > 0) this.#mergeIntoBatch(cellDeltas);
      if (formatDeltas.length > 0) this.#mergeFormatIntoBatch(formatDeltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(cellDeltas, [], formatDeltas, options);
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
    const shouldBumpContentVersion =
      Boolean(options?.sheetStructureChanged) || cellDeltas.some((d) => !cellContentEquals(d.before, d.after));

    // Apply to the canonical model first.
    for (const delta of formatDeltas) {
      // Ensure sheet exists for format-only changes.
      this.model.getCell(delta.sheetId, 0, 0);
      const sheet = this.model.sheets.get(delta.sheetId);
      if (!sheet) continue;
      if (delta.layer === "sheet") {
        sheet.defaultStyleId = delta.afterStyleId;
        continue;
      }
      const index = delta.index;
      if (index == null) continue;
      if (delta.layer === "row") {
        sheet.setRowStyleId(index, delta.afterStyleId);
        continue;
      }
      if (delta.layer === "col") {
        sheet.setColStyleId(index, delta.afterStyleId);
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
          sheet.defaultStyleId = delta.beforeStyleId;
          continue;
        }
        const index = delta.index;
        if (index == null) continue;
        if (delta.layer === "row") {
          sheet.setRowStyleId(index, delta.beforeStyleId);
          continue;
        }
        if (delta.layer === "col") {
          sheet.setColStyleId(index, delta.beforeStyleId);
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

    // Update versions before emitting events so observers can synchronously read the latest value.
    this._updateVersion += 1;
    if (shouldBumpContentVersion) this._contentVersion += 1;

    if (options.emitChange) {
      /** @type {any[]} */
      const sheetStyleDeltas = [];
      /** @type {any[]} */
      const rowStyleDeltas = [];
      /** @type {any[]} */
      const colStyleDeltas = [];
      for (const delta of formatDeltas) {
        if (!delta) continue;
        if (delta.layer === "sheet") {
          sheetStyleDeltas.push({
            sheetId: delta.sheetId,
            beforeStyleId: delta.beforeStyleId,
            afterStyleId: delta.afterStyleId,
          });
          continue;
        }
        const index = delta.index;
        if (index == null) continue;
        if (delta.layer === "row") {
          rowStyleDeltas.push({
            sheetId: delta.sheetId,
            row: index,
            beforeStyleId: delta.beforeStyleId,
            afterStyleId: delta.afterStyleId,
          });
          continue;
        }
        if (delta.layer === "col") {
          colStyleDeltas.push({
            sheetId: delta.sheetId,
            col: index,
            beforeStyleId: delta.beforeStyleId,
            afterStyleId: delta.afterStyleId,
          });
        }
      }

      const payload = {
        deltas: cellDeltas.map(cloneDelta),
        sheetViewDeltas: sheetViewDeltas.map(cloneSheetViewDelta),
        formatDeltas: formatDeltas.map(cloneFormatDelta),
        // Preferred explicit delta streams for row/col/sheet formatting.
        rowStyleDeltas,
        colStyleDeltas,
        sheetStyleDeltas,
        recalc: options.recalc,
      };
      if (options.source) payload.source = options.source;
      this.#emit("change", payload);
    }

    this.#emit("update", {});
  }
}
