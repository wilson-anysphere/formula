import {
  cellStateEquals,
  cloneCellState,
  emptyCellState,
} from "./cell.js";
import { formatA1, normalizeRange, parseA1, parseRangeA1 } from "./coords.js";
import { applyStylePatch, StyleTable } from "../formatting/styleTable.js";
import { getSheetNameValidationErrorMessage } from "../../../../packages/workbook-backend/src/sheetNameValidation.js";

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

function rangeRunKey(sheetId, col) {
  return `${sheetId}:rangeRun:${col}`;
}

function sortKey(sheetId, row, col) {
  return `${sheetId}\u0000${row.toString().padStart(10, "0")}\u0000${col
    .toString()
    .padStart(10, "0")}`;
}

function normalizeSheetNameForCaseInsensitiveCompare(name) {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics we use elsewhere (`@formula/workbook-backend`).
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

function parseRowColKey(key) {
  const comma = key.indexOf(",");
  // This format is internal-only, but keep validation to avoid corrupting state silently.
  if (comma === -1) {
    throw new Error(`Invalid cell key: ${key}`);
  }

  // Preserve the historical `split(",")` semantics: ignore any additional commas.
  const secondComma = key.indexOf(",", comma + 1);
  const rowStr = key.slice(0, comma);
  const colStr = secondComma === -1 ? key.slice(comma + 1) : key.slice(comma + 1, secondComma);

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
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function encodeBase64(bytes) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  if (typeof btoa === "function") {
    // Chunk to avoid call stack limits for large images.
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }
  throw new Error("Base64 encoding is not supported in this environment");
}

/**
 * @param {string} base64
 * @returns {Uint8Array}
 */
function decodeBase64(base64) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Uint8Array.from(Buffer.from(base64, "base64"));
  if (typeof atob === "function") {
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  }
  throw new Error("Base64 decoding is not supported in this environment");
}

/**
 * Best-effort MIME type inference for workbook images.
 *
 * Many code paths provide a mime type (e.g. file import, XLSX load). Keep this lightweight and
 * conservative so callers can omit it without breaking rendering.
 *
 * @param {string} imageId
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function inferMimeTypeForImage(imageId, bytes) {
  const ext = String(imageId ?? "").split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      break;
  }

  if (bytes && bytes.length >= 4) {
    // PNG
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
    // JPEG
    if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
    // GIF
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) return "image/gif";
    // BMP
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
    // WebP: "RIFF"...."WEBP"
    if (
      bytes.length >= 12 &&
      bytes[0] === 0x52 &&
      bytes[1] === 0x49 &&
      bytes[2] === 0x46 &&
      bytes[3] === 0x46 &&
      bytes[8] === 0x57 &&
      bytes[9] === 0x45 &&
      bytes[10] === 0x42 &&
      bytes[11] === 0x50
    ) {
      return "image/webp";
    }
  }

  return "application/octet-stream";
}

/**
 * Deep-clone a JSON-serializable value.
 *
 * @template T
 * @param {T} value
 * @returns {T}
 */
function cloneJsonSerializable(value) {
  return JSON.parse(JSON.stringify(value));
}

/**
 * @param {any} a
 * @param {any} b
 * @returns {boolean}
 */
function jsonSerializableEquals(a, b) {
  if (a === b) return true;
  try {
    return JSON.stringify(a) === JSON.stringify(b);
  } catch {
    return false;
  }
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Strict "JSON object" check used for persisted drawing schemas.
 *
 * This intentionally rejects class instances like Map/Set/Date to avoid lossy snapshot encodes.
 *
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isJsonObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/**
 * Remove empty nested objects from a style tree.
 *
 * This matches the normalization used by versioning snapshot parsers so semantic diffs
 * don't treat `{ font: {} }` as meaningfully different from `{}`.
 *
 * @param {any} value
 * @returns {any}
 */
function pruneEmptyObjects(value) {
  if (!isPlainObject(value)) return value;
  /** @type {Record<string, any>} */
  const out = {};
  for (const [key, raw] of Object.entries(value)) {
    if (raw === undefined) continue;
    const pruned = pruneEmptyObjects(raw);
    if (isPlainObject(pruned) && Object.keys(pruned).length === 0) continue;
    out[key] = pruned;
  }
  return out;
}

/**
 * Normalize an effective style object for semantic diff exports.
 *
 * - Prunes empty nested objects (e.g. `{ font: {} }`).
 * - Treats empty styles as `null` for backwards compatibility.
 *
 * @param {any} style
 * @returns {any | null}
 */
function normalizeSemanticDiffFormat(style) {
  if (!isPlainObject(style)) return null;
  const pruned = pruneEmptyObjects(style);
  if (!isPlainObject(pruned) || Object.keys(pruned).length === 0) return null;
  return pruned;
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

// Above this number of cells, `setRangeFormat` will use the compressed range-run formatting
// layer instead of enumerating every cell in the rectangle.
const RANGE_RUN_FORMAT_THRESHOLD = 50_000;

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
 *    * Merged-cell regions for this sheet (Excel-style).
 *    *
 *    * Ranges use inclusive end coordinates (`endRow`/`endCol` are inclusive) and
 *    * are anchored at the top-left cell (`startRow`/`startCol`).
 *    *
 *    * This lives on the sheet view state so it can reuse the existing
 *    * undo/redo + snapshot plumbing for sheet view deltas.
 *    *\/
 *   mergedRanges?: Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>,
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
 * Excel-style worksheet visibility.
 *
 * @typedef {"visible" | "hidden" | "veryHidden"} SheetVisibility
 */

/**
 * Excel-style tab color, round-trippable through XLSX.
 *
 * Keep in sync with `apps/desktop/src/sheets/workbookSheetStore.ts` and
 * `apps/desktop/src/workbook/workbook.ts`.
 *
 * @typedef {{
 *   rgb?: string,
 *   theme?: number,
 *   indexed?: number,
 *   tint?: number,
 *   auto?: boolean,
 * }} TabColor
 */

/**
 * Sheet metadata tracked by the controller and persisted through `encodeState`.
 *
 * Note: the sheet id is stored separately (as the key in `DocumentController.sheetMeta`).
 *
 * @typedef {{
 *   name: string,
 *   visibility: SheetVisibility,
 *   tabColor?: TabColor,
 * }} SheetMetaState
 */

/**
 * Sheet metadata delta (undoable).
 *
 * `before`/`after` are `null` for add/delete operations respectively.
 *
 * @typedef {{
 *   sheetId: string,
 *   before: SheetMetaState | null,
 *   after: SheetMetaState | null,
 * }} SheetMetaDelta
 */

/**
 * Sheet order delta (undoable).
 *
 * @typedef {{
 *   before: string[],
 *   after: string[],
 * }} SheetOrderDelta
 */

/**
 * Workbook-scoped image store entry.
 *
 * `bytes` should be the original binary payload (PNG/JPEG/etc) and is treated as opaque by the controller.
 *
 * @typedef {{
 *   bytes: Uint8Array,
 *   mimeType?: string | null,
 * }} ImageEntry
 */

/**
 * Image store delta (undoable).
 *
 * `before`/`after` are `null` for add/delete operations respectively.
 *
 * @typedef {{
 *   imageId: string,
 *   before: ImageEntry | null,
 *   after: ImageEntry | null,
 * }} ImageDelta
 */

/**
 * Drawings delta (undoable).
 *
 * This controller treats drawing entries as opaque JSON-serializable objects. Consumers should use a
 * stable schema (aligned with `formula-model` when possible).
 *
 * @typedef {{
 *   sheetId: string,
 *   before: any[],
 *   after: any[],
 * }} DrawingDelta
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

  const normalizeMergedRanges = (raw) => {
    if (!raw) return null;

    // Accept:
    // - Array<{startRow,endRow,startCol,endCol}>
    // - Array<{start:{row,col},end:{row,col}}>
    // - { regions: Array<{ range: {...} }> } (formula-model shape)
    /** @type {any[]} */
    const entries = (() => {
      if (Array.isArray(raw)) return raw;
      if (typeof raw === "object" && Array.isArray(raw?.regions)) {
        return raw.regions.map((r) => r?.range ?? r);
      }
      return [];
    })();

    /** @type {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>} */
    const out = [];
    const seen = new Set();

    for (const entry of entries) {
      if (!entry) continue;
      const range = entry?.range ?? entry;
      const startRowNum = Number(range?.startRow ?? range?.start?.row);
      const endRowNum = Number(range?.endRow ?? range?.end?.row);
      const startColNum = Number(range?.startCol ?? range?.start?.col);
      const endColNum = Number(range?.endCol ?? range?.end?.col);
      if (!Number.isInteger(startRowNum) || startRowNum < 0) continue;
      if (!Number.isInteger(endRowNum) || endRowNum < 0) continue;
      if (!Number.isInteger(startColNum) || startColNum < 0) continue;
      if (!Number.isInteger(endColNum) || endColNum < 0) continue;

      const startRow = Math.min(startRowNum, endRowNum);
      const endRow = Math.max(startRowNum, endRowNum);
      const startCol = Math.min(startColNum, endColNum);
      const endCol = Math.max(startColNum, endColNum);

      // Ignore single-cell merges (no-op).
      if (startRow === endRow && startCol === endCol) continue;

      const key = `${startRow},${startCol},${endRow},${endCol}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ startRow, endRow, startCol, endCol });
    }

    out.sort((a, b) => (a.startRow - b.startRow === 0 ? a.startCol - b.startCol : a.startRow - b.startRow));
    return out.length === 0 ? null : out;
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
  const mergedRanges = normalizeMergedRanges(
    view?.mergedRanges ?? view?.merged_ranges ?? view?.mergedRegions ?? view?.merged_regions,
  );

  return {
    frozenRows: normalizeFrozenCount(view?.frozenRows),
    frozenCols: normalizeFrozenCount(view?.frozenCols),
    ...(mergedRanges ? { mergedRanges } : {}),
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
  if (Array.isArray(view.mergedRanges)) {
    next.mergedRanges = view.mergedRanges.map((r) => ({
      startRow: r.startRow,
      endRow: r.endRow,
      startCol: r.startCol,
      endCol: r.endCol,
    }));
  }
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

  const mergedRangesEqual = (left, right) => {
    const la = Array.isArray(left) ? left : [];
    const ra = Array.isArray(right) ? right : [];
    if (la.length !== ra.length) return false;
    for (let i = 0; i < la.length; i += 1) {
      const l = la[i];
      const r = ra[i];
      if (!l || !r) return false;
      if (l.startRow !== r.startRow) return false;
      if (l.endRow !== r.endRow) return false;
      if (l.startCol !== r.startCol) return false;
      if (l.endCol !== r.endCol) return false;
    }
    return true;
  };

  return (
    a.frozenRows === b.frozenRows &&
    a.frozenCols === b.frozenCols &&
    mergedRangesEqual(a.mergedRanges, b.mergedRanges) &&
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
 * `sheet < col < row < range-run < cell`.
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
 * A compressed formatting segment for a column in a sheet.
 *
 * The segment covers the half-open row interval `[startRow, endRowExclusive)`.
 *
 * `styleId` references an entry in the document's `StyleTable` and represents a patch that
 * participates in `getCellFormat()` merge semantics.
 *
 * Runs are:
 * - non-overlapping
 * - sorted by `startRow`
 * - stored only for non-default styles (`styleId !== 0`)
 *
 * @typedef {{
 *   startRow: number,
 *   endRowExclusive: number,
 *   styleId: number,
 * }} FormatRun
 */

/**
 * Deltas for edits to the range-run formatting layer.
 *
 * These are tracked per-column, since the underlying storage is `sheet.formatRunsByCol`.
 *
 * @typedef {{
 *   sheetId: string,
 *   col: number,
 *   /**
 *    * Inclusive start row of the union of all updates captured in this delta.
 *    *\/
 *   startRow: number,
 *   /**
 *    * Exclusive end row of the union of all updates captured in this delta.
 *    *\/
 *   endRowExclusive: number,
 *   beforeRuns: FormatRun[],
 *   afterRuns: FormatRun[],
 * }} RangeRunDelta
 */

/**
 * @typedef {{
 *   label?: string,
 *   mergeKey?: string,
 *   timestamp: number,
 *   deltasByCell: Map<string, CellDelta>,
 *   deltasBySheetView: Map<string, SheetViewDelta>,
 *   deltasByFormat: Map<string, FormatDelta>,
 *   deltasByRangeRun: Map<string, RangeRunDelta>,
 *   deltasByDrawing: Map<string, DrawingDelta>,
 *   deltasByImage: Map<string, ImageDelta>,
 *   deltasBySheetMeta: Map<string, SheetMetaDelta>,
 *   sheetOrderDelta: SheetOrderDelta | null,
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
 * @param {FormatRun} run
 * @returns {FormatRun}
 */
function cloneFormatRun(run) {
  return {
    startRow: run.startRow,
    endRowExclusive: run.endRowExclusive,
    styleId: run.styleId,
  };
}

/**
 * @param {RangeRunDelta} delta
 * @returns {RangeRunDelta}
 */
function cloneRangeRunDelta(delta) {
  return {
    sheetId: delta.sheetId,
    col: delta.col,
    startRow: delta.startRow,
    endRowExclusive: delta.endRowExclusive,
    beforeRuns: Array.isArray(delta.beforeRuns) ? delta.beforeRuns.map(cloneFormatRun) : [],
    afterRuns: Array.isArray(delta.afterRuns) ? delta.afterRuns.map(cloneFormatRun) : [],
  };
}

/**
 * Clone a tab color payload.
 *
 * Note: Some snapshot producers (e.g. BranchService) represent tab colors as an ARGB
 * string. Be tolerant and normalize strings into `{ rgb }` objects so downstream
 * consumers can treat `TabColor` as an object shape.
 *
 * @param {any} color
 * @returns {TabColor | undefined}
 */
function cloneTabColor(color) {
  if (!color) return undefined;
  if (typeof color === "string") {
    return { rgb: color.toUpperCase() };
  }
  if (typeof color === "object") {
    const out = { ...color };
    if (typeof out.rgb === "string") out.rgb = out.rgb.toUpperCase();
    return out;
  }
  return undefined;
}

/**
 * @param {SheetMetaState} meta
 * @returns {SheetMetaState}
 */
function cloneSheetMetaState(meta) {
  const out = { name: meta.name, visibility: meta.visibility };
  if (meta.tabColor) out.tabColor = cloneTabColor(meta.tabColor);
  return out;
}

/**
 * @param {SheetMetaDelta} delta
 * @returns {SheetMetaDelta}
 */
function cloneSheetMetaDelta(delta) {
  return {
    sheetId: delta.sheetId,
    before: delta.before ? cloneSheetMetaState(delta.before) : null,
    after: delta.after ? cloneSheetMetaState(delta.after) : null,
  };
}

/**
 * @param {SheetOrderDelta | null | undefined} delta
 * @returns {SheetOrderDelta | null}
 */
function cloneSheetOrderDelta(delta) {
  if (!delta) return null;
  return { before: Array.isArray(delta.before) ? delta.before.slice() : [], after: Array.isArray(delta.after) ? delta.after.slice() : [] };
}

/**
 * @param {Uint8Array | null | undefined} a
 * @param {Uint8Array | null | undefined} b
 * @returns {boolean}
 */
function bytesEqual(a, b) {
  if (a === b) return true;
  const al = a ? a.length : 0;
  const bl = b ? b.length : 0;
  if (al !== bl) return false;
  if (!a || !b) return false;
  for (let i = 0; i < al; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * @param {ImageEntry} entry
 * @returns {ImageEntry}
 */
function cloneImageEntry(entry) {
  const bytes = entry?.bytes instanceof Uint8Array ? entry.bytes.slice() : new Uint8Array();
  /** @type {ImageEntry} */
  const out = { bytes };
  if (entry && "mimeType" in entry) out.mimeType = entry.mimeType ?? null;
  return out;
}

/**
 * @param {ImageEntry | null | undefined} a
 * @param {ImageEntry | null | undefined} b
 * @returns {boolean}
 */
function imageEntryEquals(a, b) {
  if (a === b) return true;
  if (!a || !b) return a == null && b == null;
  const aMime = "mimeType" in a ? a.mimeType ?? null : null;
  const bMime = "mimeType" in b ? b.mimeType ?? null : null;
  return aMime === bMime && bytesEqual(a.bytes, b.bytes);
}

/**
 * @param {ImageDelta} delta
 * @returns {ImageDelta}
 */
function cloneImageDelta(delta) {
  return {
    imageId: delta.imageId,
    before: delta.before ? cloneImageEntry(delta.before) : null,
    after: delta.after ? cloneImageEntry(delta.after) : null,
  };
}

/**
 * @param {DrawingDelta} delta
 * @returns {DrawingDelta}
 */
function cloneDrawingDelta(delta) {
  return {
    sheetId: delta.sheetId,
    before: Array.isArray(delta.before) ? cloneJsonSerializable(delta.before) : [],
    after: Array.isArray(delta.after) ? cloneJsonSerializable(delta.after) : [],
  };
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
 * @param {HistoryEntry} entry
 * @returns {RangeRunDelta[]}
 */
function entryRangeRunDeltas(entry) {
  const deltas = Array.from(entry.deltasByRangeRun.values()).map(cloneRangeRunDelta);
  deltas.sort((a, b) => {
    if (a.sheetId !== b.sheetId) return a.sheetId < b.sheetId ? -1 : 1;
    if (a.col !== b.col) return a.col - b.col;
    if (a.startRow !== b.startRow) return a.startRow - b.startRow;
    return a.endRowExclusive - b.endRowExclusive;
  });
  return deltas;
}

/**
 * @param {HistoryEntry} entry
 * @returns {SheetMetaDelta[]}
 */
function entrySheetMetaDeltas(entry) {
  const deltas = Array.from(entry.deltasBySheetMeta.values()).map(cloneSheetMetaDelta);
  deltas.sort((a, b) => (a.sheetId < b.sheetId ? -1 : a.sheetId > b.sheetId ? 1 : 0));
  return deltas;
}

/**
 * @param {HistoryEntry} entry
 * @returns {DrawingDelta[]}
 */
function entryDrawingDeltas(entry) {
  const deltas = Array.from(entry.deltasByDrawing.values()).map(cloneDrawingDelta);
  deltas.sort((a, b) => (a.sheetId < b.sheetId ? -1 : a.sheetId > b.sheetId ? 1 : 0));
  return deltas;
}

/**
 * @param {HistoryEntry} entry
 * @returns {ImageDelta[]}
 */
function entryImageDeltas(entry) {
  const deltas = Array.from(entry.deltasByImage.values()).map(cloneImageDelta);
  deltas.sort((a, b) => (a.imageId < b.imageId ? -1 : a.imageId > b.imageId ? 1 : 0));
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

/**
 * @param {TabColor | undefined | null} a
 * @param {TabColor | undefined | null} b
 * @returns {boolean}
 */
function tabColorEquals(a, b) {
  if (a === b) return true;
  if (!a || !b) return a == null && b == null;
  return (
    (a.rgb ?? null) === (b.rgb ?? null) &&
    (a.theme ?? null) === (b.theme ?? null) &&
    (a.indexed ?? null) === (b.indexed ?? null) &&
    (a.tint ?? null) === (b.tint ?? null) &&
    (a.auto ?? null) === (b.auto ?? null)
  );
}

/**
 * @param {SheetMetaState | null | undefined} a
 * @param {SheetMetaState | null | undefined} b
 * @returns {boolean}
 */
function sheetMetaStateEquals(a, b) {
  if (a === b) return true;
  if (!a || !b) return a == null && b == null;
  return a.name === b.name && a.visibility === b.visibility && tabColorEquals(a.tabColor, b.tabColor);
}

/**
 * `SheetMetaDelta`s can affect recalculation when they add/remove sheets, since formulas
 * may refer to a sheet that now exists/doesn't exist.
 *
 * @param {SheetMetaDelta[]} deltas
 * @returns {boolean}
 */
function sheetMetaDeltasAffectRecalc(deltas) {
  for (const d of deltas) {
    if (!d) continue;
    if (d.before == null || d.after == null) return true;
  }
  return false;
}

/**
 * @param {RangeRunDelta[]} deltas
 * @returns {RangeRunDelta[]}
 */
function invertRangeRunDeltas(deltas) {
  return deltas.map((d) => ({
    sheetId: d.sheetId,
    col: d.col,
    startRow: d.startRow,
    endRowExclusive: d.endRowExclusive,
    beforeRuns: Array.isArray(d.afterRuns) ? d.afterRuns.map(cloneFormatRun) : [],
    afterRuns: Array.isArray(d.beforeRuns) ? d.beforeRuns.map(cloneFormatRun) : [],
  }));
}

/**
 * @param {SheetMetaDelta[]} deltas
 * @returns {SheetMetaDelta[]}
 */
function invertSheetMetaDeltas(deltas) {
  return deltas.map((d) => ({
    sheetId: d.sheetId,
    before: d.after ? cloneSheetMetaState(d.after) : null,
    after: d.before ? cloneSheetMetaState(d.before) : null,
  }));
}

/**
 * @param {DrawingDelta[]} deltas
 * @returns {DrawingDelta[]}
 */
function invertDrawingDeltas(deltas) {
  return deltas.map((d) => ({
    sheetId: d.sheetId,
    before: Array.isArray(d.after) ? cloneJsonSerializable(d.after) : [],
    after: Array.isArray(d.before) ? cloneJsonSerializable(d.before) : [],
  }));
}

/**
 * @param {ImageDelta[]} deltas
 * @returns {ImageDelta[]}
 */
function invertImageDeltas(deltas) {
  return deltas.map((d) => ({
    imageId: d.imageId,
    before: d.after ? cloneImageEntry(d.after) : null,
    after: d.before ? cloneImageEntry(d.before) : null,
  }));
}

/**
 * @param {SheetOrderDelta | null | undefined} delta
 * @returns {SheetOrderDelta | null}
 */
function invertSheetOrderDelta(delta) {
  if (!delta) return null;
  return { before: Array.isArray(delta.after) ? delta.after.slice() : [], after: Array.isArray(delta.before) ? delta.before.slice() : [] };
}

/**
 * @param {FormatRun[] | undefined | null} runs
 * @param {number} row
 * @returns {number}
 */
function styleIdForRowInRuns(runs, row) {
  if (!runs || runs.length === 0) return 0;
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const run = runs[mid];
    if (row < run.startRow) {
      hi = mid - 1;
    } else if (row >= run.endRowExclusive) {
      lo = mid + 1;
    } else {
      return run.styleId;
    }
  }
  return 0;
}

/**
 * @param {FormatRun[] | undefined | null} a
 * @param {FormatRun[] | undefined | null} b
 * @returns {boolean}
 */
function formatRunsEqual(a, b) {
  if (a === b) return true;
  const al = a ? a.length : 0;
  const bl = b ? b.length : 0;
  if (al !== bl) return false;
  for (let i = 0; i < al; i++) {
    const ar = a[i];
    const br = b[i];
    if (!ar || !br) return false;
    if (ar.startRow !== br.startRow) return false;
    if (ar.endRowExclusive !== br.endRowExclusive) return false;
    if (ar.styleId !== br.styleId) return false;
  }
  return true;
}

/**
 * Normalize a run list:
 * - drop invalid/default runs
 * - merge adjacent runs with identical `styleId`
 *
 * Assumes `runs` are already sorted by `startRow`.
 *
 * @param {FormatRun[]} runs
 * @returns {FormatRun[]}
 */
function normalizeFormatRuns(runs) {
  /** @type {FormatRun[]} */
  const out = [];
  for (const run of runs) {
    if (!run) continue;
    const startRow = Number(run.startRow);
    const endRowExclusive = Number(run.endRowExclusive);
    const styleId = Number(run.styleId);
    if (!Number.isInteger(startRow) || startRow < 0) continue;
    if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;
    if (!Number.isInteger(styleId) || styleId <= 0) continue;
    const last = out[out.length - 1];
    if (last && last.styleId === styleId && last.endRowExclusive === startRow) {
      last.endRowExclusive = endRowExclusive;
      continue;
    }
    out.push({ startRow, endRowExclusive, styleId });
  }
  return out;
}

/**
 * Apply a style patch to a column's run list over a row interval.
 *
 * This is the core of the "compressed range formatting" layer used by `setRangeFormat`
 * for large rectangles.
 *
 * @param {FormatRun[]} runs
 * @param {number} startRow
 * @param {number} endRowExclusive
 * @param {Record<string, any> | null} stylePatch
 * @param {StyleTable} styleTable
 * @returns {FormatRun[]}
 */
function patchFormatRuns(runs, startRow, endRowExclusive, stylePatch, styleTable) {
  const clampedStart = Math.max(0, Math.trunc(startRow));
  const clampedEnd = Math.max(clampedStart, Math.trunc(endRowExclusive));
  if (clampedEnd <= clampedStart) return runs ? runs.slice() : [];

  const input = Array.isArray(runs) ? runs : [];
  /** @type {FormatRun[]} */
  const out = [];

  let i = 0;
  // Copy runs strictly before the target interval.
  while (i < input.length && input[i].endRowExclusive <= clampedStart) {
    out.push(cloneFormatRun(input[i]));
    i += 1;
  }

  // If we start inside an existing run, preserve the prefix.
  if (i < input.length) {
    const run = input[i];
    if (run.startRow < clampedStart && run.endRowExclusive > clampedStart) {
      out.push({ startRow: run.startRow, endRowExclusive: clampedStart, styleId: run.styleId });
    }
  }

  let cursor = clampedStart;
  while (cursor < clampedEnd) {
    const run = i < input.length ? input[i] : null;

    // No more runs overlap: fill the rest of the interval as a gap.
    if (!run || run.startRow >= clampedEnd) {
      const baseStyle = styleTable.get(0);
      const merged = applyStylePatch(baseStyle, stylePatch);
      const styleId = styleTable.intern(merged);
      if (styleId !== 0) out.push({ startRow: cursor, endRowExclusive: clampedEnd, styleId });
      cursor = clampedEnd;
      break;
    }

    // Gap before the next run.
    if (run.startRow > cursor) {
      const gapEnd = Math.min(run.startRow, clampedEnd);
      const baseStyle = styleTable.get(0);
      const merged = applyStylePatch(baseStyle, stylePatch);
      const styleId = styleTable.intern(merged);
      if (styleId !== 0) out.push({ startRow: cursor, endRowExclusive: gapEnd, styleId });
      cursor = gapEnd;
      continue;
    }

    // Overlap with current run.
    const overlapEnd = Math.min(run.endRowExclusive, clampedEnd);
    const baseStyle = styleTable.get(run.styleId);
    const merged = applyStylePatch(baseStyle, stylePatch);
    const styleId = styleTable.intern(merged);
    if (styleId !== 0) out.push({ startRow: cursor, endRowExclusive: overlapEnd, styleId });
    cursor = overlapEnd;

    // Advance past fully-consumed runs. If the run extends beyond the interval,
    // we'll preserve its suffix after the loop.
    if (run.endRowExclusive <= clampedEnd) {
      i += 1;
    } else {
      break;
    }
  }

  // Preserve the suffix of the current overlapping run (if any) and all remaining runs.
  if (i < input.length) {
    const run = input[i];
    if (run.startRow < clampedEnd && run.endRowExclusive > clampedEnd) {
      out.push({ startRow: clampedEnd, endRowExclusive: run.endRowExclusive, styleId: run.styleId });
      i += 1;
    }
    for (; i < input.length; i++) {
      out.push(cloneFormatRun(input[i]));
    }
  }

  // Runs may now contain adjacent segments with identical style ids (e.g. patch is a no-op or
  // patching a gap created the same style as a neighbor). Merge + drop defaults.
  out.sort((a, b) => a.startRow - b.startRow);
  return normalizeFormatRuns(out);
}

/**
 * Apply a style patch to only the *existing* runs over a row interval.
 *
 * Unlike {@link patchFormatRuns}, this does NOT create new runs for gaps. This is useful when
 * applying formatting to full rows/cols/sheet where the underlying formatting should stay stored
 * in the sheet/row/col layers, but we still need the patch to override any pre-existing range-run
 * formatting (since range-runs have higher precedence than row/col defaults).
 *
 * @param {FormatRun[]} runs
 * @param {number} startRow
 * @param {number} endRowExclusive
 * @param {Record<string, any> | null} stylePatch
 * @param {StyleTable} styleTable
 * @returns {FormatRun[]}
 */
function patchExistingFormatRuns(runs, startRow, endRowExclusive, stylePatch, styleTable) {
  const clampedStart = Math.max(0, Math.trunc(startRow));
  const clampedEnd = Math.max(clampedStart, Math.trunc(endRowExclusive));
  if (clampedEnd <= clampedStart) return Array.isArray(runs) ? runs.slice() : [];

  const input = Array.isArray(runs) ? runs : [];
  /** @type {FormatRun[]} */
  const out = [];

  for (const run of input) {
    if (!run) continue;

    // No overlap.
    if (run.endRowExclusive <= clampedStart || run.startRow >= clampedEnd) {
      out.push(cloneFormatRun(run));
      continue;
    }

    // Prefix.
    if (run.startRow < clampedStart) {
      out.push({ startRow: run.startRow, endRowExclusive: clampedStart, styleId: run.styleId });
    }

    // Overlap.
    const overlapStart = Math.max(run.startRow, clampedStart);
    const overlapEnd = Math.min(run.endRowExclusive, clampedEnd);
    const baseStyle = styleTable.get(run.styleId);
    const merged = applyStylePatch(baseStyle, stylePatch);
    const styleId = styleTable.intern(merged);
    if (styleId !== 0) out.push({ startRow: overlapStart, endRowExclusive: overlapEnd, styleId });

    // Suffix.
    if (run.endRowExclusive > clampedEnd) {
      out.push({ startRow: clampedEnd, endRowExclusive: run.endRowExclusive, styleId: run.styleId });
    }
  }

  out.sort((a, b) => a.startRow - b.startRow);
  return normalizeFormatRuns(out);
}

class SheetModel {
  constructor() {
    /** @type {Map<string, CellState>} */
    this.cells = new Map();
    /**
     * Keys of cells with an explicit cell-level style override (`styleId !== 0`).
     *
     * This is a derived index used to keep formatting operations O(#formatted cells)
     * instead of O(#stored cells). (Stored cells can be large due to value/formula
     * entries with `styleId === 0`.)
     *
     * @type {Set<string>}
     */
    this.styledCells = new Set();
    /** @type {Map<number, Set<number>>} */
    this.styledCellsByRow = new Map();
    /** @type {Map<number, Set<number>>} */
    this.styledCellsByCol = new Map();
    /** @type {SheetViewState} */
    this.view = emptySheetViewState();

    /**
     * Layered formatting.
     *
     * We store formatting at multiple granularities (sheet/col/row/range-run/cell) so the UI can apply
     * formatting to whole rows/columns and large rectangles without eagerly materializing every cell.
     *
     * The per-cell layer continues to live on `CellState.styleId`. The remaining layers live here.
     */
    this.defaultStyleId = 0;
    /** @type {Map<number, number>} */
    this.rowStyleIds = new Map();
    /** @type {Map<number, number>} */
    this.colStyleIds = new Map();

    /**
     * Range-based formatting layer for large rectangles (sparse, compressed).
     *
     * Stored as per-column sorted, non-overlapping row interval runs.
     *
     * @type {Map<number, FormatRun[]>}
     */
    this.formatRunsByCol = new Map();

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
     * Bounding box of range-run formatting overrides (formatRunsByCol).
     *
     * This represents the used-range impact of rectangular range formatting that may not span
     * full rows/cols.
     *
     * @type {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
     */
    this.rangeRunBounds = null;

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
    this.rangeRunBoundsDirty = false;

    // Track the number of cells that contribute to `contentBounds` so we can fast-path the
    // empty case (common when clearing contents but preserving styles).
    this.contentCellCount = 0;

    // Debug counters for unit tests to verify recomputation only occurs when required.
    this.__contentBoundsRecomputeCount = 0;
    this.__formatBoundsRecomputeCount = 0;
    this.__rowStyleBoundsRecomputeCount = 0;
    this.__colStyleBoundsRecomputeCount = 0;
    this.__rangeRunBoundsRecomputeCount = 0;
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
    const beforeHasStyle = Boolean(before && before.styleId !== 0);

    const after = cloneCellState(cell);
    const afterIsEmpty = after.value == null && after.formula == null && after.styleId === 0;
    const afterHasContent = Boolean(after.value != null || after.formula != null);
    const afterHasFormat = !afterIsEmpty;
    const afterHasStyle = !afterIsEmpty && after.styleId !== 0;

    // Update the canonical cell map first.
    if (afterIsEmpty) {
      this.cells.delete(key);
    } else {
      this.cells.set(key, after);
    }

    // Maintain the derived set of styled cells.
    if (beforeHasStyle !== afterHasStyle) {
      if (afterHasStyle) {
        this.styledCells.add(key);

        let cols = this.styledCellsByRow.get(row);
        if (!cols) {
          cols = new Set();
          this.styledCellsByRow.set(row, cols);
        }
        cols.add(col);

        let rows = this.styledCellsByCol.get(col);
        if (!rows) {
          rows = new Set();
          this.styledCellsByCol.set(col, rows);
        }
        rows.add(row);
      } else {
        this.styledCells.delete(key);

        const cols = this.styledCellsByRow.get(row);
        if (cols) {
          cols.delete(col);
          if (cols.size === 0) this.styledCellsByRow.delete(row);
        }

        const rows = this.styledCellsByCol.get(col);
        if (rows) {
          rows.delete(row);
          if (rows.size === 0) this.styledCellsByCol.delete(col);
        }
      }
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
   * Set the compressed range-run formatting list for a single column.
   *
   * @param {number} col
   * @param {FormatRun[] | null | undefined} runs
   */
  setFormatRunsForCol(col, runs) {
    const colIdx = Number(col);
    if (!Number.isInteger(colIdx) || colIdx < 0) return;

    const beforeRuns = this.formatRunsByCol.get(colIdx) ?? [];
    const beforeHad = beforeRuns.length > 0;

    /** @type {FormatRun[]} */
    const normalized = Array.isArray(runs) ? runs.map(cloneFormatRun) : [];
    normalized.sort((a, b) => a.startRow - b.startRow);
    const afterRuns = normalizeFormatRuns(normalized);
    const afterHas = afterRuns.length > 0;

    if (formatRunsEqual(beforeRuns, afterRuns)) return;

    if (afterHas) {
      this.formatRunsByCol.set(colIdx, afterRuns);
    } else {
      this.formatRunsByCol.delete(colIdx);
    }

    const bounds = this.rangeRunBounds;
    if (!bounds || this.rangeRunBoundsDirty) {
      // If bounds are missing/dirty, we'll recompute lazily when requested.
      if (this.formatRunsByCol.size === 0) {
        this.rangeRunBounds = null;
        this.rangeRunBoundsDirty = false;
      } else {
        this.rangeRunBoundsDirty = true;
      }
      return;
    }

    const beforeTouched =
      beforeHad &&
      (colIdx === bounds.startCol ||
        colIdx === bounds.endCol ||
        beforeRuns[0]?.startRow === bounds.startRow ||
        beforeRuns[beforeRuns.length - 1]?.endRowExclusive - 1 === bounds.endRow);

    if (!afterHas) {
      if (this.formatRunsByCol.size === 0) {
        this.rangeRunBounds = null;
        this.rangeRunBoundsDirty = false;
      } else if (beforeTouched) {
        this.rangeRunBoundsDirty = true;
      }
      return;
    }

    // afterHas: update bounds for expansion, but mark dirty for potential shrink.
    const afterMinRow = afterRuns[0].startRow;
    const afterMaxRow = afterRuns[afterRuns.length - 1].endRowExclusive - 1;
    bounds.startCol = Math.min(bounds.startCol, colIdx);
    bounds.endCol = Math.max(bounds.endCol, colIdx);
    bounds.startRow = Math.min(bounds.startRow, afterMinRow);
    bounds.endRow = Math.max(bounds.endRow, afterMaxRow);

    if (beforeTouched) {
      const beforeMinRow = beforeRuns[0].startRow;
      const beforeMaxRow = beforeRuns[beforeRuns.length - 1].endRowExclusive - 1;
      if (afterMinRow > beforeMinRow || afterMaxRow < beforeMaxRow) {
        this.rangeRunBoundsDirty = true;
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
  getRangeRunBounds() {
    if (this.formatRunsByCol.size === 0) {
      this.rangeRunBounds = null;
      this.rangeRunBoundsDirty = false;
      return null;
    }
    if (this.rangeRunBoundsDirty || !this.rangeRunBounds) {
      this.__rangeRunBoundsRecomputeCount += 1;
      this.rangeRunBounds = this.#recomputeRangeRunBounds();
      this.rangeRunBoundsDirty = false;
    }
    return this.rangeRunBounds ? { ...this.rangeRunBounds } : null;
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
   * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
   */
  #recomputeRangeRunBounds() {
    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;

    for (const [col, runs] of this.formatRunsByCol.entries()) {
      if (!runs || runs.length === 0) continue;
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
      const first = runs[0];
      const last = runs[runs.length - 1];
      if (first) minRow = Math.min(minRow, first.startRow);
      if (last) maxRow = Math.max(maxRow, last.endRowExclusive - 1);
    }

    if (minRow === Infinity || minCol === Infinity) return null;
    return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
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

    /** @type {Map<string, SheetMetaState>} */
    this.sheetMeta = new Map();

    /**
     * Workbook-scoped image store for in-cell/floating images.
     *
     * Keys are stable `imageId` strings. Values contain the binary payload plus optional mime type.
     *
     * @type {Map<string, ImageEntry>}
     */
    this.images = new Map();

    /**
     * Per-sheet drawings list (floating images/shapes/chart placeholders).
     *
     * Drawing objects must be JSON-serializable (so they can survive snapshots/undo/redo).
     *
     * @type {Map<string, any[]>}
     */
    this.drawingsBySheet = new Map();

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

    /**
     * Per-sheet monotonically increasing counter that increments whenever that sheet's
     * *content* changes (value/formula), but ignores formatting/view changes.
     *
     * This is used by downstream caches (e.g. AI workbook context building) to avoid
     * invalidating work for unrelated sheets.
     *
     * @type {Map<string, number>}
     */
    this.contentVersionBySheet = new Map();
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
   *     rangeRunDeltas: RangeRunDelta[],
   *     drawingDeltas: DrawingDelta[],
   *     imageDeltas: Array<{ imageId: string, before: { mimeType: string | null, byteLength: number } | null, after: { mimeType: string | null, byteLength: number } | null }>,
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
        this.activeBatch.deltasByFormat.size > 0 ||
        this.activeBatch.deltasByRangeRun.size > 0 ||
        this.activeBatch.deltasByDrawing.size > 0 ||
        this.activeBatch.deltasByImage.size > 0 ||
        this.activeBatch.deltasBySheetMeta.size > 0 ||
        this.activeBatch.sheetOrderDelta != null)
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
   * Monotonically increasing workbook mutation counter.
   *
   * @returns {number}
   */
  getUpdateVersion() {
    return this.updateVersion;
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
   * Read a cell without materializing the sheet.
   *
   * `DocumentController.getCell()` creates sheets lazily the first time a sheet id is
   * referenced. Some callers (e.g. tab completion previews) want to probe for values
   * without mutating workbook structure (and without creating "phantom" sheets).
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {CellState}
   */
  peekCell(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return emptyCellState();
    return sheet.getCell(c.row, c.col);
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
    *   `sheet < col < row < range-run < cell` (later layers override earlier layers for the same property).
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
    // Avoid cloning the full cell state; `getCellFormat` only needs style ids.
    let sheet = this.model.sheets.get(sheetId);
    if (!sheet) {
      this.model.getCell(sheetId, 0, 0);
      sheet = this.model.sheets.get(sheetId);
    }

    // Read the stored style id directly from the sheet's sparse cell map. This avoids
    // cloning/normalizing rich cell values for callers that only need formatting.
    const cell = sheet?.cells?.get?.(`${c.row},${c.col}`) ?? null;
    const cellStyleId = typeof cell?.styleId === "number" ? cell.styleId : 0;

    const sheetStyle = this.styleTable.get(sheet?.defaultStyleId ?? 0);
    const colStyle = this.styleTable.get(sheet?.colStyleIds.get(c.col) ?? 0);
    const rowStyle = this.styleTable.get(sheet?.rowStyleIds.get(c.row) ?? 0);
    const runStyleId =
      sheet && sheet.formatRunsByCol ? styleIdForRowInRuns(sheet.formatRunsByCol.get(c.col), c.row) : 0;
    const runStyle = this.styleTable.get(runStyleId);
    const cellStyle = this.styleTable.get(cellStyleId);

    // Precedence: sheet < col < row < range-run < cell.
    const sheetCol = applyStylePatch(sheetStyle, colStyle);
    const sheetColRow = applyStylePatch(sheetCol, rowStyle);
    const sheetColRowRun = applyStylePatch(sheetColRow, runStyle);
    return applyStylePatch(sheetColRowRun, cellStyle);
  }

  /**
   * Return the set of style ids contributing to a cell's effective formatting.
   *
   * This is useful for callers that want to cache derived formatting (clipboard export,
   * render caches, etc) without needing to stringify full style objects.
   *
   * Tuple order is:
   * `[sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId, rangeRunStyleId]`.
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {[number, number, number, number, number]}
   */
  getCellFormatStyleIds(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    // Ensure the sheet is materialized (DocumentController is lazily sheet-creating).
    // Avoid cloning the full cell state; we only need the stored style ids.
    let sheet = this.model.sheets.get(sheetId);
    if (!sheet) {
      this.model.getCell(sheetId, 0, 0);
      sheet = this.model.sheets.get(sheetId);
    }

    // Read the stored cell styleId directly from the sheet's sparse cell map. This avoids
    // cloning/normalizing rich cell values for callers that only need the style tuple
    // (clipboard caches, render caches, etc).
    const cell = sheet?.cells?.get?.(`${c.row},${c.col}`) ?? null;
    const cellStyleId = typeof cell?.styleId === "number" ? cell.styleId : 0;
    const rangeRunStyleId = styleIdForRowInRuns(sheet?.formatRunsByCol?.get?.(c.col), c.row);
    return [
      sheet?.defaultStyleId ?? 0,
      sheet?.rowStyleIds.get(c.row) ?? 0,
      sheet?.colStyleIds.get(c.col) ?? 0,
      cellStyleId,
      rangeRunStyleId,
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
   * Return sheet metadata (name/visibility/tab color).
   *
   * Note: DocumentController historically created sheets lazily. If a sheet has no explicit
   * metadata entry, we treat it as `{ name: sheetId, visibility: "visible" }`.
   *
   * @param {string} sheetId
   * @returns {SheetMetaState | null}
   */
  getSheetMeta(sheetId) {
    const id = String(sheetId ?? "").trim();
    if (!id) return null;
    if (!this.model.sheets.has(id) && !this.sheetMeta.has(id)) return null;
    const meta = this.sheetMeta.get(id) ?? { name: id, visibility: "visible" };
    return cloneSheetMetaState(meta);
  }

  /**
   * Return the drawings list for a sheet.
   *
   * The returned value is a deep clone so callers can treat it as immutable.
   *
   * @param {string} sheetId
   * @returns {any[]}
   */
  getSheetDrawings(sheetId) {
    const id = String(sheetId ?? "").trim();
    if (!id) return [];
    const drawings = this.drawingsBySheet.get(id);
    return Array.isArray(drawings) ? cloneJsonSerializable(drawings) : [];
  }

  /**
   * Replace the drawings list for a sheet.
   *
   * This is undoable and persisted in `encodeState()` snapshots.
   *
   * @param {string} sheetId
   * @param {any[]} nextDrawings
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setSheetDrawings(sheetId, nextDrawings, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    // Ensure the sheet exists (DocumentController historically materializes sheets lazily).
    this.model.getCell(id, 0, 0);

    const beforeRaw = this.drawingsBySheet.get(id) ?? [];
    const before = Array.isArray(beforeRaw) ? cloneJsonSerializable(beforeRaw) : [];

    const rawList = Array.isArray(nextDrawings) ? nextDrawings : [];
    /** @type {any[]} */
    const after = [];
    for (const raw of rawList) {
      if (!isJsonObject(raw)) {
        throw new Error("Drawings must be JSON objects");
      }
      const drawingId = String(raw.id ?? "").trim();
      if (!drawingId) throw new Error("Drawing.id must be a non-empty string");
      if (!("anchor" in raw)) throw new Error("Drawing.anchor is required");
      if (!("kind" in raw)) throw new Error("Drawing.kind is required");
      const zOrder = Number(raw.zOrder);
      if (!Number.isFinite(zOrder)) throw new Error("Drawing.zOrder must be a finite number");
      const cloned = cloneJsonSerializable(raw);
      cloned.id = drawingId;
      cloned.zOrder = zOrder;
      after.push(cloned);
    }

    if (jsonSerializableEquals(before, after)) return;
    this.#applyUserWorkbookEdits({ drawingDeltas: [{ sheetId: id, before, after }] }, options);
  }

  /**
   * Convenience helper to append a drawing to the end of a sheet's drawing list.
   *
   * @param {string} sheetId
   * @param {any} drawing
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  insertDrawing(sheetId, drawing, options = {}) {
    const existing = this.getSheetDrawings(sheetId);
    this.setSheetDrawings(sheetId, [...existing, drawing], options);
  }

  /**
   * Convenience helper to update a drawing by id.
   *
   * @param {string} sheetId
   * @param {string} drawingId
   * @param {Record<string, any> | ((drawing: any) => any)} patchOrUpdater
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  updateDrawing(sheetId, drawingId, patchOrUpdater, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    const targetId = String(drawingId ?? "").trim();
    if (!targetId) throw new Error("Drawing id cannot be empty");

    const existing = this.getSheetDrawings(id);
    const idx = existing.findIndex((d) => isJsonObject(d) && d.id === targetId);
    if (idx === -1) return;

    const current = existing[idx];
    let next;
    if (typeof patchOrUpdater === "function") {
      next = patchOrUpdater(cloneJsonSerializable(current));
    } else if (patchOrUpdater && typeof patchOrUpdater === "object") {
      next = { ...current, ...patchOrUpdater };
    } else {
      throw new Error("patchOrUpdater must be an object or function");
    }

    if (!isJsonObject(next)) throw new Error("Drawing updates must produce a JSON object");
    next.id = targetId;

    const updated = existing.slice();
    updated[idx] = next;
    this.setSheetDrawings(id, updated, options);
  }

  /**
   * Convenience helper to delete a drawing by id.
   *
   * @param {string} sheetId
   * @param {string} drawingId
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  deleteDrawing(sheetId, drawingId, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    const targetId = String(drawingId ?? "").trim();
    if (!targetId) throw new Error("Drawing id cannot be empty");

    const existing = this.getSheetDrawings(id);
    const next = existing.filter((d) => !(isJsonObject(d) && d.id === targetId));
    if (next.length === existing.length) return;
    this.setSheetDrawings(id, next, options);
  }

  /**
   * Get an image entry from the workbook-scoped image store.
   *
   * The returned value is a deep clone so callers can treat it as immutable.
   *
   * @param {string} imageId
   * @returns {ImageEntry | null}
   */
  getImage(imageId) {
    const id = String(imageId ?? "").trim();
    if (!id) return null;
    const entry = this.images.get(id);
    return entry ? cloneImageEntry(entry) : null;
  }

  /**
   * Convenience helper for UI layers: get a browser `Blob` for an image store entry.
   *
   * This avoids every caller re-implementing byte → Blob conversion and lets render layers
   * treat the controller as the single source of truth for workbook media.
   *
   * @param {string} imageId
   * @returns {Blob | null}
   */
  getImageBlob(imageId) {
    const id = String(imageId ?? "").trim();
    if (!id) return null;
    const entry = this.images.get(id);
    if (!entry) return null;
    if (typeof Blob === "undefined") return null;
    try {
      const explicit =
        entry && typeof entry.mimeType === "string" && entry.mimeType.trim().length > 0 ? entry.mimeType : null;
      const mimeType = explicit ?? inferMimeTypeForImage(id, entry.bytes);
      return new Blob([entry.bytes], { type: mimeType });
    } catch {
      return null;
    }
  }

  /**
   * Set an image entry in the workbook-scoped image store.
   *
   * This is undoable and persisted in `encodeState()` snapshots.
   *
   * @param {string} imageId
   * @param {{ bytes: Uint8Array, mimeType?: string | null }} entry
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setImage(imageId, entry, options = {}) {
    const id = String(imageId ?? "").trim();
    if (!id) throw new Error("Image id cannot be empty");
    if (!entry || typeof entry !== "object") throw new Error("Image entry must be an object");
    const bytes = entry.bytes;
    if (!(bytes instanceof Uint8Array)) throw new Error("Image entry bytes must be a Uint8Array");

    const mimeTypeRaw = "mimeType" in entry ? entry.mimeType : undefined;
    let mimeType = undefined;
    if (mimeTypeRaw === undefined) {
      mimeType = undefined;
    } else if (mimeTypeRaw === null) {
      mimeType = null;
    } else if (typeof mimeTypeRaw === "string") {
      mimeType = mimeTypeRaw;
    } else {
      throw new Error("Image mimeType must be a string or null");
    }

    /** @type {ImageEntry} */
    const after = { bytes: bytes.slice() };
    if (mimeTypeRaw !== undefined) after.mimeType = mimeType;

    const beforeEntry = this.images.get(id) ?? null;
    const before = beforeEntry ? cloneImageEntry(beforeEntry) : null;
    if (imageEntryEquals(before, after)) return;

    this.#applyUserWorkbookEdits({ imageDeltas: [{ imageId: id, before, after }] }, options);
  }

  /**
   * Delete an image from the workbook-scoped image store.
   *
   * @param {string} imageId
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  deleteImage(imageId, options = {}) {
    const id = String(imageId ?? "").trim();
    if (!id) throw new Error("Image id cannot be empty");
    const beforeEntry = this.images.get(id);
    if (!beforeEntry) return;
    const before = cloneImageEntry(beforeEntry);
    this.#applyUserWorkbookEdits({ imageDeltas: [{ imageId: id, before, after: null }] }, options);
  }

  /**
   * Return sheet ids whose visibility is "visible", in the current sheet order.
   *
   * @returns {string[]}
   */
  getVisibleSheetIds() {
    const ids = this.getSheetIds();
    const out = [];
    for (const id of ids) {
      const meta = this.getSheetMeta(id);
      const visibility = meta?.visibility ?? "visible";
      if (visibility === "visible") out.push(id);
    }
    return out;
  }

  /**
   * @param {string} sheetId
   * @returns {SheetMetaState}
   */
  #defaultSheetMeta(sheetId) {
    return { name: sheetId, visibility: "visible" };
  }

  /**
   * Normalize + validate a sheet name using Excel-like constraints.
   *
   * This intentionally mirrors `WorkbookSheetStore.validateSheetName` so callers that mutate
   * metadata via DocumentController (scripts, future UI wiring) get consistent behavior.
   *
   * @param {string} name
   * @param {{ ignoreSheetId?: string | null }} [options]
   * @returns {string}
   */
  #validateSheetName(name, options = {}) {
    const normalized = String(name ?? "").trim();
    const ignore = options.ignoreSheetId ?? null;
    const existingNames = this.getSheetIds()
      .filter((id) => !(ignore && id === ignore))
      .map((id) => (this.getSheetMeta(id) ?? this.#defaultSheetMeta(id)).name);

    const err = getSheetNameValidationErrorMessage(normalized, { existingNames });
    if (err) throw new Error(err);

    return normalized;
  }

  /**
   * Rename a sheet (metadata-only; does not rewrite formulas by itself).
   *
   * Callers that also rewrite formulas should wrap the rename + cell edits in `beginBatch/endBatch`
   * so undo/redo treats them as a single action.
   *
   * @param {string} sheetId
   * @param {string} newName
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  renameSheet(sheetId, newName, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    // Ensure the sheet exists (DocumentController historically materializes sheets lazily).
    this.model.getCell(id, 0, 0);

    const validated = this.#validateSheetName(newName, { ignoreSheetId: id });
    const before = this.getSheetMeta(id) ?? this.#defaultSheetMeta(id);
    const after = { ...before, name: validated };
    if (sheetMetaStateEquals(before, after)) return;

    this.#applyUserWorkbookEdits(
      { sheetMetaDeltas: [{ sheetId: id, before, after }] },
      { ...options, label: options.label ?? "Rename Sheet" },
    );
  }

  /**
   * @param {string} sheetId
   * @param {SheetVisibility} visibility
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  setSheetVisibility(sheetId, visibility, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    this.model.getCell(id, 0, 0);

    const before = this.getSheetMeta(id) ?? this.#defaultSheetMeta(id);
    const after = { ...before, visibility: visibility ?? "visible" };
    if (sheetMetaStateEquals(before, after)) return;

    this.#applyUserWorkbookEdits(
      { sheetMetaDeltas: [{ sheetId: id, before, after }] },
      { ...options, label: options.label ?? "Set Sheet Visibility" },
    );
  }

  /**
   * @param {string} sheetId
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  hideSheet(sheetId, options = {}) {
    this.setSheetVisibility(sheetId, "hidden", { ...options, label: options.label ?? "Hide Sheet" });
  }

  /**
   * @param {string} sheetId
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  unhideSheet(sheetId, options = {}) {
    this.setSheetVisibility(sheetId, "visible", { ...options, label: options.label ?? "Unhide Sheet" });
  }

  /**
   * @param {string} sheetId
   * @param {TabColor | string | undefined | null} tabColor
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  setSheetTabColor(sheetId, tabColor, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    this.model.getCell(id, 0, 0);

    const before = this.getSheetMeta(id) ?? this.#defaultSheetMeta(id);
    const after = { ...before, tabColor: tabColor ? cloneTabColor(tabColor) : undefined };
    if (sheetMetaStateEquals(before, after)) return;

    this.#applyUserWorkbookEdits(
      { sheetMetaDeltas: [{ sheetId: id, before, after }] },
      { ...options, label: options.label ?? "Set Tab Color" },
    );
  }

  /**
   * Reorder sheet ids by changing the workbook sheet iteration order.
   *
   * @param {string[]} sheetIdsInOrder
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  reorderSheets(sheetIdsInOrder, options = {}) {
    if (!Array.isArray(sheetIdsInOrder) || sheetIdsInOrder.length === 0) return;

    // Ensure all ids are materialized (DocumentController is lazily sheet-creating).
    for (const id of sheetIdsInOrder) {
      const trimmed = typeof id === "string" ? id.trim() : "";
      if (!trimmed) continue;
      this.model.getCell(trimmed, 0, 0);
    }

    const before = this.getSheetIds();
    const seen = new Set();
    /** @type {string[]} */
    const desired = [];

    for (const raw of sheetIdsInOrder) {
      const id = typeof raw === "string" ? raw.trim() : "";
      if (!id) continue;
      if (seen.has(id)) continue;
      if (!this.model.sheets.has(id)) continue;
      seen.add(id);
      desired.push(id);
    }

    for (const id of before) {
      if (seen.has(id)) continue;
      seen.add(id);
      desired.push(id);
    }

    if (desired.length === 0) return;
    if (desired.length === before.length && desired.every((id, i) => id === before[i])) return;

    this.#applyUserWorkbookEdits(
      { sheetOrderDelta: { before, after: desired } },
      { ...options, label: options.label ?? "Reorder Sheets" },
    );
  }

  /**
   * Add a new sheet.
   *
   * When `sheetId` is omitted, a random stable id is generated and the default visible name
   * ("Sheet1", "Sheet2", ...) is used.
   *
   * @param {{
   *   sheetId?: string,
   *   name?: string,
   *   insertAfterId?: string | null,
   * }} [params]
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   * @returns {string} The new sheet id
   */
  addSheet(params = {}, options = {}) {
    const randomUuid = (globalThis?.crypto?.randomUUID ?? null);
    const generateId = () => {
      if (typeof randomUuid === "function") return `sheet_${randomUuid.call(globalThis.crypto)}`;
      return `sheet_${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
    };

    const beforeOrder = this.getSheetIds();

    let sheetId = String(params.sheetId ?? "").trim();
    if (!sheetId) sheetId = generateId();
    if (this.model.sheets.has(sheetId)) throw new Error(`Duplicate sheet id: ${sheetId}`);

    const existingNames = new Set(
      beforeOrder.map((id) =>
        normalizeSheetNameForCaseInsensitiveCompare((this.getSheetMeta(id) ?? this.#defaultSheetMeta(id)).name),
      ),
    );
    const defaultName = (() => {
      if (params.name != null) return this.#validateSheetName(params.name, { ignoreSheetId: null });
      for (let n = 1; ; n += 1) {
        const candidate = `Sheet${n}`;
        if (!existingNames.has(normalizeSheetNameForCaseInsensitiveCompare(candidate))) return candidate;
      }
    })();

    const insertAfterId = params.insertAfterId ?? null;
    const afterOrder = beforeOrder.slice();
    const insertIdx = insertAfterId ? afterOrder.indexOf(insertAfterId) + 1 : afterOrder.length;
    const clampedIdx = insertIdx >= 0 && insertIdx <= afterOrder.length ? insertIdx : afterOrder.length;
    afterOrder.splice(clampedIdx, 0, sheetId);

    const meta = { name: defaultName, visibility: "visible" };

    this.#applyUserWorkbookEdits(
      {
        sheetMetaDeltas: [{ sheetId, before: null, after: meta }],
        sheetOrderDelta: { before: beforeOrder, after: afterOrder },
      },
      { ...options, label: options.label ?? "Add Sheet" },
    );

    return sheetId;
  }

  /**
   * Delete a sheet (cells + view + formatting + metadata).
   *
   * @param {string} sheetId
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  deleteSheet(sheetId, options = {}) {
    const id = String(sheetId ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");

    const beforeOrder = this.getSheetIds();
    if (!this.model.sheets.has(id)) throw new Error(`Sheet not found: ${id}`);
    if (beforeOrder.length <= 1) throw new Error("Cannot delete the last sheet");

    const sheet = this.model.sheets.get(id);
    if (!sheet) throw new Error(`Sheet not found: ${id}`);

    const beforeMeta = this.getSheetMeta(id) ?? this.#defaultSheetMeta(id);
    const afterOrder = beforeOrder.filter((s) => s !== id);
    const beforeDrawingsRaw = this.drawingsBySheet.get(id) ?? [];
    const beforeDrawings = Array.isArray(beforeDrawingsRaw) ? cloneJsonSerializable(beforeDrawingsRaw) : [];

    /** @type {CellDelta[]} */
    const cellDeltas = [];
    for (const [key] of sheet.cells.entries()) {
      const { row, col } = parseRowColKey(key);
      const before = this.model.getCell(id, row, col);
      const after = emptyCellState();
      cellDeltas.push({ sheetId: id, row, col, before, after: cloneCellState(after) });
    }

    /** @type {SheetViewDelta[]} */
    const sheetViewDeltas = [];
    const beforeView = this.model.getSheetView(id);
    const afterView = emptySheetViewState();
    if (!sheetViewStateEquals(beforeView, afterView)) {
      sheetViewDeltas.push({ sheetId: id, before: beforeView, after: afterView });
    }

    /** @type {FormatDelta[]} */
    const formatDeltas = [];
    if (sheet.defaultStyleId !== 0) {
      formatDeltas.push({ sheetId: id, layer: "sheet", beforeStyleId: sheet.defaultStyleId, afterStyleId: 0 });
    }
    for (const [row, styleId] of sheet.rowStyleIds.entries()) {
      if (styleId === 0) continue;
      formatDeltas.push({ sheetId: id, layer: "row", index: row, beforeStyleId: styleId, afterStyleId: 0 });
    }
    for (const [col, styleId] of sheet.colStyleIds.entries()) {
      if (styleId === 0) continue;
      formatDeltas.push({ sheetId: id, layer: "col", index: col, beforeStyleId: styleId, afterStyleId: 0 });
    }

    /** @type {RangeRunDelta[]} */
    const rangeRunDeltas = [];
    for (const [col, runs] of sheet.formatRunsByCol.entries()) {
      if (!runs || runs.length === 0) continue;
      rangeRunDeltas.push({
        sheetId: id,
        col,
        startRow: 0,
        endRowExclusive: EXCEL_MAX_ROWS,
        beforeRuns: runs.map(cloneFormatRun),
        afterRuns: [],
      });
    }

    this.#applyUserWorkbookEdits(
      {
        cellDeltas,
        sheetViewDeltas,
        formatDeltas,
        rangeRunDeltas,
        drawingDeltas: beforeDrawings.length > 0 ? [{ sheetId: id, before: beforeDrawings, after: [] }] : [],
        sheetMetaDeltas: [{ sheetId: id, before: beforeMeta, after: null }],
        sheetOrderDelta: { before: beforeOrder, after: afterOrder },
      },
      { ...options, label: options.label ?? "Delete Sheet" },
    );
  }

  /**
   * Return the current content version counter for a sheet.
   *
   * This starts at 0 and increments whenever the sheet's value/formula grid changes.
   *
   * @param {string} sheetId
   * @returns {number}
   */
  getSheetContentVersion(sheetId) {
    return this.contentVersionBySheet.get(sheetId) ?? 0;
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
    mergeBounds(sheet.getRangeRunBounds());

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
    // Avoid `Math.max(...rows.map(...))` spread: large writes can include tens of thousands of
    // rows (e.g. Python scalar fill), which would exceed JS engines' argument limits.
    let colCount = 0;
    for (const row of values) {
      const len = Array.isArray(row) ? row.length : 0;
      if (len > colCount) colCount = len;
    }
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
   * @param {{ label?: string, maxEnumeratedCells?: number, maxEnumeratedRows?: number }} [options]
   * @returns {boolean} True when the request was processed; false when skipped (e.g. safety caps).
   */
  setRangeFormat(sheetId, range, stylePatch, options = {}) {
    const maxEnumeratedCells = Number.isFinite(options?.maxEnumeratedCells) ? options.maxEnumeratedCells : 200_000;
    const maxEnumeratedRows = Number.isFinite(options?.maxEnumeratedRows) ? options.maxEnumeratedRows : 50_000;

    const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
    const isFullSheet = r.start.row === 0 && r.end.row === EXCEL_MAX_ROW && r.start.col === 0 && r.end.col === EXCEL_MAX_COL;
    const isFullHeightCols = r.start.row === 0 && r.end.row === EXCEL_MAX_ROW;
    const isFullWidthRows = r.start.col === 0 && r.end.col === EXCEL_MAX_COL;
    const rowCount = r.end.row - r.start.row + 1;
    const colCount = r.end.col - r.start.col + 1;
    const area = rowCount * colCount;

    const formatRangeDebug = `${formatA1(r.start)}:${formatA1(r.end)}`;

    // Ensure sheet exists so we can mutate format layers without materializing cells.
    this.model.getCell(sheetId, 0, 0);
    const sheet = this.model.sheets.get(sheetId);
    if (!sheet) return false;

    /** @type {CellDelta[]} */
    const cellDeltas = [];
    /** @type {FormatDelta[]} */
    const formatDeltas = [];
    /** @type {RangeRunDelta[]} */
    const rangeRunDeltas = [];

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

    const forEachStyledCell = (visitor) => {
      const styledKeys = sheet.styledCells;
      // Prefer the derived styled-cell index (styleId != 0) so full-row/col/sheet formatting
      // doesn't need to scan potentially huge value/formula grids where styleId==0.
      if (styledKeys && styledKeys.size >= 0) {
        for (const key of styledKeys) {
          const cell = sheet.cells.get(key);
          if (!cell || cell.styleId === 0) continue;
          visitor(key, cell);
        }
        return;
      }

      // Extremely defensive fallback for older model encodings.
      for (const [key, cell] of sheet.cells.entries()) {
        if (!cell || cell.styleId === 0) continue;
        visitor(key, cell);
      }
    };

    const styledCellsByRow = sheet.styledCellsByRow;
    const styledCellsByCol = sheet.styledCellsByCol;
    const hasAxisIndex =
      styledCellsByRow && typeof styledCellsByRow.get === "function" && styledCellsByCol && typeof styledCellsByCol.get === "function";

    const forEachStyledCellInColRange = (startCol, endCol, visitor) => {
      if (!hasAxisIndex) {
        forEachStyledCell((key, cell) => {
          const { col } = parseRowColKey(key);
          if (col < startCol || col > endCol) return;
          visitor(key, cell);
        });
        return;
      }

      for (let col = startCol; col <= endCol; col++) {
        const rows = styledCellsByCol.get(col);
        if (!rows || rows.size === 0) continue;
        for (const row of rows) {
          const key = `${row},${col}`;
          const cell = sheet.cells.get(key);
          if (!cell || cell.styleId === 0) continue;
          visitor(key, cell);
        }
      }
    };

    const forEachStyledCellInRowRange = (startRow, endRow, visitor) => {
      if (!hasAxisIndex) {
        forEachStyledCell((key, cell) => {
          const { row } = parseRowColKey(key);
          if (row < startRow || row > endRow) return;
          visitor(key, cell);
        });
        return;
      }

      const rowLen = endRow - startRow + 1;
      if (rowLen <= maxEnumeratedRows) {
        for (let row = startRow; row <= endRow; row++) {
          const cols = styledCellsByRow.get(row);
          if (!cols || cols.size === 0) continue;
          for (const col of cols) {
            const key = `${row},${col}`;
            const cell = sheet.cells.get(key);
            if (!cell || cell.styleId === 0) continue;
            visitor(key, cell);
          }
        }
        return;
      }

      // Avoid enumerating huge row ranges; scan only rows that actually contain styled cells.
      for (const [row, cols] of styledCellsByRow.entries()) {
        if (row < startRow || row > endRow) continue;
        for (const col of cols) {
          const key = `${row},${col}`;
          const cell = sheet.cells.get(key);
          if (!cell || cell.styleId === 0) continue;
          visitor(key, cell);
        }
      }
    };

    const forEachStyledCellInRect = (startRow, endRow, startCol, endCol, visitor) => {
      if (!hasAxisIndex) {
        forEachStyledCell((key, cell) => {
          const { row, col } = parseRowColKey(key);
          if (row < startRow || row > endRow) return;
          if (col < startCol || col > endCol) return;
          visitor(key, cell);
        });
        return;
      }

      const rowCount = endRow - startRow + 1;
      const colCount = endCol - startCol + 1;
      const canEnumerateRows = rowCount <= maxEnumeratedRows;

      if (canEnumerateRows && rowCount <= colCount) {
        for (let row = startRow; row <= endRow; row++) {
          const cols = styledCellsByRow.get(row);
          if (!cols || cols.size === 0) continue;
          for (const col of cols) {
            if (col < startCol || col > endCol) continue;
            const key = `${row},${col}`;
            const cell = sheet.cells.get(key);
            if (!cell || cell.styleId === 0) continue;
            visitor(key, cell);
          }
        }
        return;
      }

      // Column iteration is bounded by Excel's maximum column count (16,384), and avoids
      // enumerating huge row ranges.
      for (let col = startCol; col <= endCol; col++) {
        const rows = styledCellsByCol.get(col);
        if (!rows || rows.size === 0) continue;
        for (const row of rows) {
          if (row < startRow || row > endRow) continue;
          const key = `${row},${col}`;
          const cell = sheet.cells.get(key);
          if (!cell || cell.styleId === 0) continue;
          visitor(key, cell);
        }
      }
    };

    if (isFullSheet) {
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
      forEachStyledCell((key) => {
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) return;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      });

      // Patch existing range-run formatting so the patch applies everywhere.
      for (const [col, beforeRuns] of sheet.formatRunsByCol.entries()) {
        const afterRuns = patchExistingFormatRuns(beforeRuns, 0, EXCEL_MAX_ROWS, stylePatch, this.styleTable);
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;
        rangeRunDeltas.push({
          sheetId,
          col,
          startRow: 0,
          endRowExclusive: EXCEL_MAX_ROWS,
          beforeRuns: beforeRuns.map(cloneFormatRun),
          afterRuns: afterRuns.map(cloneFormatRun),
        });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, rangeRunDeltas, { label: options.label });
      return true;
    }

    if (isFullHeightCols) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const beforeStyleId = sheet.colStyleIds.get(col) ?? 0;
        const afterStyleId = patchStyleId(beforeStyleId);
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "col", index: col, beforeStyleId, afterStyleId });
      }

      // Ensure the patch overrides explicit cell formatting (sparse overrides only).
      forEachStyledCellInColRange(r.start.col, r.end.col, (key) => {
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) return;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      });

      // Patch any existing range-run formatting in the selected columns so the patch applies everywhere.
      for (let col = r.start.col; col <= r.end.col; col++) {
        const beforeRuns = sheet.formatRunsByCol.get(col) ?? [];
        if (beforeRuns.length === 0) continue;
        const afterRuns = patchExistingFormatRuns(beforeRuns, 0, EXCEL_MAX_ROWS, stylePatch, this.styleTable);
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;
        rangeRunDeltas.push({
          sheetId,
          col,
          startRow: 0,
          endRowExclusive: EXCEL_MAX_ROWS,
          beforeRuns: beforeRuns.map(cloneFormatRun),
          afterRuns: afterRuns.map(cloneFormatRun),
        });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, rangeRunDeltas, { label: options.label });
      return true;
    }

    if (isFullWidthRows) {
      // Formatting full-width row ranges requires enumerating each row to update the row formatting layer.
      // For huge selections this can freeze the UI and allocate enormous delta arrays; cap the work.
      if (rowCount > maxEnumeratedRows) {
        console.warn(
          `[DocumentController.setRangeFormat] Skipping row formatting (too many rows): sheetId=${sheetId}, range=${formatRangeDebug}, rows=${rowCount}, cols=${colCount}, area=${area}, maxEnumeratedRows=${maxEnumeratedRows}, maxEnumeratedCells=${maxEnumeratedCells}`,
        );
        return false;
      }

      for (let row = r.start.row; row <= r.end.row; row++) {
        const beforeStyleId = sheet.rowStyleIds.get(row) ?? 0;
        const afterStyleId = patchStyleId(beforeStyleId);
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "row", index: row, beforeStyleId, afterStyleId });
      }

      // Ensure the patch overrides explicit cell formatting (sparse overrides only).
      forEachStyledCellInRowRange(r.start.row, r.end.row, (key) => {
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) return;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      });

      // Patch any existing range-run formatting that overlaps the selected rows so the patch applies everywhere.
      const startRow = r.start.row;
      const endRowExclusive = r.end.row + 1;
      for (const [col, beforeRuns] of sheet.formatRunsByCol.entries()) {
        const afterRuns = patchExistingFormatRuns(beforeRuns, startRow, endRowExclusive, stylePatch, this.styleTable);
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;
        rangeRunDeltas.push({
          sheetId,
          col,
          startRow,
          endRowExclusive,
          beforeRuns: beforeRuns.map(cloneFormatRun),
          afterRuns: afterRuns.map(cloneFormatRun),
        });
      }

      this.#applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, rangeRunDeltas, { label: options.label });
      return true;
    }

    // Large rectangular ranges should not enumerate every cell. Instead, store the
    // patch as compressed per-column row interval runs (`sheet.formatRunsByCol`).
    if (area > RANGE_RUN_FORMAT_THRESHOLD) {
      const startRow = r.start.row;
      const endRowExclusive = r.end.row + 1;

      /** @type {RangeRunDelta[]} */
      const rangeRunDeltas = [];

      for (let col = r.start.col; col <= r.end.col; col++) {
        const beforeRuns = sheet.formatRunsByCol.get(col) ?? [];
        const afterRuns = patchFormatRuns(beforeRuns, startRow, endRowExclusive, stylePatch, this.styleTable);
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;
        rangeRunDeltas.push({
          sheetId,
          col,
          startRow,
          endRowExclusive,
          beforeRuns: beforeRuns.map(cloneFormatRun),
          afterRuns: afterRuns.map(cloneFormatRun),
        });
      }

      // Ensure patches apply to cells that already have explicit per-cell styles inside the
      // rectangle (cell formatting is higher precedence than range runs).
      forEachStyledCellInRect(r.start.row, r.end.row, r.start.col, r.end.col, (key, cell) => {
        if (!cell || !cell.styleId || cell.styleId === 0) return;
        const { row, col } = parseRowColKey(key);
        const before = this.model.getCell(sheetId, row, col);
        const cellAfterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: cellAfterStyleId };
        if (cellStateEquals(before, after)) return;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      });

      // No-op: patch did not change any existing runs or explicit cell styles.
      // This is not a "skipped due to safety caps" case, so return true.
      if (rangeRunDeltas.length === 0 && cellDeltas.length === 0) return true;
      this.#applyUserRangeRunDeltas(cellDeltas, rangeRunDeltas, { label: options.label });
      return true;
    }

    if (area > maxEnumeratedCells) {
      console.warn(
        `[DocumentController.setRangeFormat] Skipping cell formatting (too many cells): sheetId=${sheetId}, range=${formatRangeDebug}, rows=${rowCount}, cols=${colCount}, area=${area}, maxEnumeratedCells=${maxEnumeratedCells}, maxEnumeratedRows=${maxEnumeratedRows}`,
      );
      return false;
    }

    // Fallback: sparse per-cell overrides.
    // Also patch existing range-run formatting that overlaps this rectangle so clearing formatting works
    // (a cell with no explicit styleId should still clear underlying range-run formatting).
    const startRow = r.start.row;
    const endRowExclusive = r.end.row + 1;
    for (let col = r.start.col; col <= r.end.col; col++) {
      const beforeRuns = sheet.formatRunsByCol.get(col) ?? [];
      if (beforeRuns.length === 0) continue;
      const afterRuns = patchExistingFormatRuns(beforeRuns, startRow, endRowExclusive, stylePatch, this.styleTable);
      if (formatRunsEqual(beforeRuns, afterRuns)) continue;
      rangeRunDeltas.push({
        sheetId,
        col,
        startRow,
        endRowExclusive,
        beforeRuns: beforeRuns.map(cloneFormatRun),
        afterRuns: afterRuns.map(cloneFormatRun),
      });
    }

    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const before = this.model.getCell(sheetId, row, col);
        const afterStyleId = patchStyleId(before.styleId);
        const after = { value: before.value, formula: before.formula, styleId: afterStyleId };
        if (cellStateEquals(before, after)) continue;
        cellDeltas.push({ sheetId, row, col, before, after: cloneCellState(after) });
      }
    }

    this.#applyUserCellAndFormatDeltas(cellDeltas, [], rangeRunDeltas, { label: options.label });
    return true;
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
   * Sheet-level view mutations (freeze panes, row/col sizes) are treated as view/UI
   * interactions and are allowed even when cell edits are blocked via `canEditCell`.
   *
   * In collaborative scenarios, the collaboration binders are responsible for deciding
   * whether these mutations should be persisted into shared state (e.g. viewers/commenters
   * should not write them into Yjs, but can still adjust their local view).
   *
   * @param {string} sheetId
   * @returns {boolean}
   */
  #canEditSheetView(sheetId) {
    return true;
  }

  /**
   * List merged-cell regions for a sheet.
   *
   * Ranges use inclusive end coordinates (`endRow`/`endCol` are inclusive).
   *
   * @param {string} sheetId
   * @returns {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>}
   */
  getMergedRanges(sheetId) {
    const view = this.model.getSheetView(sheetId);
    return Array.isArray(view?.mergedRanges)
      ? view.mergedRanges.map((r) => ({
          startRow: r.startRow,
          endRow: r.endRow,
          startCol: r.startCol,
          endCol: r.endCol,
        }))
      : [];
  }

  /**
   * Resolve a cell inside a merged region to its anchor cell (top-left).
   *
   * @param {string} sheetId
   * @param {CellCoord | string} coord
   * @returns {{ row: number, col: number } | null}
   */
  getMergedMasterCell(sheetId, coord) {
    const c = typeof coord === "string" ? parseA1(coord) : coord;
    const ranges = this.getMergedRanges(sheetId);
    for (const r of ranges) {
      if (c.row < r.startRow || c.row > r.endRow) continue;
      if (c.col < r.startCol || c.col > r.endCol) continue;
      return { row: r.startRow, col: r.startCol };
    }
    return null;
  }

  /**
   * Replace merged-cell regions for a sheet.
   *
   * @param {string} sheetId
   * @param {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>} mergedRanges
   * @param {{ label?: string, mergeKey?: string }} [options]
   */
  setMergedRanges(sheetId, mergedRanges, options = {}) {
    if (!this.#canEditSheetView(sheetId)) return;
    const before = this.model.getSheetView(sheetId);
    const after = cloneSheetViewState(before);
    if (Array.isArray(mergedRanges) && mergedRanges.length > 0) {
      after.mergedRanges = mergedRanges.map((r) => ({
        startRow: r.startRow,
        endRow: r.endRow,
        startCol: r.startCol,
        endCol: r.endCol,
      }));
    } else {
      delete after.mergedRanges;
    }
    const normalized = normalizeSheetViewState(after);
    if (sheetViewStateEquals(before, normalized)) return;
    this.#applyUserSheetViewDeltas([{ sheetId, before, after: normalized }], options);
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
    if (!this.#canEditSheetView(sheetId)) return;
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
    if (!this.#canEditSheetView(sheetId)) return;
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
    if (!this.#canEditSheetView(sheetId)) return;
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
      const effectiveStyle = normalizeSemanticDiffFormat(this.getCellFormat(sheetId, { row, col }));
      cells.set(semanticDiffCellKey(row, col), {
        value: cell.value ?? null,
        formula: cell.formula ?? null,
        // Semantic diff consumers expect the *effective* format for stored cells so inherited
        // row/col/sheet formatting is visible even when `styleId === 0`.
        format: effectiveStyle,
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
      const meta = this.getSheetMeta(id) ?? this.#defaultSheetMeta(id);
      /** @type {any} */
      const out = { id, frozenRows: view.frozenRows, frozenCols: view.frozenCols, cells };
      out.name = meta.name;
      out.visibility = meta.visibility;
      if (meta.tabColor) out.tabColor = cloneTabColor(meta.tabColor);
      if (view.colWidths && Object.keys(view.colWidths).length > 0) out.colWidths = view.colWidths;
      if (view.rowHeights && Object.keys(view.rowHeights).length > 0) out.rowHeights = view.rowHeights;
      if (Array.isArray(view.mergedRanges) && view.mergedRanges.length > 0) {
        out.mergedRanges = view.mergedRanges.map((r) => ({
          startRow: r.startRow,
          endRow: r.endRow,
          startCol: r.startCol,
          endCol: r.endCol,
        }));
      }

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

      // Range-run formatting (compressed rectangular formatting).
      const formatRunsByCol = Array.from(sheet?.formatRunsByCol?.entries?.() ?? [])
        .map(([col, runs]) => ({
          col,
          runs: Array.isArray(runs)
            ? runs.map((run) => ({
                startRow: run.startRow,
                endRowExclusive: run.endRowExclusive,
                format: this.styleTable.get(run.styleId),
              }))
            : [],
        }))
        .filter((entry) => entry.runs.length > 0);
      formatRunsByCol.sort((a, b) => a.col - b.col);
      if (formatRunsByCol.length > 0) out.formatRunsByCol = formatRunsByCol;
      return out;
    });

    // Include a redundant explicit `sheetOrder` so downstream snapshot consumers can preserve
    // ordering even if they manipulate/sort the `sheets` array.
    /** @type {any} */
    const snapshot = { schemaVersion: 1, sheetOrder: sheetIds, sheets };

    const images = Array.from(this.images.entries())
      .map(([id, entry]) => {
        /** @type {any} */
        const out = { id, bytesBase64: encodeBase64(entry.bytes) };
        if (entry && "mimeType" in entry) out.mimeType = entry.mimeType ?? null;
        return out;
      })
      .sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
    if (images.length > 0) snapshot.images = images;

    /** @type {Record<string, any[]>} */
    const drawingsBySheet = {};
    for (const sheetId of sheetIds) {
      const drawings = this.drawingsBySheet.get(sheetId);
      if (!Array.isArray(drawings) || drawings.length === 0) continue;
      drawingsBySheet[sheetId] = cloneJsonSerializable(drawings);
    }
    if (Object.keys(drawingsBySheet).length > 0) snapshot.drawingsBySheet = drawingsBySheet;

    return encodeUtf8(JSON.stringify(snapshot));
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
    let nextSheets = new Map();
    /** @type {Map<string, SheetViewState>} */
    let nextViews = new Map();
    /** @type {Map<string, { defaultStyleId: number, rowStyleIds: Map<number, number>, colStyleIds: Map<number, number> }>} */
    let nextFormats = new Map();
    /** @type {Map<string, Map<number, FormatRun[]>>} */
    let nextRangeRuns = new Map();
    /** @type {Map<string, SheetMetaState>} */
    const nextSheetMeta = new Map();
    /** @type {Map<string, any[]>} */
    let nextDrawingsBySheet = new Map();
    /** @type {Map<string, ImageEntry>} */
    const nextImages = new Map();

    const rawImages = Array.isArray(parsed?.images) ? parsed.images : [];
    for (const image of rawImages) {
      const id = typeof image?.id === "string" ? image.id.trim() : "";
      if (!id) continue;
      const bytesBase64 = typeof image?.bytesBase64 === "string" ? image.bytesBase64 : null;
      if (!bytesBase64) continue;
      let bytes;
      try {
        bytes = decodeBase64(bytesBase64);
      } catch {
        continue;
      }
      const mimeTypeRaw = "mimeType" in (image ?? {}) ? image.mimeType : undefined;
      let mimeType = undefined;
      if (mimeTypeRaw === undefined) {
        mimeType = undefined;
      } else if (mimeTypeRaw === null) {
        mimeType = null;
      } else if (typeof mimeTypeRaw === "string") {
        mimeType = mimeTypeRaw;
      } else {
        mimeType = null;
      }
      /** @type {ImageEntry} */
      const entry = { bytes };
      if (mimeTypeRaw !== undefined) entry.mimeType = mimeType;
      nextImages.set(id, entry);
    }

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

    const normalizeFormatRunsByCol = (raw) => {
      /** @type {Map<number, FormatRun[]>} */
      const out = new Map();
      if (!raw) return out;

      const addColRuns = (colKey, rawRuns) => {
        const col = Number(colKey);
        if (!Number.isInteger(col) || col < 0) return;
        if (!Array.isArray(rawRuns) || rawRuns.length === 0) return;
        /** @type {FormatRun[]} */
        const runs = [];
        for (const entry of rawRuns) {
          const startRow = Number(entry?.startRow);
          const endRowExclusiveNum = Number(entry?.endRowExclusive);
          const endRowNum = Number(entry?.endRow);
          const endRowExclusive = Number.isInteger(endRowExclusiveNum)
            ? endRowExclusiveNum
            : Number.isInteger(endRowNum)
              ? endRowNum + 1
              : NaN;
          if (!Number.isInteger(startRow) || startRow < 0) continue;
          if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;
          const format = entry?.format ?? null;
          const styleId = format == null ? 0 : this.styleTable.intern(format);
          if (styleId === 0) continue;
          runs.push({ startRow, endRowExclusive, styleId });
        }
        runs.sort((a, b) => a.startRow - b.startRow);
        const normalized = normalizeFormatRuns(runs);
        if (normalized.length > 0) out.set(col, normalized);
      };

      if (Array.isArray(raw)) {
        for (const entry of raw) {
          const col = entry?.col ?? entry?.index;
          const runs = entry?.runs ?? entry?.formatRuns ?? entry?.segments;
          addColRuns(col, runs);
        }
        return out;
      }

      if (typeof raw === "object") {
        for (const [key, value] of Object.entries(raw)) {
          addColRuns(key, value);
        }
      }

      return out;
    };
    for (const sheet of sheets) {
      if (!sheet?.id) continue;
      const rawName = typeof sheet?.name === "string" ? sheet.name : null;
      const name = rawName && rawName.trim() ? rawName.trim() : sheet.id;
      const visibilityRaw = sheet?.visibility;
      const visibility =
        visibilityRaw === "hidden" || visibilityRaw === "veryHidden" || visibilityRaw === "visible" ? visibilityRaw : "visible";
      const tabColorRaw = sheet?.tabColor;
      const tabColor =
        typeof tabColorRaw === "string" || isPlainObject(tabColorRaw) ? cloneTabColor(tabColorRaw) : undefined;
      nextSheetMeta.set(sheet.id, { name, visibility, ...(tabColor ? { tabColor } : {}) });

      const cellList = Array.isArray(sheet.cells) ? sheet.cells : [];
      const view = normalizeSheetViewState({
        frozenRows: sheet?.frozenRows,
        frozenCols: sheet?.frozenCols,
        colWidths: sheet?.colWidths,
        rowHeights: sheet?.rowHeights,
        mergedRanges: sheet?.mergedRanges ?? sheet?.merged_ranges ?? sheet?.mergedRegions ?? sheet?.merged_regions,
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
      nextRangeRuns.set(sheet.id, normalizeFormatRunsByCol(sheet?.formatRunsByCol));
    }

    const rawDrawingsBySheet = parsed?.drawingsBySheet;
    if (rawDrawingsBySheet && typeof rawDrawingsBySheet === "object" && !Array.isArray(rawDrawingsBySheet)) {
      for (const [rawSheetId, rawList] of Object.entries(rawDrawingsBySheet)) {
        if (typeof rawSheetId !== "string" || !rawSheetId) continue;
        if (!nextSheets.has(rawSheetId)) continue;
        if (!Array.isArray(rawList)) continue;

        /** @type {any[]} */
        const drawings = [];
        for (const raw of rawList) {
          if (!isJsonObject(raw)) continue;
          const drawingId = typeof raw.id === "string" ? raw.id.trim() : "";
          if (!drawingId) continue;
          if (!("anchor" in raw) || !("kind" in raw)) continue;
          const zOrder = Number(raw.zOrder);
          if (!Number.isFinite(zOrder)) continue;
          const cloned = cloneJsonSerializable(raw);
          cloned.id = drawingId;
          cloned.zOrder = zOrder;
          drawings.push(cloned);
        }
        if (drawings.length > 0) {
          nextDrawingsBySheet.set(rawSheetId, drawings);
        }
      }
    }

    // Prefer an explicit sheet order field when present, falling back to the ordering
    // of the `sheets` array itself (legacy behavior).
    const rawSheetOrder = Array.isArray(parsed?.sheetOrder) ? parsed.sheetOrder : [];
    if (rawSheetOrder.length > 0 && nextSheets.size > 0) {
      /** @type {string[]} */
      const desiredOrder = [];
      const seen = new Set();
      for (const raw of rawSheetOrder) {
        if (typeof raw !== "string") continue;
        const id = raw;
        if (seen.has(id)) continue;
        if (!nextSheets.has(id)) continue;
        seen.add(id);
        desiredOrder.push(id);
      }
      if (desiredOrder.length > 0) {
        for (const id of nextSheets.keys()) {
          if (seen.has(id)) continue;
          seen.add(id);
          desiredOrder.push(id);
        }

        const reorderBySheetIds = (map) => {
          const out = new Map();
          for (const id of desiredOrder) {
            if (!map.has(id)) continue;
            out.set(id, map.get(id));
          }
          return out;
        };

        nextSheets = reorderBySheetIds(nextSheets);
        nextViews = reorderBySheetIds(nextViews);
        nextFormats = reorderBySheetIds(nextFormats);
        nextRangeRuns = reorderBySheetIds(nextRangeRuns);
        nextDrawingsBySheet = reorderBySheetIds(nextDrawingsBySheet);
      }
    }

    const existingSheetIds = new Set(this.model.sheets.keys());
    const nextSheetIds = new Set(nextSheets.keys());
    const allSheetIds = new Set([...existingSheetIds, ...nextSheetIds]);
    const addedSheetIds = Array.from(nextSheetIds).filter((id) => !existingSheetIds.has(id));
    const removedSheetIds = Array.from(existingSheetIds).filter((id) => !nextSheetIds.has(id));
    const sheetStructureChanged = removedSheetIds.length > 0 || addedSheetIds.length > 0;

    /** @type {CellDelta[]} */
    const deltas = [];
    const contentChangedSheetIds = new Set();
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
        if (!cellContentEquals(before, after)) contentChangedSheetIds.add(sheetId);
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

    /** @type {RangeRunDelta[]} */
    const rangeRunDeltas = [];
    for (const sheetId of allSheetIds) {
      const existingSheet = this.model.sheets.get(sheetId);
      const beforeRunsByCol = existingSheet?.formatRunsByCol ?? new Map();
      const afterRunsByCol = nextRangeRuns.get(sheetId) ?? new Map();
      const colKeys = new Set([...beforeRunsByCol.keys(), ...afterRunsByCol.keys()]);
      for (const col of colKeys) {
        const beforeRuns = beforeRunsByCol.get(col) ?? [];
        const afterRuns = afterRunsByCol.get(col) ?? [];
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;
        rangeRunDeltas.push({
          sheetId,
          col,
          startRow: 0,
          endRowExclusive: EXCEL_MAX_ROWS,
          beforeRuns: beforeRuns.map(cloneFormatRun),
          afterRuns: afterRuns.map(cloneFormatRun),
        });
      }
    }

    /** @type {DrawingDelta[]} */
    const drawingDeltas = [];
    const drawingSheetIds = new Set([...this.drawingsBySheet.keys(), ...nextDrawingsBySheet.keys()]);
    for (const sheetId of drawingSheetIds) {
      const beforeRaw = this.drawingsBySheet.get(sheetId) ?? [];
      const afterRaw = nextDrawingsBySheet.get(sheetId) ?? [];
      const before = Array.isArray(beforeRaw) ? cloneJsonSerializable(beforeRaw) : [];
      const after = Array.isArray(afterRaw) ? cloneJsonSerializable(afterRaw) : [];
      if (jsonSerializableEquals(before, after)) continue;
      drawingDeltas.push({ sheetId, before, after });
    }
    drawingDeltas.sort((a, b) => (a.sheetId < b.sheetId ? -1 : a.sheetId > b.sheetId ? 1 : 0));

    /** @type {ImageDelta[]} */
    const imageDeltas = [];
    const imageIds = new Set([...this.images.keys(), ...nextImages.keys()]);
    for (const imageId of imageIds) {
      const beforeEntry = this.images.get(imageId) ?? null;
      const afterEntry = nextImages.get(imageId) ?? null;
      if (imageEntryEquals(beforeEntry, afterEntry)) continue;
      imageDeltas.push({
        imageId,
        before: beforeEntry ? cloneImageEntry(beforeEntry) : null,
        after: afterEntry ? cloneImageEntry(afterEntry) : null,
      });
    }
    imageDeltas.sort((a, b) => (a.imageId < b.imageId ? -1 : a.imageId > b.imageId ? 1 : 0));

    // Ensure all snapshot sheet ids exist even when they contain no cells (the model is otherwise
    // lazily materialized via reads/writes).
    for (const sheetId of nextSheetIds) {
      this.model.getCell(sheetId, 0, 0);
    }

    // Match the sheet Map iteration order to the snapshot ordering so sheet tab order
    // roundtrips through encodeState/applyState.
    //
    // Important: do this *before* emitting the `applyState` change event so listeners
    // (UI tabs, search adapters, etc) can observe the restored order synchronously.
    //
    // Include any sheets that will be removed (present in the existing model but not in
    // the snapshot) after the snapshot-ordered ids so they remain reachable until the
    // end of `applyState`.
    const orderedSheetIds = [
      ...Array.from(nextSheets.keys()),
      ...Array.from(existingSheetIds).filter((id) => !nextSheetIds.has(id)),
    ];
    for (const sheetId of orderedSheetIds) {
      const sheet = this.model.sheets.get(sheetId);
      if (!sheet) continue;
      // Re-insert to update insertion order without changing sheet identity.
      this.model.sheets.delete(sheetId);
      this.model.sheets.set(sheetId, sheet);
    }

    // Apply sheet metadata (name/visibility/tabColor) before emitting the `applyState` change event.
    // This lets UI layers read the restored sheet display names synchronously during the event.
    this.sheetMeta.clear();
    for (const sheetId of nextSheets.keys()) {
      const meta = nextSheetMeta.get(sheetId) ?? this.#defaultSheetMeta(sheetId);
      this.sheetMeta.set(sheetId, cloneSheetMetaState(meta));
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
    this.#applyEdits(deltas, sheetViewDeltas, formatDeltas, rangeRunDeltas, drawingDeltas, imageDeltas, [], null, {
      recalc: false,
      emitChange: true,
      source: "applyState",
      sheetStructureChanged,
    });
    this.engine?.endBatch?.();
    this.engine?.recalculate();

    // `applyState` can create/delete sheets without emitting any cell deltas (e.g. empty sheets).
    // Treat those as content changes for sheet-level caches.
    for (const sheetId of addedSheetIds) {
      if (!contentChangedSheetIds.has(sheetId)) {
        this.contentVersionBySheet.set(sheetId, (this.contentVersionBySheet.get(sheetId) ?? 0) + 1);
      }
    }
    for (const sheetId of removedSheetIds) {
      if (!contentChangedSheetIds.has(sheetId)) {
        this.contentVersionBySheet.set(sheetId, (this.contentVersionBySheet.get(sheetId) ?? 0) + 1);
      }
    }

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
    this.#applyEdits(deltas, [], [], [], [], [], [], null, { recalc, emitChange: true, source });

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
    this.#applyEdits([], deltas, [], [], [], [], [], null, { recalc: false, emitChange: true, source });

    // Mark dirty even though we didn't advance the undo cursor.
    if (options.markDirty !== false) {
      this.savedCursor = null;
    }
    this.#emitDirty();
  }

  /**
   * Apply a set of drawing deltas that originated externally (e.g. collaboration sync).
   *
   * Like {@link applyExternalSheetViewDeltas}, these updates:
   * - bypass undo/redo history (not user-editable)
   * - emit `change` + `update` events so UI layers can react
   * - mark the document dirty by default
   *
   * @param {DrawingDelta[]} deltas
   * @param {{ source?: string, markDirty?: boolean }} [options]
   */
  applyExternalDrawingDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const source = typeof options.source === "string" ? options.source : undefined;
    this.#applyEdits([], [], [], [], deltas, [], [], null, { recalc: false, emitChange: true, source });

    // Mark dirty even though we didn't advance the undo cursor.
    if (options.markDirty !== false) {
      this.savedCursor = null;
    }
    this.#emitDirty();
  }

  /**
   * Apply a set of layered formatting deltas that originated externally (e.g. collaboration sync).
   *
   * Like {@link applyExternalDeltas}, these updates:
   * - bypass undo/redo history (not user-editable)
   * - bypass `canEditCell` (permissions should be enforced at the collaboration layer)
   * - emit `change` + `update` events so UI + versioning layers can react
   * - mark the document dirty by default
   *
   * @param {FormatDelta[]} deltas
   * @param {{ source?: string, markDirty?: boolean }} [options]
   */
  applyExternalFormatDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const source = typeof options.source === "string" ? options.source : undefined;
    this.#applyEdits([], [], deltas, [], [], [], [], null, { recalc: false, emitChange: true, source });

    // Mark dirty even though we didn't advance the undo cursor.
    if (options.markDirty !== false) {
      this.savedCursor = null;
    }
    this.#emitDirty();
  }

  /**
   * Apply a set of range-run formatting deltas that originated externally (e.g. collaboration sync).
   *
   * Like {@link applyExternalFormatDeltas}, these updates:
   * - bypass `canEditCell` (permissions should be enforced at the collaboration layer)
   * - bypass undo/redo history (not user-editable)
   * - emit `change` + `update` events so UI + versioning layers can react
   * - mark the document dirty by default
   *
   * @param {RangeRunDelta[]} deltas
   * @param {{ source?: string, markDirty?: boolean }} [options]
   */
  applyExternalRangeRunDeltas(deltas, options = {}) {
    if (!deltas || deltas.length === 0) return;

    // External updates should never merge with user edits.
    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const source = typeof options.source === "string" ? options.source : undefined;
    this.#applyEdits([], [], [], deltas, [], [], [], null, { recalc: false, emitChange: true, source });

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
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
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
      (batch.deltasByCell.size === 0 &&
        batch.deltasBySheetView.size === 0 &&
        batch.deltasByFormat.size === 0 &&
        batch.deltasByRangeRun.size === 0 &&
        batch.deltasByDrawing.size === 0 &&
        batch.deltasByImage.size === 0 &&
        batch.deltasBySheetMeta.size === 0 &&
        batch.sheetOrderDelta == null)
    ) {
      this.#emitHistory();
      this.#emitDirty();
      return;
    }

    this.#commitHistoryEntry(batch);

    // Only recalculate for batches that included cell input edits. Sheet view changes
    // (frozen panes, row/col sizes, etc.) do not affect formula results.
    const shouldRecalc =
      cellDeltasAffectRecalc(Array.from(batch.deltasByCell.values())) ||
      sheetMetaDeltasAffectRecalc(Array.from(batch.deltasBySheetMeta.values()));
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
        rangeRunDeltas: [],
        drawingDeltas: [],
        imageDeltas: [],
        sheetMetaDeltas: [],
        sheetOrderDelta: null,
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
      batch &&
        (batch.deltasByCell.size > 0 ||
          batch.deltasBySheetView.size > 0 ||
          batch.deltasByFormat.size > 0 ||
          batch.deltasByRangeRun.size > 0 ||
          batch.deltasByDrawing.size > 0 ||
          batch.deltasByImage.size > 0 ||
          batch.deltasBySheetMeta.size > 0 ||
          batch.sheetOrderDelta != null),
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
      const inverseRangeRuns = invertRangeRunDeltas(entryRangeRunDeltas(batch));
      const inverseDrawings = invertDrawingDeltas(entryDrawingDeltas(batch));
      const inverseImages = invertImageDeltas(entryImageDeltas(batch));
      const inverseSheetMeta = invertSheetMetaDeltas(entrySheetMetaDeltas(batch));
      const inverseSheetOrder = invertSheetOrderDelta(batch.sheetOrderDelta);
      const sheetStructureChanged = sheetMetaDeltasAffectRecalc(inverseSheetMeta);
      this.#applyEdits(
        inverseCells,
        inverseViews,
        inverseFormats,
        inverseRangeRuns,
        inverseDrawings,
        inverseImages,
        inverseSheetMeta,
        inverseSheetOrder,
        {
          recalc: false,
          emitChange: true,
          source: "cancelBatch",
          sheetStructureChanged,
        },
      );
    }

    this.engine?.endBatch?.();

    // Only recalculate when canceling a batch that mutated cell inputs or sheet structure.
    const shouldRecalc = Boolean(
      batch &&
        (cellDeltasAffectRecalc(Array.from(batch.deltasByCell.values())) ||
          sheetMetaDeltasAffectRecalc(Array.from(batch.deltasBySheetMeta.values()))),
    );
    if (shouldRecalc) {
      this.engine?.recalculate();
      this.#emit("change", {
        deltas: [],
        sheetViewDeltas: [],
        formatDeltas: [],
        rowStyleDeltas: [],
        colStyleDeltas: [],
        sheetStyleDeltas: [],
        rangeRunDeltas: [],
        drawingDeltas: [],
        imageDeltas: [],
        sheetMetaDeltas: [],
        sheetOrderDelta: null,
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
    const rangeRunDeltas = entryRangeRunDeltas(entry);
    const drawingDeltas = entryDrawingDeltas(entry);
    const imageDeltas = entryImageDeltas(entry);
    const sheetMetaDeltas = entrySheetMetaDeltas(entry);
    const sheetOrderDelta = cloneSheetOrderDelta(entry.sheetOrderDelta);
    const inverseCells = invertDeltas(cellDeltas);
    const inverseViews = invertSheetViewDeltas(viewDeltas);
    const inverseFormats = invertFormatDeltas(formatDeltas);
    const inverseRangeRuns = invertRangeRunDeltas(rangeRunDeltas);
    const inverseDrawings = invertDrawingDeltas(drawingDeltas);
    const inverseImages = invertImageDeltas(imageDeltas);
    const inverseSheetMeta = invertSheetMetaDeltas(sheetMetaDeltas);
    const inverseSheetOrder = invertSheetOrderDelta(sheetOrderDelta);
    this.cursor -= 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const sheetStructureChanged = sheetMetaDeltasAffectRecalc(sheetMetaDeltas);
    const shouldRecalc = cellDeltasAffectRecalc(cellDeltas) || sheetStructureChanged;
    this.#applyEdits(
      inverseCells,
      inverseViews,
      inverseFormats,
      inverseRangeRuns,
      inverseDrawings,
      inverseImages,
      inverseSheetMeta,
      inverseSheetOrder,
      {
      recalc: shouldRecalc,
      emitChange: true,
      source: "undo",
      sheetStructureChanged,
      },
    );
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
    const rangeRunDeltas = entryRangeRunDeltas(entry);
    const drawingDeltas = entryDrawingDeltas(entry);
    const imageDeltas = entryImageDeltas(entry);
    const sheetMetaDeltas = entrySheetMetaDeltas(entry);
    const sheetOrderDelta = cloneSheetOrderDelta(entry.sheetOrderDelta);
    this.cursor += 1;

    this.lastMergeKey = null;
    this.lastMergeTime = 0;

    const sheetStructureChanged = sheetMetaDeltasAffectRecalc(sheetMetaDeltas);
    const shouldRecalc = cellDeltasAffectRecalc(cellDeltas) || sheetStructureChanged;
    this.#applyEdits(
      cellDeltas,
      viewDeltas,
      formatDeltas,
      rangeRunDeltas,
      drawingDeltas,
      imageDeltas,
      sheetMetaDeltas,
      sheetOrderDelta,
      {
      recalc: shouldRecalc,
      emitChange: true,
      source: "redo",
      sheetStructureChanged,
      },
    );
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
    this.#applyEdits(deltas, [], [], [], [], [], [], null, { recalc: shouldRecalc, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(deltas, [], [], [], [], [], [], null, options);
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
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
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
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
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
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
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
   * @param {RangeRunDelta[]} deltas
   */
  #mergeRangeRunIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
      };
    }
    for (const delta of deltas) {
      const key = rangeRunKey(delta.sheetId, delta.col);
      const existing = this.activeBatch.deltasByRangeRun.get(key);
      if (!existing) {
        this.activeBatch.deltasByRangeRun.set(key, cloneRangeRunDelta(delta));
      } else {
        existing.afterRuns = Array.isArray(delta.afterRuns) ? delta.afterRuns.map(cloneFormatRun) : [];
        existing.startRow = Math.min(existing.startRow, delta.startRow);
        existing.endRowExclusive = Math.max(existing.endRowExclusive, delta.endRowExclusive);
      }
    }
  }

  /**
   * @param {DrawingDelta[]} deltas
   */
  #mergeDrawingsIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
      };
    }
    for (const delta of deltas) {
      const existing = this.activeBatch.deltasByDrawing.get(delta.sheetId);
      if (!existing) {
        this.activeBatch.deltasByDrawing.set(delta.sheetId, cloneDrawingDelta(delta));
      } else {
        existing.after = Array.isArray(delta.after) ? cloneJsonSerializable(delta.after) : [];
      }
    }
  }

  /**
   * @param {ImageDelta[]} deltas
   */
  #mergeImagesIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
      };
    }
    for (const delta of deltas) {
      const existing = this.activeBatch.deltasByImage.get(delta.imageId);
      if (!existing) {
        this.activeBatch.deltasByImage.set(delta.imageId, cloneImageDelta(delta));
      } else {
        existing.after = delta.after ? cloneImageEntry(delta.after) : null;
      }
    }
  }

  /**
   * @param {SheetMetaDelta[]} deltas
   */
  #mergeSheetMetaIntoBatch(deltas) {
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
      };
    }
    for (const delta of deltas) {
      const existing = this.activeBatch.deltasBySheetMeta.get(delta.sheetId);
      if (!existing) {
        this.activeBatch.deltasBySheetMeta.set(delta.sheetId, cloneSheetMetaDelta(delta));
      } else {
        existing.after = delta.after ? cloneSheetMetaState(delta.after) : null;
      }
    }
  }

  /**
   * @param {SheetOrderDelta | null | undefined} delta
   */
  #mergeSheetOrderIntoBatch(delta) {
    if (!delta) return;
    if (!this.activeBatch) {
      this.activeBatch = {
        timestamp: Date.now(),
        deltasByCell: new Map(),
        deltasBySheetView: new Map(),
        deltasByFormat: new Map(),
        deltasByRangeRun: new Map(),
        deltasByDrawing: new Map(),
        deltasByImage: new Map(),
        deltasBySheetMeta: new Map(),
        sheetOrderDelta: null,
      };
    }

    if (!this.activeBatch.sheetOrderDelta) {
      this.activeBatch.sheetOrderDelta = cloneSheetOrderDelta(delta);
      return;
    }
    this.activeBatch.sheetOrderDelta.after = Array.isArray(delta.after) ? delta.after.slice() : [];
  }

  /**
   * @param {SheetViewDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserSheetViewDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    this.#applyEdits([], deltas, [], [], [], [], [], null, { recalc: false, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeSheetViewIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry([], deltas, [], [], [], [], [], null, options);
  }

  /**
   * @param {FormatDelta[]} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserFormatDeltas(deltas, options) {
    if (!deltas || deltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    this.#applyEdits([], [], deltas, [], [], [], [], null, { recalc: false, emitChange: true, source });

    if (this.batchDepth > 0) {
      this.#mergeFormatIntoBatch(deltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry([], [], deltas, [], [], [], [], null, options);
  }

  /**
   * Apply combined per-cell style deltas and range-run formatting deltas.
   *
   * This is used by `setRangeFormat()` for large non-full-row/col rectangles to avoid O(area)
   * cell materialization.
   *
   * @param {CellDelta[]} cellDeltas
   * @param {RangeRunDelta[]} rangeRunDeltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserRangeRunDeltas(cellDeltas, rangeRunDeltas, options) {
    const hasCells = Array.isArray(cellDeltas) && cellDeltas.length > 0;
    const hasRuns = Array.isArray(rangeRunDeltas) && rangeRunDeltas.length > 0;
    if (!hasCells && !hasRuns) return;

    if (hasCells && this.canEditCell) {
      cellDeltas = cellDeltas.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
    }

    const filteredHasCells = Array.isArray(cellDeltas) && cellDeltas.length > 0;
    if (!filteredHasCells && !hasRuns) return;

    const source = typeof options?.source === "string" ? options.source : undefined;

    // Formatting changes should never trigger formula recalc.
    this.#applyEdits(cellDeltas ?? [], [], [], rangeRunDeltas ?? [], [], [], [], null, {
      recalc: false,
      emitChange: true,
      source,
    });

    if (this.batchDepth > 0) {
      if (filteredHasCells) this.#mergeIntoBatch(cellDeltas);
      if (hasRuns) this.#mergeRangeRunIntoBatch(rangeRunDeltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(cellDeltas ?? [], [], [], rangeRunDeltas ?? [], [], [], [], null, options);
  }

  /**
   * Apply a set of cell deltas, layered format deltas (sheet/row/col), and range-run deltas as a
   * single user edit (one change event / one undo step), merging into an active batch if present.
   *
   * This is primarily used for range formatting operations that need to update multiple formatting
   * layers at once (e.g. full-column formatting should also override any existing range-run formatting
   * in that column).
   *
   * @param {CellDelta[]} cellDeltas
   * @param {FormatDelta[]} formatDeltas
   * @param {RangeRunDelta[]} rangeRunDeltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} options
   */
  #applyUserCellAndFormatDeltas(cellDeltas, formatDeltas, rangeRunDeltas, options) {
    cellDeltas = Array.isArray(cellDeltas) ? cellDeltas : [];
    formatDeltas = Array.isArray(formatDeltas) ? formatDeltas : [];
    rangeRunDeltas = Array.isArray(rangeRunDeltas) ? rangeRunDeltas : [];

    if (cellDeltas.length > 0 && this.canEditCell) {
      cellDeltas = cellDeltas.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
    }

    if (cellDeltas.length === 0 && formatDeltas.length === 0 && rangeRunDeltas.length === 0) return;

    const source = typeof options?.source === "string" ? options.source : undefined;
    const shouldRecalc = this.batchDepth === 0 && cellDeltasAffectRecalc(cellDeltas);
    this.#applyEdits(cellDeltas, [], formatDeltas, rangeRunDeltas, [], [], [], null, {
      recalc: shouldRecalc,
      emitChange: true,
      source,
    });

    if (this.batchDepth > 0) {
      if (cellDeltas.length > 0) this.#mergeIntoBatch(cellDeltas);
      if (formatDeltas.length > 0) this.#mergeFormatIntoBatch(formatDeltas);
      if (rangeRunDeltas.length > 0) this.#mergeRangeRunIntoBatch(rangeRunDeltas);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(cellDeltas, [], formatDeltas, rangeRunDeltas, [], [], [], null, options);
  }

  /**
   * Apply an arbitrary mix of workbook deltas as a single undoable user edit.
   *
   * This is the common plumbing used by sheet metadata operations (rename/reorder/hide/tab color/add/delete)
   * so they participate in the same history batching + merge logic as cell edits.
   *
   * @param {{
   *   cellDeltas?: CellDelta[],
   *   sheetViewDeltas?: SheetViewDelta[],
   *   formatDeltas?: FormatDelta[],
   *   rangeRunDeltas?: RangeRunDelta[],
   *   drawingDeltas?: DrawingDelta[],
   *   imageDeltas?: ImageDelta[],
   *   sheetMetaDeltas?: SheetMetaDelta[],
   *   sheetOrderDelta?: SheetOrderDelta | null,
   * }} deltas
   * @param {{ label?: string, mergeKey?: string, source?: string }} [options]
   */
  #applyUserWorkbookEdits(deltas, options = {}) {
    const cellDeltas = Array.isArray(deltas?.cellDeltas) ? deltas.cellDeltas : [];
    const sheetViewDeltas = Array.isArray(deltas?.sheetViewDeltas) ? deltas.sheetViewDeltas : [];
    const formatDeltas = Array.isArray(deltas?.formatDeltas) ? deltas.formatDeltas : [];
    const rangeRunDeltas = Array.isArray(deltas?.rangeRunDeltas) ? deltas.rangeRunDeltas : [];
    const drawingDeltas = Array.isArray(deltas?.drawingDeltas) ? deltas.drawingDeltas : [];
    const imageDeltas = Array.isArray(deltas?.imageDeltas) ? deltas.imageDeltas : [];
    const sheetMetaDeltas = Array.isArray(deltas?.sheetMetaDeltas) ? deltas.sheetMetaDeltas : [];
    const sheetOrderDelta = deltas?.sheetOrderDelta ? cloneSheetOrderDelta(deltas.sheetOrderDelta) : null;

    let filteredCells = cellDeltas;
    if (filteredCells.length > 0 && this.canEditCell) {
      filteredCells = filteredCells.filter((delta) =>
        this.canEditCell({ sheetId: delta.sheetId, row: delta.row, col: delta.col })
      );
    }

    const hasCells = filteredCells.length > 0;
    const hasViews = sheetViewDeltas.length > 0;
    const hasFormats = formatDeltas.length > 0;
    const hasRangeRuns = rangeRunDeltas.length > 0;
    const hasDrawings = drawingDeltas.length > 0;
    const hasImages = imageDeltas.length > 0;
    const hasSheetMeta = sheetMetaDeltas.length > 0;
    const hasSheetOrder = Boolean(sheetOrderDelta);
    if (!hasCells && !hasViews && !hasFormats && !hasRangeRuns && !hasDrawings && !hasImages && !hasSheetMeta && !hasSheetOrder) {
      return;
    }

    const source = typeof options?.source === "string" ? options.source : undefined;
    const sheetStructureChanged = sheetMetaDeltasAffectRecalc(sheetMetaDeltas);
    const shouldRecalc =
      this.batchDepth === 0 && (cellDeltasAffectRecalc(filteredCells) || sheetStructureChanged);

    this.#applyEdits(filteredCells, sheetViewDeltas, formatDeltas, rangeRunDeltas, drawingDeltas, imageDeltas, sheetMetaDeltas, sheetOrderDelta, {
      recalc: shouldRecalc,
      emitChange: true,
      source,
      sheetStructureChanged,
    });

    if (this.batchDepth > 0) {
      if (hasCells) this.#mergeIntoBatch(filteredCells);
      if (hasViews) this.#mergeSheetViewIntoBatch(sheetViewDeltas);
      if (hasFormats) this.#mergeFormatIntoBatch(formatDeltas);
      if (hasRangeRuns) this.#mergeRangeRunIntoBatch(rangeRunDeltas);
      if (hasDrawings) this.#mergeDrawingsIntoBatch(drawingDeltas);
      if (hasImages) this.#mergeImagesIntoBatch(imageDeltas);
      if (hasSheetMeta) this.#mergeSheetMetaIntoBatch(sheetMetaDeltas);
      if (hasSheetOrder) this.#mergeSheetOrderIntoBatch(sheetOrderDelta);
      this.#emitDirty();
      return;
    }

    this.#commitOrMergeHistoryEntry(
      filteredCells,
      sheetViewDeltas,
      formatDeltas,
      rangeRunDeltas,
      drawingDeltas,
      imageDeltas,
      sheetMetaDeltas,
      sheetOrderDelta,
      options,
    );
  }

  /**
   * @param {CellDelta[]} cellDeltas
   * @param {SheetViewDelta[]} sheetViewDeltas
   * @param {FormatDelta[]} formatDeltas
   * @param {RangeRunDelta[]} rangeRunDeltas
   * @param {DrawingDelta[]} drawingDeltas
   * @param {ImageDelta[]} imageDeltas
   * @param {SheetMetaDelta[]} sheetMetaDeltas
   * @param {SheetOrderDelta | null} sheetOrderDelta
   * @param {{ label?: string, mergeKey?: string }} options
   */
  #commitOrMergeHistoryEntry(
    cellDeltas,
    sheetViewDeltas,
    formatDeltas,
    rangeRunDeltas,
    drawingDeltas,
    imageDeltas,
    sheetMetaDeltas,
    sheetOrderDelta,
    options,
  ) {
    drawingDeltas = Array.isArray(drawingDeltas) ? drawingDeltas : [];
    imageDeltas = Array.isArray(imageDeltas) ? imageDeltas : [];
    sheetMetaDeltas = Array.isArray(sheetMetaDeltas) ? sheetMetaDeltas : [];
    sheetOrderDelta = sheetOrderDelta ? cloneSheetOrderDelta(sheetOrderDelta) : null;
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

      for (const delta of rangeRunDeltas) {
        const key = rangeRunKey(delta.sheetId, delta.col);
        const existing = entry.deltasByRangeRun.get(key);
        if (!existing) {
          entry.deltasByRangeRun.set(key, cloneRangeRunDelta(delta));
        } else {
          existing.afterRuns = Array.isArray(delta.afterRuns) ? delta.afterRuns.map(cloneFormatRun) : [];
          existing.startRow = Math.min(existing.startRow, delta.startRow);
          existing.endRowExclusive = Math.max(existing.endRowExclusive, delta.endRowExclusive);
        }
      }

      for (const delta of drawingDeltas) {
        const existing = entry.deltasByDrawing.get(delta.sheetId);
        if (!existing) {
          entry.deltasByDrawing.set(delta.sheetId, cloneDrawingDelta(delta));
        } else {
          existing.after = Array.isArray(delta.after) ? cloneJsonSerializable(delta.after) : [];
        }
      }

      for (const delta of imageDeltas) {
        const existing = entry.deltasByImage.get(delta.imageId);
        if (!existing) {
          entry.deltasByImage.set(delta.imageId, cloneImageDelta(delta));
        } else {
          existing.after = delta.after ? cloneImageEntry(delta.after) : null;
        }
      }

      for (const delta of sheetMetaDeltas) {
        const existing = entry.deltasBySheetMeta.get(delta.sheetId);
        if (!existing) {
          entry.deltasBySheetMeta.set(delta.sheetId, cloneSheetMetaDelta(delta));
        } else {
          existing.after = delta.after ? cloneSheetMetaState(delta.after) : null;
        }
      }

      if (sheetOrderDelta) {
        if (!entry.sheetOrderDelta) {
          entry.sheetOrderDelta = cloneSheetOrderDelta(sheetOrderDelta);
        } else {
          entry.sheetOrderDelta.after = sheetOrderDelta.after.slice();
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
      deltasByRangeRun: new Map(),
      deltasByDrawing: new Map(),
      deltasByImage: new Map(),
      deltasBySheetMeta: new Map(),
      sheetOrderDelta: null,
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

    for (const delta of rangeRunDeltas) {
      entry.deltasByRangeRun.set(rangeRunKey(delta.sheetId, delta.col), cloneRangeRunDelta(delta));
    }

    for (const delta of drawingDeltas) {
      entry.deltasByDrawing.set(delta.sheetId, cloneDrawingDelta(delta));
    }

    for (const delta of imageDeltas) {
      entry.deltasByImage.set(delta.imageId, cloneImageDelta(delta));
    }

    for (const delta of sheetMetaDeltas) {
      entry.deltasBySheetMeta.set(delta.sheetId, cloneSheetMetaDelta(delta));
    }

    if (sheetOrderDelta) {
      entry.sheetOrderDelta = cloneSheetOrderDelta(sheetOrderDelta);
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
    if (
      entry.deltasByCell.size === 0 &&
      entry.deltasBySheetView.size === 0 &&
      entry.deltasByFormat.size === 0 &&
      entry.deltasByRangeRun.size === 0 &&
      entry.deltasByDrawing.size === 0 &&
      entry.deltasByImage.size === 0 &&
      entry.deltasBySheetMeta.size === 0 &&
      entry.sheetOrderDelta == null
    ) {
      return;
    }
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
   * @param {RangeRunDelta[]} rangeRunDeltas
   * @param {DrawingDelta[]} drawingDeltas
   * @param {ImageDelta[]} imageDeltas
   * @param {SheetMetaDelta[]} sheetMetaDeltas
   * @param {SheetOrderDelta | null} sheetOrderDelta
   * @param {{ recalc: boolean, emitChange: boolean, source?: string, sheetStructureChanged?: boolean }} options
   */
  #applyEdits(
    cellDeltas,
    sheetViewDeltas,
    formatDeltas,
    rangeRunDeltas,
    drawingDeltas,
    imageDeltas,
    sheetMetaDeltas,
    sheetOrderDelta,
    options,
  ) {
    const contentChangedSheetIds = new Set();
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
    for (const delta of rangeRunDeltas) {
      // Ensure sheet exists for format-only changes.
      this.model.getCell(delta.sheetId, 0, 0);
      const sheet = this.model.sheets.get(delta.sheetId);
      if (!sheet) continue;
      const col = Number(delta.col);
      if (!Number.isInteger(col) || col < 0) continue;
      const nextRuns = Array.isArray(delta.afterRuns) ? delta.afterRuns.map(cloneFormatRun) : [];
      sheet.setFormatRunsForCol(col, nextRuns);
    }
    for (const delta of sheetViewDeltas) {
      this.model.setSheetView(delta.sheetId, delta.after);
    }
    for (const delta of cellDeltas) {
      this.model.setCell(delta.sheetId, delta.row, delta.col, delta.after);
      if (!cellContentEquals(delta.before, delta.after)) contentChangedSheetIds.add(delta.sheetId);
    }

    for (const delta of drawingDeltas) {
      if (!delta) continue;
      const sheetId = delta.sheetId;
      const after = Array.isArray(delta.after) ? cloneJsonSerializable(delta.after) : [];
      if (after.length > 0) {
        // Drawings are sheet-scoped; materialize the sheet if needed.
        this.model.getCell(sheetId, 0, 0);
        this.drawingsBySheet.set(sheetId, after);
      } else {
        this.drawingsBySheet.delete(sheetId);
      }
    }

    for (const delta of imageDeltas) {
      if (!delta) continue;
      const imageId = delta.imageId;
      if (!delta.after) {
        this.images.delete(imageId);
      } else {
        this.images.set(imageId, cloneImageEntry(delta.after));
      }
    }

    for (const delta of sheetMetaDeltas) {
      if (!delta) continue;
      if (delta.after == null) {
        // Sheet deletion.
        this.sheetMeta.delete(delta.sheetId);
        this.model.sheets.delete(delta.sheetId);
        this.drawingsBySheet.delete(delta.sheetId);
        continue;
      }

      // Sheet add / metadata update.
      this.model.getCell(delta.sheetId, 0, 0);
      this.sheetMeta.set(delta.sheetId, cloneSheetMetaState(delta.after));
    }

    if (sheetOrderDelta) {
      const desired = Array.isArray(sheetOrderDelta.after) ? sheetOrderDelta.after : [];
      const existing = this.model.sheets;
      const snapshot = new Map(existing);
      const seen = new Set();
      existing.clear();
      for (const id of desired) {
        if (!id || seen.has(id)) continue;
        const sheet = snapshot.get(id);
        if (!sheet) continue;
        existing.set(id, sheet);
        seen.add(id);
      }
      // Preserve any sheets that were omitted from the order delta (defensive).
      for (const [id, sheet] of snapshot.entries()) {
        if (seen.has(id)) continue;
        existing.set(id, sheet);
      }
    }

    // Sheet metadata changes should not bump contentVersion unless they add/remove sheets.
    // `sheetStructureChanged` covers:
    // - applyState sheet add/remove
    // - undoable sheet add/delete operations
    const shouldBumpContentVersion =
      Boolean(options?.sheetStructureChanged) ||
      contentChangedSheetIds.size > 0 ||
      (Array.isArray(drawingDeltas) && drawingDeltas.length > 0) ||
      (Array.isArray(imageDeltas) && imageDeltas.length > 0);

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
      for (const delta of rangeRunDeltas) {
        this.model.getCell(delta.sheetId, 0, 0);
        const sheet = this.model.sheets.get(delta.sheetId);
        if (!sheet) continue;
        const col = Number(delta.col);
        if (!Number.isInteger(col) || col < 0) continue;
        const beforeRuns = Array.isArray(delta.beforeRuns) ? delta.beforeRuns.map(cloneFormatRun) : [];
        sheet.setFormatRunsForCol(col, beforeRuns);
      }
      for (const delta of sheetViewDeltas) {
        this.model.setSheetView(delta.sheetId, delta.before);
      }
      for (const delta of cellDeltas) {
        this.model.setCell(delta.sheetId, delta.row, delta.col, delta.before);
      }

      for (const delta of drawingDeltas) {
        if (!delta) continue;
        const sheetId = delta.sheetId;
        const before = Array.isArray(delta.before) ? cloneJsonSerializable(delta.before) : [];
        if (before.length > 0) {
          this.model.getCell(sheetId, 0, 0);
          this.drawingsBySheet.set(sheetId, before);
        } else {
          this.drawingsBySheet.delete(sheetId);
        }
      }

      for (const delta of imageDeltas) {
        if (!delta) continue;
        const imageId = delta.imageId;
        if (!delta.before) {
          this.images.delete(imageId);
        } else {
          this.images.set(imageId, cloneImageEntry(delta.before));
        }
      }

      for (const delta of sheetMetaDeltas) {
        if (!delta) continue;
        if (delta.before == null) {
          this.sheetMeta.delete(delta.sheetId);
          this.model.sheets.delete(delta.sheetId);
          this.drawingsBySheet.delete(delta.sheetId);
          continue;
        }
        this.model.getCell(delta.sheetId, 0, 0);
        this.sheetMeta.set(delta.sheetId, cloneSheetMetaState(delta.before));
      }

      if (sheetOrderDelta) {
        const desired = Array.isArray(sheetOrderDelta.before) ? sheetOrderDelta.before : [];
        const existing = this.model.sheets;
        const snapshot = new Map(existing);
        const seen = new Set();
        existing.clear();
        for (const id of desired) {
          if (!id || seen.has(id)) continue;
          const sheet = snapshot.get(id);
          if (!sheet) continue;
          existing.set(id, sheet);
          seen.add(id);
        }
        for (const [id, sheet] of snapshot.entries()) {
          if (seen.has(id)) continue;
          existing.set(id, sheet);
        }
      }
      throw err;
    }

    // Update versions before emitting events so observers can synchronously read the latest value.
    this._updateVersion += 1;
    if (shouldBumpContentVersion) this._contentVersion += 1;

    // Bump per-sheet content versions (only after the engine accepts the changes).
    for (const sheetId of contentChangedSheetIds) {
      this.contentVersionBySheet.set(sheetId, (this.contentVersionBySheet.get(sheetId) ?? 0) + 1);
    }

    if (options.emitChange) {
      /** @type {any[]} */
      const rowStyleDeltas = [];
      /** @type {any[]} */
      const colStyleDeltas = [];
      /** @type {any[]} */
      const sheetStyleDeltas = [];
      for (const d of formatDeltas) {
        if (!d) continue;
        if (d.layer === "row" && d.index != null) {
          rowStyleDeltas.push({
            sheetId: d.sheetId,
            row: d.index,
            beforeStyleId: d.beforeStyleId,
            afterStyleId: d.afterStyleId,
          });
        } else if (d.layer === "col" && d.index != null) {
          colStyleDeltas.push({
            sheetId: d.sheetId,
            col: d.index,
            beforeStyleId: d.beforeStyleId,
            afterStyleId: d.afterStyleId,
          });
        } else if (d.layer === "sheet") {
          sheetStyleDeltas.push({
            sheetId: d.sheetId,
            beforeStyleId: d.beforeStyleId,
            afterStyleId: d.afterStyleId,
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
        rangeRunDeltas: rangeRunDeltas.map(cloneRangeRunDelta),
        drawingDeltas: drawingDeltas.map(cloneDrawingDelta),
        imageDeltas: imageDeltas.map((d) => ({
          imageId: d.imageId,
          before: d.before
            ? { mimeType: ("mimeType" in d.before ? d.before.mimeType : null) ?? null, byteLength: d.before.bytes?.length ?? 0 }
            : null,
          after: d.after
            ? { mimeType: ("mimeType" in d.after ? d.after.mimeType : null) ?? null, byteLength: d.after.bytes?.length ?? 0 }
            : null,
        })),
        sheetMetaDeltas: sheetMetaDeltas.map(cloneSheetMetaDelta),
        sheetOrderDelta: sheetOrderDelta ? cloneSheetOrderDelta(sheetOrderDelta) : null,
        recalc: options.recalc,
      };
      if (options.source) payload.source = options.source;
      this.#emit("change", payload);
    }

    this.#emit("update", { version: this.updateVersion });
  }
}
