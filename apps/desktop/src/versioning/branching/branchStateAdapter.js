import { formatA1, parseA1 } from "../../document/coords.js";
import { normalizeDocumentState } from "../../../../../packages/versioning/branches/src/state.js";

/**
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").DocumentState} DocumentState
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").Cell} BranchCell
 */

const structuredCloneFn =
  typeof globalThis.structuredClone === "function" ? globalThis.structuredClone : null;

// Collab masking (permissions/encryption) renders unreadable cells as a constant
// placeholder. Branching should treat these as "unknown" rather than persisting
// the placeholder as real content.
const MASKED_CELL_VALUE = "###";

/**
 * Best-effort access to DocumentController sheet metadata.
 *
 * Task 201 adds authoritative sheet metadata (order + display names) to
 * DocumentController. This adapter needs to work both before and after that
 * change lands, so we feature-detect common access patterns.
 *
 * @param {any} doc
 * @param {string} sheetId
 * @returns {any}
 */
function getDocumentControllerSheetMeta(doc, sheetId) {
  if (!doc || typeof doc !== "object") return null;

  // Preferred API (post-Task 201).
  if (typeof doc.getSheetMeta === "function") {
    try {
      return doc.getSheetMeta(sheetId);
    } catch {
      return null;
    }
  }

  // Alternate method naming (defensive).
  if (typeof doc.getSheetMetadata === "function") {
    try {
      return doc.getSheetMetadata(sheetId);
    } catch {
      return null;
    }
  }

  // Fallback: direct map/object on the controller instance.
  const byId = doc.sheetMetaById ?? doc.sheetMetadataById ?? doc.sheetMeta ?? null;
  if (byId && typeof byId === "object") {
    if (byId instanceof Map) return byId.get(sheetId) ?? null;
    return byId[sheetId] ?? null;
  }

  return null;
}

/**
 * @template T
 * @param {T} value
 * @returns {T}
 */
function cloneJsonish(value) {
  if (structuredCloneFn) return structuredCloneFn(value);
  return JSON.parse(JSON.stringify(value));
}

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {any} value
 * @returns {boolean}
 */
function isNonEmptyPlainObject(value) {
  return isPlainObject(value) && Object.keys(value).length > 0;
}

/**
 * Normalize a DocumentController style reference (style id or style object) into a style object for BranchService.
 *
 * BranchService snapshots should be self-contained, so we always store style objects (not style ids).
 *
 * @param {DocumentController} doc
 * @param {any} raw
 * @returns {Record<string, any> | null}
 */
function branchFormatFromDocFormat(doc, raw) {
  if (raw == null) return null;

  if (typeof raw === "number") {
    const styleId = Number(raw);
    if (!Number.isInteger(styleId) || styleId <= 0) return null;
    const style = doc.styleTable?.get?.(styleId);
    return isNonEmptyPlainObject(style) ? cloneJsonish(style) : null;
  }

  if (isPlainObject(raw)) {
    return isNonEmptyPlainObject(raw) ? cloneJsonish(raw) : null;
  }

  return null;
}

/**
 * Normalize a DocumentController row/col format override table (style ids or style objects)
 * into a sparse `Record<string, styleObject>` suitable for BranchService.
 *
 * @param {DocumentController} doc
 * @param {any} raw
 * @returns {Record<string, Record<string, any>> | null}
 */
function branchFormatMapFromDocFormatMap(doc, raw) {
  if (!raw) return null;

  /** @type {Record<string, Record<string, any>>} */
  const out = {};

  /**
   * @param {any} idxRaw
   * @param {any} formatRaw
   */
  const add = (idxRaw, formatRaw) => {
    const idx = Number(idxRaw);
    if (!Number.isInteger(idx) || idx < 0) return;
    const format = branchFormatFromDocFormat(doc, formatRaw);
    if (!format) return;
    out[String(idx)] = format;
  };

  if (raw instanceof Map) {
    for (const [idx, format] of raw.entries()) add(idx, format);
  } else if (Array.isArray(raw)) {
    for (const entry of raw) {
      if (Array.isArray(entry)) {
        add(entry[0], entry[1]);
        continue;
      }
      if (isPlainObject(entry)) {
        add(entry.index, entry.format ?? entry.style);
      }
    }
  } else if (isPlainObject(raw)) {
    for (const [idx, format] of Object.entries(raw)) add(idx, format);
  }

  return Object.keys(out).length > 0 ? out : null;
}

/**
 * @param {string} key
 * @returns {{ row: number, col: number } | null}
 */
function parseRowColKey(key) {
  if (typeof key !== "string") return null;
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

/**
 * Convert the current DocumentController workbook contents into a BranchService `DocumentState`.
 *
 * This is a full-fidelity adapter for:
 * - literal values (`Cell.value`)
 * - formulas (`Cell.formula`)
 * - formatting (`Cell.format`) stored in DocumentController's style table
 *
 * @param {DocumentController} doc
 * @returns {DocumentState}
 */
export function documentControllerToBranchState(doc) {
  // Preserve the DocumentController's canonical sheet order (Task 201).
  // Do not sort lexicographically, as that loses user-visible tab ordering.
  const sheetIds = doc.getSheetIds().slice();
  /** @type {Record<string, any>} */
  const metaById = {};
  /** @type {Record<string, Record<string, BranchCell>>} */
  const cells = {};
  for (const sheetId of sheetIds) {
    const sheet = doc.model.sheets.get(sheetId);
    /** @type {Record<string, BranchCell>} */
    const outSheet = {};

    if (sheet && sheet.cells && sheet.cells.size > 0) {
      for (const [key, cell] of sheet.cells.entries()) {
        const coord = parseRowColKey(key);
        if (!coord) continue;

        /** @type {BranchCell} */
        const outCell = {};

        if (cell.formula != null) {
          outCell.formula = cell.formula;
        } else if (cell.value !== null && cell.value !== undefined) {
          outCell.value = cloneJsonish(cell.value);
        }

        if (cell.styleId !== 0) {
          outCell.format = cloneJsonish(doc.styleTable.get(cell.styleId));
        }

        if (Object.keys(outCell).length === 0) continue;
        outSheet[formatA1(coord)] = outCell;
      }
    }

    cells[sheetId] = outSheet;
    const sheetMeta = getDocumentControllerSheetMeta(doc, sheetId);
    const name =
      typeof sheetMeta?.name === "string" && sheetMeta.name.length > 0
        ? sheetMeta.name
        : typeof sheetMeta?.displayName === "string" && sheetMeta.displayName.length > 0
          ? sheetMeta.displayName
          : sheetId;

    const rawView = doc.getSheetView(sheetId);
    /** @type {Record<string, any>} */
    const view = cloneJsonish(rawView);

    // --- Layered formats (sheet/row/col defaults) ---
    //
    // DocumentController's internal representation uses style ids, while BranchService stores
    // self-contained style objects. Convert any known layered formats into style objects so
    // branching roundtrips don't drop default formatting.
    //
    // Backwards compatibility: if these fields are absent, treat as no defaults.
    const defaultFormat =
      branchFormatFromDocFormat(
        doc,
        rawView?.defaultFormat ?? rawView?.defaultStyleId ?? rawView?.sheetFormat ?? rawView?.sheetStyleId,
      ) ??
      // New DocumentController layered formatting stores these on the sheet model directly:
      // - `sheet.sheetStyleId` (style id)
      // - `sheet.sheetFormat` (style object, legacy)
      branchFormatFromDocFormat(doc, sheet?.sheetFormat ?? sheet?.sheetStyleId ?? sheet?.defaultFormat ?? sheet?.defaultStyleId);
    const rowFormats =
      branchFormatMapFromDocFormatMap(doc, rawView?.rowFormats ?? rawView?.rowStyleIds) ??
      branchFormatMapFromDocFormatMap(doc, sheet?.rowFormats ?? sheet?.rowStyleIds ?? sheet?.rowStyles);
    const colFormats =
      branchFormatMapFromDocFormatMap(doc, rawView?.colFormats ?? rawView?.colStyleIds) ??
      branchFormatMapFromDocFormatMap(doc, sheet?.colFormats ?? sheet?.colStyleIds ?? sheet?.colStyles);

    // Ensure we never persist style-id based fields.
    delete view.defaultStyleId;
    delete view.rowStyleIds;
    delete view.colStyleIds;
    delete view.sheetStyleId;

    if (defaultFormat) {
      view.defaultFormat = defaultFormat;
    } else {
      delete view.defaultFormat;
    }

    if (rowFormats) {
      view.rowFormats = rowFormats;
    } else {
      delete view.rowFormats;
    }

    if (colFormats) {
      view.colFormats = colFormats;
    } else {
      delete view.colFormats;
    }

    /** @type {Record<string, any>} */
    const metaOut = { id: sheetId, name, view };

    const visibility = sheetMeta?.visibility;
    if (visibility === "visible" || visibility === "hidden" || visibility === "veryHidden") {
      metaOut.visibility = visibility;
    }

    if (sheetMeta && ("tabColor" in sheetMeta || "tab_color" in sheetMeta)) {
      // Collab schema uses an 8-digit ARGB string. Be tolerant of other shapes
      // (e.g. { rgb }) during Task 201 rollout.
      let tabColor = sheetMeta.tabColor ?? sheetMeta.tab_color;
      if (tabColor && typeof tabColor === "object" && typeof tabColor.rgb === "string") {
        tabColor = tabColor.rgb;
      }
      if (tabColor === null) {
        metaOut.tabColor = null;
      } else if (typeof tabColor === "string") {
        metaOut.tabColor = tabColor;
      }
    }

    metaById[sheetId] = metaOut;
  }

  /** @type {DocumentState} */
  const state = {
    schemaVersion: 1,
    sheets: { order: sheetIds, metaById },
    cells,
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  return state;
}

/**
 * Replace the live DocumentController workbook contents from a BranchService `DocumentState`.
 *
 * Missing keys in `state.cells[sheetId]` are treated as deletions (cells will be cleared).
 *
 * @param {DocumentController} doc
 * @param {DocumentState} state
 */
export function applyBranchStateToDocumentController(doc, state) {
  const normalized = normalizeDocumentState(state);
  const sheetIds = normalized.sheets.order.slice();

  const sheets = sheetIds.map((sheetId) => {
    const cellMap = normalized.cells[sheetId] ?? {};
    const meta =
      normalized.sheets.metaById[sheetId] ?? { id: sheetId, name: sheetId, view: { frozenRows: 0, frozenCols: 0 } };
    const view = meta.view ?? { frozenRows: 0, frozenCols: 0 };
    const name = typeof meta.name === "string" && meta.name.length > 0 ? meta.name : sheetId;
    /** @type {Array<{ row: number, col: number, value: any, formula: string | null, format: any }>} */
    const cells = [];

    for (const [addr, cell] of Object.entries(cellMap)) {
      if (!cell || typeof cell !== "object") continue;

      let coord;
      try {
        coord = parseA1(addr);
      } catch {
        continue;
      }

      const hasEnc = cell.enc !== undefined && cell.enc !== null;
      const formula = !hasEnc && typeof cell.formula === "string" ? cell.formula : null;
      const value = hasEnc ? MASKED_CELL_VALUE : (formula !== null ? null : cell.value ?? null);
      const format = cell.format ?? null;

      if (formula === null && value === null && format === null) continue;

      cells.push({
        row: coord.row,
        col: coord.col,
        value,
        formula,
        format: format === null ? null : cloneJsonish(format),
      });
    }

    cells.sort((a, b) => (a.row - b.row === 0 ? a.col - b.col : a.row - b.row));
    /** @type {Record<string, any>} */
    const outSheet = {
      id: sheetId,
      name,
      // Include both the legacy top-level fields (consumed by DocumentController.applyState
      // today) and the nested `view` object (used by BranchService/Yjs schema) so we
      // can round-trip per-sheet UI state and migrate snapshot consumers safely.
      view: cloneJsonish(view),
      frozenRows: view.frozenRows ?? 0,
      frozenCols: view.frozenCols ?? 0,
      colWidths: view.colWidths,
      rowHeights: view.rowHeights,
      cells
    };
    if (meta.visibility === "visible" || meta.visibility === "hidden" || meta.visibility === "veryHidden") {
      outSheet.visibility = meta.visibility;
    }
    if (meta.tabColor === null) {
      outSheet.tabColor = null;
    } else if (typeof meta.tabColor === "string") {
      outSheet.tabColor = meta.tabColor;
    }

    // --- Layered formats (sheet/row/col defaults) ---
    //
    // Backwards compatibility: if absent, treat as no defaults.
    if (isNonEmptyPlainObject(view.defaultFormat)) {
      const format = cloneJsonish(view.defaultFormat);
      // Legacy adapters used `defaultFormat` while the current DocumentController snapshot
      // schema uses `sheetFormat`. Include both so either reader can restore sheet defaults.
      outSheet.defaultFormat = format;
      outSheet.sheetFormat = format;
    }

    if (isNonEmptyPlainObject(view.rowFormats)) {
      outSheet.rowFormats = cloneJsonish(view.rowFormats);
    }

    if (isNonEmptyPlainObject(view.colFormats)) {
      outSheet.colFormats = cloneJsonish(view.colFormats);
    }

    return outSheet;
  });

  const snapshot = { schemaVersion: 1, sheets };

  const encoded =
    typeof TextEncoder !== "undefined"
      ? new TextEncoder().encode(JSON.stringify(snapshot))
      : // eslint-disable-next-line no-undef
        Buffer.from(JSON.stringify(snapshot), "utf8");

  doc.applyState(encoded);
}
