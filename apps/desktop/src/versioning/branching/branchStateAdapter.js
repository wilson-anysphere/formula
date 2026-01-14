import { formatA1, parseA1 } from "../../document/coords.js";
import { normalizeDocumentState } from "../../../../../packages/versioning/branches/src/browser.js";

/**
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").DocumentState} DocumentState
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").Cell} BranchCell
 */

const structuredCloneFn =
  typeof globalThis.structuredClone === "function" ? globalThis.structuredClone : null;

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function encodeBase64(bytes) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  if (typeof btoa === "function") {
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }
  throw new Error("Base64 encoding is not supported in this environment");
}

// Collab masking (permissions/encryption) renders unreadable cells as a constant
// placeholder. Branching should treat these as "unknown" rather than persisting
// the placeholder as real content.
const MASKED_CELL_VALUE = "###";

/**
 * Best-effort access to DocumentController sheet metadata.
 *
 * DocumentController tracks authoritative sheet metadata (order + display names).
 *
 * This adapter still feature-detects common access patterns so it can tolerate older
 * controller instances during gradual rollouts and while reading historical snapshots.
 *
 * @param {any} doc
 * @param {string} sheetId
 * @returns {any}
 */
function getDocumentControllerSheetMeta(doc, sheetId) {
  if (!doc || typeof doc !== "object") return null;

  // Preferred API.
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
 * Read a field off an unknown sheet meta value.
 *
 * Supports both plain JS objects and Map-like objects (including Yjs maps)
 * during gradual rollouts / historical snapshot replay.
 *
 * @param {any} meta
 * @param {string} key
 * @returns {any}
 */
function readSheetMetaField(meta, key) {
  if (!meta) return undefined;
  if (typeof meta.get === "function") {
    try {
      return meta.get(key);
    } catch {
      return undefined;
    }
  }
  if (typeof meta === "object") return meta[key];
  return undefined;
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
 * Normalize a DocumentController per-column format-run store into the BranchService `formatRunsByCol`
 * representation (style objects, not style ids).
 *
 * @param {DocumentController} doc
 * @param {any} sheet
 * @returns {Array<{ col: number, runs: Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }> }>}
 */
function branchFormatRunsByColFromDocSheet(doc, sheet) {
  const runsByCol = sheet?.formatRunsByCol ?? sheet?.formatRunsByColumn ?? sheet?.rangeRunsByCol ?? null;
  if (!runsByCol) return [];

  /** @type {Array<{ col: number, runs: Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }> }>} */
  const out = [];

  const iter =
    typeof runsByCol?.entries === "function"
      ? runsByCol.entries()
      : isPlainObject(runsByCol)
        ? Object.entries(runsByCol)
        : [];

  for (const [colKey, runs] of iter) {
    const col = Number(colKey);
    if (!Number.isInteger(col) || col < 0) continue;
    if (!Array.isArray(runs) || runs.length === 0) continue;

    /** @type {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} */
    const outRuns = [];
    for (const run of runs) {
      const startRow = Number(run?.startRow);
      const endRowExclusive = Number(run?.endRowExclusive);
      const styleId = Number(run?.styleId);
      if (!Number.isInteger(startRow) || startRow < 0) continue;
      if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;
      if (!Number.isInteger(styleId) || styleId <= 0) continue;
      const format = branchFormatFromDocFormat(doc, styleId);
      if (!format) continue;
      outRuns.push({ startRow, endRowExclusive, format });
    }
    outRuns.sort((a, b) => a.startRow - b.startRow);
    if (outRuns.length > 0) out.push({ col, runs: outRuns });
  }

  out.sort((a, b) => a.col - b.col);
  return out;
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
  // Preserve the DocumentController's canonical sheet order (tab order).
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

        const cellFormat = branchFormatFromDocFormat(doc, cell.styleId);
        if (cellFormat) outCell.format = cellFormat;

        if (Object.keys(outCell).length === 0) continue;
        outSheet[formatA1(coord)] = outCell;
      }
    }

    cells[sheetId] = outSheet;
    const sheetMeta = getDocumentControllerSheetMeta(doc, sheetId);
    const rawName = readSheetMetaField(sheetMeta, "name");
    const rawDisplayName = readSheetMetaField(sheetMeta, "displayName");
    const name =
      typeof rawName === "string" && rawName.length > 0
        ? rawName
        : typeof rawDisplayName === "string" && rawDisplayName.length > 0
          ? rawDisplayName
          : sheetId;

    const rawView = doc.getSheetView(sheetId);
    /** @type {Record<string, any>} */
    const view = cloneJsonish(rawView);

    // --- Worksheet background images ---
    //
    // Background images are tracked via a sheet-level `backgroundImageId` that references
    // a workbook-scoped image entry (stored in `doc.images` and persisted via BranchService
    // metadata). When the controller supports this field, treat "missing" as an explicit
    // clear (`null`) so branching commits can distinguish omission (older clients) from
    // "clear background".
    const supportsBackgroundImageId =
      typeof doc.getSheetBackgroundImageId === "function" ||
      typeof doc.setSheetBackgroundImageId === "function" ||
      (rawView &&
        typeof rawView === "object" &&
        (Object.prototype.hasOwnProperty.call(rawView, "backgroundImageId") ||
          Object.prototype.hasOwnProperty.call(rawView, "background_image_id")));
    if (supportsBackgroundImageId) {
      let bg = null;
      if (typeof doc.getSheetBackgroundImageId === "function") {
        try {
          bg = doc.getSheetBackgroundImageId(sheetId);
        } catch {
          bg = null;
        }
      } else if (rawView && typeof rawView === "object") {
        bg = rawView.backgroundImageId ?? rawView.background_image_id ?? null;
      }
      const normalized = typeof bg === "string" ? bg.trim() : null;
      view.backgroundImageId = normalized || null;
      delete view.background_image_id;
    }

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

    // --- Range-run formatting (compressed rectangular formatting) ---
    //
    // DocumentController stores large-rectangle formatting as per-column row runs (`sheet.formatRunsByCol`)
    // referencing style ids. Convert these to self-contained style objects so BranchService snapshots can
    // round-trip without depending on the DocumentController style table.
    //
    // Important: we always include the key (even if empty) so BranchService can distinguish
    // "caller omitted the field" from "caller explicitly cleared all runs" during commit overlays.
    view.formatRunsByCol = branchFormatRunsByColFromDocSheet(doc, sheet);

    /** @type {Record<string, any>} */
    const metaOut = { id: sheetId, name, view };

    const visibility = readSheetMetaField(sheetMeta, "visibility");
    if (visibility === "visible" || visibility === "hidden" || visibility === "veryHidden") {
      metaOut.visibility = visibility;
    }

    // Tab color is optional metadata. DocumentController uses an object shape (e.g. `{ rgb }`),
    // while BranchService stores an ARGB string.
    const rawTabColor = readSheetMetaField(sheetMeta, "tabColor");
    const rawTabColorLegacy = readSheetMetaField(sheetMeta, "tab_color");
    const supportsTabColor =
      // DocumentController API
      typeof doc.setSheetTabColor === "function" ||
      // Fallback for alternate metadata shapes
      rawTabColor !== undefined ||
      rawTabColorLegacy !== undefined;

    // Treat "missing" as an explicit clear (null) *when* we know tabColor is supported, so
    // tab-color removals can be committed.
    if (sheetMeta && supportsTabColor) {
      // Collab schema uses an 8-digit ARGB string. Be tolerant of other shapes (e.g. `{ rgb }`)
      // Be tolerant of legacy shapes / historical snapshots.
      let tabColor;
      if (rawTabColor !== undefined) tabColor = rawTabColor;
      else if (rawTabColorLegacy !== undefined) tabColor = rawTabColorLegacy;
      else tabColor = null;

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

  /** @type {Record<string, any>} */
  const metadata = {};

  // --- Workbook images + sheet drawings ---
  //
  // BranchService doesn't have first-class workbook media types yet, but it provides a
  // generic `metadata` map that is versioned + merged. Store the DocumentController
  // image store and per-sheet drawings there so branching doesn't drop embedded media.
  //
  // Shape intentionally mirrors DocumentController's snapshot schema so we can rehydrate
  // via `DocumentController.applyState` without additional transforms.
  const imagesMap = doc?.images;
  if (imagesMap instanceof Map) {
    const images = Array.from(imagesMap.entries())
      .map(([id, entry]) => {
        const imageId = typeof id === "string" ? id : String(id);
        const bytes = entry?.bytes instanceof Uint8Array ? entry.bytes : null;
        if (!bytes) return null;
        /** @type {any} */
        const out = { id: imageId, bytesBase64: encodeBase64(bytes) };
        if ("mimeType" in (entry ?? {})) out.mimeType = entry.mimeType ?? null;
        return out;
      })
      .filter(Boolean)
      .sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
    if (images.length > 0) metadata.images = images;
  }

  const drawingsBySheet = {};
  const canReadDrawings = typeof doc.getSheetDrawings === "function" || doc?.drawingsBySheet instanceof Map;
  if (canReadDrawings) {
    for (const sheetId of sheetIds) {
      let drawings = [];
      if (typeof doc.getSheetDrawings === "function") {
        try {
          drawings = doc.getSheetDrawings(sheetId) ?? [];
        } catch {
          drawings = [];
        }
      } else {
        const raw = doc.drawingsBySheet?.get?.(sheetId);
        drawings = Array.isArray(raw) ? cloneJsonish(raw) : [];
      }
      if (Array.isArray(drawings) && drawings.length > 0) {
        drawingsBySheet[sheetId] = cloneJsonish(drawings);
      }
    }
    if (Object.keys(drawingsBySheet).length > 0) metadata.drawingsBySheet = drawingsBySheet;
  }

  /** @type {DocumentState} */
  const state = {
    schemaVersion: 1,
    sheets: { order: sheetIds, metaById },
    cells,
    metadata,
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

      // Treat any `enc` marker (including `null`) as encrypted so we never fall
      // back to plaintext fields when a marker exists.
      const hasEnc = cell.enc !== undefined;
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
      ...(Object.prototype.hasOwnProperty.call(view, "backgroundImageId")
        ? { backgroundImageId: view.backgroundImageId }
        : Object.prototype.hasOwnProperty.call(view, "background_image_id")
          ? { backgroundImageId: view.background_image_id }
          : {}),
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
      // DocumentController expects a tabColor object (see `TabColor` in documentController).
      // BranchService stores a canonical ARGB string, so wrap it.
      outSheet.tabColor = { rgb: meta.tabColor };
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

    // --- Range-run formatting (compressed rectangular formatting) ---
    // DocumentController expects this at the sheet top-level (not nested under `view`).
    if (Array.isArray(view.formatRunsByCol) && view.formatRunsByCol.length > 0) {
      outSheet.formatRunsByCol = cloneJsonish(view.formatRunsByCol);
    }

    return outSheet;
  });

  // Match DocumentController's snapshot shape: include an explicit `sheetOrder`
  // array so ordering survives even if consumers manipulate/sort the `sheets`
  // list.
  /** @type {any} */
  const snapshot = { schemaVersion: 1, sheetOrder: sheetIds, sheets };

  // Restore workbook images + drawings (stored inside BranchService metadata).
  // These are optional; omit when empty so older controllers/snapshots remain compact.
  const rawImages = normalized.metadata?.images;
  if (Array.isArray(rawImages) && rawImages.length > 0) {
    snapshot.images = cloneJsonish(rawImages);
  }

  const rawDrawingsBySheet = normalized.metadata?.drawingsBySheet;
  if (rawDrawingsBySheet && typeof rawDrawingsBySheet === "object" && !Array.isArray(rawDrawingsBySheet)) {
    /** @type {Record<string, any[]>} */
    const out = {};
    for (const [sheetId, drawings] of Object.entries(rawDrawingsBySheet)) {
      if (!sheetIds.includes(sheetId)) continue;
      if (!Array.isArray(drawings) || drawings.length === 0) continue;
      out[sheetId] = cloneJsonish(drawings);
    }
    if (Object.keys(out).length > 0) snapshot.drawingsBySheet = out;
  }

  const encoded =
    typeof TextEncoder !== "undefined"
      ? new TextEncoder().encode(JSON.stringify(snapshot))
      : // eslint-disable-next-line no-undef
        Buffer.from(JSON.stringify(snapshot), "utf8");

  doc.applyState(encoded);
}
