/**
 * Workbook state helpers for BranchService.
 *
 * Branching originally stored only `{ sheets: Record<sheetId, CellMap> }`. As
 * branching expanded to include full workbook metadata (sheet order/names,
 * metadata map, named ranges, comments), persisted stores may still contain
 * commits using the legacy schema. These helpers provide a single
 * normalization/migration path so old histories can still be loaded.
 */

/**
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").LegacyDocumentState} LegacyDocumentState
 * @typedef {import("./types.js").SheetsState} SheetsState
 * @typedef {import("./types.js").SheetMeta} SheetMeta
 * @typedef {import("./types.js").SheetVisibility} SheetVisibility
 * @typedef {import("./types.js").CellMap} CellMap
 */

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function getYText(value) {
  if (!value || typeof value !== "object") return null;
  const maybe = value;
  // Duck-type against multiple `yjs` module instances (ESM/CJS) and constructor renaming.
  // Keep this aligned with `@formula/collab-yjs-utils` so we don't accidentally treat
  // plain JS objects (or Maps) as Y.Text.
  if (typeof maybe.toString !== "function") return null;
  if (typeof maybe.toDelta !== "function") return null;
  if (typeof maybe.applyDelta !== "function") return null;
  if (typeof maybe.insert !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return maybe;
}

/**
 * @param {any} value
 * @returns {number}
 */
function normalizeFrozenCount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

// Drawing ids can be authored via remote/shared state (sheet view state). Keep validation strict
// so BranchService normalization/merges don't deep-clone pathological ids (e.g. multi-megabyte strings).
const MAX_DRAWING_ID_STRING_CHARS = 4096;

/**
 * @param {any} value
 * @returns {{
 *   frozenRows: number,
 *   frozenCols: number,
 *   backgroundImageId?: string | null,
 *   colWidths?: Record<string, number>,
 *   rowHeights?: Record<string, number>,
 *   mergedRanges?: Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>,
 *   drawings?: unknown[],
 *   defaultFormat?: Record<string, any>,
 *   rowFormats?: Record<string, Record<string, any>>,
 *   colFormats?: Record<string, Record<string, any>>,
 *   formatRunsByCol?: Array<{ col: number, runs: Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }> }>,
 * }}
 */
function normalizeSheetView(value) {
  const frozenRows = normalizeFrozenCount(isRecord(value) ? value.frozenRows : undefined);
  const frozenCols = normalizeFrozenCount(isRecord(value) ? value.frozenCols : undefined);

  const hasBackgroundImageId =
    isRecord(value) &&
    (Object.prototype.hasOwnProperty.call(value, "backgroundImageId") ||
      Object.prototype.hasOwnProperty.call(value, "background_image_id"));
  const backgroundImageIdRaw = isRecord(value) ? value.backgroundImageId ?? value.background_image_id : undefined;
  /** @type {string | null} */
  let backgroundImageId = null;
  if (typeof backgroundImageIdRaw === "string") {
    const trimmed = backgroundImageIdRaw.trim();
    if (trimmed) backgroundImageId = trimmed;
  } else if (backgroundImageIdRaw === null) {
    // Preserve explicit clears so semantic merges can distinguish "omitted" from "cleared".
    backgroundImageId = null;
  }

  const hasMergedRanges =
    isRecord(value) &&
    (Object.prototype.hasOwnProperty.call(value, "mergedRanges") ||
      Object.prototype.hasOwnProperty.call(value, "mergedCells") ||
      Object.prototype.hasOwnProperty.call(value, "merged_cells"));
  const mergedRangesRaw = isRecord(value)
    ? value.mergedRanges ?? value.mergedCells ?? value.merged_cells
    : undefined;

  /**
   * @param {any} raw
   * @returns {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>}
   */
  const normalizeMergedRanges = (raw) => {
    if (raw == null) return [];
    let arr = raw;
    const isYArrayLike = Boolean(arr && typeof arr === "object" && typeof arr.length === "number" && typeof arr.get === "function");
    if (!Array.isArray(arr) && !isYArrayLike && typeof arr?.toArray === "function") {
      try {
        arr = arr.toArray();
      } catch {
        // ignore
      }
    }
    const len = Array.isArray(arr) ? arr.length : isYArrayLike ? arr.length : 0;
    if (len === 0) return [];

    const overlaps = (a, b) =>
      a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;

    /** @type {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>} */
    const out = [];

    for (let idx = 0; idx < len; idx += 1) {
      const entry = Array.isArray(arr) ? arr[idx] : arr.get(idx);
      const sr = Number(entry?.startRow ?? entry?.start_row ?? entry?.sr);
      const er = Number(entry?.endRow ?? entry?.end_row ?? entry?.er);
      const sc = Number(entry?.startCol ?? entry?.start_col ?? entry?.sc);
      const ec = Number(entry?.endCol ?? entry?.end_col ?? entry?.ec);
      if (!Number.isInteger(sr) || sr < 0) continue;
      if (!Number.isInteger(er) || er < 0) continue;
      if (!Number.isInteger(sc) || sc < 0) continue;
      if (!Number.isInteger(ec) || ec < 0) continue;

      const startRow = Math.min(sr, er);
      const endRow = Math.max(sr, er);
      const startCol = Math.min(sc, ec);
      const endCol = Math.max(sc, ec);
      // Skip degenerate single-cell merges.
      if (startRow === endRow && startCol === endCol) continue;

      const candidate = { startRow, endRow, startCol, endCol };
      for (let i = out.length - 1; i >= 0; i -= 1) {
        if (overlaps(out[i], candidate)) out.splice(i, 1);
      }
      out.push(candidate);
    }

    if (out.length === 0) return [];

    out.sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol || a.endRow - b.endRow || a.endCol - b.endCol);

    /** @type {Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>} */
    const deduped = [];
    let lastKey = null;
    for (const r of out) {
      const key = `${r.startRow},${r.endRow},${r.startCol},${r.endCol}`;
      if (key === lastKey) continue;
      lastKey = key;
      deduped.push(r);
    }

    return deduped;
  };

  const mergedRanges = hasMergedRanges ? normalizeMergedRanges(mergedRangesRaw) : null;

  const hasDrawings = isRecord(value) && Object.prototype.hasOwnProperty.call(value, "drawings");
  const drawingsRaw = isRecord(value) ? value.drawings : undefined;

  /**
   * @param {any} raw
   * @returns {unknown[]}
   */
  const normalizeDrawings = (raw) => {
    if (raw == null) return [];
    let arr = raw;
    const isYArrayLike = Boolean(arr && typeof arr === "object" && typeof arr.length === "number" && typeof arr.get === "function");
    // Avoid Yjs dev-mode warnings for unintegrated arrays by reading from `_prelimContent`.
    if (isYArrayLike && arr?.doc == null && Array.isArray(arr?._prelimContent)) {
      arr = arr._prelimContent;
    }
    if (!Array.isArray(arr) && !isYArrayLike && typeof arr?.toArray === "function") {
      try {
        arr = arr.toArray();
      } catch {
        // ignore
      }
    }
    const len = Array.isArray(arr) ? arr.length : isYArrayLike ? arr.length : 0;
    if (len === 0) return [];

    /** @type {any[]} */
    const out = [];

    for (let idx = 0; idx < len; idx += 1) {
      const entry = Array.isArray(arr) ? arr[idx] : arr.get(idx);

      // Validate `id` *before* materializing via `.toJSON()`. Yjs' `toJSON()` will stringify
      // nested Y.Text values (including drawing ids), which can allocate large strings (or
      // throw) for pathological ids.
      let rawId = entry?.id;
      if (entry && typeof entry === "object" && typeof entry.get === "function") {
        const prelim = entry?.doc == null ? entry?._prelimContent : null;
        if (prelim instanceof Map) {
          rawId = prelim.get("id");
        } else {
          try {
            rawId = entry.get("id");
          } catch {
            rawId = undefined;
          }
        }
      }
      let normalizedId;
      if (typeof rawId === "string") {
        if (rawId.length > MAX_DRAWING_ID_STRING_CHARS) continue;
        const trimmed = rawId.trim();
        if (!trimmed) continue;
        normalizedId = trimmed;
      } else if (typeof rawId === "number") {
        if (!Number.isSafeInteger(rawId)) continue;
        normalizedId = rawId;
      } else {
        const text = getYText(rawId);
        if (!text) continue;
        // Unintegrated Y.Text instances warn on `.length`/`.toString()` reads and may have a
        // stale `_length` (pending ops). Treat them as invalid ids.
        if (text.doc == null) continue;
        if (typeof text.length === "number" && text.length > MAX_DRAWING_ID_STRING_CHARS) continue;
        let str;
        try {
          str = text.toString();
        } catch {
          continue;
        }
        if (typeof str !== "string") continue;
        if (str.length > MAX_DRAWING_ID_STRING_CHARS) continue;
        const trimmed = str.trim();
        if (!trimmed) continue;
        normalizedId = trimmed;
      }

      let json = entry;
      if (json && typeof json === "object" && typeof json.toJSON === "function") {
        try {
          json = json.toJSON();
        } catch {
          // ignore
        }
      }
      if (!isRecord(json)) continue;

      try {
        const cloned = structuredClone(json);
        cloned.id = normalizedId;
        out.push(cloned);
      } catch {
        // ignore malformed/non-cloneable drawing entries
      }
    }

    out.sort((a, b) => {
      const za = Number.isFinite(Number(a?.zOrder)) ? Number(a.zOrder) : 0;
      const zb = Number.isFinite(Number(b?.zOrder)) ? Number(b.zOrder) : 0;
      if (za !== zb) return za - zb;
      const ida = a?.id == null ? "" : String(a.id);
      const idb = b?.id == null ? "" : String(b.id);
      return ida < idb ? -1 : ida > idb ? 1 : 0;
    });

    return out;
  };

  const drawings = hasDrawings ? normalizeDrawings(drawingsRaw) : null;

  const normalizeFormatObject = (raw) => {
    if (!isRecord(raw)) return null;
    // Treat empty objects as "no format" so we don't bloat state with meaningless defaults.
    if (Object.keys(raw).length === 0) return null;
    return structuredClone(raw);
  };

  const normalizeFormatRunsByCol = (raw) => {
    if (!raw) return null;
    /** @type {Array<{ col: number, runs: Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }> }>} */
    const out = [];

    const normalizeRunList = (rawRuns) => {
      if (!Array.isArray(rawRuns) || rawRuns.length === 0) return [];
      /** @type {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} */
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
        const format = normalizeFormatObject(entry?.format ?? entry?.style);
        if (!format) continue;
        runs.push({ startRow, endRowExclusive, format });
      }
      runs.sort((a, b) => a.startRow - b.startRow);
      return runs;
    };

    const addColRuns = (colRaw, runsRaw) => {
      const col = Number(colRaw);
      if (!Number.isInteger(col) || col < 0) return;
      const runs = normalizeRunList(runsRaw);
      if (runs.length === 0) return;
      out.push({ col, runs });
    };

    if (raw instanceof Map) {
      for (const [col, runs] of raw.entries()) addColRuns(col, runs);
    } else if (Array.isArray(raw)) {
      for (const entry of raw) {
        if (!entry) continue;
        if (Array.isArray(entry)) {
          addColRuns(entry[0], entry[1]);
          continue;
        }
        addColRuns(entry?.col ?? entry?.index ?? entry?.column, entry?.runs ?? entry?.formatRuns ?? entry?.segments);
      }
    } else if (isRecord(raw)) {
      for (const [key, value] of Object.entries(raw)) {
        if (Array.isArray(value)) {
          addColRuns(key, value);
          continue;
        }
        if (isRecord(value)) {
          addColRuns(key, value.runs ?? value.formatRuns ?? value.segments);
        }
      }
    }

    out.sort((a, b) => a.col - b.col);
    return out.length === 0 ? null : out;
  };

  const normalizeFormatOverrides = (raw) => {
    if (!raw) return null;
    /** @type {Record<string, any>} */
    const out = {};

    if (Array.isArray(raw)) {
      for (const entry of raw) {
        const index = Array.isArray(entry) ? entry[0] : entry?.index;
        const format = Array.isArray(entry) ? entry[1] : (entry?.format ?? entry?.style);
        const idx = Number(index);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeFormatObject(format);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    } else if (isRecord(raw)) {
      for (const [key, format] of Object.entries(raw)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeFormatObject(format);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    }

    return Object.keys(out).length === 0 ? null : out;
  };

  const normalizeAxisSize = (raw) => {
    const num = Number(raw);
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
    } else if (isRecord(raw)) {
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

  const colWidths = isRecord(value) ? normalizeAxisOverrides(value.colWidths) : null;
  const rowHeights = isRecord(value) ? normalizeAxisOverrides(value.rowHeights) : null;

  const defaultFormat = isRecord(value) ? normalizeFormatObject(value.defaultFormat) : null;
  const rowFormats = isRecord(value) ? normalizeFormatOverrides(value.rowFormats) : null;
  const colFormats = isRecord(value) ? normalizeFormatOverrides(value.colFormats) : null;
  let formatRunsByCol = null;
  if (isRecord(value)) {
    const hasKey = Object.prototype.hasOwnProperty.call(value, "formatRunsByCol");
    formatRunsByCol = normalizeFormatRunsByCol(value.formatRunsByCol);
    // Preserve explicit clears/empties so callers can distinguish omission ("don't change")
    // from "set to empty" when performing back-compat overlays in BranchService.commit().
    if (formatRunsByCol == null && hasKey) formatRunsByCol = [];
  }

  return {
    frozenRows,
    frozenCols,
    ...(hasBackgroundImageId ? { backgroundImageId } : {}),
    ...(colWidths ? { colWidths } : {}),
    ...(rowHeights ? { rowHeights } : {}),
    ...(hasMergedRanges ? { mergedRanges } : {}),
    ...(hasDrawings ? { drawings } : {}),
    ...(defaultFormat ? { defaultFormat } : {}),
    ...(rowFormats ? { rowFormats } : {}),
    ...(colFormats ? { colFormats } : {}),
    ...(formatRunsByCol != null ? { formatRunsByCol } : {}),
  };
}

/**
 * @param {any} value
 * @returns {SheetVisibility | undefined}
 */
function normalizeSheetVisibility(value) {
  if (value === "visible" || value === "hidden" || value === "veryHidden") return value;
  return undefined;
}

/**
 * @param {any} value
 * @returns {string | null | undefined}
 */
function normalizeTabColor(value) {
  if (value == null) return value === null ? null : undefined;
  const str = String(value);
  if (!/^[0-9A-Fa-f]{8}$/.test(str)) return undefined;
  return str.toUpperCase();
}

/**
 * @returns {DocumentState}
 */
export function emptyDocumentState() {
  return {
    schemaVersion: 1,
    sheets: { order: [], metaById: {} },
    cells: {},
    metadata: {},
    namedRanges: {},
    comments: {},
  };
}

/**
 * @param {any} input
 * @returns {input is DocumentState}
 */
function isWorkbookDocumentState(input) {
  return (
    isRecord(input) &&
    input.schemaVersion === 1 &&
    isRecord(input.sheets) &&
    Array.isArray(input.sheets.order) &&
    isRecord(input.sheets.metaById) &&
    isRecord(input.cells)
  );
}

/**
 * @param {any} input
 * @returns {input is LegacyDocumentState}
 */
function isLegacyDocumentState(input) {
  return isRecord(input) && isRecord(input.sheets) && !("cells" in input) && !("schemaVersion" in input);
}

/**
 * Normalize (and if needed, migrate) an arbitrary value into a valid BranchService
 * {@link DocumentState}.
 *
 * @param {any} input
 * @returns {DocumentState}
 */
export function normalizeDocumentState(input) {
  /** @type {DocumentState} */
  let state;

  if (isWorkbookDocumentState(input)) {
    // Avoid `structuredClone(input)` here: `sheets.metaById[*].view.drawings` is user-authored shared
    // state and may contain pathological ids (e.g. multi-megabyte strings). We rebuild normalized
    // sheet metadata below, so deep-cloning the raw sheet meta up-front is unnecessary and can
    // amplify memory usage.
    const raw = /** @type {any} */ (input);
    state = {
      schemaVersion: 1,
      // Keep a reference to the raw sheets object for read-only normalization below.
      // (We overwrite `state.sheets` with a new normalized object before returning.)
      sheets: raw.sheets,
      // Cells/metadata collections are still cloned so the returned state is safe to mutate.
      cells: structuredClone(raw.cells),
      metadata: isRecord(raw.metadata) ? structuredClone(raw.metadata) : {},
      namedRanges: isRecord(raw.namedRanges) ? structuredClone(raw.namedRanges) : {},
      comments: isRecord(raw.comments) ? structuredClone(raw.comments) : {},
    };
  } else if (isLegacyDocumentState(input)) {
    const legacy = /** @type {LegacyDocumentState} */ (input);
    const cells = isRecord(legacy.sheets) ? structuredClone(legacy.sheets) : {};
    state = {
      schemaVersion: 1,
      sheets: { order: Object.keys(cells), metaById: {} },
      cells,
      metadata: {},
      namedRanges: {},
      comments: {},
    };
  } else {
    // Be forgiving: best-effort coerce partial objects (useful for tests).
    const raw = isRecord(input) ? input : {};
    const sheetsRaw = raw.sheets;
    state = {
      schemaVersion: 1,
      // Avoid deep-cloning sheets: we rebuild normalized sheet metadata below.
      sheets: isRecord(sheetsRaw) ? /** @type {any} */ (sheetsRaw) : { order: [], metaById: {} },
      cells: isRecord(raw.cells)
        ? structuredClone(raw.cells)
        : // Some callers might still pass the legacy `sheets` field for cells.
          (isRecord(sheetsRaw) && !Array.isArray(sheetsRaw?.order) ? structuredClone(sheetsRaw) : {}),
      metadata: isRecord(raw.metadata) ? structuredClone(raw.metadata) : {},
      namedRanges: isRecord(raw.namedRanges) ? structuredClone(raw.namedRanges) : {},
      comments: isRecord(raw.comments) ? structuredClone(raw.comments) : {},
    };
  }

  // --- Normalize workbook-level collections ---

  if (!isRecord(state.cells)) state.cells = {};
  if (!isRecord(state.metadata)) state.metadata = {};
  if (!isRecord(state.namedRanges)) state.namedRanges = {};
  if (!isRecord(state.comments)) state.comments = {};

  /** @type {SheetsState} */
  const sheetsState = isRecord(state.sheets) ? state.sheets : { order: [], metaById: {} };
  const rawOrder = Array.isArray(sheetsState.order) ? sheetsState.order : [];
  const rawMetaById = isRecord(sheetsState.metaById) ? sheetsState.metaById : {};

  // Collect sheet ids from both cells and metadata.
  const sheetIds = new Set([
    ...Object.keys(state.cells ?? {}),
    ...Object.keys(rawMetaById ?? {}),
  ]);

  /** @type {Record<string, SheetMeta>} */
  const metaById = {};
  for (const sheetId of sheetIds) {
    const rawMeta = rawMetaById[sheetId];
    if (isRecord(rawMeta)) {
      const rawView = isRecord(rawMeta.view) ? rawMeta.view : rawMeta;
      /** @type {SheetMeta} */
      const meta = {
        id: typeof rawMeta.id === "string" && rawMeta.id.length > 0 ? rawMeta.id : sheetId,
        name: rawMeta.name == null ? null : String(rawMeta.name),
        view: normalizeSheetView(rawView),
      };

      const visibility = normalizeSheetVisibility(rawMeta.visibility);
      if (visibility !== undefined) meta.visibility = visibility;

      const tabColor = normalizeTabColor(rawMeta.tabColor);
      if (tabColor !== undefined) meta.tabColor = tabColor;

      metaById[sheetId] = meta;
    } else {
      // Legacy histories have no separate sheet name; treat id as the display name.
      metaById[sheetId] = { id: sheetId, name: sheetId, view: { frozenRows: 0, frozenCols: 0 } };
    }
  }

  // Ensure `cells` has an entry for every sheet (even empty sheets).
  for (const sheetId of Object.keys(metaById)) {
    if (!isRecord(state.cells[sheetId])) state.cells[sheetId] = /** @type {CellMap} */ ({});
  }

  /** @type {string[]} */
  const order = [];
  const seen = new Set();
  for (const id of rawOrder) {
    if (typeof id !== "string" || id.length === 0) continue;
    if (seen.has(id)) continue;
    if (!metaById[id]) continue;
    order.push(id);
    seen.add(id);
  }

  // Append any sheets not present in the order (stable iteration order).
  for (const id of Object.keys(metaById)) {
    if (seen.has(id)) continue;
    order.push(id);
    seen.add(id);
  }

  state.schemaVersion = 1;
  state.sheets = { order, metaById };

  return state;
}
