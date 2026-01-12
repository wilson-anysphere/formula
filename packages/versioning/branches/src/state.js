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
 * @typedef {import("./types.js").CellMap} CellMap
 */

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
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

/**
 * @param {any} value
 * @returns {{ frozenRows: number, frozenCols: number, colWidths?: Record<string, number>, rowHeights?: Record<string, number> }}
 */
function normalizeSheetView(value) {
  const frozenRows = normalizeFrozenCount(isRecord(value) ? value.frozenRows : undefined);
  const frozenCols = normalizeFrozenCount(isRecord(value) ? value.frozenCols : undefined);

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

  return {
    frozenRows,
    frozenCols,
    ...(colWidths ? { colWidths } : {}),
    ...(rowHeights ? { rowHeights } : {}),
  };
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
    state = structuredClone(input);
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
    state = {
      schemaVersion: 1,
      sheets: isRecord(raw.sheets) ? /** @type {any} */ (structuredClone(raw.sheets)) : { order: [], metaById: {} },
      cells: isRecord(raw.cells)
        ? structuredClone(raw.cells)
        : // Some callers might still pass the legacy `sheets` field for cells.
          (isRecord(raw.sheets) && !Array.isArray(raw.sheets?.order) ? structuredClone(raw.sheets) : {}),
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
      metaById[sheetId] = {
        id: typeof rawMeta.id === "string" && rawMeta.id.length > 0 ? rawMeta.id : sheetId,
        name: rawMeta.name == null ? null : String(rawMeta.name),
        view: normalizeSheetView(rawView),
      };
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
