import { deepEqual, normalizeCell } from "./cell.js";
import { normalizeDocumentState } from "./state.js";

/**
 * @typedef {import("./types.js").Cell} Cell
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").SheetMeta} SheetMeta
 */

/**
 * A patch is a sparse update to a {@link DocumentState}.
 *
 * `null` indicates deletion for map-like collections.
 *
 * Legacy patches (v0) used the shape `{ sheets: Record<sheetId, CellPatch> }`
 * where `sheets` meant *cell* updates. BranchService v2 reserves `sheets` for
 * workbook metadata (sheet order/names, metadata map, named ranges, comments)
 * and uses `cells` for cell updates; {@link applyPatch} still accepts legacy
 * patches for store migration.
 *
 * @typedef {{
 *   schemaVersion?: 1,
 *   sheets?: {
 *     order?: string[],
 *     metaById?: Record<string, SheetMeta | null>
 *   } | Record<string, Record<string, Cell | null>>,
 *   cells?: Record<string, Record<string, Cell | null>>,
 *   metadata?: Record<string, any | null>,
 *   namedRanges?: Record<string, any | null>,
 *   comments?: Record<string, any | null>
 * }} Patch
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
 * @returns {value is Record<string, SheetMeta | null>}
 */
function looksLikeSheetMetaByIdPatch(value) {
  if (!isRecord(value)) return false;
  const entries = Object.entries(value);
  if (entries.length === 0) return false;
  for (const [_sheetId, meta] of entries) {
    if (meta === null) continue;
    if (!isRecord(meta)) return false;
    if (!("id" in meta) && !("name" in meta)) return false;
  }
  return true;
}

/**
 * @param {Patch} patch
 */
function isLegacyCellPatch(patch) {
  if (!isRecord(patch)) return false;
  if (patch.schemaVersion === 1) return false;
  if (!isRecord(patch.sheets)) return false;
  if ("cells" in patch) return false;

  // BranchService v2 uses `patch.sheets.order` / `patch.sheets.metaById` for sheet
  // metadata. Legacy patches store cell updates under `patch.sheets[sheetId]`.
  //
  // Detect legacy patches by checking whether the `sheets` object looks like the
  // newer metadata shape. This intentionally allows legacy cell patches to be
  // combined with newer top-level sections like `metadata` (future-proofing).
  const sheets = /** @type {any} */ (patch.sheets);
  if (Array.isArray(sheets.order)) return false;
  if (looksLikeSheetMetaByIdPatch(sheets.metaById)) return false;
  return true;
}

/**
 * @param {DocumentState} base
 * @param {DocumentState} next
 * @returns {Patch}
 */
export function diffDocumentStates(base, next) {
  const baseState = normalizeDocumentState(base);
  const nextState = normalizeDocumentState(next);

  /** @type {Patch} */
  const patch = { schemaVersion: 1 };

  // --- Sheets metadata ---

  const allSheetIds = new Set([
    ...Object.keys(baseState.sheets.metaById ?? {}),
    ...Object.keys(nextState.sheets.metaById ?? {}),
  ]);

  /** @type {Record<string, SheetMeta | null>} */
  const metaByIdPatch = {};

  for (const sheetId of allSheetIds) {
    const baseMeta = baseState.sheets.metaById?.[sheetId];
    const nextMeta = nextState.sheets.metaById?.[sheetId];

    if (baseMeta && !nextMeta) {
      metaByIdPatch[sheetId] = null;
      continue;
    }
    if (!baseMeta && nextMeta) {
      metaByIdPatch[sheetId] = structuredClone(nextMeta);
      continue;
    }
    if (baseMeta && nextMeta && !deepEqual(baseMeta, nextMeta)) {
      metaByIdPatch[sheetId] = structuredClone(nextMeta);
    }
  }

  if (Object.keys(metaByIdPatch).length > 0) {
    patch.sheets = patch.sheets && !isRecord(patch.sheets) ? {} : (patch.sheets ?? {});
    // @ts-expect-error - Patch.sheets union type.
    patch.sheets.metaById = metaByIdPatch;
  }

  if (!deepEqual(baseState.sheets.order, nextState.sheets.order)) {
    patch.sheets = patch.sheets && !isRecord(patch.sheets) ? {} : (patch.sheets ?? {});
    // @ts-expect-error - Patch.sheets union type.
    patch.sheets.order = structuredClone(nextState.sheets.order);
  }

  // --- Workbook metadata ---

  /** @type {Record<string, any | null>} */
  const metadataPatch = {};
  const metadataKeys = new Set([
    ...Object.keys(baseState.metadata ?? {}),
    ...Object.keys(nextState.metadata ?? {}),
  ]);

  for (const key of metadataKeys) {
    const baseVal = baseState.metadata?.[key];
    const nextVal = nextState.metadata?.[key];
    const baseNorm = baseVal === undefined ? null : baseVal;
    const nextNorm = nextVal === undefined ? null : nextVal;
    if (deepEqual(baseNorm, nextNorm)) continue;
    metadataPatch[key] = nextVal === undefined ? null : structuredClone(nextVal);
  }

  if (Object.keys(metadataPatch).length > 0) patch.metadata = metadataPatch;

  // --- Named ranges ---

  /** @type {Record<string, any | null>} */
  const namedRangesPatch = {};
  const namedRangeKeys = new Set([
    ...Object.keys(baseState.namedRanges ?? {}),
    ...Object.keys(nextState.namedRanges ?? {}),
  ]);

  for (const key of namedRangeKeys) {
    const baseVal = baseState.namedRanges?.[key];
    const nextVal = nextState.namedRanges?.[key];
    const baseNorm = baseVal === undefined ? null : baseVal;
    const nextNorm = nextVal === undefined ? null : nextVal;
    if (deepEqual(baseNorm, nextNorm)) continue;
    namedRangesPatch[key] = nextVal === undefined ? null : structuredClone(nextVal);
  }

  if (Object.keys(namedRangesPatch).length > 0) patch.namedRanges = namedRangesPatch;

  // --- Comments ---

  /** @type {Record<string, any | null>} */
  const commentsPatch = {};
  const commentIds = new Set([
    ...Object.keys(baseState.comments ?? {}),
    ...Object.keys(nextState.comments ?? {}),
  ]);

  for (const id of commentIds) {
    const baseVal = baseState.comments?.[id];
    const nextVal = nextState.comments?.[id];
    const baseNorm = baseVal === undefined ? null : baseVal;
    const nextNorm = nextVal === undefined ? null : nextVal;
    if (deepEqual(baseNorm, nextNorm)) continue;
    commentsPatch[id] = nextVal === undefined ? null : structuredClone(nextVal);
  }

  if (Object.keys(commentsPatch).length > 0) patch.comments = commentsPatch;

  // --- Cells ---

  const removedSheets = new Set(
    Object.keys(baseState.sheets.metaById ?? {}).filter((id) => !(id in (nextState.sheets.metaById ?? {})))
  );

  const sheetIds = new Set([
    ...Object.keys(baseState.cells ?? {}),
    ...Object.keys(nextState.cells ?? {}),
  ]);

  /** @type {Record<string, Record<string, Cell | null>>} */
  const cellsPatch = {};

  for (const sheetId of sheetIds) {
    // If the sheet was removed entirely, rely on the metaById deletion patch
    // instead of writing per-cell deletions (keeps patches sparse).
    if (removedSheets.has(sheetId)) continue;

    const baseSheet = baseState.cells[sheetId] ?? {};
    const nextSheet = nextState.cells[sheetId] ?? {};
    const cellAddrs = new Set([...Object.keys(baseSheet), ...Object.keys(nextSheet)]);

    /** @type {Record<string, Cell | null>} */
    const sheetPatch = {};

    for (const addr of cellAddrs) {
      const baseCell = normalizeCell(baseSheet[addr]);
      const nextCell = normalizeCell(nextSheet[addr]);
      if (deepEqual(baseCell, nextCell)) continue;
      sheetPatch[addr] = nextCell;
    }

    if (Object.keys(sheetPatch).length > 0) cellsPatch[sheetId] = sheetPatch;
  }

  if (Object.keys(cellsPatch).length > 0) patch.cells = cellsPatch;

  return patch;
}

/**
 * @param {DocumentState} state
 * @param {Patch} patch
 * @returns {DocumentState}
 */
export function applyPatch(state, patch) {
  const out = structuredClone(normalizeDocumentState(state));

  /** @type {Record<string, Record<string, Cell | null>> | null} */
  let cellsPatch = null;

  if (isLegacyCellPatch(patch)) {
    // Legacy patch stored cell updates at `patch.sheets`.
    // @ts-expect-error - legacy shape.
    cellsPatch = patch.sheets ?? null;
  } else if (isRecord(patch.cells)) {
    cellsPatch = /** @type {any} */ (patch.cells);
  }

  // --- Apply sheet metadata ---
  if (!isLegacyCellPatch(patch) && isRecord(patch.sheets)) {
    const sheetsPatch = /** @type {any} */ (patch.sheets);
    if (isRecord(sheetsPatch.metaById)) {
      for (const [sheetId, meta] of Object.entries(sheetsPatch.metaById)) {
        if (meta === null) {
          delete out.sheets.metaById[sheetId];
          delete out.cells[sheetId];
          continue;
        }
        if (!isRecord(meta)) continue;
        /** @type {SheetMeta} */
        const nextMeta = {
          id: typeof meta.id === "string" && meta.id.length > 0 ? meta.id : sheetId,
          name: meta.name == null ? null : String(meta.name),
        };
        if (meta.visibility === "visible" || meta.visibility === "hidden" || meta.visibility === "veryHidden") {
          nextMeta.visibility = meta.visibility;
        }
        if ("tabColor" in meta) {
          if (meta.tabColor === null) {
            nextMeta.tabColor = null;
          } else if (typeof meta.tabColor === "string") {
            nextMeta.tabColor = meta.tabColor;
          }
        }
        if (isRecord(meta.view)) {
          // Avoid deep-cloning untrusted view payloads here; we normalize the full document
          // state (including drawing id validation) before returning.
          nextMeta.view = /** @type {any} */ (meta.view);
        } else if (
          "frozenRows" in meta ||
          "frozenCols" in meta ||
          "backgroundImageId" in meta ||
          "background_image_id" in meta ||
          "backgroundImage" in meta ||
          "background_image" in meta
        ) {
          // Back-compat: some older patches may have stored view fields directly on the sheet meta.
          const bg =
            typeof meta.backgroundImageId === "string" || meta.backgroundImageId === null
              ? meta.backgroundImageId
              : typeof meta.background_image_id === "string" || meta.background_image_id === null
                ? meta.background_image_id
                : typeof meta.backgroundImage === "string" || meta.backgroundImage === null
                  ? meta.backgroundImage
                  : typeof meta.background_image === "string" || meta.background_image === null
                    ? meta.background_image
                  : undefined;
          nextMeta.view = structuredClone({
            frozenRows: meta.frozenRows ?? 0,
            frozenCols: meta.frozenCols ?? 0,
            ...(bg !== undefined ? { backgroundImageId: bg } : {}),
          });
        }
        out.sheets.metaById[sheetId] = nextMeta;
        if (!isRecord(out.cells[sheetId])) out.cells[sheetId] = {};
      }
    }

    if (Array.isArray(sheetsPatch.order)) {
      out.sheets.order = sheetsPatch.order.filter((id) => typeof id === "string");
    }
  }

  // --- Apply cells ---
  for (const [sheetId, sheetPatch] of Object.entries(cellsPatch ?? {})) {
    if (!isRecord(sheetPatch)) continue;
    const sheet = out.cells[sheetId] ?? {};
    out.cells[sheetId] = sheet;

    for (const [cell, cellValue] of Object.entries(sheetPatch)) {
      if (cellValue === null) {
        delete sheet[cell];
      } else {
        sheet[cell] = cellValue;
      }
    }
  }

  // --- Apply workbook metadata ---
  if (isRecord(patch.metadata)) {
    for (const [key, value] of Object.entries(patch.metadata)) {
      if (value === null) delete out.metadata[key];
      else out.metadata[key] = value;
    }
  }

  // --- Apply named ranges ---
  if (isRecord(patch.namedRanges)) {
    for (const [key, value] of Object.entries(patch.namedRanges)) {
      if (value === null) delete out.namedRanges[key];
      else out.namedRanges[key] = value;
    }
  }

  // --- Apply comments ---
  if (isRecord(patch.comments)) {
    for (const [id, value] of Object.entries(patch.comments)) {
      if (value === null) delete out.comments[id];
      else out.comments[id] = value;
    }
  }

  return normalizeDocumentState(out);
}

export {};
