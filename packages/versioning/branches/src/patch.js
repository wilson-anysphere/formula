import { normalizeCell } from "./cell.js";

/**
 * @typedef {import("./types.js").Cell} Cell
 * @typedef {import("./types.js").DocumentState} DocumentState
 */

/**
 * A patch is a sparse mapping of sheet+cell to the new cell value. `null`
 * indicates deletion.
 *
 * @typedef {{
 *   sheets: Record<string, Record<string, Cell | null>>
 * }} Patch
 */

/**
 * @param {DocumentState} base
 * @param {DocumentState} next
 * @returns {Patch}
 */
export function diffDocumentStates(base, next) {
  /** @type {Patch} */
  const patch = { sheets: {} };

  const sheetIds = new Set([
    ...Object.keys(base.sheets ?? {}),
    ...Object.keys(next.sheets ?? {})
  ]);

  for (const sheetId of sheetIds) {
    const baseSheet = base.sheets[sheetId] ?? {};
    const nextSheet = next.sheets[sheetId] ?? {};
    const cellAddrs = new Set([
      ...Object.keys(baseSheet),
      ...Object.keys(nextSheet)
    ]);

    /** @type {Record<string, Cell | null>} */
    const sheetPatch = {};

    for (const cell of cellAddrs) {
      const baseCell = normalizeCell(baseSheet[cell]);
      const nextCell = normalizeCell(nextSheet[cell]);
      if (JSON.stringify(baseCell) === JSON.stringify(nextCell)) continue;
      sheetPatch[cell] = nextCell;
    }

    if (Object.keys(sheetPatch).length > 0) patch.sheets[sheetId] = sheetPatch;
  }

  return patch;
}

/**
 * @param {DocumentState} state
 * @param {Patch} patch
 * @returns {DocumentState}
 */
export function applyPatch(state, patch) {
  /** @type {DocumentState} */
  const out = structuredClone(state);

  for (const [sheetId, sheetPatch] of Object.entries(patch.sheets ?? {})) {
    const sheet = out.sheets[sheetId] ?? {};
    out.sheets[sheetId] = sheet;

    for (const [cell, cellValue] of Object.entries(sheetPatch)) {
      if (cellValue === null) {
        delete sheet[cell];
      } else {
        sheet[cell] = cellValue;
      }
    }
  }

  return out;
}

export {};

