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
 * @template T
 * @param {T} value
 * @returns {T}
 */
function cloneJsonish(value) {
  if (structuredCloneFn) return structuredCloneFn(value);
  return JSON.parse(JSON.stringify(value));
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
  const sheetIds = doc.getSheetIds().slice().sort();
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
    // DocumentController doesn't currently track display names separately from ids.
    const view = doc.getSheetView(sheetId);
    metaById[sheetId] = { id: sheetId, name: sheetId, view: cloneJsonish(view) };
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
    const meta = normalized.sheets.metaById[sheetId] ?? { id: sheetId, name: sheetId, view: { frozenRows: 0, frozenCols: 0 } };
    const view = meta.view ?? { frozenRows: 0, frozenCols: 0 };
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
    return {
      id: sheetId,
      frozenRows: view.frozenRows ?? 0,
      frozenCols: view.frozenCols ?? 0,
      colWidths: view.colWidths,
      rowHeights: view.rowHeights,
      cells
    };
  });

  const snapshot = { schemaVersion: 1, sheets };

  const encoded =
    typeof TextEncoder !== "undefined"
      ? new TextEncoder().encode(JSON.stringify(snapshot))
      : // eslint-disable-next-line no-undef
        Buffer.from(JSON.stringify(snapshot), "utf8");

  doc.applyState(encoded);
}
