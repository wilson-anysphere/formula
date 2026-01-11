import { formatA1, parseA1 } from "../../document/coords.js";

/**
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").DocumentState} DocumentState
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").Cell} BranchCell
 */

const structuredCloneFn =
  typeof globalThis.structuredClone === "function" ? globalThis.structuredClone : null;

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
  /** @type {DocumentState} */
  const state = { sheets: {} };

  const sheetIds = doc.getSheetIds().slice().sort();
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

    state.sheets[sheetId] = outSheet;
  }

  return state;
}

/**
 * Replace the live DocumentController workbook contents from a BranchService `DocumentState`.
 *
 * Missing keys in `state.sheets[sheetId]` are treated as deletions (cells will be cleared).
 *
 * @param {DocumentController} doc
 * @param {DocumentState} state
 */
export function applyBranchStateToDocumentController(doc, state) {
  const sheetsObj = state?.sheets ?? {};
  const sheetIds = Object.keys(sheetsObj).sort();

  const sheets = sheetIds.map((sheetId) => {
    const cellMap = sheetsObj[sheetId] ?? {};
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

      const formula = typeof cell.formula === "string" ? cell.formula : null;
      const value = formula !== null ? null : cell.value ?? null;
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
    return { id: sheetId, cells };
  });

  const snapshot = { schemaVersion: 1, sheets };

  const encoded =
    typeof TextEncoder !== "undefined"
      ? new TextEncoder().encode(JSON.stringify(snapshot))
      : // eslint-disable-next-line no-undef
        Buffer.from(JSON.stringify(snapshot), "utf8");

  doc.applyState(encoded);

  // DocumentController lazily creates sheets on access. If a branch state includes
  // an explicit empty sheet, ensure it is reflected in the controller model.
  for (const sheet of sheets) {
    if (sheet.cells.length > 0) continue;
    doc.getCell(sheet.id, "A1");
  }
}

