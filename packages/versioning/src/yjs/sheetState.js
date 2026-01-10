import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";

/**
 * Parse a spreadsheet cell key. Supports:
 * - `${sheetId}:${row}:${col}` (docs/06-collaboration.md)
 * - `r{row}c{col}` (unit-test convenience)
 *
 * @param {string} key
 * @returns {{ sheetId: string | null, row: number, col: number } | null}
 */
export function parseSpreadsheetCellKey(key) {
  const colon = key.split(":");
  if (colon.length === 3) {
    const row = Number(colon[1]);
    const col = Number(colon[2]);
    if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
    return { sheetId: colon[0], row, col };
  }

  const m = key.match(/^r(\d+)c(\d+)$/);
  if (m) {
    return { sheetId: null, row: Number(m[1]), col: Number(m[2]) };
  }

  return null;
}

/**
 * @param {any} cellData
 */
function extractCell(cellData) {
  if (cellData instanceof Y.Map) {
    return {
      value: cellData.get("value") ?? null,
      formula: cellData.get("formula") ?? null,
      format: cellData.get("format") ?? cellData.get("style") ?? null,
    };
  }
  if (cellData && typeof cellData === "object") {
    return {
      value: cellData.value ?? null,
      formula: cellData.formula ?? null,
      format: cellData.format ?? cellData.style ?? null,
    };
  }
  return { value: cellData ?? null, formula: null, format: null };
}

/**
 * Convert a Yjs doc into a per-sheet state suitable for semantic diff.
 *
 * @param {Y.Doc} doc
 * @param {{ sheetId?: string | null }} [opts]
 */
export function sheetStateFromYjsDoc(doc, opts = {}) {
  const targetSheetId = opts.sheetId ?? null;
  const cellsMap = doc.getMap("cells");

  /** @type {Map<string, any>} */
  const cells = new Map();
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed) return;
    if (targetSheetId != null && parsed.sheetId !== targetSheetId) return;
    cells.set(cellKey(parsed.row, parsed.col), extractCell(cellData));
  });

  return { cells };
}

/**
 * @param {Uint8Array} snapshot
 * @param {{ sheetId?: string | null }} [opts]
 */
export function sheetStateFromYjsSnapshot(snapshot, opts = {}) {
  const doc = new Y.Doc();
  Y.applyUpdate(doc, snapshot);
  return sheetStateFromYjsDoc(doc, opts);
}

