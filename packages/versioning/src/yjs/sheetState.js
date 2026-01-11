import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";
import { parseCellKey } from "../../../collab/session/src/cell-key.js";

/**
 * Parse a spreadsheet cell key. Supports:
 * - `${sheetId}:${row}:${col}` (docs/06-collaboration.md)
 * - `${sheetId}:${row},${col}` (legacy internal encoding)
 * - `r{row}c{col}` (unit-test convenience, resolved against `defaultSheetId`)
 *
 * @param {string} key
 * @param {{ defaultSheetId?: string }} [opts]
 * @returns {{ sheetId: string, row: number, col: number } | null}
 */
export function parseSpreadsheetCellKey(key, opts = {}) {
  const parsed = parseCellKey(key, { defaultSheetId: opts.defaultSheetId ?? "Sheet1" });
  if (!parsed) return null;
  return { sheetId: parsed.sheetId, row: parsed.row, col: parsed.col };
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
