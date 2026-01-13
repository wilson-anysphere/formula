/**
 * Adapter from a spreadsheet API (e.g. packages/ai-tools SpreadsheetApi) to the
 * `packages/ai-rag` Workbook shape.
 *
 * This lets RAG indexing work without materializing full 2D matrices; callers
 * only need to provide `listNonEmptyCells`.
 */

import { throwIfAborted } from "../utils/abort.js";

/**
 * Coordinate base for `address.row/col` returned by the SpreadsheetApi.
 *
 * - `"one"`: 1-based coordinates (A1 => row=1,col=1). This matches `@formula/ai-tools`.
 * - `"zero"`: 0-based coordinates (A1 => row=0,col=0).
 * - `"auto"`: Heuristic detection. If *any* non-empty cell has row===0 or col===0,
 *   assume 0-based, otherwise assume 1-based. This is deterministic but cannot
 *   disambiguate cases where all cells are >0 (e.g. data starts at row 10).
 *
 * Note: `packages/ai-rag` always uses 0-based coordinates internally.
 *
 * @param {{
 *   spreadsheet: {
 *     listSheets(): string[],
 *     listNonEmptyCells(sheet?: string): Array<{ address: { sheet: string, row: number, col: number }, cell: { value?: any, formula?: string } }>
 *   },
 *   workbookId: string,
 *   coordinateBase?: "one" | "zero" | "auto"
 *   signal?: AbortSignal
 * }} params
 */
export function workbookFromSpreadsheetApi(params) {
  const { spreadsheet, workbookId } = params;
  const signal = params.signal;
  const coordinateBase = params.coordinateBase ?? "one";
  if (coordinateBase !== "one" && coordinateBase !== "zero" && coordinateBase !== "auto") {
    throw new Error(`workbookFromSpreadsheetApi: invalid coordinateBase "${coordinateBase}"`);
  }
  throwIfAborted(signal);
  const sheetNames = spreadsheet.listSheets();

  /** @type {Map<string, any[]>} */
  const entriesBySheet = new Map();
  let sawZeroBasedCoord = false;

  for (const sheetName of sheetNames) {
    throwIfAborted(signal);
    const entries = spreadsheet.listNonEmptyCells(sheetName) ?? [];
    entriesBySheet.set(sheetName, entries);
    for (const entry of entries) {
      throwIfAborted(signal);
      const row = entry?.address?.row;
      const col = entry?.address?.col;
      if (row === 0 || col === 0) sawZeroBasedCoord = true;
    }
  }

  throwIfAborted(signal);
  const resolvedBase =
    coordinateBase === "auto" ? (sawZeroBasedCoord ? "zero" : "one") : coordinateBase;

  const sheets = sheetNames.map((sheetName) => {
    throwIfAborted(signal);
    const cells = new Map();
    const entries = entriesBySheet.get(sheetName) ?? [];
    for (const entry of entries) {
      throwIfAborted(signal);
      const inputRow = entry?.address?.row;
      const inputCol = entry?.address?.col;
      if (!Number.isInteger(inputRow) || !Number.isInteger(inputCol)) continue;

      const row = resolvedBase === "one" ? inputRow - 1 : inputRow;
      const col = resolvedBase === "one" ? inputCol - 1 : inputCol;

      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      const cell = entry?.cell ?? {};
      const value = cell.value ?? null;
      const formula = cell.formula ?? null;
      // `SpreadsheetApi.listNonEmptyCells` may include formatting-only cells. These should
      // be dropped from the ai-rag workbook to avoid bloating sparse cell maps.
      if ((value == null || value === "") && (formula == null || String(formula).trim() === "")) continue;
      cells.set(`${row},${col}`, { value, formula });
    }
    return { name: sheetName, cells };
  });

  return { id: workbookId, sheets };
}
