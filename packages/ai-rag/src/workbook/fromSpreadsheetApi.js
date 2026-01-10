/**
 * Adapter from a spreadsheet API (e.g. packages/ai-tools SpreadsheetApi) to the
 * `packages/ai-rag` Workbook shape.
 *
 * This lets RAG indexing work without materializing full 2D matrices; callers
 * only need to provide `listNonEmptyCells`.
 */

/**
 * @param {{
 *   spreadsheet: {
 *     listSheets(): string[],
 *     listNonEmptyCells(sheet?: string): Array<{ address: { sheet: string, row: number, col: number }, cell: { value?: any, formula?: string } }>
 *   },
 *   workbookId: string
 * }} params
 */
export function workbookFromSpreadsheetApi(params) {
  const { spreadsheet, workbookId } = params;
  const sheetNames = spreadsheet.listSheets();

  const sheets = sheetNames.map((sheetName) => {
    const cells = new Map();
    const entries = spreadsheet.listNonEmptyCells(sheetName);
    for (const entry of entries) {
      const row = entry?.address?.row;
      const col = entry?.address?.col;
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      const cell = entry?.cell ?? {};
      cells.set(`${row},${col}`, { value: cell.value ?? null, formula: cell.formula ?? null });
    }
    return { name: sheetName, cells };
  });

  return { id: workbookId, sheets };
}

