import type { Range } from "../../selection/types";

import type { WorkbookSchemaProvider } from "./WorkbookContextBuilder.js";

function isValidRange(value: unknown): value is Range {
  if (!value || typeof value !== "object") return false;
  const r = value as any;
  return (
    Number.isInteger(r.startRow) &&
    Number.isInteger(r.endRow) &&
    Number.isInteger(r.startCol) &&
    Number.isInteger(r.endCol) &&
    r.startRow >= 0 &&
    r.endRow >= r.startRow &&
    r.startCol >= 0 &&
    r.endCol >= r.startCol
  );
}

/**
 * Best-effort adapter: converts the desktop `DocumentWorkbookAdapter` shape
 * (used by search, name box, and formula completion) into a `WorkbookSchemaProvider`
 * that `WorkbookContextBuilder` can consume.
 */
export function createSchemaProviderFromSearchWorkbook(workbook: any): WorkbookSchemaProvider {
  return {
    getNamedRanges: () => {
      const values = typeof workbook?.names?.values === "function" ? Array.from(workbook.names.values()) : [];
      const out: Array<{ name: string; sheetId: string; range: Range }> = [];
      for (const entry of values) {
        const name = typeof entry?.name === "string" ? entry.name.trim() : "";
        const sheetId = typeof entry?.sheetName === "string" ? entry.sheetName : "";
        const range = entry?.range;
        if (!name || !sheetId || !isValidRange(range)) continue;
        out.push({ name, sheetId, range });
      }
      return out;
    },
    getTables: () => {
      const values = typeof workbook?.tables?.values === "function" ? Array.from(workbook.tables.values()) : [];
      const out: Array<{ name: string; sheetId: string; range: Range }> = [];
      for (const table of values) {
        const name = typeof table?.name === "string" ? table.name.trim() : "";
        const sheetId = typeof table?.sheetName === "string" ? table.sheetName : "";
        const startRow = table?.startRow;
        const endRow = table?.endRow;
        const startCol = table?.startCol;
        const endCol = table?.endCol;
        const range: Range = {
          startRow: Number(startRow),
          endRow: Number(endRow),
          startCol: Number(startCol),
          endCol: Number(endCol),
        };
        if (!name || !sheetId || !isValidRange(range)) continue;
        out.push({ name, sheetId, range });
      }
      return out;
    },
  };
}

