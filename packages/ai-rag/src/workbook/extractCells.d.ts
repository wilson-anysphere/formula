import type { Rect } from "./rect.js";

/**
 * Extract a bounded window of normalized cells from a workbook sheet.
 */
export function extractCells(
  sheet: any,
  rect: Rect,
  opts?: { maxRows?: number; maxCols?: number; signal?: AbortSignal },
): any[][];

