/**
 * Normalize a raw cell into the `{ v, f }` shape used by ai-rag.
 */
export function normalizeCell(raw: any): { v?: any; f?: string };

export function getSheetMatrix(sheet: any): any[][] | null;

export function getSheetCellMap(sheet: any): Map<string, any> | null;

export function getCellRaw(sheet: any, row: number, col: number): any;

