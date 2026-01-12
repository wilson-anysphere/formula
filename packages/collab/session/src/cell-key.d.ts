export type CellAddress = { sheetId: string; row: number; col: number };

export function makeCellKey(cell: CellAddress): string;

export function parseCellKey(key: string, options?: { defaultSheetId?: string }): CellAddress | null;

export function normalizeCellKey(key: string, options?: { defaultSheetId?: string }): string | null;

