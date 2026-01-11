export type CellCoord = { row: number; col: number };
export type CellRange = { start: CellCoord; end: CellCoord };

export function columnIndexToName(colIndex: number): string;
export function columnNameToIndex(name: string): number;
export function parseA1(a1: string): CellCoord;
export function formatA1(coord: CellCoord): string;
export function normalizeRange(range: CellRange): CellRange;
export function parseRangeA1(a1Range: string): CellRange;

