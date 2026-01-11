export function columnLettersToIndex(letters: string): number | null;

export function parseCellRef(cell: string): { col: number; row: number } | null;

export function parseA1Range(rangeRef: string): {
  sheetName?: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
} | null;

