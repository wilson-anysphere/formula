import { colToName as colToNameA1, fromA1, toA1 } from "@formula/spreadsheet-frontend/a1";

export type CellAddress = string;

export function colToName(col: number): string {
  if (!Number.isFinite(col) || col < 0) return "";
  return colToNameA1(col);
}

export function nameToCol(name: string): number | null {
  try {
    const { col0 } = fromA1(`${name}1`);
    return col0;
  } catch {
    return null;
  }
}

export function parseCellAddress(addr: CellAddress): { row: number; col: number } | null {
  const noSheet = addr.includes("!") ? addr.split("!").slice(-1)[0] : addr;
  try {
    const { row0, col0 } = fromA1(noSheet);
    return { row: row0, col: col0 };
  } catch {
    return null;
  }
}

export function formatCellAddress(row: number, col: number): CellAddress {
  return toA1(row, col);
}

export function expandRange(range: string): CellAddress[] {
  const noSheet = range.includes("!") ? range.split("!").slice(-1)[0] : range;
  const parts = noSheet.split(":");
  if (parts.length === 1) return [noSheet];
  if (parts.length !== 2) return [];
  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  if (!start || !end) return [];
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);
  const minCol = Math.min(start.col, end.col);
  const maxCol = Math.max(start.col, end.col);
  const out: CellAddress[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      out.push(formatCellAddress(r, c));
    }
  }
  return out;
}
