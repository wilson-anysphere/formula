import { colToName as colToNameA1, fromA1, toA1 } from "@formula/spreadsheet-frontend/a1";

export type CellAddress = string;

export function colToName(col: number): string {
  if (!Number.isFinite(col) || col < 0) return "";
  return colToNameA1(col);
}

export function nameToCol(name: string): number | null {
  const normalized = name.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return null;
  try {
    const { col0 } = fromA1(`${normalized}1`);
    return col0;
  } catch {
    return null;
  }
}

export function parseCellAddress(addr: CellAddress, out?: { row: number; col: number }): { row: number; col: number } | null {
  const sheetSeparator = addr.lastIndexOf("!");
  const noSheet = sheetSeparator === -1 ? addr : addr.slice(sheetSeparator + 1);
  try {
    const { row0, col0 } = fromA1(noSheet);
    if (out) {
      out.row = row0;
      out.col = col0;
      return out;
    }
    return { row: row0, col: col0 };
  } catch {
    return null;
  }
}

export function formatCellAddress(row: number, col: number): CellAddress {
  return toA1(row, col);
}

export const DEFAULT_MAX_EXPAND_RANGE_CELLS = 200_000;

export function expandRange(range: string, options: { maxCells?: number } = {}): CellAddress[] {
  const sheetSeparator = range.lastIndexOf("!");
  const noSheet = sheetSeparator === -1 ? range : range.slice(sheetSeparator + 1);
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
  const rows = maxRow - minRow + 1;
  const cols = maxCol - minCol + 1;
  const cellCount = rows * cols;
  const maxCells = options.maxCells ?? DEFAULT_MAX_EXPAND_RANGE_CELLS;
  if (cellCount > maxCells) {
    // The auditing/formula-debugger UI cannot meaningfully highlight hundreds of thousands
    // of individual cells; returning a tiny set avoids catastrophic allocations.
    const startAddr = formatCellAddress(minRow, minCol);
    const endAddr = formatCellAddress(maxRow, maxCol);
    return startAddr === endAddr ? [startAddr] : [startAddr, endAddr];
  }
  const out: CellAddress[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      out.push(formatCellAddress(r, c));
    }
  }
  return out;
}
