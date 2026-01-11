import type { CellCoord, Range } from "./types";
import { toA1 } from "@formula/spreadsheet-frontend/a1";

export function cellToA1(cell: CellCoord): string {
  return toA1(cell.row, cell.col);
}

export function rangeToA1(range: Range): string {
  const start = toA1(range.startRow, range.startCol);
  const end = toA1(range.endRow, range.endCol);
  if (start === end) return start;
  return `${start}:${end}`;
}
