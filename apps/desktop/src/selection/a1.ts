import type { CellCoord, Range } from "./types";

export function colToName(col: number): string {
  if (col < 0) return "A";
  let n = col + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export function cellToA1(cell: CellCoord): string {
  return `${colToName(cell.col)}${cell.row + 1}`;
}

export function rangeToA1(range: Range): string {
  const start = cellToA1({ row: range.startRow, col: range.startCol });
  const end = cellToA1({ row: range.endRow, col: range.endCol });
  if (start === end) return start;
  return `${start}:${end}`;
}

