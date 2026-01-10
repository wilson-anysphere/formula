export type CellAddress = string;

export function colToName(col: number): string {
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function nameToCol(name: string): number | null {
  let acc = 0;
  for (const raw of name) {
    const c = raw.toUpperCase();
    if (c < "A" || c > "Z") return null;
    acc = acc * 26 + (c.charCodeAt(0) - 64);
  }
  return acc === 0 ? null : acc - 1;
}

export function parseCellAddress(addr: CellAddress): { row: number; col: number } | null {
  const noSheet = addr.includes("!") ? addr.split("!").slice(-1)[0] : addr;
  const match = /^(\$?[A-Za-z]+)(\$?\d+)$/.exec(noSheet);
  if (!match) return null;
  const colName = match[1].replace("$", "");
  const rowStr = match[2].replace("$", "");
  const col = nameToCol(colName);
  if (col == null) return null;
  const rowNum = Number.parseInt(rowStr, 10);
  if (!Number.isFinite(rowNum) || rowNum <= 0) return null;
  return { row: rowNum - 1, col };
}

export function formatCellAddress(row: number, col: number): CellAddress {
  return `${colToName(col)}${row + 1}`;
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

