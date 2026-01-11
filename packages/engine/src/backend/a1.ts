export function colToName(col0: number): string {
  if (!Number.isInteger(col0) || col0 < 0) {
    throw new Error(`Expected a 0-based column index >= 0, got ${col0}`);
  }

  let n = col0 + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function toA1(row0: number, col0: number): string {
  if (!Number.isInteger(row0) || row0 < 0) {
    throw new Error(`Expected a 0-based row index >= 0, got ${row0}`);
  }
  return `${colToName(col0)}${row0 + 1}`;
}

export function toA1Range(startRow0: number, startCol0: number, endRow0: number, endCol0: number): string {
  const start = toA1(startRow0, startCol0);
  const end = toA1(endRow0, endCol0);
  return start === end ? start : `${start}:${end}`;
}

