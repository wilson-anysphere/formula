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

function colNameToIndex(colName: string): number {
  if (colName.trim() === "") {
    throw new Error("Expected a non-empty column name");
  }

  let n = 0;
  for (const ch of colName.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`Invalid column name: ${colName}`);
    }
    n = n * 26 + (code - 64);
  }
  return n - 1;
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

export function fromA1(address: string): { row0: number; col0: number } {
  const trimmed = address.trim();
  const match = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(trimmed);
  if (!match) {
    throw new Error(`Invalid A1 address: ${address}`);
  }

  const [, colName, rowStr] = match;
  const row1 = Number(rowStr);
  if (!Number.isInteger(row1) || row1 < 1) {
    throw new Error(`Invalid row in A1 address: ${address}`);
  }

  return { row0: row1 - 1, col0: colNameToIndex(colName) };
}
