export interface Range0 {
  startRow0: number;
  endRow0Exclusive: number;
  startCol0: number;
  endCol0Exclusive: number;
}

export function colToName(col0: number): string {
  if (!Number.isSafeInteger(col0) || col0 < 0) {
    throw new Error(`colToName: col0 must be a non-negative safe integer, got ${col0}`);
  }

  let col = col0 + 1;
  let name = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    col = Math.floor((col - 1) / 26);
  }
  return name;
}

export function toA1(row0: number, col0: number): string {
  if (!Number.isSafeInteger(row0) || row0 < 0) {
    throw new Error(`toA1: row0 must be a non-negative safe integer, got ${row0}`);
  }
  return `${colToName(col0)}${row0 + 1}`;
}

export function fromA1(address: string): { row0: number; col0: number } {
  const trimmed = address.trim();
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)$/.exec(trimmed);
  if (!match) {
    throw new Error(`Invalid A1 address: "${address}"`);
  }

  const colLabel = match[1].toUpperCase();
  let col1 = 0;
  for (const ch of colLabel) {
    col1 = col1 * 26 + (ch.charCodeAt(0) - 64);
  }
  const row1 = Number.parseInt(match[2], 10);

  if (!Number.isSafeInteger(row1) || row1 <= 0) {
    throw new Error(`Invalid A1 address row: "${address}"`);
  }
  return { row0: row1 - 1, col0: col1 - 1 };
}

export function range0ToA1(range: Range0): string {
  if (range.endRow0Exclusive <= range.startRow0 || range.endCol0Exclusive <= range.startCol0) {
    throw new Error(`Invalid range0 (empty): ${JSON.stringify(range)}`);
  }
  const start = toA1(range.startRow0, range.startCol0);
  const end = toA1(range.endRow0Exclusive - 1, range.endCol0Exclusive - 1);
  return start === end ? start : `${start}:${end}`;
}

