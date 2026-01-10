function assertFiniteNonNegativeInt(value, label) {
  if (!Number.isFinite(value) || value < 0 || !Number.isInteger(value)) {
    throw new Error(`${label} must be a finite non-negative integer. Received: ${value}`);
  }
}

export function columnLabelToIndex(label) {
  if (!/^[A-Z]+$/i.test(label)) {
    throw new Error(`Invalid column label: ${label}`);
  }

  let value = 0;
  for (const ch of label.toUpperCase()) {
    const code = ch.charCodeAt(0);
    value = value * 26 + (code - 64);
  }
  return value - 1;
}

export function indexToColumnLabel(index) {
  assertFiniteNonNegativeInt(index, "index");

  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function parseCellAddress(a1) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(a1.trim());
  if (!match) {
    throw new Error(`Invalid A1 cell address: ${a1}`);
  }

  const col = columnLabelToIndex(match[1]);
  const row = Number.parseInt(match[2], 10) - 1;
  if (!Number.isFinite(row) || row < 0) {
    throw new Error(`Invalid row in A1 address: ${a1}`);
  }

  return { row, col };
}

export function formatCellAddress(coord) {
  assertFiniteNonNegativeInt(coord.row, "row");
  assertFiniteNonNegativeInt(coord.col, "col");

  return `${indexToColumnLabel(coord.col)}${coord.row + 1}`;
}

export function parseRangeAddress(a1) {
  const trimmed = a1.trim();
  const parts = trimmed.split(":");
  if (parts.length === 1) {
    const cell = parseCellAddress(parts[0]);
    return { startRow: cell.row, startCol: cell.col, endRow: cell.row, endCol: cell.col };
  }
  if (parts.length !== 2) {
    throw new Error(`Invalid A1 range address: ${a1}`);
  }

  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);

  return { startRow, startCol, endRow, endCol };
}

export function formatRangeAddress(range) {
  assertFiniteNonNegativeInt(range.startRow, "startRow");
  assertFiniteNonNegativeInt(range.startCol, "startCol");
  assertFiniteNonNegativeInt(range.endRow, "endRow");
  assertFiniteNonNegativeInt(range.endCol, "endCol");

  if (range.startRow > range.endRow || range.startCol > range.endCol) {
    throw new Error(`Invalid range coordinates: ${JSON.stringify(range)}`);
  }

  const start = formatCellAddress({ row: range.startRow, col: range.startCol });
  const end = formatCellAddress({ row: range.endRow, col: range.endCol });
  return start === end ? start : `${start}:${end}`;
}
