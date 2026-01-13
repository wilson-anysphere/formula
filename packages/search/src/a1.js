export function colToIndex(col) {
  const input = String(col).toUpperCase();
  if (!/^[A-Z]+$/.test(input)) {
    throw new Error(`Invalid column: ${col}`);
  }

  let n = 0;
  for (let i = 0; i < input.length; i++) {
    n = n * 26 + (input.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function indexToCol(index) {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`Invalid column index: ${index}`);
  }

  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function parseA1Address(input) {
  const s = String(input).trim();
  const m = s.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  if (!m) {
    throw new Error(`Invalid A1 address: ${input}`);
  }

  const [, colAbs, colLetters, rowAbs, rowStr] = m;
  const row = Number.parseInt(rowStr, 10) - 1;
  if (!Number.isFinite(row) || row < 0) {
    throw new Error(`Invalid A1 row: ${input}`);
  }

  return {
    row,
    col: colToIndex(colLetters),
    rowAbsolute: rowAbs === "$",
    colAbsolute: colAbs === "$",
  };
}

export function formatA1Address({ row, col, rowAbsolute = false, colAbsolute = false }) {
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid row/col for A1 formatting: row=${row} col=${col}`);
  }
  const colPart = `${colAbsolute ? "$" : ""}${indexToCol(col)}`;
  const rowPart = `${rowAbsolute ? "$" : ""}${row + 1}`;
  return `${colPart}${rowPart}`;
}

export function parseA1Range(input) {
  const s = String(input).trim();
  const parts = s.split(":");
  if (parts.length === 1) {
    const addr = parseA1Address(parts[0]);
    return { startRow: addr.row, endRow: addr.row, startCol: addr.col, endCol: addr.col };
  }
  if (parts.length !== 2) {
    throw new Error(`Invalid A1 range: ${input}`);
  }

  const a = parseA1Address(parts[0]);
  const b = parseA1Address(parts[1]);

  const startRow = Math.min(a.row, b.row);
  const endRow = Math.max(a.row, b.row);
  const startCol = Math.min(a.col, b.col);
  const endCol = Math.max(a.col, b.col);

  return { startRow, endRow, startCol, endCol };
}

export function formatA1Range(range) {
  // Excel-style shorthand formatting for full row/column selections.
  // This is primarily used by `parseGoTo("A:A")` / `parseGoTo("1:1")`, which expands
  // row/column references using Excel's default grid limits.
  const DEFAULT_MAX_ROWS = 1_048_576;
  const DEFAULT_MAX_COLS = 16_384;

  const isFullSheet =
    range.startRow === 0 &&
    range.endRow === DEFAULT_MAX_ROWS - 1 &&
    range.startCol === 0 &&
    range.endCol === DEFAULT_MAX_COLS - 1;
  if (!isFullSheet) {
    const isFullColumns = range.startRow === 0 && range.endRow === DEFAULT_MAX_ROWS - 1;
    if (isFullColumns) {
      const startCol = indexToCol(range.startCol);
      const endCol = indexToCol(range.endCol);
      return startCol === endCol ? `${startCol}:${startCol}` : `${startCol}:${endCol}`;
    }

    const isFullRows = range.startCol === 0 && range.endCol === DEFAULT_MAX_COLS - 1;
    if (isFullRows) {
      const startRow = String(range.startRow + 1);
      const endRow = String(range.endRow + 1);
      return startRow === endRow ? `${startRow}:${startRow}` : `${startRow}:${endRow}`;
    }
  }

  const start = formatA1Address({ row: range.startRow, col: range.startCol });
  const end = formatA1Address({ row: range.endRow, col: range.endCol });
  return start === end ? start : `${start}:${end}`;
}

export function splitSheetQualifier(input) {
  const s = String(input).trim();

  const quoted = s.match(/^'((?:[^']|'')+)'!(.+)$/);
  if (quoted) {
    return {
      sheetName: quoted[1].replace(/''/g, "'"),
      ref: quoted[2],
    };
  }

  const unquoted = s.match(/^([^!]+)!(.+)$/);
  if (unquoted) {
    return { sheetName: unquoted[1], ref: unquoted[2] };
  }

  return { sheetName: null, ref: s };
}
