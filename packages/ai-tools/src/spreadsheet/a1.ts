export const DEFAULT_SHEET_NAME = "Sheet1";

export type SheetName = string;

export interface CellAddress {
  sheet: SheetName;
  row: number;
  col: number;
}

export interface RangeAddress {
  sheet: SheetName;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

// Support absolute references (e.g. $A$1) by allowing optional `$` markers.
const CELL_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i;

export function columnLabelToIndex(label: string): number {
  const normalized = label.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    throw new Error(`Invalid column label: ${label}`);
  }

  let value = 0;
  for (const char of normalized) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value;
}

export function columnIndexToLabel(index: number): string {
  if (!Number.isInteger(index) || index <= 0) {
    throw new Error(`Invalid column index: ${index}`);
  }

  let value = index;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function parseSheetPrefix(input: string, defaultSheet: string): { sheet: string; rest: string } {
  const bangIndex = input.lastIndexOf("!");
  if (bangIndex === -1) {
    return { sheet: defaultSheet, rest: input.trim() };
  }

  const rawSheet = input.slice(0, bangIndex).trim();
  const rest = input.slice(bangIndex + 1).trim();
  if (!rawSheet) {
    throw new Error(`Invalid A1 reference: missing sheet name before "!" in "${input}"`);
  }

  // Excel style: 'Sheet Name'!A1 (single quotes, '' to escape).
  const sheet =
    rawSheet.startsWith("'") && rawSheet.endsWith("'")
      ? rawSheet.slice(1, -1).replace(/''/g, "'")
      : rawSheet;

  if (!sheet) {
    throw new Error(`Invalid A1 reference: empty sheet name in "${input}"`);
  }

  return { sheet, rest };
}

export function parseA1Cell(input: string, defaultSheet: string = DEFAULT_SHEET_NAME): CellAddress {
  const { sheet, rest } = parseSheetPrefix(input, defaultSheet);
  const match = CELL_RE.exec(rest);
  if (!match) {
    throw new Error(`Invalid cell reference: "${input}"`);
  }

  const col = columnLabelToIndex(match[1]);
  const row = Number(match[2]);
  if (!Number.isFinite(row) || row <= 0) {
    throw new Error(`Invalid row number in cell reference: "${input}"`);
  }

  return { sheet, row, col };
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiAlphaNum(ch: string): boolean {
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function isReservedUnquotedSheetName(name: string): boolean {
  // Excel boolean literals are tokenized as keywords; quoting avoids ambiguity in formulas.
  const lower = String(name ?? "").toLowerCase();
  return lower === "true" || lower === "false";
}

function looksLikeA1CellReference(name: string): boolean {
  // If an unquoted sheet name looks like a cell reference (e.g. "A1" or "XFD1048576"),
  // Excel requires quoting to disambiguate.
  let i = 0;
  let letters = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiLetter(ch)) break;
    if (letters.length >= 3) return false;
    letters += ch;
    i += 1;
  }

  if (letters.length === 0) return false;

  let digits = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiDigit(ch)) break;
    digits += ch;
    i += 1;
  }

  if (digits.length === 0) return false;
  if (i !== name.length) return false;

  const col = letters
    .split("")
    .reduce((acc, c) => acc * 26 + (c.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1), 0);
  return col <= 16_384;
}

function looksLikeR1C1CellReference(name: string): boolean {
  // In R1C1 notation, `R`/`C` are valid relative references. Excel may also treat
  // `R123C456` as a cell reference even when the workbook is in A1 mode.
  const upper = String(name ?? "").toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;

  let i = 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;

  i += 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  return i === upper.length;
}

function isValidUnquotedSheetNameForA1(name: string): boolean {
  if (!name) return false;
  const first = name[0];
  if (!first || isAsciiDigit(first)) return false;
  if (!(first === "_" || isAsciiLetter(first))) return false;
  for (let i = 1; i < name.length; i += 1) {
    const ch = name[i]!;
    if (!(isAsciiAlphaNum(ch) || ch === "_" || ch === ".")) return false;
  }

  if (isReservedUnquotedSheetName(name)) return false;
  if (looksLikeA1CellReference(name) || looksLikeR1C1CellReference(name)) return false;

  return true;
}

function formatSheetName(sheet: string): string {
  // Excel style: quote sheet names containing spaces/special characters
  // using single quotes and escaping embedded quotes via doubling.
  //
  // Avoid emitting ambiguous identifiers like `TRUE!A1`, `A1!A1`, or `R1C1!A1`.
  if (isValidUnquotedSheetNameForA1(sheet)) return sheet;
  return `'${sheet.replace(/'/g, "''")}'`;
}

export function formatA1Cell(address: CellAddress): string {
  return `${formatSheetName(address.sheet)}!${columnIndexToLabel(address.col)}${address.row}`;
}

export function parseA1Range(input: string, defaultSheet: string = DEFAULT_SHEET_NAME): RangeAddress {
  const { sheet, rest } = parseSheetPrefix(input, defaultSheet);
  const parts = rest.split(":").map((part) => part.trim());
  if (parts.length === 0 || parts.length > 2) {
    throw new Error(`Invalid range reference: "${input}"`);
  }

  const start = parseA1Cell(parts[0], sheet);
  const end = parts.length === 2 ? parseA1Cell(parts[1], sheet) : start;

  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);

  return { sheet, startRow, startCol, endRow, endCol };
}

export function formatA1Range(range: RangeAddress): string {
  const start = `${columnIndexToLabel(range.startCol)}${range.startRow}`;
  const end = `${columnIndexToLabel(range.endCol)}${range.endRow}`;
  const body = start === end ? start : `${start}:${end}`;
  return `${formatSheetName(range.sheet)}!${body}`;
}

export function rangeSize(range: RangeAddress): { rows: number; cols: number } {
  return {
    rows: range.endRow - range.startRow + 1,
    cols: range.endCol - range.startCol + 1
  };
}
