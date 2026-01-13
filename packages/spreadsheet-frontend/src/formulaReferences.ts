import { tokenizeFormula } from "./formula/tokenizeFormula.ts";
import { parseStructuredReferenceText } from "./formula/structuredReferences.ts";

export type FormulaReferenceRange = {
  /**
   * 0-based row/column indices (inclusive).
   *
   * Note: Grid renderers typically use end-exclusive ranges; callers should
   * convert as needed (`endRow + 1`, `endCol + 1`).
   */
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  sheet?: string;
};

export type FormulaReference = {
  /** The original reference text as written in the formula (e.g. "A1", "Sheet1!$A$1:$B$2"). */
  text: string;
  /** Normalized 0-based, inclusive coordinates. */
  range: FormulaReferenceRange;
  /** Stable ordering by appearance in the formula string. */
  index: number;
  /** Start offset in the formula string (0-based, inclusive). */
  start: number;
  /** End offset in the formula string (0-based, exclusive). */
  end: number;
};

export type ExtractedFormulaReferences = {
  references: FormulaReference[];
  /**
   * Index of the reference that should be treated as "active" (for replacement),
   * based on the current cursor/selection. `null` when the caret is not within
   * a reference token.
   */
  activeIndex: number | null;
};

export type ColoredFormulaReference = FormulaReference & { color: string };

export type StructuredTableInfo = {
  /** Table name as used in structured references (e.g. `Table1`). */
  name: string;
  /** Column names in table order. */
  columns: readonly string[];
  /** Optional sheet name the table belongs to. */
  sheet?: string;
  /** Legacy alias (matches the completion schema provider contract). */
  sheetName?: string;
  /** 0-based inclusive coordinates of the full table range (including header row). */
  startRow?: number;
  startCol?: number;
  endRow?: number;
  endCol?: number;
};

export type ExtractFormulaReferencesOptions = {
  /**
   * Optional resolver for non-A1 identifiers (e.g. named ranges).
   *
   * If provided, any identifier token that is not a function call (no immediate `(`)
   * and is not the TRUE/FALSE literals will be passed through this resolver. Only
   * identifiers that resolve to a range will be included in the output.
   */
  resolveName?: (name: string) => FormulaReferenceRange | null;
  /**
   * Optional resolver for structured references (e.g. `Table1[Amount]`).
   *
   * When provided, this is consulted after A1 parsing fails.
   */
  resolveStructuredRef?: (refText: string) => FormulaReferenceRange | null;
  /**
   * Optional table metadata used to resolve Excel structured references into A1-style
   * coordinates for highlighting.
   */
  tables?: ReadonlyMap<string, StructuredTableInfo> | ReadonlyArray<StructuredTableInfo>;
};

export const FORMULA_REFERENCE_PALETTE: readonly string[] = [
  // Excel-ish palette (blue, red, green, purple, teal, orange, …).
  "#4F81BD",
  "#C0504D",
  "#9BBB59",
  "#8064A2",
  "#4BACC6",
  "#F79646",
  "#1F497D",
  "#943634"
];

export function extractFormulaReferences(
  input: string,
  cursorStart?: number,
  cursorEnd?: number,
  opts?: ExtractFormulaReferencesOptions
): ExtractedFormulaReferences {
  const tokens = tokenizeFormula(input);
  const references: FormulaReference[] = [];
  let refIndex = 0;

  for (const token of tokens) {
    if (token.type === "reference") {
      const parsed = parseA1RangeWithSheet(token.text) ?? resolveStructuredReference(token.text, opts);
      if (!parsed) continue;
      references.push({
        text: token.text,
        range: parsed,
        index: refIndex++,
        start: token.start,
        end: token.end
      });
      continue;
    }

    if (token.type === "identifier" && opts?.resolveName) {
      const lower = token.text.toLowerCase();
      if (lower === "true" || lower === "false") continue;
      const resolved = opts.resolveName(token.text);
      if (!resolved) continue;
      references.push({
        text: token.text,
        range: resolved,
        index: refIndex++,
        start: token.start,
        end: token.end
      });
      continue;
    }
  }

  const activeIndex =
    cursorStart === undefined || cursorEnd === undefined ? null : findActiveReferenceIndex(references, cursorStart, cursorEnd);

  return { references, activeIndex };
}

export function assignFormulaReferenceColors(
  references: readonly FormulaReference[],
  previousByText: ReadonlyMap<string, string> | null | undefined
): { colored: ColoredFormulaReference[]; nextByText: Map<string, string> } {
  const prev = previousByText ?? new Map<string, string>();
  const usedColors = new Set<string>();
  const nextByText = new Map<string, string>();

  // Build a stable list of unique references (by text), preserving first-appearance order.
  const uniqueByText: Array<{ text: string; firstIndex: number }> = [];
  for (const reference of references) {
    if (nextByText.has(reference.text)) continue;
    nextByText.set(reference.text, "");
    uniqueByText.push({ text: reference.text, firstIndex: reference.index });
  }
  // Reset placeholders; we'll fill them in deterministically below.
  nextByText.clear();

  // First pass: reuse previous colors for any references that still exist, so inserting a new
  // reference earlier in the formula doesn't "steal" colors from existing refs.
  for (const entry of uniqueByText) {
    const color = prev.get(entry.text);
    if (!color) continue;
    if (usedColors.has(color)) continue;
    nextByText.set(entry.text, color);
    usedColors.add(color);
  }

  // Second pass: assign fresh colors to new references (or ones whose previous color
  // collided), walking the Excel-ish palette in order.
  for (const entry of uniqueByText) {
    if (nextByText.has(entry.text)) continue;

    const color =
      FORMULA_REFERENCE_PALETTE.find((candidate) => !usedColors.has(candidate)) ??
      FORMULA_REFERENCE_PALETTE[entry.firstIndex % FORMULA_REFERENCE_PALETTE.length]!;

    nextByText.set(entry.text, color);
    usedColors.add(color);
  }

  const colored = references.map((reference) => ({ ...reference, color: nextByText.get(reference.text)! }));
  return { colored, nextByText };
}

function findActiveReferenceIndex(references: readonly FormulaReference[], cursorStart: number, cursorEnd: number): number | null {
  const start = Math.min(cursorStart, cursorEnd);
  const end = Math.max(cursorStart, cursorEnd);

  // If text is selected, treat a reference as active only when the selection is
  // contained within that reference token.
  if (start !== end) {
    const active = references.find((ref) => start >= ref.start && end <= ref.end);
    return active ? active.index : null;
  }

  // Caret: treat the reference containing either the character at the caret or
  // immediately before it as active. This matches typical editor behavior where
  // being at the end of a token still counts as "in" the token.
  const positions = start === 0 ? [0] : [start, start - 1];
  for (const pos of positions) {
    const active = references.find((ref) => ref.start <= pos && pos < ref.end);
    if (active) return active.index;
  }
  return null;
}

function columnLettersToIndex(letters: string): number | null {
  let col = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return null;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

function parseCellRef(cell: string): { row: number; col: number } | null {
  const match = /^\$?([A-Z]+)\$?([0-9]+)$/.exec(cell.toUpperCase());
  if (!match) return null;
  const col = columnLettersToIndex(match[1]!);
  if (col == null) return null;
  const row = Number.parseInt(match[2]!, 10) - 1;
  if (!Number.isFinite(row) || row < 0) return null;
  return { col, row };
}

function parseSheetAndRef(rangeRef: string): { sheet: string | undefined; ref: string } {
  const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(rangeRef);
  if (!match) return { sheet: undefined, ref: rangeRef };

  const rawSheet = match[1] ?? match[2];
  const sheetName = rawSheet ? rawSheet.replace(/''/g, "'") : undefined;
  return { sheet: sheetName, ref: match[3]! };
}

function parseA1RangeWithSheet(rangeRef: string): FormulaReferenceRange | null {
  const { sheet, ref } = parseSheetAndRef(rangeRef.trim());
  const [startRef, endRef] = ref.split(":", 2);
  const start = parseCellRef(startRef!);
  const end = parseCellRef(endRef ?? startRef!);
  if (!start || !end) return null;

  return {
    sheet,
    startCol: Math.min(start.col, end.col),
    startRow: Math.min(start.row, end.row),
    endCol: Math.max(start.col, end.col),
    endRow: Math.max(start.row, end.row)
  };
}

function resolveStructuredReference(refText: string, opts: ExtractFormulaReferencesOptions | undefined): FormulaReferenceRange | null {
  if (!opts) return null;

  if (typeof opts.resolveStructuredRef === "function") {
    const resolved = opts.resolveStructuredRef(refText);
    if (resolved) return resolved;
  }

  const tables = opts.tables;
  if (!tables) return null;
  const parsed = parseStructuredReferenceText(refText);
  if (!parsed) return null;

  const table = findTableInfo(tables, parsed.tableName);
  if (!table) return null;

  const startRow = table.startRow;
  const endRow = table.endRow;
  const startCol = table.startCol;
  const endCol = table.endCol;
  if (![startRow, endRow, startCol, endCol].every((v) => Number.isFinite(v))) return null;

  const baseStartCol = Math.min(startCol!, endCol!);
  const baseEndCol = Math.max(startCol!, endCol!);

  const baseStartRow = Math.min(startRow!, endRow!);
  const baseEndRow = Math.max(startRow!, endRow!);

  // Handle special structured reference specifiers like `Table1[#All]`, `Table1[#Data]`, etc.
  // These reference whole-table ranges (not a specific column).
  const maybeSpecifier = parsed.columnName.toLowerCase();
  if (maybeSpecifier.startsWith("#")) {
    const sheet = table.sheet ?? table.sheetName;
    if (!sheet) return null;

    if (maybeSpecifier === "#all") {
      return { sheet, startRow: baseStartRow, endRow: baseEndRow, startCol: baseStartCol, endCol: baseEndCol };
    }

    if (maybeSpecifier === "#headers") {
      return { sheet, startRow: baseStartRow, endRow: baseStartRow, startCol: baseStartCol, endCol: baseEndCol };
    }

    if (maybeSpecifier === "#totals") {
      return { sheet, startRow: baseEndRow, endRow: baseEndRow, startCol: baseStartCol, endCol: baseEndCol };
    }

    if (maybeSpecifier === "#data") {
      let refStartRow = baseStartRow;
      const refEndRow = baseEndRow;
      if (refEndRow > refStartRow) {
        // Exclude header row when the table has at least one data row.
        refStartRow = refStartRow + 1;
      }
      return {
        sheet,
        startRow: Math.min(refStartRow, refEndRow),
        endRow: Math.max(refStartRow, refEndRow),
        startCol: baseStartCol,
        endCol: baseEndCol
      };
    }
  }

  const columnIndex = table.columns.findIndex((c) => String(c).toLowerCase() === parsed.columnName.toLowerCase());
  if (columnIndex < 0) return null;

  const col = baseStartCol + columnIndex;
  if (col < baseStartCol || col > baseEndCol) return null;

  let refStartRow = baseStartRow;
  let refEndRow = baseEndRow;

  const selector = parsed.selector?.toLowerCase() ?? null;
  if (selector === "#headers") {
    refEndRow = refStartRow;
  } else if (selector === "#totals") {
    // Best-effort: treat totals as the last row of the table range.
    refStartRow = baseEndRow;
    refEndRow = baseEndRow;
  } else if (selector === "#all") {
    // Keep the full range, including headers.
  } else {
    // Default / #Data: exclude header row when the table has at least one data row.
    if (refEndRow > refStartRow) {
      refStartRow = refStartRow + 1;
    }
  }

  return {
    sheet: table.sheet ?? table.sheetName,
    startRow: Math.min(refStartRow, refEndRow),
    endRow: Math.max(refStartRow, refEndRow),
    startCol: col,
    endCol: col
  };
}

function findTableInfo(
  tables: ReadonlyMap<string, StructuredTableInfo> | ReadonlyArray<StructuredTableInfo>,
  tableName: string
): StructuredTableInfo | null {
  const target = tableName.toLowerCase();

  if (isStructuredTableArray(tables)) {
    for (const table of tables) {
      const name = (table?.name ?? "").toString();
      if (name.toLowerCase() === target) return table;
    }
    return null;
  }

  const direct = tables.get(tableName);
  if (direct) return direct;

  for (const [key, table] of tables.entries()) {
    if (key.toLowerCase() === target) return table;
    if ((table?.name ?? "").toLowerCase() === target) return table;
  }
  return null;
}

function isStructuredTableArray(
  tables: ReadonlyMap<string, StructuredTableInfo> | ReadonlyArray<StructuredTableInfo>
): tables is ReadonlyArray<StructuredTableInfo> {
  return Array.isArray(tables);
}
