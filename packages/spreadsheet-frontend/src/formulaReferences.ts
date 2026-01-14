import { tokenizeFormula, type FormulaToken } from "./formula/tokenizeFormula.ts";
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
  const references = extractFormulaReferencesFromTokens(tokens, input, opts);

  const activeIndex =
    cursorStart === undefined || cursorEnd === undefined ? null : findActiveReferenceIndex(references, cursorStart, cursorEnd);

  return { references, activeIndex };
}

export function extractFormulaReferencesFromTokens(
  tokens: readonly FormulaToken[],
  input: string,
  opts?: ExtractFormulaReferencesOptions
): FormulaReference[] {
  const references: FormulaReference[] = [];
  let refIndex = 0;

  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]!;

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

      // Treat `IDENT (` as a function call even when there is whitespace between the name and "(".
      // Use the token stream rather than scanning the source string to avoid re-walking long inputs.
      let j = i + 1;
      while (j < tokens.length && tokens[j]!.type === "whitespace") j += 1;
      if (j < tokens.length) {
        const next = tokens[j]!;
        if (next.type === "punctuation" && next.text === "(") continue;
      }

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

  return references;
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

  const findContainingReference = (needleStart: number, needleEnd: number): number | null => {
    if (references.length === 0) return null;
    let lo = 0;
    let hi = references.length - 1;
    let candidate = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const ref = references[mid]!;
      if (ref.start <= needleStart) {
        candidate = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    if (candidate < 0) return null;
    const ref = references[candidate]!;
    if (ref.start <= needleStart && needleEnd <= ref.end) return ref.index;
    return null;
  };

  // If text is selected, treat a reference as active only when the selection is
  // contained within that reference token.
  if (start !== end) {
    return findContainingReference(start, end);
  }

  // Caret: treat the reference containing either the character at the caret or
  // immediately before it as active. This matches typical editor behavior where
  // being at the end of a token still counts as "in" the token.
  const positions = start === 0 ? [0] : [start, start - 1];
  for (const pos of positions) {
    const active = findContainingReference(pos, pos + 1);
    if (active != null) return active;
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
  if (!rangeRef.includes("!")) return { sheet: undefined, ref: rangeRef };

  const scanQuotedToken = (start: number): { text: string; end: number } | null => {
    if (rangeRef[start] !== "'") return null;
    let i = start + 1;
    while (i < rangeRef.length) {
      if (rangeRef[i] === "'") {
        // Escaped quote inside quoted identifier: '' -> '
        if (rangeRef[i + 1] === "'") {
          i += 2;
          continue;
        }
        const raw = rangeRef.slice(start + 1, i);
        return { text: raw.replace(/''/g, "'"), end: i + 1 };
      }
      i += 1;
    }
    return null;
  };

  const scanUnquotedToken = (start: number): { text: string; end: number } | null => {
    let i = start;
    while (i < rangeRef.length) {
      const ch = rangeRef[i] ?? "";
      if (ch === ":" || ch === "!") break;
      i += 1;
    }
    if (i === start) return null;
    return { text: rangeRef.slice(start, i), end: i };
  };

  const scanNameToken = (start: number): { text: string; end: number } | null => {
    if (rangeRef[start] === "'") return scanQuotedToken(start);
    return scanUnquotedToken(start);
  };

  // Parse either:
  // - Sheet!A1
  // - Sheet1:Sheet3!A1
  // - 'Sheet 1'!A1
  // - 'Sheet 1':'Sheet 3'!A1
  // - mixed quoting: Sheet1:'Sheet 3'!A1 / 'Sheet 1':Sheet3!A1
  const first = scanNameToken(0);
  if (!first) return { sheet: undefined, ref: rangeRef };

  let pos = first.end;
  if (rangeRef[pos] === ":") {
    const second = scanNameToken(pos + 1);
    if (!second) return { sheet: undefined, ref: rangeRef };
    pos = second.end;
    if (rangeRef[pos] !== "!") return { sheet: undefined, ref: rangeRef };
    const sheet = `${first.text}:${second.text}`;
    return { sheet, ref: rangeRef.slice(pos + 1) };
  }

  if (rangeRef[pos] === "!") {
    return { sheet: first.text, ref: rangeRef.slice(pos + 1) };
  }

  // Not actually a sheet-qualified ref (e.g. might be `A1:B2`); treat as unqualified.
  return { sheet: undefined, ref: rangeRef };
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

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function unescapeStructuredRefItem(text: string): string {
  // Excel escapes `]` inside structured reference items by doubling it: `]]` -> `]`.
  return text.replaceAll("]]", "]");
}

function findMatchingStructuredRefBracketEnd(src: string, start: number): number | null {
  // Structured references escape closing brackets inside items by doubling: `]]` -> literal `]`.
  // That makes naive depth counting incorrect (it will pop twice for an escaped bracket).
  //
  // This matches the bracket span using a small backtracking parser:
  // - On `[` increase depth.
  // - On `]]`, prefer treating it as an escape (consume 2 chars; depth unchanged), but remember
  //   a choice point. If we later fail to close all brackets, backtrack and reinterpret that
  //   `]]` as a real closing bracket.
  if (src[start] !== "[") return null;

  let i = start;
  let depth = 0;
  const escapeChoices: Array<{ i: number; depth: number }> = [];

  const backtrack = (): boolean => {
    const choice = escapeChoices.pop();
    if (!choice) return false;
    i = choice.i;
    depth = choice.depth;
    // Reinterpret the first `]` of the `]]` pair as a real closing bracket.
    depth -= 1;
    i += 1;
    return true;
  };

  while (true) {
    if (i >= src.length) {
      // Unclosed bracket span.
      if (!backtrack()) return null;
      if (depth === 0) return i;
      continue;
    }

    const ch = src[i] ?? "";
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }

    if (ch === "]") {
      if (src[i + 1] === "]" && depth > 0) {
        // Prefer treating `]]` as an escaped literal `]` inside an item. Record a choice point
        // so we can reinterpret it as a closing bracket if needed.
        escapeChoices.push({ i, depth });
        i += 2;
        continue;
      }

      depth -= 1;
      i += 1;
      if (depth === 0) return i;
      if (depth < 0) {
        // Too many closing brackets - try reinterpreting an earlier escape.
        if (!backtrack()) return null;
        if (depth === 0) return i;
      }
      continue;
    }

    i += 1;
  }
}

type ParsedNestedStructuredReference = {
  tableName: string;
  /**
   * Structured reference selector items (e.g. `#All`, `#Headers`, `#Data`).
   *
   * Excel permits multiple selector items in a single structured reference, for example:
   *   Table1[[#Headers],[#Data],[Col1]]
   *   Table1[[#All],[#Totals]]
   *
   * These represent a union of table areas. We only resolve them when the resulting
   * range is a single contiguous rectangle (otherwise we'd highlight a misleading
   * bounding box).
   */
  selectors: string[];
  /**
   * Column tokens referenced by the structured reference.
   *
   * When empty, the structured reference selects all table columns.
   */
  columnItems: string[];
  /** Separators between `columnItems` (length = `columnItems.length - 1`). */
  columnSeparators: Array<"," | ":">;
};

const KNOWN_SELECTOR_ITEMS = new Set(["#all", "#headers", "#data", "#totals", "#this row"]);

function normalizeStructuredRefItem(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function parseNestedStructuredReference(refText: string): ParsedNestedStructuredReference | null {
  const firstBracket = refText.indexOf("[");
  if (firstBracket < 0) return null;
  const tableName = refText.slice(0, firstBracket);
  const suffix = refText.slice(firstBracket);
  if (!suffix.startsWith("[[")) return null;

  // Parse the nested `[[...]]` list using bracket-aware scanning so `]]` escapes in column names
  // don't confuse the separator detection.
  const rawItems: string[] = [];
  const separators: Array<"," | ":"> = [];

  let pos = 1; // Start at the second `[` (beginning of the first item).
  while (true) {
    if (suffix[pos] !== "[") return null;
    const end = findMatchingStructuredRefBracketEnd(suffix, pos);
    if (!end) return null;
    rawItems.push(suffix.slice(pos + 1, end - 1));
    pos = end;

    while (pos < suffix.length && isWhitespace(suffix[pos] ?? "")) pos += 1;
    const next = suffix[pos] ?? "";
    if (next === "," || next === ":") {
      separators.push(next);
      pos += 1;
      while (pos < suffix.length && isWhitespace(suffix[pos] ?? "")) pos += 1;
      continue;
    }

    if (next === "]") {
      // Outer `[[...]]` close. Must be the final char.
      pos += 1;
      while (pos < suffix.length && isWhitespace(suffix[pos] ?? "")) pos += 1;
      if (pos !== suffix.length) return null;
      break;
    }

    return null;
  }

  if (rawItems.length === 0) return null;

  const items = rawItems.map((item) => unescapeStructuredRefItem(item).trim());

  // Parse leading selector items (`#All`, `#Data`, …). These must be comma-separated.
  let selectorCount = 0;
  while (selectorCount < items.length) {
    const item = items[selectorCount] ?? "";
    const normalized = normalizeStructuredRefItem(item);
    if (!KNOWN_SELECTOR_ITEMS.has(normalized)) break;
    // Selector items cannot be joined via ":".
    if (selectorCount < separators.length && separators[selectorCount] === ":") return null;
    selectorCount += 1;
  }

  if (selectorCount > 0 && selectorCount < items.length) {
    // The boundary between the last selector and the first column token must be a comma.
    if (separators[selectorCount - 1] !== ",") return null;
  }

  const selectors = items.slice(0, selectorCount);
  const columnItems = items.slice(selectorCount);
  const columnSeparators = separators.slice(selectorCount);

  return { tableName, selectors, columnItems, columnSeparators };
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
  const nested = parsed ? null : parseNestedStructuredReference(refText);
  if (!parsed && !nested) return null;

  const table = findTableInfo(tables, parsed ? parsed.tableName : nested!.tableName);
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

  if (nested) {
    const normalizedSelectors =
      nested.selectors.length === 0 ? ["" /* default -> #Data */] : nested.selectors.map((s) => normalizeStructuredRefItem(s));

    type RowInterval = { startRow: number; endRow: number };
    const intervals: RowInterval[] = [];

    for (const selector of normalizedSelectors) {
      if (selector === "#this row") {
        // `#This Row` depends on the edited row context and is handled by UI-specific resolvers.
        return null;
      }
      if (selector !== "" && selector !== "#data" && selector !== "#headers" && selector !== "#totals" && selector !== "#all") {
        return null;
      }

      let start = baseStartRow;
      let end = baseEndRow;

      if (selector === "#headers") {
        end = start;
      } else if (selector === "#totals") {
        start = baseEndRow;
        end = baseEndRow;
      } else if (selector === "#all") {
        // Keep full range.
      } else {
        // Default / #Data: exclude header row when the table has at least one data row.
        if (end > start) start = start + 1;
      }

      intervals.push({ startRow: Math.min(start, end), endRow: Math.max(start, end) });
    }

    // Merge intervals; only resolve when the union is a single contiguous interval.
    intervals.sort((a, b) => a.startRow - b.startRow);
    const merged: RowInterval[] = [];
    for (const interval of intervals) {
      const last = merged[merged.length - 1];
      if (!last) {
        merged.push({ ...interval });
        continue;
      }
      if (interval.startRow <= last.endRow + 1) {
        last.endRow = Math.max(last.endRow, interval.endRow);
        continue;
      }
      merged.push({ ...interval });
    }
    if (merged.length !== 1) return null;
    const rowSpan = merged[0]!;

    const tableColumns = Array.isArray(table.columns) ? table.columns : [];

    const findColumnIndex = (name: string): number => {
      const target = String(name ?? "").toLowerCase();
      return tableColumns.findIndex((c) => String(c).toLowerCase() === target);
    };

    let startColRef = baseStartCol;
    let endColRef = baseEndCol;

    if (nested.columnItems.length > 0) {
      const indices = new Set<number>();
      const cols = nested.columnItems;
      const seps = nested.columnSeparators;

      let i = 0;
      while (i < cols.length) {
        const sep = i < seps.length ? seps[i] : null;
        if (sep === ":") {
          if (i + 1 >= cols.length) return null;
          const startName = cols[i] ?? "";
          const endName = cols[i + 1] ?? "";
          const startIdx = findColumnIndex(startName);
          const endIdx = findColumnIndex(endName);
          if (startIdx < 0 || endIdx < 0) return null;
          const lo = Math.min(startIdx, endIdx);
          const hi = Math.max(startIdx, endIdx);
          for (let idx = lo; idx <= hi; idx += 1) indices.add(idx);
          i += 2;
          if (i < cols.length && seps[i - 1] !== ",") return null;
          continue;
        }

        const idx = findColumnIndex(cols[i] ?? "");
        if (idx < 0) return null;
        indices.add(idx);
        i += 1;
        if (i < cols.length && seps[i - 1] !== ",") return null;
      }

      const uniqSorted = Array.from(indices).sort((a, b) => a - b);
      if (uniqSorted.length === 0) return null;
      const minIdx = uniqSorted[0]!;
      const maxIdx = uniqSorted[uniqSorted.length - 1]!;

      // Multi-column refs can represent unions. Only resolve when contiguous so we don't
      // highlight a misleading bounding rectangle.
      if (maxIdx - minIdx + 1 !== uniqSorted.length) return null;

      startColRef = baseStartCol + minIdx;
      endColRef = baseStartCol + maxIdx;
      if (startColRef < baseStartCol || endColRef > baseEndCol) return null;
    }

    return {
      sheet: table.sheet ?? table.sheetName,
      startRow: rowSpan.startRow,
      endRow: rowSpan.endRow,
      startCol: startColRef,
      endCol: endColRef,
    };
  }

  // Handle special structured reference specifiers like `Table1[#All]`, `Table1[#Data]`, etc.
  // These reference whole-table ranges (not a specific column).
  const maybeSpecifier = parsed!.columnName.toLowerCase();
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

  const columnIndex = table.columns.findIndex((c) => String(c).toLowerCase() === parsed!.columnName.toLowerCase());
  if (columnIndex < 0) return null;

  const col = baseStartCol + columnIndex;
  if (col < baseStartCol || col > baseEndCol) return null;

  let refStartRow = baseStartRow;
  let refEndRow = baseEndRow;

  const selector = parsed!.selector?.toLowerCase() ?? null;
  if (selector === "#headers") {
    refEndRow = refStartRow;
  } else if (selector === "#totals") {
    // Best-effort: treat totals as the last row of the table range.
    refStartRow = baseEndRow;
    refEndRow = baseEndRow;
  } else if (selector === "#all") {
    // Keep the full range, including headers.
  } else if (selector === null || selector === "#data") {
    // Default / #Data: exclude header row when the table has at least one data row.
    if (refEndRow > refStartRow) {
      refStartRow = refStartRow + 1;
    }
  } else {
    // Unsupported selector (e.g. `#This Row`) - we can't resolve it to a stable
    // rectangular range without additional context.
    return null;
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
