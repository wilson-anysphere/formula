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

export function extractFormulaReferences(input: string, cursorStart?: number, cursorEnd?: number): ExtractedFormulaReferences {
  const tokens = tokenizeFormula(input);
  const references: FormulaReference[] = [];
  let refIndex = 0;

  for (const token of tokens) {
    if (token.type !== "reference") continue;
    const parsed = parseA1RangeWithSheet(token.text);
    if (!parsed) continue;
    references.push({
      text: token.text,
      range: parsed,
      index: refIndex++,
      start: token.start,
      end: token.end
    });
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

type FormulaTokenType =
  | "whitespace"
  | "operator"
  | "punctuation"
  | "number"
  | "string"
  | "function"
  | "identifier"
  | "reference"
  | "error"
  | "unknown";

type FormulaToken = {
  type: FormulaTokenType;
  text: string;
  start: number;
  end: number;
};

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

const UNICODE_LETTER_RE: RegExp | null = (() => {
  try {
    return new RegExp("^\\p{Alphabetic}$", "u");
  } catch {
    return null;
  }
})();

const UNICODE_ALNUM_RE: RegExp | null = (() => {
  try {
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    return null;
  }
})();

function isUnicodeAlphabetic(ch: string): boolean {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return isAsciiLetter(ch);
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isAsciiLetter(ch) || isDigit(ch);
}

function isIdentifierStart(ch: string): boolean {
  return ch === "_" || isUnicodeAlphabetic(ch);
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || ch === "." || isUnicodeAlphanumeric(ch);
}

function isErrorBodyChar(ch: string): boolean {
  return (
    ch === "_" ||
    ch === "/" ||
    ch === "." ||
    (ch >= "0" && ch <= "9") ||
    (ch >= "A" && ch <= "Z") ||
    (ch >= "a" && ch <= "z")
  );
}

function tryReadString(input: string, start: number): { text: string; end: number } | null {
  if (input[start] !== '"') return null;
  let i = start + 1;
  while (i < input.length) {
    if (input[i] === '"') {
      if (input[i + 1] === '"') {
        i += 2;
        continue;
      }
      return { text: input.slice(start, i + 1), end: i + 1 };
    }
    i += 1;
  }
  return { text: input.slice(start), end: input.length };
}

function tryReadErrorCode(input: string, start: number): { text: string; end: number } | null {
  if (input[start] !== "#") return null;

  // Mirror the engine lexer: treat `#` as an error literal when followed by a
  // plausible error body (letters/digits/_/./). Stop before the first
  // non-body character so we don't accidentally consume trailing punctuation.
  if (!isErrorBodyChar(input[start + 1] ?? "")) return null;

  let i = start + 1;
  while (i < input.length && isErrorBodyChar(input[i] ?? "")) i += 1;
  if (input[i] === "!" || input[i] === "?") i += 1;
  return { text: input.slice(start, i), end: i };
}

function tryReadNumber(input: string, start: number): { text: string; end: number } | null {
  const ch = input[start];
  if (!isDigit(ch) && !(ch === "." && isDigit(input[start + 1] ?? ""))) return null;

  let i = start;
  while (isDigit(input[i] ?? "")) i += 1;
  if (input[i] === ".") {
    i += 1;
    while (isDigit(input[i] ?? "")) i += 1;
  }

  if (input[i] === "e" || input[i] === "E") {
    const sign = input[i + 1];
    const signLen = sign === "+" || sign === "-" ? 1 : 0;
    if (isDigit(input[i + 1 + signLen] ?? "")) {
      i += 1 + signLen;
      while (isDigit(input[i] ?? "")) i += 1;
    }
  }

  return { text: input.slice(start, i), end: i };
}

function tryReadSheetPrefix(input: string, start: number): { text: string; end: number } | null {
  if (input[start] === "'") {
    // Excel escapes apostrophes inside sheet names using doubled quotes: ''.
    let i = start + 1;
    while (i < input.length) {
      if (input[i] === "'") {
        if (input[i + 1] === "'") {
          i += 2;
          continue;
        }
        if (input[i + 1] === "!") {
          return { text: input.slice(start, i + 2), end: i + 2 };
        }
        return null;
      }
      i += 1;
    }
    return null;
  }

  // Only treat `Sheet!A1` as a sheet-qualified ref when the `Sheet` token starts
  // at a natural boundary. This avoids incorrectly highlighting the tail of an
  // invalid unquoted sheet name that contains spaces (e.g. `My Sheet!A1` should
  // not be tokenized as `Sheet!A1`).
  let prev = start - 1;
  while (prev >= 0 && isWhitespace(input[prev]!)) prev -= 1;
  if (prev >= 0 && isIdentifierPart(input[prev]!)) return null;

  if (!isIdentifierStart(input[start] ?? "")) return null;

  let i = start + 1;
  while (i < input.length && isIdentifierPart(input[i]!)) i += 1;
  if (input[i] === "!") {
    return { text: input.slice(start, i + 1), end: i + 1 };
  }
  return null;
}

function tryReadCellRef(input: string, start: number): { text: string; end: number } | null {
  let i = start;
  if (input[i] === "$") i += 1;

  const colStart = i;
  while (i < input.length) {
    const ch = input[i];
    if (!((ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z"))) break;
    i += 1;
  }
  if (i === colStart) return null;
  const colLen = i - colStart;
  if (colLen > 3) return null;

  if (input[i] === "$") i += 1;

  const rowStart = i;
  while (isDigit(input[i] ?? "")) i += 1;
  if (i === rowStart) return null;

  return { text: input.slice(start, i), end: i };
}

function tryReadReference(input: string, start: number): { text: string; end: number } | null {
  let i = start;
  const sheet = tryReadSheetPrefix(input, i);
  if (sheet) i = sheet.end;

  const first = tryReadCellRef(input, i);
  if (!first) return null;
  i = first.end;

  if (input[i] === ":") {
    const second = tryReadCellRef(input, i + 1);
    if (second) {
      i = second.end;
      const prefix = sheet ? sheet.text : "";
      return { text: prefix + input.slice(sheet ? sheet.end : start, i), end: i };
    }
  }

  const prefix = sheet ? sheet.text : "";
  return { text: prefix + first.text, end: i };
}

function tokenizeFormula(input: string): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let i = 0;

  while (i < input.length) {
    const ch = input[i]!;

    if (isWhitespace(ch)) {
      const start = i;
      while (isWhitespace(input[i] ?? "")) i += 1;
      tokens.push({ type: "whitespace", text: input.slice(start, i), start, end: i });
      continue;
    }

    const str = tryReadString(input, i);
    if (str) {
      tokens.push({ type: "string", text: str.text, start: i, end: str.end });
      i = str.end;
      continue;
    }

    if (ch === "#") {
      // Spill-range postfix operator (`A1#`) vs error literal (`#REF!`).
      const prev = tokens[tokens.length - 1];
      const isPostfixSpill =
        prev &&
        prev.end === i &&
        prev.type !== "whitespace" &&
        (prev.type === "reference" ||
          prev.type === "identifier" ||
          prev.type === "function" ||
          (prev.type === "punctuation" && (prev.text === ")" || prev.text === "]")));
      if (isPostfixSpill) {
        tokens.push({ type: "operator", text: "#", start: i, end: i + 1 });
        i += 1;
        continue;
      }

      const err = tryReadErrorCode(input, i);
      if (err) {
        tokens.push({ type: "error", text: err.text, start: i, end: err.end });
        i = err.end;
        continue;
      }

      tokens.push({ type: "operator", text: "#", start: i, end: i + 1 });
      i += 1;
      continue;
    }

    const num = tryReadNumber(input, i);
    if (num) {
      tokens.push({ type: "number", text: num.text, start: i, end: num.end });
      i = num.end;
      continue;
    }

    // Disambiguation: `My Sheet!A1` is an invalid unquoted sheet-qualified reference
    // (sheet names containing spaces must be quoted as `'My Sheet'!A1`). When users
    // type it anyway, highlight just the cell reference (`A1`) rather than treating
    // `Sheet!A1` as a sheet-qualified reference and ignoring the `My ` prefix.
    const lastToken = tokens[tokens.length - 1];
    const precededByWhitespace = lastToken?.type === "whitespace";
    const prevNonWhitespace = (() => {
      for (let j = tokens.length - 1; j >= 0; j--) {
        if (tokens[j]?.type !== "whitespace") return tokens[j] ?? null;
      }
      return null;
    })();
    const possibleSheetPrefix = input[i] !== "'" && precededByWhitespace ? tryReadSheetPrefix(input, i) : null;

    if (!(possibleSheetPrefix && prevNonWhitespace?.type === "identifier")) {
      const ref = tryReadReference(input, i);
      if (ref) {
        tokens.push({ type: "reference", text: ref.text, start: i, end: ref.end });
        i = ref.end;
        continue;
      }
    }

    if (isIdentifierStart(ch)) {
      const start = i;
      i += 1;
      while (i < input.length && isIdentifierPart(input[i]!)) i += 1;
      const ident = input.slice(start, i);
      if (input[i] === "(") {
        tokens.push({ type: "function", text: ident, start, end: i });
      } else {
        tokens.push({ type: "identifier", text: ident, start, end: i });
      }
      continue;
    }

    const twoChar = input.slice(i, i + 2);
    if (twoChar === ">=" || twoChar === "<=" || twoChar === "<>") {
      tokens.push({ type: "operator", text: twoChar, start: i, end: i + 2 });
      i += 2;
      continue;
    }

    if ("+-*/^&=><%@".includes(ch)) {
      tokens.push({ type: "operator", text: ch, start: i, end: i + 1 });
      i += 1;
      continue;
    }

    if ("(),;:[]{}.!".includes(ch)) {
      tokens.push({ type: "punctuation", text: ch, start: i, end: i + 1 });
      i += 1;
      continue;
    }

    tokens.push({ type: "unknown", text: ch, start: i, end: i + 1 });
    i += 1;
  }

  return tokens;
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
