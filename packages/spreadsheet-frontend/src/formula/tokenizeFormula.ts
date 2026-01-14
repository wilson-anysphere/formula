export type FormulaTokenType =
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

export type FormulaToken = {
  type: FormulaTokenType;
  text: string;
  start: number;
  end: number;
};

function codePointAt(str: string, index: number): { ch: string; nextIndex: number } | null {
  if (index < 0 || index >= str.length) return null;
  const cp = str.codePointAt(index);
  if (cp == null) return null;
  return { ch: String.fromCodePoint(cp), nextIndex: index + (cp > 0xffff ? 2 : 1) };
}

function prevCodePointAt(str: string, index: number): { ch: string; startIndex: number } | null {
  if (index <= 0 || index > str.length) return null;
  let i = index - 1;

  // If we're at the second code unit of a surrogate pair, step back to the high surrogate.
  const codeUnit = str.charCodeAt(i);
  if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff && i - 1 >= 0) {
    const prev = str.charCodeAt(i - 1);
    if (prev >= 0xd800 && prev <= 0xdbff) i -= 1;
  }

  const cp = str.codePointAt(i);
  if (cp == null) return null;
  return { ch: String.fromCodePoint(cp), startIndex: i };
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isReservedUnquotedSheetName(name: string): boolean {
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
    if (!ch || !isDigit(ch)) break;
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
  const upper = String(name ?? "").toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;

  let i = 1;
  while (i < upper.length && isDigit(upper[i] ?? "")) i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;

  i += 1;
  while (i < upper.length && isDigit(upper[i] ?? "")) i += 1;
  return i === upper.length;
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
    ch === "¡" ||
    ch === "¿" ||
    isUnicodeAlphanumeric(ch)
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
  // non-body character so we don't accidentally consume trailing punctuation
  // like `]` in `Table1[#All]`.
  if (!isErrorBodyChar(input[start + 1] ?? "")) return null;

  let i = start + 1;
  while (i < input.length && isErrorBodyChar(input[i] ?? "")) i += 1;
  // Error literals may optionally end in `!` or `?` (e.g. `#REF!`, `#NAME?`).
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
  let scan = start;
  while (true) {
    const prev = prevCodePointAt(input, scan);
    if (!prev) break;
    if (isWhitespace(prev.ch)) {
      scan = prev.startIndex;
      continue;
    }
    if (isIdentifierPart(prev.ch)) return null;
    break;
  }

  const first = codePointAt(input, start);
  if (!first) return null;

  if (first.ch === "[") {
    // External workbook prefix: `[Book1.xlsx]Sheet1!A1`
    let i = first.nextIndex;
    while (i < input.length) {
      // Excel escapes `]` in external workbook names by doubling: `]]` -> literal `]`.
      // Continue scanning for the *real* closing bracket of the workbook prefix.
      if (input[i] === "]") {
        if (input[i + 1] === "]") {
          i += 2;
          continue;
        }
        break;
      }
      i += 1;
    }
    if (i >= input.length || input[i] !== "]") return null;
    i += 1;
    const sheetStart = codePointAt(input, i);
    if (!sheetStart || !isIdentifierStart(sheetStart.ch)) return null;

    let j = sheetStart.nextIndex;
    while (j < input.length) {
      const next = codePointAt(input, j);
      if (!next) break;
      if (next.ch === ":" || isIdentifierPart(next.ch)) {
        j = next.nextIndex;
        continue;
      }
      break;
    }
    if (input[j] === "!") {
      const sheetSpec = input.slice(i, j);
      const sheetNames = sheetSpec.split(":");
      if (
        sheetNames.some(
          (name) =>
            isReservedUnquotedSheetName(name) ||
            looksLikeA1CellReference(name) ||
            looksLikeR1C1CellReference(name)
        )
      ) {
        return null;
      }
      return { text: input.slice(start, j + 1), end: j + 1 };
    }
    return null;
  }

  if (!isIdentifierStart(first.ch)) return null;

  let i = first.nextIndex;
  while (i < input.length) {
    const next = codePointAt(input, i);
    if (!next) break;
    if (next.ch === ":" || isIdentifierPart(next.ch)) {
      i = next.nextIndex;
      continue;
    }
    break;
  }
  if (input[i] === "!") {
    const sheetSpec = input.slice(start, i);
    const sheetNames = sheetSpec.split(":");
    if (
      sheetNames.some(
        (name) =>
          isReservedUnquotedSheetName(name) ||
          looksLikeA1CellReference(name) ||
          looksLikeR1C1CellReference(name)
      )
    ) {
      return null;
    }
    return { text: input.slice(start, i + 1), end: i + 1 };
  }
  return null;
}

function tryReadExternalWorkbookNameRef(input: string, start: number): { text: string; end: number } | null {
  // Workbook-scoped external defined names can appear in unquoted form:
  //   [Book.xlsx]MyName
  //
  // The engine serializer prefers the quoted form (`'[Book.xlsx]MyName'`) to avoid structured-ref
  // ambiguity, but the parser accepts the unquoted form and users may type it directly.
  if (input[start] !== "[") return null;

  // Scan the workbook prefix, treating `[` as a literal character and honoring Excel's `]]` escape
  // for literal `]` characters inside the workbook id.
  let i = start + 1;
  while (i < input.length) {
    if (input[i] === "]") {
      if (input[i + 1] === "]") {
        i += 2;
        continue;
      }
      i += 1;
      break;
    }
    i += 1;
  }
  if (i <= start + 1 || i > input.length) return null;

  // Optional whitespace between the workbook prefix and the name token.
  while (i < input.length && isWhitespace(input[i] ?? "")) i += 1;

  const first = codePointAt(input, i);
  if (!first || !isIdentifierStart(first.ch)) return null;

  let end = first.nextIndex;
  while (end < input.length) {
    const next = codePointAt(input, end);
    if (!next) break;
    if (!isIdentifierPart(next.ch)) break;
    end = next.nextIndex;
  }

  return { text: input.slice(start, end), end };
}

function tryReadQuotedIdentifier(input: string, start: number): { text: string; end: number } | null {
  if (input[start] !== "'") return null;

  // Quoted identifiers use Excel-style escaping where `''` represents a literal `'` inside the
  // identifier. This is primarily used for quoted sheet names (`'My Sheet'!A1`), but workbook-scoped
  // external defined-name references are also commonly represented as a single quoted token:
  //   `'[Book.xlsx]MyName'`
  // so we tokenize it as an `identifier` for syntax highlighting.
  let i = start + 1;
  while (i < input.length) {
    if (input[i] === "'") {
      if (input[i + 1] === "'") {
        i += 2;
        continue;
      }
      return { text: input.slice(start, i + 1), end: i + 1 };
    }
    i += 1;
  }

  // Unterminated quote - best-effort: treat the rest of the input as the identifier.
  return { text: input.slice(start), end: input.length };
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

  // Mirror the engine lexer: avoid mis-highlighting defined names that start with a cell-ref
  // prefix (e.g. `A1FOO`, `R1C1Sheet`), since Excel allows such names because they do not fully
  // match the cell-reference grammar.
  const next = input[i] ?? "";
  if (next === "(" || (next !== "." && isIdentifierPart(next))) return null;

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

const MODE_NORMAL = 0;
const MODE_STRING = 1;
const MODE_QUOTED_IDENT = 2;

function encodeState(depth: number, mode: number): number {
  return (depth << 2) | (mode & 0b11);
}

function decodeDepth(state: number): number {
  return state >> 2;
}

function decodeMode(state: number): number {
  return state & 0b11;
}

function pushUnique(list: number[], value: number): void {
  if (!list.includes(value)) list.push(value);
}

function pushDepth(states: number[][], pos: number, depth: number): void {
  if (pos < 0 || pos >= states.length) return;
  pushUnique(states[pos]!, depth);
}

function pushState(states: number[][], pos: number, depth: number, mode: number): void {
  if (mode !== MODE_NORMAL && depth !== 0) return;
  if (pos < 0 || pos >= states.length) return;
  pushUnique(states[pos]!, encodeState(depth, mode));
}

/**
 * Find the end position (exclusive) for a bracketed segment starting at `start`.
 *
 * This mirrors the Rust engine's lexer disambiguation for `]]` inside structured refs/workbook
 * references:
 * - `]]` can be either an escaped literal `]` (consume 2 chars; depth unchanged), or
 * - a closing bracket followed by another `]` (consume 1 char; depth-1; next `]` processed normally).
 *
 * The scanner explores both interpretations and returns the earliest closing bracket that can
 * still lead to a globally valid parse of the remainder (e.g. doesn't end inside an unterminated
 * string literal and doesn't leave stray `]` tokens outside bracketed segments).
 */
function findBracketEnd(src: string, start: number): number | null {
  if (src[start] !== "[") return null;

  // Track code-unit indices so returned offsets match DOM selection semantics.
  const chars: string[] = [];
  for (let i = start; i < src.length; i += 1) chars.push(src[i]!);
  if (chars[0] !== "[") return null;

  const n = chars.length;
  if (n < 2) return null;

  const fwd: number[][] = Array.from({ length: n + 1 }, () => []);
  fwd[1]!.push(1);
  for (let i = 1; i < n; i += 1) {
    const depths = fwd[i]!;
    if (depths.length === 0) continue;
    for (const depth of depths) {
      if (depth === 0) continue;
      const ch = chars[i]!;
      if (ch === "[") {
        pushDepth(fwd, i + 1, depth + 1);
      } else if (ch === "]") {
        pushDepth(fwd, i + 1, depth - 1);
        if (chars[i + 1] === "]") pushDepth(fwd, i + 2, depth);
      } else {
        pushDepth(fwd, i + 1, depth);
      }
    }
  }

  const bwd: number[][] = Array.from({ length: n + 1 }, () => []);
  bwd[n]!.push(encodeState(0, MODE_NORMAL));
  for (let i = n - 1; i >= 0; i -= 1) {
    const ch = chars[i]!;
    if (ch === "[") {
      for (const state of bwd[i + 1]!) {
        const depthAfter = decodeDepth(state);
        const modeAfter = decodeMode(state);
        if (modeAfter !== MODE_NORMAL) {
          pushState(bwd, i, 0, modeAfter);
        } else if (depthAfter > 0) {
          pushState(bwd, i, depthAfter - 1, MODE_NORMAL);
        }
      }
      continue;
    }

    if (ch === "]") {
      for (const state of bwd[i + 1]!) {
        const depthAfter = decodeDepth(state);
        const modeAfter = decodeMode(state);
        if (modeAfter !== MODE_NORMAL) {
          pushState(bwd, i, 0, modeAfter);
        } else {
          pushState(bwd, i, depthAfter + 1, MODE_NORMAL);
        }
      }

      // Escape transitions are only valid while inside brackets (depth > 0).
      if (chars[i + 1] === "]") {
        for (const state of bwd[i + 2] ?? []) {
          const depthAfter = decodeDepth(state);
          const modeAfter = decodeMode(state);
          if (modeAfter === MODE_NORMAL && depthAfter > 0) pushState(bwd, i, depthAfter, MODE_NORMAL);
        }
      }
      continue;
    }

    if (ch === '"') {
      for (const state of bwd[i + 1]!) {
        const depthAfter = decodeDepth(state);
        const modeAfter = decodeMode(state);
        if (modeAfter === MODE_STRING && depthAfter === 0) {
          // Opening quote (`"`), entering string literal.
          pushState(bwd, i, 0, MODE_NORMAL);
          continue;
        }
        if (modeAfter === MODE_NORMAL) {
          if (depthAfter > 0) {
            // Quotes are literal characters inside brackets.
            pushState(bwd, i, depthAfter, MODE_NORMAL);
            continue;
          }
          if (chars[i + 1] !== '"') {
            // Closing quote (`"`), exiting string literal.
            pushState(bwd, i, 0, MODE_STRING);
          }
          continue;
        }
        if (modeAfter === MODE_QUOTED_IDENT && depthAfter === 0) {
          // Quotes are literal characters inside quoted identifiers.
          pushState(bwd, i, 0, MODE_QUOTED_IDENT);
        }
      }

      // Escaped quote (`""`) within a string literal.
      if (chars[i + 1] === '"') {
        for (const state of bwd[i + 2] ?? []) {
          const depthAfter = decodeDepth(state);
          const modeAfter = decodeMode(state);
          if (depthAfter === 0 && modeAfter === MODE_STRING) pushState(bwd, i, 0, MODE_STRING);
        }
      }
      continue;
    }

    if (ch === "'") {
      for (const state of bwd[i + 1]!) {
        const depthAfter = decodeDepth(state);
        const modeAfter = decodeMode(state);
        if (modeAfter === MODE_QUOTED_IDENT && depthAfter === 0) {
          // Opening quote (`'`), entering quoted identifier.
          pushState(bwd, i, 0, MODE_NORMAL);
          continue;
        }
        if (modeAfter === MODE_NORMAL) {
          if (depthAfter > 0) {
            // Quotes are literal characters inside brackets.
            pushState(bwd, i, depthAfter, MODE_NORMAL);
            continue;
          }
          if (chars[i + 1] !== "'") {
            // Closing quote (`'`), exiting quoted identifier.
            pushState(bwd, i, 0, MODE_QUOTED_IDENT);
          }
          continue;
        }
        if (modeAfter === MODE_STRING && depthAfter === 0) {
          // Quotes are literal characters inside string literals.
          pushState(bwd, i, 0, MODE_STRING);
        }
      }

      // Escaped quote (`''`) within a quoted identifier.
      if (chars[i + 1] === "'") {
        for (const state of bwd[i + 2] ?? []) {
          const depthAfter = decodeDepth(state);
          const modeAfter = decodeMode(state);
          if (depthAfter === 0 && modeAfter === MODE_QUOTED_IDENT) pushState(bwd, i, 0, MODE_QUOTED_IDENT);
        }
      }
      continue;
    }

    for (const state of bwd[i + 1]!) {
      const depthAfter = decodeDepth(state);
      const modeAfter = decodeMode(state);
      pushState(bwd, i, depthAfter, modeAfter);
    }
  }

  const goal = encodeState(0, MODE_NORMAL);
  for (let endIdx = 1; endIdx <= n; endIdx += 1) {
    if (fwd[endIdx]!.includes(0) && bwd[endIdx]!.includes(goal)) {
      const end = start + endIdx;
      return Math.min(end, src.length);
    }
  }
  return null;
}

function tryReadStructuredReference(input: string, start: number): { text: string; end: number } | null {
  const first = codePointAt(input, start);
  if (!first) return null;
  if (!isIdentifierStart(first.ch)) return null;

  let i = first.nextIndex;
  while (i < input.length) {
    const next = codePointAt(input, i);
    if (!next) break;
    if (!isIdentifierPart(next.ch)) break;
    i = next.nextIndex;
  }
  if (input[i] !== "[") return null;

  const end = findBracketEnd(input, i);
  if (!end) return null;
  const text = input.slice(start, end);
  return { text, end };
}

function tryReadImplicitStructuredReference(input: string, start: number): { text: string; end: number } | null {
  if (input[start] !== "[") return null;

  // Implicit structured references only apply in table context. To avoid mis-tokenizing other
  // bracket constructs (e.g. external workbook prefixes), only recognize:
  //   [@Column]
  //   [@[Column Name]]
  //   [@]
  //   [[#This Row],[Column]]
  // (and related selector-qualified forms).
  let scan = start + 1;
  while (scan < input.length && isWhitespace(input[scan] ?? "")) scan += 1;
  const next = input[scan] ?? "";
  if (next !== "@" && next !== "[") return null;

  // Use the same `]]` disambiguation logic as table-qualified structured references, otherwise
  // implicit refs like `[[#All],[Amount]]` could incorrectly "eat" into later string literals
  // containing `]]` (e.g. `... & "]]"`).
  const end = findBracketEnd(input, start);
  if (!end) return null;
  return { text: input.slice(start, end), end };
}

export function tokenizeFormula(input: string): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let i = 0;

  while (i < input.length) {
    const cp = codePointAt(input, i);
    const ch = cp?.ch ?? "";
    const nextIndex = cp?.nextIndex ?? i + 1;

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
      // Best-effort: treat `#` as a postfix operator only when immediately after
      // an expression-like token (no intervening whitespace).
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
        tokens.push({ type: "operator", text: "#", start: i, end: nextIndex });
        i = nextIndex;
        continue;
      }

      const err = tryReadErrorCode(input, i);
      if (err) {
        tokens.push({ type: "error", text: err.text, start: i, end: err.end });
        i = err.end;
        continue;
      }

      // Standalone `#` (or followed by whitespace) is still a spill operator.
      tokens.push({ type: "operator", text: "#", start: i, end: nextIndex });
      i = nextIndex;
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
    const possibleSheetPrefix = ch !== "'" && precededByWhitespace ? tryReadSheetPrefix(input, i) : null;

    if (!(possibleSheetPrefix && prevNonWhitespace?.type === "identifier")) {
      const ref = tryReadReference(input, i);
      if (ref) {
        tokens.push({ type: "reference", text: ref.text, start: i, end: ref.end });
        i = ref.end;
        continue;
      }
    }

    const structured = tryReadStructuredReference(input, i);
    if (structured) {
      tokens.push({ type: "reference", text: structured.text, start: i, end: structured.end });
      i = structured.end;
      continue;
    }

    const implicitStructured = tryReadImplicitStructuredReference(input, i);
    if (implicitStructured) {
      tokens.push({ type: "reference", text: implicitStructured.text, start: i, end: implicitStructured.end });
      i = implicitStructured.end;
      continue;
    }

    const externalNameRef = tryReadExternalWorkbookNameRef(input, i);
    if (externalNameRef) {
      tokens.push({ type: "identifier", text: externalNameRef.text, start: i, end: externalNameRef.end });
      i = externalNameRef.end;
      continue;
    }

    const quotedIdent = tryReadQuotedIdentifier(input, i);
    if (quotedIdent) {
      tokens.push({ type: "identifier", text: quotedIdent.text, start: i, end: quotedIdent.end });
      i = quotedIdent.end;
      continue;
    }

    if (isIdentifierStart(ch)) {
      const start = i;
      i = nextIndex;
      while (i < input.length) {
        const next = codePointAt(input, i);
        if (!next) break;
        if (!isIdentifierPart(next.ch)) break;
        i = next.nextIndex;
      }
      const ident = input.slice(start, i);
      // Excel permits whitespace between a function name and the opening paren (e.g. `SUM (A1)`).
      // Treat that as a function token for highlighting/autocomplete parity with the engine lexer.
      let scan = i;
      while (scan < input.length && isWhitespace(input[scan] ?? "")) scan += 1;
      if (input[scan] === "(") {
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
      tokens.push({ type: "operator", text: ch, start: i, end: nextIndex });
      i = nextIndex;
      continue;
    }

    if ("(),;:[]{}.!".includes(ch)) {
      tokens.push({ type: "punctuation", text: ch, start: i, end: nextIndex });
      i = nextIndex;
      continue;
    }

    tokens.push({ type: "unknown", text: ch, start: i, end: nextIndex });
    i = nextIndex;
  }

  return tokens;
}
