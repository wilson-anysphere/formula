import type { FormulaToken } from "./types.js";

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
  let prev = start - 1;
  while (prev >= 0 && isWhitespace(input[prev])) prev -= 1;
  if (prev >= 0 && isIdentifierPart(input[prev])) return null;

  if (!isIdentifierStart(input[start] ?? "")) return null;

  let i = start + 1;
  while (i < input.length && isIdentifierPart(input[i])) i += 1;
  if (input[i] === "!") {
    const sheetName = input.slice(start, i);
    if (
      isReservedUnquotedSheetName(sheetName) ||
      looksLikeA1CellReference(sheetName) ||
      looksLikeR1C1CellReference(sheetName)
    ) {
      return null;
    }
    return { text: input.slice(start, i + 1), end: i + 1 };
  }
  return null;
}

function tryReadCellRef(input: string, start: number): { text: string; end: number } | null {
  let i = start;
  if (input[i] === "$") i += 1;

  let colStart = i;
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

export function tokenizeFormula(input: string): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let i = 0;

  while (i < input.length) {
    const ch = input[i];

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

      // Standalone `#` (or followed by whitespace) is still a spill operator.
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
        if (tokens[j]?.type !== "whitespace") return tokens[j];
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
      while (i < input.length && isIdentifierPart(input[i])) i += 1;
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
