import type { FormulaToken } from "./types.js";

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isIdentifierStart(ch: string): boolean {
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z") || ch === "_";
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || isDigit(ch) || ch === ".";
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
  let i = start + 1;
  while (i < input.length) {
    const ch = input[i];
    if (isWhitespace(ch) || ch === "," || ch === ")" || ch === "(" || ch === "+" || ch === "-" || ch === "*" || ch === "/") {
      break;
    }
    i += 1;
  }
  if (i === start + 1) return null;
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

  if (!isIdentifierStart(input[start] ?? "")) return null;

  let i = start + 1;
  while (i < input.length && isIdentifierPart(input[i])) i += 1;
  if (input[i] === "!") {
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

    const err = tryReadErrorCode(input, i);
    if (err) {
      tokens.push({ type: "error", text: err.text, start: i, end: err.end });
      i = err.end;
      continue;
    }

    const num = tryReadNumber(input, i);
    if (num) {
      tokens.push({ type: "number", text: num.text, start: i, end: num.end });
      i = num.end;
      continue;
    }

    const ref = tryReadReference(input, i);
    if (ref) {
      tokens.push({ type: "reference", text: ref.text, start: i, end: ref.end });
      i = ref.end;
      continue;
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

    if ("+-*/^&=><".includes(ch)) {
      tokens.push({ type: "operator", text: ch, start: i, end: i + 1 });
      i += 1;
      continue;
    }

    if ("(),;:".includes(ch)) {
      tokens.push({ type: "punctuation", text: ch, start: i, end: i + 1 });
      i += 1;
      continue;
    }

    tokens.push({ type: "unknown", text: ch, start: i, end: i + 1 });
    i += 1;
  }

  return tokens;
}
