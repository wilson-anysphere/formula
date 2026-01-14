type ActiveArgumentSpan = {
  fnName: string;
  argIndex: number;
  argText: string;
  span: { start: number; end: number };
};

type StackFrame =
  | { kind: "function"; name: string; argIndex: number; argStart: number; parenIndex: number }
  | { kind: "group" }
  | { kind: "bracket" }
  | { kind: "brace" };

function isIdentifierStart(ch: string): boolean {
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z") || ch === "_";
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || (ch >= "0" && ch <= "9") || ch === ".";
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function popUntilKind(stack: StackFrame[], kind: StackFrame["kind"]): void {
  for (let i = stack.length - 1; i >= 0; i -= 1) {
    if (stack[i]?.kind === kind) {
      stack.length = i;
      return;
    }
  }
}

function popParenFrame(stack: StackFrame[]): void {
  for (let i = stack.length - 1; i >= 0; i -= 1) {
    const frame = stack[i];
    if (frame?.kind === "function" || frame?.kind === "group") {
      stack.length = i;
      return;
    }
  }
}

function isEscapedStructuredRefBracket(formulaText: string, index: number, bracketDepth: number): boolean {
  // Excel escapes `]` inside structured reference items by doubling it: `]]` -> `]`.
  //
  // We only need a best-effort heuristic here: treat `]]` as an escaped literal `]`
  // when it's followed by a character that couldn't immediately follow a closing
  // bracket group. This matches the heuristic used by the formula tokenizer.
  if (formulaText[index] !== "]" || formulaText[index + 1] !== "]") return false;
  // When only a single bracket group is open, `]]` can't represent closing multiple
  // nested groups (there's nothing to pop twice). Treat it as an escaped literal.
  if (bracketDepth === 1) return true;
  // `]]]...` implies at least one escaped `]` (consume the first two as the escape).
  if (formulaText[index + 2] === "]") return true;
  let k = index + 2;
  while (k < formulaText.length && isWhitespace(formulaText[k] ?? "")) k += 1;
  const after = formulaText[k] ?? "";
  const isDelimiterAfterClose = after === "" || after === "," || after === ";" || after === "]" || after === ")";
  return !isDelimiterAfterClose;
}

function findArgumentEnd(formulaText: string, start: number): number {
  let inString = false;
  let inSheetQuote = false;
  let parenDepth = 0;
  let bracketDepth = 0;
  let braceDepth = 0;

  for (let i = start; i < formulaText.length; i += 1) {
    const ch = formulaText[i];

    if (inString) {
      if (ch === '"') {
        if (formulaText[i + 1] === '"') {
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }

    if (inSheetQuote) {
      if (ch === "'") {
        // Excel escapes apostrophes in sheet names by doubling them: '' -> '
        if (formulaText[i + 1] === "'") {
          i += 1;
          continue;
        }
        inSheetQuote = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }

    if (ch === "'") {
      inSheetQuote = true;
      continue;
    }

    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
      if (bracketDepth > 0 && isEscapedStructuredRefBracket(formulaText, i, bracketDepth)) {
        // Skip the escaped `]]` sequence without closing the bracket scope.
        i += 1;
        continue;
      }
      if (bracketDepth > 0) bracketDepth -= 1;
      continue;
    }

    if (ch === "{") {
      braceDepth += 1;
      continue;
    }
    if (ch === "}") {
      if (braceDepth > 0) braceDepth -= 1;
      continue;
    }

    // Treat parentheses inside `[]` / `{}` as plain characters so structured references
    // like `Table1[Amount)]` don't accidentally break argument parsing.
    if (ch === "(" && bracketDepth === 0) {
      parenDepth += 1;
      continue;
    }
    if (ch === ")" && bracketDepth === 0) {
      if (parenDepth > 0) {
        parenDepth -= 1;
        continue;
      }
      if (bracketDepth === 0 && braceDepth === 0) return i;
      continue;
    }

    if ((ch === "," || ch === ";") && parenDepth === 0 && bracketDepth === 0 && braceDepth === 0) return i;
  }

  return formulaText.length;
}

/**
 * Returns the innermost function call at `cursorIndex` and the current argument span.
 *
 * This is used for argument-hint rendering and live preview evaluation. It is
 * intentionally shallow (does not build an AST), but it properly ignores commas
 * inside:
 *  - string literals
 *  - nested parentheses
 *  - square brackets (`[]`) (structured / external references)
 *  - curly braces (`{}`) (array literals)
 */
export function getActiveArgumentSpan(formulaText: string, cursorIndex: number): ActiveArgumentSpan | null {
  const cursor = Math.max(0, Math.min(cursorIndex, formulaText.length));
  const stack: StackFrame[] = [];

  let i = 0;
  let inString = false;
  let inSheetQuote = false;

  while (i < cursor) {
    const ch = formulaText[i];

    if (inString) {
      if (ch === '"') {
        if (formulaText[i + 1] === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }

    if (inSheetQuote) {
      if (ch === "'") {
        // Excel escapes apostrophes in sheet names by doubling them: '' -> '
        if (formulaText[i + 1] === "'") {
          i += 2;
          continue;
        }
        inSheetQuote = false;
      }
      i += 1;
      continue;
    }

    if (ch === '"') {
      inString = true;
      i += 1;
      continue;
    }

    if (ch === "'") {
      inSheetQuote = true;
      i += 1;
      continue;
    }

    // Bracketed structured references/external refs should be treated as opaque
    // (except for nested bracket matching). Otherwise column names like
    // `Table1[Amount(USD)]` could be misread as function calls.
    if (ch === "[") {
      stack.push({ kind: "bracket" });
      i += 1;
      continue;
    }
    if (ch === "]") {
      const bracketDepth = stack.filter((frame) => frame.kind === "bracket").length;
      if (bracketDepth > 0 && isEscapedStructuredRefBracket(formulaText, i, bracketDepth)) {
        // Skip escaped bracket sequences inside structured reference items so
        // delimiters like commas within column names don't break arg parsing.
        i += 2;
        continue;
      }
      popUntilKind(stack, "bracket");
      i += 1;
      continue;
    }

    const inBracket = stack.some((frame) => frame.kind === "bracket");
    if (inBracket) {
      i += 1;
      continue;
    }

    if (isIdentifierStart(ch)) {
      const start = i;
      i += 1;
      while (i < cursor && isIdentifierPart(formulaText[i])) i += 1;
      const name = formulaText.slice(start, i).toUpperCase();

      const next = formulaText[i];
      if (next === "(" && i < cursor) {
        stack.push({ kind: "function", name, argIndex: 0, argStart: i + 1, parenIndex: i });
        i += 1;
        continue;
      }

      continue;
    }

    if (ch === "(") {
      stack.push({ kind: "group" });
      i += 1;
      continue;
    }

    if (ch === ")") {
      popParenFrame(stack);
      i += 1;
      continue;
    }

    if (ch === "{") {
      stack.push({ kind: "brace" });
      i += 1;
      continue;
    }

    if (ch === "}") {
      popUntilKind(stack, "brace");
      i += 1;
      continue;
    }

    if (ch === "," || ch === ";") {
      const top = stack[stack.length - 1];
      if (top?.kind === "function") {
        top.argIndex += 1;
        top.argStart = i + 1;
      }
      i += 1;
      continue;
    }

    i += 1;
  }

  for (let s = stack.length - 1; s >= 0; s -= 1) {
    const frame = stack[s];
    if (frame?.kind !== "function") continue;

    const boundary = findArgumentEnd(formulaText, frame.argStart);

    let start = frame.argStart;
    while (start < boundary && isWhitespace(formulaText[start] ?? "")) start += 1;

    let end = boundary;
    while (end > start && isWhitespace(formulaText[end - 1] ?? "")) end -= 1;

    return {
      fnName: frame.name,
      argIndex: frame.argIndex,
      argText: formulaText.slice(start, end),
      span: { start, end },
    };
  }

  return null;
}
