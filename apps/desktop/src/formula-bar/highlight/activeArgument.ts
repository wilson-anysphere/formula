type ActiveArgumentSpan = {
  fnName: string;
  argIndex: number;
  argText: string;
  span: { start: number; end: number };
};

type StackFrame =
  | { kind: "function"; name: string; argIndex: number; argStart: number; parenIndex: number }
  | { kind: "group" }
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

function findMatchingBracketEnd(formulaText: string, start: number): number | null {
  // Structured references escape closing brackets inside items by doubling: `]]` -> literal `]`.
  // That makes naive depth counting incorrect (it will pop twice for an escaped bracket).
  //
  // We match the bracket span using a small backtracking parser:
  // - On `[[` (or any nested brackets), increase depth.
  // - On `]]`, prefer treating it as an escape (consume both, depth unchanged), but remember
  //   a choice point. If we later fail to close all brackets, backtrack and reinterpret that
  //   `]]` as a real closing bracket.
  //
  // This keeps argument parsing stable for structured refs like:
  //   Table1[[#Headers],[A]]B],[Col2]]
  if (formulaText[start] !== "[") return null;

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
    if (i >= formulaText.length) {
      // Unclosed bracket span.
      if (!backtrack()) return null;
      continue;
    }

    const ch = formulaText[i] ?? "";
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }

    if (ch === "]") {
      if (formulaText[i + 1] === "]" && depth > 0) {
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
      }
      continue;
    }

    i += 1;
  }
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

function findArgumentEnd(formulaText: string, start: number): number {
  let inString = false;
  let inSheetQuote = false;
  let parenDepth = 0;
  let braceDepth = 0;

  let i = start;
  while (i < formulaText.length) {
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

    if (ch === "[") {
      const end = findMatchingBracketEnd(formulaText, i);
      if (!end) return formulaText.length;
      i = end;
      continue;
    }

    if (ch === "{") {
      braceDepth += 1;
      i += 1;
      continue;
    }
    if (ch === "}") {
      if (braceDepth > 0) braceDepth -= 1;
      i += 1;
      continue;
    }

    // Treat parentheses inside `[]` / `{}` as plain characters so structured references
    // like `Table1[Amount)]` don't accidentally break argument parsing.
    if (ch === "(") {
      parenDepth += 1;
      i += 1;
      continue;
    }
    if (ch === ")") {
      if (parenDepth > 0) {
        parenDepth -= 1;
        i += 1;
        continue;
      }
      if (braceDepth === 0) return i;
      i += 1;
      continue;
    }

    if ((ch === "," || ch === ";") && parenDepth === 0 && braceDepth === 0) return i;

    i += 1;
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

    // Bracketed structured references/external refs should be treated as opaque. Otherwise
    // column names like `Table1[Amount(USD)]` or escaped brackets like `A]]B` could be
    // misread as nested calls / argument separators.
    if (ch === "[") {
      const end = findMatchingBracketEnd(formulaText, i);
      if (!end || end > cursor) {
        // Cursor is inside an unterminated bracket span - stop scanning, but keep the
        // function stack intact so we can still return the enclosing function call.
        i = cursor;
        continue;
      }
      i = end;
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
