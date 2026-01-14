type ActiveArgumentSpan = {
  fnName: string;
  argIndex: number;
  argText: string;
  span: { start: number; end: number };
};

type GetActiveArgumentSpanOptions = {
  /**
   * Argument separator characters to treat as delimiting function arguments.
   *
   * Defaults to both `,` and `;` so the parser works reasonably across locales.
   * Callers that know the workbook/UI locale should provide the active list separator
   * (e.g. `;` for many comma-decimal locales) so we don't misinterpret decimal commas
   * inside numeric literals as argument separators.
   */
  argSeparators?: string | readonly string[];
};

type StackFrame =
  | { kind: "function"; name: string; argIndex: number; argStart: number; parenIndex: number }
  | { kind: "group" }
  | { kind: "brace" };

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
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isUnicodeAlphabetic(ch) || (ch >= "0" && ch <= "9");
}

function isIdentifierStart(ch: string): boolean {
  return ch === "_" || isUnicodeAlphabetic(ch);
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || ch === "." || isUnicodeAlphanumeric(ch);
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function findWorkbookPrefixEnd(formulaText: string, start: number): number | null {
  // External workbook prefixes escape closing brackets by doubling: `]]` -> literal `]`.
  //
  // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
  if (formulaText[start] !== "[") return null;
  let i = start + 1;
  while (i < formulaText.length) {
    if (formulaText[i] === "]") {
      if (formulaText[i + 1] === "]") {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i += 1;
  }
  return null;
}

function findWorkbookPrefixEndIfValid(formulaText: string, start: number): number | null {
  const end = findWorkbookPrefixEnd(formulaText, start);
  if (!end) return null;

  const skipWs = (idx: number): number => {
    let i = idx;
    while (i < formulaText.length && isWhitespace(formulaText[i] ?? "")) i += 1;
    return i;
  };

  const scanQuotedSheetName = (idx: number): number | null => {
    if (formulaText[idx] !== "'") return null;
    let i = idx + 1;
    while (i < formulaText.length) {
      const ch = formulaText[i] ?? "";
      if (ch === "'") {
        // Excel escapes apostrophes inside quoted sheet names by doubling: '' -> '
        if (i + 1 < formulaText.length && formulaText[i + 1] === "'") {
          i += 2;
          continue;
        }
        return i + 1;
      }
      i += 1;
    }
    return null;
  };

  const scanUnquotedName = (idx: number): number | null => {
    if (idx >= formulaText.length) return null;
    const first = formulaText[idx] ?? "";
    if (!(first === "_" || isUnicodeAlphabetic(first))) return null;

    let i = idx + 1;
    while (i < formulaText.length) {
      const ch = formulaText[i] ?? "";
      // Be conservative: align with the Rust parser's unquoted identifier rules.
      if (ch === "_" || ch === "." || ch === "$" || isUnicodeAlphanumeric(ch)) {
        i += 1;
        continue;
      }
      break;
    }
    return i;
  };

  const scanSheetNameToken = (idx: number): number | null => {
    const i = skipWs(idx);
    if (i >= formulaText.length) return null;
    if (formulaText[i] === "'") return scanQuotedSheetName(i);
    return scanUnquotedName(i);
  };

  // Heuristic: only treat this as an external workbook prefix if it is immediately followed by:
  // - a sheet spec and `!` (e.g. `[Book.xlsx]Sheet1!A1`), OR
  // - a defined name identifier (e.g. `[Book.xlsx]MyName`).
  //
  // This avoids incorrectly treating nested structured references (which *are* nested) as workbook
  // prefixes while still supporting workbook names that contain `[` characters (Excel treats `[` as
  // plain text within workbook ids).
  const sheetEnd = scanSheetNameToken(end);
  if (sheetEnd != null) {
    let i = skipWs(sheetEnd);

    // External 3D span: `[Book.xlsx]Sheet1:Sheet3!A1`
    if (i < formulaText.length && formulaText[i] === ":") {
      i = scanSheetNameToken(i + 1) ?? i;
      i = skipWs(i);
    }

    if (i < formulaText.length && formulaText[i] === "!") return end;
  }

  // Workbook-scoped external defined name: `[Book.xlsx]MyName`.
  const nameStart = skipWs(end);
  if (scanUnquotedName(nameStart) != null) return end;

  return null;
}

function findMatchingStructuredRefBracketEnd(formulaText: string, start: number): number | null {
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

function findMatchingBracketEnd(formulaText: string, start: number): number | null {
  // Prefer structured-ref matching (supports nested `[[...]]`). If that fails, fall back to
  // workbook-prefix scanning which treats `[` as a literal character.
  return (
    findMatchingStructuredRefBracketEnd(formulaText, start) ??
    findWorkbookPrefixEndIfValid(formulaText, start)
  );
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

function findArgumentEnd(formulaText: string, start: number, isArgSeparator: (ch: string) => boolean): number {
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
      // Array literals (`{...}`) can contain unbalanced parentheses while the user is typing.
      // Treat them as plain characters so we don't leak `parenDepth` out of the brace scope and
      // mis-detect argument boundaries.
      if (braceDepth > 0) {
        i += 1;
        continue;
      }
      parenDepth += 1;
      i += 1;
      continue;
    }
    if (ch === ")") {
      if (braceDepth > 0) {
        i += 1;
        continue;
      }
      if (parenDepth > 0) {
        parenDepth -= 1;
        i += 1;
        continue;
      }
      return i;
    }

    if (isArgSeparator(ch ?? "") && parenDepth === 0 && braceDepth === 0) return i;

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
export function getActiveArgumentSpan(
  formulaText: string,
  cursorIndex: number,
  opts: GetActiveArgumentSpanOptions = {}
): ActiveArgumentSpan | null {
  const cursor = Math.max(0, Math.min(cursorIndex, formulaText.length));
  const stack: StackFrame[] = [];

  let i = 0;
  let inString = false;
  let inSheetQuote = false;

  const separatorsRaw = opts.argSeparators ?? [",", ";"];
  const separators = Array.isArray(separatorsRaw) ? separatorsRaw : [separatorsRaw];
  const isArgSeparator = (ch: string): boolean => separators.some((sep) => sep === ch);

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

      // Excel permits whitespace between a function name and the opening paren, e.g. `SUM (A1)`.
      // Skip whitespace so argument hints and previews remain stable even when users insert spaces/newlines.
      let scan = i;
      while (scan < cursor && isWhitespace(formulaText[scan] ?? "")) scan += 1;
      const next = formulaText[scan];
      if (next === "(" && scan < cursor) {
        stack.push({ kind: "function", name, argIndex: 0, argStart: scan + 1, parenIndex: scan });
        i = scan + 1;
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

    if (isArgSeparator(ch ?? "")) {
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

    const boundary = findArgumentEnd(formulaText, frame.argStart, isArgSeparator);

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
