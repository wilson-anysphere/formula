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

function findArgumentEnd(formulaText: string, start: number): number {
  let inString = false;
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

    if (ch === '"') {
      inString = true;
      continue;
    }

    if (ch === "(") {
      parenDepth += 1;
      continue;
    }
    if (ch === ")") {
      if (parenDepth > 0) {
        parenDepth -= 1;
        continue;
      }
      if (bracketDepth === 0 && braceDepth === 0) return i;
      continue;
    }

    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
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

    if (ch === '"') {
      inString = true;
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

    if (ch === "[") {
      stack.push({ kind: "bracket" });
      i += 1;
      continue;
    }

    if (ch === "]") {
      popUntilKind(stack, "bracket");
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
