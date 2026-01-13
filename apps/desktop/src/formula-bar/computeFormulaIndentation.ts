const INDENT_WIDTH = 2;
const MAX_INDENT_LEVEL = 20;

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

/**
 * Computes indentation (spaces) to insert after an Alt+Enter newline while editing
 * a formula.
 *
 * The indentation level is derived from formula structure up to the caret:
 * - Count `(` minus `)` while ignoring any parentheses that appear inside string literals.
 * - Clamp indentation to a reasonable maximum so pathological inputs don't create huge whitespace runs.
 */
export function computeFormulaIndentation(text: string, cursor: number): string {
  const cursorPos = Math.max(0, Math.min(cursor, text.length));

  let inString = false;
  let parenDepth = 0;
  let lastSignificantOutsideString: string | null = null;

  for (let i = 0; i < cursorPos; i++) {
    const ch = text[i]!;

    if (ch === '"') {
      if (inString) {
        // Excel-style escaping: `""` inside a string literal is an escaped `"`.
        if (text[i + 1] === '"') {
          i++; // Skip the escaped quote.
          continue;
        }
        inString = false;
      } else {
        inString = true;
      }
      continue;
    }

    if (inString) continue;

    if (ch === "(") {
      parenDepth++;
    } else if (ch === ")") {
      parenDepth = Math.max(0, parenDepth - 1);
    }

    if (!isWhitespace(ch)) {
      lastSignificantOutsideString = ch;
    }
  }

  // Don't auto-indent when breaking a string literal; extra spaces would become part of the string.
  if (inString) return "";

  // Optional continuation indent: if the preceding token is a comma but we're not inside parentheses
  // (e.g. array constants / union expressions), add a single indentation level.
  if (parenDepth === 0 && lastSignificantOutsideString === ",") {
    parenDepth = 1;
  }

  const clamped = Math.max(0, Math.min(parenDepth, MAX_INDENT_LEVEL));
  return " ".repeat(clamped * INDENT_WIDTH);
}

