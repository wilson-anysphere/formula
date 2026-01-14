/**
 * @typedef {{
 *   isFormula: boolean,
 *   inFunctionCall: boolean,
 *   functionName?: string,
 *   argIndex?: number,
 *   expectingRange?: boolean,
 *   functionNamePrefix?: { text: string, start: number, end: number },
 *   currentArg?: { text: string, start: number, end: number }
 * }} PartialFormulaContext
 */

/**
 * Extremely lightweight "partial parser" used for tab-completion.
 *
 * This is not a full Excel parser. It exists solely to answer questions like:
 * - Are we currently inside a function call?
 * - What is the function name?
 * - Which argument index are we on?
 * - Does that argument typically want a range?
 *
 * @param {string} input
 * @param {number} cursorPosition
 * @param {{ isRangeArg: (fnName: string, argIndex: number) => boolean }} functionRegistry
 * @returns {PartialFormulaContext}
 */
export function parsePartialFormula(input, cursorPosition, functionRegistry) {
  // This parser runs on every keystroke; it must be crash-proof even if called
  // with unexpected inputs.
  const safeInput = typeof input === "string" ? input : "";
  const safeCursor = clampCursor(safeInput, cursorPosition);
  const prefix = safeInput.slice(0, safeCursor);

  if (!prefix.startsWith("=")) {
    return { isFormula: false, inFunctionCall: false };
  }

  // Scan for unbalanced parentheses in the prefix to determine whether the
  // cursor is currently inside a (...) argument list.
  /** @type {number[]} */
  const openParens = [];
  let inString = false;
  let inSheetQuote = false;
  let bracketDepth = 0;
  let braceDepth = 0;
  for (let i = 0; i < prefix.length; i++) {
    const ch = prefix[i];
    if (inString) {
      if (ch === '"') {
        // Excel escapes quotes inside string literals via doubled quotes: "".
        if (prefix[i + 1] === '"') {
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }
    if (inSheetQuote) {
      if (ch === "'") {
        // Excel escapes apostrophes inside sheet names via doubled quotes: ''.
        if (prefix[i + 1] === "'") {
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
    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      continue;
    }
    if (ch === "{") {
      braceDepth += 1;
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      continue;
    }
    if (bracketDepth !== 0) continue;
    if (ch === "(") {
      openParens.push(i);
    } else if (ch === ")") {
      openParens.pop();
    }
  }

  if (openParens.length === 0) {
    // Not in a function call; still might be typing a function name.
    const functionPrefix = findTokenAtCursor(prefix, safeCursor);
    if (functionPrefix && functionPrefix.text.length > 0) {
      return {
        isFormula: true,
        inFunctionCall: false,
        functionNamePrefix: functionPrefix,
      };
    }
    return { isFormula: true, inFunctionCall: false };
  }

  const openParenIndex = openParens[openParens.length - 1];
  const fnName = findFunctionNameBeforeParen(prefix, openParenIndex);
  const argContext = getArgContext(prefix, openParenIndex, safeCursor);

  return {
    isFormula: true,
    inFunctionCall: Boolean(fnName),
    functionName: fnName ?? undefined,
    argIndex: argContext.argIndex,
    currentArg: argContext.currentArg,
    expectingRange: Boolean(fnName && functionRegistry?.isRangeArg?.(fnName, argContext.argIndex)),
  };
}

const UNICODE_ALNUM_RE = (() => {
  try {
    // Match the Rust backend's `char::is_alphanumeric` (Alphabetic || Number).
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    // Older JS engines may not support Unicode property escapes.
    return null;
  }
})();

/**
 * True when `ch` is a valid identifier character for function names / name prefixes.
 *
 * This is intentionally conservative and only needs to support Excel-like function identifiers:
 * letters/digits plus `.` and `_` (e.g. `_xlfn.XLOOKUP`, `WEIBULL.DIST`).
 *
 * @param {string} ch
 */
function isIdentChar(ch) {
  if (!ch) return false;
  // Fast path for ASCII.
  const code = ch.charCodeAt(0);
  if (
    (code >= 48 && code <= 57) || // 0-9
    (code >= 65 && code <= 90) || // A-Z
    (code >= 97 && code <= 122) || // a-z
    code === 46 || // .
    code === 95 // _
  ) {
    return true;
  }

  // Best-effort Unicode support for localized function names (e.g. ZÄHLENWENN).
  return Boolean(UNICODE_ALNUM_RE && UNICODE_ALNUM_RE.test(ch));
}

function clampCursor(input, cursorPosition) {
  const len = typeof input === "string" ? input.length : 0;
  if (!Number.isInteger(cursorPosition)) return len;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > len) return len;
  return cursorPosition;
}

/**
 * Finds a plausible function name token immediately preceding the given '('.
 * @param {string} input
 * @param {number} openParenIndex
 * @returns {string | null}
 */
function findFunctionNameBeforeParen(input, openParenIndex) {
  let i = openParenIndex - 1;
  while (i >= 0 && /\s/.test(input[i])) i--;
  const end = i + 1;
  while (i >= 0 && isIdentChar(input[i])) i--;
  const start = i + 1;
  const token = input.slice(start, end);
  if (!token) return null;
  // Avoid returning something that is obviously a cell ref like "A1".
  if (/^[A-Za-z]{1,3}\d+$/.test(token)) return null;
  return token.toUpperCase();
}

/**
 * Determine current argument index and span inside the current function call.
 * @param {string} input
 * @param {number} openParenIndex
 * @param {number} cursorPosition
 */
function getArgContext(input, openParenIndex, cursorPosition) {
  const baseDepth = 1;
  let depth = baseDepth;
  // Excel locales that use `;` as the argument separator typically use `,` as the
  // decimal separator. To keep completions working in those locales, we prefer
  // semicolons as separators when any are present at the base depth, and otherwise
  // fall back to commas.
  let commaArgIndex = 0;
  let lastCommaIndex = -1;
  let semicolonArgIndex = 0;
  let lastSemicolonIndex = -1;
  let inString = false;
  let inSheetQuote = false;
  let bracketDepth = 0;
  let braceDepth = 0;

  for (let i = openParenIndex + 1; i < cursorPosition; i++) {
    const ch = input[i];
    if (inString) {
      if (ch === '"') {
        if (input[i + 1] === '"') {
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }
    if (inSheetQuote) {
      if (ch === "'") {
        if (input[i + 1] === "'") {
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
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      continue;
    }
    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      continue;
    }
    if (ch === "{") {
      braceDepth += 1;
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }
    if (ch === "(" && bracketDepth === 0) depth++;
    else if (ch === ")" && bracketDepth === 0) depth = Math.max(baseDepth, depth - 1);
    else if (depth === baseDepth && bracketDepth === 0 && braceDepth === 0) {
      if (ch === ",") {
        commaArgIndex++;
        lastCommaIndex = i;
      } else if (ch === ";") {
        semicolonArgIndex++;
        lastSemicolonIndex = i;
      }
    }
  }

  const useSemicolons = lastSemicolonIndex !== -1;
  const argIndex = useSemicolons ? semicolonArgIndex : commaArgIndex;
  const lastArgSeparatorIndex = useSemicolons ? lastSemicolonIndex : lastCommaIndex;

  let rawStart = lastArgSeparatorIndex === -1 ? openParenIndex + 1 : lastArgSeparatorIndex + 1;
  let start = rawStart;
  while (start < cursorPosition && /\s/.test(input[start])) start++;
  const currentArg = {
    start,
    end: cursorPosition,
    text: input.slice(start, cursorPosition),
  };

  return { argIndex, currentArg };
}

/**
 * Finds the last token ending at cursor for function-name completion.
 * @param {string} inputPrefix prefix up to cursor
 * @param {number} cursorPosition
 * @returns {{text:string,start:number,end:number} | null}
 */
function findTokenAtCursor(inputPrefix, cursorPosition) {
  // When completing function names we look at the token at cursor in the formula.
  // Example: "=VLO" => token "VLO" spanning [1, 4).
  let i = cursorPosition - 1;
  while (i >= 0 && isIdentChar(inputPrefix[i])) i--;
  const start = i + 1;
  const end = cursorPosition;
  const text = inputPrefix.slice(start, end);

  // Must be preceded by '=' or an operator/whitespace.
  const before = start - 1 >= 0 ? inputPrefix[start - 1] : "";
  if (before && !/[=\s(,;{+\-*/^]/.test(before)) return null;

  if (!text) return null;
  if (/^\d+$/.test(text)) return null;
  // Avoid treating cell references (A1, BC23, etc.) as function name prefixes.
  if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(text)) return null;
  return { text, start, end };
}
