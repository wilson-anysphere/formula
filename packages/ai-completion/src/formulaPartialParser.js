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
  /** @type {{ index: number, functionName: string | null }[]} */
  const openParens = [];
  let inString = false;
  let inSheetQuote = false;
  // Track whether the cursor is currently inside a `[...]` segment.
  //
  // Note: In Excel formulas, `]` inside structured references and external workbook prefixes
  // is escaped as `]]`, which is ambiguous with nested bracket closure (e.g. `[[Col]]`).
  // Use `findMatchingBracketEnd` to skip complete bracket segments and avoid naive depth
  // counting errors that would treat `]]` as two closings.
  let bracketDepth = 0;
  let braceDepth = 0;
  // Track the most recent identifier token so we can cheaply associate it with a following '('
  // (function call). This avoids O(n^2) rescans for formulas with many nested/grouping parens.
  let identStart = null;
  let pendingIdent = null;
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
    // Only track identifiers outside structured references. Identifiers inside `[...]` are
    // table/column names and shouldn't be considered function names.
    if (bracketDepth === 0 && isIdentChar(ch)) {
      if (identStart === null) identStart = i;
      continue;
    }
    if (identStart !== null) {
      pendingIdent = prefix.slice(identStart, i);
      identStart = null;
    }
    if (ch === '"') {
      inString = true;
      pendingIdent = null;
      continue;
    }
    if (ch === "[") {
      pendingIdent = null;
      // Skip a full bracket segment when it closes before the cursor. If it does not close
      // within the prefix, we know the cursor is inside a structured reference / external ref.
      const end = findMatchingBracketEnd(prefix, i, prefix.length);
      if (end == null) {
        bracketDepth = 1;
        break;
      }
      i = end - 1;
      pendingIdent = null;
      continue;
    }
    if (ch === "{") {
      braceDepth += 1;
      pendingIdent = null;
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      pendingIdent = null;
      continue;
    }
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      pendingIdent = null;
      continue;
    }
    if (bracketDepth !== 0) continue;
    if (ch === "(") {
      openParens.push({ index: i, functionName: functionNameFromIdent(pendingIdent, functionRegistry) });
      pendingIdent = null;
    } else if (ch === ")") {
      openParens.pop();
      pendingIdent = null;
    } else if (!/\s/.test(ch)) {
      // Any other non-whitespace token breaks the identifier->'(' link.
      pendingIdent = null;
    }
  }

  const openFunctionCall = findInnermostFunctionCall(openParens);
  if (!openFunctionCall) {
    // When the cursor is inside a string literal / quoted sheet name / structured reference,
    // do not attempt function-name completion. Tab completion should not suggest functions while
    // the user is typing plain text or table column names.
    if (inString || inSheetQuote || bracketDepth !== 0) {
      return { isFormula: true, inFunctionCall: false };
    }
    // Not in a function call; still might be typing a function name.
    const functionPrefix = findTokenAtCursor(prefix, safeCursor, functionRegistry);
    if (functionPrefix && functionPrefix.text.length > 0) {
      return {
        isFormula: true,
        inFunctionCall: false,
        functionNamePrefix: functionPrefix,
      };
    }
    return { isFormula: true, inFunctionCall: false };
  }

  const openParenIndex = openFunctionCall.index;
  const fnName = openFunctionCall.functionName;
  const argContext = getArgContext(prefix, openParenIndex, safeCursor);

  return {
    isFormula: true,
    inFunctionCall: true,
    functionName: fnName,
    argIndex: argContext.argIndex,
    currentArg: argContext.currentArg,
    expectingRange: Boolean(fnName && functionRegistry?.isRangeArg?.(fnName, argContext.argIndex)),
  };
}

/**
 * Returns the innermost open paren that looks like a function call (has an identifier before it).
 *
 * Grouping parentheses (e.g. `=SUM((A1+B1))`) should not shadow the outer function call. We
 * still track them for arg-separator scanning via `getArgContext`, but for completion context
 * we need the *function* paren.
 *
 * @param {{ index: number, functionName: string | null }[]} openParens
 * @returns {{ index: number, functionName: string } | null}
 */
function findInnermostFunctionCall(openParens) {
  for (let i = openParens.length - 1; i >= 0; i--) {
    const entry = openParens[i];
    const fnName = entry?.functionName;
    if (typeof fnName === "string" && fnName.length > 0) {
      return { index: entry.index, functionName: fnName };
    }
  }
  return null;
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
 * Best-effort check: does the host function registry contain any functions that start with
 * the given prefix?
 *
 * This is used to disambiguate a small set of Excel functions that look like A1 cell references
 * (e.g. `LOG10` looks like column `LOG`, row `10`). When the prefix matches a known function,
 * treat it as a function token so completion can still suggest it.
 *
 * @param {unknown} functionRegistry
 * @param {string} prefix
 */
function hasFunctionPrefix(functionRegistry, prefix) {
  const search = functionRegistry && typeof functionRegistry.search === "function" ? functionRegistry.search : null;
  if (!search) return false;
  try {
    const matches = search.call(functionRegistry, prefix, { limit: 1 });
    return Array.isArray(matches) && matches.length > 0;
  } catch {
    return false;
  }
}

/**
 * Normalize a previously-parsed identifier token into a function name candidate.
 *
 * @param {string | null} identToken
 * @param {unknown} functionRegistry
 * @returns {string | null}
 */
function functionNameFromIdent(identToken, functionRegistry) {
  const token = typeof identToken === "string" ? identToken : "";
  if (!token) return null;
  // Avoid returning something that is obviously a cell ref like "A1".
  if (/^[A-Za-z]{1,3}\d+$/.test(token) && !hasFunctionPrefix(functionRegistry, token)) return null;
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
    if (ch === "'") {
      inSheetQuote = true;
      continue;
    }
    if (ch === "[") {
      // Skip bracket segments so we don't treat commas/semicolons inside structured refs
      // as function argument separators.
      const end = findMatchingBracketEnd(input, i, cursorPosition);
      if (end == null) break;
      i = end - 1;
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
    if (ch === "(") depth++;
    else if (ch === ")") depth = Math.max(baseDepth, depth - 1);
    else if (depth === baseDepth && braceDepth === 0) {
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
 * Find the end position (exclusive) for a bracketed segment starting at `startIndex`.
 *
 * Excel uses `]]` to encode a literal `]` inside structured references and external workbook
 * prefixes, which is ambiguous with nested closure (e.g. `[[Col]]`).
 *
 * This matcher prefers treating `]]` as an escaped `]` but will backtrack if that interpretation
 * makes it impossible to close all brackets before `limit`.
 *
 * @param {string} src
 * @param {number} startIndex
 * @param {number} limit exclusive upper bound for scanning
 * @returns {number | null}
 */
function findMatchingBracketEnd(src, startIndex, limit) {
  if (typeof src !== "string") return null;
  const max = typeof limit === "number" && Number.isFinite(limit) ? Math.max(0, Math.min(src.length, limit)) : src.length;
  if (startIndex < 0 || startIndex >= max) return null;
  if (src[startIndex] !== "[") return null;

  let i = startIndex;
  let depth = 0;
  /** @type {Array<{ i: number, depth: number }>} */
  const escapeChoices = [];

  const backtrack = () => {
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
    if (i >= max) {
      if (!backtrack()) return null;
      continue;
    }

    const ch = src[i];
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }
    if (ch === "]") {
      if (src[i + 1] === "]" && depth > 0 && i + 1 < max) {
        // Prefer treating `]]` as an escaped literal `]`. Record a choice point so we can
        // reinterpret it as a real closing bracket if we later fail to close everything.
        escapeChoices.push({ i, depth });
        i += 2;
        continue;
      }
      depth -= 1;
      i += 1;
      if (depth === 0) return i;
      if (depth < 0) {
        // Too many closing brackets; try reinterpreting an earlier escape.
        if (!backtrack()) return null;
      }
      continue;
    }

    i += 1;
  }
}

/**
 * Finds the last token ending at cursor for function-name completion.
 * @param {string} inputPrefix prefix up to cursor
 * @param {number} cursorPosition
 * @returns {{text:string,start:number,end:number} | null}
 */
function findTokenAtCursor(inputPrefix, cursorPosition, functionRegistry) {
  // When completing function names we look at the token at cursor in the formula.
  // Example: "=VLO" => token "VLO" spanning [1, 4).
  let i = cursorPosition - 1;
  while (i >= 0 && isIdentChar(inputPrefix[i])) i--;
  const start = i + 1;
  const end = cursorPosition;
  const text = inputPrefix.slice(start, end);

  // Must be preceded by '=' or an operator/whitespace.
  const before = start - 1 >= 0 ? inputPrefix[start - 1] : "";
  if (before && !/[=\s(,;{+\\\-*/^@<>&]/.test(before)) return null;

  if (!text) return null;
  if (/^\d+$/.test(text)) return null;
  // Avoid treating cell references (A1, BC23, etc.) as function name prefixes.
  if (/^[A-Za-z]{1,3}\d+$/.test(text) && !hasFunctionPrefix(functionRegistry, text)) return null;
  return { text, start, end };
}
