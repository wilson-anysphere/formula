/**
 * Find the end position (exclusive) for a bracketed segment starting at `startIndex`.
 *
 * Excel uses `]]` to encode a literal `]` inside structured references and external workbook
 * prefixes, which is ambiguous with nested closure (e.g. `[[Col]]`).
 *
 * This matcher prefers treating `]]` as an escaped `]` but will backtrack if that interpretation
 * makes it impossible to close all brackets before `limit`.
 *
 * Workbook prefixes (`[Book.xlsx]Sheet1!A1`) are not nested: `[` is treated as a plain character
 * inside the workbook name, so we fall back to a non-nesting scan when the structured-ref matcher
 * fails. The fallback is gated by a heuristic that checks for either:
 * - a sheet reference suffix ending in `!` (e.g. `[Book.xlsx]Sheet1!A1`), OR
 * - a workbook-scoped defined name (e.g. `[Book.xlsx]MyName`).
 *
 * This is intentionally a small, dependency-free helper that can be reused by the lightweight
 * formula partial parser and completion utilities without pulling in the full formula tokenizer.
 *
 * @param {string} src
 * @param {number} startIndex
 * @param {number} limit exclusive upper bound for scanning
 * @returns {number | null}
 */
export function findMatchingBracketEnd(src, startIndex, limit) {
  if (typeof src !== "string") return null;
  const max =
    typeof limit === "number" && Number.isFinite(limit) ? Math.max(0, Math.min(src.length, Math.trunc(limit))) : src.length;
  if (startIndex < 0 || startIndex >= max) return null;
  if (src[startIndex] !== "[") return null;

  // Prefer structured-ref-style matching (supports nested `[[...]]`). If that fails, fall back to
  // workbook-prefix matching which treats `[` as a literal character.
  return findMatchingStructuredRefBracketEnd(src, startIndex, max) ?? findWorkbookPrefixEndIfValid(src, startIndex, max);
}

const UNICODE_LETTER_RE = (() => {
  try {
    return new RegExp("^\\p{Alphabetic}$", "u");
  } catch {
    return null;
  }
})();

const UNICODE_ALNUM_RE = (() => {
  try {
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    return null;
  }
})();

function isUnicodeAlphabetic(ch) {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isUnicodeAlphanumeric(ch) {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isUnicodeAlphabetic(ch) || (ch >= "0" && ch <= "9");
}

function findWorkbookPrefixEnd(src, startIndex, max) {
  // External workbook prefixes escape closing brackets by doubling: `]]` -> literal `]`.
  //
  // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
  if (src[startIndex] !== "[") return null;
  let i = startIndex + 1;
  while (i < max) {
    if (src[i] === "]") {
      if (i + 1 < max && src[i + 1] === "]") {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i += 1;
  }
  return null;
}

function findWorkbookPrefixEndIfValid(src, startIndex, max) {
  const end = findWorkbookPrefixEnd(src, startIndex, max);
  if (!end) return null;

  const skipWs = (idx) => {
    let i = idx;
    while (i < max && /\s/.test(src[i] ?? "")) i += 1;
    return i;
  };

  const scanQuotedSheetName = (idx) => {
    if (src[idx] !== "'") return null;
    let i = idx + 1;
    while (i < max) {
      const ch = src[i] ?? "";
      if (ch === "'") {
        // Excel escapes apostrophes inside quoted sheet names by doubling: '' -> '
        if (i + 1 < max && src[i + 1] === "'") {
          i += 2;
          continue;
        }
        return i + 1;
      }
      i += 1;
    }
    return null;
  };

  const scanUnquotedName = (idx) => {
    if (idx >= max) return null;
    const first = src[idx] ?? "";
    if (!(first === "_" || isUnicodeAlphabetic(first))) return null;

    let i = idx + 1;
    while (i < max) {
      const ch = src[i] ?? "";
      // Conservative identifier scan: align with Excel-like identifier rules.
      if (ch === "_" || ch === "." || ch === "$" || isUnicodeAlphanumeric(ch)) {
        i += 1;
        continue;
      }
      break;
    }
    return i;
  };

  const scanSheetNameToken = (idx) => {
    const i = skipWs(idx);
    if (i >= max) return null;
    if (src[i] === "'") return scanQuotedSheetName(i);
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
    if (i < max && src[i] === ":") {
      i = scanSheetNameToken(i + 1) ?? i;
      i = skipWs(i);
    }

    if (i < max && src[i] === "!") return end;
  }

  // Workbook-scoped external defined name: `[Book.xlsx]MyName`.
  const nameStart = skipWs(end);
  if (scanUnquotedName(nameStart) != null) return end;

  return null;
}

function findMatchingStructuredRefBracketEnd(src, startIndex, max) {
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
