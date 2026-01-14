import { colToName } from "../a1.ts";

function colNameToIndex(col: string): number {
  const normalized = col.toUpperCase();
  let col1 = 0;
  for (const ch of normalized) {
    col1 = col1 * 26 + (ch.charCodeAt(0) - 64);
  }
  return col1 - 1;
}

const MASK_CHAR = "\u0000";

type MaskedSpan = { length: number; text: string };

function findWorkbookPrefixEnd(src: string, start: number): number | null {
  // External workbook prefixes escape closing brackets by doubling: `]]` -> literal `]`.
  //
  // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
  if (src[start] !== "[") return null;
  let i = start + 1;
  while (i < src.length) {
    if (src[i] === "]") {
      if (src[i + 1] === "]") {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i += 1;
  }
  return null;
}

function findMatchingStructuredRefBracketEnd(src: string, start: number): number | null {
  // Structured references escape closing brackets inside items by doubling: `]]` -> literal `]`.
  // That makes naive depth counting incorrect (it will pop twice for an escaped bracket).
  //
  // Match the span using a small backtracking parser:
  // - On `[` increase depth.
  // - On `]]`, prefer treating it as an escape (consume both, depth unchanged), but remember
  //   a choice point. If we later fail to close all brackets, backtrack and reinterpret that
  //   `]]` as a real closing bracket.
  if (src[start] !== "[") return null;

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
    if (i >= src.length) {
      // Unclosed bracket span.
      if (!backtrack()) return null;
      continue;
    }

    const ch = src[i] ?? "";
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }

    if (ch === "]") {
      if (src[i + 1] === "]" && depth > 0) {
        escapeChoices.push({ i, depth });
        i += 2;
        continue;
      }

      depth -= 1;
      i += 1;
      if (depth === 0) return i;
      if (depth < 0) {
        if (!backtrack()) return null;
      }
      continue;
    }

    i += 1;
  }
}

function findBracketSpanEnd(src: string, start: number): number | null {
  // Prefer structured-ref-style matching first (handles nested `[[...]]` spans). If it fails,
  // fall back to workbook-prefix scanning, which treats `[` as a literal character.
  return findMatchingStructuredRefBracketEnd(src, start) ?? findWorkbookPrefixEnd(src, start);
}

function maskBracketSpans(segment: string): { masked: string; spans: MaskedSpan[] } {
  const spans: MaskedSpan[] = [];
  let out = "";

  let i = 0;
  while (i < segment.length) {
    const ch = segment[i];
    if (ch === "[") {
      const end = findBracketSpanEnd(segment, i);
      if (end && end > i) {
        const original = segment.slice(i, end);
        spans.push({ length: end - i, text: original });
        out += MASK_CHAR.repeat(end - i);
        i = end;
        continue;
      }
    }

    out += ch;
    i += 1;
  }

  return { masked: out, spans };
}

function unmaskBracketSpans(segment: string, spans: readonly MaskedSpan[]): string {
  if (spans.length === 0) return segment;

  let out = "";
  let spanIndex = 0;
  let i = 0;

  while (i < segment.length) {
    const ch = segment[i];
    if (ch !== MASK_CHAR) {
      out += ch;
      i += 1;
      continue;
    }

    // Consume a contiguous run of mask characters.
    let j = i;
    while (j < segment.length && segment[j] === MASK_CHAR) j += 1;
    let runLen = j - i;

    while (runLen > 0 && spanIndex < spans.length) {
      const span = spans[spanIndex]!;
      // If the mask run doesn't match our stored spans, fall back to emitting the raw mask chars.
      if (span.length > runLen) break;
      out += span.text;
      runLen -= span.length;
      spanIndex += 1;
    }

    if (runLen > 0) {
      out += MASK_CHAR.repeat(runLen);
    }

    i = j;
  }

  // If we couldn't restore all spans, leave the remainder masked (best-effort).
  return out;
}

/**
 * Best-effort A1 reference shifter used by drag-fill (fill handle).
 *
 * Supports:
 * - A1 refs with/without `$` (e.g. `A1`, `$A$1`, `$A1`, `A$1`)
 * - simple ranges (`A1:B2`) by shifting each endpoint independently
 * - whole-row / whole-column references (`A:A`, `A:B`, `1:1`, `1:10`)
 * - basic sheet-qualified refs (`Sheet1!A1`, `'Sheet Name'!A1`)
 *
 * Known limitations (intentionally, for now):
 * - Does not understand structured references / tables (`Table1[Col]`)
 * - Does not parse R1C1 references
 * - May incorrectly treat named ranges that look like cell refs (e.g. `LOG10`)
 *   as cell references unless they are followed by `(`.
 * - Sheet names are matched using a best-effort regex that supports Excel-style
 *   escaped quotes (doubled `'`), but the formula is not fully parsed.
 */
export function shiftA1References(formula: string, deltaRows: number, deltaCols: number): string {
  if ((deltaRows === 0 && deltaCols === 0) || formula.length === 0) return formula;

  // Only shift outside of double-quoted string literals.
  let result = "";
  let cursor = 0;

  while (cursor < formula.length) {
    const nextQuote = formula.indexOf('"', cursor);
    const end = nextQuote === -1 ? formula.length : nextQuote;
    result += shiftSegment(formula.slice(cursor, end), deltaRows, deltaCols);

    if (nextQuote === -1) break;

    // Copy the string literal verbatim, handling Excel's `""` escape.
    let i = nextQuote;
    let literalEnd = i + 1;
    while (literalEnd < formula.length) {
      if (formula[literalEnd] !== '"') {
        literalEnd++;
        continue;
      }

      if (formula[literalEnd + 1] === '"') {
        literalEnd += 2;
        continue;
      }

      literalEnd++;
      break;
    }

    result += formula.slice(i, literalEnd);
    cursor = literalEnd;
  }

  return result;
}

function shiftSegment(segment: string, deltaRows: number, deltaCols: number): string {
  // Bracket spans (`[...]`) can appear in structured references and external workbook prefixes.
  // Treat them as opaque so we don't accidentally rewrite workbook names or table column names
  // that happen to look like A1 references (e.g. `[A1.xlsx]Sheet1!A1` or `Table1[A1]`).
  const { masked, spans } = maskBracketSpans(segment);

  const sheetPrefixRe = "(?:(?:'(?:[^']|'')+'|[A-Za-z0-9_]+)!)?";
  const tokenBoundaryPrefixRe = "(^|[^A-Za-z0-9_])";

  const shiftCol = (col: string, isAbs: boolean) => (isAbs ? colNameToIndex(col) : colNameToIndex(col) + deltaCols);
  const shiftRow = (row: number, isAbs: boolean) => (isAbs ? row : row + deltaRows);

  const replaceColRange = (
    _match: string,
    prefix: string,
    sheetPrefix: string,
    startAbs: string,
    startCol: string,
    endAbs: string,
    endCol: string
  ) => {
    const nextStart = shiftCol(startCol, Boolean(startAbs));
    const nextEnd = shiftCol(endCol, Boolean(endAbs));
    // The engine formula grammar does not accept sheet-qualified error literals like `Sheet1!#REF!`.
    // Drop the sheet prefix when the rewritten reference becomes invalid.
    if (nextStart < 0 || nextEnd < 0) return `${prefix}#REF!`;
    return `${prefix}${sheetPrefix}${startAbs}${colToName(nextStart)}:${endAbs}${colToName(nextEnd)}`;
  };

  const replaceRowRange = (
    _match: string,
    prefix: string,
    sheetPrefix: string,
    startAbs: string,
    startRow: string,
    endAbs: string,
    endRow: string
  ) => {
    const startRow0 = Number.parseInt(startRow, 10) - 1;
    const endRow0 = Number.parseInt(endRow, 10) - 1;
    const nextStart = shiftRow(startRow0, Boolean(startAbs));
    const nextEnd = shiftRow(endRow0, Boolean(endAbs));
    if (nextStart < 0 || nextEnd < 0) return `${prefix}#REF!`;
    return `${prefix}${sheetPrefix}${startAbs}${nextStart + 1}:${endAbs}${nextEnd + 1}`;
  };

  // The leading group captures either start-of-string (empty) or a delimiter
  // character so we can enforce a "token boundary" without needing lookbehind.
  //
  // We avoid matching tokens followed by `(` to reduce false-positives on
  // functions like `LOG10(`.
  const colRangeRegex = new RegExp(
    `${tokenBoundaryPrefixRe}(${sheetPrefixRe})(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_])(?!\\s*\\()`,
    "g"
  );

  const rowRangeRegex = new RegExp(
    `${tokenBoundaryPrefixRe}(${sheetPrefixRe})(\\$?)([1-9]\\d*):(\\$?)([1-9]\\d*)(?!\\d)(?!\\s*\\()`,
    "g"
  );

  const cellRefRegex = new RegExp(
    `${tokenBoundaryPrefixRe}(${sheetPrefixRe})(\\$?)([A-Za-z]{1,3})(\\$?)([1-9]\\d*)(?!\\d)(?!\\s*\\()`,
    "g"
  );

  const withColRanges = masked.replace(
    colRangeRegex,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    replaceColRange as any
  );

  const withRowRanges = withColRanges.replace(
    rowRangeRegex,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    replaceRowRange as any
  );

  const shifted = withRowRanges.replace(cellRefRegex, (_match, prefix: string, sheetPrefix: string, colAbs: string, col: string, rowAbs: string, row: string) => {
    const col0 = colNameToIndex(col);
    const row0 = Number.parseInt(row, 10) - 1;

    const nextCol0 = colAbs ? col0 : col0 + deltaCols;
    const nextRow0 = rowAbs ? row0 : row0 + deltaRows;

    if (nextCol0 < 0 || nextRow0 < 0) {
      return `${prefix}#REF!`;
    }

    const next = `${sheetPrefix}${colAbs}${colToName(nextCol0)}${rowAbs}${nextRow0 + 1}`;
    return `${prefix}${next}`;
  });

  // Excel drops the spill-range operator (`#`) once the base reference becomes invalid.
  // Our shifter can rewrite `A1#` into `#REF!#`; normalize that to `#REF!` for closer
  // parity with the engine's AST-based rewrite.
  const normalized = shifted.replace(/#REF!#+/g, "#REF!");
  return unmaskBracketSpans(normalized, spans);
}
