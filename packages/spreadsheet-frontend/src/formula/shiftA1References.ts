import { colToName } from "../a1";

function colNameToIndex(col: string): number {
  const normalized = col.toUpperCase();
  let col1 = 0;
  for (const ch of normalized) {
    col1 = col1 * 26 + (ch.charCodeAt(0) - 64);
  }
  return col1 - 1;
}

/**
 * Best-effort A1 reference shifter used by drag-fill (fill handle).
 *
 * Supports:
 * - A1 refs with/without `$` (e.g. `A1`, `$A$1`, `$A1`, `A$1`)
 * - simple ranges (`A1:B2`) by shifting each endpoint independently
 * - basic sheet-qualified refs (`Sheet1!A1`, `'Sheet Name'!A1`)
 *
 * Known limitations (intentionally, for now):
 * - Does not understand structured references / tables (`Table1[Col]`)
 * - Does not parse R1C1 references
 * - May incorrectly treat named ranges that look like cell refs (e.g. `LOG10`)
 *   as cell references unless they are followed by `(`.
 * - Does not fully support escaped quotes in sheet names (Excel uses doubled `'`).
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
  // The leading group captures either start-of-string (empty) or a delimiter
  // character so we can enforce a "token boundary" without needing lookbehind.
  //
  // We avoid matching tokens followed by `(` to reduce false-positives on
  // functions like `LOG10(`.
  const regex =
    /(^|[^A-Za-z0-9_])((?:(?:'[^']+'|[A-Za-z0-9_]+)!)?)(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?!\d)(?!\s*\()/g;

  return segment.replace(regex, (_match, prefix: string, sheetPrefix: string, colAbs: string, col: string, rowAbs: string, row: string) => {
    const col0 = colNameToIndex(col);
    const row0 = Number.parseInt(row, 10) - 1;

    const nextCol0 = colAbs ? col0 : col0 + deltaCols;
    const nextRow0 = rowAbs ? row0 : row0 + deltaRows;

    if (nextCol0 < 0 || nextRow0 < 0) {
      return `${prefix}${sheetPrefix}#REF!`;
    }

    const next = `${sheetPrefix}${colAbs}${colToName(nextCol0)}${rowAbs}${nextRow0 + 1}`;
    return `${prefix}${next}`;
  });
}
