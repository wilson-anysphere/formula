export function rewriteSheetNamesInFormula(
  formula: string,
  oldName: string,
  newName: string,
): string {
  const out: string[] = [];
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i];

    if (inString) {
      out.push(ch);
      if (ch === '"') {
        if (formula[i + 1] === '"') {
          out.push('"');
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
      out.push('"');
      i += 1;
      continue;
    }

    if (ch === "'") {
      const parsed = parseQuotedSheetSpec(formula, i);
      if (parsed) {
        const { nextIndex, sheetSpec } = parsed;
        const rewritten =
          rewriteSheetSpec(sheetSpec, oldName, newName) ?? quoteSheetSpec(sheetSpec);
        out.push(rewritten, "!");
        i = nextIndex;
        continue;
      }
    }

    const parsedUnquoted = parseUnquotedSheetSpec(formula, i);
    if (parsedUnquoted) {
      const { nextIndex, sheetSpec } = parsedUnquoted;
      out.push(rewriteSheetSpec(sheetSpec, oldName, newName) ?? sheetSpec, "!");
      i = nextIndex;
      continue;
    }

    out.push(ch);
    i += 1;
  }

  return out.join("");
}

function normalizeSheetNameForCaseInsensitiveCompare(name: string): string {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  //
  // Match the semantics used by `@formula/workbook-backend` and the Rust backend.
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

/**
 * Rewrite sheet references inside a formula after deleting a sheet.
 *
 * Excel behavior (approximated):
 * - Direct references to the deleted sheet become `#REF!` (e.g. `=Sheet2!A1` → `=#REF!`).
 * - 3D references shift boundaries when the deleted sheet is a boundary
 *   (e.g. `=SUM(Sheet1:Sheet3!A1)` with `Sheet1` deleted → `=SUM(Sheet2:Sheet3!A1)`).
 *
 * This implementation is intentionally conservative: it only rewrites tokens that parse as sheet
 * references and it does not touch string literals.
 */
export function rewriteDeletedSheetReferencesInFormula(
  formula: string,
  deletedSheet: string,
  sheetOrder: string[],
): string {
  const out: string[] = [];
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i];

    if (inString) {
      out.push(ch);
      if (ch === '"') {
        if (formula[i + 1] === '"') {
          out.push('"');
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
      out.push('"');
      i += 1;
      continue;
    }

    if (ch === "#") {
      const parsed = parseErrorLiteral(formula, i);
      if (parsed) {
        out.push(parsed.raw);
        i = parsed.nextIndex;
        continue;
      }
    }

    if (ch === "'") {
      const parsed = parseQuotedSheetSpec(formula, i);
      if (parsed) {
        const raw = formula.slice(i, parsed.nextIndex); // includes trailing '!'
        const rewrite = rewriteSheetSpecForDelete(parsed.sheetSpec, deletedSheet, sheetOrder);
        if (rewrite.kind === "unchanged") {
          out.push(raw);
          i = parsed.nextIndex;
          continue;
        }
        if (rewrite.kind === "adjusted") {
          out.push(rewrite.spec, "!");
          i = parsed.nextIndex;
          continue;
        }
        // invalidate
        out.push("#REF!");
        i = sheetRefTailEnd(formula, parsed.nextIndex);
        continue;
      }
    }

    const parsedUnquoted = parseUnquotedSheetSpec(formula, i);
    if (parsedUnquoted) {
      const raw = formula.slice(i, parsedUnquoted.nextIndex); // includes trailing '!'
      const rewrite = rewriteSheetSpecForDelete(parsedUnquoted.sheetSpec, deletedSheet, sheetOrder);
      if (rewrite.kind === "unchanged") {
        out.push(raw);
        i = parsedUnquoted.nextIndex;
        continue;
      }
      if (rewrite.kind === "adjusted") {
        out.push(rewrite.spec, "!");
        i = parsedUnquoted.nextIndex;
        continue;
      }
      out.push("#REF!");
      i = sheetRefTailEnd(formula, parsedUnquoted.nextIndex);
      continue;
    }

    out.push(ch);
    i += 1;
  }

  return out.join("");
}

function parseQuotedSheetSpec(
  formula: string,
  startIndex: number,
): { nextIndex: number; sheetSpec: string } | null {
  if (formula[startIndex] !== "'") return null;

  let i = startIndex + 1;
  const content: string[] = [];

  while (i < formula.length) {
    const ch = formula[i];
    if (ch === "'") {
      if (formula[i + 1] === "'") {
        content.push("'");
        i += 2;
        continue;
      }
      i += 1;
      break;
    }
    content.push(ch);
    i += 1;
  }

  if (formula[i] !== "!") return null;

  return { nextIndex: i + 1, sheetSpec: content.join("") };
}

function parseUnquotedSheetSpec(
  formula: string,
  startIndex: number,
): { nextIndex: number; sheetSpec: string } | null {
  const first = codePointAt(formula, startIndex);
  if (!first) return null;

  // Match the Rust backend's behavior:
  // - Accept Unicode letters (not just ASCII) for unquoted sheet names.
  // - Support external workbook prefixes like `[Book.xlsx]Sheet1!A1`.
  if (first.ch !== "[" && first.ch !== "_" && !isUnicodeAlphabetic(first.ch)) return null;

  let i = first.nextIndex;

  // External workbook prefix: `[Book1.xlsx]Sheet1!A1`
  if (first.ch === "[") {
    while (i < formula.length) {
      const next = codePointAt(formula, i);
      if (!next) return null;
      if (next.ch === "]") {
        i = next.nextIndex;
        break;
      }
      i = next.nextIndex;
    }

    if (i >= formula.length) return null;

    const after = codePointAt(formula, i);
    if (!after) return null;
    if (after.ch !== "_" && !isUnicodeAlphabetic(after.ch)) {
      // Likely a structured reference rather than an external workbook reference.
      return null;
    }
  }

  while (i < formula.length) {
    const next = codePointAt(formula, i);
    if (!next) return null;
    if (next.ch === "!") {
      return { nextIndex: next.nextIndex, sheetSpec: formula.slice(startIndex, i) };
    }
    if (next.ch === "_" || next.ch === "." || next.ch === ":" || isUnicodeAlphanumeric(next.ch)) {
      i = next.nextIndex;
      continue;
    }
    break;
  }

  return null;
}

function rewriteSheetSpec(sheetSpec: string, oldName: string, newName: string): string | null {
  const { workbookPrefix, remainder } = splitWorkbookPrefix(sheetSpec);
  const [start, end] = split3d(remainder);

  const renamedStart = startEquals(start, oldName) ? newName : start;
  const renamedEnd = end && startEquals(end, oldName) ? newName : end;

  if (renamedStart === start && renamedEnd === end) return null;

  return formatSheetReference(workbookPrefix, renamedStart, renamedEnd);
}

function splitWorkbookPrefix(sheetSpec: string): { workbookPrefix: string | null; remainder: string } {
  // External references can include `[` / `]` inside a path component
  // (e.g. `'C:\\[foo]\\[Book.xlsx]Sheet1'!A1`). The workbook delimiter is the last `[...]` pair.
  const openIdx = sheetSpec.lastIndexOf("[");
  if (openIdx === -1) return { workbookPrefix: null, remainder: sheetSpec };
  const closeIdx = sheetSpec.indexOf("]", openIdx);
  if (closeIdx === -1) return { workbookPrefix: null, remainder: sheetSpec };
  const prefixEnd = closeIdx + 1;
  if (prefixEnd >= sheetSpec.length) return { workbookPrefix: null, remainder: sheetSpec };
  return { workbookPrefix: sheetSpec.slice(0, prefixEnd), remainder: sheetSpec.slice(prefixEnd) };
}

function split3d(remainder: string): [string, string | null] {
  const idx = remainder.indexOf(":");
  if (idx === -1) return [remainder, null];
  return [remainder.slice(0, idx), remainder.slice(idx + 1)];
}

function startEquals(a: string, b: string): boolean {
  return normalizeSheetNameForCaseInsensitiveCompare(a) === normalizeSheetNameForCaseInsensitiveCompare(b);
}

function quoteSheetSpec(sheetSpec: string): string {
  return `'${sheetSpec.replace(/'/g, "''")}'`;
}

function isValidUnquotedSheetName(name: string): boolean {
  if (!name) return false;
  const first = name[0];
  if (!first || isAsciiDigit(first)) return false;
  if (!(isAsciiLetter(first) || first === "_")) return false;
  for (let i = 1; i < name.length; i += 1) {
    const ch = name[i];
    if (!(isAsciiAlphaNum(ch) || ch === "_" || ch === ".")) return false;
  }
  if (isReservedUnquotedSheetName(name)) return false;
  if (looksLikeA1CellReference(name) || looksLikeR1C1CellReference(name)) return false;
  return true;
}

function needsQuotingForSheetReference(name: string): boolean {
  const [start, end] = split3d(name);
  if (end !== null) {
    return !(isValidUnquotedSheetName(start) && isValidUnquotedSheetName(end));
  }
  return !isValidUnquotedSheetName(name);
}

function formatSheetReference(
  workbookPrefix: string | null,
  start: string,
  end: string | null,
): string {
  const content = `${workbookPrefix ?? ""}${start}${end ? `:${end}` : ""}`;
  return needsQuotingForSheetReference(content) ? quoteSheetSpec(content) : content;
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiAlphaNum(ch: string): boolean {
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function isReservedUnquotedSheetName(name: string): boolean {
  // Excel boolean literals (`TRUE`/`FALSE`) are tokenized as keywords; quoting avoids ambiguity.
  // Match the Rust backend's `is_reserved_unquoted_sheet_name`.
  return name.toLowerCase() === "true" || name.toLowerCase() === "false";
}

function looksLikeA1CellReference(name: string): boolean {
  // If an unquoted sheet name looks like a cell reference (e.g. "A1" or "XFD1048576"),
  // Excel requires quoting to disambiguate.
  //
  // Match the Rust backend's `looks_like_a1_cell_reference`.
  let i = 0;
  let letters = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiLetter(ch)) break;
    if (letters.length >= 3) return false;
    letters += ch;
    i += 1;
  }

  if (letters.length === 0) return false;

  let digits = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiDigit(ch)) break;
    digits += ch;
    i += 1;
  }

  if (digits.length === 0) return false;
  if (i !== name.length) return false;

  const col = letters
    .split("")
    .reduce((acc, c) => acc * 26 + (c.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1), 0);
  return col <= 16_384;
}

function looksLikeR1C1CellReference(name: string): boolean {
  // In R1C1 notation, `R`/`C` are valid relative references. Excel may also treat
  // `R123C456` as a cell reference even when the workbook is in A1 mode.
  //
  // Match the Rust backend's `looks_like_r1c1_cell_reference`.
  const upper = name.toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;

  let i = 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;

  i += 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  return i === upper.length;
}

const UNICODE_LETTER_RE: RegExp | null = (() => {
  try {
    return new RegExp("^\\p{L}$", "u");
  } catch {
    return null;
  }
})();

const UNICODE_ALNUM_RE: RegExp | null = (() => {
  try {
    return new RegExp("^[\\p{L}\\p{N}]$", "u");
  } catch {
    return null;
  }
})();

function isUnicodeAlphabetic(ch: string): boolean {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return isAsciiLetter(ch);
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isAsciiAlphaNum(ch);
}

function codePointAt(str: string, index: number): { ch: string; nextIndex: number } | null {
  if (index < 0 || index >= str.length) return null;
  const cp = str.codePointAt(index);
  if (cp == null) return null;
  return { ch: String.fromCodePoint(cp), nextIndex: index + (cp > 0xffff ? 2 : 1) };
}

type DeleteSheetSpecRewrite =
  | { kind: "unchanged" }
  | { kind: "adjusted"; spec: string }
  | { kind: "invalidate" };

function rewriteSheetSpecForDelete(
  sheetSpec: string,
  deletedSheet: string,
  sheetOrder: string[],
): DeleteSheetSpecRewrite {
  const { workbookPrefix, remainder } = splitWorkbookPrefix(sheetSpec);
  const [start, end] = split3d(remainder);

  if (end == null) {
    return startEquals(start, deletedSheet) ? { kind: "invalidate" } : { kind: "unchanged" };
  }

  const startMatches = startEquals(start, deletedSheet);
  const endMatches = startEquals(end, deletedSheet);
  if (!startMatches && !endMatches) return { kind: "unchanged" };

  const startIdx = sheetIndexInOrder(sheetOrder, start);
  const endIdx = sheetIndexInOrder(sheetOrder, end);
  if (startIdx == null || endIdx == null) return { kind: "invalidate" };

  // Span references only the deleted sheet.
  if (startIdx === endIdx) return { kind: "invalidate" };

  const dir = endIdx > startIdx ? 1 : -1;
  let newStartIdx = startIdx;
  let newEndIdx = endIdx;

  if (startMatches) newStartIdx += dir;
  if (endMatches) newEndIdx -= dir;

  const newStart = sheetOrder[newStartIdx];
  const newEnd = sheetOrder[newEndIdx];
  if (!newStart || !newEnd) return { kind: "invalidate" };

  const nextEnd = startEquals(newStart, newEnd) ? null : newEnd;
  return { kind: "adjusted", spec: formatSheetReference(workbookPrefix, newStart, nextEnd) };
}

function sheetIndexInOrder(sheetOrder: string[], name: string): number | null {
  const target = normalizeSheetNameForCaseInsensitiveCompare(name);
  for (let i = 0; i < sheetOrder.length; i += 1) {
    const candidate = sheetOrder[i];
    if (candidate && normalizeSheetNameForCaseInsensitiveCompare(candidate) === target) return i;
  }
  return null;
}

function parseErrorLiteral(formula: string, startIndex: number): { nextIndex: number; raw: string } | null {
  if (formula[startIndex] !== "#") return null;
  let i = startIndex + 1;
  while (i < formula.length) {
    const ch = formula[i];
    if (isAsciiAlphaNum(ch) || ch === "/" || ch === "_" || ch === ".") {
      i += 1;
      continue;
    }
    if (ch === "!" || ch === "?") {
      i += 1;
      break;
    }
    break;
  }
  if (i === startIndex + 1) return null;
  return { nextIndex: i, raw: formula.slice(startIndex, i) };
}

function sheetRefTailEnd(formula: string, startIndex: number): number {
  let i = startIndex;
  let bracketDepth = 0;
  let parenDepth = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i];

    if (inString) {
      if (ch === '"') {
        if (formula[i + 1] === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }

    switch (ch) {
      case '"':
        inString = true;
        i += 1;
        continue;
      case "[":
        bracketDepth += 1;
        i += 1;
        continue;
      case "]":
        bracketDepth = Math.max(0, bracketDepth - 1);
        i += 1;
        continue;
      case "(":
        parenDepth += 1;
        i += 1;
        continue;
      case ")":
        if (parenDepth === 0) return i;
        parenDepth = Math.max(0, parenDepth - 1);
        i += 1;
        continue;
      default:
        break;
    }

    if (bracketDepth === 0 && parenDepth === 0) {
      if (
        ch === " " ||
        ch === "\t" ||
        ch === "\n" ||
        ch === "\r" ||
        ch === "," ||
        ch === ";" ||
        ch === "+" ||
        ch === "-" ||
        ch === "*" ||
        ch === "/" ||
        ch === "^" ||
        ch === "&" ||
        ch === "=" ||
        ch === "<" ||
        ch === ">" ||
        ch === "{" ||
        ch === "}" ||
        ch === "%"
      ) {
        return i;
      }
    }

    i += 1;
  }

  return i;
}
