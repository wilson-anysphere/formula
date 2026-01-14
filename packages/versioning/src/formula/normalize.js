import { parseFormula } from "./parse.js";

/**
 * @param {number} value
 */
function normalizeNumber(value) {
  if (!Number.isFinite(value)) return String(value);
  // Avoid scientific notation for small ints (purely cosmetic)
  if (Number.isInteger(value)) return String(value);
  // Normalize -0 to 0
  if (Object.is(value, -0)) return "0";
  return String(value);
}

const COMMUTATIVE_BINARY_OPS = new Set(["+", "*"]);
const COMMUTATIVE_FUNCTIONS = new Set(["SUM", "PRODUCT", "MAX", "MIN", "AND", "OR"]);

/**
 * @param {import("./parse.js").AstNode} node
 * @returns {string}
 */
function serializeAst(node) {
  switch (node.type) {
    case "Number":
      return `N(${normalizeNumber(node.value)})`;
    case "String":
      return `S(${JSON.stringify(node.value)})`;
    case "Cell":
      return `C(${node.ref.toUpperCase()})`;
    case "Name":
      return `I(${node.name.toUpperCase()})`;
    case "Percent":
      return `P(${serializeAst(node.expr)})`;
    case "Unary":
      return `U(${node.op},${serializeAst(node.expr)})`;
    case "Range":
      return `R(${serializeAst(node.start)},${serializeAst(node.end)})`;
    case "Binary": {
      if (COMMUTATIVE_BINARY_OPS.has(node.op)) {
        const parts = flattenBinary(node.op, node).map(serializeAst);
        parts.sort();
        return `B(${node.op},[${parts.join(",")}])`;
      }
      return `B(${node.op},${serializeAst(node.left)},${serializeAst(node.right)})`;
    }
    case "Function": {
      const name = node.name.toUpperCase();
      if (COMMUTATIVE_FUNCTIONS.has(name)) {
        const args = flattenFunction(name, node).map(serializeAst);
        args.sort();
        return `F(${name},[${args.join(",")}])`;
      }
      return `F(${name},${node.args.map(serializeAst).join(",")})`;
    }
    default: {
      /** @type {never} */
      const _exhaustive = node;
      return _exhaustive;
    }
  }
}

/**
 * @param {string} op
 * @param {import("./parse.js").AstNode} node
 * @returns {import("./parse.js").AstNode[]}
 */
function flattenBinary(op, node) {
  if (node.type === "Binary" && node.op === op) {
    return [...flattenBinary(op, node.left), ...flattenBinary(op, node.right)];
  }
  return [node];
}

/**
 * @param {string} name
 * @param {import("./parse.js").AstNode} node
 * @returns {import("./parse.js").AstNode[]}
 */
function flattenFunction(name, node) {
  if (node.type === "Function" && node.name.toUpperCase() === name) {
    return node.args.flatMap((arg) => flattenFunction(name, arg));
  }
  return [node];
}

/**
 * Normalize formula to a canonical AST serialization so we can detect semantic
 * equivalence across formatting/case/whitespace and simple commutativity.
 *
 * @param {string | null | undefined} formula
 * @returns {string | null}
 */
export function normalizeFormula(formula) {
  if (formula == null) return null;
  const trimmed = String(formula).trim();
  if (!trimmed) return null;
  try {
    return serializeAst(parseFormula(trimmed));
  } catch {
    // If we can't parse, fall back to a conservative normalization:
    // - remove whitespace *outside* of string literals / quoted sheet names
    // - uppercase *outside* of string literals
    //
    // This avoids incorrectly treating string literals as case/whitespace-insensitive
    // (e.g. ="Hello World" vs ="HelloWorld") while still allowing benign formatting
    // differences to compare equal.
    return normalizeFormulaFallbackText(trimmed);
  }
}

/**
 * Light formula normalization intended for diff/rendering purposes:
 * - trim leading/trailing whitespace
 * - treat empty/whitespace-only as `null`
 * - ensure a leading `=` so callers can consistently tokenize/render formulas
 * - treat a bare `=` (or `=   `) as empty (`null`)
 *
 * This is intentionally *not* semantic normalization; use {@link normalizeFormula}
 * when you need to compare formulas for semantic equivalence.
 *
 * @param {string | null | undefined} formula
 * @returns {string | null}
 */
export function normalizeFormulaText(formula) {
  if (formula == null) return null;
  const trimmed = String(formula).trim();
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = withoutEquals.trim();
  if (!stripped) return null;
  return `=${stripped}`;
}

/**
 * Best-effort textual normalization when the minimal AST parser can't handle a
 * formula (e.g. comparisons, structured references, etc).
 *
 * The key invariant is that we must not change the contents of string literals
 * (double quotes) since they are semantically significant in Excel.
 *
 * Quoted sheet names in formulas use single quotes; those are case-insensitive,
 * but whitespace is significant, so we preserve it.
 *
 * @param {string} input
 */
function normalizeFormulaFallbackText(input) {
  let out = "";
  let inString = false;
  let inQuotedSheet = false;

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

  const isUnicodeAlphabetic = (ch) => {
    if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
    return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
  };

  const isUnicodeAlphanumeric = (ch) => {
    if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
    return isUnicodeAlphabetic(ch) || (ch >= "0" && ch <= "9");
  };

  const findMatchingStructuredRefBracketEnd = (start) => {
    // Excel structured references escape closing brackets inside items by doubling: `]]` -> `]`.
    // That makes naive depth counting incorrect when trying to find the end of a bracket span.
    //
    // We use a small backtracking matcher:
    // - On `[` increase depth.
    // - On `]]`, prefer treating it as an escape (consume both, depth unchanged), but remember
    //   a choice point. If we later fail to close all brackets, backtrack and reinterpret that
    //   `]]` as a real closing bracket.
    if (input[start] !== "[") return null;

    let i = start;
    let depth = 0;
    /** @type {Array<{ i: number; depth: number }>} */
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
      if (i >= input.length) {
        if (!backtrack()) return null;
        if (depth === 0) return i;
        continue;
      }

      const ch = input[i];
      if (ch === "[") {
        depth += 1;
        i += 1;
        continue;
      }

      if (ch === "]") {
        if (input[i + 1] === "]" && depth > 0) {
          escapeChoices.push({ i, depth });
          i += 2;
          continue;
        }

        depth -= 1;
        i += 1;
        if (depth === 0) return i;
        if (depth < 0) {
          if (!backtrack()) return null;
          if (depth === 0) return i;
        }
        continue;
      }

      i += 1;
    }
  };

  const findWorkbookPrefixEnd = (start) => {
    // External workbook prefixes escape closing brackets by doubling: `]]` -> literal `]`.
    //
    // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
    if (input[start] !== "[") return null;
    let i = start + 1;
    while (i < input.length) {
      if (input[i] === "]") {
        if (input[i + 1] === "]") {
          i += 2;
          continue;
        }
        return i + 1;
      }
      i += 1;
    }
    return null;
  };

  const findWorkbookPrefixEndIfValid = (start) => {
    const end = findWorkbookPrefixEnd(start);
    if (!end) return null;

    const skipWs = (idx) => {
      let i = idx;
      while (i < input.length && /\s/.test(input[i] ?? "")) i += 1;
      return i;
    };

    const scanQuotedSheetName = (idx) => {
      if (input[idx] !== "'") return null;
      let i = idx + 1;
      while (i < input.length) {
        const ch = input[i] ?? "";
        if (ch === "'") {
          // Excel escapes apostrophes inside quoted sheet names by doubling: '' -> '
          if (i + 1 < input.length && input[i + 1] === "'") {
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
      if (idx >= input.length) return null;
      const first = input[idx] ?? "";
      if (!(first === "_" || isUnicodeAlphabetic(first))) return null;

      let i = idx + 1;
      while (i < input.length) {
        const ch = input[i] ?? "";
        // Be conservative: align with the Rust parser's unquoted identifier rules.
        // (Names and unquoted sheet identifiers share similar constraints in formula text.)
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
      if (i >= input.length) return null;
      if (input[i] === "'") return scanQuotedSheetName(i);
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
      if (i < input.length && input[i] === ":") {
        i = scanSheetNameToken(i + 1) ?? i;
        i = skipWs(i);
      }

      if (i < input.length && input[i] === "!") return end;
    }

    // Workbook-scoped external defined name: `[Book.xlsx]MyName`.
    const nameStart = skipWs(end);
    if (scanUnquotedName(nameStart) != null) return end;

    return null;
  };

  const findMatchingBracketEnd = (start) =>
    findMatchingStructuredRefBracketEnd(start) ?? findWorkbookPrefixEndIfValid(start);

  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];

    if (inString) {
      out += ch;
      if (ch === "\"") {
        // Escaped quote inside string literal: ""
        if (input[i + 1] === "\"") {
          out += "\"";
          i += 1;
        } else {
          inString = false;
        }
      }
      continue;
    }

    if (inQuotedSheet) {
      if (ch === "'") {
        out += "'";
        // Escaped single quote inside sheet name: ''
        if (input[i + 1] === "'") {
          out += "'";
          i += 1;
        } else {
          inQuotedSheet = false;
        }
      } else {
        // Sheet names are case-insensitive, but whitespace is significant.
      out += ch.toUpperCase();
      }
      continue;
    }

    if (ch === "\"") {
      inString = true;
      out += ch;
      continue;
    }

    if (ch === "'") {
      inQuotedSheet = true;
      out += ch;
      continue;
    }

    if (ch === "[") {
      const end = findMatchingBracketEnd(i);
      if (!end) {
        // Unterminated bracket span. Be conservative: preserve the remaining contents (including
        // whitespace) since it may be part of a structured reference item name.
        out += input.slice(i).toUpperCase();
        break;
      }
      // Preserve whitespace inside bracket spans because it can be significant in structured
      // reference item names (e.g. `Table1[Total Amount]`).
      out += input.slice(i, end).toUpperCase();
      i = end - 1;
      continue;
    }

    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      continue;
    }

    out += ch.toUpperCase();
  }

  return out;
}
