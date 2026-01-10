import { deepEqual } from "./cell.js";

/**
 * Very small Excel-like formula parser intended only for semantic equivalence
 * during merges.
 *
 * It purposefully does **not** aim for full Excel compatibility; it covers the
 * minimal surface area needed for merge semantics:
 * - whitespace / case-insensitive equivalence
 * - stable parsing of arithmetic expressions and function calls
 * - cell references and ranges
 */

/**
 * @typedef {{
 *   type:
 *     | "number"
 *     | "string"
 *     | "cell"
 *     | "range"
 *     | "name"
 *     | "call"
 *     | "unary"
 *     | "binary",
 *   [key: string]: any
 * }} FormulaAst
 */

/**
 * @typedef {{ type: string, value: string }} Token
 */

/**
 * @param {string} input
 * @returns {Token[]}
 */
function tokenize(input) {
  /** @type {Token[]} */
  const tokens = [];
  let i = 0;

  const push = (type, value) => tokens.push({ type, value });

  const isAlpha = (c) => /[A-Za-z_]/.test(c);
  const isDigit = (c) => /[0-9]/.test(c);

  while (i < input.length) {
    const c = input[i];
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i += 1;
      continue;
    }

    if (c === ",") {
      push("comma", c);
      i += 1;
      continue;
    }
    if (c === "(" || c === ")") {
      push(c, c);
      i += 1;
      continue;
    }

    if (c === "+" || c === "-" || c === "*" || c === "/" || c === "^" || c === "&") {
      push("op", c);
      i += 1;
      continue;
    }

    if (c === '"') {
      let j = i + 1;
      let out = "";
      while (j < input.length) {
        if (input[j] === '"') {
          if (input[j + 1] === '"') {
            out += '"';
            j += 2;
            continue;
          }
          break;
        }
        out += input[j];
        j += 1;
      }
      if (j >= input.length || input[j] !== '"') {
        throw new Error("Unterminated string literal in formula");
      }
      push("string", out);
      i = j + 1;
      continue;
    }

    // number literal: digits[.digits]
    if (isDigit(c) || (c === "." && isDigit(input[i + 1] ?? ""))) {
      let j = i;
      while (isDigit(input[j] ?? "")) j += 1;
      if (input[j] === ".") {
        j += 1;
        while (isDigit(input[j] ?? "")) j += 1;
      }
      push("number", input.slice(i, j));
      i = j;
      continue;
    }

    // Cell reference or identifier
    if (c === "$" || isAlpha(c)) {
      // Range/cell reference pattern.
      const cellMatch = input
        .slice(i)
        .match(/^\$?[A-Za-z]{1,3}\$?\d+/);
      if (cellMatch) {
        const cellRef = cellMatch[0];
        i += cellRef.length;
        if (input[i] === ":" && input.slice(i + 1).match(/^\$?[A-Za-z]{1,3}\$?\d+/)) {
          const other = input.slice(i + 1).match(/^\$?[A-Za-z]{1,3}\$?\d+/)[0];
          push("range", normalizeRef(`${cellRef}:${other}`));
          i += 1 + other.length;
          continue;
        }
        push("cell", normalizeRef(cellRef));
        continue;
      }

      // Identifier (function name / named range)
      let j = i;
      while (isAlpha(input[j] ?? "") || isDigit(input[j] ?? "")) j += 1;
      push("ident", input.slice(i, j).toUpperCase());
      i = j;
      continue;
    }

    throw new Error(`Unexpected character in formula: ${c}`);
  }

  push("eof", "");
  return tokens;
}

/**
 * Normalizes cell references / ranges for case-insensitive comparison.
 * Keeps `$` markers intact.
 *
 * @param {string} ref
 */
function normalizeRef(ref) {
  return ref.toUpperCase();
}

/**
 * @param {Token[]} tokens
 */
function parser(tokens) {
  let pos = 0;

  const peek = () => tokens[pos];
  const consume = (type, value) => {
    const t = tokens[pos];
    if (!t || t.type !== type || (value !== undefined && t.value !== value)) {
      throw new Error(`Expected token ${type}${value ? `(${value})` : ""}`);
    }
    pos += 1;
    return t;
  };

  const PRECEDENCE = {
    "^": 4,
    "*": 3,
    "/": 3,
    "+": 2,
    "-": 2,
    "&": 1
  };

  /**
   * @param {number} minPrec
   * @returns {FormulaAst}
   */
  const parseExpr = (minPrec = 0) => {
    let left = parseUnary();

    while (peek().type === "op" && PRECEDENCE[peek().value] !== undefined) {
      const op = peek().value;
      const prec = PRECEDENCE[op];
      if (prec < minPrec) break;

      consume("op");
      const right = parseExpr(op === "^" ? prec : prec + 1);
      left = { type: "binary", op, left, right };
    }

    return left;
  };

  const parseUnary = () => {
    if (peek().type === "op" && (peek().value === "+" || peek().value === "-")) {
      const op = consume("op").value;
      const expr = parseUnary();
      return { type: "unary", op, expr };
    }
    return parsePrimary();
  };

  const parsePrimary = () => {
    const t = peek();
    if (t.type === "number") {
      consume("number");
      return { type: "number", value: Number(t.value) };
    }
    if (t.type === "string") {
      consume("string");
      return { type: "string", value: t.value };
    }
    if (t.type === "cell") {
      consume("cell");
      return { type: "cell", ref: t.value };
    }
    if (t.type === "range") {
      consume("range");
      return { type: "range", ref: t.value };
    }
    if (t.type === "ident") {
      consume("ident");
      const name = t.value;
      if (peek().type === "(") {
        consume("(");
        /** @type {FormulaAst[]} */
        const args = [];
        if (peek().type !== ")") {
          while (true) {
            args.push(parseExpr(0));
            if (peek().type === "comma") {
              consume("comma");
              continue;
            }
            break;
          }
        }
        consume(")");
        return { type: "call", name, args };
      }
      return { type: "name", name };
    }
    if (t.type === "(") {
      consume("(");
      const expr = parseExpr(0);
      consume(")");
      return expr;
    }
    throw new Error(`Unexpected token in formula: ${t.type}`);
  };

  const expr = parseExpr(0);
  consume("eof");
  return expr;
}

/**
 * @param {string} formula
 * @returns {FormulaAst | null}
 */
export function parseFormulaAst(formula) {
  if (typeof formula !== "string") return null;

  const trimmed = formula.trim();
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  try {
    const tokens = tokenize(withoutEquals);
    return parser(tokens);
  } catch {
    return null;
  }
}

/**
 * @param {string} a
 * @param {string} b
 */
export function areFormulasAstEquivalent(a, b) {
  if (a === b) return true;
  const astA = parseFormulaAst(a);
  const astB = parseFormulaAst(b);
  if (astA && astB) return deepEqual(astA, astB);

  // Fallback to simple normalization: ignore whitespace, normalize case.
  const normalize = (s) => s.replace(/\s+/g, "").toUpperCase();
  return normalize(a) === normalize(b);
}

