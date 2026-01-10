/**
 * A deliberately small formula parser used only for conflict resolution
 * heuristics (AST equivalence + subtree detection).
 *
 * This is NOT intended to be a full Excel parser – it only supports enough
 * syntax to power deterministic tests and provide useful conflict resolution
 * defaults (see docs/06-collaboration.md).
 *
 * Supported:
 * - Numbers (e.g. 1, 3.14)
 * - Cell refs (e.g. A1, $B$2) and ranges (A1:B2)
 * - Function calls (e.g. SUM(A1, 2))
 * - Unary +/- and binary + - * /
 * - Parentheses and commas
 */

/**
 * @typedef {ReturnType<typeof parseFormula>} ParsedFormula
 */

/**
 * @param {string} input
 * @returns {FormulaAst}
 */
export function parseFormula(input) {
  const tokens = tokenize(input);
  const parser = new Parser(tokens);
  const ast = parser.parseFormula();
  parser.assertEof();
  return ast;
}

/**
 * @param {FormulaAst} ast
 * @returns {string}
 */
export function formulaToString(ast) {
  switch (ast.type) {
    case "number":
      return ast.value;
    case "cell":
      return `${ast.col}${ast.row}`;
    case "range":
      return `${formulaToString(ast.start)}:${formulaToString(ast.end)}`;
    case "name":
      return ast.name;
    case "call":
      return `${ast.name}(${ast.args.map(formulaToString).join(",")})`;
    case "unary":
      return `${ast.op}${formulaToString(ast.expr)}`;
    case "binary":
      return `(${formulaToString(ast.left)}${ast.op}${formulaToString(ast.right)})`;
    default:
      return "<unknown>";
  }
}

/**
 * @typedef {FormulaNumber|FormulaCell|FormulaRange|FormulaName|FormulaCall|FormulaUnary|FormulaBinary} FormulaAst
 *
 * @typedef {{type: "number", value: string}} FormulaNumber
 * @typedef {{type: "cell", col: string, row: number}} FormulaCell
 * @typedef {{type: "range", start: FormulaCell, end: FormulaCell}} FormulaRange
 * @typedef {{type: "name", name: string}} FormulaName
 * @typedef {{type: "call", name: string, args: Array<FormulaAst>}} FormulaCall
 * @typedef {{type: "unary", op: "+"|"-", expr: FormulaAst}} FormulaUnary
 * @typedef {{type: "binary", op: "+"|"-"|"*"|"/", left: FormulaAst, right: FormulaAst}} FormulaBinary
 */

/**
 * @param {string} input
 * @returns {Array<Token>}
 */
function tokenize(input) {
  const src = input.trim();
  /** @type {Array<Token>} */
  const out = [];

  let i = 0;
  // Ignore leading '='.
  if (src[i] === "=") i += 1;

  while (i < src.length) {
    const ch = src[i];

    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    if ("()+-*/,:".includes(ch)) {
      out.push({ type: ch, value: ch });
      i += 1;
      continue;
    }

    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < src.length && /[0-9.]/.test(src[j])) j += 1;
      const num = src.slice(i, j);
      if (!/^\d+(\.\d+)?$/.test(num)) {
        throw new Error(`Invalid number literal: ${num}`);
      }
      out.push({ type: "NUMBER", value: num });
      i = j;
      continue;
    }

    if (/[A-Za-z_$]/.test(ch)) {
      let j = i;
      while (j < src.length && /[A-Za-z0-9_$]/.test(src[j])) j += 1;
      const ident = src.slice(i, j);
      out.push({ type: "IDENT", value: ident });
      i = j;
      continue;
    }

    throw new Error(`Unexpected character '${ch}' at ${i}`);
  }

  return out;
}

/**
 * @typedef {object} Token
 * @property {string} type
 * @property {string} value
 */

class Parser {
  /**
   * @param {Array<Token>} tokens
   */
  constructor(tokens) {
    this.tokens = tokens;
    this.i = 0;
  }

  /** @returns {Token|null} */
  peek() {
    return this.tokens[this.i] ?? null;
  }

  /** @returns {Token|null} */
  consume() {
    const t = this.peek();
    if (!t) return null;
    this.i += 1;
    return t;
  }

  /**
   * @param {string} type
   * @returns {Token}
   */
  expect(type) {
    const t = this.consume();
    if (!t || t.type !== type) {
      throw new Error(`Expected token ${type} but got ${t ? t.type : "EOF"}`);
    }
    return t;
  }

  assertEof() {
    if (this.peek() != null) {
      throw new Error(`Unexpected trailing tokens starting at ${this.peek().value}`);
    }
  }

  /** @returns {FormulaAst} */
  parseFormula() {
    return this.parseAdditive();
  }

  /** @returns {FormulaAst} */
  parseAdditive() {
    let left = this.parseMultiplicative();
    // eslint-disable-next-line no-constant-condition
    while (true) {
      const t = this.peek();
      if (!t || (t.type !== "+" && t.type !== "-")) break;
      this.consume();
      const right = this.parseMultiplicative();
      left = { type: "binary", op: /** @type {"+"|"-"} */ (t.type), left, right };
    }
    return left;
  }

  /** @returns {FormulaAst} */
  parseMultiplicative() {
    let left = this.parseUnary();
    // eslint-disable-next-line no-constant-condition
    while (true) {
      const t = this.peek();
      if (!t || (t.type !== "*" && t.type !== "/")) break;
      this.consume();
      const right = this.parseUnary();
      left = { type: "binary", op: /** @type {"*"|"/"} */ (t.type), left, right };
    }
    return left;
  }

  /** @returns {FormulaAst} */
  parseUnary() {
    const t = this.peek();
    if (t && (t.type === "+" || t.type === "-")) {
      this.consume();
      const expr = this.parseUnary();
      return { type: "unary", op: /** @type {"+"|"-"} */ (t.type), expr };
    }
    return this.parsePrimary();
  }

  /** @returns {FormulaAst} */
  parsePrimary() {
    const t = this.peek();
    if (!t) throw new Error("Unexpected EOF");

    if (t.type === "NUMBER") {
      this.consume();
      return { type: "number", value: t.value };
    }

    if (t.type === "IDENT") {
      // Function call if followed by '('.
      const identToken = this.consume();
      const maybeParen = this.peek();
      if (maybeParen?.type === "(") {
        this.consume(); // '('
        /** @type {Array<FormulaAst>} */
        const args = [];
        if (this.peek()?.type !== ")") {
          args.push(this.parseFormula());
          while (this.peek()?.type === ",") {
            this.consume();
            args.push(this.parseFormula());
          }
        }
        this.expect(")");
        return {
          type: "call",
          name: identToken.value.toUpperCase(),
          args
        };
      }

      // Range or cell ref?
      const maybeCell = parseCellIdent(identToken.value);
      if (maybeCell) {
        const next = this.peek();
        if (next?.type === ":") {
          this.consume();
          const endToken = this.expect("IDENT");
          const endCell = parseCellIdent(endToken.value);
          if (!endCell) {
            throw new Error(`Invalid range end cell ref: ${endToken.value}`);
          }
          return { type: "range", start: maybeCell, end: endCell };
        }
        return maybeCell;
      }

      // Named range / identifier.
      return { type: "name", name: identToken.value.toUpperCase() };
    }

    if (t.type === "(") {
      this.consume();
      const expr = this.parseFormula();
      this.expect(")");
      return expr;
    }

    throw new Error(`Unexpected token: ${t.type}`);
  }
}

/**
 * @param {string} ident
 * @returns {import("./formula-parser.js").FormulaCell|null}
 */
function parseCellIdent(ident) {
  // Strip $ anchors.
  const stripped = ident.replaceAll("$", "");
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(stripped);
  if (!match) return null;
  const [, col, rowStr] = match;
  return { type: "cell", col: col.toUpperCase(), row: Number.parseInt(rowStr, 10) };
}

