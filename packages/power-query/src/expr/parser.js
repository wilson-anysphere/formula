import "./ast.js";

import { tokenizeExpr } from "./tokenizer.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 * @typedef {import("./ast.js").ExprToken} ExprToken
 */

const PRECEDENCE = new Map([
  ["||", 2],
  ["&&", 3],
  ["==", 4],
  ["!=", 4],
  ["===", 4],
  ["!==", 4],
  ["<", 5],
  ["<=", 5],
  [">", 5],
  [">=", 5],
  ["+", 6],
  ["-", 6],
  ["*", 7],
  ["/", 7],
  ["%", 7],
]);

const TERNARY_BP = 1;
const UNARY_BP = 8;

const ALLOWED_FUNCTIONS = new Set([
  "date",
  "date_from_text",
  "date_add_days",
  "text_upper",
  "text_lower",
  "text_trim",
  "text_length",
  "text_contains",
  "number_round",
]);

/**
 * @param {string} message
 * @param {string} input
 * @param {ExprToken} token
 * @returns {never}
 */
function parseError(message, input, token) {
  const pos = token.span?.start ?? 0;
  const caret = `${" ".repeat(Math.max(0, pos))}^`;
  throw new Error(`${message}\n${input}\n${caret}`);
}

/**
 * @param {string} op
 * @returns {[number, number] | null}
 */
function infixBindingPower(op) {
  const prec = PRECEDENCE.get(op);
  if (!prec) return null;
  // All supported binary operators are left-associative.
  return [prec, prec + 1];
}

class Parser {
  /**
   * @param {ExprToken[]} tokens
   * @param {string} input
   */
  constructor(tokens, input) {
    this.tokens = tokens;
    this.input = input;
    this.pos = 0;
  }

  /** @returns {ExprToken} */
  peek() {
    return this.tokens[this.pos] ?? { type: "eof", span: { start: this.input.length, end: this.input.length } };
  }

  /** @returns {ExprToken} */
  next() {
    const tok = this.peek();
    this.pos += 1;
    return tok;
  }

  /**
   * @param {ExprToken["type"]} type
   * @param {string} [value]
   * @returns {boolean}
   */
  match(type, value) {
    const tok = this.peek();
    if (tok.type !== type) return false;
    if (value !== undefined && tok.value !== value) return false;
    this.pos += 1;
    return true;
  }

  /**
   * @param {ExprToken["type"]} type
   * @param {string} [value]
   */
  expect(type, value) {
    const tok = this.peek();
    if (tok.type !== type || (value !== undefined && tok.value !== value)) {
      parseError(`Expected ${value ?? type}`, this.input, tok);
    }
    this.pos += 1;
  }

  /** @returns {ExprNode} */
  parse() {
    const expr = this.parseExpression(0, new Set());
    this.expect("eof");
    return expr;
  }

  /**
   * Pratt parser.
   *
   * @param {number} minBp
   * @param {Set<string>} stopOps Operators that should terminate parsing (not consumed).
   * @returns {ExprNode}
   */
  parseExpression(minBp, stopOps) {
    let left = this.parsePrefix(stopOps);

    for (;;) {
      const tok = this.peek();
      if (tok.type === "eof") break;
      if (tok.type !== "operator") break;
      const op = String(tok.value);
      if (stopOps.has(op)) break;

      if (op === "?") {
        if (TERNARY_BP < minBp) break;
        this.next(); // consume '?'

        const consequent = this.parseExpression(0, new Set([...stopOps, ":"]));
        this.expect("operator", ":");
        const alternate = this.parseExpression(TERNARY_BP, stopOps);
        left = { type: "ternary", test: left, consequent, alternate };
        continue;
      }

      const bp = infixBindingPower(op);
      if (!bp) {
        parseError(`Unsupported operator '${op}'`, this.input, tok);
      }
      const [lBp, rBp] = bp;
      if (lBp < minBp) break;

      this.next(); // consume operator
      const right = this.parseExpression(rBp, stopOps);
      left = { type: "binary", op, left, right };
    }

    return left;
  }

  /**
   * @param {Set<string>} stopOps
   * @returns {ExprNode}
   */
  parsePrefix(stopOps) {
    const tok = this.peek();
    if (tok.type === "operator") {
      const op = String(tok.value);
      if (op === "!" || op === "+" || op === "-") {
        this.next();
        const arg = this.parseExpression(UNARY_BP, stopOps);
        return { type: "unary", op, arg };
      }
    }
    return this.parsePrimary(stopOps);
  }

  /**
   * @param {Set<string>} stopOps
   * @returns {ExprNode}
   */
  parsePrimary(stopOps) {
    const tok = this.peek();
    switch (tok.type) {
      case "number":
        this.next();
        return { type: "literal", value: /** @type {number} */ (tok.value) };
      case "string":
        this.next();
        return { type: "literal", value: /** @type {string} */ (tok.value) };
      case "column":
        this.next();
        return { type: "column", name: String(tok.value) };
      case "identifier": {
        this.next();
        const raw = String(tok.value);
        if (raw === "_") return { type: "value" };
        const ident = raw.toLowerCase();
        if (ident === "true") return { type: "literal", value: true };
        if (ident === "false") return { type: "literal", value: false };
        if (ident === "null") return { type: "literal", value: null };

        if (this.match("operator", "(")) {
          // Only allow a small, explicitly whitelisted set of safe functions.
          if (!ALLOWED_FUNCTIONS.has(ident)) {
            parseError(`Unsupported function '${raw}'`, this.input, tok);
          }
          /** @type {ExprNode[]} */
          const args = [];
          if (!this.match("operator", ")")) {
            for (;;) {
              args.push(this.parseExpression(0, new Set([",", ")"])));
              if (this.match("operator", ",")) continue;
              this.expect("operator", ")");
              break;
            }
          }
          return { type: "call", callee: raw, args };
        }

        parseError(`Unsupported identifier '${raw}'`, this.input, tok);
      }
      case "operator":
        if (tok.value === "(") {
          this.next();
          const expr = this.parseExpression(0, new Set([")"]));
          this.expect("operator", ")");
          return expr;
        }
        break;
      default:
        break;
    }

    parseError(`Unexpected token '${tok.type}'`, this.input, tok);
  }
}

/**
 * Parse an expression string into an AST.
 *
 * @param {string} input
 * @returns {ExprNode}
 */
export function parseExpr(input) {
  const tokens = tokenizeExpr(input);
  const parser = new Parser(tokens, input);
  return parser.parse();
}
