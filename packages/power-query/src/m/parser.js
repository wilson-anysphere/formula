import { span } from "./ast.js";
import { MLanguageSyntaxError } from "./errors.js";
import { tokenizeM } from "./tokenizer.js";

/**
 * @typedef {import("./ast.js").MProgram} MProgram
 * @typedef {import("./ast.js").MExpression} MExpression
 * @typedef {import("./ast.js").MIdentifier} MIdentifier
 * @typedef {import("./ast.js").MLetExpression} MLetExpression
 * @typedef {import("./ast.js").MLetBinding} MLetBinding
 * @typedef {import("./ast.js").MIdentifierName} MIdentifierName
 * @typedef {import("./tokenizer.js").Token} Token
 * @typedef {import("./errors.js").MLocation} MLocation
 */

/**
 * Pragmatic recursive-descent parser for a small M language subset.
 *
 * NOTE: This is not intended to be a full, spec-complete M implementation.
 * It focuses on the expression forms produced by common Power Query workflows.
 */
class Parser {
  /**
   * @param {string} source
   * @param {Token[]} tokens
   */
  constructor(source, tokens) {
    this.source = source;
    this.tokens = tokens;
    this.pos = 0;
  }

  /** @returns {Token} */
  peek() {
    return this.tokens[this.pos] ?? this.tokens[this.tokens.length - 1];
  }

  /** @returns {Token} */
  next() {
    const t = this.peek();
    this.pos = Math.min(this.pos + 1, this.tokens.length - 1);
    return t;
  }

  /**
   * @param {Token["type"]} type
   * @param {string} [value]
   * @returns {boolean}
   */
  match(type, value) {
    const t = this.peek();
    if (t.type !== type) return false;
    if (value != null && t.value !== value) return false;
    return true;
  }

  /**
   * @param {Token["type"]} type
   * @param {string} [value]
   * @returns {Token | null}
   */
  consume(type, value) {
    if (!this.match(type, value)) return null;
    return this.next();
  }

  /**
   * @param {string[]} expected
   * @returns {never}
   */
  unexpected(expected) {
    const found = this.peek();
    throw new MLanguageSyntaxError("Unexpected token", {
      location: found.start,
      expected,
      found: { type: found.type, value: found.value },
      source: this.source,
    });
  }

  /**
   * @param {Token["type"]} type
   * @param {string} [value]
   * @param {string[]} [expected]
   * @returns {Token}
   */
  expect(type, value, expected) {
    const t = this.consume(type, value);
    if (t) return t;
    this.unexpected(expected ?? [value ?? type]);
  }

  /** @returns {MProgram} */
  parseProgram() {
    const expr = this.parseExpression();
    const end = this.expect("eof", undefined, ["end of file"]);
    return { type: "Program", expression: expr, span: span(expr.span.start, end.end) };
  }

  /** @returns {MExpression} */
  parseExpression() {
    if (this.match("keyword", "let")) return this.parseLetExpression();
    if (this.match("keyword", "each")) return this.parseEachExpression();
    return this.parseBinaryExpression(0);
  }

  /** @returns {MLetExpression} */
  parseLetExpression() {
    const start = this.expect("keyword", "let").start;
    /** @type {MLetBinding[]} */
    const bindings = [];

    while (!this.match("keyword", "in")) {
      if (this.match("eof")) this.unexpected(["in"]);

      const name = this.parseIdentifierName();
      this.expect("operator", "=", ["="]);
      const value = this.parseExpression();
      const end = value.span.end;
      bindings.push({ name, value, span: span(name.span.start, end) });

      if (this.consume("punct", ",")) continue;
      if (this.match("keyword", "in")) break;
      // Power Query frequently ends bindings with a comma; if we don't see one,
      // treat the next token as an error.
      this.unexpected([",", "in"]);
    }

    this.expect("keyword", "in");
    const body = this.parseExpression();
    return { type: "LetExpression", bindings, body, span: span(start, body.span.end) };
  }

  /** @returns {MExpression} */
  parseEachExpression() {
    const startToken = this.expect("keyword", "each");
    const body = this.parseExpression();
    return { type: "EachExpression", body, span: span(startToken.start, body.span.end) };
  }

  /**
   * Pratt parser (precedence climbing).
   * @param {number} minPrec
   * @returns {MExpression}
   */
  parseBinaryExpression(minPrec) {
    let left = this.parseUnaryExpression();
    while (true) {
      const op = this.peekBinaryOperator();
      if (!op) break;
      const prec = BINARY_PRECEDENCE[op] ?? -1;
      if (prec < minPrec) break;
      const opToken = this.next();
      const right = this.parseBinaryExpression(prec + 1);
      left = {
        type: "BinaryExpression",
        operator: /** @type {any} */ (op),
        left,
        right,
        span: span(left.span.start, right.span.end),
      };

      // Prevent infinite loops in case of zero-width progress.
      if (this.peek() === opToken) break;
    }
    return left;
  }

  /** @returns {MExpression} */
  parseUnaryExpression() {
    if (this.match("keyword", "not")) {
      const start = this.next().start;
      const argument = this.parseUnaryExpression();
      return { type: "UnaryExpression", operator: "not", argument, span: span(start, argument.span.end) };
    }
    if (this.match("operator", "+") || this.match("operator", "-")) {
      const op = this.next();
      const argument = this.parseUnaryExpression();
      return {
        type: "UnaryExpression",
        operator: /** @type {"+" | "-"} */ (op.value),
        argument,
        span: span(op.start, argument.span.end),
      };
    }
    return this.parsePostfixExpression();
  }

  /** @returns {MExpression} */
  parsePostfixExpression() {
    let expr = this.parsePrimary();
    while (true) {
      if (this.match("punct", "(")) {
        expr = this.parseCallExpression(expr);
        continue;
      }
      if (this.match("punct", "[")) {
        this.next();
        const field = this.parseFieldName();
        const end = this.expect("punct", "]", ["]"]).end;
        expr = { type: "FieldAccessExpression", base: expr, field, span: span(expr.span.start, end) };
        continue;
      }
      if (this.match("punct", "{")) {
        this.next();
        const key = this.parseExpression();
        const end = this.expect("punct", "}", ["}"]).end;
        expr = { type: "ItemAccessExpression", base: expr, key, span: span(expr.span.start, end) };
        continue;
      }
      break;
    }
    return expr;
  }

  /**
   * @param {MExpression} callee
   * @returns {MExpression}
   */
  parseCallExpression(callee) {
    this.expect("punct", "(", ["("]);
    /** @type {MExpression[]} */
    const args = [];
    if (!this.match("punct", ")")) {
      while (true) {
        args.push(this.parseExpression());
        if (this.consume("punct", ",")) continue;
        break;
      }
    }
    const end = this.expect("punct", ")", [")"]).end;
    return { type: "CallExpression", callee, args, span: span(callee.span.start, end) };
  }

  /** @returns {MExpression} */
  parsePrimary() {
    const t = this.peek();
    switch (t.type) {
      case "number": {
        const token = this.next();
        return {
          type: "Literal",
          value: Number(token.value),
          literalType: "number",
          span: span(token.start, token.end),
        };
      }
      case "string": {
        const token = this.next();
        return { type: "Literal", value: token.value, literalType: "string", span: span(token.start, token.end) };
      }
      case "keyword": {
        if (t.value === "true" || t.value === "false") {
          const token = this.next();
          return {
            type: "Literal",
            value: token.value === "true",
            literalType: "boolean",
            span: span(token.start, token.end),
          };
        }
        if (t.value === "null") {
          const token = this.next();
          return { type: "Literal", value: null, literalType: "null", span: span(token.start, token.end) };
        }
        if (t.value === "type") return this.parseTypeExpression();
        break;
      }
      case "identifier":
      case "quotedIdentifier":
        return this.parseIdentifierExpression();
      case "punct": {
        if (t.value === "(") {
          const start = this.next().start;
          const expr = this.parseExpression();
          const end = this.expect("punct", ")", [")"]).end;
          return { type: "ParenthesizedExpression", expression: expr, span: span(start, end) };
        }
        if (t.value === "{") return this.parseListExpression();
        if (t.value === "[") return this.parseRecordOrImplicitFieldAccess();
        break;
      }
      default:
        break;
    }
    this.unexpected(["expression"]);
  }

  /** @returns {MExpression} */
  parseTypeExpression() {
    const start = this.expect("keyword", "type").start;
    const name = this.parseQualifiedName();
    const end = this.tokens[this.pos - 1]?.end ?? this.peek().start;
    return { type: "TypeExpression", name, span: span(start, end) };
  }

  /** @returns {MExpression} */
  parseIdentifierExpression() {
    const start = this.peek().start;
    const parts = this.parseQualifiedNameParts();
    const end = this.tokens[this.pos - 1]?.end ?? this.peek().start;
    return { type: "Identifier", parts, span: span(start, end) };
  }

  /** @returns {string} */
  parseQualifiedName() {
    return this.parseQualifiedNameParts().join(".");
  }

  /** @returns {string[]} */
  parseQualifiedNameParts() {
    const first = this.next();
    if (first.type !== "identifier" && first.type !== "quotedIdentifier") {
      this.unexpected(["identifier"]);
    }
    const parts = [first.value];
    while (this.consume("punct", ".")) {
      const seg = this.expect("identifier", undefined, ["identifier"]);
      parts.push(seg.value);
    }
    return parts;
  }

  /** @returns {MExpression} */
  parseListExpression() {
    const start = this.expect("punct", "{", ["{"]).start;
    /** @type {MExpression[]} */
    const elements = [];
    if (!this.match("punct", "}")) {
      while (true) {
        elements.push(this.parseExpression());
        if (this.consume("punct", ",")) continue;
        break;
      }
    }
    const end = this.expect("punct", "}", ["}"]).end;
    return { type: "ListExpression", elements, span: span(start, end) };
  }

  /** @returns {MExpression} */
  parseRecordOrImplicitFieldAccess() {
    const start = this.expect("punct", "[", ["["]).start;
    if (this.match("punct", "]")) {
      const end = this.next().end;
      return { type: "RecordExpression", fields: [], span: span(start, end) };
    }

    // Read the first field name. If followed by '=', this is a record literal.
    /** @type {Token} */
    let keyToken = this.peek();
    /** @type {string} */
    let key = this.parseFieldName();

    if (this.match("operator", "=")) {
      /** @type {{ key: string; value: MExpression; span: import("./errors.js").MSpan }[]} */
      const fields = [];
      while (true) {
        this.expect("operator", "=", ["="]);
        const value = this.parseExpression();
        fields.push({ key, value, span: span(keyToken.start, value.span.end) });

        if (!this.consume("punct", ",")) break;
        if (this.match("punct", "]")) break;

        keyToken = this.peek();
        key = this.parseFieldName();
      }
      const end = this.expect("punct", "]", ["]"]).end;
      return { type: "RecordExpression", fields, span: span(start, end) };
    }

    // Implicit field access: [Field]
    const end = this.expect("punct", "]", ["]"]).end;
    return { type: "FieldAccessExpression", base: null, field: key, span: span(start, end) };
  }

  /** @returns {string} */
  parseFieldName() {
    const t = this.peek();
    if (t.type === "identifier" || t.type === "quotedIdentifier") {
      return this.next().value;
    }
    if (t.type === "string") return this.next().value;
    this.unexpected(["field name"]);
  }

  /** @returns {MIdentifierName} */
  parseIdentifierName() {
    const t = this.peek();
    if (t.type === "identifier") {
      const tok = this.next();
      return { name: tok.value, quoted: false, span: span(tok.start, tok.end) };
    }
    if (t.type === "quotedIdentifier") {
      const tok = this.next();
      return { name: tok.value, quoted: true, span: span(tok.start, tok.end) };
    }
    this.unexpected(["identifier"]);
  }

  /** @returns {string | null} */
  peekBinaryOperator() {
    const t = this.peek();
    if (t.type === "operator" && BINARY_PRECEDENCE[t.value] != null) return t.value;
    if (t.type === "keyword" && (t.value === "and" || t.value === "or")) return t.value;
    return null;
  }
}

const BINARY_PRECEDENCE = {
  or: 1,
  and: 2,
  "=": 3,
  "<>": 3,
  "<": 3,
  "<=": 3,
  ">": 3,
  ">=": 3,
  "+": 4,
  "-": 4,
  "&": 4,
  "*": 5,
  "/": 5,
};

/**
 * Parse an M script into a typed AST.
 *
 * @param {string} source
 * @returns {MProgram}
 */
export function parseM(source) {
  const tokens = tokenizeM(source);
  const parser = new Parser(source, tokens);
  return parser.parseProgram();
}
