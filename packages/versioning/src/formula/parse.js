import { tokenizeFormula } from "./tokenize.js";

/**
 * @typedef {{ type: "Number", value: number }} NumberNode
 * @typedef {{ type: "String", value: string }} StringNode
 * @typedef {{ type: "Cell", ref: string }} CellNode
 * @typedef {{ type: "Name", name: string }} NameNode
 * @typedef {{ type: "Unary", op: string, expr: AstNode }} UnaryNode
 * @typedef {{ type: "Binary", op: string, left: AstNode, right: AstNode }} BinaryNode
 * @typedef {{ type: "Function", name: string, args: AstNode[] }} FunctionNode
 * @typedef {{ type: "Range", start: AstNode, end: AstNode }} RangeNode
 * @typedef {{ type: "Percent", expr: AstNode }} PercentNode
 * @typedef {NumberNode | StringNode | CellNode | NameNode | UnaryNode | BinaryNode | FunctionNode | RangeNode | PercentNode} AstNode
 */

/**
 * Minimal parser for semantic diff normalization. Not a full Excel grammar.
 *
 * @param {string} formula
 * @returns {AstNode}
 */
export function parseFormula(formula) {
  const trimmed = formula.trim();
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const tokens = tokenizeFormula(withoutEquals);
  let pos = 0;

  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  const expect = (type, value) => {
    const t = next();
    if (t.type !== type || (value !== undefined && t.value !== value)) {
      throw new Error(`Expected ${type}${value ? `(${value})` : ""} but got ${t.type}(${t.value})`);
    }
    return t;
  };

  const isOp = (v) => peek().type === "op" && peek().value === v;
  const isPunct = (v) => peek().type === "punct" && peek().value === v;
  const isIdent = () => peek().type === "ident";

  function parseExpression() {
    return parseAdditive();
  }

  function parseAdditive() {
    let node = parseMultiplicative();
    while (isOp("+") || isOp("-")) {
      const op = next().value;
      const right = parseMultiplicative();
      node = { type: "Binary", op, left: node, right };
    }
    return node;
  }

  function parseMultiplicative() {
    let node = parsePower();
    while (isOp("*") || isOp("/")) {
      const op = next().value;
      const right = parsePower();
      node = { type: "Binary", op, left: node, right };
    }
    return node;
  }

  // Exponentiation is right-associative
  function parsePower() {
    let node = parseUnary();
    if (isOp("^")) {
      const op = next().value;
      const right = parsePower();
      node = { type: "Binary", op, left: node, right };
    }
    return node;
  }

  function parseUnary() {
    if (isOp("+") || isOp("-")) {
      const op = next().value;
      return { type: "Unary", op, expr: parseUnary() };
    }
    let node = parsePrimary();
    // Postfix percent operator
    while (isOp("%")) {
      next();
      node = { type: "Percent", expr: node };
    }
    return node;
  }

  function parsePrimary() {
    const t = peek();

    if (isPunct("(")) {
      next();
      const node = parseExpression();
      expect("punct", ")");
      return node;
    }

    if (t.type === "number") {
      next();
      const num = Number(t.value);
      if (Number.isNaN(num)) {
        throw new Error(`Invalid number literal: ${t.value}`);
      }
      return { type: "Number", value: num };
    }

    if (t.type === "string") {
      next();
      return { type: "String", value: t.value };
    }

    if (t.type === "ident") {
      return parseIdentifierOrReferenceOrFunction();
    }

    throw new Error(`Unexpected token: ${t.type}(${t.value})`);
  }

  function parseIdentifierOrReferenceOrFunction() {
    const first = expect("ident").value;

    // Sheet ref: Sheet1!A1 or 'My Sheet'!A1
    if (isOp("!")) {
      next();
      const second = expect("ident").value;
      const ref = `${first}!${second}`;
      return parseMaybeRangeOrReturnRef(ref);
    }

    // Function call
    if (isPunct("(")) {
      next();
        /** @type {AstNode[]} */
        const args = [];
        if (!isPunct(")")) {
          while (true) {
            args.push(parseExpression());
            // In Excel, argument separators are typically `,` but some locales use `;`.
            // Our tokenizer classifies both as punctuation.
            if (isPunct(",") || isPunct(";")) {
              next();
              continue;
            }
            break;
          }
        }
      expect("punct", ")");
      return { type: "Function", name: first, args };
    }

    return parseMaybeRangeOrReturnRef(first);
  }

  function parseMaybeRangeOrReturnRef(identOrRef) {
    // Range operator: A1:B2
    if (isOp(":")) {
      next();
      const end = expect("ident").value;
      return {
        type: "Range",
        start: identifierToNode(identOrRef),
        end: identifierToNode(end),
      };
    }
    return identifierToNode(identOrRef);
  }

  function identifierToNode(ident) {
    // Detect cell reference. This regex covers basic A1 refs with optional $.
    // It intentionally does NOT validate bounds (A..XFD, 1..1048576) since we
    // only need canonicalization for diff.
    const cellRefRegex = /^(\$?[A-Za-z]{1,3}\$?\d+)(?:\:(\$?[A-Za-z]{1,3}\$?\d+))?$/;
    const m = ident.match(cellRefRegex);
    if (m) {
      return { type: "Cell", ref: ident };
    }
    return { type: "Name", name: ident };
  }

  const ast = parseExpression();
  if (peek().type !== "eof") {
    throw new Error(`Unexpected trailing input: ${peek().type}(${peek().value})`);
  }
  return ast;
}
