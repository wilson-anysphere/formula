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
    // - trim whitespace
    // - collapse internal whitespace
    // - uppercase (Excel functions are case-insensitive)
    return trimmed.replace(/\s+/g, "").toUpperCase();
  }
}

