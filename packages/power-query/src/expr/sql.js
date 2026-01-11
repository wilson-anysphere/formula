import "./ast.js";

import { parseDateLiteral } from "./date.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 * @typedef {import("../folding/dialect.js").SqlDialect} SqlDialect
 */

/**
 * @typedef {{
 *   sql: string;
 *   params: unknown[];
 * }} SqlFragment
 */

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {string} op
 * @returns {string}
 */
function binaryOpToSql(op) {
  switch (op) {
    case "&&":
      return "AND";
    case "||":
      return "OR";
    case "==":
    case "===":
      return "=";
    case "!=":
    case "!==":
      return "!=";
    default:
      return op;
  }
}

/**
 * Compile an expression AST to a SQL fragment.
 *
 * All non-null literals are parameterized (`?`) and returned via `params` to
 * avoid SQL injection via formulas.
 *
 * @param {ExprNode} expr
 * @param {{
 *   alias: string;
 *   quoteIdentifier: (identifier: string) => string;
 *   dialect: SqlDialect;
 *   knownColumns: string[] | null;
 * }} ctx
 * @returns {SqlFragment}
 */
export function compileExprToSql(expr, ctx) {
  /** @type {unknown[]} */
  const params = [];

  /** @param {unknown} value */
  const param = (value) => {
    if (isDate(value)) {
      params.push(ctx.dialect.formatDateParam(value));
      return "?";
    }
    params.push(value === undefined ? null : value);
    return "?";
  };

  /**
   * @param {ExprNode} node
   * @returns {string}
   */
  function toSql(node) {
    switch (node.type) {
      case "value":
        throw new Error("Value placeholder '_' is not supported in SQL folding");
      case "column": {
        if (ctx.knownColumns && !ctx.knownColumns.includes(node.name)) {
          throw new Error(`Unknown column '${node.name}'`);
        }
        return `${ctx.alias}.${ctx.quoteIdentifier(node.name)}`;
      }
      case "literal":
        if (node.value == null) return "NULL";
        return param(node.value);
      case "unary": {
        const rhs = toSql(node.arg);
        switch (node.op) {
          case "!":
            return `(NOT ${rhs})`;
          case "+":
            return `(+${rhs})`;
          case "-":
            return `(-${rhs})`;
          default:
            throw new Error(`Unsupported unary operator '${node.op}'`);
        }
      }
      case "binary": {
        const left = toSql(node.left);
        const right = toSql(node.right);
        const op = binaryOpToSql(node.op);
        return `(${left} ${op} ${right})`;
      }
      case "ternary": {
        const test = toSql(node.test);
        const cons = toSql(node.consequent);
        const alt = toSql(node.alternate);
        return `(CASE WHEN ${test} THEN ${cons} ELSE ${alt} END)`;
      }
      case "call":
        if (node.callee.toLowerCase() === "date") {
          if (node.args.length !== 1) {
            throw new Error("date() expects exactly 1 argument");
          }
          const arg0 = node.args[0];
          if (arg0.type !== "literal" || typeof arg0.value !== "string") {
            throw new Error('date() expects a string literal like date("2020-01-01")');
          }
          return param(parseDateLiteral(arg0.value));
        }
        throw new Error(`Unsupported function '${node.callee}'`);
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported node '${exhausted.type}'`);
      }
    }
  }

  return { sql: toSql(expr), params };
}
