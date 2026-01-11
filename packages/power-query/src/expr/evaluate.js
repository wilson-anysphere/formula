import "./ast.js";

import { parseDateLiteral } from "./date.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 */

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {ExprNode} expr
 * @param {(name: string) => number} getColumnIndex
 * @returns {ExprNode}
 */
export function bindExprColumns(expr, getColumnIndex) {
  switch (expr.type) {
    case "column":
      return { ...expr, index: getColumnIndex(expr.name) };
    case "value":
    case "literal":
      return expr;
    case "unary":
      return { ...expr, arg: bindExprColumns(expr.arg, getColumnIndex) };
    case "binary":
      return {
        ...expr,
        left: bindExprColumns(expr.left, getColumnIndex),
        right: bindExprColumns(expr.right, getColumnIndex),
      };
    case "ternary":
      return {
        ...expr,
        test: bindExprColumns(expr.test, getColumnIndex),
        consequent: bindExprColumns(expr.consequent, getColumnIndex),
        alternate: bindExprColumns(expr.alternate, getColumnIndex),
      };
    case "call":
      return { ...expr, args: expr.args.map((a) => bindExprColumns(a, getColumnIndex)) };
    default: {
      /** @type {never} */
      const exhausted = expr;
      throw new Error(`Unsupported expression node '${exhausted.type}'`);
    }
  }
}

/**
 * Evaluate an expression against a row.
 *
 * `bindExprColumns()` can be used ahead of time to avoid repeated column name lookups.
 *
 * @param {ExprNode} expr
 * @param {unknown[]} values
 * @param {Map<string, number> | null} [columnIndex]
 * @param {unknown} [value]
 * @returns {unknown}
 */
export function evaluateExpr(expr, values, columnIndex = null, value = undefined) {
  switch (expr.type) {
    case "literal":
      return expr.value;
    case "value":
      if (value === undefined) {
        throw new Error("Formula references '_' but no value was provided");
      }
      return value;
    case "column": {
      const idx =
        expr.index != null ? expr.index : columnIndex?.get(expr.name) ?? (() => {
          throw new Error(`Unknown column '${expr.name}'`);
        })();
      return values[idx];
    }
    case "unary": {
      const arg = evaluateExpr(expr.arg, values, columnIndex, value);
      switch (expr.op) {
        case "!":
          return !arg;
        case "+":
          // eslint-disable-next-line no-implicit-coercion
          return +/** @type {any} */ (arg);
        case "-":
          // eslint-disable-next-line no-implicit-coercion
          return -/** @type {any} */ (arg);
        default: {
          /** @type {never} */
          const exhausted = expr.op;
          throw new Error(`Unsupported unary operator '${exhausted}'`);
        }
      }
    }
    case "binary": {
      switch (expr.op) {
        case "&&": {
          const left = evaluateExpr(expr.left, values, columnIndex, value);
          return left ? evaluateExpr(expr.right, values, columnIndex, value) : left;
        }
        case "||": {
          const left = evaluateExpr(expr.left, values, columnIndex, value);
          return left ? left : evaluateExpr(expr.right, values, columnIndex, value);
        }
        default:
          break;
      }

      const left = evaluateExpr(expr.left, values, columnIndex, value);
      const right = evaluateExpr(expr.right, values, columnIndex, value);
      switch (expr.op) {
        case "+":
          // eslint-disable-next-line no-implicit-coercion
          return /** @type {any} */ (left) + /** @type {any} */ (right);
        case "-":
          // eslint-disable-next-line no-implicit-coercion
          return /** @type {any} */ (left) - /** @type {any} */ (right);
        case "*":
          // eslint-disable-next-line no-implicit-coercion
          return /** @type {any} */ (left) * /** @type {any} */ (right);
        case "/":
          // eslint-disable-next-line no-implicit-coercion
          return /** @type {any} */ (left) / /** @type {any} */ (right);
        case "%":
          // eslint-disable-next-line no-implicit-coercion
          return /** @type {any} */ (left) % /** @type {any} */ (right);
        case "<":
          return /** @type {any} */ (left) < /** @type {any} */ (right);
        case "<=":
          return /** @type {any} */ (left) <= /** @type {any} */ (right);
        case ">":
          return /** @type {any} */ (left) > /** @type {any} */ (right);
        case ">=":
          return /** @type {any} */ (left) >= /** @type {any} */ (right);
        case "==":
          if (isDate(left) && isDate(right)) return left.getTime() === right.getTime();
          // eslint-disable-next-line eqeqeq
          return /** @type {any} */ (left) == /** @type {any} */ (right);
        case "!=":
          if (isDate(left) && isDate(right)) return left.getTime() !== right.getTime();
          // eslint-disable-next-line eqeqeq
          return /** @type {any} */ (left) != /** @type {any} */ (right);
        case "===":
          if (isDate(left) && isDate(right)) return left.getTime() === right.getTime();
          return left === right;
        case "!==":
          if (isDate(left) && isDate(right)) return left.getTime() !== right.getTime();
          return left !== right;
        default:
          throw new Error(`Unsupported binary operator '${expr.op}'`);
      }
    }
    case "ternary": {
      const test = evaluateExpr(expr.test, values, columnIndex, value);
      return test
        ? evaluateExpr(expr.consequent, values, columnIndex, value)
        : evaluateExpr(expr.alternate, values, columnIndex, value);
    }
    case "call": {
      const callee = expr.callee.toLowerCase();
      switch (callee) {
        case "date": {
          if (expr.args.length !== 1) {
            throw new Error("date() expects exactly 1 argument");
          }
          const arg = evaluateExpr(expr.args[0], values, columnIndex, value);
          if (typeof arg !== "string") {
            throw new Error('date() expects a string like date("2020-01-01")');
          }
          return parseDateLiteral(arg);
        }
        default:
          throw new Error(`Unsupported function '${expr.callee}'`);
      }
    }
    default: {
      /** @type {never} */
      const exhausted = expr;
      throw new Error(`Unsupported expression node '${exhausted.type}'`);
    }
  }
}
