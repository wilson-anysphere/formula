import "./ast.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 */

/**
 * @param {ExprNode} expr
 * @param {(name: string) => number} getColumnIndex
 * @returns {ExprNode}
 */
export function bindExprColumns(expr, getColumnIndex) {
  switch (expr.type) {
    case "column":
      return { ...expr, index: getColumnIndex(expr.name) };
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
 * @returns {unknown}
 */
export function evaluateExpr(expr, values, columnIndex = null) {
  switch (expr.type) {
    case "literal":
      return expr.value;
    case "column": {
      const idx =
        expr.index != null ? expr.index : columnIndex?.get(expr.name) ?? (() => {
          throw new Error(`Unknown column '${expr.name}'`);
        })();
      return values[idx];
    }
    case "unary": {
      const arg = evaluateExpr(expr.arg, values, columnIndex);
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
          const left = evaluateExpr(expr.left, values, columnIndex);
          return left ? evaluateExpr(expr.right, values, columnIndex) : left;
        }
        case "||": {
          const left = evaluateExpr(expr.left, values, columnIndex);
          return left ? left : evaluateExpr(expr.right, values, columnIndex);
        }
        default:
          break;
      }

      const left = evaluateExpr(expr.left, values, columnIndex);
      const right = evaluateExpr(expr.right, values, columnIndex);
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
          // eslint-disable-next-line eqeqeq
          return /** @type {any} */ (left) == /** @type {any} */ (right);
        case "!=":
          // eslint-disable-next-line eqeqeq
          return /** @type {any} */ (left) != /** @type {any} */ (right);
        case "===":
          return left === right;
        case "!==":
          return left !== right;
        default:
          throw new Error(`Unsupported binary operator '${expr.op}'`);
      }
    }
    case "ternary": {
      const test = evaluateExpr(expr.test, values, columnIndex);
      return test ? evaluateExpr(expr.consequent, values, columnIndex) : evaluateExpr(expr.alternate, values, columnIndex);
    }
    case "call": {
      // No functions are currently supported in the formula surface area.
      // This node type exists so we can give a targeted error message and
      // potentially add safe functions in the future.
      throw new Error(`Unsupported function '${expr.callee}'`);
    }
    default: {
      /** @type {never} */
      const exhausted = expr;
      throw new Error(`Unsupported expression node '${exhausted.type}'`);
    }
  }
}

