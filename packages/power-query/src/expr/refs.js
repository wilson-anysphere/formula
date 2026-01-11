import "./ast.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 */

/**
 * @param {ExprNode} expr
 * @param {Set<string>} out
 */
function visit(expr, out) {
  switch (expr.type) {
    case "value":
      return;
    case "column":
      out.add(expr.name);
      return;
    case "literal":
      return;
    case "unary":
      visit(expr.arg, out);
      return;
    case "binary":
      visit(expr.left, out);
      visit(expr.right, out);
      return;
    case "ternary":
      visit(expr.test, out);
      visit(expr.consequent, out);
      visit(expr.alternate, out);
      return;
    case "call":
      expr.args.forEach((arg) => visit(arg, out));
      return;
    default: {
      /** @type {never} */
      const exhausted = expr;
      throw new Error(`Unsupported expression node '${exhausted.type}'`);
    }
  }
}

/**
 * Collect column references (`[Column]`) from an expression AST.
 *
 * @param {ExprNode} expr
 * @returns {Set<string>}
 */
export function collectExprColumnRefs(expr) {
  const out = new Set();
  visit(expr, out);
  return out;
}
